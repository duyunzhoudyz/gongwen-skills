# -*- coding: utf-8 -*-
"""改稿润色的辅助统计：篇幅统计、改稿前后对照、合并稿溯源、提取 Word 批注、导出正文。

属于 gongwen-polish，源文件就在 skills/gongwen-polish/scripts/，不由同步脚本生成。
只读输入文件，不访问网络；除 body --out 外不写任何文件，body --out 不覆盖已有文件（同名时加 -v2、-v3）。

用法：
    python scripts/polish_stat.py count <文件> [--target 字数] [--max] [--json] [--strict]
    python scripts/polish_stat.py compare <原稿> <改稿> [--json]
    python scripts/polish_stat.py trace <合并稿> <来源稿> [<来源稿>...] [--threshold 0.6] [--json] [--strict]
    python scripts/polish_stat.py comments <带批注的.docx> [--json]
    python scripts/polish_stat.py body <文件> [<文件>...] [--out 输出.txt]

count：统计正文字数，按一级标题（一、二、……）分部分统计；给了 --target 时，算出与目标的差距和各部分
       按比例分配的参考字数。默认允许目标±10%；加 --max 表示目标是上限（用户说“以内”“不超过”），
       超过即报“超出”，低于上限七成时提示复查是否删了必保内容。
compare：对照原稿和改稿：总字数和各部分字数的变化，逐段列出修改、移动、拆分、删除、新增的段落，
         用来核对每一处改动都在任务范围内（如每处改动都对应某条批注）。“新增”指在原稿中找不到来源的段落；
         “修改”“移动”“拆分”另列出改稿段比原稿段多出的文字，用来查句内追加的新事实；
         并比对状态和义务强度词（已、正在、拟、将、必须、应当、要、可以、力争、基本、全面等）的增减，
         单字的改动（如“拟”改“已”）也列出。
         段落按顺序配对（带权最长公共子序列），再把剩下的新增段落与全部原稿段比一次，认出移动和拆分。
trace：把合并稿逐句与各来源稿比对，标出每句来自哪一稿；找不到来源的句子可能是新增内容，须逐条核对；
       有来源的句子再按逗号切成小句，列出找不到来源的小句；同时列出来源稿中没有被采用的句子。
comments：从 .docx 中提取 Word 批注：批注内容、所在段落、批注圈定的原文。不输出批注人姓名。
body：导出【起草说明】之前的全部文字（含标题、版头），给几份文件时依次拼接。用于把 docx 转成文字稿，
      或把原稿和材料拼成 fact_diff --ledger 的参照文件（每份文件的【起草说明】及其后内容都不带入）。

统计口径：标题之后、【起草说明】之前的正文，不含空白字符。
标题行：开头的版头行（以“文件”结尾的发文机关标志、发文字号行）之后的第一个非空行，不超过 60 字、
不以句末标点或冒号结尾、不以层次序号开头的，视为标题。标题折行沿用机检 gongwen_lint.py 的规则：
标题还没有以文种（通知、报告、汇报等）结尾时，后面不超过 40 字、不含句读和括号的行算标题续行，
标题最多 3 行；紧跟标题以破折号开头的行视为副标题。只有标题、副标题的稿件，篇幅与机检一致；
带份号、密级、版记等版头版记要素的正式文件，以机检的篇幅为准。

输入：.txt、.md（UTF-8，兼容 GB18030）、.docx。读 .docx 只用标准库，按“接受全部修订”后的文字读取
（含修订插入的文字，不含修订删除的文字），表格一行为一段、单元格之间用制表符分隔。
退出码：0 正常；1 使用 --strict 时 count 超出目标（或低于目标±10%的下限），或 trace 有找不到来源的句子、小句；
2 输入错误。
"""
from __future__ import annotations

import argparse
import difflib
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Set, Tuple

BOUNDARY = "【起草说明】"
# 分界行写法与 fact_diff.py、gongwen_lint.py 保持一致：兼容 Markdown 标题、加粗包裹
DRAFT_NOTE_RE = re.compile(r"^\s*(?:#{1,6}\s*)?(?:(?:\*\*|__)\s*)?" + re.escape(BOUNDARY))
SUPPORTED_SUFFIXES = (".txt", ".md", ".docx")

H1_RE = re.compile(r"^\s*[一二三四五六七八九十]{1,3}、")
HEADING_PREFIX_RE = re.compile(
    r"^\s*(?:[一二三四五六七八九十]{1,3}、|[（(][一二三四五六七八九十]{1,3}[）)]"
    r"|\d{1,2}[.．、]|[（(]\d{1,2}[）)]|第[一二三四五六七八九十]{1,3}[，、,])"
)
HEAD_LINE_RE = re.compile(r"(?:文件|〔\s*\d{4}\s*〕\s*\d+\s*号)\s*$")
SUBTITLE_RE = re.compile(r"^\s*(?:——|—|－－|--)")
# 标题折行：与 gongwen_lint.py 的 TITLE_END_RE、_title_continuation 保持一致
TITLE_DOC_END_RE = re.compile(
    r"(?:决议|决定|命令|公报|公告|通告|意见|通知|通报|报告|请示|批复|议案|函|纪要|签报|总结|计划|要点|材料|方案"
    r"|讲话|发言|致辞|信息|简报|办法|规定|细则|规则|制度|章程|预案|规划|说明|汇报|提纲)$"
)
TITLE_TAIL_PUNCT = "。，；：、！？.,;:!?"
TITLE_CONT_BAD_RE = re.compile(r"[。；：:，,（(]")
TITLE_HEADING_START_RE = re.compile(
    r"^\s*(?:[一二三四五六七八九十]{1,3}\s*[、，,.．]|[（(]\s*[一二三四五六七八九十\d]{1,3}\s*[）)]|\d{1,2}\s*[.．、])"
)
DATE_ONLY_RE = re.compile(
    r"^\s*[\d〇○零一二三四五六七八九OＯ]{4}\s*年\s*[\d一二三四五六七八九十]{1,3}\s*月\s*[\d一二三四五六七八九十]{1,3}\s*日\s*$"
)
TITLE_MAX_LINES = 3
TITLE_CONT_MAX_CHARS = 40
TITLE_END_PUNCT = "。！？；：!?;:，,"
MARKER_RE = re.compile(r"【(待补|待核实)(?:[：:][^】\n]*)?】")
SENT_SPLIT_RE = re.compile(r"(?<=[。！？；!?;])")
CLAUSE_SPLIT_RE = re.compile(r"(?<=[，,])")
# 比对前去掉的字符：空白和常见标点
STRIP_RE = re.compile(
    r"[\s，。、；：？！“”‘’「」『』《》〈〉（）()【】\[\]〔〕—…·．.,;:?!'\"\-－～~/／|｜]"
)

TITLE_MAX_CHARS = 60
TARGET_TOLERANCE = 0.10
UPPER_FLOOR = 0.70          # 目标为上限时，低于上限的这个比例提示复查必保内容
EXCERPT_CHARS = 40
MIN_TRACE_CHARS = 6
MIN_CLAUSE_CHARS = 4        # trace 逐小句核对来源时，去掉标点后不足 4 字的小句不比对
PAIR_THRESHOLD = 0.6        # 按顺序配对为“修改”的最低重合度
MOVE_THRESHOLD = 0.75       # 不按顺序认定为“移动”“拆分”的最低重合度
SHORT_PAIR_THRESHOLD = 0.6  # 短段（如小标题）按整体相似度配对的门槛
SHORT_CONTAIN_THRESHOLD = 0.9  # 短段的文字有九成以上出现在原稿段中，也可配对（压缩时只留小标题句）
CONTAIN_WEIGHT = 0.7        # 按包含率配对的权重低于整体相似配对，免得小标题被含同样字眼的长段抢走
SHORT_PARA_CHARS = 12
MAX_INSERTED_SHOWN = 3
# 状态和义务强度词：compare 逐对比较增减，单字也比（“拟”改“已”只多一个字，“多出的文字”列不出来）。
# 单字词排除常见的复合词，免得“重要”“要求”“模拟”之类的词干扰计数。
STATUS_WORDS: Tuple[Tuple[str, "re.Pattern[str]"], ...] = tuple(
    (word, re.compile(pattern)) for word, pattern in (
        ("已", r"(?<![而得])已"),
        ("正在", r"正在"),
        ("尚未", r"尚未"),
        ("拟", r"(?<![模虚草])拟(?![定写稿制])"),
        ("将", r"将(?![军士领])"),
        ("必须", r"必须"),
        ("务必", r"务必"),
        ("应当", r"应当"),
        ("不得", r"不得"),
        ("要", r"(?<![重主需纲摘首紧必简提扼只])要(?![求点素件害闻务领])"),
        ("可以", r"可以"),
        ("力争", r"力争"),
        ("基本", r"基本"),
        ("全面", r"全面"),
    )
)
DEFAULT_THRESHOLD = 0.6
MAX_SOURCES = 6
MAX_LIST_ITEMS = 60
MAX_XML_BYTES = 50 * 1024 * 1024
SOURCE_LABELS = ("甲稿", "乙稿", "丙稿", "丁稿", "戊稿", "己稿")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_P = f"{{{W_NS}}}p"
W_T = f"{{{W_NS}}}t"
W_ID = f"{{{W_NS}}}id"
W_DATE = f"{{{W_NS}}}date"
W_COMMENT = f"{{{W_NS}}}comment"
W_CRS = f"{{{W_NS}}}commentRangeStart"
W_CRE = f"{{{W_NS}}}commentRangeEnd"
W_CREF = f"{{{W_NS}}}commentReference"
W_INS = f"{{{W_NS}}}ins"
W_DEL = f"{{{W_NS}}}del"
W_MOVE_FROM = f"{{{W_NS}}}moveFrom"
W_MOVE_TO = f"{{{W_NS}}}moveTo"
W_BODY = f"{{{W_NS}}}body"
W_TBL = f"{{{W_NS}}}tbl"
W_TR = f"{{{W_NS}}}tr"
W_TC = f"{{{W_NS}}}tc"
W_TAB = f"{{{W_NS}}}tab"
W_BR = f"{{{W_NS}}}br"
W_CR = f"{{{W_NS}}}cr"
W_SECTPR = f"{{{W_NS}}}sectPr"


class InputError(Exception):
    """输入文件不存在、格式不支持或读取失败。"""


# ---------------------------------------------------------------------------
# 读取
# ---------------------------------------------------------------------------

def _decode(data: bytes, path: Path) -> str:
    for encoding in ("utf-8-sig", "gb18030"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise InputError(f"无法识别文件编码（请另存为 UTF-8）：{path.name}")


def _check_path(path: Path) -> None:
    if not path.exists():
        raise InputError(f"未找到文件：{path}")
    if not path.is_file():
        raise InputError(f"不是文件：{path}")
    if path.suffix.lower() not in SUPPORTED_SUFFIXES:
        raise InputError(f"只支持 .txt、.md、.docx：{path.name}")


def _read_xml(archive: zipfile.ZipFile, name: str) -> Optional[ET.Element]:
    try:
        info = archive.getinfo(name)
    except KeyError:
        return None
    if info.file_size > MAX_XML_BYTES:
        raise InputError(f"{name} 过大，已停止读取")
    try:
        return ET.fromstring(archive.read(name))
    except ET.ParseError as exc:
        raise InputError(f"{name} 不是合法的 XML：{exc}") from exc


def _open_docx(path: Path) -> Tuple[ET.Element, Optional[ET.Element]]:
    """返回（document.xml 根元素，comments.xml 根元素或 None）。"""
    try:
        archive = zipfile.ZipFile(str(path))
    except (zipfile.BadZipFile, OSError) as exc:
        raise InputError(f"无法打开 docx 文件：{exc}") from exc
    with archive:
        document = _read_xml(archive, "word/document.xml")
        if document is None:
            raise InputError("不是有效的 Word 文档：缺少 word/document.xml")
        comments_root = _read_xml(archive, "word/comments.xml")
    return document, comments_root


def _docx_body(document: ET.Element) -> ET.Element:
    body = document.find(W_BODY)
    return document if body is None else body


def _para_text(paragraph: ET.Element) -> str:
    """段落文字，按“接受全部修订”取：含 w:ins、w:moveTo 内的文字，不含 w:del、w:moveFrom 内的文字。
    文本框等嵌套段落不计入外层段落。"""
    parts: List[str] = []

    def walk(element: ET.Element) -> None:
        for child in element:
            tag = child.tag
            if tag in (W_DEL, W_MOVE_FROM, W_P):
                continue
            if tag == W_T:
                parts.append(child.text or "")
            elif tag == W_TAB:
                parts.append("\t")
            elif tag in (W_BR, W_CR):
                parts.append("\n")
            else:
                walk(child)

    walk(paragraph)
    return "".join(parts)


def _docx_paragraphs(document: ET.Element) -> List[ET.Element]:
    """按文档顺序返回正文中的段落元素（含表格、内容控件中的段落，不含文本框内的嵌套段落）。"""
    result: List[ET.Element] = []

    def walk(element: ET.Element) -> None:
        for child in element:
            if child.tag == W_P:
                result.append(child)
            elif child.tag != W_SECTPR:
                walk(child)

    walk(_docx_body(document))
    return result


def _count_revisions(document: ET.Element) -> int:
    return sum(1 for el in document.iter() if el.tag in (W_INS, W_DEL, W_MOVE_FROM, W_MOVE_TO))


def _docx_lines(document: ET.Element) -> List[str]:
    """正文逐段取文字；表格一行为一段，单元格之间用制表符分隔。"""
    lines: List[str] = []

    def cell_text(cell: ET.Element) -> str:
        texts: List[str] = []

        def walk(element: ET.Element) -> None:
            for child in element:
                if child.tag == W_P:
                    texts.append(_para_text(child).replace("\n", " ").strip())
                else:
                    walk(child)

        walk(cell)
        return " ".join(t for t in texts if t)

    def walk(element: ET.Element) -> None:
        for child in element:
            tag = child.tag
            if tag == W_P:
                lines.extend(_para_text(child).split("\n"))
            elif tag == W_TBL:
                for row in (r for r in child if r.tag == W_TR):  # 只取本表的行，嵌套表格并入单元格文字
                    cells = [cell_text(tc) for tc in row if tc.tag == W_TC]
                    lines.append("\t".join(cells))
            elif tag != W_SECTPR:
                walk(child)

    walk(_docx_body(document))
    return lines


def read_source(path: Path) -> Tuple[List[str], int]:
    """读取文件，返回（按行切分的文本，docx 中的修订痕迹处数）。docx 按“接受全部修订”后的文字读取。"""
    _check_path(path)
    if path.suffix.lower() == ".docx":
        document, _ = _open_docx(path)
        return _docx_lines(document), _count_revisions(document)
    try:
        data = path.read_bytes()
    except OSError as exc:
        raise InputError(f"读取失败：{path.name}（{exc}）") from exc
    return _decode(data, path).splitlines(), 0


def read_lines(path: Path) -> List[str]:
    """读取文件，返回按行（docx 按段落）切分的文本。"""
    return read_source(path)[0]


# ---------------------------------------------------------------------------
# 正文结构
# ---------------------------------------------------------------------------

def nonspace_len(text: str) -> int:
    return len(re.sub(r"\s", "", text))


def _looks_like_title(text: str) -> bool:
    stripped = text.strip()
    if not stripped or nonspace_len(stripped) > TITLE_MAX_CHARS:
        return False
    if stripped[-1] in TITLE_END_PUNCT:
        return False
    if HEADING_PREFIX_RE.match(stripped) or MARKER_RE.search(stripped):
        return False
    return True


def _title_continuation(text: str) -> bool:
    """标题续行（与 gongwen_lint.py 的 _title_continuation 一致）：不超过 40 字，不含句读和括号。"""
    stripped = text.strip()
    if not stripped or len(stripped) > TITLE_CONT_MAX_CHARS or TITLE_HEADING_START_RE.match(stripped):
        return False
    if SUBTITLE_RE.match(stripped) or DATE_ONLY_RE.match(stripped) or TITLE_CONT_BAD_RE.search(stripped):
        return False
    return not MARKER_RE.search(stripped)


class Body:
    """正文（【起草说明】之前）的标题和段落。"""

    def __init__(self, name: str, lines: Sequence[str], include_title: bool = False,
                 revisions: int = 0) -> None:
        self.name = name
        self.revisions = revisions  # docx 中的修订痕迹处数；已按接受全部修订后的文字读取
        paras: List[str] = []
        for line in lines:
            if DRAFT_NOTE_RE.match(line):
                break
            if line.strip():
                paras.append(line.strip())
        start = 0
        while start < len(paras) and start < 3 and nonspace_len(paras[start]) <= 40 \
                and HEAD_LINE_RE.search(paras[start]):
            start += 1  # 版头：发文机关标志、发文字号
        self.title: List[str] = []
        if not include_title and len(paras) - start >= 2 and _looks_like_title(paras[start]):
            self.title.append(paras[start])
            j = start + 1
            # 折行标题：至少留一段正文
            while (j < len(paras) - 1 and len(self.title) < TITLE_MAX_LINES
                   and not TITLE_DOC_END_RE.search("".join(self.title).rstrip(TITLE_TAIL_PUNCT))
                   and _title_continuation(paras[j])):
                self.title.append(paras[j])
                j += 1
            if j < len(paras) and SUBTITLE_RE.match(paras[j]):
                self.title.append(paras[j])
        self.paras: List[str] = paras[start + len(self.title):]

    @property
    def char_count(self) -> int:
        return sum(nonspace_len(p) for p in self.paras)

    def markers(self) -> Dict[str, int]:
        counts = {"待补": 0, "待核实": 0}
        for para in self.paras:
            for match in MARKER_RE.finditer(para):
                counts[match.group(1)] += 1
        return counts

    def sections(self) -> List[Dict[str, object]]:
        """按一级标题切分：第一个一级标题之前为“开头”。"""
        result: List[Dict[str, object]] = []
        current: Dict[str, object] = {"heading": "开头", "char_count": 0, "paragraphs": 0, "first_para": 1}
        for index, para in enumerate(self.paras, start=1):
            if H1_RE.match(para):
                if current["paragraphs"]:
                    result.append(current)
                current = {"heading": _heading_text(para), "char_count": 0, "paragraphs": 0,
                           "first_para": index}
            current["char_count"] = int(current["char_count"]) + nonspace_len(para)
            current["paragraphs"] = int(current["paragraphs"]) + 1
        if current["paragraphs"]:
            result.append(current)
        for number, sec in enumerate(result, start=1):
            sec["index"] = number
        return result


def _heading_text(para: str) -> str:
    """标题与正文同段时，只取标题部分。"""
    text = para.strip()
    cut = re.search(r"[。！？；]", text)
    if cut:
        text = text[:cut.start()]
    return _excerpt(text, 30)


def _excerpt(text: str, limit: int = EXCERPT_CHARS) -> str:
    text = re.sub(r"\s+", "", text)
    return text if len(text) <= limit else text[:limit] + "…"


def load_body(path: Path, include_title: bool = False) -> Body:
    lines, revisions = read_source(path)
    return Body(path.name, lines, include_title=include_title, revisions=revisions)


def _revision_note(name: str, revisions: int) -> str:
    return (f"{name} 含修订痕迹 {revisions} 处，已按“接受全部修订”后的文字统计；"
            f"fact_diff.py 读 docx 时会漏掉修订插入的文字，比对前先用 body 子命令另存为 .txt。")


# ---------------------------------------------------------------------------
# count
# ---------------------------------------------------------------------------

BASIS_NOTE = ("统计口径：标题之后、【起草说明】之前的正文，不含空白字符；标题（含折行标题、副标题）的识别与机检 "
              "gongwen_lint.py 一致，带份号、密级、版记等要素的正式文件以机检的篇幅为准。")
FAIL_STATUSES = ("超出", "不足")


def count_report(path: Path, target: Optional[int] = None, include_title: bool = False,
                 upper: bool = False) -> Dict[str, object]:
    if target is not None and target <= 0:
        raise InputError("--target 必须是正整数")
    if upper and target is None:
        raise InputError("--max 需要和 --target 一起使用")
    body = load_body(path, include_title=include_title)
    total = body.char_count
    sections = body.sections()
    for sec in sections:
        sec["share"] = round(int(sec["char_count"]) / total, 3) if total else 0.0
    report: Dict[str, object] = {
        "file": body.name,
        "title": "".join(body.title),
        "char_count": total,
        "paragraphs": len(body.paras),
        "markers": body.markers(),
        "sections": sections,
        "target": None,
        "revisions": body.revisions,
        "basis": BASIS_NOTE,
    }
    if target is not None:
        diff = total - target
        if upper:
            # 用户说“以内”“不超过”：目标就是上限
            low, high, mode = int(round(target * UPPER_FLOOR)), target, "上限"
            if total > high:
                status, advice = "超出", f"比上限多 {diff} 字，至少还要删减 {diff} 字"
            elif total < low:
                status, advice = "偏少", (f"不到上限的七成（{low} 字），复查是否删去了事实、数据、措施、"
                                          f"责任主体、时限等必保内容；没有误删的照常交付，不为凑字数加内容")
            else:
                status, advice = "在范围内", "未超过上限"
        else:
            low = int(round(target * (1 - TARGET_TOLERANCE)))
            high = int(round(target * (1 + TARGET_TOLERANCE)))
            mode = "±10%"
            if total > high:
                status, advice = "超出", f"比目标多 {diff} 字，至少删减 {total - high} 字才进入允许范围"
            elif total < low:
                status, advice = "不足", f"比目标少 {-diff} 字；扩写只能展开已有事实，素材不够时按实际篇幅交付"
            else:
                status, advice = "在范围内", "已在目标±10%以内"
        for sec in sections:
            sec["target"] = int(round(target * float(sec["share"])))
        report["target"] = {"value": target, "mode": mode, "low": low, "high": high, "diff": diff,
                            "status": status, "advice": advice}
    return report


def render_count(report: Dict[str, object]) -> str:
    lines = [f"篇幅统计：{report['file']}"]
    if report["title"]:
        lines.append(f"标题（不计入）：{report['title']}")
    markers = report["markers"]
    lines.append(f"正文 {report['char_count']} 字，{report['paragraphs']} 段；"
                 f"【待补】{markers['待补']} 处，【待核实】{markers['待核实']} 处。")
    target = report["target"]
    if target:
        if target["mode"] == "上限":
            scope = f"不超过 {target['value']} 字（低于 {target['low']} 字时复查必保内容）"
        else:
            scope = f"{target['value']} 字，允许 {target['low']}～{target['high']} 字"
        lines.append(f"目标 {scope}：{target['status']}。{target['advice']}。")
    if report.get("revisions"):
        lines.append("提示：" + _revision_note(str(report["file"]), int(report["revisions"])))
    lines.append("各部分：")
    for sec in report["sections"]:
        extra = f"，按比例参考 {sec['target']} 字" if target else ""
        lines.append(f"  {sec['index']}. {sec['heading']}：{sec['char_count']} 字，{sec['paragraphs']} 段，"
                     f"占 {float(sec['share']) * 100:.1f}%{extra}")
    if target:
        lines.append("提示：按比例分配只作参考。压缩按“铺垫、复述、空泛表态、重复举例”的顺序删，不平均砍。")
    lines.append(str(report["basis"]))
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# compare
# ---------------------------------------------------------------------------

def _norm(text: str) -> str:
    return STRIP_RE.sub("", text)


def _norm_heading(text: str) -> str:
    return _norm(HEADING_PREFIX_RE.sub("", text))


def _containment(original: str, revised: str) -> float:
    """改稿段中有多大比例的文字按顺序出现在原稿段中（只计两字以上的连续片段）。"""
    if not revised:
        return 0.0
    blocks = difflib.SequenceMatcher(None, original, revised, autojunk=False).get_matching_blocks()
    return sum(block.size for block in blocks if block.size >= 2) / len(revised)


def _pair_score(original: str, revised: str, need: float = 0.0) -> float:
    """改稿段与原稿段能否配成一对：达不到门槛的返回 0。
    长段看改稿段的文字有多少来自原稿段；短段（如小标题）看两段整体相似度，免得短标题被长段落“包含”而误配。
    先用相邻两字重合率粗筛，减少逐字比对。"""
    if not revised or not original:
        return 0.0
    if len(revised) < SHORT_PARA_CHARS:
        score = difflib.SequenceMatcher(None, original, revised, autojunk=False).ratio()
        if score >= max(SHORT_PAIR_THRESHOLD, need):
            return score
        # 压缩时一段只留下小标题句：短段的文字几乎都出现在原稿段中，按较低的权重配对，不算“新增”
        if _containment(original, revised) >= SHORT_CONTAIN_THRESHOLD:
            return max(need, CONTAIN_WEIGHT)
        return 0.0
    if need:
        # 认定“移动”“拆分”：改稿段的文字要大部分来自原稿段
        if _coverage(_grams(revised), _grams(original)) < need * 0.8:
            return 0.0
        score = _containment(original, revised)
        return score if score >= need else 0.0
    # 按顺序配对“修改”：改稿段大部分来自原稿段（删改），或原稿段大部分保留在改稿段中（补写），满足其一即可
    rev_grams, ori_grams = _grams(revised), _grams(original)
    rough = max(_coverage(rev_grams, ori_grams), _coverage(ori_grams, rev_grams))
    if rough < PAIR_THRESHOLD * 0.8:
        return 0.0
    blocks = difflib.SequenceMatcher(None, original, revised, autojunk=False).get_matching_blocks()
    matched = sum(block.size for block in blocks if block.size >= 2)
    score = max(matched / len(revised), matched / len(original))
    return score if score >= PAIR_THRESHOLD else 0.0


def _inserted(original: str, revised: str) -> List[str]:
    """改稿段比原稿段多出的文字片段（去掉标点后不少于 2 字的）。
    相邻片段之间只隔一个相同字的，并为一段，免得“的”“和”之类的单字把一句新增内容切碎。"""
    opcodes = difflib.SequenceMatcher(None, original, revised, autojunk=False).get_opcodes()
    spans: List[List[int]] = []
    current: Optional[List[int]] = None
    for tag, _, _, j1, j2 in opcodes:
        if tag in ("insert", "replace"):
            if current is None:
                current = [j1, j2]
            else:
                current[1] = j2
        elif tag == "equal" and current is not None and j2 - j1 >= 2:
            spans.append(current)
            current = None
    if current is not None:
        spans.append(current)
    fragments: List[str] = []
    for j1, j2 in spans:
        piece = revised[j1:j2].strip()
        if len(_norm(piece)) >= 2:
            fragments.append(piece)
    return fragments


def _status_counts(text: str) -> Dict[str, int]:
    text = MARKER_RE.sub("", text)  # 【待核实：甲稿为基本完成……】里的词不算
    return {word: len(pattern.findall(text)) for word, pattern in STATUS_WORDS}


def _status_changes(original: str, revised: str, increases_only: bool = False) -> List[Dict[str, object]]:
    """状态和义务强度词在原稿段、改稿段中的出现次数有变化的，逐词列出。"""
    before, after = _status_counts(original), _status_counts(revised)
    changes: List[Dict[str, object]] = []
    for word, _ in STATUS_WORDS:
        a, b = before[word], after[word]
        if a != b and (b > a or not increases_only):
            changes.append({"word": word, "original": a, "revised": b})
    return changes


def _with_inserted(item: Dict[str, object], original: str, revised: str) -> Dict[str, object]:
    fragments = _inserted(original, revised)
    item["inserted"] = [_excerpt(f) for f in fragments]
    item["inserted_chars"] = sum(len(_norm(f)) for f in fragments)
    # 拆分时一段原稿对应几段改稿，各段的词数自然变少，只报增加的
    item["status_changes"] = _status_changes(original, revised, increases_only=item["type"] == "拆分")
    return item


def _align_block(a_keys: Sequence[str], b_keys: Sequence[str]) -> List[Tuple[int, int, float]]:
    """按顺序配对两组段落，使配对得分之和最大（带权最长公共子序列），返回（原稿下标，改稿下标，得分）。"""
    n, m = len(a_keys), len(b_keys)
    if not n or not m:
        return []
    weight = [[_pair_score(a_keys[i], b_keys[j]) for j in range(m)] for i in range(n)]
    dp = [[0.0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            best = max(dp[i - 1][j], dp[i][j - 1])
            w = weight[i - 1][j - 1]
            if w > 0:
                best = max(best, dp[i - 1][j - 1] + w)
            dp[i][j] = best
    pairs: List[Tuple[int, int, float]] = []
    i, j = n, m
    while i > 0 and j > 0:
        w = weight[i - 1][j - 1]
        if w > 0 and abs(dp[i][j] - (dp[i - 1][j - 1] + w)) < 1e-9:
            pairs.append((i - 1, j - 1, w))
            i, j = i - 1, j - 1
        elif dp[i - 1][j] >= dp[i][j - 1]:
            i -= 1
        else:
            j -= 1
    pairs.reverse()
    return pairs


def _mark_moves_and_splits(changes: List[Dict[str, object]], a: "Body", b: "Body") -> List[Dict[str, object]]:
    """“新增”段落再与全部原稿段比一次：来自某个被删段落的记“移动”，来自已配对段落的记“拆分”。
    处理后仍为“新增”的，才是原稿中找不到来源的段落。"""
    deleted_nos = {int(c["original_no"]) for c in changes if c["type"] == "删除"}
    a_keys = [_norm_heading(p) for p in a.paras]
    consumed: Set[int] = set()
    result: List[Dict[str, object]] = []
    for item in changes:
        if item["type"] != "新增":
            result.append(item)
            continue
        key = _norm_heading(b.paras[int(item["revised_no"]) - 1])
        best_no, best = 0, 0.0
        for index, original in enumerate(a_keys, start=1):
            score = _pair_score(original, key, MOVE_THRESHOLD)
            if score > best:
                best_no, best = index, score
        if not best_no:
            result.append(item)
            continue
        kind = "移动" if best_no in deleted_nos and best_no not in consumed else "拆分"
        if kind == "移动":
            consumed.add(best_no)
        result.append(_with_inserted(
            {"type": kind, "original_no": best_no, "revised_no": item["revised_no"],
             "similarity": round(best, 2), "original": _excerpt(a.paras[best_no - 1]),
             "revised": item["revised"]},
            a.paras[best_no - 1], b.paras[int(item["revised_no"]) - 1]))
    return [c for c in result if not (c["type"] == "删除" and int(c["original_no"]) in consumed)]


def compare_report(original: Path, revised: Path) -> Dict[str, object]:
    a, b = load_body(original), load_body(revised)
    a_secs, b_secs = a.sections(), b.sections()
    used: Set[int] = set()
    section_rows: List[Dict[str, object]] = []
    for sec in a_secs:
        key = _norm_heading(str(sec["heading"]))
        match_index: Optional[int] = None
        best = 0.0
        for j, other in enumerate(b_secs):
            if j in used:
                continue
            ratio = difflib.SequenceMatcher(None, key, _norm_heading(str(other["heading"]))).ratio()
            if ratio > best:
                best, match_index = ratio, j
        if match_index is not None and best >= 0.6:
            used.add(match_index)
            other = b_secs[match_index]
            section_rows.append({"original": sec["heading"], "revised": other["heading"],
                                 "original_chars": sec["char_count"], "revised_chars": other["char_count"]})
        else:
            section_rows.append({"original": sec["heading"], "revised": None,
                                 "original_chars": sec["char_count"], "revised_chars": 0})
    for j, other in enumerate(b_secs):
        if j not in used:
            section_rows.append({"original": None, "revised": other["heading"],
                                 "original_chars": 0, "revised_chars": other["char_count"]})

    a_norm = [_norm(p) for p in a.paras]
    b_norm = [_norm(p) for p in b.paras]
    matcher = difflib.SequenceMatcher(None, a_norm, b_norm, autojunk=False)
    changes: List[Dict[str, object]] = []
    unchanged = 0

    def deleted(i: int) -> Dict[str, object]:
        return {"type": "删除", "original_no": i + 1, "revised_no": None, "similarity": None,
                "original": _excerpt(a.paras[i]), "revised": ""}

    def added(j: int) -> Dict[str, object]:
        return {"type": "新增", "original_no": None, "revised_no": j + 1, "similarity": None,
                "original": "", "revised": _excerpt(b.paras[j])}

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            unchanged += i2 - i1
            continue
        # 不一致的区段内按顺序配对：改稿段的文字大部分来自某个原稿段，才算“修改”；
        # 没配上的改稿段记“新增”，没配上的原稿段记“删除”。标题比较时去掉层次序号。
        a_keys = [_norm_heading(p) for p in a.paras[i1:i2]]
        b_keys = [_norm_heading(p) for p in b.paras[j1:j2]]
        i, j = i1, j1
        for pi, pj, score in _align_block(a_keys, b_keys):
            changes.extend(deleted(k) for k in range(i, i1 + pi))
            changes.extend(added(k) for k in range(j, j1 + pj))
            changes.append(_with_inserted(
                {"type": "修改", "original_no": i1 + pi + 1, "revised_no": j1 + pj + 1,
                 "similarity": round(score, 2),
                 "original": _excerpt(a.paras[i1 + pi]), "revised": _excerpt(b.paras[j1 + pj])},
                a.paras[i1 + pi], b.paras[j1 + pj]))
            i, j = i1 + pi + 1, j1 + pj + 1
        changes.extend(deleted(k) for k in range(i, i2))
        changes.extend(added(k) for k in range(j, j2))
    changes = _mark_moves_and_splits(changes, a, b)
    counts = {"未改": unchanged, "修改": 0, "移动": 0, "拆分": 0, "删除": 0, "新增": 0}
    for item in changes:
        counts[str(item["type"])] += 1
    a_total, b_total = a.char_count, b.char_count
    return {
        "original": a.name,
        "revised": b.name,
        "original_chars": a_total,
        "revised_chars": b_total,
        "ratio": round(b_total / a_total, 3) if a_total else None,
        "original_markers": a.markers(),
        "revised_markers": b.markers(),
        "sections": section_rows,
        "paragraph_summary": counts,
        "inserted_chars": sum(int(c.get("inserted_chars", 0)) for c in changes),
        "status_changed": sum(1 for c in changes if c.get("status_changes")),
        "changes": changes,
        "revisions": {a.name: a.revisions, b.name: b.revisions},
        "basis": BASIS_NOTE,
    }


def render_compare(report: Dict[str, object]) -> str:
    ratio = report["ratio"]
    ratio_text = f"，改稿为原稿的 {float(ratio) * 100:.1f}%" if ratio is not None else ""
    lines = [f"改稿对照：{report['original']} → {report['revised']}",
             f"字数：{report['original_chars']} → {report['revised_chars']}{ratio_text}"]
    om, rm = report["original_markers"], report["revised_markers"]
    lines.append(f"【待补】{om['待补']} → {rm['待补']}；【待核实】{om['待核实']} → {rm['待核实']}")
    lines.append("各部分：")
    for row in report["sections"]:
        if row["original"] is None:
            lines.append(f"  新增部分：{row['revised']}（{row['revised_chars']} 字）")
        elif row["revised"] is None:
            lines.append(f"  删去部分：{row['original']}（{row['original_chars']} 字）")
        else:
            lines.append(f"  {row['original']}：{row['original_chars']} → {row['revised_chars']} 字")
    s = report["paragraph_summary"]
    lines.append(f"段落：未改 {s['未改']}，修改 {s['修改']}，移动 {s['移动']}，拆分 {s['拆分']}，"
                 f"删除 {s['删除']}，新增 {s['新增']}")
    for item in report["changes"][:MAX_LIST_ITEMS]:
        if item["type"] in ("修改", "移动", "拆分"):
            lines.append(f"  [{item['type']}] 原第{item['original_no']}段 → 改稿第{item['revised_no']}段"
                         f"（重合度 {item['similarity']}）：{item['original']} ⇒ {item['revised']}")
            inserted = item.get("inserted") or []
            if inserted:
                shown = "".join(f"「{f}」" for f in inserted[:MAX_INSERTED_SHOWN])
                more = f"等 {len(inserted)} 处" if len(inserted) > MAX_INSERTED_SHOWN else ""
                lines.append(f"      多出的文字（{item['inserted_chars']} 字）：{shown}{more}")
            status = item.get("status_changes") or []
            if status:
                lines.append("      状态/强度词变化：" + "；".join(
                    f"{c['word']} {c['original']}→{c['revised']}" for c in status))
        elif item["type"] == "删除":
            lines.append(f"  [删除] 原第{item['original_no']}段：{item['original']}")
        else:
            lines.append(f"  [新增] 改稿第{item['revised_no']}段：{item['revised']}")
    if len(report["changes"]) > MAX_LIST_ITEMS:
        lines.append(f"  ……另有 {len(report['changes']) - MAX_LIST_ITEMS} 处，用 --json 查看全部")
    for name, revisions in report.get("revisions", {}).items():
        if revisions:
            lines.append("提示：" + _revision_note(str(name), int(revisions)))
    lines.append("提示：逐条确认每处改动都在任务范围内；按批注改稿时，每处改动应能对应到某条批注。"
                 "“新增”段落在原稿中找不到来源，须逐条核对有没有新增事实；“多出的文字”只能是措辞调整、"
                 "新增的【待补】【待核实】或脱敏替换，带新数字、荣誉、评价、排名的删去或改为【待补】。")
    if report.get("status_changed"):
        lines.append(f"提示：{report['status_changed']} 处段落的状态/强度词有变化，逐条对照原文：拟、将、正在写成已，"
                     "基本、力争被删掉，可以、要改成必须之类的，不是本次任务明确要求的一律恢复；"
                     "整句删去造成的减少，写进“删去的数据”或修改说明即可。")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# trace
# ---------------------------------------------------------------------------

def _sentences(body: Body) -> List[Tuple[int, str]]:
    result: List[Tuple[int, str]] = []
    for index, para in enumerate(body.paras, start=1):
        for piece in SENT_SPLIT_RE.split(para):
            if piece.strip():
                result.append((index, piece.strip()))
    return result


def _grams(text: str) -> Set[str]:
    norm = _norm(MARKER_RE.sub("", text))
    if len(norm) < 2:
        return {norm} if norm else set()
    return {norm[i:i + 2] for i in range(len(norm) - 1)}


def _coverage(target: Set[str], *others: Set[str]) -> float:
    if not target:
        return 0.0
    pool: Set[str] = set()
    for other in others:
        pool |= other
    return len(target & pool) / len(target)


def _best_source(grams: Set[str], pool: List[Tuple[str, int, Set[str], str]]) -> Tuple[float, str, str]:
    """返回单句覆盖率最高的来源：（覆盖率，来源标签，来源句原文）。"""
    best: Tuple[float, str, str] = (0.0, "", "")
    for label, _, g, text in pool:
        cov = _coverage(grams, g)
        if cov > best[0]:
            best = (cov, label, text)
    return best


def _uncovered_clauses(sentence: str, pool: List[Tuple[str, int, Set[str], str]],
                       threshold: float) -> List[str]:
    """整句找到了来源，再按逗号切成小句，逐一找来源：在原句后追加的无数字新内容（如荣誉、评价）
    整句覆盖率仍可能过线，要靠小句核对才查得出。去掉标点后不足 4 字的小句不比对。"""
    clauses = [c.strip() for c in CLAUSE_SPLIT_RE.split(sentence) if c.strip()]
    if len(clauses) < 2:
        return []
    result: List[str] = []
    for clause in clauses:
        if len(_norm(MARKER_RE.sub("", clause))) < MIN_CLAUSE_CHARS:
            continue
        if _best_source(_grams(clause), pool)[0] < threshold:
            result.append(_excerpt(clause))
    return result


def trace_report(merged: Path, sources: Sequence[Path], threshold: float = DEFAULT_THRESHOLD) -> Dict[str, object]:
    if not sources:
        raise InputError("至少需要一份来源稿")
    if len(sources) > MAX_SOURCES:
        raise InputError(f"来源稿最多 {MAX_SOURCES} 份")
    if not 0 < threshold <= 1:
        raise InputError("--threshold 必须在 0 到 1 之间")
    m_body = load_body(merged)
    source_bodies = [load_body(p) for p in sources]
    labels = {SOURCE_LABELS[i]: source_bodies[i].name for i in range(len(source_bodies))}

    pool: List[Tuple[str, int, Set[str], str]] = []
    source_sents: List[Tuple[str, int, str, Set[str]]] = []
    for i, body in enumerate(source_bodies):
        label = SOURCE_LABELS[i]
        for para_no, sent in _sentences(body):
            grams = _grams(sent)
            if grams:
                pool.append((label, para_no, grams, sent))
                source_sents.append((label, para_no, sent, grams))

    rows: List[Dict[str, object]] = []
    merged_grams: List[Set[str]] = []
    per_label: Dict[str, int] = {label: 0 for label in labels}
    per_label["多稿合成"] = 0
    unmatched = 0
    partial = 0
    skipped = 0
    for para_no, sent in _sentences(m_body):
        grams = _grams(sent)
        norm_len = len(_norm(MARKER_RE.sub("", sent)))
        if norm_len < MIN_TRACE_CHARS:
            skipped += 1
            continue
        merged_grams.append(grams)
        cov, label, matched = _best_source(grams, pool)
        source = label
        if cov < threshold:
            # 一句由两稿各取一部分拼成：取覆盖最多的两句合并计算
            ranked = sorted(pool, key=lambda item: _coverage(grams, item[2]), reverse=True)[:2]
            if len(ranked) == 2:
                combined = _coverage(grams, ranked[0][2], ranked[1][2])
                if combined >= threshold:
                    cov = combined
                    names = sorted({ranked[0][0], ranked[1][0]})
                    source = "＋".join(names)
                    matched = _excerpt(ranked[0][3]) + "｜" + _excerpt(ranked[1][3])
        uncovered: List[str] = []
        if cov >= threshold:
            if "＋" in source:
                per_label["多稿合成"] += 1
            else:
                per_label[source] += 1
            uncovered = _uncovered_clauses(sent, pool, threshold)
            if uncovered:
                partial += 1
        else:
            source = "未找到来源"
            unmatched += 1
        if matched and "｜" not in matched:
            matched = _excerpt(matched)
        rows.append({"paragraph": para_no, "sentence": _excerpt(sent, 60), "source": source,
                     "coverage": round(cov, 2), "matched": matched, "uncovered_clauses": uncovered})

    unused: List[Dict[str, object]] = []
    for label, para_no, sent, grams in source_sents:
        if len(_norm(sent)) < MIN_TRACE_CHARS:
            continue
        best = max((_coverage(grams, g) for g in merged_grams), default=0.0)
        if best < threshold:
            unused.append({"source": label, "paragraph": para_no, "sentence": _excerpt(sent, 60),
                           "coverage": round(best, 2)})
    return {
        "merged": m_body.name,
        "sources": labels,
        "threshold": threshold,
        "sentences": len(rows),
        "skipped_short": skipped,
        "by_source": per_label,
        "unmatched": unmatched,
        "partial": partial,
        "rows": rows,
        "unused": unused,
        "note": "覆盖率按去掉标点和【待补】【待核实】后的相邻两字计算，只说明文字来自哪一稿；"
                "有来源的句子再按逗号切成小句逐一核对；数字、日期是否被改动，用 fact_diff.py 核对。",
    }


def render_trace(report: Dict[str, object]) -> str:
    lines = [f"合并稿溯源：{report['merged']}（阈值 {report['threshold']}）"]
    for label, name in report["sources"].items():
        lines.append(f"  {label}：{name}")
    by = report["by_source"]
    parts = "，".join(f"{k} {v} 句" for k, v in by.items())
    lines.append(f"共比对 {report['sentences']} 句：{parts}，未找到来源 {report['unmatched']} 句，"
                 f"句内有找不到来源的小句 {report['partial']} 句（过短未比对 {report['skipped_short']} 句）。")
    missing = [r for r in report["rows"] if r["source"] == "未找到来源"]
    if missing:
        lines.append("未找到来源的句子（逐条核对：新增事实须删去或改为【待补】）：")
        for row in missing[:MAX_LIST_ITEMS]:
            lines.append(f"  第{row['paragraph']}段（覆盖率 {row['coverage']}）：{row['sentence']}")
    partial_rows = [r for r in report["rows"] if r.get("uncovered_clauses")]
    if partial_rows:
        lines.append("句内找不到来源的小句（逐条核对：只是改了措辞的保留，带新事实的删去或改为【待补】）：")
        for row in partial_rows[:MAX_LIST_ITEMS]:
            clauses = "".join(f"「{c}」" for c in row["uncovered_clauses"])
            lines.append(f"  第{row['paragraph']}段（整句来自{row['source']}）：{clauses}")
    unused = report["unused"]
    if unused:
        lines.append(f"来源稿中未被采用的句子 {len(unused)} 句（确认是有意舍弃，写进合并说明）：")
        for row in unused[:MAX_LIST_ITEMS]:
            lines.append(f"  {row['source']}第{row['paragraph']}段：{row['sentence']}")
        if len(unused) > MAX_LIST_ITEMS:
            lines.append(f"  ……另有 {len(unused) - MAX_LIST_ITEMS} 句，用 --json 查看全部")
    lines.append(str(report["note"]))
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# comments
# ---------------------------------------------------------------------------

def _text_of(element: ET.Element) -> str:
    return "".join(t.text or "" for t in element.iter(W_T))


def comments_report(path: Path) -> Dict[str, object]:
    _check_path(path)
    if path.suffix.lower() != ".docx":
        raise InputError("comments 只支持 .docx 文件")
    document, comments_root = _open_docx(path)
    body = _docx_body(document)
    paragraphs = _docx_paragraphs(document)
    para_texts: List[str] = [_para_text(p) for p in paragraphs]
    para_index = {id(p): index for index, p in enumerate(paragraphs, start=1)}
    ordinal: Dict[int, int] = {}
    count = 0
    for index, text in enumerate(para_texts, start=1):
        if text.strip():
            count += 1
            ordinal[index] = count
    # 圈定原文按接受全部修订后的文字取：修订删除、移走的文字不计入
    removed = {id(t) for el in body.iter() if el.tag in (W_DEL, W_MOVE_FROM) for t in el.iter(W_T)}

    active: Set[str] = set()
    anchors: Dict[str, List[str]] = {}
    start_para: Dict[str, int] = {}
    ref_para: Dict[str, int] = {}
    p_index = 0
    for element in body.iter():
        tag = element.tag
        if tag == W_P:
            p_index = para_index.get(id(element), p_index)
        elif tag == W_CRS:
            cid = element.get(W_ID, "")
            active.add(cid)
            start_para.setdefault(cid, p_index)
            anchors.setdefault(cid, [])
        elif tag == W_CRE:
            active.discard(element.get(W_ID, ""))
        elif tag == W_T and active and id(element) not in removed:
            for cid in active:
                anchors[cid].append(element.text or "")
        elif tag == W_CREF:
            ref_para.setdefault(element.get(W_ID, ""), p_index)
    revisions = _count_revisions(document)

    items: List[Dict[str, object]] = []
    if comments_root is not None:
        for comment in comments_root.iter(W_COMMENT):
            cid = comment.get(W_ID, "")
            text = " ".join(_text_of(p) for p in comment.iter(W_P) if _text_of(p).strip()).strip()
            para_no = start_para.get(cid) or ref_para.get(cid) or 0
            anchor = "".join(anchors.get(cid, [])).strip()
            if not anchor and para_no:
                anchor = para_texts[para_no - 1].strip()
            date = (comment.get(W_DATE) or "")[:10]
            items.append({"id": cid, "paragraph": ordinal.get(para_no), "anchor": _excerpt(anchor, 80),
                          "comment": text, "date": date})
    items.sort(key=lambda item: (item["paragraph"] is None, item["paragraph"] or 0))
    for number, item in enumerate(items, start=1):
        item["no"] = number
    note = "段落序号按非空段落计算（含表格内段落）；不输出批注人姓名。"
    if revisions:
        note += (f"文档含修订痕迹 {revisions} 处：本 skill 按“接受全部修订”后的文字改稿，"
                 f"先运行 body 子命令把正文另存为 .txt（如 body {path.name} --out 原稿.txt），"
                 f"之后的统计和比对都用这份 .txt；不接受某处修订的，请用户先在 Word 中处理。"
                 f"学习领导的修订习惯请用 gongwen-profile。")
    return {
        "file": path.name,
        "comment_count": len(items),
        "revision_marks": revisions,
        "comments": items,
        "note": note,
    }


def render_comments(report: Dict[str, object]) -> str:
    lines = [f"批注提取：{report['file']}，共 {report['comment_count']} 条批注；修订痕迹 {report['revision_marks']} 处。"]
    for item in report["comments"]:
        where = f"第{item['paragraph']}段" if item["paragraph"] else "位置未定"
        date = f"（{item['date']}）" if item["date"] else ""
        lines.append(f"{item['no']}. {where}{date}")
        lines.append(f"   原文：{item['anchor'] or '（未圈定文字）'}")
        lines.append(f"   批注：{item['comment'] or '（空批注）'}")
    if not report["comments"]:
        lines.append("没有找到批注。手写批注或截图批注，请把批注原话和位置转述给模型。")
    lines.append(str(report["note"]))
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# body
# ---------------------------------------------------------------------------

def body_text(paths: Sequence[Path]) -> Tuple[str, Dict[str, object]]:
    """每份文件取【起草说明】之前的全部文字（含标题、版头，去掉首尾空行），依次拼接，文件之间空一行。"""
    if not paths:
        raise InputError("至少需要一份文件")
    blocks: List[str] = []
    files: List[Dict[str, object]] = []
    for path in paths:
        lines, revisions = read_source(path)
        kept: List[str] = []
        cut = False
        for line in lines:
            if DRAFT_NOTE_RE.match(line):
                cut = True
                break
            kept.append(line.rstrip())
        while kept and not kept[0].strip():
            kept.pop(0)
        while kept and not kept[-1].strip():
            kept.pop()
        blocks.append("\n".join(kept))
        files.append({"file": path.name, "lines": len(kept), "draft_note_removed": cut,
                      "revisions": revisions})
    return "\n\n".join(b for b in blocks if b) + "\n", {"files": files}


def _free_path(path: Path) -> Path:
    """同名文件已存在时依次加 -v2、-v3……，不覆盖。"""
    if not path.exists():
        return path
    for number in range(2, 1000):
        candidate = path.with_name(f"{path.stem}-v{number}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise InputError(f"同名文件过多，请换一个文件名：{path.name}")


def write_body(paths: Sequence[Path], out: Path) -> Dict[str, object]:
    text, info = body_text(paths)
    if out.suffix.lower() not in (".txt", ".md"):
        raise InputError("--out 只能是 .txt 或 .md 文件")
    if not out.parent.exists():
        raise InputError(f"输出目录不存在：{out.parent}")
    target = _free_path(out)
    try:
        with open(target, "x", encoding="utf-8", newline="\n") as handle:  # "x"：文件已存在时报错，不覆盖
            handle.write(text)
    except OSError as exc:
        raise InputError(f"写入失败：{target}（{exc}）") from exc
    info["out"] = str(target)
    return info


def render_body_info(info: Dict[str, object]) -> str:
    lines = [f"已写入：{info['out']}（UTF-8，不覆盖已有文件）"]
    for item in info["files"]:  # type: ignore[union-attr]
        extra = []
        if item["draft_note_removed"]:
            extra.append("已去掉【起草说明】及其后内容")
        if item["revisions"]:
            extra.append(f"含修订痕迹 {item['revisions']} 处，已按接受全部修订后的文字导出")
        tail = "；" + "；".join(extra) if extra else ""
        lines.append(f"  {item['file']}：{item['lines']} 行{tail}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 命令行
# ---------------------------------------------------------------------------

def _setup_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
        except (AttributeError, ValueError, OSError):
            pass


class _ChineseHelpFormatter(argparse.RawDescriptionHelpFormatter):
    def add_usage(self, usage, actions, groups, prefix=None):
        return super().add_usage(usage, actions, groups, "用法：" if prefix is None else prefix)


def _new_parser(prog: str, description: str, epilog: str = "") -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog=prog, description=description, epilog=epilog,
                                     formatter_class=_ChineseHelpFormatter, add_help=False)
    parser._optionals.title = "选项"
    parser._positionals.title = "参数"
    parser.add_argument("-h", "--help", action="help", help="显示本帮助信息并退出")
    return parser


def make_parser() -> argparse.ArgumentParser:
    parser = _new_parser(
        "polish_stat.py",
        "改稿润色的辅助统计：篇幅统计（count）、改稿前后对照（compare）、合并稿溯源（trace）、"
        "提取 Word 批注（comments）、导出正文（body）。只读输入文件；除 body --out 外不写文件。",
        "退出码：0 正常；1 使用 --strict 且 count 超出目标（或低于±10%的下限）、trace 有找不到来源的句子或小句；"
        "2 输入错误。",
    )
    sub = parser.add_subparsers(dest="command", title="子命令", metavar="子命令")

    def add_sub(name: str, help_text: str) -> argparse.ArgumentParser:
        sp = sub.add_parser(name, help=help_text, description=help_text, add_help=False,
                            formatter_class=_ChineseHelpFormatter)
        sp._optionals.title = "选项"
        sp._positionals.title = "参数"
        sp.add_argument("-h", "--help", action="help", help="显示本帮助信息并退出")
        return sp

    p_count = add_sub("count", "统计正文字数和各部分字数，可与目标字数比较")
    p_count.add_argument("file", help="稿件（.txt、.md、.docx）")
    p_count.add_argument("--target", type=int, help="目标字数；默认允许±10%%，加 --max 时为上限")
    p_count.add_argument("--max", action="store_true", dest="upper",
                         help="目标是上限（用户说“以内”“不超过”）：超过即报超出，低于上限七成时提示复查必保内容")
    p_count.add_argument("--include-title", action="store_true", help="把标题也计入字数")
    p_count.add_argument("--json", action="store_true", help="以 JSON 格式输出")
    p_count.add_argument("--strict", action="store_true", help="超出目标（或低于±10%%的下限）时返回 1")

    p_cmp = add_sub("compare", "对照原稿和改稿：字数变化、逐段改动清单和每段多出的文字")
    p_cmp.add_argument("original", help="原稿")
    p_cmp.add_argument("revised", help="改稿")
    p_cmp.add_argument("--json", action="store_true", help="以 JSON 格式输出")

    p_trace = add_sub("trace", "合并稿逐句溯源：每句来自哪一稿，哪些句子、小句找不到来源")
    p_trace.add_argument("merged", help="合并稿")
    p_trace.add_argument("sources", nargs="+", help="来源稿，按甲稿、乙稿……的顺序，最多 6 份")
    p_trace.add_argument("--threshold", type=float, default=DEFAULT_THRESHOLD,
                         help=f"判定有来源的最低覆盖率，默认 {DEFAULT_THRESHOLD}")
    p_trace.add_argument("--json", action="store_true", help="以 JSON 格式输出")
    p_trace.add_argument("--strict", action="store_true", help="有找不到来源的句子或小句时返回 1")

    p_com = add_sub("comments", "从 .docx 中提取 Word 批注，统计修订痕迹")
    p_com.add_argument("file", help="带批注的 .docx")
    p_com.add_argument("--json", action="store_true", help="以 JSON 格式输出")

    p_body = add_sub("body", "导出【起草说明】之前的全部文字；给几份文件时依次拼接（docx 按接受全部修订后的文字）")
    p_body.add_argument("files", nargs="+", help="稿件或材料（.txt、.md、.docx），按拼接顺序")
    p_body.add_argument("--out", help="写入的 .txt 文件（UTF-8）；同名文件已存在时加 -v2，不覆盖；不给时打印到屏幕")
    p_body.add_argument("--json", action="store_true", help="以 JSON 格式输出（与 --out 同用时输出写入结果）")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    _setup_console()
    parser = make_parser()
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:
        return 0 if exc.code == 0 else 2
    if not args.command:
        parser.print_help()
        return 2
    try:
        if args.command == "count":
            report = count_report(Path(args.file), args.target, args.include_title, args.upper)
            text = render_count(report)
            failed = bool(args.strict and report["target"] and report["target"]["status"] in FAIL_STATUSES)
        elif args.command == "compare":
            report = compare_report(Path(args.original), Path(args.revised))
            text, failed = render_compare(report), False
        elif args.command == "trace":
            report = trace_report(Path(args.merged), [Path(p) for p in args.sources], args.threshold)
            text = render_trace(report)
            failed = bool(args.strict and (report["unmatched"] or report["partial"]))
        elif args.command == "body":
            paths = [Path(p) for p in args.files]
            if args.out:
                report = write_body(paths, Path(args.out))
                text = render_body_info(report)
            else:
                body, report = body_text(paths)
                report["text"] = body
                text = body.rstrip()
            failed = False
        else:
            report = comments_report(Path(args.file))
            text, failed = render_comments(report), False
    except InputError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(text)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
