# -*- coding: utf-8 -*-
"""本单位历史稿文风参数统计（gongwen-profile）。

用法：
    python scripts/style_profile.py <文件或目录> [<文件或目录>...] [--json] [--top N]
                                    [--min-count N] [--examples] [--output 文件]

输入：.txt、.md、.docx。目录只读取其下一层的 .txt、.md、.docx，按文件名排序。
      .docx 用标准库解析，不需要 python-docx；带修订痕迹的按"接受全部修订"后的文本统计。
      每篇只统计【起草说明】之前的正文。建议 3～10 篇，少于 3 篇或多于 10 篇时给出提示，照常统计。

统计口径：
    字数：不含空白的字符数（含标点）。
    句子：按句号、问号、感叹号切分；段末没有句末标点的部分算一句。句长不含句末标点。
    段落：一行为一段。标题、主送机关行、落款、日期行、"附件："行、表格不计入段落和句子统计。
    标题：一级"一、"，二级"（一）"，三级"1."，四级"（1）"；与正文同段的标题取到第一个句号为止。
    密度：每千字出现次数，分母为该篇正文字数。
    汇总：中位数和四分位范围（P25～P75），不用平均数，避免个别长稿拉偏。

保密与隐私：
    发现密级标识，"内部资料""工作秘密""严禁外传""内部参阅"等标注，或"未公开""未披露"说明的，该篇不统计，
    提示改用内网模型或删除后再统计；docx 同时查页眉、页脚和修订中删除的文字（规则见同目录 secret_marks.py，
    优先取 lint_rules.json 中 PRI-005、PRI-006、PRI-007，与共享机检一致）。
    发现身份证号、手机号、带标签的个人编号、家庭住址、"姓名＋性别＋年龄"组合的，照常统计，
    只报告每篇的命中类别和次数，不输出命中内容；这些片段在提取高频短语前已屏蔽。

输出：不输出原文段落和句子。高频四字格、惯用语、领起词只输出短语本身；
      标题示例只在使用 --examples 时输出，供模型归纳句式，不写入文风参数文件。

退出码：0 正常；1 没有可统计的文件（全部被跳过或为空）；2 输入错误。
只依赖标准库。
"""
from __future__ import annotations

import argparse
import json
import math
import os
import re
import statistics
import sys
import zipfile
from collections import Counter, OrderedDict, defaultdict
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET

_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))
import secret_marks  # noqa: E402  （同目录的共用模块）

VERSION = "1.1.0"
BOUNDARY = "【起草说明】"
# 与 gongwen_lint.py、fact_diff.py 的分界行写法保持一致
DRAFT_NOTE_RE = re.compile(r"^\s*(?:#{1,6}\s*)?(?:(?:\*\*|__)\s*)?" + re.escape(BOUNDARY))
SUPPORTED_SUFFIXES = (".txt", ".md", ".docx")
RECOMMENDED_MIN, RECOMMENDED_MAX = 3, 10

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS

HAN = r"\u4e00-\u9fff"
HAN_RE = re.compile(r"[\u4e00-\u9fff]")
WS_RE = re.compile(r"\s+")

# 层次标题
HEADING_RES = OrderedDict([
    ("1", re.compile(r"^[一二三四五六七八九十]{1,3}、")),
    ("2", re.compile(r"^[（(][一二三四五六七八九十]{1,3}[）)]")),
    ("3", re.compile(r"^\d{1,2}[.．](?!\d)")),
    ("4", re.compile(r"^[（(]\d{1,2}[）)]")),
])
SENT_END = "。！？"
SENT_SPLIT_RE = re.compile(r"[^。！？]+[。！？]?")

DATE_LINE_RE = re.compile(r"^\d{4}年\d{1,2}月\d{1,2}日$|^[二〇一二三四五六七八九十○]{4}年[一二三四五六七八九十]{1,3}月[一二三四五六七八九十]{1,3}日$")
ATTACH_RE = re.compile(r"^附件[：:\d]")
MAIN_SEND_RE = re.compile(r"^[^。！？，]{2,40}[：:]$")
TABLE_LINE_RE = re.compile(r"^\s*\|.*\|\s*$")

# 一是、二是……：前面是行首、标点或空白
YISHI_RE = re.compile(r"(?:^|(?<=[，。；：！？、\s（(]))([一二三四五六七八九十])是")
DASH_RE = re.compile(r"[—―]{1,2}")
NUMBER_RE = re.compile(r"\d+(?:[.．]\d+)?")
PERCENT_RE = re.compile(r"\d+(?:[.．]\d+)?\s*[%％]")

PUNCT_NAMES = OrderedDict([
    ("顿号", "、"),
    ("逗号", "，"),
    ("分号", "；"),
    ("冒号", "："),
    ("问号", "？"),
    ("感叹号", "！"),
    ("左引号", "“"),
    ("书名号", "《"),
    ("左括号", "（"),
])

# 措辞层级（见 references/_shared/style/defensive.md 第六部分）
WORDING = OrderedDict([
    ("强制词", re.compile(r"必须|严禁|不得|一律|务必")),
    ("规范词", re.compile(r"应当")),
    ("授权词", re.compile(r"(?<!不)可以(?!考虑)")),
    ("建议词", re.compile(r"建议|可以考虑|尽量|酌情")),
])

# 机关常用语：多了才是问题，统计密度
JIGUAN_WORDS = (
    "进一步", "切实", "扎实", "不断", "积极", "全面", "深入", "有效", "着力", "大力",
    "高度重视", "抓手", "压实", "落实落细", "走深走实", "见行见效", "赋能", "助力",
    "打造", "闭环", "统筹", "持续", "显著", "明显",
)

# 惯用语（开头、结尾、过渡），只统计出现与否和次数
SET_PHRASES = OrderedDict([
    ("现将……如下", re.compile(r"现将[^。]{0,30}如下")),
    ("特此报告", re.compile(r"特此报告")),
    ("特此通知", re.compile(r"特此通知")),
    ("特此函告", re.compile(r"特此函告|特此函复")),
    ("妥否，请批示", re.compile(r"妥否[，,]\s*请批示")),
    ("请予审批", re.compile(r"请予(?:审批|批准)")),
    ("以上汇报……批评指正", re.compile(r"以上汇报[^。]{0,20}批评指正")),
    ("在……领导下", re.compile(r"在[^。，]{1,24}(?:正确|坚强|亲切)?领导下")),
    ("按照……部署（要求）", re.compile(r"按照[^。]{1,30}(?:部署|要求)")),
    ("根据……要求（精神）", re.compile(r"根据[^。]{1,30}(?:要求|精神|规定)")),
    ("为深入贯彻落实", re.compile(r"为(?:深入)?贯彻落实")),
    ("下面，我讲", re.compile(r"下面[，,]?\s*我(?:讲|谈|就)")),
    ("总的来看", re.compile(r"总的来看|总体来看")),
    ("下一步", re.compile(r"下一步")),
    ("截至……", re.compile(r"截至")),
    ("据统计", re.compile(r"据统计")),
    ("同比", re.compile(r"同比")),
    ("百分点", re.compile(r"百分点")),
])

# 开头方式
OPENING_RULES = (
    ("铺垫式", re.compile(r"^(?:随着|近年来|当前[，,]|在[^。，]{1,24}领导下|为(?:深入|进一步)?(?:贯彻|落实)|"
                       r"[^。]{0,10}高度重视)")),
    ("依据式", re.compile(r"^(?:根据|按照|依据|为[^。]{1,30}(?:根据|按照))")),
    ("引出式", re.compile(r"现将[^。]{0,30}如下")),
)
# 开头 40 字内出现带单位的数据，视为数据式开头（日期不算）
DATA_HEAD_RE = re.compile(r"[0-9][0-9.．]*\s*(?:%|％|万|亿|元|个|项|件|户|人|次|家|条|处|台|天|篇|期|所|名)")
# 首句中出现即视为铺垫式开头
PREAMBLE_IN_SENTENCE_RE = re.compile(r"在[^。，]{1,24}(?:正确|坚强|亲切)领导下|随着[^。]{2,30}[，,]|近年来[，,]|"
                                     r"高度重视|重要意义|新形势下|新时代以来")
# 结构性标题，不计入领起词
STRUCTURAL_HEADING_RE = re.compile(r"^(?:存在的?(?:主要)?问题|下一步|主要(?:做法|工作|成效)|基本情况|工作打算|"
                                   r"\d{4}年工作(?:打算|安排)|下半年工作|重点任务完成情况)")
ENDING_RULES = (
    ("公文结束语", re.compile(r"特此(?:报告|通知|函告|函复|通报|批复)|妥否[，,]\s*请批示|请予(?:审批|批准)|此复")),
    ("请批评指正", re.compile(r"批评指正")),
    ("号召式", re.compile(r"让我们|携手|共同谱写|再创佳绩|新的更大")),
    ("表态式", re.compile(r"^(?:我们将|我[局院校委办镇区县]将|下一步|今后)")),
    ("分条打算", re.compile(r"^一是")),
)

# 常见姓氏（用于识别"姓名＋性别＋年龄""××同志"），只计数，不输出内容
SURNAMES = ("[王李张刘陈杨黄赵吴周徐孙马朱胡郭何高林罗郑梁谢宋唐许韩冯邓曹彭曾肖田董袁潘于蒋蔡余杜叶程苏魏吕"
            "丁任沈姚卢姜崔钟谭陆汪范金石廖贾夏韦付方白邹孟熊秦邱江尹薛闫段雷侯龙史陶黎贺顾毛郝龚邵万钱严覃武戴莫孔向汤]")

# 隐私：只计数，不输出内容
PRIVACY_RES = OrderedDict([
    ("身份证号", re.compile(r"(?<![0-9A-Za-z])\d{17}[\dXx](?![0-9A-Za-z])")),
    ("手机号", re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")),
    ("个人编号", re.compile(r"(?:病案号|住院号|门诊号|学号|学籍号|考生号|准考证号|医保卡号|社保卡号|护照号|工号)"
                         r"\s*[：:为是]?\s*[A-Za-z0-9]{4,}")),
    ("家庭住址", re.compile(r"(?:家庭住址|现住址|住址|户籍地址)\s*[：:为是]?\s*[^，。；\s]{0,20}\d+\s*(?:号|室|栋|幢|单元|楼)")),
    ("姓名＋性别＋年龄", re.compile(SURNAMES + r"[\u4e00-\u9fff]{1,2}(?<!某)[，,（(]\s*[男女][，,、]\s*\d{1,3}\s*岁")),
    ("疑似姓名（××同志）", re.compile(SURNAMES + r"(?!某)[\u4e00-\u9fff]{1,2}(?<!某)同志")),
])
MASK_CHAR = "\u3000"  # 屏蔽隐私片段用的全角空格：不是汉字，不会进入短语统计


class InputError(Exception):
    """输入错误：文件不存在、类型不支持、无法解析。"""


# ---------------------------------------------------------------------------
# 读取
# ---------------------------------------------------------------------------

def _decode(data: bytes, name: str, notes: List[str]) -> str:
    for encoding in ("utf-8-sig", "gb18030"):
        try:
            text = data.decode(encoding)
        except UnicodeDecodeError:
            continue
        if encoding != "utf-8-sig":
            notes.append(f"{name} 不是 UTF-8 编码，已按 GB18030 读取")
        return text
    raise InputError(f"无法识别文件编码，请另存为 UTF-8：{name}")


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _paragraph_text(p: ET.Element) -> str:
    """段落文字，按"接受全部修订"处理：跳过删除和移出的内容。"""
    parts: List[str] = []

    def walk(node: ET.Element, hidden: bool) -> None:
        name = _local(node.tag)
        if name in ("del", "moveFrom"):
            hidden = True
        if not hidden:
            if name == "t" and node.text:
                parts.append(node.text)
            elif name == "tab":
                parts.append("\t")
            elif name in ("br", "cr"):
                parts.append("\n")
        for child in node:
            if _local(child.tag) in ("pPr", "rPr", "instrText", "delInstrText"):
                continue
            walk(child, hidden)

    walk(p, False)
    return "".join(parts)


def read_docx_lines(path: Path) -> List[str]:
    """用标准库读取 docx 正文：段落一行；表格每行一行，单元格用制表符连接，行首加制表符以便识别。"""
    try:
        with zipfile.ZipFile(str(path)) as zf:
            data = zf.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError, OSError) as exc:
        raise InputError(f"无法打开 docx 文件：{path.name}（{exc}）")
    try:
        root = ET.fromstring(data)
    except ET.ParseError as exc:
        raise InputError(f"docx 内容无法解析：{path.name}（{exc}）")
    body = root.find(W + "body")
    lines: List[str] = []
    if body is None:
        return lines

    def visit(node: ET.Element) -> None:
        for child in node:
            name = _local(child.tag)
            if name == "p":
                lines.extend(_paragraph_text(child).split("\n"))
            elif name == "tbl":
                for tr in child.iter(W + "tr"):
                    cells = []
                    for tc in tr.findall(W + "tc"):
                        cells.append("".join(_paragraph_text(p) for p in tc.iter(W + "p")))
                    lines.append("\t" + "\t".join(cells))
            elif name in ("sdt", "sdtContent", "customXml", "ins", "smartTag"):
                visit(child)

    visit(body)
    return lines


def read_lines(path: Path, notes: List[str]) -> List[str]:
    if not path.exists():
        raise InputError(f"文件不存在：{path}")
    if not path.is_file():
        raise InputError(f"不是文件：{path}")
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise InputError(f"不支持的文件类型 {suffix or '（无扩展名）'}，只支持 .txt、.md、.docx：{path.name}")
    if suffix == ".docx":
        lines = read_docx_lines(path)
    else:
        try:
            data = path.read_bytes()
        except OSError as exc:
            raise InputError(f"读取失败：{path.name}（{exc}）")
        text = _decode(data, path.name, notes)
        lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    body: List[str] = []
    for line in lines:
        if DRAFT_NOTE_RE.match(line):
            break
        body.append(line)
    return body


def expand_inputs(items: Sequence[str]) -> List[Path]:
    paths: List[Path] = []
    for raw in items:
        path = Path(raw)
        if path.is_dir():
            found = sorted(p for p in path.iterdir()
                           if p.is_file() and p.suffix.lower() in SUPPORTED_SUFFIXES
                           and not p.name.startswith(("~$", ".")))
            if not found:
                raise InputError(f"目录中没有 .txt、.md、.docx 文件：{path}")
            paths.extend(found)
        else:
            paths.append(path)
    seen = set()
    unique: List[Path] = []
    for p in paths:
        key = os.path.normcase(str(p.resolve())) if p.exists() else str(p)
        if key not in seen:
            seen.add(key)
            unique.append(p)
    return unique


# ---------------------------------------------------------------------------
# 结构切分
# ---------------------------------------------------------------------------

def nchars(text: str) -> int:
    return len(WS_RE.sub("", text))


def heading_level(line: str) -> Optional[str]:
    for level, regex in HEADING_RES.items():
        if regex.match(line):
            return level
    return None


def strip_heading_marker(line: str, level: str) -> str:
    return HEADING_RES[level].sub("", line, count=1).strip()


class Doc:
    """一篇稿件切分后的结构。"""

    def __init__(self, name: str) -> None:
        self.name = name
        self.title: Optional[str] = None
        self.subtitle = False
        self.headings: List[Tuple[str, str]] = []   # (级别, 标题文字)
        self.paragraphs: List[str] = []             # 正文段落（不含标题行）
        self.text_for_density = ""                  # 标题与正文段落，用于密度和字数
        self.skipped_lines = 0


def split_structure(name: str, raw_lines: Sequence[str]) -> Doc:
    doc = Doc(name)
    lines = [line.strip() for line in raw_lines]
    # 去掉 Markdown 标题符号
    lines = [re.sub(r"^#{1,6}\s*", "", line) for line in lines]
    nonempty = [(i, line) for i, line in enumerate(lines) if line]
    if not nonempty:
        return doc
    first_idx = 0
    first = nonempty[0][1]
    if (heading_level(first) is None and len(first) <= 60 and not any(ch in first for ch in SENT_END)
            and len(nonempty) > 1):
        doc.title = first
        first_idx = 1
        if len(nonempty) > 1 and re.match(r"^[—―]{1,2}", nonempty[1][1]):
            doc.subtitle = True
            first_idx = 2
    # 标题之后、正文之前的短行（如讲话人、日期括注）不计入正文
    while (first_idx < len(nonempty) - 1 and len(nonempty[first_idx][1]) <= 20
           and heading_level(nonempty[first_idx][1]) is None
           and not re.search(r"[，。；！？、]", nonempty[first_idx][1])
           and not MAIN_SEND_RE.match(nonempty[first_idx][1])):
        first_idx += 1
    density_parts: List[str] = []
    body_started = False
    tail_meta = _tail_meta_indexes(nonempty)
    for pos, (_i, line) in enumerate(nonempty[first_idx:], start=first_idx):
        if TABLE_LINE_RE.match(line) or raw_lines[_i].startswith("\t"):
            doc.skipped_lines += 1
            continue
        if not body_started and MAIN_SEND_RE.match(line) and heading_level(line) is None:
            doc.skipped_lines += 1  # 主送机关行
            body_started = True
            continue
        body_started = True
        if pos in tail_meta or DATE_LINE_RE.match(line) or ATTACH_RE.match(line):
            doc.skipped_lines += 1
            continue
        level = heading_level(line)
        if level:
            content = strip_heading_marker(line, level)
            cut = _first_sentence_end(content)
            if cut is not None and cut < len(content) - 1:
                # 标题与正文同段：标题取到第一个句号
                head, rest = content[:cut], content[cut + 1:].strip()
                doc.headings.append((level, head))
                density_parts.append(head)
                if rest:
                    doc.paragraphs.append(rest)
                    density_parts.append(rest)
            else:
                ends_like_sentence = cut is not None and (nchars(content) > 15 or "，" in content)
                if len(content) > 80 or ends_like_sentence:
                    # 带句末标点的单句条目、很长的条目，一般是正文，按段落统计
                    doc.paragraphs.append(content)
                else:
                    doc.headings.append((level, content.rstrip("。")))
                density_parts.append(content)
            continue
        doc.paragraphs.append(line)
        density_parts.append(line)
    doc.text_for_density = "\n".join(density_parts)
    return doc


def _first_sentence_end(text: str) -> Optional[int]:
    for i, ch in enumerate(text):
        if ch in SENT_END:
            return i
    return None


def _tail_meta_indexes(nonempty: Sequence[Tuple[int, str]]) -> set:
    """文末落款：最后几行中不含标点的短行（单位名称）和日期行。"""
    result = set()
    for pos in range(len(nonempty) - 1, max(-1, len(nonempty) - 4), -1):
        line = nonempty[pos][1]
        if ENDING_RULES[0][1].search(line):
            break  # "特此通知"等结束语属于正文
        if DATE_LINE_RE.match(line) or (len(line) <= 24 and not re.search(r"[，。；：！？、]", line)
                                         and heading_level(line) is None):
            result.add(pos)
        else:
            break
    return result


def split_sentences(paragraph: str) -> List[str]:
    sentences = []
    for match in SENT_SPLIT_RE.finditer(paragraph):
        piece = match.group(0)
        core = piece.rstrip(SENT_END)
        if nchars(core) > 0:
            sentences.append(core)
    return sentences


# ---------------------------------------------------------------------------
# 统计
# ---------------------------------------------------------------------------

def quantiles(values: Sequence[float]) -> Dict[str, float]:
    vals = sorted(values)
    if not vals:
        return {}
    if len(vals) == 1:
        v = vals[0]
        return {"min": v, "p25": v, "median": v, "p75": v, "max": v, "mean": v, "n": 1}
    q = statistics.quantiles(vals, n=4, method="inclusive")
    return {
        "min": round(vals[0], 1),
        "p25": round(q[0], 1),
        "median": round(statistics.median(vals), 1),
        "p75": round(q[2], 1),
        "max": round(vals[-1], 1),
        "mean": round(statistics.mean(vals), 1),
        "n": len(vals),
    }


def per_k(count: int, chars: int) -> float:
    return round(count * 1000.0 / chars, 2) if chars else 0.0


def heading_style(text: str) -> str:
    parts = [p for p in re.split(r"[，,\s　]+", text) if p]
    if len(parts) == 2 and len(parts[0]) == len(parts[1]) and len(parts[0]) >= 3:
        return "对仗式"
    if len(parts) >= 2:
        return "分句式"
    return "短语式（8字以内）" if nchars(text) <= 8 else "长短语式"


NUM_SUMMARY_RE = re.compile(r"[一二三四五六七八九十两]+[个项大条点化抓]|\d+[个项条]")


def privacy_scan(lines: Sequence[str], extra_lines: Sequence[str] = ()) -> Tuple[Dict[str, int], List[str]]:
    """返回（个人信息类别和次数，命中的密级或内部标注类别）。extra_lines 只用于查标注（docx 的页眉页脚等）。"""
    counts: Dict[str, int] = {}
    marks = secret_marks.find_marks(list(lines) + list(extra_lines))
    for line in lines:
        for label, regex in PRIVACY_RES.items():
            n = len(regex.findall(line))
            if n:
                counts[label] = counts.get(label, 0) + n
    return counts, marks


def mask_privacy(text: str) -> str:
    for regex in PRIVACY_RES.values():
        text = regex.sub(lambda m: MASK_CHAR * len(m.group(0)), text)
    return text


def classify_opening(paragraph: str) -> str:
    first = split_sentences(paragraph)
    head_sentence = first[0] if first else paragraph
    # 顺序：铺垫式、依据式、数据式、引出式（只有"现将……如下"一类引出语）、直述式
    if PREAMBLE_IN_SENTENCE_RE.search(head_sentence):
        return "铺垫式"
    rules = dict(OPENING_RULES)
    if rules["铺垫式"].search(paragraph):
        return "铺垫式"
    if rules["依据式"].search(paragraph):
        return "依据式"
    if DATA_HEAD_RE.search(paragraph[:40]):
        return "数据式"
    if rules["引出式"].search(paragraph):
        return "引出式"
    return "直述式"


def classify_ending(paragraph: str) -> str:
    for label, regex in ENDING_RULES:
        if regex.search(paragraph):
            return label
    return "其他"


def analyze_doc(doc: Doc) -> dict:
    text = doc.text_for_density
    chars = nchars(text)
    sentences: List[str] = []
    for para in doc.paragraphs:
        sentences.extend(split_sentences(para))
    sent_lens = [nchars(s) for s in sentences]
    para_lens = [nchars(p) for p in doc.paragraphs]
    levels = Counter(level for level, _ in doc.headings)
    yishi = YISHI_RE.findall(text)
    max_enum = 0
    for ch in yishi:
        max_enum = max(max_enum, "一二三四五六七八九十".index(ch) + 1)
    punct = OrderedDict((name, per_k(text.count(sym), chars)) for name, sym in PUNCT_NAMES.items())
    punct["破折号"] = per_k(len(DASH_RE.findall(text)), chars)
    wording = OrderedDict((name, per_k(len(regex.findall(text)), chars)) for name, regex in WORDING.items())
    return {
        "chars": chars,
        "title": bool(doc.title),
        "subtitle": doc.subtitle,
        "paragraphs": len(doc.paragraphs),
        "sentences": len(sentences),
        "sentence_len": quantiles(sent_lens),
        "short_sentence_ratio": round(sum(1 for n in sent_lens if n <= 20) / len(sent_lens), 3) if sent_lens else 0.0,
        "long_sentence_ratio": round(sum(1 for n in sent_lens if n >= 80) / len(sent_lens), 3) if sent_lens else 0.0,
        "paragraph_len": quantiles(para_lens),
        "headings": {lv: levels.get(lv, 0) for lv in HEADING_RES},
        "yishi": {"count": len(yishi), "per_1000": per_k(len(yishi), chars), "max_enum": max_enum},
        "punctuation_per_1000": punct,
        "wording_per_1000": wording,
        "numbers_per_1000": per_k(len(NUMBER_RE.findall(text)), chars),
        "percent_count": len(PERCENT_RE.findall(text)),
        "opening": classify_opening(doc.paragraphs[0]) if doc.paragraphs else "无正文",
        "ending": classify_ending(doc.paragraphs[-1]) if doc.paragraphs else "无正文",
    }


def four_char_phrases(texts: Sequence[str], min_count: int, top: int) -> List[dict]:
    """高频四字格：出现次数够、跨篇出现、左右邻字多样（排除长短语被截断的片段）。"""
    counts: Counter = Counter()
    docs_of: Dict[str, set] = defaultdict(set)
    left: Dict[str, Counter] = defaultdict(Counter)
    right: Dict[str, Counter] = defaultdict(Counter)
    for di, text in enumerate(texts):
        for line in text.split("\n"):
            for run in re.finditer(r"[\u4e00-\u9fff]{4,}", line):
                seg, start = run.group(0), run.start()
                after_digit = start > 0 and (line[start - 1].isdigit() or line[start - 1] in ".．%％")
                for k in range(len(seg) - 3):
                    gram = seg[k:k + 4]
                    if k == 0 and after_digit:
                        continue  # "3.1个工作日"中的"个工作日"是计量单位，不是四字格
                    if "的" in gram or "了" in gram or re.match(r"^[一二三四五六七八九十]是", gram):
                        continue
                    counts[gram] += 1
                    docs_of[gram].add(di)
                    left[gram][seg[k - 1] if k > 0 else "|"] += 1
                    right[gram][seg[k + 4] if k + 4 < len(seg) else "|"] += 1
    need_docs = 2 if len(texts) >= 2 else 1
    result = []
    for gram, n in counts.items():
        if n < min_count or len(docs_of[gram]) < need_docs:
            continue
        if not (_side_ok(left[gram], n) and _side_ok(right[gram], n)):
            continue
        result.append({"phrase": gram, "count": n, "docs": len(docs_of[gram])})
    result.sort(key=lambda x: (-x["count"], -x["docs"], x["phrase"]))
    return result[:top]


def _side_ok(neighbors: Counter, total: int) -> bool:
    if neighbors.get("|", 0) / total >= 0.5:
        return True
    return len(neighbors) >= 2 and max(neighbors.values()) / total <= 0.75


def lead_words(items: Iterable[Tuple[int, str]], top: int, need_docs: int) -> List[dict]:
    counts: Counter = Counter()
    docs_of: Dict[str, set] = defaultdict(set)
    for di, text in items:
        han = "".join(HAN_RE.findall(text[:6]))
        if len(han) >= 2:
            word = han[:2]
            counts[word] += 1
            docs_of[word].add(di)
    rows = [{"word": w, "count": n, "docs": len(docs_of[w])}
            for w, n in counts.items() if n >= 2 and len(docs_of[w]) >= need_docs]
    rows.sort(key=lambda x: (-x["count"], -x["docs"], x["word"]))
    return rows[:top]


def paragraph_openers(docs: Sequence["Doc"], top: int, need_docs: int) -> List[dict]:
    """段首用语：段落第一个逗号前不超过 6 字的分句，如"今年以来""同时"，以及"一是"。数字记作×。"""
    counts: Counter = Counter()
    docs_of: Dict[str, set] = defaultdict(set)
    for di, d in enumerate(docs):
        for p in d.paragraphs:
            match = re.match(r"^([^，,。；：！？]{1,6})[，,]", p)
            if match and HAN_RE.search(match.group(1)):
                key = re.sub(r"\d+", "×", match.group(1))
            elif re.match(r"^[一二三四五六七八九十]是", p):
                key = "一是"
            else:
                continue
            counts[key] += 1
            docs_of[key].add(di)
    rows = [{"word": w, "count": n, "docs": len(docs_of[w])}
            for w, n in counts.items() if n >= 2 and len(docs_of[w]) >= need_docs]
    rows.sort(key=lambda x: (-x["count"], -x["docs"], x["word"]))
    return rows[:top]


def build_report(paths: Sequence[Path], top: int = 20, min_count: int = 3, examples: bool = False) -> dict:
    notes: List[str] = []
    documents: List[dict] = []
    skipped: List[dict] = []
    docs: List[Doc] = []
    privacy_rows: List[dict] = []
    for path in paths:
        raw = read_lines(path, notes)
        extra = secret_marks.docx_mark_lines(path) if path.suffix.lower() == ".docx" else []
        counts, marks = privacy_scan(raw, extra)
        if marks:
            skipped.append({"file": path.name, "marks": marks,
                            "reason": f"发现{secret_marks.describe(marks)}，未统计；请改用单位内网部署、允许处理此类材料的模型，"
                                      "或删除相关内容后再统计"})
            continue
        if counts:
            privacy_rows.append({"file": path.name, "hits": counts})
        masked = [mask_privacy(line) for line in raw]
        doc = split_structure(path.name, masked)
        if nchars(doc.text_for_density) == 0:
            skipped.append({"file": path.name, "reason": "没有可统计的正文"})
            continue
        stats = analyze_doc(doc)
        stats = OrderedDict([("file", path.name)] + list(stats.items()))
        documents.append(stats)
        docs.append(doc)

    counted = len(documents)
    if counted and counted < RECOMMENDED_MIN:
        notes.append(f"可统计的样本只有 {counted} 篇，少于 {RECOMMENDED_MIN} 篇，参数代表性有限，写入文风参数时注明样本数")
    if counted > RECOMMENDED_MAX:
        notes.append(f"样本 {counted} 篇，多于 {RECOMMENDED_MAX} 篇；文书类型不同的，建议按类型分组分别统计")

    summary: dict = {}
    if documents:
        need_docs = 2 if counted >= 2 else 1
        all_sent = [nchars(s) for d in docs for p in d.paragraphs for s in split_sentences(p)]
        all_para = [nchars(p) for d in docs for p in d.paragraphs]
        heading_texts = [(lv, h) for d in docs for lv, h in d.headings]
        styles = Counter(heading_style(h) for lv, h in heading_texts if lv in ("1", "2"))
        summary = OrderedDict()
        summary["doc_chars"] = quantiles([d["chars"] for d in documents])
        summary["sentence_len"] = quantiles(all_sent)
        summary["short_sentence_ratio"] = round(sum(1 for n in all_sent if n <= 20) / len(all_sent), 3) if all_sent else 0.0
        summary["long_sentence_ratio"] = round(sum(1 for n in all_sent if n >= 80) / len(all_sent), 3) if all_sent else 0.0
        summary["paragraph_len"] = quantiles(all_para)
        summary["headings_per_doc"] = {lv: quantiles([d["headings"][lv] for d in documents]) for lv in HEADING_RES}
        heading_info = OrderedDict()
        for lv in HEADING_RES:
            texts = [h for level, h in heading_texts if level == lv]
            if texts:
                heading_info[lv] = {"count": len(texts), "length": quantiles([nchars(h) for h in texts])}
        summary["heading"] = OrderedDict([
            ("by_level", heading_info),
            ("styles_level1_2", dict(styles.most_common())),
            ("number_summary", sum(1 for _, h in heading_texts if NUM_SUMMARY_RE.search(h))),
            ("lead_words", lead_words(((i, h) for i, d in enumerate(docs) for lv, h in d.headings
                                       if not STRUCTURAL_HEADING_RE.match(h)), top, need_docs)),
            ("titles", sum(1 for d in documents if d["title"])),
            ("subtitles", sum(1 for d in documents if d["subtitle"])),
        ])
        if examples:
            sample: Dict[str, List[str]] = {}
            for lv, h in heading_texts:
                if lv in ("1", "2") and len(sample.setdefault(lv, [])) < 5:
                    sample[lv].append(h)
            summary["heading"]["examples"] = sample
        summary["yishi"] = {
            "per_1000": quantiles([d["yishi"]["per_1000"] for d in documents]),
            "docs_using": sum(1 for d in documents if d["yishi"]["count"]),
            "max_enum": max(d["yishi"]["max_enum"] for d in documents),
        }
        summary["punctuation_per_1000"] = OrderedDict(
            (name, quantiles([d["punctuation_per_1000"][name] for d in documents]))
            for name in documents[0]["punctuation_per_1000"])
        summary["wording_per_1000"] = OrderedDict(
            (name, quantiles([d["wording_per_1000"][name] for d in documents])) for name in WORDING)
        summary["numbers_per_1000"] = quantiles([d["numbers_per_1000"] for d in documents])
        summary["docs_with_percent"] = sum(1 for d in documents if d["percent_count"])
        summary["openings"] = dict(Counter(d["opening"] for d in documents).most_common())
        summary["endings"] = dict(Counter(d["ending"] for d in documents).most_common())
        texts = [d.text_for_density for d in docs]
        summary["four_char_phrases"] = four_char_phrases(texts, min_count, top)
        set_rows = []
        for label, regex in SET_PHRASES.items():
            per_doc = [len(regex.findall(t)) for t in texts]
            if sum(per_doc):
                set_rows.append({"phrase": label, "count": sum(per_doc), "docs": sum(1 for n in per_doc if n)})
        set_rows.sort(key=lambda x: (-x["docs"], -x["count"]))
        summary["set_phrases"] = set_rows
        jiguan = []
        total_chars = sum(d["chars"] for d in documents)
        for word in JIGUAN_WORDS:
            n = sum(t.count(word) for t in texts)
            if n:
                jiguan.append({"word": word, "count": n, "per_1000": per_k(n, total_chars),
                               "docs": sum(1 for t in texts if word in t)})
        jiguan.sort(key=lambda x: -x["count"])
        summary["jiguan_words"] = jiguan[:top]
        summary["paragraph_openers"] = paragraph_openers(docs, top, need_docs)

    report = OrderedDict()
    report["tool"] = "style_profile.py"
    report["version"] = VERSION
    report["scope"] = "【起草说明】之前的正文；标题、主送机关行、落款、日期行、“附件：”行、表格不计入段落和句子统计"
    report["definitions"] = {
        "字数": "不含空白的字符数（含标点）",
        "句长": "按句号、问号、感叹号切分，不含句末标点",
        "密度": "每千字出现次数",
        "汇总": "中位数与四分位范围（P25～P75）",
        "短句、长句": "短句不超过 20 字，长句不少于 80 字",
    }
    report["sample"] = {"files": len(paths), "counted": counted, "skipped": skipped}
    report["notes"] = notes
    report["documents"] = documents
    report["summary"] = summary
    report["privacy"] = {
        "files_with_hits": len(privacy_rows),
        "by_file": privacy_rows,
        "note": "只报告类别和次数，不输出命中内容；命中片段已在统计前屏蔽" if privacy_rows else "",
    }
    return report


# ---------------------------------------------------------------------------
# 文本输出
# ---------------------------------------------------------------------------

def _q(stat: dict, unit: str = "") -> str:
    if not stat:
        return "无"
    if stat.get("n", 0) <= 1:
        return f"{_fmt(stat['median'])}{unit}"
    return f"{_fmt(stat['median'])}{unit}（{_fmt(stat['p25'])}～{_fmt(stat['p75'])}）"


def _fmt(value: float) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def render_text(report: dict) -> str:
    out: List[str] = []
    sample = report["sample"]
    out.append(f"文风参数统计（style_profile.py {report['version']}）")
    out.append(f"样本：{sample['files']} 篇，可统计 {sample['counted']} 篇，跳过 {len(sample['skipped'])} 篇")
    out.append("口径：" + report["scope"] + "；汇总为中位数（P25～P75）")
    for item in sample["skipped"]:
        out.append(f"已跳过：{item['file']}：{item['reason']}")
    for note in report["notes"]:
        out.append("提示：" + note)
    if report["privacy"]["by_file"]:
        out.append("隐私提示（只列类别和次数，已在统计前屏蔽，请先脱敏再保存样本）：")
        for row in report["privacy"]["by_file"]:
            hits = "、".join(f"{k} {v} 处" for k, v in row["hits"].items())
            out.append(f"  {row['file']}：{hits}")
    docs = report["documents"]
    if not docs:
        out.append("没有可统计的正文。")
        return "\n".join(out)
    out.append("")
    out.append("一、每篇概况")
    out.append("| 文件 | 字数 | 段落 | 句子 | 句长中位数 | 一级标题 | 二级标题 | 一是…（每千字） | 开头 | 结尾 |")
    out.append("|---|---|---|---|---|---|---|---|---|---|")
    for d in docs:
        out.append(f"| {d['file']} | {d['chars']} | {d['paragraphs']} | {d['sentences']} | "
                   f"{_fmt(d['sentence_len'].get('median', 0))} | {d['headings']['1']} | {d['headings']['2']} | "
                   f"{_fmt(d['yishi']['per_1000'])} | {d['opening']} | {d['ending']} |")
    s = report["summary"]
    out.append("")
    out.append("二、量化参数（中位数，括号内为 P25～P75）")
    out.append("| 参数 | 本批样本 |")
    out.append("|---|---|")
    out.append(f"| 篇幅（字） | {_q(s['doc_chars'])}，最短 {_fmt(s['doc_chars']['min'])}，最长 {_fmt(s['doc_chars']['max'])} |")
    out.append(f"| 句长（字） | {_q(s['sentence_len'])}；短句（≤20字）占 {round(s['short_sentence_ratio'] * 100)}%，"
               f"长句（≥80字）占 {round(s['long_sentence_ratio'] * 100)}% |")
    out.append(f"| 段长（字） | {_q(s['paragraph_len'])} |")
    for lv, name in (("1", "一级标题（个/篇）"), ("2", "二级标题（个/篇）"), ("3", "三级标题（个/篇）")):
        out.append(f"| {name} | {_q(s['headings_per_doc'][lv])} |")
    by_level = s["heading"]["by_level"]
    for lv, name in (("1", "一级标题长度（字）"), ("2", "二级标题长度（字）")):
        if lv in by_level:
            out.append(f"| {name} | {_q(by_level[lv]['length'])} |")
    if s["heading"]["styles_level1_2"]:
        styles = "、".join(f"{k} {v}" for k, v in s["heading"]["styles_level1_2"].items())
        out.append(f"| 一、二级标题句式 | {styles}；含数字概括 {s['heading']['number_summary']} 个 |")
    out.append(f"| “一是……”（处/千字） | {_q(s['yishi']['per_1000'])}；{s['yishi']['docs_using']} 篇使用，最多列到第 {s['yishi']['max_enum']} 项 |")
    for name, stat in s["punctuation_per_1000"].items():
        if name in ("顿号", "破折号", "冒号", "分号", "感叹号", "问号"):
            out.append(f"| {name}（个/千字） | {_q(stat)} |")
    for name, stat in s["wording_per_1000"].items():
        out.append(f"| {name}（次/千字） | {_q(stat)} |")
    out.append(f"| 数字（个/千字） | {_q(s['numbers_per_1000'])}；{s['docs_with_percent']} 篇用了百分数 |")
    out.append("")
    out.append("三、高频表述（只列短语，不含原文句子）")
    if s["four_char_phrases"]:
        out.append("四字格：" + "、".join(f"{x['phrase']}（{x['count']}次/{x['docs']}篇）" for x in s["four_char_phrases"]))
    else:
        out.append("四字格：未达到统计门槛")
    if s["set_phrases"]:
        out.append("惯用语：" + "、".join(f"{x['phrase']}（{x['docs']}篇）" for x in s["set_phrases"]))
    if s["jiguan_words"]:
        out.append("机关常用语：" + "、".join(f"{x['word']}（{x['count']}次，{_fmt(x['per_1000'])}/千字）" for x in s["jiguan_words"]))
    if s["heading"]["lead_words"]:
        out.append("标题领起词：" + "、".join(f"{x['word']}（{x['count']}次/{x['docs']}篇）" for x in s["heading"]["lead_words"]))
    if s["paragraph_openers"]:
        out.append("段首用语：" + "、".join(f"{x['word']}（{x['count']}次/{x['docs']}篇）" for x in s["paragraph_openers"]))
    out.append("开头方式：" + "、".join(f"{k} {v} 篇" for k, v in s["openings"].items()))
    out.append("结尾方式：" + "、".join(f"{k} {v} 篇" for k, v in s["endings"].items()))
    examples = s["heading"].get("examples")
    if examples:
        out.append("")
        out.append("标题示例（仅供归纳句式，不写入文风参数文件）：")
        for lv, items in examples.items():
            out.append(f"  {lv}级：" + "；".join(items))
    out.append("")
    out.append("说明：参数是本批样本的参考范围，起草时不必逐项凑到中位数；与公文规范冲突的习惯不模仿。")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    k = 2
    while True:
        candidate = path.with_name(f"{path.stem}-v{k}{path.suffix}")
        if not candidate.exists():
            return candidate
        k += 1


def _setup_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


class _ChineseHelpFormatter(argparse.RawDescriptionHelpFormatter):
    def add_usage(self, usage, actions, groups, prefix=None):
        return super().add_usage(usage, actions, groups, "用法：" if prefix is None else prefix)


def make_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="style_profile.py",
        description="统计本单位历史稿的文风参数：句长、段长、各级标题、“一是……”使用率、标点密度、"
                    "措辞层级、高频四字格和惯用语、篇幅分布。只输出参数和短语，不输出原文段落。",
        epilog="退出码：0 正常；1 没有可统计的文件；2 输入错误。",
        formatter_class=_ChineseHelpFormatter,
        add_help=False,
    )
    parser._optionals.title = "选项"
    parser._positionals.title = "参数"
    parser.add_argument("-h", "--help", action="help", help="显示本帮助信息并退出")
    parser.add_argument("inputs", nargs="+", help="历史稿文件（.txt、.md、.docx）或所在目录，建议 3～10 篇")
    parser.add_argument("--json", action="store_true", help="以 JSON 格式输出")
    parser.add_argument("--top", type=int, default=20, help="高频短语最多列出几条（默认 20）")
    parser.add_argument("--min-count", type=int, default=3, help="四字格至少出现几次才列出（默认 3）")
    parser.add_argument("--examples", action="store_true", help="附一、二级标题示例各 5 个，供归纳句式")
    parser.add_argument("--output", help="把文本结果另存为文件；文件已存在时加 -v2、-v3，不覆盖")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    _setup_console()
    parser = make_parser()
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:
        return 0 if exc.code == 0 else 2
    if args.top < 1 or args.min_count < 1:
        print("错误：--top 和 --min-count 必须是正整数", file=sys.stderr)
        return 2
    try:
        paths = expand_inputs(args.inputs)
        report = build_report(paths, top=args.top, min_count=args.min_count, examples=args.examples)
    except InputError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    text = json.dumps(report, ensure_ascii=False, indent=2) if args.json else render_text(report)
    if args.output:
        target = unique_path(Path(args.output))
        try:
            target.write_text(text + "\n", encoding="utf-8")
        except OSError as exc:
            print(f"错误：无法写入 {target}（{exc}）", file=sys.stderr)
            return 2
        print(f"已写入：{target}")
    else:
        print(text)
    return 0 if report["sample"]["counted"] else 1


if __name__ == "__main__":
    sys.exit(main())
