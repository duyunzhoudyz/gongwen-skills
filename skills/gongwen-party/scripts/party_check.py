#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""党建材料的确定性检查（gongwen-party 自带脚本）。

子命令：
  measures  <稿件>
      查整改措施：每条措施有没有时限；措施条数是否少于查摆问题条数。
      措施开头复述问题的引导分句（“针对7月12日……问题，”）中的日期不算时限。
      只检查【起草说明】之前的正文。
  opinions  <征求意见清单> <稿件>
      查征求意见逐条回应表：清单中的每条意见是否都有对应行，
      意见是否照录原文，整改措施栏是否为空。回应表可以在正文中，也可以在【起草说明】之后。
      清单按“（一）班子建设方面”等分类、每类重新编号的，跳过分类标题，按出现顺序统一编号；
      回应表的行先按意见原文对应，其次按原文相近，最后按序号。

输入支持 .txt、.md、.docx（docx 需要 python-docx）。只用标准库，不联网，不写文件。
退出码：0 完成检查；1 使用 --strict 且存在 P0、P1；2 输入错误。
"""
from __future__ import annotations

import argparse
import difflib
import json
import os
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

BOUNDARY = "【起草说明】"
# 分界行写法与 gongwen_lint.py、fact_diff.py 保持一致：行首可有空白、Markdown 标题符号或加粗符号
DRAFT_NOTE_RE = re.compile(r"^\s*(?:#{1,6}\s*)?(?:(?:\*\*|__)\s*)?" + re.escape(BOUNDARY))
SUPPORTED_SUFFIXES = (".txt", ".md", ".docx")
EXCERPT_LIMIT = 30

CN_DIGITS = "一二三四五六七八九十"

# 层次序数（output-format.md 第一部分）：一、→（一）→ 1. →（1）
LEVEL_PATTERNS = (
    (1, re.compile(r"^\s*[一二三四五六七八九十]+、")),
    (2, re.compile(r"^\s*[（(][一二三四五六七八九十]+[）)]")),
    (3, re.compile(r"^\s*\d{1,2}[.．](?!\d)")),
    (4, re.compile(r"^\s*[（(]\d{1,2}[）)]")),
)

# 部分标题关键词
MEASURE_TITLE_RE = re.compile(r"整改|措施|努力方向|改进方向|下一步|今后|打算")
PROBLEM_TITLE_RE = re.compile(r"问题|不足|差距|查摆|检视|短板|对照检查")
PROBLEM_EXCLUDE_RE = re.compile(r"原因|根源|剖析|措施|整改|落实|打算|方向")
MEASURE_EXCLUDE_RE = re.compile(r"情况")

# 问题和措施写在同一部分的标题，如“存在问题及整改措施”
MIXED_TITLE_RE = re.compile(r"(?:问题|不足|差距)[^，。]{0,6}(?:和|及|与|、|以及)[^，。]{0,6}(?:整改|措施|方向|打算)")

# 段内列举“一是……二是……”
YISHI_RE = re.compile(r"(?:^|(?<=[。；;！？：:，,\s]))([一二三四五六七八九十]+是)")

# 时限：具体日期、期限、频次、整改状态用语，或标明待补、待核实的时限
DEADLINE_RE = re.compile(
    r"(?:\d{4}|[〇零一二三四五六七八九]{4})年"
    r"|\d{1,2}月份?"
    # 汉字月份：前面不是其他汉字（排除“统一月报”的“一月”），或前面是年份、介词；也可以后面跟“底、前、份”等
    r"|(?:(?<![一-鿿])|(?<=[年从于在至到自及和、]))[一二三四五六七八九十]{1,3}月"
    r"|[一二三四五六七八九十]{1,3}月(?:份|底|末|初|前|上旬|中旬|下旬)"
    r"|年底|年末|年内|月底|月末|季度末|季末|学期末|半年内"
    # 年初、年中：前面是“今、明、本”或介词，或后面跟“前、完成”（排除“落实年初定的制度”）
    r"|(?<=[今明本次当到于在至])年(?:初|中)|年(?:初|中)(?=前|以前|完成)"
    r"|\d+\s*个?(?:工作日|日|天|周|星期|月|年)(?:内|以内|之内)"
    r"|[一二三四五六七八九十两半]+个?(?:工作日|日|天|周|星期|月|年)(?:内|以内|之内)"
    r"|每(?:年|半年|季度|季|个月|月|周|星期|旬|天|日|学期)"
    r"|立行立改|即知即改"
    r"|(?:本|下)(?:月|季度|学期|周)"
    r"|【待(?:补|核实)：[^】]*(?:时限|时间|期限|日期|完成|频次)[^】]*】"
)

# 只写“长期坚持”的措施不可检验，单独提示补检查频次
LONG_TERM_RE = re.compile(r"长期坚持")

# 措施开头复述问题的引导分句（“针对7月12日河堤管涌暴露的排查不实问题，”），其中的日期不是时限
LEAD_CLAUSE_RE = re.compile(
    r"(?:针对|对于|对照|关于|就)[^，。；,;]*?"
    r"(?:问题|不足|短板|差距|弱项|情况|意见|现象|事件)的?[，,：:]"
)

# 征求意见清单的编号行
NUMBERED_LINE_RE = re.compile(
    r"^\s*(?:"
    r"[（(]\s*(?P<a>\d{1,3})\s*[）)]"
    r"|(?P<b>\d{1,3})\s*[.．、:：)）]"
    r"|(?P<c>[一二三四五六七八九十]{1,3})\s*[、.．:：]"
    r"|[（(]\s*(?P<d>[一二三四五六七八九十]{1,3})\s*[）)]"
    r")\s*(?P<text>.*)$"
)

# 分组式清单的分类标题（“（一）班子建设方面”“作风方面：”）：没有句读，不超过 GROUP_TITLE_LIMIT 字。
# 编号的分类标题看编号层级判断（其后紧跟另一种编号的意见）；不编号的还须以“方面、类”等或冒号结尾
GROUP_STRICT_RE = re.compile(r"(?:方面|类|类别|领域|板块)(?:[（(][^）)]*[）)])?[：:]?$|[：:]$")
SENTENCE_PUNCT_RE = re.compile(r"[，。；！？,;!?]")
GROUP_TITLE_LIMIT = 30
# 清单末尾的说明行（“以上意见由组织部门反馈。”）不算意见
NOTE_LINE_RE = re.compile(r"^(?:以上|上述|注[：:]|备注|说明[：:])")
# 不编号清单的标题行（“某单位征求意见汇总”“征求意见清单（反馈给某副镇长）”）不算意见
LIST_TITLE_RE = re.compile(r"(?:汇总|清单)(?:[（(][^）)]*[）)])?$")
FUZZY_MATCH_RATIO = 0.6

TABLE_ROW_RE = re.compile(r"^\s*\|.*\|\s*$")
TABLE_SEP_RE = re.compile(r"^\s*\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)*\|?\s*$")
# 比对意见原文时忽略的字符：空白、标点、Markdown 符号
NORMALIZE_RE = re.compile(r"[\s　，。、；：？！“”‘’\"'（）()《》〈〉【】\[\],.;:?!…—\-·*_`|]")
EMPTY_CELL_RE = re.compile(r"^[\s\-—–/／无暂无空]*$")


class InputError(Exception):
    """输入错误，对应退出码 2。"""


@dataclass
class Issue:
    id: str
    level: str
    category: str
    line: Optional[int]
    excerpt: str
    message: str
    suggestion: str

    def to_dict(self) -> Dict[str, object]:
        return {
            "id": self.id,
            "level": self.level,
            "category": self.category,
            "line": self.line,
            "excerpt": self.excerpt,
            "message": self.message,
            "suggestion": self.suggestion,
        }


@dataclass
class Item:
    """一个列举项：起始行号和全文。"""
    line: int
    text: str


@dataclass
class Section:
    level: int
    line: int
    title: str
    lines: List[Tuple[int, str]] = field(default_factory=list)  # （行号, 文字），含标题行之后的内容
    role: str = "other"  # problem / measure / mixed / other


# ---------------------------------------------------------------------------
# 读取
# ---------------------------------------------------------------------------

def _decode(data: bytes, path: Path) -> str:
    for encoding in ("utf-8-sig", "gb18030"):
        try:
            text = data.decode(encoding)
        except UnicodeDecodeError:
            continue
        if encoding != "utf-8-sig":
            print(f"提示：{path.name} 不是 UTF-8 编码，已按 GB18030 读取。", file=sys.stderr)
        return text
    raise InputError(f"无法识别文件编码，请另存为 UTF-8：{path}")


def _read_docx(path: Path) -> str:
    try:
        import docx  # type: ignore
        from docx.table import Table  # type: ignore
        from docx.text.paragraph import Paragraph  # type: ignore
    except ImportError:
        raise InputError("读取 .docx 需要 python-docx，请先运行：pip install python-docx；"
                         "或把内容另存为 .txt 后再检查。")
    try:
        document = docx.Document(str(path))
    except Exception as exc:  # 文件损坏、不是 docx 等
        raise InputError(f"无法打开 docx 文件：{path}（{exc}）")
    lines: List[str] = []
    for child in document.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            lines.append(Paragraph(child, document).text)
        elif tag == "tbl":
            # Word 表格转成 Markdown 表格行，便于与文本输入统一解析；合并单元格按底层 tc 去重
            seen = set()
            for index, row in enumerate(Table(child, document).rows):
                cells = []
                for cell in row.cells:
                    if cell._tc in seen:
                        continue
                    seen.add(cell._tc)
                    cells.append(cell.text.replace("\n", " ").replace("|", "／"))
                lines.append("| " + " | ".join(cells) + " |")
                if index == 0:
                    lines.append("|" + "---|" * max(len(cells), 1))
            lines.append("")
    return "\n".join(lines)


def _long_path(path: Path) -> Path:
    """Windows 下完整路径接近 260 字符上限时加长路径前缀。"""
    if os.name != "nt":
        return path
    try:
        full = str(path.resolve())
    except OSError:
        return path
    prefix = "\\\\?\\"
    if len(full) < 240 or full.startswith(prefix):
        return path
    if full.startswith("\\\\"):
        return Path(prefix + "UNC\\" + full[2:])
    return Path(prefix + full)


def read_text(path: Path) -> str:
    """读取 .txt、.md、.docx 文件全文，统一换行符。"""
    shown = path
    path = _long_path(path)
    if not path.exists():
        raise InputError(f"文件不存在：{shown}")
    if not path.is_file():
        raise InputError(f"不是文件：{shown}")
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise InputError(f"不支持的文件类型 {suffix or '（无扩展名）'}，只支持 .txt、.md、.docx：{shown}")
    if suffix == ".docx":
        text = _read_docx(path)
    else:
        try:
            data = path.read_bytes()
        except OSError as exc:
            raise InputError(f"读取失败：{shown}（{exc}）")
        text = _decode(data, path)
    return text.replace("\r\n", "\n").replace("\r", "\n")


def body_lines(text: str) -> List[str]:
    """【起草说明】之前的正文行。"""
    lines: List[str] = []
    for line in text.split("\n"):
        if DRAFT_NOTE_RE.match(line):
            break
        lines.append(line)
    return lines


def _excerpt(text: str) -> str:
    text = re.sub(r"\s+", "", text)
    return text if len(text) <= EXCERPT_LIMIT else text[:EXCERPT_LIMIT] + "……"


def cn_to_int(value: str) -> Optional[int]:
    """把一至九十九的汉字数字转成整数，失败返回 None。"""
    if not value or any(ch not in CN_DIGITS for ch in value):
        return None
    if value == "十":
        return 10
    if "十" not in value:
        return CN_DIGITS.index(value) + 1 if len(value) == 1 else None
    tens, _, ones = value.partition("十")
    if len(tens) > 1 or len(ones) > 1:
        return None
    tens_value = CN_DIGITS.index(tens) + 1 if tens else 1
    ones_value = CN_DIGITS.index(ones) + 1 if ones else 0
    if tens_value > 9:
        return None
    return tens_value * 10 + ones_value


def heading_level(line: str) -> Optional[int]:
    for level, pattern in LEVEL_PATTERNS:
        if pattern.match(line):
            return level
    return None


# ---------------------------------------------------------------------------
# measures：整改措施检查
# ---------------------------------------------------------------------------

def _classify(title: str) -> str:
    is_measure = bool(MEASURE_TITLE_RE.search(title)) and not MEASURE_EXCLUDE_RE.search(title)
    is_problem = bool(PROBLEM_TITLE_RE.search(title)) and not PROBLEM_EXCLUDE_RE.search(title)
    if is_measure and MIXED_TITLE_RE.search(title):
        return "mixed"
    if is_measure:
        return "measure"
    if is_problem:
        return "problem"
    return "other"


def split_sections(lines: Sequence[str]) -> List[Section]:
    """按正文中最高一级的层次序数切分部分。"""
    levels = [heading_level(line) for line in lines]
    present = [lv for lv in levels if lv is not None]
    if not present:
        return []
    top = min(present)
    sections: List[Section] = []
    current: Optional[Section] = None
    for number, (line, level) in enumerate(zip(lines, levels), start=1):
        if level == top:
            # 标题与正文同段时（“（一）强化学习。每月……”），标题取第一个句号之前
            title = re.split(r"[。；]", line.strip(), maxsplit=1)[0]
            current = Section(level=top, line=number, title=title)
            current.role = _classify(title)
            rest = line.strip()[len(title):].lstrip("。；")
            if rest:
                current.lines.append((number, rest))
            sections.append(current)
        elif current is not None:
            current.lines.append((number, line))
    return sections


def split_items(section: Section) -> List[Item]:
    """把一个部分拆成列举项：下一级层次序数 → “一是……二是……” → 自然段。"""
    content = [(n, t) for n, t in section.lines if t.strip()]
    if not content:
        return []
    child_levels = sorted({lv for lv in (heading_level(t) for _, t in content)
                           if lv is not None and lv > section.level})
    if child_levels:
        child = child_levels[0]
        items: List[Item] = []
        for number, text in content:
            if heading_level(text) == child:
                items.append(Item(number, text.strip()))
            elif items:
                items[-1].text += text.strip()
        if items:
            return items
    # 段内列举：只取“一是……”起的各段，“一是”之前的引导语不算一项
    items = []
    for number, text in content:
        stripped = text.strip()
        starts = [m.start(1) for m in YISHI_RE.finditer(stripped)]
        if len(starts) >= 2 or (starts and starts[0] == 0):
            for k, start in enumerate(starts):
                end = starts[k + 1] if k + 1 < len(starts) else len(stripped)
                items.append(Item(number, stripped[start:end]))
    if items:
        return items
    # 自然段：以冒号结尾的引导语不算一项
    return [Item(number, text.strip()) for number, text in content
            if not text.strip().endswith(("：", ":"))]


def check_measures(text: str) -> Dict[str, object]:
    lines = body_lines(text)
    sections = split_sections(lines)
    issues: List[Issue] = []
    problem_count = 0
    measure_count = 0
    measures_without_deadline = 0
    found_problem = found_measure = False

    for section in sections:
        if section.role == "problem":
            found_problem = True
            problem_count += len(split_items(section))
        elif section.role == "measure":
            found_measure = True
            items = split_items(section)
            measure_count += len(items)
            for item in items:
                if DEADLINE_RE.search(LEAD_CLAUSE_RE.sub("", item.text)):
                    continue
                measures_without_deadline += 1
                if LONG_TERM_RE.search(item.text):
                    issues.append(Issue(
                        "MEA-006", "P2", "整改措施", item.line, _excerpt(item.text),
                        "这条措施只写了“长期坚持”，无法检验",
                        "补检查频次（如“每季度对照检查1次”）或阶段目标；本人没有给出的不强行补，在起草说明“待补”一行列出",
                    ))
                else:
                    issues.append(Issue(
                        "MEA-001", "P2", "整改措施", item.line, _excerpt(item.text),
                        "这条措施没有写时限，不便检验是否完成",
                        "本人材料里有完成时间或频次的写进去（如“2027年3月底前”“每月”）；"
                        "本人没有给出的不强行补，也不插【待补】，在起草说明“待补”一行列出",
                    ))
        elif section.role == "mixed":
            issues.append(Issue(
                "MEA-004", "P2", "结构", section.line, _excerpt(section.title),
                "问题和措施写在同一部分，未自动核对措施时限",
                "人工核对每个问题后面是否有带时限的措施；按通知要求分开写的，问题、原因、措施各成一部分",
            ))

    if found_problem and not found_measure:
        issues.append(Issue(
            "MEA-003", "P1", "结构", None, "",
            "有查摆问题部分，没有找到整改措施部分",
            "增加整改措施（努力方向）部分，与问题一一对应；本人没有提供措施的，写“针对……问题，【待补：具体措施】”",
        ))
    if found_problem and found_measure and measure_count < problem_count:
        issues.append(Issue(
            "MEA-002", "P2", "整改措施", None, "",
            f"查摆问题 {problem_count} 条，整改措施 {measure_count} 条，措施少于问题",
            "逐个核对每个问题是否都有对应措施；一条措施对应几个问题的，在措施中点明对应的问题",
        ))
    if not found_problem and not found_measure and not any(s.role == "mixed" for s in sections):
        issues.append(Issue(
            "MEA-005", "P2", "结构", None, "",
            "没有识别到问题部分和措施部分（部分标题需用“一、”或“（一）”等层次序数，并含“问题”“措施”等字样）",
            "按 references/checklist.md 人工检查措施时限",
        ))
    return {
        "check": "measures",
        "scope": "【起草说明】之前的正文",
        "stats": {
            "sections": len(sections),
            "problems": problem_count,
            "measures": measure_count,
            "measures_without_deadline": measures_without_deadline,
        },
        "issues": [issue.to_dict() for issue in issues],
    }


# ---------------------------------------------------------------------------
# opinions：征求意见逐条回应检查
# ---------------------------------------------------------------------------

@dataclass
class Opinion:
    number: Optional[int]  # 解析后一律为整数
    text: str
    line: int


@dataclass
class Table:
    line: int
    header: List[str]
    rows: List[Tuple[int, List[str]]]


def _cells(row: str) -> List[str]:
    return [cell.strip() for cell in row.strip().strip("|").split("|")]


def find_tables(lines: Sequence[str]) -> List[Table]:
    tables: List[Table] = []
    index = 0
    while index < len(lines):
        line = lines[index]
        nxt = lines[index + 1] if index + 1 < len(lines) else ""
        if TABLE_ROW_RE.match(line) and TABLE_SEP_RE.match(nxt):
            table = Table(line=index + 1, header=_cells(line), rows=[])
            index += 2
            while index < len(lines) and TABLE_ROW_RE.match(lines[index]):
                if not TABLE_SEP_RE.match(lines[index]):
                    table.rows.append((index + 1, _cells(lines[index])))
                index += 1
            tables.append(table)
            continue
        index += 1
    return tables


def _column(header: Sequence[str], patterns: Sequence[str]) -> Optional[int]:
    names = [re.sub(r"[\s*_`]", "", h) for h in header]
    for pattern in patterns:
        for k, name in enumerate(names):
            if re.search(pattern, name):
                return k
    return None


def _parse_number(value: str) -> Optional[int]:
    value = re.sub(r"[\s*_（）()、.．第条号意见]", "", value)
    if value.isdigit():
        return int(value)
    return cn_to_int(value)


@dataclass
class OpinionList:
    opinions: List[Opinion]
    group_titles: int = 0     # 跳过的分类标题行数
    renumbered: bool = False  # 编号有重复，已按出现顺序重新编号


def normalize(text: str) -> str:
    return NORMALIZE_RE.sub("", text)


def _numbered_line(line: str) -> Optional[Tuple[str, Optional[int], str]]:
    """编号行返回（编号样式, 编号, 正文），不是编号行返回 None。样式：a（1） b 1. c 一、 d（一）。"""
    match = NUMBERED_LINE_RE.match(line)
    if not match or not match.group("text").strip():
        return None
    for style in ("a", "b", "c", "d"):
        raw = match.group(style)
        if raw:
            number = int(raw) if style in ("a", "b") else cn_to_int(raw)
            return style, number, match.group("text").strip()
    return None


def _is_group_title(text: str, strict: bool) -> bool:
    """分类标题：没有句读、不长。strict 时还须以“方面、类”等或冒号结尾。"""
    body = text.strip().rstrip("：:")
    if not body or SENTENCE_PUNCT_RE.search(body) or len(normalize(body)) > GROUP_TITLE_LIMIT:
        return False
    return bool(GROUP_STRICT_RE.search(text.strip())) or not strict


def _next_nonempty(lines: Sequence[str], index: int) -> Optional[int]:
    for k in range(index + 1, len(lines)):
        if lines[k].strip():
            return k
    return None


def _group_styles(lines: Sequence[str], parsed: Sequence[Optional[Tuple[str, Optional[int], str]]]
                  ) -> Tuple[List[str], bool]:
    """判断分组式清单：返回分类标题所用的编号样式，以及分类标题下的意见是否不编号。

    分类标题的样式先出现，其后紧跟另一种样式的编号行（“（一）班子建设方面”下接“1.”）；
    多级分组逐级判断。某样式的行只要有一行不像标题（有句读或过长），就不当作分类标题。
    """
    seq = [info for info in parsed if info]
    if not seq:
        return [], False

    def all_titles(style: str, strict: bool) -> bool:
        return all(_is_group_title(info[2], strict) for info in seq if info[0] == style)

    groups: List[str] = []
    pos = 0
    while pos + 1 < len(seq):
        style, next_style = seq[pos][0], seq[pos + 1][0]
        if next_style == style or next_style in groups or not all_titles(style, strict=False):
            break
        groups.append(style)
        pos += 1
    if groups:
        return groups, False

    # 只有一种编号，编号行都是“××方面”，其下是不编号、一行一条的意见
    style = seq[0][0]
    if all(info[0] == style for info in seq) and all_titles(style, strict=True):
        for index, info in enumerate(parsed):
            if not info:
                continue
            nxt = _next_nonempty(lines, index)
            if nxt is None or parsed[nxt] or NOTE_LINE_RE.match(lines[nxt].strip()):
                return [], False
        return [style], True
    return [], False


def _renumber_if_needed(opinions: List[Opinion]) -> bool:
    """编号重复（分组式清单每类重新编号）或部分缺失时，按出现顺序重新编号，返回 True。"""
    numbers = [op.number for op in opinions]
    if all(n is None for n in numbers):
        for position, op in enumerate(opinions, start=1):
            op.number = position
        return False
    if None not in numbers and len(set(numbers)) == len(numbers):
        return False
    for position, op in enumerate(opinions, start=1):
        op.number = position
    return True


def parse_opinion_list(text: str) -> OpinionList:
    """解析征求意见清单：优先读表头含“意见”的表格，其次读编号行，最后按非空行。

    分组式清单（“（一）班子建设方面”下每类重新从 1 编号）跳过分类标题；编号重复时按出现顺序重新编号。
    """
    lines = text.split("\n")
    for table in find_tables(lines):
        op_col = _column(table.header, [r"意见|建议|问题"])
        if op_col is None:
            continue
        idx_col = _column(table.header, [r"序号|编号"])
        opinions: List[Opinion] = []
        for line_no, cells in table.rows:
            text_cell = cells[op_col].strip() if op_col < len(cells) else ""
            if not text_cell:
                continue
            number = None
            if idx_col is not None and idx_col < len(cells):
                number = _parse_number(cells[idx_col])
            opinions.append(Opinion(number, text_cell, line_no))
        if opinions:
            result = OpinionList(opinions)
            result.renumbered = _renumber_if_needed(opinions)
            return result

    parsed = [_numbered_line(line) for line in lines]
    result = OpinionList([])
    if any(parsed):
        group_styles, plain_items = _group_styles(lines, parsed)
        continuing = False  # 上一行属于最后一条意见（中间没有空行），本行可能是折行
        for index, (line, info) in enumerate(zip(lines, parsed)):
            stripped = line.strip()
            if info and info[0] in group_styles:
                result.group_titles += 1
                continuing = False
            elif info:
                result.opinions.append(Opinion(info[1], info[2], index + 1))
                continuing = True
            elif not stripped or NOTE_LINE_RE.match(stripped):
                continuing = False
            elif _is_group_title(stripped, strict=True):
                # 不编号的分类标题（“作风方面”）或引导语，不是上一条意见的折行
                nxt = _next_nonempty(lines, index)
                if nxt is not None and parsed[nxt]:
                    result.group_titles += 1
                continuing = False
            elif plain_items and result.group_titles:
                result.opinions.append(Opinion(None, stripped, index + 1))  # 分类标题下一行一条
            elif continuing and not stripped.endswith(("：", ":")):
                result.opinions[-1].text += stripped  # 一条意见折行的续行
            else:
                continuing = False
        result.renumbered = _renumber_if_needed(result.opinions)
        return result

    for line_no, line in enumerate(lines, start=1):
        stripped = line.strip()
        if (not stripped or stripped.endswith(("：", ":")) or stripped.startswith("#")
                or NOTE_LINE_RE.match(stripped)):
            continue
        if _is_group_title(stripped, strict=True):
            result.group_titles += 1
            continue
        if LIST_TITLE_RE.search(stripped) and not SENTENCE_PUNCT_RE.search(stripped):
            continue
        result.opinions.append(Opinion(None, stripped, line_no))
    _renumber_if_needed(result.opinions)
    return result


def parse_opinions(text: str) -> List[Opinion]:
    """只返回意见列表，供简单调用。"""
    return parse_opinion_list(text).opinions


@dataclass
class Row:
    line: int
    cells: List[str]
    number: Optional[int]
    key: str  # 意见栏去掉空白和标点后的文字


def _match_rows(opinions: Sequence[Opinion], rows: Sequence[Row]) -> Dict[int, int]:
    """把清单意见与回应表行一一对应，返回 {意见下标: 行下标}。

    依次按：意见原文完全一致 → 原文相近（相似度不低于 FUZZY_MATCH_RATIO）→ 序号一致。
    序号只在回应表序号不重复时使用，避免分组编号错位。
    """
    assigned: Dict[int, int] = {}
    used = set()
    keys = [normalize(op.text) for op in opinions]
    for i, key in enumerate(keys):
        for j, row in enumerate(rows):
            if j not in used and row.key and row.key == key:
                assigned[i] = j
                used.add(j)
                break

    numbers = [row.number for row in rows if row.number is not None]
    numbers_reliable = len(numbers) == len(set(numbers))
    candidates = []
    for i, key in enumerate(keys):
        if i in assigned or not key:
            continue
        for j, row in enumerate(rows):
            if j in used or not row.key:
                continue
            ratio = difflib.SequenceMatcher(None, key, row.key).ratio()
            if ratio >= FUZZY_MATCH_RATIO:
                bonus = 0.05 if numbers_reliable and row.number == opinions[i].number else 0.0
                candidates.append((ratio + bonus, i, j))
    for _score, i, j in sorted(candidates, key=lambda c: (-c[0], c[1], c[2])):
        if i not in assigned and j not in used:
            assigned[i] = j
            used.add(j)

    if numbers_reliable:
        for i, op in enumerate(opinions):
            if i in assigned:
                continue
            for j, row in enumerate(rows):
                if j not in used and row.number is not None and row.number == op.number:
                    assigned[i] = j
                    used.add(j)
                    break
    return assigned


def check_opinions(list_text: str, draft_text: str) -> Dict[str, object]:
    parsed_list = parse_opinion_list(list_text)
    opinions = parsed_list.opinions
    if not opinions:
        raise InputError("征求意见清单中没有读到任何意见（每条意见一行，或用“1.”“（1）”等编号）")
    issues: List[Issue] = []
    lines = draft_text.split("\n")
    table = None
    for candidate in find_tables(lines):
        if _column(candidate.header, [r"意见"]) is not None:
            table = candidate
            if _column(candidate.header, [r"整改措施|措施|回应|处理"]) is not None:
                break
    stats = {
        "opinions": len(opinions),
        "answered": 0,
        "missing": [],
        "table_line": None,
        "group_titles": parsed_list.group_titles,
        "renumbered": parsed_list.renumbered,
    }
    if table is None:
        issues.append(Issue(
            "OPN-005", "P1", "征求意见回应", None, "",
            "没有找到征求意见逐条回应表（表头含“意见”的 Markdown 表格）",
            "在起草说明末尾附“征求意见逐条回应表”：序号｜征求到的意见（照录原文）｜对应的查摆问题｜整改措施｜时限",
        ))
        stats["missing"] = [op.number for op in opinions]
        return {"check": "opinions", "stats": stats, "issues": [i.to_dict() for i in issues]}

    stats["table_line"] = table.line
    idx_col = _column(table.header, [r"序号|编号"])
    op_col = _column(table.header, [r"征求.*意见", r"意见.*原文", r"意见"])
    measure_col = _column(table.header, [r"整改措施", r"措施", r"回应", r"处理"])
    if measure_col is not None and measure_col == op_col:
        measure_col = None
    if measure_col is None:
        issues.append(Issue(
            "OPN-007", "P2", "征求意见回应", table.line, " | ".join(table.header),
            "回应表没有“整改措施”或“回应”栏，无法核对每条意见的处理",
            "增加“整改措施”栏",
        ))

    def cell(cells: List[str], col: Optional[int]) -> str:
        return cells[col] if col is not None and col < len(cells) else ""

    rows: List[Row] = []
    for line_no, cells in table.rows:
        if not any(c.strip() for c in cells):
            continue
        number = _parse_number(cell(cells, idx_col)) if idx_col is not None else None
        rows.append(Row(line_no, cells, number, normalize(cell(cells, op_col))))

    assigned = _match_rows(opinions, rows)
    for i, op in enumerate(opinions):
        if i not in assigned:
            stats["missing"].append(op.number)
            issues.append(Issue(
                "OPN-001", "P1", "征求意见回应", op.line, _excerpt(op.text),
                f"清单第 {op.number} 条意见在回应表中没有对应行",
                "补上这一条：照录原文，写明对应的问题和整改措施；不属于本人（班子）职责的，写明转交哪里及理由",
            ))
            continue
        stats["answered"] += 1
        row = rows[assigned[i]]
        if op_col is not None and row.key != normalize(op.text):
            issues.append(Issue(
                "OPN-004", "P1", "征求意见回应", row.line, _excerpt(cell(row.cells, op_col)),
                f"第 {op.number} 条意见的表述与清单原文不一致",
                f"照录清单原文“{_excerpt(op.text)}”，不概括、不弱化",
            ))
        if measure_col is not None:
            measure = cell(row.cells, measure_col)
            if EMPTY_CELL_RE.match(measure):
                issues.append(Issue(
                    "OPN-002", "P1", "征求意见回应", row.line, _excerpt(cell(row.cells, op_col) or op.text),
                    f"第 {op.number} 条意见没有写整改措施或处理意见",
                    "写明整改措施和时限；本人还没想好的，写【待补：整改措施】",
                ))
            elif "【待补" in measure:
                issues.append(Issue(
                    "OPN-003", "P2", "征求意见回应", row.line, _excerpt(measure),
                    f"第 {op.number} 条意见的整改措施待补",
                    "定稿前由本人补齐",
                ))
    used_rows = set(assigned.values())
    for j, row in enumerate(rows):
        if j in used_rows:
            continue
        label = f"序号 {row.number} 的一行" if row.number is not None else "这一行"
        issues.append(Issue(
            "OPN-006", "P2", "征求意见回应", row.line, _excerpt(cell(row.cells, op_col)),
            f"回应表中{label}在征求意见清单中找不到对应的意见",
            "核对是否照录了清单原文；清单之外的内容不放进回应表",
        ))
    return {"check": "opinions", "stats": stats, "issues": [i.to_dict() for i in issues]}


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------

def summarize(issues: Sequence[Dict[str, object]]) -> Dict[str, int]:
    summary = {"P0": 0, "P1": 0, "P2": 0}
    for issue in issues:
        summary[str(issue["level"])] = summary.get(str(issue["level"]), 0) + 1
    return summary


def render_text(report: Dict[str, object]) -> str:
    stats = report["stats"]
    if report["check"] == "measures":
        head = (f"整改措施检查：{report['file']}（{report['scope']}）\n"
                f"识别到查摆问题 {stats['problems']} 条，整改措施 {stats['measures']} 条，"
                f"其中没有时限的 {stats['measures_without_deadline']} 条。")
    else:
        head = (f"征求意见回应检查：清单 {report['list_file']}，稿件 {report['file']}\n"
                f"清单共 {stats['opinions']} 条意见，回应表覆盖 {stats['answered']} 条。")
        if stats.get("group_titles"):
            head += f"\n清单中有 {stats['group_titles']} 行分类标题，未计入意见。"
        if stats.get("renumbered"):
            head += ("\n清单编号有重复（分类后每类重新编号），已按出现顺序统一编号为 1 至 "
                     f"{stats['opinions']}，回应表按意见原文对应。")
    lines = [head, ""]
    issues = report["issues"]
    if not issues:
        lines.append("结论：未发现问题。措施是否具体、可检验，仍需人工判断。")
        return "\n".join(lines)
    for k, issue in enumerate(issues, start=1):
        where = f"第 {issue['line']} 行" if issue["line"] else "全文"
        lines.append(f"{k}. [{issue['level']}] {issue['id']} {where}：{issue['message']}")
        if issue["excerpt"]:
            lines.append(f"   原文：{issue['excerpt']}")
        lines.append(f"   建议：{issue['suggestion']}")
    summary = report["summary"]
    lines.append("")
    lines.append(f"合计：P0 {summary['P0']} 条，P1 {summary['P1']} 条，P2 {summary['P2']} 条。")
    return "\n".join(lines)


def _setup_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:  # 被重定向为 StringIO 等对象时没有 reconfigure
            pass


class _ChineseHelpFormatter(argparse.RawDescriptionHelpFormatter):
    def add_usage(self, usage, actions, groups, prefix=None):
        return super().add_usage(usage, actions, groups, "用法：" if prefix is None else prefix)


def _common(parser: argparse.ArgumentParser) -> None:
    parser._optionals.title = "选项"
    parser._positionals.title = "参数"
    parser.add_argument("-h", "--help", action="help", help="显示本帮助信息并退出")
    parser.add_argument("--json", action="store_true", help="以 JSON 格式输出")
    parser.add_argument("--strict", action="store_true", help="存在 P0、P1 时返回 1")


def make_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="party_check.py",
        description="党建材料的确定性检查：整改措施时限、征求意见逐条回应。",
        epilog="退出码：0 完成检查；1 使用 --strict 且存在 P0、P1；2 输入错误。",
        formatter_class=_ChineseHelpFormatter,
        add_help=False,
    )
    parser._optionals.title = "选项"
    parser.add_argument("-h", "--help", action="help", help="显示本帮助信息并退出")
    sub = parser.add_subparsers(dest="command", title="子命令", metavar="{measures,opinions}")
    measures = sub.add_parser(
        "measures", add_help=False, formatter_class=_ChineseHelpFormatter,
        help="查整改措施有没有时限、措施条数是否少于问题条数",
        description="查整改措施有没有时限、措施条数是否少于问题条数。只检查【起草说明】之前的正文。",
    )
    measures.add_argument("draft", help="稿件（.txt、.md、.docx）")
    _common(measures)
    opinions = sub.add_parser(
        "opinions", add_help=False, formatter_class=_ChineseHelpFormatter,
        help="查征求意见逐条回应表是否逐条覆盖、照录原文、写了措施",
        description="查征求意见逐条回应表：是否逐条覆盖、意见是否照录原文、整改措施栏是否为空。"
                    "回应表为表头含“意见”的 Markdown 表格，可放在【起草说明】之后。",
    )
    opinions.add_argument("opinion_list", help="征求意见清单（每条一行，或带编号，或表头含“意见”的表格）")
    opinions.add_argument("draft", help="含征求意见逐条回应表的稿件（.txt、.md、.docx）")
    _common(opinions)
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
        if args.command == "measures":
            report = check_measures(read_text(Path(args.draft)))
            report["file"] = Path(args.draft).name
        else:
            list_text = read_text(Path(args.opinion_list))
            report = check_opinions(list_text, read_text(Path(args.draft)))
            report["file"] = Path(args.draft).name
            report["list_file"] = Path(args.opinion_list).name
    except InputError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    report["summary"] = summarize(report["issues"])
    if args.json:
        ordered = {"file": report["file"]}
        if "list_file" in report:
            ordered["list_file"] = report["list_file"]
        ordered.update({k: v for k, v in report.items() if k not in ordered})
        print(json.dumps(ordered, ensure_ascii=False, indent=2))
    else:
        print(render_text(report))
    if args.strict and (report["summary"]["P0"] or report["summary"]["P1"]):
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
