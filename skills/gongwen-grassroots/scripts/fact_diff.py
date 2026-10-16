# -*- coding: utf-8 -*-
"""改稿前后事实一致性检查。

源文件位于仓库 shared/scripts/，各 skill 中的副本由 tools/sync_shared.py 生成，请在源文件处修改。
维护者：gongwen-humanize。接口见 docs/开发规范.md 8.1 节。

用法：
    python scripts/fact_diff.py <原稿> <改稿> [--json] [--strict]
    python scripts/fact_diff.py <素材台账.md> <正文.txt> --ledger [--json] [--strict]

--ledger（台账对账）：只输出"改稿新增"，即正文中有、台账中没有的事实；按集合比较，同一个数在正文中
用两次不报；台账中 Markdown 表格表头为"序号""来源"的列不参与比对。

比对内容（只比对【起草说明】之前的正文）：
    数字：阿拉伯数字和带量词的汉字数字，连同单位一起比对，如"3.5天""12.5%""2个百分点"；
          汉字数字换算成阿拉伯数字后比对，"三家"与"3家"视为相同；"三分之一"按"1/3"比对；
          数值范围中前一个数省略的单位按后一个数补全，"3～5万元"与"3万～5万元"视为相同；
    日期：如"2026年10月16日""2026-10-16"；时刻如"9:30"；月、日的虚位 0 不计，"11月05日"与"11月5日"视为相同；
    标记：【待补：……】【待核实：……】；
    书名号：《……》整体比对；
    引号：“……”‘……’「……」『……』及成对的英文双引号，整体比对；
    发文字号：如"某政发〔2026〕12号"。

不算事实、不参与比对的内容：
    层次序号 一、（一）1.（1），段内的 1. 1、1）（1）列举序号，作列举用的"第一，""第二、"；
    单独的汉字"一"加量词（如"一个""一批"），以及不带量词的汉字数字（如"统一""一是"）。

输出两张清单："原稿有而改稿没有""改稿新增"。按出现次数比对：原稿出现 2 次、改稿 1 次，也会列出。

退出码：0 正常；--strict 时任一清单非空返回 1；输入错误（文件不存在、格式不支持、读取失败）返回 2。
只依赖标准库；读取 .docx 需要 python-docx，未安装时给出安装提示并返回 2。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import Counter, OrderedDict
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

BOUNDARY = "【起草说明】"
# 分界行：去掉行首空白后以【起草说明】开头；兼容 Markdown 标题、加粗包裹的写法。
# 与 gongwen_lint.py 的 DRAFT_NOTE_RE、gongwen-format 的 gw_layout.is_draft_note_line 保持一致。
DRAFT_NOTE_RE = re.compile(r"^\s*(?:#{1,6}\s*)?(?:(?:\*\*|__)\s*)?" + re.escape(BOUNDARY))
SUPPORTED_SUFFIXES = (".txt", ".md", ".docx")
MASK = "\ue000"  # 私用区字符，用于屏蔽已识别的片段，长度不变，便于回溯原文位置
MAX_LOCATIONS = 3
EXCERPT_RADIUS = 14
VALUE_DISPLAY_LIMIT = 60

# 类型名称（输出用）
T_NUMBER = "数字"
T_DATE = "日期"
T_MARKER = "标记"
T_TITLE = "书名号"
T_QUOTE = "引号"
T_DOCNO = "发文字号"
TYPE_ORDER = (T_NUMBER, T_DATE, T_MARKER, T_TITLE, T_QUOTE, T_DOCNO)

# 全角数字和符号转半角（不用 NFKC，避免把 ① 之类的序号转成数字）
FULLWIDTH_TABLE = str.maketrans("０１２３４５６７８９％．", "0123456789%.")

# ---------------------------------------------------------------------------
# 正则
# ---------------------------------------------------------------------------

MARKER_RE = re.compile(r"【(?:待补|待核实)(?:[：:][^】\n]*)?】")
DOCNO_RE = re.compile(
    r"([\u4e00-\u9fffA-Za-z]{0,12})"
    r"([〔\[［(（【])(\d{4})([〕\]］)）】])第?(\d{1,5})号"
)
# 发文字号前面常见的非代字成分，截掉最后一次出现及其之前的部分，使改写上下文不影响比对。
# 单字只收不会出现在机关代字中的介词、连词（按、依、据、以、在、为、是、由、从、见、即、并、把）
DOCNO_PREFIX_CUT_RE = re.compile(
    r".*(?:根据|按照|依据|依照|参照|对照|遵照|收到|接到|贯彻|落实|学习|传达|印发|转发|执行|你局|你委|你厅|你办|你院|你校|"
    r"你公司|你单位|贵单位|贵局|贵委|贵厅|我局|我委|我厅|我院|我校|本局|本院|关于|文件|的|了|和|及|与|将|"
    r"按|依|据|以|在|为|是|由|从|见|即|并|把)"
)
TITLE_RE = re.compile(r"《[^《》\n]{1,200}》")
QUOTE_RES = (
    re.compile(r"“[^“”\n]{1,300}”"),
    re.compile(r"「[^「」\n]{1,300}」"),
    re.compile(r"『[^『』\n]{1,300}』"),
    re.compile(r"‘[^‘’\n]{1,300}’"),
    re.compile(r"\"[^\"\n]{1,300}\""),
)

# 层次序号与列举序号
MD_PREFIX = r"(?:#{1,6}\s*|[-*+]\s+|>\s*)?(?:\*\*)?"
SEQ_LINE_START_RE = re.compile(
    r"^\s*" + MD_PREFIX + r"(?:[一二三四五六七八九十]{1,3}[、.．]|"
    r"[（(][一二三四五六七八九十]{1,3}[）)]|"
    r"\d{1,3}[.．、](?!\d)|"
    r"[（(]\d{1,3}[）)]|"
    r"\d{1,3}[)）])"
)
SEQ_INLINE_RES = (
    re.compile(r"[（(][一二三四五六七八九十]{1,3}[）)]"),
    re.compile(r"[（(]\d{1,2}[）)]"),
    re.compile(r"(?:(?<=^)|(?<=[：:；;。！？!?，,\s\ue000]))\d{1,2}[.．、)）](?!\d)"),
    # 作列举用的序数词："第一，""第二、""第三：""其一，"
    re.compile(r"第[一二三四五六七八九十\d]{1,3}(?=[，,、：:；;])"),
)

DATE_RES = (
    re.compile(r"(?<!\d)\d{4}年\d{1,2}月(?:\d{1,2}[日号])?"),
    re.compile(r"(?<!\d)\d{1,2}月\d{1,2}[日号]"),
    re.compile(r"(?<![\d.])\d{4}[-/.]\d{1,2}[-/.]\d{1,2}(?![\d.])"),
)
TIME_RE = re.compile(r"(?<![\d.])(\d{1,2})[:：](\d{2})(?![\d])")
# 日期中月、日的虚位 0（年份中的 0 前面是数字，不受影响）
LEADING_ZERO_RE = re.compile(r"(?<!\d)0+(?=\d)")
# 数值范围的连接号：前一个数没写单位或"万""亿"时，沿用后一个数的（"3～5万元"与"3万～5万元"视为相同）
RANGE_SEP_RE = re.compile(r"\s*(?:[～~〜－\-]|—{1,2}|至|到)\s*")
# 汉字书写的日期，如"二〇二六年十一月十六日""十月十六日"，换算成阿拉伯数字后与阿拉伯数字日期统一比对
CN_DATE_RES = (
    re.compile(r"([〇零一二三四五六七八九]{4})年([一二三四五六七八九十]{1,3})月(?:([一二三四五六七八九十]{1,3})[日号])?"),
    re.compile(r"(?<![〇零一二三四五六七八九十])()([一二三四五六七八九十]{1,3})月([一二三四五六七八九十]{1,3})[日号]"),
)

CN_DIGITS = "零〇一二两三四五六七八九"
CN_SMALL_UNITS = "十百千"
CN_BIG_UNITS = "万亿"
CN_NUM_CHARS = CN_DIGITS + CN_SMALL_UNITS
PERCENT_CN_RE = re.compile(r"百分之([" + CN_NUM_CHARS + r"]+)")
# 汉字分数“三分之一”“千分之五”（“百分之”由 PERCENT_CN_RE 先行处理）
FRACTION_CN_RE = re.compile(r"([" + CN_NUM_CHARS + r"万]+)分之([" + CN_NUM_CHARS + r"]+)")
_FRACTION_UNIT_DENOM = {"十": 10, "百": 100, "千": 1000, "万": 10000}

ARABIC_RE = re.compile(
    r"(?<![\d.])(第)?(\d+(?:,\d{3})+|\d+)(\.\d+)?(余|多)?([十百千万亿]{1,2})?(余|多)?"
)
CN_RE = re.compile(
    r"(第)?([" + CN_NUM_CHARS + r"]+(?:[万亿][" + CN_NUM_CHARS + r"]+)*)(余|多)?([万亿]{1,2})?(余|多)?"
)

# 单位：先匹配多字单位，再匹配单字单位
MULTI_UNITS = (
    "个百分点", "百分点", "个工作日", "工作日", "个月", "万人次", "人次", "例次", "台次", "户次", "架次",
    "平方公里", "平方千米", "平方米", "立方米", "公里", "千米", "公顷", "千克", "公斤", "小时", "分钟",
    "课时", "学时", "学分", "季度", "周岁", "周年",
)
SINGLE_UNITS = frozenset(
    "%‰个项次家名人位所件条批期轮场起篇份台辆处座户村镇张本床例套栋间册门节"
    "年月日号时秒天周岁倍元角吨亩米层级类种支队组届点股斤箱瓶块段部章成"
)
# 含数字但不表示数量的成语和惯用语，抽取汉字数字前先屏蔽
CN_IDIOMS = ("第一时间", "千家万户", "百家争鸣", "三番五次", "三天两头", "五花八门", "一五一十", "三令五申")
LATIN_UNIT_RE = re.compile(r"[A-Za-z]{1,3}[²³]?(?![A-Za-z])")


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------

class Fact:
    __slots__ = ("kind", "value", "line", "pos", "excerpt")

    def __init__(self, kind: str, value: str, line: int, pos: int, excerpt: str) -> None:
        self.kind = kind
        self.value = value
        self.line = line
        self.pos = pos
        self.excerpt = excerpt

    @property
    def key(self) -> Tuple[str, str]:
        return (self.kind, self.value)


class InputError(Exception):
    """输入错误，对应退出码 2。"""


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
                         "或把正文另存为 .txt 后再比对。")
    try:
        document = docx.Document(str(path))
    except Exception as exc:  # 文件损坏、不是 docx 等
        raise InputError(f"无法打开 docx 文件：{path}（{exc}）")
    lines: List[str] = []
    body = document.element.body
    for child in body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            lines.append(Paragraph(child, document).text)
        elif tag == "tbl":
            # python-docx 对横向、纵向合并的单元格逐格重复返回同一个单元格，按底层 tc 元素去重，避免一处数字计成多处
            seen = set()
            for row in Table(child, document).rows:
                texts = []
                for cell in row.cells:
                    if cell._tc in seen:
                        continue
                    seen.add(cell._tc)
                    texts.append(cell.text.replace("\n", " "))
                lines.append("\t".join(texts))
    return "\n".join(lines)


def _long_path(path: Path) -> Path:
    """Windows 下完整路径接近 260 字符上限时加长路径前缀，避免"文件不存在"误报。"""
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
    """读取 .txt、.md、.docx 文件的全文。"""
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
        return _read_docx(path)
    try:
        data = path.read_bytes()
    except OSError as exc:
        raise InputError(f"读取失败：{path}（{exc}）")
    return _decode(data, path)


TABLE_ROW_RE = re.compile(r"^\s*\|.*\|\s*$")
TABLE_SEP_RE = re.compile(r"^\s*\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)*\|?\s*$")
LEDGER_SKIP_COLUMNS = ("序号", "来源")


def _table_cells(row: str) -> List[str]:
    return [cell.strip() for cell in row.strip().strip("|").split("|")]


def ledger_lines(lines: Sequence[str]) -> List[str]:
    """台账对账用：去掉 Markdown 表格中表头为"序号""来源"的列，其余单元格用制表符连接。

    分隔行（|---|）变为空行，行数不变，便于按行号定位。表格以外的行原样保留。
    """
    result: List[str] = []
    skip: Optional[set] = None  # 当前表格中要去掉的列号；None 表示不在表格中
    for i, line in enumerate(lines):
        if not TABLE_ROW_RE.match(line):
            skip = None
            result.append(line)
            continue
        if TABLE_SEP_RE.match(line):
            result.append("")
            continue
        cells = _table_cells(line)
        if skip is None:  # 表头：下一行是分隔行
            nxt = lines[i + 1] if i + 1 < len(lines) else ""
            names = [re.sub(r"[\s*_]", "", c) for c in cells]
            skip = {k for k, name in enumerate(names) if name in LEDGER_SKIP_COLUMNS} \
                if TABLE_SEP_RE.match(nxt) else set()
        result.append("\t".join(c for k, c in enumerate(cells) if k not in skip))
    return result


def body_lines(text: str) -> List[str]:
    """返回【起草说明】之前的正文行。"""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines: List[str] = []
    for line in text.split("\n"):
        if DRAFT_NOTE_RE.match(line):
            break
        lines.append(line)
    return lines


# ---------------------------------------------------------------------------
# 汉字数字
# ---------------------------------------------------------------------------

_CN_DIGIT_VALUE = {"零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
                   "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
_CN_SMALL_VALUE = {"十": 10, "百": 100, "千": 1000}
_CN_BIG_VALUE = {"万": 10 ** 4, "亿": 10 ** 8}


def cn_to_int(text: str) -> Optional[int]:
    """把汉字数字转成整数，不合法时返回 None。

    支持"十五""三百零五""一万二千"，以及逐位书写的年份"二〇二六"。
    "三四"（概数）、"十四五"之类不是确定数值，返回 None。
    """
    if not text:
        return None
    if all(ch in _CN_DIGIT_VALUE for ch in text):
        if len(text) == 1:
            return _CN_DIGIT_VALUE[text]
        if "〇" in text or "零" in text or len(text) == 4:
            if "两" in text:
                return None
            return int("".join(str(_CN_DIGIT_VALUE[ch]) for ch in text))
        return None
    no_unit = 10 ** 9
    total = 0
    section = 0
    number: Optional[int] = None
    prev = ""
    last_small_unit = no_unit
    for ch in text:
        if ch in _CN_DIGIT_VALUE:
            if number is not None and prev not in ("零", "〇"):
                return None  # 连续数字，如"三四"（概数）
            number = _CN_DIGIT_VALUE[ch]
        elif ch in _CN_SMALL_VALUE:
            unit = _CN_SMALL_VALUE[ch]
            if unit >= last_small_unit:
                return None  # 单位不递减，如"十百""千千"
            if number is None:
                if ch != "十":
                    return None
                number = 1  # "十五"中的"十"
            section += number * unit
            number = None
            last_small_unit = unit
        elif ch in _CN_BIG_VALUE:
            if number is not None:
                section += number
            if section == 0:
                return None
            total += section * _CN_BIG_VALUE[ch]
            section = 0
            number = None
            last_small_unit = no_unit
        else:
            return None
        prev = ch
    if number is not None:
        before = text[-2] if len(text) >= 2 else ""
        if before in ("百", "千"):
            number = number * _CN_SMALL_VALUE[before] // 10  # "三百五"按 350
        elif before in _CN_BIG_VALUE:
            number = number * _CN_BIG_VALUE[before] // 10  # "一万五"按 15000
        section += number
    return total + section


# ---------------------------------------------------------------------------
# 抽取
# ---------------------------------------------------------------------------

def _excerpt(line: str, start: int, end: int) -> str:
    left = max(0, start - EXCERPT_RADIUS)
    right = min(len(line), end + EXCERPT_RADIUS)
    text = line[left:right].strip()
    return ("……" if left > 0 else "") + text + ("……" if right < len(line) else "")


def _mask(buf: List[str], start: int, end: int) -> None:
    for i in range(start, end):
        buf[i] = MASK


def _norm_number(int_part: str, frac: Optional[str]) -> str:
    int_part = int_part.replace(",", "")
    int_part = int_part.lstrip("0") or "0"
    if frac:
        frac_digits = frac[1:].rstrip("0")
        if frac_digits:
            return f"{int_part}.{frac_digits}"
    return int_part


def _match_unit(text: str, pos: int) -> str:
    """从 pos 开始匹配单位，返回单位文字（可能为空）。允许数字与 % 之间有一个空格。"""
    probe = pos
    if probe < len(text) and text[probe] == " " and probe + 1 < len(text) and text[probe + 1] in "%‰":
        probe += 1
    for unit in MULTI_UNITS:
        if text.startswith(unit, probe):
            return text[pos:probe] + unit
    if probe < len(text) and text[probe] in SINGLE_UNITS:
        return text[pos:probe] + text[probe]
    match = LATIN_UNIT_RE.match(text, pos)
    if match:
        return match.group(0)
    return ""


def _scan_spans(regex: "re.Pattern[str]", masked: str,
                allow_inner_mask: bool = False) -> List[Tuple[int, int, "re.Match[str]"]]:
    """在已屏蔽的文本中查找匹配。默认跳过含屏蔽字符的匹配；引号、书名号允许内部含已识别片段。"""
    spans = []
    for m in regex.finditer(masked):
        text = m.group(0)
        if MASK in text:
            if not allow_inner_mask or text[0] == MASK or text[-1] == MASK:
                continue
        spans.append((m.start(), m.end(), m))
    return spans


def _docno_value(match: "re.Match[str]") -> Tuple[str, int]:
    """返回 (发文字号文本, 起始位置偏移)，截掉代字前面的动词、介词等。"""
    prefix = match.group(1)
    cut = DOCNO_PREFIX_CUT_RE.match(prefix)
    offset = cut.end() if cut else 0
    return match.group(0)[offset:], offset


def extract_line(line: str, lineno: int) -> List[Fact]:
    """抽取一行正文中的事实。"""
    facts: List[Fact] = []
    original = line
    norm = line.translate(FULLWIDTH_TABLE)
    buf = list(norm)

    def current() -> str:
        return "".join(buf)

    def add(kind: str, value: str, start: int, end: int) -> None:
        facts.append(Fact(kind, value, lineno, start, _excerpt(original, start, end)))

    # 1. 标记
    for start, end, m in _scan_spans(MARKER_RE, current()):
        add(T_MARKER, original[start:end], start, end)
        _mask(buf, start, end)
    # 2. 发文字号
    for start, end, m in _scan_spans(DOCNO_RE, current()):
        value, offset = _docno_value(m)
        add(T_DOCNO, value, start + offset, end)
        _mask(buf, start + offset, end)
    # 3. 书名号（内部可含已识别的发文字号、标记）
    for start, end, m in _scan_spans(TITLE_RE, current(), allow_inner_mask=True):
        add(T_TITLE, original[start:end], start, end)
        _mask(buf, start, end)
    # 4. 引号（内部可含已识别的书名号等）
    for regex in QUOTE_RES:
        for start, end, m in _scan_spans(regex, current(), allow_inner_mask=True):
            add(T_QUOTE, original[start:end], start, end)
            _mask(buf, start, end)
    # 5. 层次序号和列举序号（不算事实，只屏蔽）
    m = SEQ_LINE_START_RE.match(current())
    if m:
        _mask(buf, m.start(), m.end())
    for regex in SEQ_INLINE_RES:
        for start, end, _m in _scan_spans(regex, current()):
            _mask(buf, start, end)
    # 6. 日期、时刻（"16号"与"16日"视为相同）
    for regex in DATE_RES:
        for start, end, m in _scan_spans(regex, current()):
            value = m.group(0)
            if value.endswith("号"):
                value = value[:-1] + "日"
            value = LEADING_ZERO_RE.sub("", value)  # "11月05日"与"11月5日"视为相同
            add(T_DATE, value, start, end)
            _mask(buf, start, end)
    for regex in CN_DATE_RES:
        for start, end, m in _scan_spans(regex, current()):
            year, month, day = m.group(1), m.group(2), m.group(3)
            y = cn_to_int(year) if year else None
            mo = cn_to_int(month)
            d = cn_to_int(day) if day else None
            if mo is None or not 1 <= mo <= 12 or (year and y is None) or (day and (d is None or not 1 <= d <= 31)):
                continue
            value = (f"{y}年" if year else "") + f"{mo}月" + (f"{d}日" if day else "")
            add(T_DATE, value, start, end)
            _mask(buf, start, end)
    for start, end, m in _scan_spans(TIME_RE, current()):
        add(T_DATE, f"{int(m.group(1))}:{m.group(2)}", start, end)
        _mask(buf, start, end)
    # 7. "百分之三十""百分之百"
    for start, end, m in _scan_spans(PERCENT_CN_RE, current()):
        value = 100 if m.group(1) == "百" else cn_to_int(m.group(1))
        if value is not None:
            add(T_NUMBER, f"{value}%", start, end)
            _mask(buf, start, end)
    # 7.1 "三分之一""千分之五"：按"分子/分母"比对，千分之按 ‰
    for start, end, m in _scan_spans(FRACTION_CN_RE, current()):
        denom = _FRACTION_UNIT_DENOM.get(m.group(1)) or cn_to_int(m.group(1))
        numer = cn_to_int(m.group(2))
        if not denom or numer is None:
            continue
        add(T_NUMBER, f"{numer}‰" if denom == 1000 else f"{numer}/{denom}", start, end)
        _mask(buf, start, end)
    # 8. 阿拉伯数字
    text = current()
    numbers = []  # [match, 倍数（万、亿）, 单位, 结束位置]
    for m in ARABIC_RE.finditer(text):
        if MASK in m.group(0):
            continue
        unit = _match_unit(text, m.end())
        numbers.append([m, m.group(5) or "", unit, m.end() + len(unit)])
    # 数值范围：前一个数缺倍数或单位的，沿用后一个数的（只改写法、数值不变的不报差异）
    for cur, nxt in zip(numbers, numbers[1:]):
        if not cur[2] and cur[3] < nxt[0].start() and RANGE_SEP_RE.fullmatch(text, cur[3], nxt[0].start()):
            cur[1] = cur[1] or nxt[1]
            cur[2] = nxt[2]
    for m, mult, unit, end in numbers:
        ordinal, int_part, frac, approx1, _mult, approx2 = m.groups()
        value = (ordinal or "") + _norm_number(int_part, frac) + (approx1 or "") + mult + \
            (approx2 or "") + unit.replace(" ", "")
        add(T_NUMBER, value, m.start(), end)
        _mask(buf, m.start(), end)
    # 9. 带量词的汉字数字（先屏蔽成语）
    for idiom in CN_IDIOMS:
        for m in re.finditer(re.escape(idiom), current()):
            _mask(buf, m.start(), m.end())
    text = current()
    for m in CN_RE.finditer(text):
        ordinal, digits, approx1, mult, approx2 = m.groups()
        number = cn_to_int(digits)
        if number is None:
            continue
        unit = _match_unit(text, m.end())
        if not ordinal:
            if not unit or LATIN_UNIT_RE.fullmatch(unit):
                continue  # 不带量词的汉字数字（统一、一是、十四五）不算
            if digits == "一" and not (approx1 or mult or approx2):
                continue  # "一个""一批""一次"多为泛指
        value = (ordinal or "") + str(number) + (approx1 or "") + (mult or "") + (approx2 or "") + \
            unit.replace(" ", "")
        end = m.end() + len(unit)
        add(T_NUMBER, value, m.start(), end)
        _mask(buf, m.start(), end)
    facts.sort(key=lambda f: f.pos)
    return facts


def extract_facts(lines: Sequence[str]) -> List[Fact]:
    facts: List[Fact] = []
    for index, line in enumerate(lines, start=1):
        if line.strip():
            facts.extend(extract_line(line, index))
    return facts


# ---------------------------------------------------------------------------
# 比对
# ---------------------------------------------------------------------------

def _diff_entries(base: List[Fact], base_counts: Counter, other_counts: Counter,
                  base_is_original: bool) -> List[dict]:
    """列出 base 中出现次数多于另一稿的项（按首次出现顺序），位置取自 base。"""
    grouped: "OrderedDict[Tuple[str, str], List[Fact]]" = OrderedDict()
    for fact in base:
        grouped.setdefault(fact.key, []).append(fact)
    entries: List[dict] = []
    for key, items in grouped.items():
        mine, theirs = base_counts[key], other_counts.get(key, 0)
        if mine <= theirs:
            continue
        kind, value = key
        entries.append({
            "type": kind,
            "value": value,
            "original_count": mine if base_is_original else theirs,
            "revised_count": theirs if base_is_original else mine,
            "locations": [{"line": f.line, "excerpt": f.excerpt} for f in items[:MAX_LOCATIONS]],
        })
    return entries


def compare(original: List[Fact], revised: List[Fact]) -> Tuple[List[dict], List[dict]]:
    """返回 (原稿有而改稿没有, 改稿新增)。按出现次数比对。"""
    orig_counts = Counter(f.key for f in original)
    rev_counts = Counter(f.key for f in revised)
    missing = _diff_entries(original, orig_counts, rev_counts, base_is_original=True)
    added = _diff_entries(revised, rev_counts, orig_counts, base_is_original=False)
    return missing, added


def compare_ledger(ledger: List[Fact], body: List[Fact]) -> List[dict]:
    """台账对账：返回正文中有、台账中没有的事实。按集合比较，同一事实在正文中用几次都只看有没有。"""
    known = {f.key for f in ledger}
    body_counts = Counter(f.key for f in body)
    grouped: "OrderedDict[Tuple[str, str], List[Fact]]" = OrderedDict()
    for fact in body:
        if fact.key not in known:
            grouped.setdefault(fact.key, []).append(fact)
    return [{
        "type": kind,
        "value": value,
        "original_count": 0,
        "revised_count": body_counts[(kind, value)],
        "locations": [{"line": f.line, "excerpt": f.excerpt} for f in items[:MAX_LOCATIONS]],
    } for (kind, value), items in grouped.items()]


def _type_counts(facts: List[Fact]) -> Dict[str, int]:
    counts = Counter(f.kind for f in facts)
    return {kind: counts.get(kind, 0) for kind in TYPE_ORDER}


def build_report(original_path: Path, revised_path: Path, ledger: bool = False) -> dict:
    orig_lines = body_lines(read_text(original_path))
    rev_lines = body_lines(read_text(revised_path))
    if ledger:
        orig_lines = ledger_lines(orig_lines)
    orig_facts = extract_facts(orig_lines)
    rev_facts = extract_facts(rev_lines)
    if ledger:
        added = compare_ledger(orig_facts, rev_facts)
        return {
            "mode": "ledger",
            "original_file": original_path.name,
            "revised_file": revised_path.name,
            "scope": "【起草说明】之前的正文；台账表格的“序号”“来源”列不参与比对",
            "original_lines": len(orig_lines),
            "revised_lines": len(rev_lines),
            "fact_counts": {"original": _type_counts(orig_facts), "revised": _type_counts(rev_facts)},
            "added": added,
            "consistent": not added,
        }
    missing, added = compare(orig_facts, rev_facts)
    return {
        "original_file": original_path.name,
        "revised_file": revised_path.name,
        "scope": "【起草说明】之前的正文",
        "original_lines": len(orig_lines),
        "revised_lines": len(rev_lines),
        "fact_counts": {"original": _type_counts(orig_facts), "revised": _type_counts(rev_facts)},
        "missing": missing,
        "added": added,
        "consistent": not missing and not added,
    }


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------

def _short(value: str) -> str:
    if len(value) <= VALUE_DISPLAY_LIMIT:
        return value
    return value[:VALUE_DISPLAY_LIMIT] + "……"


def _render_list(title: str, entries: List[dict], where: str,
                 names: Tuple[str, str] = ("原稿", "改稿")) -> List[str]:
    lines = [f"{title}（{len(entries)} 项）"]
    if not entries:
        lines.append("无")
        return lines
    for index, item in enumerate(entries, start=1):
        lines.append(f"{index}. [{item['type']}] {_short(item['value'])}"
                     f"（{names[0]} {item['original_count']} 处，{names[1]} {item['revised_count']} 处）")
        for loc in item["locations"]:
            lines.append(f"   {where}第 {loc['line']} 行：{loc['excerpt']}")
    return lines


def render_ledger_text(report: dict) -> str:
    counts_o = report["fact_counts"]["original"]
    counts_r = report["fact_counts"]["revised"]
    summary = "，".join(f"{k} {counts_o[k]}/{counts_r[k]}" for k in TYPE_ORDER)
    lines = [
        f"台账对账：台账 {report['original_file']}，正文 {report['revised_file']}",
        f"比对范围：{report['scope']}（台账 {report['original_lines']} 行，正文 {report['revised_lines']} 行）",
        f"识别到的事实项（台账/正文）：{summary}",
        "",
    ]
    lines += _render_list("正文中有、台账中没有", report["added"], "正文", ("台账", "正文"))
    lines.append("")
    if report["added"]:
        lines.append(f"结论：{len(report['added'])} 项不在台账中。逐项处理：补进台账并写明出处，"
                     "或在正文中改为【待补：说明】；新加的【待补】【待核实】标记核对无误即可。")
    else:
        lines.append("结论：正文中的数字、日期、标记等均能在台账中找到。")
    return "\n".join(lines)


def render_text(report: dict) -> str:
    if report.get("mode") == "ledger":
        return render_ledger_text(report)
    counts_o = report["fact_counts"]["original"]
    counts_r = report["fact_counts"]["revised"]
    summary = "，".join(f"{k} {counts_o[k]}/{counts_r[k]}" for k in TYPE_ORDER)
    lines = [
        f"事实一致性检查：原稿 {report['original_file']}，改稿 {report['revised_file']}",
        f"比对范围：{report['scope']}（原稿 {report['original_lines']} 行，改稿 {report['revised_lines']} 行）",
        f"识别到的事实项（原稿/改稿）：{summary}",
        "",
    ]
    lines += _render_list("一、原稿有而改稿没有", report["missing"], "原稿")
    lines.append("")
    lines += _render_list("二、改稿新增", report["added"], "改稿")
    lines.append("")
    total = len(report["missing"]) + len(report["added"])
    if total:
        lines.append(f"结论：发现差异 {total} 项。逐项核对：属于事实改动的，恢复原稿写法；"
                     "确属表达调整的（如去掉强调用的引号、删去复述句后重复的数字少了一处、按保密要求脱敏），"
                     "在修改说明中写明。")
    else:
        lines.append("结论：未发现事实差异。人名、单位名、承诺事项不在本脚本比对范围内，请人工核对。")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

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
        prog="fact_diff.py",
        description="比对改稿前后的数字、日期、【待补】【待核实】标记、书名号和引号内容、发文字号，"
                    "确认改稿没有改动事实。只比对【起草说明】之前的正文。",
        epilog="退出码：0 正常；1 使用 --strict 且发现差异；2 输入错误。",
        formatter_class=_ChineseHelpFormatter,
        add_help=False,
    )
    parser._optionals.title = "选项"
    parser._positionals.title = "参数"
    parser.add_argument("-h", "--help", action="help", help="显示本帮助信息并退出")
    parser.add_argument("original", help="原稿文件（.txt、.md、.docx）")
    parser.add_argument("revised", help="改稿文件（.txt、.md、.docx）")
    parser.add_argument("--json", action="store_true", help="以 JSON 格式输出")
    parser.add_argument("--strict", action="store_true", help="任一清单非空时返回 1")
    parser.add_argument("--ledger", action="store_true",
                        help="台账对账：第一个文件为素材台账、第二个为正文，只列正文中有、台账中没有的事实；"
                             "按是否出现比较，不计次数；台账表格的“序号”“来源”列不参与比对")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    _setup_console()
    parser = make_parser()
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:
        return 0 if exc.code == 0 else 2
    try:
        report = build_report(Path(args.original), Path(args.revised), ledger=args.ledger)
    except InputError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(render_text(report))
    if args.strict and not report["consistent"]:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
