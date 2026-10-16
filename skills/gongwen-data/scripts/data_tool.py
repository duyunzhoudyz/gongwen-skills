#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""公文数据整理工具：报表体检、多表合并、指标计算、数据转文字。

用法：
  python scripts/data_tool.py inspect <表格> [--sheet 名称 | --all-sheets] [--header-row N] [--key 列名] [--json] [--strict]
  python scripts/data_tool.py merge <表格> [<表格>...] [--all-sheets] [--key 列名] [--layout auto|record|indicator]
                                    [--expect 名单] [--out 汇总.xlsx] [--json] [--strict]
  python scripts/data_tool.py calc --type 类型 --current 本期 --base 基期 [其他参数] [--json]
  python scripts/data_tool.py calc --type 类型 --table <表格> --name-col 列名 --current-col 列名 --base-col 列名 [--json]
  python scripts/data_tool.py phrase （参数同 calc，另输出可直接写进公文的句子） [--md]
  python scripts/data_tool.py count <表格> --by 列名[,列名] [--unit 件] [--json]

输入：.csv（自动识别 UTF-8、GB18030）和 .xlsx（需要 openpyxl）。不支持 .xls，请在 Excel 或 WPS 中另存为 .xlsx。
约定：
  - 不联网，不调用任何外部服务，不修改输入文件；--out 写出的文件已存在时加 -v2、-v3，不覆盖。
  - 输出中的姓名、身份证号、手机号等个人信息一律脱敏。
  - 计算一律用十进制精确运算，四舍五入；每个结果附算式，便于复核。
  - 材料带密级或内部标注（秘密★、内部资料、不得外传等）时停止处理（DT-SEC-01）：inspect 只报标注位置，
    其他子命令拦截并退出；用户说明可以处理后，加 --marking-ok 重新运行。
退出码：0 正常；1 --strict 时存在 P0 问题；2 输入错误（文件不存在、格式不支持、参数错误、缺少 openpyxl、
       材料带密级或内部标注而未加 --marking-ok）。

计算类型（--type）：
  growth 增速（同比、环比等，用 --basis 指明）    change 增减量          pp 百分点
  share 占比     pull 拉动（百分点）     contrib 贡献率     completion 完成率（可配时间进度）
  multiple 倍数和番数     cagr 年均增速     mean 简单平均     wmean 加权平均     ratio 比值
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import math
import re
import sys
from collections import Counter, OrderedDict
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

VERSION = "1.0.0"

LEVELS = ("P0", "P1", "P2")
MAX_LIST = 10  # 同类问题在输出中最多列出的位置数

# ---------------------------------------------------------------------------
# 词表
# ---------------------------------------------------------------------------

# 伪装成数值的空值
NULL_TOKENS = {
    "-", "--", "—", "——", "―", "－", "/", "／", "\\", "n/a", "na", "n.a.", "null", "none", "nan",
    "无", "空", "暂无", "未知", "不适用", "待定", "?", "？", "…", "...", "×", "x",
}
ERROR_RE = re.compile(r"^#(?:N/A|VALUE!|DIV/0!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!)$", re.I)
NOTE_PREFIXES = (
    "注：", "注:", "注 ", "备注", "说明", "数据来源", "资料来源", "来源：", "来源:", "制表", "填表",
    "填报人", "审核", "负责人", "联系人", "联系电话", "报送", "*", "※",
)
TOTAL_WORDS = ("合计", "总计", "总和", "汇总", "共计", "总数")
TOTAL_SCOPE_RE = re.compile(r"^全(?:院|市|县|区|省|镇|乡|街道|校|公司|集团|局|系统|行业|部)(?:合计|总计)?$")
SUBTOTAL_WORDS = ("小计",)
OF_WHICH_RE = re.compile(r"^(?:其中|内[：:]|#)")
AVG_ROW_RE = re.compile(r"^(?:全\S{0,4})?(?:平均|均值)$")


def is_summary_label(label: str) -> bool:
    """合计、小计、平均行的标签（如"合计""全院合计""内科小计"）。"""
    key = re.sub(r"[\s　]+", "", str(label))
    return bool(key) and (any(w in key for w in SUBTOTAL_WORDS) or bool(AVG_ROW_RE.match(key))
                          or (len(key) <= 12 and (any(w in key for w in TOTAL_WORDS) or bool(TOTAL_SCOPE_RE.match(key)))))

RATE_NAME_RE = re.compile(r"率|占比|比重|比例|百分比|%|％|增幅|增速|同比|环比|构成")
AVG_NAME_RE = re.compile(r"平均|均值|人均|次均|例均|日均|床均|户均|^均")
ID_NAME_RE = re.compile(r"序号|编号|编码|代码|号码|工号|单号|卡号|证号|邮编|电话|手机|排名|名次|位次|排序|"
                        r"病案号|住院号|门诊号|就诊号|身份证|学号|学籍号|账号")
# 由列名推断计量单位：列名以单位结尾，或"××数"
NAME_UNIT_SUFFIX = (("床日数", "床日"), ("人次", "人次"), ("台次", "台次"), ("例次", "例次"), ("床日", "床日"), ("人数", "人"),
                    ("例数", "例"), ("台数", "台"), ("件数", "件"), ("次数", "次"), ("户数", "户"), ("个数", "个"),
                    ("家数", "家"), ("项数", "项"), ("篇数", "篇"), ("天数", "天"))
PERIOD_NAME_RE = re.compile(r"^(?:年份|年度|月份|季度|年|月|日期|时间|统计时间|统计时段)$")
UNIT_COL_RE = re.compile(r"^(?:计量)?单位$")

# 个人信息
SENSITIVE_NAME_RE = re.compile(
    r"姓名|患者|病人|病员|学生姓名|身份证|证件号|手机|电话|联系方式|病案号|住院号|门诊号|就诊号|学号|学籍号|"
    r"银行卡|账号|住址|家庭地址|居住地址"
)
NAME_COL_RE = re.compile(r"姓名|患者|病人|病员|联系人")
ID_CARD_RE = re.compile(r"(?<!\d)\d{17}[\dXx](?!\d)")
MOBILE_RE = re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")

# 密级和内部标注（guardrails 第一条）：文件名、工作表名、标题、表头前文字、单元格、表注中出现即停止处理。
# "内部资料""内部使用"后面紧跟汉字的是项目名称（如"内部资料印刷费"），不算标注；"保密工作经费"不命中。
MARKING_RES = (
    re.compile(r"(?:绝密|机密|秘密)\s*[★☆*＊]"),
    re.compile(r"密级\s*[:：]?\s*(?:绝密|机密|秘密)"),
    re.compile(r"工作秘密"),
    re.compile(r"不得外传"),
    re.compile(r"仅供内部"),
    re.compile(r"(?<!是否)内部(?:资料|使用)(?:(?![一-鿿])|(?=不得|注意|请勿|严禁|禁止|勿))"),
)

# 统计时段
PERIOD_RE = re.compile(
    r"(?:19|20)\d{2}年(?:度)?"
    r"(?:(?:第?[一二三四1-4]季度|[上下]半年|(?:\d{1,2}|[一二三四五六七八九十]{1,3})(?:[—\-－~～至到](?:\d{1,2}|[一二三四五六七八九十]{1,3}))?月(?:份)?|前[一二三]季度))?"
    r"(?:[春秋]季学期|第[一二]学期)?"
    r"|(?:19|20)\d{2}[—\-－~～至到](?:19|20)?\d{2}学年(?:第[一二]学期)?"
)

# 计量单位：同一量纲内可换算（换算到基础单位的倍数）
UNIT_FACTORS: Dict[str, Tuple[str, Decimal]] = {}


def _reg(dim: str, pairs: Sequence[Tuple[str, str]]) -> None:
    for unit, factor in pairs:
        UNIT_FACTORS[unit] = (dim, Decimal(factor))


_reg("金额", [("元", "1"), ("千元", "1000"), ("万元", "10000"), ("百万元", "1000000"), ("亿元", "100000000")])
_reg("人", [("人", "1"), ("万人", "10000")])
_reg("人次", [("人次", "1"), ("万人次", "10000")])
_reg("次", [("次", "1"), ("万次", "10000")])
_reg("件", [("件", "1"), ("万件", "10000")])
_reg("户", [("户", "1"), ("万户", "10000")])
_reg("个", [("个", "1"), ("万个", "10000")])
_reg("台", [("台", "1"), ("台次", "1")])
_reg("例", [("例", "1"), ("例次", "1")])
_reg("重量", [("千克", "1"), ("公斤", "1"), ("吨", "1000"), ("万吨", "10000000")])
_reg("面积", [("平方米", "1"), ("万平方米", "10000"), ("公顷", "10000"), ("亩", "666.6666666667"), ("万亩", "6666666.666667")])
_reg("电量", [("千瓦时", "1"), ("度", "1"), ("万千瓦时", "10000"), ("亿千瓦时", "100000000")])

KNOWN_UNITS = sorted(
    set(UNIT_FACTORS) | {"%", "‰", "天", "日", "小时", "分钟", "床日", "张", "名", "所", "座", "辆", "篇", "项", "家",
                         "床", "元/人", "元/次", "元/例", "万元/人", "倍", "分", "个百分点", "立方米", "公里", "千米", "米"},
    key=len, reverse=True,
)
COUNT_UNITS = {"人", "人次", "次", "件", "户", "个", "台", "台次", "例", "例次", "家", "篇", "项", "名", "所", "座",
               "辆", "张", "床"}
CURRENCY_CHARS = "¥￥$€£"
APPROX_PREFIX = ("大约", "将近", "约", "近", "超过", "逾", "不足", "不到", "接近")
APPROX_SUFFIX = ("左右", "上下", "以上", "以下", "余", "多")

FULLWIDTH = str.maketrans("０１２３４５６７８９．，％－＋（）", "0123456789.,%-+()")


# ---------------------------------------------------------------------------
# 通用
# ---------------------------------------------------------------------------

def setup_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
        except Exception:  # noqa: BLE001  控制台不支持时保持原样
            pass


class InputError(Exception):
    """输入错误，退出码 2。"""


def need_openpyxl():
    try:
        import openpyxl  # type: ignore
    except ImportError as exc:
        raise InputError("读取或写出 .xlsx 需要 openpyxl，请先运行：pip install openpyxl") from exc
    return openpyxl


def unique_path(path: Path) -> Path:
    """文件已存在时依次尝试 -v2、-v3……，不覆盖已有文件。"""
    if not path.exists():
        return path
    n = 2
    while True:
        candidate = path.with_name(f"{path.stem}-v{n}{path.suffix}")
        if not candidate.exists():
            return candidate
        n += 1


def col_letter(index: int) -> str:
    """0 起的列号转 Excel 列字母。"""
    index += 1
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def to_decimal(value: Any) -> Decimal:
    if isinstance(value, Decimal):
        return value
    if isinstance(value, float):
        return Decimal(repr(value))
    return Decimal(str(value))


def decimals_of(value: Decimal) -> int:
    """小数位数（去掉末尾的 0 之前按书写计）。"""
    exp = value.as_tuple().exponent
    return -exp if isinstance(exp, int) and exp < 0 else 0


def fmt_num(value: Decimal, places: Optional[int] = None) -> str:
    """四舍五入到指定位数；places 为 None 时按原样输出（去掉科学计数）。"""
    if places is not None:
        quant = Decimal(1).scaleb(-places)
        value = value.quantize(quant, rounding=ROUND_HALF_UP)
    text = format(value, "f")
    if places is None and "." in text:
        text = text.rstrip("0").rstrip(".")
    if text.startswith("-") and float(value) == 0:
        text = text[1:]
    return text


def fmt_raw(value: Decimal) -> str:
    """输入值原样显示：不带多余的 0，负数加括号以免与减号混淆。"""
    text = fmt_num(value)
    return f"（{text}）" if value < 0 else text


def norm_text(text: str) -> str:
    """名称归一：去空白、全角转半角、去掉末尾括注单位。"""
    text = str(text).translate(FULLWIDTH)
    text = re.sub(r"[\s　]+", "", text)
    return text


def mask_name(name: str) -> str:
    name = str(name).strip()
    if not name:
        return name
    return name[0] + "某"


def mask_digits(text: str) -> str:
    text = ID_CARD_RE.sub(lambda m: m.group(0)[:6] + "********" + m.group(0)[-4:], str(text))
    text = MOBILE_RE.sub(lambda m: m.group(0)[:3] + "****" + m.group(0)[-4:], text)
    return text


# ---------------------------------------------------------------------------
# 单元格解析
# ---------------------------------------------------------------------------

@dataclass
class Cell:
    raw: Any = None
    kind: str = "empty"  # num 数值、textnum 文本型数字、empty 空、null 占位符、error 错误值、text 文本、bool、date
    value: Optional[Decimal] = None
    flags: List[str] = field(default_factory=list)  # percent、unit、approx、currency、fullwidth、thousands、accounting、space
    unit: str = ""
    number_format: str = ""

    @property
    def is_number(self) -> bool:
        return self.kind in ("num", "textnum") and self.value is not None

    @property
    def is_empty(self) -> bool:
        return self.kind == "empty"

    @property
    def text(self) -> str:
        if self.raw is None:
            return ""
        if isinstance(self.raw, float) and self.raw.is_integer():
            return str(int(self.raw))
        return str(self.raw).strip()


NUM_BODY_RE = re.compile(r"^[+\-]?(?:\d+(?:\.\d+)?|\.\d+)$")
THOUSANDS_RE = re.compile(r"^[+\-]?\d{1,3}(?:[,，]\d{3})+(?:\.\d+)?$")


def parse_cell(raw: Any, number_format: str = "") -> Cell:
    """把单元格原值解析为 Cell。不改动原值，只记录类型和标记。"""
    cell = Cell(raw=raw, number_format=number_format or "")
    if raw is None:
        return cell
    if isinstance(raw, bool):
        cell.kind = "bool"
        return cell
    if isinstance(raw, (int, float, Decimal)):
        if isinstance(raw, float) and (math.isnan(raw) or math.isinf(raw)):
            cell.kind = "error"
            return cell
        if isinstance(raw, float) and raw.is_integer() and abs(raw) < 1e15:
            value = Decimal(int(raw))
        else:
            value = to_decimal(raw)
        if "%" in cell.number_format:
            value = (value * 100).normalize()
            cell.flags.append("percent")
        cell.kind, cell.value = "num", value
        return cell
    if hasattr(raw, "year") and hasattr(raw, "month"):
        cell.kind = "date"
        return cell

    text = str(raw)
    if not text.strip():
        return cell
    original = text
    if re.search(r"[　]|^\s|\s$", text):
        cell.flags.append("space")
    text = text.strip().strip("　")
    if text.lower() in NULL_TOKENS:
        cell.kind = "null"
        return cell
    if ERROR_RE.match(text):
        cell.kind = "error"
        return cell
    if re.search(r"[０-９．％－，]", text):
        cell.flags.append("fullwidth")
    text = text.translate(FULLWIDTH).replace(" ", "")
    for prefix in APPROX_PREFIX:
        if text.startswith(prefix):
            text = text[len(prefix):]
            cell.flags.append("approx")
            break
    for suffix in APPROX_SUFFIX:
        if text.endswith(suffix) and len(text) > len(suffix):
            head = text[: -len(suffix)]
            # "多"只在数字后面才算约数，如"20多"
            if re.search(r"\d$", head) or any(head.endswith(u) for u in KNOWN_UNITS):
                text = head
                cell.flags.append("approx")
                break
    if text and text[0] in CURRENCY_CHARS:
        text = text[1:]
        cell.flags.append("currency")
    if re.fullmatch(r"\([+\-]?[\d,.]+\)", text):
        text = "-" + text[1:-1]
        cell.flags.append("accounting")
    for unit in KNOWN_UNITS:
        if text.endswith(unit) and len(text) > len(unit) and re.search(r"\d$", text[: -len(unit)]):
            if unit == "%":
                cell.flags.append("percent")
            elif unit == "‰":
                cell.flags.append("permille")
            else:
                cell.flags.append("unit")
                cell.unit = unit
            text = text[: -len(unit)]
            break
    if THOUSANDS_RE.match(text):
        text = re.sub(r"[,，]", "", text)
        cell.flags.append("thousands")
    if NUM_BODY_RE.match(text):
        try:
            value = Decimal(text)
        except InvalidOperation:
            cell.kind = "text"
            return cell
        if "permille" in cell.flags:
            value = value / 10
            cell.flags.append("percent")
        cell.kind, cell.value = "textnum", value
        cell.raw = original
        return cell
    cell.kind = "text"
    return cell


# ---------------------------------------------------------------------------
# 读取
# ---------------------------------------------------------------------------

@dataclass
class Sheet:
    path: Path
    sheet: Optional[str]
    grid: List[List[Cell]]
    merged: List[Tuple[int, int, int, int]] = field(default_factory=list)  # 0 起：r1, c1, r2, c2
    hidden_rows: List[int] = field(default_factory=list)  # 1 起
    hidden_cols: List[str] = field(default_factory=list)
    no_cache: List[Tuple[int, int]] = field(default_factory=list)  # 公式无缓存值的单元格（0 起）
    title: str = ""  # 工作表名（单工作表时 sheet 为空，这里仍记下，用于查密级标注）
    marks: List[Tuple[str, str]] = field(default_factory=list)  # 密级和内部标注：（位置, 命中文字）

    @property
    def source(self) -> str:
        return self.path.name if not self.sheet else f"{self.path.name}［{self.sheet}］"


def scan_marking(sheet: Sheet) -> List[Tuple[str, str]]:
    """查找密级和内部标注。只返回位置和命中的标注文字，不带单元格其余内容。"""
    hits: List[Tuple[str, str]] = []

    def check(where: str, text: str) -> None:
        for regex in MARKING_RES:
            m = regex.search(text)
            if m:
                hits.append((where, re.sub(r"\s+", "", m.group(0))))
                return

    check("文件名", sheet.path.stem)
    if sheet.title or sheet.sheet:
        check("工作表名", sheet.title or sheet.sheet or "")
    for r, row in enumerate(sheet.grid):
        for c, cell in enumerate(row):
            if isinstance(cell.raw, str) and cell.raw.strip():
                check(f"{col_letter(c)}{r + 1}", cell.raw)
    return hits


def marking_refs(sheet: Sheet) -> str:
    return _limit([f"{where}“{text}”" for where, text in sheet.marks])


MARKING_STOP = ("按 guardrails 第一条停止处理：不读数、不复述、不汇总。告诉用户涉密信息不得输入联网 AI 工具；"
                "用户说明使用的是单位内网部署、允许处理此类材料的模型，或已删去相关内容，或标识系误加、材料已公开后，"
                "才可加 --marking-ok 重新运行，不得自行添加")


MARKING_OK_HELP = ("材料带密级或内部标注（DT-SEC-01）时，用户已说明可以处理后才加："
                   "内网部署模型、已删去相关内容、标识系误加或材料已公开")


class MarkingError(InputError):
    """材料带密级或内部标注，未经用户说明不处理（退出码 2）。"""


def guard_marking(sheets: Sequence[Sheet], ok: bool) -> None:
    """merge、count、calc --table、chart 读表前拦截带密级或内部标注的材料。"""
    if ok:
        return
    marked = [s for s in sheets if s.marks]
    if marked:
        refs = "；".join(f"{s.source} {marking_refs(s)}" for s in marked)
        raise MarkingError(f"DT-SEC-01 材料带有密级或内部标注（{refs}）。{MARKING_STOP}")


def read_csv(path: Path) -> List[List[Cell]]:
    data = path.read_bytes()
    text = None
    for encoding in ("utf-8-sig", "gb18030"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise InputError(f"无法识别 {path.name} 的编码，请另存为 UTF-8 编码的 CSV")
    sample = text[:4096]
    delimiter = ","
    if sample.count("\t") > sample.count(","):
        delimiter = "\t"
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    grid = []
    for row in reader:
        cells = []
        for v in row:
            cell = parse_cell(v if v != "" else None)
            # CSV 没有单元格类型，干净的数字字符串按数值处理；带单位、约数、货币符号的照样报告
            if cell.kind == "textnum" and not ({"unit", "approx", "currency"} & set(cell.flags)):
                cell.kind = "num"
            cells.append(cell)
        grid.append(cells)
    return grid


def _sniff_format(path: Path) -> str:
    head = path.read_bytes()[:8]
    if head.startswith(b"PK"):
        return "xlsx"
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        return "xls"
    return "text"


def read_xlsx(path: Path, sheet: Optional[str] = None, all_sheets: bool = False) -> List[Sheet]:
    openpyxl = need_openpyxl()
    try:
        wb_val = openpyxl.load_workbook(str(path), data_only=True)
        wb_fx = openpyxl.load_workbook(str(path), data_only=False)
    except Exception as exc:  # noqa: BLE001
        raise InputError(f"无法打开 {path.name}：{exc}") from exc
    names = wb_val.sheetnames
    if sheet:
        if sheet not in names:
            raise InputError(f"{path.name} 中没有名为“{sheet}”的工作表，现有：{'、'.join(names)}")
        names = [sheet]
    elif not all_sheets:
        names = [wb_val.active.title if wb_val.active is not None else names[0]]
    result = []
    for name in names:
        ws, wf = wb_val[name], wb_fx[name]
        if ws.max_row == 1 and ws.max_column == 1 and ws.cell(1, 1).value is None:
            continue
        grid: List[List[Cell]] = []
        no_cache: List[Tuple[int, int]] = []
        for r, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column)):
            cells = []
            for c, cell in enumerate(row):
                value = cell.value
                if value is None:
                    fx = wf.cell(row=r + 1, column=c + 1).value
                    if isinstance(fx, str) and fx.startswith("="):
                        no_cache.append((r, c))
                cells.append(parse_cell(value, getattr(cell, "number_format", "") or ""))
            grid.append(cells)
        merged = [(m.min_row - 1, m.min_col - 1, m.max_row - 1, m.max_col - 1) for m in ws.merged_cells.ranges]
        hidden_rows = [i for i, dim in ws.row_dimensions.items() if dim.hidden]
        hidden_cols = [k for k, dim in ws.column_dimensions.items() if dim.hidden]
        result.append(Sheet(path, name if (all_sheets or sheet or len(wb_val.sheetnames) > 1) else None,
                            grid, merged, sorted(hidden_rows), sorted(hidden_cols), no_cache, title=name))
    if not result:
        raise InputError(f"{path.name} 中没有可读取的数据")
    # 没有读取的工作表（如"封面""说明"）也查密级和内部标注
    others: List[Tuple[str, str]] = []
    for name in wb_val.sheetnames:
        if name in names:
            continue
        probe = Sheet(path, name, [], title=name)
        for r, row in enumerate(wb_val[name].iter_rows(values_only=True)):
            probe.grid.append([Cell(raw=v) for v in row])
        others.extend((f"工作表“{name}”" + ("" if where == "工作表名" else f" {where}"), text)
                      for where, text in scan_marking(probe) if where != "文件名")
    for s in result:
        s.marks = list(others)
    return result


def load_sheets(path_text: str, sheet: Optional[str] = None, all_sheets: bool = False) -> List[Sheet]:
    """读入表格并查密级和内部标注（结果在 Sheet.marks）。"""
    path = Path(path_text)
    if not path.is_file():
        raise InputError(f"找不到文件：{path_text}")
    suffix = path.suffix.lower()
    kind = _sniff_format(path)
    if suffix == ".xls" or kind == "xls":
        raise InputError(f"{path.name} 是旧版 .xls 格式，请在 Excel 或 WPS 中另存为 .xlsx 后再处理")
    if suffix in (".xlsx", ".xlsm") or kind == "xlsx":
        sheets = read_xlsx(path, sheet, all_sheets)
    elif suffix in (".csv", ".txt", ".tsv"):
        grid = read_csv(path)
        if not grid:
            raise InputError(f"{path.name} 是空文件")
        sheets = [Sheet(path, None, grid)]
    else:
        raise InputError(f"不支持的格式：{path.name}（只支持 .csv、.xlsx）")
    for s in sheets:
        s.marks = scan_marking(s) + s.marks
    return sheets


# ---------------------------------------------------------------------------
# 表格结构分析
# ---------------------------------------------------------------------------

@dataclass
class Column:
    index: int
    name: str
    base: str
    unit: str = ""
    unit_source: str = ""
    kind: str = "text"  # key、number、rate、avg、id、period、unitcol、text
    sensitive: bool = False

    def to_json(self) -> dict:
        return {"index": self.index, "letter": col_letter(self.index), "name": self.name, "unit": self.unit,
                "kind": self.kind}


@dataclass
class Row:
    index: int  # grid 中 0 起的行号
    kind: str  # data、total、subtotal、of_which、avg、note、blank、header_repeat
    label: str
    cells: List[Cell]

    @property
    def excel_row(self) -> int:
        return self.index + 1


@dataclass
class Issue:
    id: str
    level: str
    category: str
    location: str
    message: str
    suggestion: str = ""
    detail: Optional[dict] = None

    def to_json(self) -> dict:
        data = {"id": self.id, "level": self.level, "category": self.category, "location": self.location,
                "message": self.message, "suggestion": self.suggestion}
        if self.detail:
            data["detail"] = self.detail
        return data


@dataclass
class Table:
    sheet: Sheet
    label: str = "表A"
    header_rows: List[int] = field(default_factory=list)
    columns: List[Column] = field(default_factory=list)
    rows: List[Row] = field(default_factory=list)
    title: str = ""
    unit_note: str = ""
    periods: List[str] = field(default_factory=list)
    filler: str = ""
    source_note: str = ""
    notes: List[str] = field(default_factory=list)  # 表注（表格下方的说明行）
    pre_texts: List[str] = field(default_factory=list)  # 标题与表头之间的文字（如填报单位、口径说明）
    issues: List[Issue] = field(default_factory=list)
    checks: List[dict] = field(default_factory=list)  # 合计校验结果
    key_col: Optional[int] = None

    @property
    def source(self) -> str:
        return self.sheet.source

    def add(self, *args, **kwargs) -> None:
        self.issues.append(Issue(*args, **kwargs))

    def data_rows(self) -> List[Row]:
        return [r for r in self.rows if r.kind == "data"]

    def column(self, name: str) -> Optional[Column]:
        target = norm_text(name)
        for col in self.columns:
            if norm_text(col.name) == target or norm_text(col.base) == target:
                return col
        return None

    def cell(self, row: Row, col: Column) -> Cell:
        return row.cells[col.index] if col.index < len(row.cells) else Cell()

    def where(self, row: Row, col: Optional[Column] = None) -> str:
        text = f"{self.label} 第{row.excel_row}行"
        if row.label:
            text += f"（{self.display_label(row.label)}）"
        if col is not None:
            text += f" {col.name}"
        return text

    def display_label(self, label: str) -> str:
        """行标签脱敏。合计、小计、平均行的标签不是人名，不做姓名脱敏（不写成"合某"）。"""
        if self.key_col is not None and self.columns and self.columns[self.key_col].sensitive \
                and not is_summary_label(label):
            return mask_name(label)
        return mask_digits(label)

    def mask_free(self, text: str) -> str:
        """标题、表注等自由文字脱敏：身份证号、手机号打码；姓名列中出现的姓名改为"张某"。"""
        text = mask_digits(text)
        for col in self.columns:
            if not (col.sensitive and NAME_COL_RE.search(col.name)):
                continue
            for row in self.rows:
                value = self.cell(row, col).text
                if len(value) >= 2 and value in text:
                    text = text.replace(value, mask_name(value))
        return text


UNIT_IN_NAME_RE = re.compile(r"[（(]\s*(?:单位\s*[:：])?\s*([^（）()]{1,10}?)\s*[）)]\s*$")
UNIT_NOTE_RE = re.compile(r"(?<!填报)(?<!报送)(?<!统计)(?<!填表)(?:计量)?单位\s*[:：]\s*([^\s，,；;）)]{1,10})")
SOURCE_NOTE_RE = re.compile(r"(?:数据来源|资料来源|来源)\s*(?:为|是|[:：])\s*([^；;。]+)")
FILLER_RE = re.compile(r"(?:填报|报送|统计)单位\s*[:：]\s*([^\s，,；;]{1,30})")


def _is_yearlike(cell: Cell) -> bool:
    """表头中的年份（如 2024、2025）按文字看待。"""
    return cell.is_number and not cell.flags and cell.value == cell.value.to_integral_value() \
        and 1900 <= cell.value <= 2100  # type: ignore[union-attr, operator]


def _row_stats(cells: List[Cell], header_mode: bool = False) -> Tuple[int, int]:
    nonempty = sum(1 for c in cells if not c.is_empty)
    numeric = sum(1 for c in cells if c.is_number and not (header_mode and _is_yearlike(c)))
    return nonempty, numeric


def _is_unit_row(cells: List[Cell]) -> bool:
    texts = [c.text.strip("（）()") for c in cells if not c.is_empty]
    texts = [t for t in texts if t not in ("单位", "计量单位")]
    if len(texts) < 2:
        return False
    hits = sum(1 for t in texts if t in KNOWN_UNITS)
    return hits >= max(2, math.ceil(len(texts) * 0.6))


def _fill_merged(sheet: Sheet) -> None:
    """表头区的合并单元格向右、向下填充，便于组合多级表头；数据区不填充，只报告。"""
    for r1, c1, r2, c2 in sheet.merged:
        anchor = sheet.grid[r1][c1] if r1 < len(sheet.grid) and c1 < len(sheet.grid[r1]) else None
        if anchor is None:
            continue
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if (r, c) == (r1, c1) or r >= len(sheet.grid) or c >= len(sheet.grid[r]):
                    continue
                cell = sheet.grid[r][c]
                if cell.is_empty and anchor.kind in ("text",):
                    sheet.grid[r][c] = Cell(raw=anchor.raw, kind="text", flags=["merged"])


def analyze(sheet: Sheet, label: str = "表A", header_row: Optional[int] = None, key_name: Optional[str] = None) -> Table:
    table = Table(sheet=sheet, label=label)
    grid = sheet.grid
    width = max((len(r) for r in grid), default=0)
    for row in grid:
        while len(row) < width:
            row.append(Cell())
    if width == 0:
        raise InputError(f"{sheet.source} 中没有数据")

    # 1. 定表头
    stats = [_row_stats(r) for r in grid]
    hstats = [_row_stats(r, header_mode=True) for r in grid]
    eff_width = max((s[0] for s in stats), default=0)
    merged_header_rows = {r1 for r1, c1, r2, c2 in sheet.merged if c2 > c1}
    if header_row is not None:
        if header_row < 1 or header_row > len(grid):
            raise InputError(f"--header-row {header_row} 超出范围（共 {len(grid)} 行）")
        start = header_row - 1
    else:
        start = None
        for r in range(min(40, len(grid))):
            nonempty, numeric = hstats[r]
            if nonempty >= max(2, eff_width // 3) and numeric <= nonempty // 5:
                follow = [stats[i] for i in range(r + 1, min(r + 6, len(grid)))]
                if any(n >= 1 and ne >= 2 for ne, n in follow):
                    start = r
                    break
        if start is None:
            start = 0
            table.add("DT-STR-01", "P1", "表格结构", f"{label}", "未识别出表头，按第 1 行作表头处理",
                      "用 --header-row 指定表头所在行")
    _fill_merged(sheet)
    header_rows = [start]
    units_row: Optional[int] = None
    r = start + 1
    while r < len(grid) and len(header_rows) < 3:
        nonempty, numeric = _row_stats(grid[r], header_mode=True)
        if nonempty == 0 or numeric > 0:
            break
        if _is_unit_row(grid[r]):
            units_row = r
            r += 1
            break
        # 第二级表头：上一行有横向合并单元格，或本行填补了上一行的空位（至少两个非空单元格）
        if header_row is None and nonempty >= 2 and (r - 1 in merged_header_rows or any(
                grid[r - 1][c].is_empty and not grid[r][c].is_empty for c in range(width))):
            header_rows.append(r)
            r += 1
            continue
        break
    table.header_rows = header_rows
    body_start = r

    # 2. 表头名称（多级表头用"-"连接）
    names: List[str] = []
    for c in range(width):
        parts: List[str] = []
        for hr in header_rows:
            text = grid[hr][c].text
            if not text and hr != header_rows[-1] and not grid[header_rows[-1]][c].is_empty:
                # 上级表头为空：取左侧最近的、下面有子表头的上级表头
                for cc in range(c - 1, -1, -1):
                    left = grid[hr][cc].text
                    if left:
                        if not grid[header_rows[-1]][cc].is_empty:
                            text = left
                        break
            if text and (not parts or parts[-1] != text):
                parts.append(text)
        names.append("-".join(parts) if parts else f"第{col_letter(c)}列")
    if len(header_rows) > 1:
        table.add("DT-STR-02", "P2", "表格结构", f"{label} 第{header_rows[0] + 1}—{header_rows[-1] + 1}行",
                  "多级表头，已合并为“上级-下级”的列名", "核对列名是否正确")

    # 3. 表头以上：标题、单位、时段、填报单位
    pre_texts = []
    for rr in range(0, start):
        pre_texts.extend(c.text for c in grid[rr] if c.text)
    for text in pre_texts:
        if not table.title and len(text) >= 4 and not UNIT_NOTE_RE.search(text) and not FILLER_RE.search(text):
            table.title = text
        m = UNIT_NOTE_RE.search(text)
        if m and not table.unit_note:
            table.unit_note = m.group(1)
        m = FILLER_RE.search(text)
        if m and not table.filler:
            table.filler = m.group(1)
    table.pre_texts = [t for t in pre_texts if t != table.title]

    # 4. 列
    columns: List[Column] = []
    for c, name in enumerate(names):
        base, unit, unit_source = name, "", ""
        m = UNIT_IN_NAME_RE.search(name)
        if m and (m.group(1) in KNOWN_UNITS or len(m.group(1)) <= 4):
            unit, unit_source = m.group(1).replace("％", "%"), "表头"
            base = name[: m.start()].strip()
        elif units_row is not None and grid[units_row][c].text:
            unit, unit_source = grid[units_row][c].text.strip("（）()"), "单位行"
        columns.append(Column(index=c, name=name, base=base, unit=unit, unit_source=unit_source))
    table.columns = columns

    # 5. 行分类
    header_norm = [norm_text(n) for n in names]
    rows: List[Row] = []
    for rr in range(body_start, len(grid)):
        cells = grid[rr]
        texts = [c.text for c in cells if c.text]
        nonempty, numeric = stats[rr]
        first_text = next((c.text for c in cells if c.kind == "text" and c.text), "")
        label_text = first_text
        if nonempty == 0:
            kind = "blank"
        elif texts and any(texts[0].startswith(p) for p in NOTE_PREFIXES) and numeric <= 1:
            kind = "note"
        elif sum(1 for c, h in zip(cells, header_norm) if c.text and norm_text(c.text) == h) >= max(2, nonempty * 0.8):
            kind = "header_repeat"
        else:
            key = norm_text(first_text)
            if any(w in key for w in SUBTOTAL_WORDS):
                kind = "subtotal"
            elif len(key) <= 12 and (any(w in key for w in TOTAL_WORDS) or TOTAL_SCOPE_RE.match(key)):
                kind = "total"
            elif OF_WHICH_RE.match(key):
                kind = "of_which"
            elif AVG_ROW_RE.match(key):
                kind = "avg"
            else:
                kind = "data"
        rows.append(Row(index=rr, kind=kind, label=label_text, cells=cells))
    # 末尾的说明行、表内说明
    for row in rows:
        if row.kind == "note":
            text = " ".join(c.text for c in row.cells if c.text)
            table.notes.append(text)
            m = SOURCE_NOTE_RE.search(text)
            if m and not table.source_note:
                table.source_note = m.group(1).strip()
            m = UNIT_NOTE_RE.search(text)
            if m and not table.unit_note:
                table.unit_note = m.group(1)
    table.rows = rows

    # 6. 列类型
    data_rows = table.data_rows()
    for col in columns:
        cells = [row.cells[col.index] for row in data_rows]
        filled = [c for c in cells if not c.is_empty]
        numeric = [c for c in filled if c.is_number]
        texts = [c for c in filled if c.kind == "text"]
        if SENSITIVE_NAME_RE.search(col.name):
            col.sensitive = True
        unit_values = [c for c in filled if c.text.strip("（）()") in KNOWN_UNITS]
        numeric_like = len(numeric) + sum(1 for c in filled if c.kind in ("null", "error"))
        if UNIT_COL_RE.match(norm_text(col.base)) and filled and len(unit_values) >= len(filled) / 2:
            col.kind = "unitcol"
        elif PERIOD_NAME_RE.match(norm_text(col.base)):
            col.kind = "period"
        elif ID_NAME_RE.search(col.base):
            col.kind = "id"
        elif filled and numeric and numeric_like >= max(1, (len(filled) + 1) // 2):
            if RATE_NAME_RE.search(col.name) or col.unit in ("%", "‰") or \
                    sum(1 for c in numeric if "percent" in c.flags) > len(numeric) / 2:
                col.kind = "rate"
                if not col.unit:
                    col.unit, col.unit_source = "%", "推断"
            elif AVG_NAME_RE.search(col.base):
                col.kind = "avg"
            else:
                col.kind = "number"
            if not col.unit and table.unit_note and col.kind in ("number", "avg"):
                col.unit, col.unit_source = table.unit_note, "表注"
            if not col.unit and col.kind in ("number", "avg"):
                last = re.split(r"[-－]", col.base)[0] if "-" in col.base else col.base
                for suffix, unit in NAME_UNIT_SUFFIX:
                    if last.endswith(suffix):
                        col.unit, col.unit_source = unit, "列名推断"
                        break
            if not col.unit:
                suffixes = Counter(c.unit for c in numeric if c.unit)
                if suffixes:
                    col.unit, col.unit_source = suffixes.most_common(1)[0][0], "单元格"
        elif texts or not filled:
            col.kind = "text"
    key_candidates = [c for c in columns[:3] if c.kind == "text"]
    if key_name:
        chosen = table.column(key_name)
        if chosen is None:
            raise InputError(f"{sheet.source} 中没有名为“{key_name}”的列；现有列：{'、'.join(c.name for c in columns)}")
    else:
        def grouping(col: Column) -> bool:
            """分组列不宜作关键列：有合并单元格（如"片区"），或一半以上的值重复（如"类别"）。"""
            cells = [row.cells[col.index] for row in data_rows]
            texts = [norm_text(c.text) for c in cells if c.text]
            if any("merged" in c.flags for c in cells):
                return True
            return len(texts) >= 4 and len(set(texts)) * 2 <= len(texts)

        preferred = [c for c in key_candidates if not grouping(c)]
        chosen = (preferred or key_candidates or [None])[0]
    if chosen is not None:
        chosen.kind = "key"
        table.key_col = chosen.index
    # 行标签以关键列为准
    if table.key_col is not None:
        for row in rows:
            text = row.cells[table.key_col].text
            if text:
                row.label = text

    # 7. 时段
    sources = [table.title] + pre_texts + table.notes + [sheet.path.stem] + ([sheet.sheet] if sheet.sheet else [])
    found: List[str] = []
    for text in sources:
        for m in PERIOD_RE.finditer(text or ""):
            value = m.group(0)
            if value not in found:
                found.append(value)
    table.periods = found
    return table


# ---------------------------------------------------------------------------
# 体检
# ---------------------------------------------------------------------------

def _limit(items: List[str]) -> str:
    shown = "；".join(items[:MAX_LIST])
    return shown + (f"；等共 {len(items)} 处" if len(items) > MAX_LIST else "")


def _sum_cells(cells: List[Cell]) -> Tuple[Decimal, int, List[Cell]]:
    total = Decimal(0)
    places = 0
    bad = []
    for cell in cells:
        if cell.is_number:
            total += cell.value  # type: ignore[operator]
            places = max(places, decimals_of(cell.value))  # type: ignore[arg-type]
        elif cell.kind not in ("empty", "null"):
            bad.append(cell)
    return total, places, bad


def check_totals(table: Table) -> None:
    """把表中自带的合计、小计行当作校验和，与明细求和对账。"""
    num_cols = [c for c in table.columns if c.kind == "number"]
    skip_cols = [c for c in table.columns if c.kind in ("rate", "avg")]
    rows = table.rows
    group: List[Row] = []
    details_all: List[Row] = []
    subtotals: List[Row] = []
    pending_top: List[Row] = []  # 位于明细之前的合计行（如首行"全院"）

    def compare(target: Row, parts: List[Row], what: str, kind_label: str) -> None:
        for col in num_cols:
            stated = table.cell(target, col)
            if stated.is_empty or stated.kind == "null":
                continue
            if not stated.is_number:
                continue
            cells = [table.cell(r, col) for r in parts]
            if any("approx" in c.flags for c in cells + [stated]):
                table.checks.append({"row": target.excel_row, "column": col.name, "status": "skipped",
                                     "reason": "含约数，不能精确校验"})
                continue
            total, places, bad = _sum_cells(cells)
            empties = sum(1 for c in cells if c.is_empty or c.kind == "null")
            places = max(places, decimals_of(stated.value))  # type: ignore[arg-type]
            diff = stated.value - total  # type: ignore[operator]
            # 修约容差：末位不超过 2 个单位（P2）；超过 2 个单位、但在各项修约误差累计上限内的，报 P1 请核对
            unit = Decimal(1).scaleb(-places)
            theory = Decimal(0) if places == 0 else Decimal("0.5") * unit * (len(cells) + 1)
            tol = min(theory, 2 * unit)
            record = {"row": target.excel_row, "label": table.display_label(target.label), "column": col.name,
                      "stated": fmt_num(stated.value), "computed": fmt_num(total), "rows": len(parts),  # type: ignore[arg-type]
                      "formula": f"{what}{len(parts)}行求和＝{fmt_num(total)}，表中{kind_label}为{fmt_num(stated.value)}"}  # type: ignore[arg-type]
            if bad:
                record["status"] = "skipped"
                record["reason"] = "明细中有无法识别为数值的单元格"
            elif diff == 0:
                record["status"] = "ok"
                if empties:
                    record["note"] = f"明细中有{empties}个空值，合计不含这些行"
                    table.add("DT-SUM-05", "P2", "合计校验", table.where(target, col),
                              f"{kind_label}与明细相符，但明细中有{empties}个空值，{kind_label}不含这些行",
                              "用这一列与其他列对比（如算环比、占比）时口径不一致，先补齐或两边同口径剔除")
            elif abs(diff) <= tol:
                record["status"] = "rounding"
                table.add("DT-SUM-03", "P2", "合计校验", table.where(target, col),
                          f"{kind_label}与明细求和相差{fmt_num(diff)}，不超过末位2个单位，在修约误差范围内",
                          "确认是四舍五入造成的；表注可写“因四舍五入，分项之和与合计略有出入”",
                          {"stated": record["stated"], "computed": record["computed"]})
            elif abs(diff) <= theory:
                record["status"] = "uncertain"
                table.add("DT-SUM-06", "P1", "合计校验", table.where(target, col),
                          f"{kind_label}为{fmt_num(stated.value)}，明细{len(parts)}行求和为{fmt_num(total)}，"  # type: ignore[arg-type]
                          f"相差{fmt_num(diff)}，超过末位2个单位；可能是修约，也可能填报有误",
                          "请核对；查清前不直接引用表中合计，需要时写明细相加数并标【待核实】",
                          {"stated": record["stated"], "computed": record["computed"], "diff": fmt_num(diff)})
            else:
                record["status"] = "mismatch"
                level_id = "DT-SUM-01" if kind_label == "合计" else "DT-SUM-02"
                extra = f"，明细中有{empties}个空值未计入" if empties else ""
                table.add(level_id, "P0", "合计校验", table.where(target, col),
                          f"{kind_label}为{fmt_num(stated.value)}，明细{len(parts)}行求和为{fmt_num(total)}，"  # type: ignore[arg-type]
                          f"相差{fmt_num(diff)}{extra}",
                          "核对漏填、重复填报或手工改数；查清前不直接引用表中合计，需要时写明细相加数并标【待核实】", {
                              "stated": record["stated"], "computed": record["computed"], "diff": fmt_num(diff)})
            table.checks.append(record)
        for col in skip_cols:
            stated = table.cell(target, col)
            if stated.is_number:
                table.checks.append({"row": target.excel_row, "column": col.name, "status": "skipped",
                                     "reason": "比率或平均数不能用求和校验"})
                if not any(i.id == "DT-SUM-04" and col.name in i.location for i in table.issues):
                    table.add("DT-SUM-04", "P2", "合计校验", table.where(target, col),
                              f"“{col.name}”是比率或平均数，合计行不能用明细相加或简单平均核对",
                              "用分子、分母的合计重新计算后核对（见 calc --type ratio）")

    for row in rows:
        if row.kind == "data":
            group.append(row)
            details_all.append(row)
        elif row.kind == "subtotal":
            if group:
                compare(row, group, "上方明细", "小计")
            subtotals.append(row)
            group = []
        elif row.kind == "total":
            if not details_all:
                pending_top.append(row)
                continue
            if subtotals:
                # 合计＝各小计＋最后一个小计之后的明细
                compare(row, subtotals + group, "各小计及其余明细", "合计")
            else:
                compare(row, details_all, "全部明细", "合计")
            group = []
    for row in pending_top:
        if details_all:
            compare(row, details_all, "全部明细", "合计")


def note_hints(table: Table, rows: List[Row]) -> List[str]:
    """表注、表头前说明中提到这些行的句子（如"中医科不开展四级手术"），用来说明空值的含义。"""
    pieces: List[str] = []
    for text in table.pre_texts + table.notes:
        body = text.strip()
        if re.fullmatch(r"[（(].*[）)]", body):
            body = body[1:-1]  # 整句括起来的说明，如"（中医科不开展四级手术）"
        body = re.sub(r"^\s*(?:注|备注|说明)\s*[:：]?", "", body)
        pieces.extend(p.strip() for p in re.split(r"[；;。\n]", body) if p.strip())
    hits: List[str] = []
    for row in rows:
        key = norm_text(row.label)
        if len(key) < 2:
            continue
        for piece in pieces:
            if key in norm_text(piece) and piece not in hits:
                hits.append(piece)
    return [table.mask_free(h) for h in hits]


def _with_hints(hints: List[str], default: str) -> str:
    if not hints:
        return default
    quoted = "；".join(f"“{h}”" for h in hints[:3])
    return f"表注已说明：{quoted}，按表注处理并写进起草说明；表注没有说到的，{default}"


def inspect_table(table: Table, marking_ok: bool = False) -> Table:
    sheet = table.sheet
    label = table.label
    data_rows = table.data_rows()
    body_indexes = {r.index for r in table.rows if r.kind not in ("blank", "note")}

    # 密级和内部标注（inspect 未经确认时在 main 中直接停止，不会走到这里）
    if sheet.marks:
        if marking_ok:
            table.add("DT-SEC-01", "P2", "保密", f"{label} {marking_refs(sheet)}",
                      "材料带有密级或内部标注，用户已说明可以处理",
                      "处理结果和输出文件仍按内部材料管理：不进对外材料，不上传联网工具")
        else:
            table.add("DT-SEC-01", "P0", "保密", f"{label} {marking_refs(sheet)}",
                      f"材料带有密级或内部标注。{MARKING_STOP}", "回到开工检查第1条")

    if not data_rows:
        table.add("DT-STR-03", "P1", "表格结构", label, "没有识别出数据行", "用 --header-row 指定表头所在行")

    # 合并单元格（数据区）
    in_body = [m for m in sheet.merged if m[0] in body_indexes or m[2] in body_indexes]
    if in_body:
        refs = [f"{col_letter(c1)}{r1 + 1}:{col_letter(c2)}{r2 + 1}" for r1, c1, r2, c2 in in_body]
        table.add("DT-MRG-01", "P1", "合并单元格", f"{label} {_limit(refs)}",
                  f"数据区有 {len(refs)} 处合并单元格，合并后只有左上角单元格有值，其余按空值读取",
                  "取消合并并逐行填写，再汇总或排序", {"ranges": refs})
    if sheet.hidden_rows or sheet.hidden_cols:
        parts = []
        if sheet.hidden_rows:
            parts.append("隐藏行 " + "、".join(str(i) for i in sheet.hidden_rows[:MAX_LIST]))
        if sheet.hidden_cols:
            parts.append("隐藏列 " + "、".join(sheet.hidden_cols[:MAX_LIST]))
        table.add("DT-STR-04", "P2", "表格结构", f"{label} {'；'.join(parts)}",
                  "表中有隐藏的行或列，已一并读取", "确认隐藏内容是否应计入")
    if sheet.no_cache:
        refs = [f"{col_letter(c)}{r + 1}" for r, c in sheet.no_cache]
        table.add("DT-ERR-02", "P1", "公式", f"{label} {_limit(refs)}",
                  "公式单元格没有计算结果（文件可能由程序生成、未在 Excel 中保存过），按空值读取",
                  "在 Excel 或 WPS 中打开并保存后再处理")

    # 重复表头
    repeats = [r for r in table.rows if r.kind == "header_repeat"]
    if repeats:
        table.add("DT-STR-05", "P1", "表格结构", f"{label} 第" + "、".join(str(r.excel_row) for r in repeats) + "行",
                  "数据中夹有重复的表头行（多见于复制粘贴拼表）", "删除重复表头行")

    # 个人信息
    sens_cols = [c for c in table.columns if c.sensitive]
    if sens_cols:
        table.add("DT-PRI-01", "P0", "个人信息", f"{label} " + "、".join(c.name for c in sens_cols),
                  "表中有个人信息列，输出中已脱敏", "汇总分析只需要统计数；删去或脱敏这些列后再交给他人或上传")
    found_ids = []
    for row in table.rows:
        for c, cell in enumerate(row.cells):
            if cell.is_empty or table.columns[c].sensitive or table.columns[c].kind in ("number", "rate", "avg"):
                continue
            if ID_CARD_RE.search(cell.text) or MOBILE_RE.search(cell.text):
                found_ids.append(f"{col_letter(c)}{row.excel_row}")
    if found_ids:
        table.add("DT-PRI-02", "P0", "个人信息", f"{label} {_limit(found_ids)}",
                  "单元格中有疑似身份证号或手机号，输出中已脱敏", "删除或改为 138****0000 一类写法")

    # 逐列检查
    for col in table.columns:
        if col.kind not in ("number", "rate", "avg"):
            continue
        empties, nulls, textnums, units, approx, errors, negatives, over100, wide, words = \
            [], [], [], [], [], [], [], [], [], []
        empty_rows: List[Row] = []
        null_rows: List[Row] = []
        for row in data_rows:
            cell = table.cell(row, col)
            ref = f"{col_letter(col.index)}{row.excel_row}"
            if cell.is_empty:
                empties.append(ref)
                empty_rows.append(row)
            elif cell.kind == "null":
                nulls.append(f"{ref}“{cell.text}”")
                null_rows.append(row)
            elif cell.kind == "error":
                errors.append(f"{ref}“{cell.text}”")
            elif cell.kind in ("text", "bool", "date"):
                words.append(f"{ref}“{mask_digits(cell.text)}”")
            else:
                if cell.kind == "textnum" and not ({"approx", "unit"} & set(cell.flags)):
                    textnums.append(f"{ref}“{cell.text}”")
                if "approx" in cell.flags:
                    approx.append(f"{ref}“{cell.text}”")
                if cell.unit and col.unit and cell.unit != col.unit:
                    units.append(f"{ref}“{cell.text}”")
                elif "unit" in cell.flags or "currency" in cell.flags:
                    wide.append(f"{ref}“{cell.text}”")
                if cell.value is not None and cell.value < 0 and col.kind == "number":
                    negatives.append(f"{ref}（{fmt_num(cell.value)}）")
                if col.kind == "rate" and cell.value is not None and cell.value > 100 and \
                        not re.search(r"完成|增长|增速|增幅|同比|环比|比上年|超", col.name):
                    over100.append(f"{ref}（{fmt_num(cell.value)}%）")
        where = f"{label} {col.name}列"
        if errors:
            table.add("DT-ERR-01", "P0", "公式", f"{where}：{_limit(errors)}", "单元格是公式错误值",
                      "在原表中修正公式（常见原因是除数为 0 或引用失效）")
        if empties:
            hints = note_hints(table, empty_rows)
            table.add("DT-NUL-01", "P1", "空值", f"{where}：{_limit(empties)}",
                      f"{len(empties)} 个空白单元格。空白不等于 0，本工具不按 0 计算",
                      _with_hints(hints, "向填报单位确认是未填报还是 0；正文涉及的写【待补：说明】"),
                      {"notes": hints} if hints else None)
        if nulls:
            hints = note_hints(table, null_rows)
            table.add("DT-NUL-02", "P1", "空值", f"{where}：{_limit(nulls)}",
                      f"{len(nulls)} 个单元格用符号或文字代替数值（如“—”“/”“无”），按空值处理",
                      _with_hints(hints, "确认含义：是 0、不适用还是未填报；正文涉及的写【待补：说明】"),
                      {"notes": hints} if hints else None)
        if textnums:
            table.add("DT-TXT-01", "P1", "文本型数字", f"{where}：{_limit(textnums)}",
                      f"{len(textnums)} 个数值以文本形式保存，Excel 求和时会被跳过，合计偏小",
                      "在 Excel 中把该列转换为数字（数据→分列，或选中后“转换为数字”）")
        if words:
            table.add("DT-TXT-05", "P1", "文本型数字", f"{where}：{_limit(words)}",
                      f"数值列中有 {len(words)} 个单元格是文字，不能参与计算", "改填数值，或在备注栏说明")
        if wide:
            table.add("DT-TXT-02", "P1", "文本型数字", f"{where}：{_limit(wide)}",
                      "数值后带单位或货币符号，Excel 不能直接计算", "单位写进表头，单元格只填数字")
        if approx:
            table.add("DT-TXT-03", "P1", "约数", f"{where}：{_limit(approx)}",
                      "约数（如“约20”“30多”）不能参与精确计算，合计和增速都不可靠",
                      "请填报单位给出准确数；正文引用时写【待核实：准确数】")
        if units:
            table.add("DT-UNIT-01", "P0", "单位", f"{where}：{_limit(units)}",
                      f"单元格单位与本列单位“{col.unit}”不一致", "统一换算后再计算，换算过程写进计算底稿")
        if negatives:
            table.add("DT-VAL-01", "P2", "数值", f"{where}：{_limit(negatives)}", "出现负数",
                      "确认是否为减少量、亏损或录入错误")
        if over100:
            table.add("DT-VAL-02", "P2", "数值", f"{where}：{_limit(over100)}", "比率超过 100%",
                      "核对分子分母是否取反、口径是否一致")
        if not col.unit and col.kind in ("number", "avg") and not any(c.kind == "unitcol" for c in table.columns):
            table.add("DT-UNIT-02", "P2", "单位", where, "未标注计量单位",
                      "在表头或表注写明单位；正文引用前向填报单位确认")
        # 全角数字、首尾空格
        spaces = [f"{col_letter(col.index)}{r.excel_row}" for r in data_rows
                  if {"fullwidth", "space"} & set(table.cell(r, col).flags)]
        if spaces:
            table.add("DT-TXT-04", "P2", "文本型数字", f"{where}：{_limit(spaces)}", "数值含全角字符或首尾空格",
                      "改为半角数字，删除空格")

    # 关键列：重复、写法不一致
    if table.key_col is not None:
        key_col = table.columns[table.key_col]
        seen: Dict[str, Row] = {}
        variants: Dict[str, set] = {}
        dups = []
        for row in data_rows:
            text = row.cells[key_col.index].text
            if not text or "merged" in row.cells[key_col.index].flags:
                continue  # 合并单元格向下填充出来的值不算重复
            norm = re.sub(r"[（(].*?[）)]", "", norm_text(text))
            raw = row.cells[key_col.index].raw
            variants.setdefault(norm, set()).add(str(raw) if raw is not None else text)
            if norm in seen:
                dups.append(f"第{seen[norm].excel_row}行与第{row.excel_row}行（{table.display_label(text)}）")
            else:
                seen[norm] = row
        if dups:
            table.add("DT-DUP-01", "P1", "重复", f"{label} {key_col.name}列：{_limit(dups)}",
                      "同一名称出现多次，汇总时会重复计算", "确认是否重复报送，删除重复行或合并为一行")
        odd = ["、".join(f"“{table.display_label(v)}”" for v in sorted(vs)) for vs in variants.values() if len(vs) > 1]
        if odd:
            table.add("DT-DUP-02", "P1", "重复", f"{label} {key_col.name}列：{_limit(odd)}",
                      "同一名称写法不一致（空格、全角半角、括注不同），分组汇总时会被当成不同单位",
                      "统一写法")
    # 整行重复
    sigs: Dict[Tuple, int] = {}
    whole = []
    for row in data_rows:
        sig = tuple(c.text for c in row.cells)
        if sig in sigs and any(c.is_number for c in row.cells):
            whole.append(f"第{sigs[sig]}行与第{row.excel_row}行")
        else:
            sigs[sig] = row.excel_row
    if whole:
        table.add("DT-DUP-03", "P1", "重复", f"{label} {_limit(whole)}", "整行完全相同", "确认是否重复粘贴")

    # 合计校验
    check_totals(table)
    has_total = any(r.kind in ("total", "subtotal") for r in table.rows)
    if not has_total and data_rows:
        table.checks.append({"status": "none", "reason": "表中没有合计行，无法对账"})
    of_which = [r for r in table.rows if r.kind == "of_which"]
    if of_which:
        table.add("DT-STR-06", "P2", "表格结构",
                  f"{label} 第" + "、".join(str(r.excel_row) for r in of_which[:MAX_LIST]) + "行",
                  "“其中”行是上一行的组成部分，已排除在求和之外", "确认这些行没有被重复加总")

    # 时段
    if not table.periods:
        table.add("DT-PER-01", "P2", "时段", label, "表中未写明统计时段", "在标题或表注写明，如“2026年1—9月”")
    elif len(table.periods) > 1:
        table.add("DT-PER-03", "P2", "时段", label, "表中出现多个时段：" + "、".join(table.periods),
                  "确认各列分别对应哪个时段")
    return table


def table_summary(table: Table) -> dict:
    counts = Counter(i.level for i in table.issues)
    numeric_stats = {}
    data_rows = table.data_rows()
    for col in table.columns:
        if col.kind in ("number", "rate", "avg"):
            values = [table.cell(r, col).value for r in data_rows if table.cell(r, col).is_number]
            if values:
                numeric_stats[col.name] = {
                    "count": len(values), "missing": len(data_rows) - len(values),
                    "sum": fmt_num(sum(values)) if col.kind == "number" else None,  # type: ignore[arg-type]
                    "min": fmt_num(min(values)), "max": fmt_num(max(values)),  # type: ignore[type-var]
                }
    return {
        "label": table.label,
        "source": table.source,
        "stopped": False,
        "title": table.mask_free(table.title),
        "pre_texts": [table.mask_free(t) for t in table.pre_texts],
        "filler": table.filler,
        "unit_note": table.unit_note,
        "periods": table.periods,
        "source_note": table.source_note,
        "notes": [table.mask_free(t) for t in table.notes],
        "header_rows": [r + 1 for r in table.header_rows],
        "columns": [c.to_json() for c in table.columns],
        "data_rows": len(data_rows),
        "total_rows": [{"row": r.excel_row, "label": table.display_label(r.label), "kind": r.kind}
                       for r in table.rows if r.kind in ("total", "subtotal")],
        "checks": table.checks,
        "stats": numeric_stats,
        "summary": {lv: counts.get(lv, 0) for lv in LEVELS},
        "issues": [i.to_json() for i in sorted(table.issues, key=issue_order)],
    }


def issue_order(issue: Issue) -> Tuple[int, int]:
    """级别高的在前；同级中保密问题（DT-SEC）排第一。"""
    return LEVELS.index(issue.level), 0 if issue.id.startswith("DT-SEC") else 1


def stopped_summary(sheet: Sheet, label: str) -> dict:
    """带密级或内部标注、用户尚未说明时的体检结果：只报标注的位置，不输出标题、表注、统计和其他问题。"""
    issue = Issue("DT-SEC-01", "P0", "保密", f"{label} {marking_refs(sheet)}",
                  f"材料带有密级或内部标注。{MARKING_STOP}", "回到开工检查第1条")
    return {"label": label, "source": sheet.source, "stopped": True,
            "marks": [{"location": where, "text": text} for where, text in sheet.marks[:MAX_LIST]],
            "summary": {"P0": 1, "P1": 0, "P2": 0}, "issues": [issue.to_json()]}


# ---------------------------------------------------------------------------
# 合并
# ---------------------------------------------------------------------------

def table_letter(i: int) -> str:
    return "表" + (chr(65 + i) if i < 26 else f"{chr(65 + i // 26 - 1)}{chr(65 + i % 26)}")


def unit_dim(unit: str) -> Optional[Tuple[str, Decimal]]:
    return UNIT_FACTORS.get(unit)


def convert(value: Decimal, src: str, dst: str) -> Optional[Decimal]:
    a, b = unit_dim(src), unit_dim(dst)
    if not a or not b or a[0] != b[0]:
        return None
    return value * a[1] / b[1]


def _factor_text(src: str, dst: str) -> str:
    a, b = unit_dim(src), unit_dim(dst)
    if not a or not b:
        return ""
    ratio = a[1] / b[1]
    if ratio >= 1:
        return f"×{fmt_num(ratio.normalize())}"
    return f"÷{fmt_num((1 / ratio).normalize())}"


def read_expect(text: Optional[str]) -> List[str]:
    if not text:
        return []
    path = Path(text)
    if path.is_file():
        raw = path.read_text(encoding="utf-8-sig")
        items = re.split(r"[\r\n,，、；;]+", raw)
    else:
        items = re.split(r"[,，、；;]+", text)
    return [i.strip() for i in items if i.strip()]


def norm_period(text: str) -> str:
    return re.sub(r"[—\-－~～至到]", "-", norm_text(text)).replace("份", "")


def strip_paren(text: str) -> str:
    return re.sub(r"[（(].*?[）)]", "", norm_text(text))


def merge_tables(tables: List[Table], key: Optional[str] = None, layout: str = "auto",
                 expect: Optional[List[str]] = None) -> dict:
    """合并多张下级报表。口径冲突不合并：时段不同的表不计入汇总，单位不能换算的列不汇总，比率和平均数不相加。"""
    issues: List[Issue] = []

    def add(*args, **kwargs):
        issues.append(Issue(*args, **kwargs))

    # 1. 关键列
    for t in tables:
        if key:
            col = t.column(key)
            if col is None:
                raise InputError(f"{t.label}（{t.source}）中没有名为“{key}”的列")
            if t.key_col is not None and t.key_col != col.index:
                t.columns[t.key_col].kind = "text"
            col.kind = "key"
            t.key_col = col.index
            for row in t.rows:
                if row.cells[col.index].text:
                    row.label = row.cells[col.index].text
    key_names = sorted({norm_text(t.columns[t.key_col].base) for t in tables if t.key_col is not None})
    if len(key_names) > 1:
        add("DT-MG-10", "P2", "表头", "各表关键列", "各表的关键列名称不同：" + "、".join(key_names),
            "已按各表的第一列文字列对应，请确认含义相同")

    # 2. 布局
    if layout == "auto":
        layout = "record"
        counter = Counter()
        for t in tables:
            if t.key_col is not None:
                counter.update({strip_paren(r.cells[t.key_col].text) for r in t.data_rows() if r.cells[t.key_col].text})
        if len(tables) >= 2 and counter:
            shared = sum(1 for n in counter.values() if n >= 2)
            if shared / len(counter) >= 0.5:
                layout = "indicator"

    # 3. 时段：以多数表的时段为基准，不一致的表不计入汇总
    period_of = {t.label: (t.periods[0] if t.periods else "") for t in tables}
    known = Counter(norm_period(p) for p in period_of.values() if p)
    baseline_norm = known.most_common(1)[0][0] if known else ""
    baseline = next((p for p in period_of.values() if p and norm_period(p) == baseline_norm), "")
    excluded: List[str] = []
    for t in tables:
        p = period_of[t.label]
        if not p:
            if baseline:
                add("DT-MG-04", "P1", "时段", f"{t.label}（{t.source}）",
                    f"未写明统计时段，按与多数表相同（{baseline}）处理", "向填报单位确认")
        elif norm_period(p) != baseline_norm:
            excluded.append(t.label)
            add("DT-MG-03", "P0", "时段", f"{t.label}（{t.source}）",
                f"统计时段为“{p}”，与多数表的“{baseline}”不一致，未计入汇总",
                "请填报单位按统一时段重报，或单独列示并在正文注明")

    # 4. 数值列按归一名称对齐
    col_order: List[str] = []
    col_info: "OrderedDict[str, Dict[str, Column]]" = OrderedDict()
    for t in tables:
        for col in t.columns:
            if col.kind not in ("number", "rate", "avg"):
                continue
            nb = norm_text(col.base)
            if nb not in col_info:
                col_info[nb] = {}
                col_order.append(nb)
            col_info[nb][t.label] = col
    for nb in col_order:
        missing = [t.label for t in tables if t.label not in col_info[nb]]
        if missing:
            name = next(iter(col_info[nb].values())).base
            add("DT-MG-01", "P1", "表头", f"“{name}”列", f"{'、'.join(missing)} 没有这一列，该列汇总不完整",
                "确认是否漏填或列名写法不同")

    # 5. 收集数值
    entries: List[dict] = []
    merged_rows: List[dict] = []
    for t in tables:
        unit_col = next((c for c in t.columns if c.kind == "unitcol"), None)
        for row in t.data_rows():
            key_text = row.cells[t.key_col].text if t.key_col is not None else ""
            out = {"来源": t.label, "单位名称": t.filler or t.sheet.sheet or t.sheet.path.stem,
                   "行": row.excel_row, "关键列": t.display_label(key_text), "数值": OrderedDict()}
            row_unit = row.cells[unit_col.index].text if unit_col is not None else ""
            for nb in col_order:
                col = col_info[nb].get(t.label)
                if col is None:
                    continue
                cell = t.cell(row, col)
                ok = cell.is_number and "approx" not in cell.flags
                kind = col.kind
                if layout == "indicator" and (RATE_NAME_RE.search(key_text) or AVG_NAME_RE.search(key_text)):
                    kind = "rate"
                entries.append({
                    "table": t.label, "row": row.excel_row, "key": key_text,
                    "group": strip_paren(key_text) if layout == "indicator" else "合计",
                    "nb": nb, "name": col.base, "kind": kind,
                    "value": cell.value if ok else None,
                    "unit": row_unit or cell.unit or col.unit, "excluded": t.label in excluded, "out": out,
                })
            merged_rows.append(out)

    # 6. 定目标单位：同组同列取最常见的单位；不同量纲不能换算的标为冲突
    targets: Dict[Tuple[str, str], str] = {}
    blocked: Dict[Tuple[str, str], str] = {}
    by_slot: "OrderedDict[Tuple[str, str], List[dict]]" = OrderedDict()
    for e in entries:
        by_slot.setdefault((e["group"], e["nb"]), []).append(e)
    for slot, items in by_slot.items():
        units = Counter(e["unit"] for e in items if e["unit"])
        if not units:
            targets[slot] = ""
            continue
        target = units.most_common(1)[0][0]
        targets[slot] = target
        bad = sorted({e["unit"] for e in items if e["unit"] and e["unit"] != target
                      and convert(Decimal(1), e["unit"], target) is None})
        if bad:
            where = sorted({e["table"] for e in items if e["unit"] in bad})
            blocked[slot] = f"{'、'.join(where)}的单位“{'、'.join(bad)}”与“{target}”不能换算"
            name = items[0]["name"] if layout == "record" else f"{items[0]['key']} {items[0]['name']}"
            add("DT-MG-05", "P0", "单位", f"{'、'.join(where)} “{name}”",
                f"单位“{'、'.join(bad)}”与其他表的“{target}”不是同一计量（如人与人次、元与件），不能相加",
                "请填报单位按统一口径重报；该项未汇总")

    # 7. 换算并写回明细
    conversions: List[dict] = []
    for e in entries:
        slot = (e["group"], e["nb"])
        value, unit, target = e["value"], e["unit"], targets.get(slot, "")
        if value is not None and unit and target and unit != target and slot not in blocked:
            new = convert(value, unit, target)
            if new is not None:
                new = new.normalize()
                conversions.append({"table": e["table"], "row": e["row"], "column": e["name"], "from": unit,
                                    "to": target,
                                    "formula": f"{fmt_num(value)}{unit}{_factor_text(unit, target)}＝{fmt_num(new)}{target}"})
                value = new
        e["value"] = value
        e["out"]["数值"][e["name"]] = None if value is None else fmt_num(value)
        if layout == "indicator":
            e["out"]["计量单位"] = target or unit

    # 8. 重复报送：同一名称出现在两张以上计入汇总的表中，先不计入合计，另列含重复的合计，确认后再加
    dup_tables: "OrderedDict[str, List[str]]" = OrderedDict()
    if layout == "record":
        for e in entries:
            k = strip_paren(e["key"])
            if k and not e["excluded"]:
                dup_tables.setdefault(k, [])
                if e["table"] not in dup_tables[k]:
                    dup_tables[k].append(e["table"])
    dup_keys = {k for k, ts in dup_tables.items() if len(ts) >= 2}

    def is_dup(e: dict) -> bool:
        return bool(dup_keys) and strip_paren(e["key"]) in dup_keys

    sens_key = any(t.key_col is not None and t.columns[t.key_col].sensitive for t in tables)

    def mask_dup(k: str) -> str:
        return mask_name(k) if sens_key else mask_digits(k)

    # 9. 汇总
    aggregate: "OrderedDict[str, dict]" = OrderedDict()
    for slot, items in by_slot.items():
        group, nb = slot
        first = items[0]
        bucket = aggregate.setdefault(group, {"key": first["key"] if layout == "indicator" else "合计",
                                              "values": OrderedDict()})
        used = [e for e in items if not e["excluded"]]
        entry: Dict[str, Any] = {"value": None, "unit": targets.get(slot, ""), "formula": "", "complete": True,
                                 "note": ""}
        kinds = {e["kind"] for e in items}
        if kinds & {"rate", "avg"}:
            entry["complete"] = False
            entry["note"] = "比率或平均数不能相加，用分子、分母的合计重算"
        elif slot in blocked:
            entry["complete"] = False
            entry["note"] = "单位冲突，未汇总：" + blocked[slot]
        else:
            parts = [(e["table"], e["value"]) for e in used if e["value"] is not None and not is_dup(e)]
            dups = [(e["table"], e["value"]) for e in used if e["value"] is not None and is_dup(e)]
            missing = [f"{e['table']}第{e['row']}行" for e in used if e["value"] is None]
            total = sum((v for _, v in parts), Decimal(0))
            entry["value"] = fmt_num(total)
            if len(parts) <= 6:
                entry["formula"] = "＋".join(f"{lb} {fmt_raw(v)}" for lb, v in parts) + f"＝{fmt_num(total)}"
            else:
                entry["formula"] = f"{len(parts)}个数值求和＝{fmt_num(total)}"
            notes = []
            absent = [t.label for t in tables if t.label not in excluded and t.label not in col_info[nb]]
            if missing:
                notes.append(f"{len(missing)}处空值或约数未计入：" + _limit(missing))
            if absent and layout == "record":
                notes.append(f"{'、'.join(absent)}没有这一列")
            if notes:
                add("DT-MG-07", "P1", "汇总", f"{bucket['key']} {first['name']}", "汇总不完整：" + "；".join(notes),
                    "补齐后重新汇总；正文引用这个合计时标【待核实：说明】")
            if dups:
                with_dups = total + sum((v for _, v in dups), Decimal(0))
                entry["value_with_dups"] = fmt_num(with_dups)
                note = f"重复报送待确认的{len(dups)}个数值未计入，含重复的合计为{fmt_num(with_dups)}"
                # 只计一次：同一单位在各表中报的数相同时才能确定；数不同的不替用户挑
                groups: "OrderedDict[str, List[Tuple[str, Decimal]]]" = OrderedDict()
                for e in used:
                    if e["value"] is not None and is_dup(e):
                        groups.setdefault(strip_paren(e["key"]), []).append((e["table"], e["value"]))
                differ = [f"{mask_dup(k)}（" + "、".join(f"{lb} {fmt_num(v)}" for lb, v in vs) + "）"
                          for k, vs in groups.items() if len({v for _, v in vs}) > 1]
                if differ:
                    note += f"；重复单位在各表中的数不同（{_limit(differ)}），不能自动只计一次，请向填报单位确认"
                else:
                    dedup = total + sum((vs[0][1] for vs in groups.values()), Decimal(0))
                    entry["value_dedup"] = fmt_num(dedup)
                    note += f"；重复单位只计一次的合计为{fmt_num(dedup)}，确认是重复报送的采用这个数"
                notes.append(note)
            if notes:
                entry["complete"] = False
                entry["note"] = "；".join(notes)
        bucket["values"][first["name"]] = entry

    rate_cols = sorted({(e["key"] if layout == "indicator" else e["name"]) for e in entries
                        if e["kind"] in ("rate", "avg")})
    if rate_cols:
        add("DT-MG-08", "P2", "汇总", "、".join(rate_cols[:MAX_LIST]), "比率、平均数未汇总",
            "全院（全系统）的比率用分子合计除以分母合计计算，不用各单位比率相加或简单平均")
    if conversions:
        pairs = sorted({f"{c['table']} {c['column']}：{c['from']}→{c['to']}" for c in conversions})
        add("DT-MG-06", "P1", "单位", _limit(pairs), f"已把 {len(conversions)} 个数值换算为统一单位",
            "核对换算是否正确，换算过程写进计算底稿")

    # 10. 重复报送、应报未报
    keys_present = {strip_paren(o["关键列"]) for o in merged_rows if o["关键列"]}
    if layout == "record":
        if dup_keys:
            dups = [f"{mask_dup(k)}（{'、'.join(ts)}）" for k, ts in dup_tables.items() if k in dup_keys]
            add("DT-MG-09", "P0", "重复", _limit(dups),
                "同一名称在多张表中出现，可能重复报送；这些行暂未计入合计。另列两个数：含重复的合计（value_with_dups），"
                "以及各表报数相同时重复单位只计一次的合计（value_dedup）",
                "向填报单位确认是重复报送还是同名的不同单位：确认是重复报送的，采用只计一次的合计；是不同单位的，"
                "采用含重复的合计，或用 --key 指定能区分的列重新汇总")
        present = keys_present | {norm_text(o["单位名称"]) for o in merged_rows}
    else:
        present = {norm_text(t.filler or t.sheet.sheet or t.sheet.path.stem) for t in tables}
    if expect:
        def has(name: str) -> bool:
            n = strip_paren(name)
            return any(n == p or (len(n) >= 2 and n in p) for p in present)

        missing_units = [e for e in expect if not has(e)]
        if missing_units:
            add("DT-MG-11", "P1", "应报未报", _limit(missing_units), f"名单中有 {len(missing_units)} 个单位没有数据",
                "催报；正文写全口径合计时注明“不含……”，或标【待补：说明】")
        if layout == "record":
            expected = {strip_paren(e) for e in expect}
            extra = sorted(k for k in keys_present if k not in expected)
            if extra:
                add("DT-MG-12", "P2", "应报未报", _limit(extra), "有名单之外的单位", "确认是否应计入")

    for t in tables:
        issues.extend(t.issues)
    counts = Counter(i.level for i in issues)
    columns_out = []
    for nb in col_order:
        cols = list(col_info[nb].values())
        units = Counter(c.unit for c in cols if c.unit)
        columns_out.append({"name": cols[0].base, "unit": units.most_common(1)[0][0] if units else "",
                            "kind": "rate" if any(c.kind == "rate" for c in cols) else
                            ("avg" if any(c.kind == "avg" for c in cols) else "number")})
    return {
        "layout": layout,
        "tables": [{"label": t.label, "source": t.source, "filler": t.filler, "period": period_of[t.label],
                    "data_rows": len(t.data_rows()), "excluded": t.label in excluded,
                    "summary": {lv: Counter(i.level for i in t.issues).get(lv, 0) for lv in LEVELS}} for t in tables],
        "baseline_period": baseline,
        "excluded_tables": excluded,
        "columns": columns_out,
        "conversions": conversions,
        "aggregate": list(aggregate.values()),
        "rows": merged_rows,
        "summary": {lv: counts.get(lv, 0) for lv in LEVELS},
        "issues": [i.to_json() for i in sorted(issues, key=issue_order)],
    }


def write_merge_output(result: dict, out_path: Path) -> Path:
    out_path = unique_path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    value_names: List[str] = []
    for row in result["rows"]:
        for name in row["数值"]:
            if name not in value_names:
                value_names.append(name)
    indicator = result.get("layout") == "indicator"
    units = {c["name"]: c["unit"] for c in result.get("columns", [])}
    shown = value_names if indicator else [f"{n}（{units[n]}）" if units.get(n) else n for n in value_names]
    header = ["来源", "单位名称", "原表行号", "名称"] + (["计量单位"] if indicator else []) + shown
    detail = [[r["来源"], r["单位名称"], r["行"], r["关键列"]] + ([r.get("计量单位", "")] if indicator else [])
              + [r["数值"].get(n) for n in value_names] for r in result["rows"]]
    agg_rows = []
    for item in result["aggregate"]:
        unit = next((v.get("unit") for v in item["values"].values() if v.get("unit")), "")
        # 汇总行旁写明不完整、含重复待确认等情况，不让一个"干净"的合计数单独出现
        notes = "；".join(f"{n}：{v['note']}" for n, v in item["values"].items() if v.get("note") and v.get("value"))
        agg_rows.append([item["key"]] + ([unit] if indicator else [])
                        + [item["values"].get(n, {}).get("value") or item["values"].get(n, {}).get("note", "")
                           for n in value_names] + [notes])
    if out_path.suffix.lower() == ".csv":
        with out_path.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.writer(fh)
            writer.writerow(header + ["说明"])
            writer.writerows(detail)
            for row in agg_rows:
                writer.writerow(["汇总", "", ""] + row)
        return out_path
    openpyxl = need_openpyxl()
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # type: ignore

    wb = openpyxl.Workbook()
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def write_sheet(ws, title: str, head: List[str], body: List[list]) -> None:
        ws.append([title])
        ws.cell(1, 1).font = Font(name="宋体", bold=True, size=12)
        ws.append(head)
        for c in range(1, len(head) + 1):
            cell = ws.cell(2, c)
            cell.font = Font(name="宋体", bold=True)
            cell.fill = PatternFill("solid", fgColor="D9D9D9")
        for row in body:
            ws.append([_excel_value(v) for v in row])
        for r in range(2, ws.max_row + 1):
            for c in range(1, len(head) + 1):
                cell = ws.cell(r, c)
                cell.border = border
                if r > 2:
                    cell.font = Font(name="宋体")
                cell.alignment = Alignment(wrap_text=True, vertical="center")
        for c in range(1, len(head) + 1):
            width = max(len(str(ws.cell(r, c).value or "")) for r in range(2, ws.max_row + 1))
            ws.column_dimensions[col_letter(c - 1)].width = min(40, max(8, width * 2 + 2))
        ws.freeze_panes = "A3"

    ws = wb.active
    ws.title = "汇总"
    write_sheet(ws, "汇总结果（比率、平均数和有冲突的项未汇总，见说明）",
                ["名称"] + (["计量单位"] if indicator else []) + shown + ["说明"], agg_rows)
    write_sheet(wb.create_sheet("明细"), "合并明细（已排除合计、小计、说明行）", header, detail)
    issues = result["issues"]
    write_sheet(wb.create_sheet("问题清单"), "体检和合并问题清单", ["序号", "级别", "类别", "位置", "问题", "处理建议"],
                [[i + 1, it["level"], it["category"], it["location"], it["message"], it["suggestion"]]
                 for i, it in enumerate(issues)])
    if result["conversions"]:
        write_sheet(wb.create_sheet("单位换算"), "单位换算记录", ["表", "原表行号", "列", "换算"],
                    [[c["table"], c["row"], c["column"], c["formula"]] for c in result["conversions"]])
    wb.save(str(out_path))
    return out_path


def _excel_value(value: Any) -> Any:
    if isinstance(value, str) and re.fullmatch(r"-?\d+(?:\.\d+)?", value):
        number = Decimal(value)
        return int(number) if number == number.to_integral_value() and "." not in value else float(number)
    return value


# ---------------------------------------------------------------------------
# 指标计算
# ---------------------------------------------------------------------------

BASIS_LABELS = {
    "同比": "上年同期", "环比": "上期", "比上年": "上年", "比上年同期": "上年同期", "比年初": "年初",
    "比上月": "上月", "比上季度": "上季度", "比上学年": "上学年",
}
RATE_VERB_SPEED = re.compile(r"增速|增长率|增幅|增长速度")


class CalcError(InputError):
    pass


def parse_input_number(text: Any, what: str) -> Decimal:
    if isinstance(text, Decimal):
        return text
    cell = parse_cell(str(text).strip())
    if not cell.is_number:
        raise CalcError(f"{what}“{text}”不是数值")
    if "approx" in cell.flags:
        raise CalcError(f"{what}“{text}”是约数，不能用于精确计算；请提供准确数")
    return cell.value  # type: ignore[return-value]


def parse_values(text: Optional[str], what: str) -> List[Decimal]:
    if not text:
        raise CalcError(f"缺少 {what}")
    items = [s for s in re.split(r"[,，、;；\s]+", str(text)) if s]
    return [parse_input_number(s, what) for s in items]


@dataclass
class CalcResult:
    type: str
    name: str = ""
    value: Optional[Decimal] = None
    places: int = 1
    unit: str = ""
    formula: str = ""
    inputs: Dict[str, str] = field(default_factory=dict)
    extra: Dict[str, Any] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)
    basis: str = ""

    @property
    def display(self) -> str:
        if self.value is None:
            return ""
        return f"{fmt_num(self.value, self.places)}{self.unit}"

    def to_json(self) -> dict:
        data = {"type": self.type, "name": self.name, "basis": self.basis, "inputs": self.inputs,
                "value": None if self.value is None else float(self.value),
                "rounded": None if self.value is None else fmt_num(self.value, self.places),
                "unit": self.unit, "display": self.display, "formula": self.formula, "warnings": self.warnings}
        for k, v in self.extra.items():
            data[k] = v
        return data


def _approx_tail(exact: Decimal, places: int, show: int, unit: str = "") -> str:
    """算式结尾：能整除时写"＝7.0%"，修约时写"＝6.986%≈7.0%"。"""
    rounded = fmt_num(exact, places)
    if exact == Decimal(rounded):
        return f"＝{rounded}{unit}"
    shown = fmt_num(exact, show)
    if Decimal(shown) == Decimal(rounded):
        return f"≈{rounded}{unit}"
    return f"＝{shown}{unit}≈{rounded}{unit}"


def input_places(*values: Decimal) -> int:
    return max((decimals_of(v) for v in values), default=0)


def is_count_unit(unit: str) -> bool:
    return unit in COUNT_UNITS


def calc(kind: str, a: argparse.Namespace, wording: bool = False) -> CalcResult:
    """计算并补充提示。wording 为 True（phrase 生成句子）时，另提示句子里用了哪些默认叫法。"""
    res = _calc(kind, a)
    if kind in ("growth", "change", "pp", "multiple") and not (a.basis or a.base_label):
        res.warnings.insert(0, "未指明对比基期（同比、环比、比上年等），句子中暂写“比基期”“基期”，"
                               "请用 --basis 或 --base-label 写明实际基期")
    if not wording:
        return res
    if kind == "completion" and not a.base_label:
        res.warnings.append("未用 --base-label 写明目标名称，句子中暂写“年度目标”；季度目标、专项目标请改写")
    if kind in ("share", "pull", "contrib"):
        if not (a.part_name or a.name):
            res.warnings.append("未写部分的名称（--part-name 或 --name），句子中暂写“该指标”，请改为实际名称")
        if not a.total_name:
            res.warnings.append("未写总量的名称（--total-name），句子中暂写“总量”，请改为实际名称（如“全县产值”）")
    elif not a.name:
        res.warnings.append("未写指标名称（--name），句子中暂写“该指标”，请改为实际名称")
    return res


def _calc(kind: str, a: argparse.Namespace) -> CalcResult:
    places = a.decimals
    name = a.name or ""
    unit = a.unit or ""
    res = CalcResult(type=kind, name=name)
    pct_places = 1 if places is None else places

    def need(attr: str, what: str) -> Decimal:
        value = getattr(a, attr, None)
        if value is None or value == "":
            raise CalcError(f"--type {kind} 需要 {what}（--{attr.replace('_', '-')}）")
        return parse_input_number(value, what)

    if kind in ("growth", "change"):
        cur, base = need("current", "本期数"), need("base", "基期数")
        diff = cur - base
        res.inputs = {"current": fmt_num(cur), "base": fmt_num(base)}
        res.basis = a.basis or ""
        dplaces = input_places(cur, base) if places is None else places
        res.extra["change"] = {"value": float(diff), "display": f"{fmt_num(diff, dplaces)}{unit}",
                               "formula": f"{fmt_raw(cur)}－{fmt_raw(base)}＝{fmt_num(diff, dplaces)}{unit}"}
        if kind == "change":
            res.value, res.places, res.unit = diff, dplaces, unit
            res.formula = res.extra["change"]["formula"]
            return res
        res.unit, res.places = "%", pct_places
        if base == 0:
            res.warnings.append("基期为0，增速没有意义，只写增减量（如“上年同期为0，本期为……”）")
            res.formula = f"基期为0，不计算增速；增减量 {res.extra['change']['formula']}"
            return res
        if base < 0 or cur < 0:
            res.warnings.append("基期或本期为负数（如亏损），增速百分比没有通常含义；用“扭亏为盈”“减亏”“增亏”和增减额表述")
            res.extra["sign_case"] = _sign_case(cur, base)
            res.formula = f"不计算增速；增减额 {res.extra['change']['formula']}"
            return res
        exact = diff / base * 100
        res.value = exact
        res.formula = f"（{fmt_raw(cur)}－{fmt_raw(base)}）÷{fmt_raw(base)}×100%" + _approx_tail(exact, pct_places, pct_places + 2, "%")
        if is_count_unit(unit) and base < 50:
            res.warnings.append("基数较小（小于50），增速百分比容易放大，建议同时写出增减量")
        if abs(exact) >= 100:
            res.warnings.append("增速超过100%，可同时写“是基期的X倍”；先核对数据是否录错")
        return res

    if kind == "pp":
        cur, base = need("current", "本期比率"), need("base", "基期比率")
        # 位数默认取输入位数（最多 1 位）：用户给"71%"就写"71%"，不改成"71.0%"
        if places is None:
            pct_places = min(input_places(cur, base), 1)
        # 百分点用正文中显示的两个比率相减，读者能直接对上（data-expression.md 第十部分第 3 条）
        if decimals_of(cur) > pct_places or decimals_of(base) > pct_places:
            res.warnings.append(f"比率按 {pct_places} 位小数修约后再相减，与正文显示一致")
            cur = Decimal(fmt_num(cur, pct_places))
            base = Decimal(fmt_num(base, pct_places))
        res.inputs = {"current": fmt_num(cur) + "%", "base": fmt_num(base) + "%"}
        res.basis = a.basis or ""
        diff = cur - base
        res.value, res.unit, res.places = diff, "个百分点", pct_places
        res.formula = f"{fmt_num(cur)}%－{fmt_num(base)}%" + _approx_tail(diff, pct_places, pct_places + 2, "个百分点")
        if base != 0:
            rel = diff / base * 100
            res.extra["relative_change"] = {"value": float(rel), "display": f"{fmt_num(rel, 1 if places is None else places)}%",
                                            "note": "比率自身的相对变化，一般不写进正文，避免与百分点混淆"}
        return res

    if kind == "share":
        part, total = need("part", "部分数"), need("total", "总数")
        res.inputs = {"part": fmt_num(part), "total": fmt_num(total)}
        if total == 0:
            raise CalcError("总数为0，不能计算占比")
        exact = part / total * 100
        res.value, res.unit, res.places = exact, "%", pct_places
        res.formula = f"{fmt_raw(part)}÷{fmt_raw(total)}×100%" + _approx_tail(exact, pct_places, pct_places + 2, "%")
        if exact > 100 or exact < 0:
            res.warnings.append("占比超出0～100%，核对部分与总数的口径是否一致")
        return res

    if kind == "pull":
        pc, pb, tb = need("part_current", "部分本期数"), need("part_base", "部分基期数"), need("total_base", "总量基期数")
        res.inputs = {"part_current": fmt_num(pc), "part_base": fmt_num(pb), "total_base": fmt_num(tb)}
        if tb == 0:
            raise CalcError("总量基期数为0，不能计算拉动")
        exact = (pc - pb) / tb * 100
        res.value, res.unit, res.places = exact, "个百分点", pct_places
        res.formula = f"（{fmt_raw(pc)}－{fmt_raw(pb)}）÷{fmt_raw(tb)}×100" + _approx_tail(exact, pct_places, pct_places + 2, "个百分点")
        return res

    if kind == "contrib":
        pc, pb = need("part_current", "部分本期数"), need("part_base", "部分基期数")
        tc, tb = need("total_current", "总量本期数"), need("total_base", "总量基期数")
        res.inputs = {"part_current": fmt_num(pc), "part_base": fmt_num(pb), "total_current": fmt_num(tc),
                      "total_base": fmt_num(tb)}
        if tc == tb:
            raise CalcError("总量没有变化（增量为0），不能计算贡献率")
        exact = (pc - pb) / (tc - tb) * 100
        res.value, res.unit, res.places = exact, "%", pct_places
        res.formula = (f"（{fmt_raw(pc)}－{fmt_raw(pb)}）÷（{fmt_raw(tc)}－{fmt_raw(tb)}）×100%"
                       + _approx_tail(exact, pct_places, pct_places + 2, "%"))
        if tc < tb:
            res.warnings.append("总量下降时贡献率的含义与增长时相反，正文慎用，建议直接写各部分的增减额")
        return res

    if kind == "completion":
        actual, target = need("actual", "实际完成数"), need("target", "目标数")
        res.inputs = {"actual": fmt_num(actual), "target": fmt_num(target)}
        if target <= 0:
            raise CalcError("目标数须大于0")
        exact = actual / target * 100
        res.value, res.unit, res.places = exact, "%", pct_places
        res.formula = f"{fmt_raw(actual)}÷{fmt_raw(target)}×100%" + _approx_tail(exact, pct_places, pct_places + 2, "%")
        if a.elapsed is not None:
            elapsed = parse_input_number(a.elapsed, "已过时间")
            span = parse_input_number(a.span or 12, "全期时间")
            if span <= 0:
                raise CalcError("--span 须大于0")
            progress = elapsed / span * 100
            gap = Decimal(fmt_num(exact, pct_places)) - Decimal(fmt_num(progress, pct_places))
            res.extra["time_progress"] = {
                "value": float(progress), "display": f"{fmt_num(progress, pct_places)}%",
                "formula": f"{fmt_num(elapsed)}÷{fmt_num(span)}×100%" + _approx_tail(progress, pct_places, pct_places + 2, "%"),
                "gap": float(gap), "gap_display": f"{fmt_num(abs(gap), pct_places)}个百分点",
                "gap_formula": f"{fmt_num(exact, pct_places)}%－{fmt_num(progress, pct_places)}%＝{fmt_num(gap, pct_places)}个百分点"}
        return res

    if kind == "multiple":
        cur, base = need("current", "本期数"), need("base", "基期数")
        res.inputs = {"current": fmt_num(cur), "base": fmt_num(base)}
        if base <= 0 or cur < 0:
            raise CalcError("倍数要求基期数大于0、本期数不小于0")
        exact = cur / base
        mplaces = (1 if exact >= 1 else 2) if places is None else places
        res.value, res.unit, res.places = exact, "倍", mplaces
        res.formula = f"{fmt_raw(cur)}÷{fmt_raw(base)}" + _approx_tail(exact, mplaces, mplaces + 2, "倍")
        growth = (exact - 1) * 100
        res.extra["growth"] = {"value": float(growth), "display": f"{fmt_num(growth, pct_places)}%"}
        if cur > 0:
            fan = Decimal(str(math.log2(float(exact))))
            res.extra["doublings"] = {"value": float(fan), "display": fmt_num(fan, 2),
                                      "formula": f"log2（{fmt_raw(cur)}÷{fmt_raw(base)}）＝{fmt_num(fan, 2)}"}
        if exact < 1:
            res.warnings.append("本期小于基期，不说“减少X倍”，写“下降X%”或“相当于基期的X%”")
        return res

    if kind == "cagr":
        cur, base = need("current", "期末数"), need("base", "期初数")
        years = need("years", "间隔年数")
        res.inputs = {"current": fmt_num(cur), "base": fmt_num(base), "years": fmt_num(years)}
        if base <= 0 or cur <= 0 or years <= 0:
            raise CalcError("年均增速要求期初数、期末数和年数都大于0")
        exact = Decimal(str((float(cur) / float(base)) ** (1 / float(years)) - 1)) * 100
        res.value, res.unit, res.places = exact, "%", pct_places
        res.formula = (f"（{fmt_raw(cur)}÷{fmt_raw(base)}）^（1÷{fmt_num(years)}）－1"
                       + _approx_tail(exact, pct_places, pct_places + 2, "%"))
        res.warnings.append("年均增速按水平法（几何平均）计算，不是各年增速的简单平均")
        return res

    if kind in ("mean", "wmean"):
        values = parse_values(a.values, "--values 数值")
        if kind == "mean":
            exact = sum(values, Decimal(0)) / len(values)
            res.formula = f"（{'＋'.join(fmt_raw(v) for v in values)}）÷{len(values)}"
            res.warnings.append("这是简单平均；各单位规模不同时，比率类指标要用加权平均或分子分母重算")
        else:
            weights = parse_values(a.weights, "--weights 权数")
            if len(weights) != len(values):
                raise CalcError("--values 与 --weights 的个数不同")
            wsum = sum(weights, Decimal(0))
            if wsum == 0:
                raise CalcError("权数之和为0")
            exact = sum((v * w for v, w in zip(values, weights)), Decimal(0)) / wsum
            res.formula = (f"（{'＋'.join(f'{fmt_raw(v)}×{fmt_raw(w)}' for v, w in zip(values, weights))}）"
                           f"÷（{'＋'.join(fmt_raw(w) for w in weights)}）")
        mplaces = max(1, input_places(*values)) if places is None else places
        res.value, res.unit, res.places = exact, unit, mplaces
        res.formula += _approx_tail(exact, mplaces, mplaces + 2, unit)
        res.inputs = {"values": "、".join(fmt_num(v) for v in values)}
        return res

    if kind == "ratio":
        num, den = need("numerator", "分子"), need("denominator", "分母")
        scale = parse_input_number(a.scale if a.scale is not None else 1, "倍率")
        res.inputs = {"numerator": fmt_num(num), "denominator": fmt_num(den), "scale": fmt_num(scale)}
        if den == 0:
            raise CalcError("分母为0，不能计算")
        exact = num / den * scale
        rplaces = (1 if places is None else places)
        res.value, res.places = exact, rplaces
        res.unit = unit or ("%" if scale == 100 else ("‰" if scale == 1000 else ""))
        tail = "×100%" if scale == 100 else ("×1000‰" if scale == 1000 else (f"×{fmt_num(scale)}" if scale != 1 else ""))
        res.formula = f"{fmt_raw(num)}÷{fmt_raw(den)}{tail}" + _approx_tail(exact, rplaces, rplaces + 2, res.unit)
        return res

    raise CalcError(f"不支持的计算类型：{kind}")


def _sign_case(cur: Decimal, base: Decimal) -> str:
    if base < 0 <= cur:
        return "扭亏为盈" if cur > 0 else "减亏至持平"
    if base >= 0 > cur:
        return "由盈转亏"
    if base < 0 and cur < 0:
        return "减亏" if abs(cur) < abs(base) else ("增亏" if abs(cur) > abs(base) else "亏损持平")
    return ""


# ---------------------------------------------------------------------------
# 数据转文字
# ---------------------------------------------------------------------------

# 单字单位与前一个字组成的名词：去掉"××数"时只去"数"，如"受理案件数"→"受理案件"，不写成"受理案120件"
UNIT_NOUNS = {
    "案件", "事件", "信件", "文件", "稿件", "邮件", "部件", "零件", "证件", "来件", "函件", "急件", "要件",
    "病人", "工人", "老人", "个人", "家人", "军人", "证人", "名人", "病例", "案例", "事例", "实例", "特例",
    "农户", "住户", "用户", "客户", "商户", "散户", "专户", "账户", "门户", "班次", "车次", "频次", "批次",
    "架次", "场次", "名次", "层次", "档次", "病床", "专家", "国家", "商家", "店家", "事项", "立项", "款项",
}


def subject_of(name: str, unit: str) -> str:
    """指标名以复合计量词结尾时去掉，避免"出院人次6171人次"：出院人次→出院，出院人数→出院，走访户数→走访。

    单字单位不截："不良事件""受理案件"的"件"是名称的一部分，截掉会写成"不良事12件"。
    """
    if not unit or unit in ("%", "‰"):
        return name
    tail = unit + "数"
    if name.endswith(tail) and len(name) > len(tail):
        stem = name[: -len(tail)]
        if len(unit) == 1 and (stem[-1] + unit) in UNIT_NOUNS:
            return name[:-1]  # 只去"数"：受理案件数→受理案件
        return stem
    if len(unit) >= 2 and name.endswith(unit) and len(name) > len(unit):
        return name[: -len(unit)]
    return name


def phrase(res: CalcResult, a: argparse.Namespace) -> List[str]:
    """按 shared/rules/data-expression.md 生成可直接写进公文的句子，第一句为推荐写法。"""
    raw_name = a.name or "该指标"
    unit = a.unit or ""
    name = subject_of(raw_name, unit)
    period = (a.period or "") + ("，" if a.period else "")
    basis = res.basis or a.basis or ""
    base_label = a.base_label or BASIS_LABELS.get(basis, "基期")
    out: List[str] = []
    p = res.places
    t = res.type

    if t in ("growth", "change"):
        cur = Decimal(res.inputs["current"])
        base = Decimal(res.inputs["base"])
        ch = res.extra["change"]
        diff = cur - base
        dplaces = input_places(cur, base) if a.decimals is None or t == "growth" else a.decimals
        diff_txt = f"{fmt_num(abs(diff), dplaces)}{unit}"
        cur_txt = f"{fmt_num(cur, dplaces)}{unit}"
        more, less = ("增加", "减少")
        compare = basis if basis and basis.startswith(("同比", "环比")) else f"比{base_label}"
        if "sign_case" in res.extra:
            case = res.extra["sign_case"]
            loss_cur = f"{fmt_num(abs(cur), dplaces)}{unit}"
            loss_base = f"{fmt_num(abs(base), dplaces)}{unit}"
            if case == "扭亏为盈":
                out.append(f"{period}{name}{cur_txt}，{base_label}亏损{loss_base}，实现扭亏为盈。")
            elif case == "减亏至持平":
                out.append(f"{period}{name}盈亏持平，{base_label}亏损{loss_base}。")
            elif case in ("减亏", "增亏"):
                out.append(f"{period}{name}亏损{loss_cur}，比{base_label}{case}{diff_txt}。")
            elif case == "由盈转亏":
                out.append(f"{period}{name}亏损{loss_cur}，{base_label}盈利{loss_base}，由盈转亏。")
            else:
                out.append(f"{period}{name}亏损{loss_cur}，与{base_label}持平。")
            return out
        if base == 0:
            out.append(f"{period}{name}{cur_txt}，{base_label}为0。")
            return out
        if t == "change" or res.value is None:
            if diff == 0:
                out.append(f"{period}{name}{cur_txt}，与{base_label}持平。")
            else:
                out.append(f"{period}{name}{cur_txt}，比{base_label}{more if diff > 0 else less}{diff_txt}。")
            return out
        g = Decimal(fmt_num(res.value, p))
        if g == 0:
            out.append(f"{period}{name}{cur_txt}，与{base_label}持平。")
            return out
        verb = "增长" if g > 0 else "下降"
        g_txt = f"{fmt_num(abs(g), p)}%"
        out.append(f"{period}{name}{cur_txt}，{compare}{verb}{g_txt}。")
        out.append(f"{period}{name}{cur_txt}，比{base_label}{more if diff > 0 else less}{diff_txt}，{verb}{g_txt}。")
        if any("基数较小" in w for w in res.warnings):
            out.reverse()  # 基数小时推荐同时写增减量的写法
        if abs(res.value) >= 100 and cur > 0:
            m = cur / base
            out.append(f"{period}{name}{cur_txt}，是{base_label}的{fmt_num(m, 1)}倍。")
        return out

    if t == "pp":
        cur = Decimal(res.inputs["current"].rstrip("%"))
        base = Decimal(res.inputs["base"].rstrip("%"))
        rp = max(input_places(cur, base), p)
        d = Decimal(fmt_num(res.value, p))  # type: ignore[arg-type]
        speed = bool(RATE_VERB_SPEED.search(name))
        up, down = ("加快", "回落") if speed else ("提高", "下降")
        compare = basis if basis and basis.startswith(("同比", "环比")) else f"比{base_label}"
        if d == 0:
            out.append(f"{period}{name}为{fmt_num(cur, rp)}%，与{base_label}持平。")
            return out
        verb = up if d > 0 else down
        out.append(f"{period}{name}为{fmt_num(cur, rp)}%，{compare}{verb}{fmt_num(abs(d), p)}个百分点。")
        out.append(f"{period}{name}{verb}{fmt_num(abs(d), p)}个百分点，由{base_label}的{fmt_num(base, rp)}%"
                   f"{'升至' if d > 0 else '降至'}{fmt_num(cur, rp)}%。")
        return out

    if t == "share":
        part_name = a.part_name or name
        total_name = a.total_name or "总量"
        part = Decimal(res.inputs["part"])
        out.append(f"{period}{part_name}{fmt_num(part)}{unit}，占{total_name}的{res.display}。")
        out.append(f"{period}{part_name}占{total_name}的比重为{res.display}。")
        return out

    if t == "pull":
        part_name = a.part_name or name
        total_name = a.total_name or "总量"
        v = Decimal(fmt_num(res.value, p))  # type: ignore[arg-type]
        if v >= 0:
            out.append(f"{period}{part_name}拉动{total_name}增长{fmt_num(v, p)}个百分点。")
        else:
            out.append(f"{period}{part_name}拉低{total_name}增速{fmt_num(abs(v), p)}个百分点。")
        return out

    if t == "contrib":
        part_name = a.part_name or name
        total_name = a.total_name or "总量"
        out.append(f"{period}{part_name}对{total_name}增长的贡献率为{res.display}。")
        return out

    if t == "completion":
        actual = Decimal(res.inputs["actual"])
        target = Decimal(res.inputs["target"])
        c = Decimal(fmt_num(res.value, p))  # type: ignore[arg-type]
        goal = a.base_label or "年度目标"
        sentence = f"{period}{name}完成{fmt_num(actual)}{unit}，完成{goal}{fmt_num(target)}{unit}的{fmt_num(c, p)}%"
        if c > 100:
            sentence = (f"{period}{name}完成{fmt_num(actual)}{unit}，超额完成{goal}（{fmt_num(target)}{unit}），"
                        f"完成率{fmt_num(c, p)}%")
        elif c == 100:
            sentence = f"{period}{name}完成{fmt_num(actual)}{unit}，完成{goal}"
        tp = res.extra.get("time_progress")
        if tp and c < 100:
            gap = Decimal(str(tp["gap"]))
            if gap > 0:
                sentence += f"，快于时间进度{tp['gap_display']}"
            elif gap < 0:
                sentence += f"，慢于时间进度{tp['gap_display']}"
            else:
                sentence += "，与时间进度持平"
        out.append(sentence + "。")
        return out

    if t == "multiple":
        m = res.value  # type: ignore[assignment]
        cur = Decimal(res.inputs["current"])
        if m >= 1:  # type: ignore[operator]
            out.append(f"{period}{name}{fmt_num(cur)}{unit}，是{base_label}的{res.display}。")
            out.append(f"{period}{name}{fmt_num(cur)}{unit}，比{base_label}增长{res.extra['growth']['display']}。")
            fan = res.extra.get("doublings")
            if fan and abs(fan["value"] - round(fan["value"])) < 0.02 and round(fan["value"]) >= 1:
                k = int(round(fan["value"]))
                cn = "一二三四五"[k - 1] if k <= 5 else str(k)
                out.append(f"{period}{name}{fmt_num(cur)}{unit}，比{base_label}翻了{cn}番。")
        else:
            share = Decimal(str(m)) * 100
            out.append(f"{period}{name}{fmt_num(cur)}{unit}，相当于{base_label}的{fmt_num(share, 1)}%。")
            out.append(f"{period}{name}{fmt_num(cur)}{unit}，比{base_label}下降{fmt_num(abs(Decimal(str(res.extra['growth']['value']))), 1)}%。")
        return out

    if t == "cagr":
        if (res.value or 0) >= 0:
            out.append(f"{period}{name}年均增长{res.display}。")
        else:
            out.append(f"{period}{name}年均下降{fmt_num(abs(res.value), p)}%。")  # type: ignore[arg-type]
        return out

    if t in ("mean", "wmean"):
        out.append(f"{period}{name}平均为{res.display}。")
        return out

    if t == "ratio":
        out.append(f"{period}{name}为{res.display}。")
        return out
    return out


# ---------------------------------------------------------------------------
# 表格批量计算
# ---------------------------------------------------------------------------

def _usable(cell: Cell) -> bool:
    return cell.is_number and "approx" not in cell.flags


def _cell_state(cell: Cell) -> str:
    if cell.is_empty:
        return "为空"
    if "approx" in cell.flags:
        return f"为约数“{cell.text}”"
    return f"为“{mask_digits(cell.text)}”"


def scope_gaps(table: Table, rows: List[Row], cols: List[Column]) -> List[Tuple[Row, List[Column]]]:
    """明细中有的列有数、有的列没有数的行：这些行让各列合计的口径不同（如本期含某科室、基期不含）。"""
    gaps = []
    for r in rows:
        lack = [c for c in cols if not _usable(table.cell(r, c))]
        if lack and len(lack) < len(cols):
            gaps.append((r, lack))
    return gaps


def describe_gaps(table: Table, gaps: List[Tuple[Row, List[Column]]], cols: List[Column]) -> Tuple[str, str]:
    """返回（完整说明，短说明）。短说明放进句子里的【待核实】。"""
    parts, short = [], []
    for c in cols:
        hit = [r for r, lack in gaps if c in lack]
        if not hit:
            continue
        labels = [table.display_label(r.label) or f"第{r.excel_row}行" for r in hit]
        where = "，".join(f"第{r.excel_row}行{_cell_state(table.cell(r, c))}" for r in hit[:MAX_LIST])
        others = "、".join(x.name for x in cols if x is not c)
        names = "、".join(labels[:3]) + (f"等{len(labels)}行" if len(labels) > 3 else "")
        parts.append(f"{others}含{names}，{c.name}不含（{where}）")
        short.append(f"{c.name}不含{names}")
    return "；".join(parts), "，".join(short)


TOTAL_BAD = ("mismatch", "uncertain")  # 合计校验不通过（P0 不符、P1 可能修约也可能填报有误）
UNIT_FROM = {"growth": "current", "change": "current", "multiple": "current", "share": "part",
             "pull": "part_current", "contrib": "part_current", "completion": "actual"}


def table_calc(a: argparse.Namespace, with_phrase: bool) -> dict:
    """表格模式：每行算一次。合计、小计行与明细不符时，改用明细相加数计算并在句子后标【待核实】。"""
    sheets = load_sheets(a.table, a.sheet)
    guard_marking(sheets, getattr(a, "marking_ok", False))
    table = analyze(sheets[0], "表A", a.header_row)
    inspect_table(table, getattr(a, "marking_ok", False))
    bad_totals = {(c["row"], c["column"]): c for c in table.checks if c.get("status") in TOTAL_BAD}

    def col(arg: Optional[str], what: str, required: bool = True) -> Optional[Column]:
        if not arg:
            if required:
                raise CalcError(f"表格模式需要 {what}")
            return None
        c = table.column(arg)
        if c is None:
            raise CalcError(f"表中没有名为“{arg}”的列；现有列：{'、'.join(x.name for x in table.columns)}")
        return c

    name_col = col(a.name_col, "--name-col 名称列", required=False)
    if name_col is None and table.key_col is not None:
        name_col = table.columns[table.key_col]
    kind = a.type
    mapping: Dict[str, Optional[Column]] = {}
    if kind in ("growth", "change", "pp", "multiple"):
        mapping = {"current": col(a.current_col, "--current-col 本期列"), "base": col(a.base_col, "--base-col 基期列")}
    elif kind == "share":
        mapping = {"part": col(a.part_col, "--part-col 部分数列")}
    elif kind == "completion":
        mapping = {"actual": col(a.actual_col, "--actual-col 完成数列"), "target": col(a.target_col, "--target-col 目标数列")}
    elif kind == "ratio":
        mapping = {"numerator": col(a.numerator_col, "--numerator-col 分子列"),
                   "denominator": col(a.denominator_col, "--denominator-col 分母列")}
    elif kind in ("pull", "contrib"):
        mapping = {"part_current": col(a.current_col, "--current-col 本期列"),
                   "part_base": col(a.base_col, "--base-col 基期列")}
    else:
        raise CalcError(f"表格模式不支持 --type {kind}")
    cols: List[Column] = [c for c in mapping.values() if c is not None]
    # 同口径重算只对数量列有意义；比率、平均数相加没有意义
    summable = all(c.kind == "number" for c in cols)
    top_warnings: List[str] = []

    def column_total(c: Column) -> Tuple[Decimal, str]:
        """总量：表中合计行与明细相符时取合计行，否则取明细相加（并在全表提示中写明表中合计）。"""
        totals = [r for r in table.rows if r.kind == "total"]
        bad = bad_totals.get((totals[0].excel_row, c.name)) if totals else None
        if totals and table.cell(totals[0], c).is_number and bad is None:
            return table.cell(totals[0], c).value, f"取表中合计行（第{totals[0].excel_row}行）"  # type: ignore[return-value]
        value = detail_sum(table.data_rows(), c)
        count = sum(1 for r in table.data_rows() if _usable(table.cell(r, c)))
        if bad is not None:
            top_warnings.append(f"{c.name}表中合计为{bad['stated']}，与明细相加{fmt_num(value)}不符，"
                                f"总量改用明细相加数；正文写这个总量时用明细相加数并标【待核实】")
            return value, f"明细{count}行相加（表中合计{bad['stated']}与明细不符，未采用）"
        return value, f"明细{count}行相加"

    def col_sum(rows_: List[Row], c: Column) -> Decimal:
        return sum((table.cell(r, c).value for r in rows_), Decimal(0))  # type: ignore[misc]

    def detail_sum(rows_: List[Row], c: Column) -> Decimal:
        """明细相加：只加能参与计算的数（空值、占位符、约数不计）。"""
        return sum((table.cell(r, c).value for r in rows_ if _usable(table.cell(r, c))), Decimal(0))  # type: ignore[misc]

    # 合计行、小计行各自对应的明细行
    scope_of: Dict[int, List[Row]] = {}
    group: List[Row] = []
    for r in table.rows:
        if r.kind == "data":
            group.append(r)
        elif r.kind == "subtotal":
            scope_of[r.index] = list(group)
            group = []
        elif r.kind == "total":
            scope_of[r.index] = table.data_rows()

    rows = [r for r in table.rows if r.kind in ("data", "total", "subtotal")]
    share_total: Optional[Decimal] = None
    share_total_formula = ""
    totals_used: Dict[str, str] = {}
    if kind == "share":
        share_total, share_total_formula = column_total(mapping["part"])  # type: ignore[arg-type]
        rows = table.data_rows()
    if kind in ("pull", "contrib"):
        gaps = scope_gaps(table, table.data_rows(), cols)
        if gaps:
            # 两期包含的明细不同：总量改用两期都有数的明细相加，各部分拉动之和才等于同口径的总量增速
            full, _short = describe_gaps(table, gaps, cols)
            same = [r for r in table.data_rows() if all(_usable(table.cell(r, c)) for c in cols)]
            tc = col_sum(same, mapping["part_current"])  # type: ignore[arg-type]
            tb = col_sum(same, mapping["part_base"])  # type: ignore[arg-type]
            how_c = how_b = f"两期都有数的明细{len(same)}行相加，同口径"
            top_warnings.append(f"两期口径不同：{full}。总量已按同口径（剔除这些行）计算，"
                                f"正文写明“不含……”，或补齐后重算")
        else:
            tc, how_c = column_total(mapping["part_current"])  # type: ignore[arg-type]
            tb, how_b = column_total(mapping["part_base"])  # type: ignore[arg-type]
        totals_used = {"total_current": f"{fmt_num(tc)}（{how_c}）", "total_base": f"{fmt_num(tb)}（{how_b}）"}
        a.total_current, a.total_base = tc, tb
        rows = table.data_rows()

    def row_label(row: Row) -> str:
        label = table.display_label(row.cells[name_col.index].text) if name_col is not None else ""
        if not label:
            label = table.display_label(row.label) or f"第{row.excel_row}行"
        if row.kind == "total" and len(label) > 2 and label.endswith(("合计", "总计")):
            label = label[:-2]
        return label

    def fill_unit(ns: argparse.Namespace) -> None:
        """未给 --unit 时取所用列的单位（如"本年累计（万元）"取万元），句子和底稿的"文字表述"都带单位。"""
        source = mapping.get(UNIT_FROM.get(kind, ""))
        if not ns.unit and source is not None and source.unit not in ("%", "‰"):
            ns.unit = source.unit

    def same_scope_item(row: Row, label: str, scope: List[Row], gaps: List[Tuple[Row, List[Column]]],
                        warnings: List[str]) -> Optional[Dict[str, Any]]:
        """剔除口径不一致的明细后，用两边都有数的明细相加重算一次。"""
        same = [r for r in scope if all(_usable(table.cell(r, c)) for c in cols)]
        if not summable or not same:
            return None
        dropped = [table.display_label(r.label) or f"第{r.excel_row}行" for r, _ in gaps]
        dropped_txt = "、".join(dropped[:3]) + (f"等{len(dropped)}行" if len(dropped) > 3 else "")
        ns2 = argparse.Namespace(**vars(a))
        ns2.name = f"{label}（不含{dropped_txt}）" + (a.name or "")
        for attr, c in mapping.items():
            setattr(ns2, attr, col_sum(same, c))  # type: ignore[arg-type]
        fill_unit(ns2)
        try:
            res2 = calc(kind, ns2, wording=with_phrase)
        except CalcError as exc:
            warnings.append(f"同口径重算未完成：{exc}")
            return None
        sums = "，".join(f"{c.name}{fmt_num(col_sum(same, c))}" for c in cols)
        item2: Dict[str, Any] = {"row": row.excel_row, "name": ns2.name, "row_kind": "same_scope"}
        item2.update(res2.to_json())
        item2["warnings"] = list(res2.warnings)
        item2["formula"] = f"剔除{dropped_txt}，其余{len(same)}行明细相加：{sums}；{res2.formula}"
        if with_phrase:
            item2["sentences"] = phrase(res2, ns2)
        return item2

    results = []
    for row in rows:
        ns = argparse.Namespace(**vars(a))
        label = row_label(row)
        ns.name = f"{label}{a.name or ''}" if a.name else label
        skip = ""
        fixes: List[str] = []  # 合计、小计与明细不符的列：改用明细相加数
        kind_label = "合计" if row.kind == "total" else "小计"
        for attr, c in mapping.items():
            cell = table.cell(row, c)  # type: ignore[arg-type]
            if not _usable(cell):
                skip = f"{c.name}为“{cell.text or '空'}”，不能计算"  # type: ignore[union-attr]
                break
            value = cell.value
            if row.kind in ("total", "subtotal") and (row.excel_row, c.name) in bad_totals:  # type: ignore[union-attr]
                value = detail_sum(scope_of.get(row.index, []), c)  # type: ignore[arg-type]
                fixes.append(f"{c.name}表中{kind_label}为{fmt_num(cell.value)}，明细相加为{fmt_num(value)}")  # type: ignore[union-attr, arg-type]
            setattr(ns, attr, value)
        fill_unit(ns)
        if kind == "share":
            ns.total = share_total
        item: Dict[str, Any] = {"row": row.excel_row, "name": ns.name, "row_kind": row.kind}
        if skip:
            item["skipped"] = skip
            item["sentence"] = f"{ns.name}【待补：{skip}】" if with_phrase else ""
            results.append(item)
            continue
        try:
            res = calc(kind, ns, wording=with_phrase)
        except CalcError as exc:
            item["skipped"] = str(exc)
            results.append(item)
            continue
        item.update(res.to_json())
        item["warnings"] = list(res.warnings)
        extra_item: Optional[Dict[str, Any]] = None
        verify: List[str] = []
        if fixes:
            # 表中合计、小计只放进提示；结果、句子、底稿都用明细相加数，并带【待核实】
            item["warnings"].append(f"表中{kind_label}与明细不符：{'；'.join(fixes)}。已改用明细相加数计算，"
                                    f"句子带【待核实】；查清前不直接引用表中{kind_label}")
            item["formula"] = f"用明细相加数（{'；'.join(fixes)}）：{res.formula}"
            item["total_fixes"] = fixes
            verify.extend(fixes)
        if row.kind in ("total", "subtotal"):
            scope = scope_of.get(row.index, [])
            gaps = scope_gaps(table, scope, cols) if len(cols) > 1 else []
            if gaps:
                full, short = describe_gaps(table, gaps, cols)
                verify.insert(0, f"两期口径不同，{short}")
                item["warnings"].append(f"明细口径不一致：{full}。该行结果不能直接引用；请两边同口径剔除后计算"
                                        f"（见紧随其后的“同口径”结果），或在正文标【待核实】")
                item["scope_gaps"] = [{"row": r.excel_row, "label": table.display_label(r.label),
                                       "missing": [c.name for c in lack]} for r, lack in gaps]
                extra_item = same_scope_item(row, label, scope, gaps, item["warnings"])
        if with_phrase:
            sentences = phrase(res, ns)
            if verify:
                # 合计不符或口径不同的合计行：句子带【待核实】，粘进正文也不会被当作确数
                mark = "；".join(verify)
                sentences = [(s[:-1] if s.endswith("。") else s) + f"【待核实：{mark}】。" for s in sentences]
            item["sentences"] = sentences
        results.append(item)
        if extra_item is not None:
            results.append(extra_item)

    # 每行都相同的提示（如未写基期）只在顶部说一次
    computed = [r for r in results if "skipped" not in r]
    if len(computed) > 1:
        common = [w for w in computed[0].get("warnings", []) if all(w in r.get("warnings", []) for r in computed)]
        for r in computed:
            r["warnings"] = [w for w in r.get("warnings", []) if w not in common]
        top_warnings = common + top_warnings
    data = {"table": table.source, "type": kind, "results": results, "warnings": top_warnings,
            "table_issues": [i.to_json() for i in table.issues if i.level == "P0"]}
    if kind == "share":
        data["total"] = {"value": fmt_num(share_total), "how": share_total_formula}  # type: ignore[arg-type]
    if totals_used:
        data["totals"] = totals_used
    return data


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------

def count_by(table: Table, by: List[str], unit: str = "件") -> dict:
    """分类计数：按一列或两列分组统计行数和占比，用于台账、明细表（如不良事件按类型统计）。"""
    cols: List[Column] = []
    for name in by:
        c = table.column(name)
        if c is None:
            raise InputError(f"表中没有名为“{name}”的列；现有列：{'、'.join(x.name for x in table.columns)}")
        if c.sensitive:
            raise InputError(f"“{c.name}”是个人信息列，不按它分组统计")
        cols.append(c)
    rows = table.data_rows()
    counter: Counter = Counter()
    blank: List[str] = []
    for row in rows:
        values = tuple(norm_text(table.cell(row, c).text) for c in cols)
        if any(not v for v in values):
            blank.append(f"第{row.excel_row}行")
            continue
        counter[values] += 1
    total = sum(counter.values())
    if total == 0:
        raise InputError("没有可以计数的行")
    items = []
    for values, n in sorted(counter.items(), key=lambda kv: (-kv[1], kv[0])):
        share = Decimal(n) / Decimal(total) * 100
        label = "、".join(values)
        items.append({
            "values": list(values), "count": n, "share": fmt_num(share, 1),
            "formula": f"{n}÷{total}×100%" + _approx_tail(share, 1, 3, "%"),
            "sentence": f"{label}{n}{unit}，占{fmt_num(share, 1)}%",
        })
    share_sum = sum(Decimal(i["share"]) for i in items)
    data = {"by": [c.name for c in cols], "total": total, "unit": unit, "items": items,
            "share_sum": fmt_num(share_sum, 1)}
    notes = []
    if blank:
        notes.append(f"{len(blank)}行分组列为空，未计入：" + _limit(blank))
    if share_sum != 100 and len(cols) == 1:
        notes.append(f"各项占比修约后相加为{fmt_num(share_sum, 1)}%，表注写“因四舍五入，各项占比之和不等于100%”")
    data["notes"] = notes
    return data


def print_json(data: Any) -> None:
    print(json.dumps(data, ensure_ascii=False, indent=2))


def render_inspect_text(tables: List[dict]) -> str:
    lines = []
    stopped = [t for t in tables if t.get("stopped")]
    if stopped:
        # 第一行就是停止提示，模型和用户一眼能看到
        lines.append("【停止处理】DT-SEC-01 " + stopped[0]["issues"][0]["message"])
    for t in tables:
        s = t["summary"]
        lines.append(f"== {t['label']}：{t['source']} ==")
        if t.get("stopped"):
            lines.append("密级或内部标注：" + "；".join(f"{m['location']}“{m['text']}”" for m in t["marks"]))
            lines.append("未输出标题、表注、统计数和其他问题。")
            continue
        if t["title"]:
            lines.append(f"标题：{t['title']}")
        if t.get("pre_texts"):
            lines.append(f"表头前说明：{'；'.join(t['pre_texts'])}")
        lines.append(f"统计时段：{'、'.join(t['periods']) or '未写明'}；表注单位：{t['unit_note'] or '无'}；"
                     f"数据行 {t['data_rows']} 行")
        lines.append("列：" + "；".join(f"{c['letter']} {c['name']}（{c['kind']}{'，' + c['unit'] if c['unit'] else ''}）"
                                      for c in t["columns"]))
        if t.get("notes"):
            lines.append("表注：" + "；".join(re.sub(r"^注\s*[:：]\s*", "", n) for n in t["notes"]))
        checks = [c for c in t["checks"] if c.get("status") in ("ok", "mismatch", "rounding", "uncertain")]
        if checks:
            counts = Counter(c["status"] for c in checks)
            parts = [f"相符 {counts['ok']} 项"]
            for status, word in (("mismatch", "不符"), ("uncertain", "超出修约容差待核"), ("rounding", "修约误差")):
                if counts[status]:
                    parts.append(f"{word} {counts[status]} 项")
            lines.append(f"合计校验：{len(checks)} 项，{'，'.join(parts)}")
        elif any(c.get("status") == "none" for c in t["checks"]):
            lines.append("合计校验：表中没有合计行，无法对账")
        skipped: List[str] = []
        for c in t["checks"]:
            if c.get("status") == "skipped":
                text = f"{c['column']}（{c['reason']}）"
                if text not in skipped:
                    skipped.append(text)
        if skipped:
            lines.append(f"未校验：{'；'.join(skipped)}")
        lines.append(f"问题：P0 {s['P0']} 条，P1 {s['P1']} 条，P2 {s['P2']} 条")
        for i, issue in enumerate(t["issues"], start=1):
            lines.append(f"  {i}. [{issue['level']}] {issue['category']}｜{issue['location']}｜{issue['message']}"
                         f"{'｜建议：' + issue['suggestion'] if issue['suggestion'] else ''}")
    return "\n".join(lines)


def render_merge_text(result: dict) -> str:
    lines = [f"布局：{'记录式（每行一个单位或事项）' if result['layout'] == 'record' else '指标式（每行一个指标）'}；"
             f"基准时段：{result['baseline_period'] or '未写明'}"]
    for t in result["tables"]:
        flag = "（未计入汇总）" if t["excluded"] else ""
        lines.append(f"{t['label']}：{t['source']}，时段 {t['period'] or '未写明'}，数据行 {t['data_rows']}{flag}")
    lines.append("汇总：")
    for item in result["aggregate"]:
        parts = []
        for name, v in item["values"].items():
            if v["value"] is None:
                parts.append(f"{name}（{v['note']}）")
            else:
                parts.append(f"{name} {v['value']}" + (f"（{v['note']}）" if v.get("note") else ""))
        lines.append(f"  {item['key']}：" + "；".join(parts))
    if result["conversions"]:
        lines.append("单位换算：")
        for c in result["conversions"][:MAX_LIST]:
            lines.append(f"  {c['table']}第{c['row']}行 {c['column']}：{c['formula']}")
    s = result["summary"]
    lines.append(f"问题：P0 {s['P0']} 条，P1 {s['P1']} 条，P2 {s['P2']} 条")
    for i, issue in enumerate(result["issues"], start=1):
        lines.append(f"  {i}. [{issue['level']}] {issue['category']}｜{issue['location']}｜{issue['message']}")
    if result.get("out"):
        lines.append(f"已写出：{result['out']}")
    return "\n".join(lines)


def ledger_md(items: List[dict]) -> str:
    """计算底稿表格行：序号｜指标｜算式｜结果｜文字表述｜来源。"""
    lines = ["| 序号 | 指标 | 算式 | 结果 | 文字表述 | 来源 |", "|---|---|---|---|---|---|"]
    for i, item in enumerate(items, start=1):
        sentence = (item.get("sentences") or [""])[0].rstrip("。")
        lines.append(f"| {i} | {item.get('name', '')} | {item.get('formula', '')} | {item.get('display', '')} | "
                     f"{sentence} | 【待补：表名和行号】 |")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 命令行
# ---------------------------------------------------------------------------

class ChineseHelpFormatter(argparse.RawDescriptionHelpFormatter):
    """--help 输出中文："用法："。"""

    def add_usage(self, usage, actions, groups, prefix=None):
        return super().add_usage(usage, actions, groups, "用法：" if prefix is None else prefix)


def localize(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """分组标题和 -h 说明改为中文（argparse 没有公开接口，只能改属性）。"""
    parser._optionals.title = "选项"
    parser._positionals.title = "参数"
    parser.add_argument("-h", "--help", action="help", help="显示本帮助信息并退出")
    return parser


def build_parser() -> argparse.ArgumentParser:
    parser = localize(argparse.ArgumentParser(
        prog="data_tool.py",
        description="公文数据整理：报表体检、分类计数、多表合并、指标计算、数据转文字。不联网，不修改输入文件。",
        epilog="退出码：0 正常；1 --strict 时存在 P0 问题；2 输入错误。",
        formatter_class=ChineseHelpFormatter,
        add_help=False,
    ))
    sub = parser.add_subparsers(dest="command", metavar="子命令", title="子命令")
    _add = sub.add_parser

    def add_sub(name: str, **kwargs) -> argparse.ArgumentParser:
        return localize(_add(name, add_help=False, formatter_class=ChineseHelpFormatter, **kwargs))

    sub.add_parser = add_sub  # type: ignore[assignment]

    p = sub.add_parser("inspect", help="报表体检：表头、空值、合并单元格、文本型数字、单位、合计校验、重复、个人信息")
    p.add_argument("file", help="表格文件（.csv 或 .xlsx）")
    p.add_argument("--sheet", help="指定工作表名称（默认读取活动工作表）")
    p.add_argument("--all-sheets", action="store_true", help="逐个体检全部工作表")
    p.add_argument("--header-row", type=int, help="表头所在行号（从 1 起），自动识别不准时使用")
    p.add_argument("--key", help="关键列名称（如“街道”“科室”），用于查重和行标签；默认取前三列中不是分组列的文字列")
    p.add_argument("--json", action="store_true", help="输出 JSON")
    p.add_argument("--strict", action="store_true", help="存在 P0 问题时返回 1")
    p.add_argument("--marking-ok", action="store_true", help=MARKING_OK_HELP)

    k = sub.add_parser("count", help="分类计数：按一列或两列统计行数和占比（如不良事件按类型、信访按类别）")
    k.add_argument("file", help="表格文件（.csv 或 .xlsx）")
    k.add_argument("--by", required=True, help="分组列名；两列用逗号分隔，如“事件类型,事件分级”")
    k.add_argument("--unit", default="件", help="计数单位（默认“件”）")
    k.add_argument("--sheet", help="工作表名称")
    k.add_argument("--header-row", type=int, help="表头所在行号（从 1 起）")
    k.add_argument("--json", action="store_true", help="输出 JSON")
    k.add_argument("--marking-ok", action="store_true", help=MARKING_OK_HELP)

    m = sub.add_parser("merge", help="多表合并：对齐表头、统一单位、标出时段和口径冲突，汇总下级报表")
    m.add_argument("files", nargs="+", help="多个表格文件；一个 .xlsx 内多个工作表时配合 --all-sheets")
    m.add_argument("--all-sheets", action="store_true", help="把每个工作表当作一张下级报表")
    m.add_argument("--key", help="关键列名称（如“科室”“指标名称”），默认取各表第一列文字列")
    m.add_argument("--layout", choices=("auto", "record", "indicator"), default="auto",
                   help="record：每行一个单位或事项，按列汇总；indicator：每行一个指标，按指标汇总；默认自动判断")
    m.add_argument("--expect", help="应报单位名单：文件（每行一个）或用逗号分隔的名称")
    m.add_argument("--header-row", type=int, help="各表表头所在行号（从 1 起）")
    m.add_argument("--out", help="写出汇总表（.xlsx 或 .csv），已存在时加 -v2，不覆盖")
    m.add_argument("--json", action="store_true", help="输出 JSON")
    m.add_argument("--strict", action="store_true", help="存在 P0 问题时返回 1")
    m.add_argument("--marking-ok", action="store_true", help=MARKING_OK_HELP)

    for name, helptext in (("calc", "指标计算：输出结果和算式"), ("phrase", "数据转文字：在计算结果之外输出可直接写进公文的句子")):
        c = sub.add_parser(name, help=helptext)
        c.add_argument("--type", required=True,
                       choices=("growth", "change", "pp", "share", "pull", "contrib", "completion", "multiple",
                                "cagr", "mean", "wmean", "ratio"),
                       help="计算类型，见本脚本说明")
        c.add_argument("--name", help="指标名称，如“出院人次”")
        c.add_argument("--unit", help="计量单位，如“人次”“万元”")
        c.add_argument("--period", help="统计时段，如“2026年1—9月”")
        c.add_argument("--basis", help="对比方式：同比、环比、比上年、比年初等")
        c.add_argument("--base-label", help="基期的叫法，如“上年同期”“上月”“2021年”")
        c.add_argument("--current", help="本期数（pp 时为本期比率）")
        c.add_argument("--base", help="基期数（pp 时为基期比率）")
        c.add_argument("--part", help="share：部分数")
        c.add_argument("--total", help="share：总数")
        c.add_argument("--part-current", help="pull、contrib：部分本期数")
        c.add_argument("--part-base", help="pull、contrib：部分基期数")
        c.add_argument("--total-current", help="contrib：总量本期数")
        c.add_argument("--total-base", help="pull、contrib：总量基期数")
        c.add_argument("--part-name", help="share、pull、contrib：部分的名称")
        c.add_argument("--total-name", help="share、pull、contrib：总量的名称")
        c.add_argument("--actual", help="completion：实际完成数")
        c.add_argument("--target", help="completion：目标数")
        c.add_argument("--elapsed", help="completion：已过时间（如 9 个月）")
        c.add_argument("--span", help="completion：全期时间（默认 12）")
        c.add_argument("--years", help="cagr：间隔年数（如 2021 年到 2026 年为 5）")
        c.add_argument("--values", help="mean、wmean：数值，逗号分隔")
        c.add_argument("--weights", help="wmean：权数，逗号分隔")
        c.add_argument("--numerator", help="ratio：分子")
        c.add_argument("--denominator", help="ratio：分母")
        c.add_argument("--scale", help="ratio：倍率，100 表示百分比（默认 1）")
        c.add_argument("--decimals", type=int, help="保留小数位数（默认：百分比和百分点 1 位）")
        c.add_argument("--table", help="表格模式：对表中每一行计算")
        c.add_argument("--sheet", help="表格模式：工作表名称")
        c.add_argument("--header-row", type=int, help="表格模式：表头所在行号")
        c.add_argument("--name-col", help="表格模式：名称列")
        c.add_argument("--current-col", help="表格模式：本期列")
        c.add_argument("--base-col", help="表格模式：基期列")
        c.add_argument("--part-col", help="表格模式（share）：部分数列，总数取合计行，没有合计行时取明细求和")
        c.add_argument("--actual-col", help="表格模式（completion）：完成数列")
        c.add_argument("--target-col", help="表格模式（completion）：目标数列")
        c.add_argument("--numerator-col", help="表格模式（ratio）：分子列")
        c.add_argument("--denominator-col", help="表格模式（ratio）：分母列")
        c.add_argument("--json", action="store_true", help="输出 JSON")
        c.add_argument("--marking-ok", action="store_true", help=MARKING_OK_HELP)
        if name == "phrase":
            c.add_argument("--md", action="store_true", help="另输出计算底稿表格行（Markdown）")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    setup_console()
    parser = build_parser()
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:
        return 0 if exc.code == 0 else 2
    if not args.command:
        parser.print_help()
        return 2
    try:
        if args.command == "inspect":
            sheets = load_sheets(args.file, args.sheet, args.all_sheets)
            results = []
            stopped = any(s.marks for s in sheets) and not args.marking_ok
            for i, sheet in enumerate(sheets):
                if stopped:
                    # 带密级或内部标注：整个文件停止处理，只报标注位置
                    if sheet.marks:
                        results.append(stopped_summary(sheet, table_letter(i)))
                    continue
                table = analyze(sheet, table_letter(i), args.header_row, args.key)
                inspect_table(table, args.marking_ok)
                results.append(table_summary(table))
            data = {"tool": "data_tool", "version": VERSION, "command": "inspect", "stopped": stopped,
                    "tables": results, "summary": {lv: sum(t["summary"][lv] for t in results) for lv in LEVELS}}
            if args.json:
                print_json(data)
            else:
                print(render_inspect_text(results))
            return 1 if args.strict and data["summary"]["P0"] else 0

        if args.command == "count":
            sheets = load_sheets(args.file, args.sheet)
            guard_marking(sheets, args.marking_ok)
            table = analyze(sheets[0], "表A", args.header_row)
            inspect_table(table, args.marking_ok)
            by = [b.strip() for b in re.split(r"[,，、]", args.by) if b.strip()]
            data = {"tool": "data_tool", "version": VERSION, "command": "count", "source": table.source,
                    **count_by(table, by, args.unit)}
            if args.json:
                print_json(data)
            else:
                print(f"按{'、'.join(data['by'])}计数，共{data['total']}{data['unit']}：")
                for item in data["items"]:
                    print(f"  {item['sentence']}　算式：{item['formula']}")
                for note in data["notes"]:
                    print(f"提示：{note}")
            return 0

        if args.command == "merge":
            tables: List[Table] = []
            loaded = [s for path in args.files for s in load_sheets(path, None, args.all_sheets)]
            guard_marking(loaded, args.marking_ok)  # 任何一张表带标注，整批不合并
            for sheet in loaded:
                table = analyze(sheet, table_letter(len(tables)), args.header_row, args.key)
                inspect_table(table, args.marking_ok)
                tables.append(table)
            if len(tables) < 2:
                raise InputError("合并至少需要两张表；只有一张表时请用 inspect")
            result = merge_tables(tables, args.key, args.layout, read_expect(args.expect))
            result = {"tool": "data_tool", "version": VERSION, "command": "merge", **result}
            if args.out:
                result["out"] = str(write_merge_output(result, Path(args.out)))
            if args.json:
                print_json(result)
            else:
                print(render_merge_text(result))
            return 1 if args.strict and result["summary"]["P0"] else 0

        with_phrase = args.command == "phrase"
        if args.table:
            data = table_calc(args, with_phrase)
            data = {"tool": "data_tool", "version": VERSION, "command": args.command, **data}
            if args.json:
                print_json(data)
            else:
                for item in data["results"]:
                    if "skipped" in item:
                        print(f"第{item['row']}行 {item['name']}：未计算（{item['skipped']}）")
                        continue
                    line = f"第{item['row']}行 {item['name']}：{item['display']}　算式：{item['formula']}"
                    if with_phrase:
                        line += f"\n    {item['sentences'][0] if item['sentences'] else ''}"
                    for w in item.get("warnings", []):
                        line += f"\n    提示：{w}"
                    print(line)
                for w in data.get("warnings", []):
                    print(f"提示（全表）：{w}")
                if data.get("table_issues"):
                    print("注意：表中有 P0 问题，请先运行 inspect 查看。")
            if with_phrase and getattr(args, "md", False):
                print(ledger_md([r for r in data["results"] if "skipped" not in r]))
            return 0

        res = calc(args.type, args, wording=with_phrase)
        data = {"tool": "data_tool", "version": VERSION, "command": args.command, **res.to_json()}
        if with_phrase:
            data["sentences"] = phrase(res, args)
        if args.json:
            print_json(data)
        else:
            print(f"结果：{res.display or '不计算'}")
            print(f"算式：{res.formula}")
            if "change" in res.extra and res.type == "growth":
                print(f"增减量：{res.extra['change']['formula']}")
            if "time_progress" in res.extra:
                tp = res.extra["time_progress"]
                print(f"时间进度：{tp['formula']}；差距：{tp['gap_formula']}")
            for w in res.warnings:
                print(f"提示：{w}")
            if with_phrase:
                for i, s in enumerate(data["sentences"], start=1):
                    print(f"写法{i}：{s}")
        if with_phrase and getattr(args, "md", False):
            print(ledger_md([data]))
        return 0
    except InputError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    except OSError as exc:
        print(f"错误：无法读写文件（{exc.strerror or exc}）：{exc.filename or ''}。"
              f"请检查路径是否存在、文件是否被 Excel 占用；Windows 下路径超过 260 个字符时请换到短一些的目录",
              file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
