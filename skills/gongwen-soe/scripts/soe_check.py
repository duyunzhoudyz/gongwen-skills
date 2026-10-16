#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""国有企业专用文书的行业机检（gongwen-soe）。

在通用机检 gongwen_lint.py 之外，补查国企专用文书中能写成规则的问题：
  - 决策程序：先执行后补程序、追认、职代会后置、直接提交决策机构、党委会“决定”经营事项、
    拆分规避集体决策、经理层超授权（SOE-001～007）；
  - 对外表述：无出处的“第一”“领先”、绝对化环境声明和合规结论（SOE-008～010）；
  - 对外材料中的未公开信息、商业秘密、经营数据（SOE-011～013，只在对外材料中检查）；
  - 数据表述：比率变化用百分数、同比环比用混（SOE-014、015）；党建成效归因（SOE-016）；
  - 函数检查：议题要件和方案比选（SOE-101、102）、议案要件（SOE-103）、决议出席人数和票数（SOE-104）、
    改革台账状态（SOE-105）、ESG 披露口径（SOE-106）、金额单位（SOE-107，制度标准金额不计）、
    纪要议定事项和前置程序（SOE-108、109）、融合材料“两张皮”（SOE-110）。

用法：
    python scripts/soe_check.py <文件> [--type 文书] [--external] [--as-of YYYY-MM-DD] [--json] [--strict]
    python scripts/soe_check.py --list-rules

输入 .txt、.md、.docx（docx 需要 python-docx），遇到单独成行的【起草说明】即停止读取。
输出字段与 gongwen_lint.py 一致（file、doc_type、char_count、summary、issues），可以并入核稿意见。
退出码：0 正常；--strict 且存在 P0 时为 1；输入错误为 2。
只依赖标准库；正则规则在同目录的 soe_rules.json 中，修改规则不必改代码。不访问网络，不写任何文件。
"""
from __future__ import annotations

import argparse
import calendar
import datetime as _dt
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Pattern, Sequence, Tuple

VERSION = "1.0.0"
RULES_PATH = Path(__file__).resolve().parent / "soe_rules.json"
LEVELS = ("P0", "P1", "P2")

# 【起草说明】分界行，与 gongwen_lint.py、fact_diff.py 的写法一致
DRAFT_NOTE_RE = re.compile(r"^\s*(?:#{1,6}\s*)?(?:(?:\*\*|__)\s*)?【起草说明】")
SENTENCE_END = "。！？；!?;\n"

Lines = Sequence[Tuple[int, str]]


class InputError(Exception):
    """输入错误，退出码 2。"""


@dataclass
class Rule:
    id: str
    category: str
    gate: str
    level: str
    message: str
    suggestion: str
    basis: str
    types: Tuple[str, ...] = ()
    regex: Optional[Pattern[str]] = None
    external_level: Optional[str] = None
    external_only: bool = False
    exempt: Optional[Pattern[str]] = None
    exempt_match: Optional[Pattern[str]] = None
    hit_examples: List[str] = field(default_factory=list)
    miss_examples: List[str] = field(default_factory=list)

    def applies(self, doc_type: Optional[str]) -> bool:
        """types 为空时所有文书都查；非空时只在指定文书类型下查。"""
        if not self.types:
            return True
        return doc_type in self.types


@dataclass
class RuleSet:
    types: Tuple[str, ...]
    external_types: Tuple[str, ...]
    rules: List[Rule]
    procedural: Dict[str, Rule]


# ---------------------------------------------------------------------------
# 规则加载
# ---------------------------------------------------------------------------

def _compile(pattern: Optional[str], rid: str, name: str) -> Optional[Pattern[str]]:
    if not pattern:
        return None
    try:
        return re.compile(pattern)
    except re.error as exc:
        raise InputError(f"规则 {rid} 的 {name} 不是合法的正则表达式：{exc}") from exc


def _level(value: Any, rid: str) -> str:
    if value not in LEVELS:
        raise InputError(f"规则 {rid} 的级别 {value!r} 无效，只能是 P0、P1、P2")
    return str(value)


def _build_rule(item: Dict[str, Any], procedural: bool) -> Rule:
    rid = str(item.get("id") or "").strip()
    if not rid:
        raise InputError("规则缺少 id")
    for key in ("category", "gate", "message", "basis"):
        if not item.get(key):
            raise InputError(f"规则 {rid} 缺少 {key}")
    rule = Rule(
        id=rid,
        category=item["category"],
        gate=item["gate"],
        level=_level(item.get("level"), rid),
        message=item["message"],
        suggestion=item.get("suggestion", ""),
        basis=item["basis"],
        types=tuple(item.get("types") or ()),
        external_level=_level(item["external_level"], rid) if item.get("external_level") else None,
        external_only=bool(item.get("external_only")),
        exempt=_compile(item.get("exempt_context"), rid, "exempt_context"),
        exempt_match=_compile(item.get("exempt_match"), rid, "exempt_match"),
        hit_examples=list(item.get("hit_examples") or []),
        miss_examples=list(item.get("miss_examples") or []),
    )
    if not procedural:
        rule.regex = _compile(item.get("pattern"), rid, "pattern")
        if rule.regex is None:
            raise InputError(f"规则 {rid} 缺少 pattern")
    return rule


def load_rules(path: Path = RULES_PATH) -> RuleSet:
    """读取规则文件。文件缺失或格式错误时抛出 InputError。"""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise InputError(f"找不到规则文件：{path}") from exc
    except (OSError, ValueError) as exc:
        raise InputError(f"规则文件无法读取或不是合法的 JSON：{exc}") from exc
    if not isinstance(data, dict):
        raise InputError("规则文件顶层必须是对象")
    types = tuple(data.get("types") or ())
    external_types = tuple(data.get("external_types") or ())
    rules = [_build_rule(item, False) for item in data.get("rules") or []]
    procedural_list = [_build_rule(item, True) for item in data.get("procedural") or []]
    ids = [r.id for r in rules + procedural_list]
    duplicated = sorted({i for i in ids if ids.count(i) > 1})
    if duplicated:
        raise InputError(f"规则 id 重复：{'、'.join(duplicated)}")
    for rule in rules + procedural_list:
        unknown = [t for t in rule.types if t not in types]
        if unknown:
            raise InputError(f"规则 {rule.id} 的 types 含未登记的文书类型：{'、'.join(unknown)}")
    missing = [rid for rid in PROCEDURAL_IDS if rid not in {r.id for r in procedural_list}]
    if missing:
        raise InputError(f"规则文件缺少函数检查的登记：{'、'.join(missing)}")
    return RuleSet(types=types, external_types=external_types, rules=rules,
                   procedural={r.id: r for r in procedural_list})


# ---------------------------------------------------------------------------
# 输入读取
# ---------------------------------------------------------------------------

def read_lines(path: Path) -> List[Tuple[int, str]]:
    """读取文件，返回 [(行号或段落序号, 文本)]，到【起草说明】为止。"""
    if not path.is_file():
        raise InputError(f"文件不存在：{path}")
    suffix = path.suffix.lower()
    if suffix in (".txt", ".md", ".markdown", ".text"):
        raw = path.read_bytes()
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            try:
                text = raw.decode("gb18030")
            except UnicodeDecodeError as exc:
                raise InputError("无法识别文件编码，请另存为 UTF-8 后重试。") from exc
        lines = split_lines(text)
    elif suffix == ".docx":
        lines = _read_docx(path)
    else:
        raise InputError(f"不支持的文件类型：{path.suffix or '无扩展名'}。只支持 .txt、.md、.docx。")
    return cut_at_note(lines)


def split_lines(text: str) -> List[Tuple[int, str]]:
    return list(enumerate(text.replace("\r\n", "\n").replace("\r", "\n").split("\n"), start=1))


def _read_docx(path: Path) -> List[Tuple[int, str]]:
    try:
        import docx  # type: ignore
        from docx.oxml.ns import qn  # type: ignore
        from docx.text.paragraph import Paragraph  # type: ignore
    except ImportError as exc:
        raise InputError("读取 .docx 需要 python-docx，请先运行：pip install python-docx") from exc
    try:
        document = docx.Document(str(path))
    except Exception as exc:  # 损坏或加密的文件
        raise InputError(f"无法打开 docx 文件：{exc}") from exc
    lines: List[Tuple[int, str]] = []
    number = 0
    p_tag, tbl_tag = qn("w:p"), qn("w:tbl")
    for child in document.element.body.iterchildren():
        if child.tag == p_tag:
            number += 1
            lines.append((number, Paragraph(child, document).text))
        elif child.tag == tbl_tag:
            # 表格按行转成 Markdown 风格的一行，便于台账检查
            for row in child.iter(qn("w:tr")):
                cells = []
                for cell in row.iter(qn("w:tc")):
                    texts = [Paragraph(p_el, document).text for p_el in cell.iter(p_tag)]
                    cells.append(" ".join(t for t in texts if t).replace("|", "／"))
                number += 1
                lines.append((number, "| " + " | ".join(cells) + " |"))
    return _insert_docx_separators(lines)


def _insert_docx_separators(lines: List[Tuple[int, str]]) -> List[Tuple[int, str]]:
    """docx 表格的第一行之后补一行分隔行，使其能按 Markdown 表格解析。"""
    out: List[Tuple[int, str]] = []
    previous_is_row = False
    for number, text in lines:
        is_row = text.startswith("| ")
        out.append((number, text))
        if is_row and not previous_is_row:
            columns = text.count("|") - 1
            out.append((number, "|" + "---|" * max(columns, 1)))
        previous_is_row = is_row
    return out


def cut_at_note(lines: Lines) -> List[Tuple[int, str]]:
    """只保留【起草说明】之前的正文。"""
    kept: List[Tuple[int, str]] = []
    for number, text in lines:
        if DRAFT_NOTE_RE.match(text):
            break
        kept.append((number, text))
    return kept


# ---------------------------------------------------------------------------
# 公共小工具
# ---------------------------------------------------------------------------

def _sentence(text: str, start: int, end: int) -> str:
    """命中位置所在的句子，用于豁免判断。"""
    left = max((text.rfind(ch, 0, start) for ch in SENTENCE_END), default=-1)
    rights = [pos for pos in (text.find(ch, end) for ch in SENTENCE_END) if pos != -1]
    right = min(rights) if rights else len(text)
    return text[left + 1:right]


def _clip(text: str, limit: int = 40) -> str:
    text = text.strip()
    return text if len(text) <= limit else text[:limit] + "……"


def _context(text: str, start: int, end: int, width: int = 12) -> str:
    return text[max(0, start - width):min(len(text), end + width)].strip()


def _issue(rule: Rule, level: str, line: int, excerpt: str, context: str = "",
           message: Optional[str] = None, suggestion: Optional[str] = None) -> Dict[str, Any]:
    data: Dict[str, Any] = {
        "id": rule.id,
        "level": level,
        "category": rule.category,
        "line": line,
        "excerpt": excerpt,
        "message": message or rule.message,
        "suggestion": rule.suggestion if suggestion is None else suggestion,
        "basis": rule.basis,
        "check": "soe",
        "gate": rule.gate,
    }
    if context and context != excerpt:
        data["context"] = context
    return data


def _first_text_line(lines: Lines) -> Tuple[int, str]:
    for number, text in lines:
        if text.strip():
            return number, _clip(text, 30)
    return 1, "（全文）"


CN_DIGITS = {"零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5,
             "六": 6, "七": 7, "八": 8, "九": 9}


def to_int(token: str) -> Optional[int]:
    """把阿拉伯数字或一百以内的汉字数字转成整数。"""
    token = token.strip()
    if token.isdigit():
        return int(token)
    if not token or any(ch not in CN_DIGITS and ch != "十" for ch in token):
        return None
    if "十" in token:
        tens, _, ones = token.partition("十")
        tens_value = CN_DIGITS.get(tens, None) if tens else 1
        ones_value = CN_DIGITS.get(ones, None) if ones else 0
        if tens_value is None or ones_value is None:
            return None
        return tens_value * 10 + ones_value
    if len(token) == 1:
        return CN_DIGITS[token]
    return None


NUM = r"([0-9]+|[零〇一二两三四五六七八九十]+)"


# ---------------------------------------------------------------------------
# 正则规则
# ---------------------------------------------------------------------------

def check_regex(lines: Lines, doc_type: Optional[str], external: bool, ruleset: RuleSet) -> List[Dict[str, Any]]:
    issues: List[Dict[str, Any]] = []
    for rule in ruleset.rules:
        if rule.regex is None or not rule.applies(doc_type):
            continue
        if rule.external_only and not external:
            continue
        level = rule.external_level if (external and rule.external_level) else rule.level
        for number, text in lines:
            for match in rule.regex.finditer(text):
                if match.end() == match.start():
                    continue
                if rule.exempt_match and rule.exempt_match.search(match.group(0)):
                    continue
                if rule.exempt and rule.exempt.search(_sentence(text, match.start(), match.end())):
                    continue
                issues.append(_issue(rule, level, number, _clip(match.group(0)),
                                     _context(text, match.start(), match.end())))
    return issues


# ---------------------------------------------------------------------------
# 函数检查
# ---------------------------------------------------------------------------

YITI_ELEMENTS = (
    ("背景和依据", re.compile(r"背景|依据|缘由|必要性"), "P1"),
    ("方案", re.compile(r"方案"), "P1"),
    ("需决策事项", re.compile(r"(?:需|提请|请求|拟请|请)[^。\n]{0,12}?(?:决策|审议|研究讨论|决定|审定)|决策事项|审议事项"), "P1"),
    ("前置程序完成情况", re.compile(r"前置|合法合规审查|合规审查|法律审核|法律意见|职代会|职工代表大会|可行性研究|尽职调查|资产评估|专家论证"), "P1"),
    ("风险分析", re.compile(r"风险"), "P2"),
)
OPTION_RE = re.compile(r"方案[一二三四五1-5ＡＢA-Ba-b甲乙丙]|备选方案|比选|比较|两个方案|两种方案|多个方案|几个方案")


def check_yiti(lines: Lines, ruleset: RuleSet) -> List[Dict[str, Any]]:
    text = "\n".join(t for _n, t in lines)
    line, title = _first_text_line(lines)
    issues: List[Dict[str, Any]] = []
    rule = ruleset.procedural["SOE-101"]
    for name, regex, level in YITI_ELEMENTS:
        if not regex.search(text):
            issues.append(_issue(rule, level, line, title, message=f"议题材料缺少要件：{name}"))
    option_rule = ruleset.procedural["SOE-102"]
    if "方案" in text and not OPTION_RE.search(text):
        issues.append(_issue(option_rule, option_rule.level, line, title))
    return issues


YIAN_ELEMENTS = (
    ("提请审议事项", re.compile(r"请审议|提请审议|请予审议|提请董事会审议|提请审议事项")),
    ("决策程序履行情况", re.compile(r"前置|党委|专门委员会|委员会审议|决策程序")),
)


def check_yian(lines: Lines, ruleset: RuleSet) -> List[Dict[str, Any]]:
    text = "\n".join(t for _n, t in lines)
    line, title = _first_text_line(lines)
    rule = ruleset.procedural["SOE-103"]
    return [_issue(rule, rule.level, line, title, message=f"董事会议案缺少要件：{name}")
            for name, regex in YIAN_ELEMENTS if not regex.search(text)]


SHOULD_RE = re.compile(r"应(?:出席|到会|到)(?:会议的?)?(?:董事)?\s*" + NUM + r"\s*(?:人|名|位)")
ACTUAL_RE = re.compile(r"实(?:际)?(?:出席|到会|到)(?:会议的?)?(?:董事)?\s*" + NUM + r"\s*(?:人|名|位)")
# 票数写法：“同意7票，反对0票，弃权0票”“同意票7票，反对票0票”“同意：7票；反对：0票；弃权：0票”；弃权可省略
_VOTE_SEP = r"\s*[，,、；;]?\s*"
_VOTE_COLON = r"\s*[：:]?\s*"
VOTE_RE = re.compile(r"同意(?:票)?" + _VOTE_COLON + NUM + r"\s*票" + _VOTE_SEP + r"反对(?:票)?" + _VOTE_COLON + NUM
                     + r"\s*票(?:" + _VOTE_SEP + r"弃权(?:票)?" + _VOTE_COLON + NUM + r"\s*票)?")
# 上市公司公告常见的写法：7票同意，0票反对，0票弃权
VOTE_RE_ALT = re.compile(NUM + r"\s*票\s*同意" + _VOTE_SEP + NUM + r"\s*票\s*反对(?:" + _VOTE_SEP
                         + NUM + r"\s*票\s*弃权)?")
RECUSE_COUNT_RE = re.compile(r"回避(?:表决)?[^。；\n]{0,6}?" + NUM + r"\s*(?:人|名|位)|" + NUM
                             + r"\s*(?:人|名|位)(?:关联)?(?:董事)?回避")
RECUSE_RE = re.compile(r"回避")
PASSED_RE = re.compile(r"审议通过|通过了|获得通过|予以通过|表决通过")
FAILED_RE = re.compile(r"未通过|未获通过|不予通过|否决|未审议通过")
UNANIMOUS_RE = re.compile(r"全票|一致通过|一致同意")


def _recuse_count(block: str) -> Tuple[bool, Optional[int]]:
    if not RECUSE_RE.search(block):
        return False, 0
    match = RECUSE_COUNT_RE.search(block)
    if match:
        token = next((g for g in match.groups() if g), None)
        if token is not None:
            return True, to_int(token)
    # “关联董事张某、李某回避表决”：按顿号数姓名
    named = re.search(r"(?:关联)?董事([^。；，,\n董]{1,20}?)回避", block)
    if named:
        names = [n for n in re.split(r"[、，,和及]", named.group(1)) if n.strip()]
        if names:
            return True, len(names)
    return True, None


def check_resolution(lines: Lines, ruleset: RuleSet) -> List[Dict[str, Any]]:
    rule = ruleset.procedural["SOE-104"]
    issues: List[Dict[str, Any]] = []
    should: Optional[Tuple[int, int, str]] = None
    actual: Optional[Tuple[int, int, str]] = None
    for number, text in lines:
        if should is None:
            m = SHOULD_RE.search(text)
            if m and to_int(m.group(1)) is not None:
                should = (number, to_int(m.group(1)) or 0, m.group(0))
        if actual is None:
            m = ACTUAL_RE.search(text)
            if m and to_int(m.group(1)) is not None:
                actual = (number, to_int(m.group(1)) or 0, m.group(0))
    first_line, title = _first_text_line(lines)
    if should is None or actual is None:
        issues.append(_issue(rule, "P1", first_line, title,
                             message="决议未写明应出席和实际出席董事人数，无法核对表决",
                             suggestion="按会议记录写明应出席、实际出席（含委托出席）人数；记录没有的标【待补】"))
    else:
        n_should, n_actual = should[1], actual[1]
        if n_actual > n_should:
            issues.append(_issue(rule, "P0", actual[0], actual[2],
                                 message=f"实际出席 {n_actual} 人多于应出席 {n_should} 人"))
        elif n_should and n_actual * 2 <= n_should:
            issues.append(_issue(rule, "P0", actual[0], actual[2],
                                 message=f"实际出席 {n_actual} 人未超过应出席 {n_should} 人的半数，会议不能举行",
                                 suggestion="核对出席和委托出席情况；人数确实不足的，不写“会议合法有效”，不写“审议通过”"))
    votes_found = False
    block_start = 0
    flat: List[Tuple[int, str]] = list(lines)
    for index, (number, text) in enumerate(flat):
        matches = sorted(list(VOTE_RE.finditer(text)) + list(VOTE_RE_ALT.finditer(text)),
                         key=lambda item: item.start())
        for match in matches:
            votes_found = True
            g_yes, g_no, g_abstain = match.groups()
            yes, no = to_int(g_yes), to_int(g_no)
            abstain = to_int(g_abstain) if g_abstain is not None else 0
            if yes is None or no is None or abstain is None:
                continue
            block_lines = [t for _n, t in flat[block_start:index]] + [text]
            block = "\n".join(block_lines)
            total = yes + no + abstain
            has_recuse, recused = _recuse_count(block)
            excerpt = match.group(0)
            if actual is not None:
                if has_recuse and recused is None:
                    if total > actual[1]:
                        issues.append(_issue(rule, "P0", number, excerpt,
                                             message=f"票数之和 {total} 超过实际出席人数 {actual[1]}"))
                    else:
                        issues.append(_issue(rule, "P2", number, excerpt,
                                             message="有董事回避表决，但没有写明回避人数，无法核对票数",
                                             suggestion="写明回避的董事（职务或“某董事”）和人数"))
                else:
                    expected = actual[1] - (recused or 0)
                    if total != expected:
                        detail = f"实际出席 {actual[1]} 人" + (f"、回避 {recused} 人" if recused else "")
                        issues.append(_issue(rule, "P0", number, excerpt,
                                             message=f"票数之和 {total} 与应参加表决的人数 {expected} 不符（{detail}）"))
            if UNANIMOUS_RE.search(block) and (no or abstain):
                issues.append(_issue(rule, "P0", number, excerpt,
                                     message="写了“全票”“一致”，但表决结果有反对或弃权票"))
            passed = bool(PASSED_RE.search(block)) and not FAILED_RE.search(block)
            # 回避人数不明时无法确定基数，只报上面的 P2，不判是否通过
            if should is not None and passed and not (has_recuse and recused is None):
                base = should[1] - (recused or 0)
                if base and yes * 2 <= base:
                    issues.append(_issue(rule, "P0", number, excerpt,
                                         message=f"同意 {yes} 票未超过{'无关联关系董事' if has_recuse else '全体董事'} {base} 人的半数，不能写“通过”",
                                         suggestion="无回避的，同意票须超过全体董事的半数；有关联董事回避的，同意票须超过无关联关系董事人数的半数"
                                                    "（《公司法》第一百三十九条、第一百八十五条）；确未通过的写“未通过”；章程规定更高比例的按章程"))
            if has_recuse and recused and actual is not None and passed and actual[1] - recused < 3:
                issues.append(_issue(rule, "P1", number, excerpt,
                                     message=f"回避后出席的无关联关系董事只有 {actual[1] - recused} 人，不足三人",
                                     suggestion="出席董事会会议的无关联关系董事人数不足三人的，该事项应提交股东会审议"
                                                "（《公司法》第一百三十九条、第一百八十五条）；国有独资公司按出资人机构的规定"))
            block_start = index + 1
    if not votes_found:
        issues.append(_issue(rule, "P1", first_line, title,
                             message="决议未写表决结果",
                             suggestion="每项议案写“表决结果：同意X票，反对X票，弃权X票”；记录没有的标【待补】"))
    return issues


TABLE_SEP_RE = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(?:\|\s*:?-{2,}:?\s*)*\|?\s*$")
ALLOWED_STATUS = ("已完成", "按计划推进", "滞后", "未启动", "持续推进")
EMPTY_CELL_RE = re.compile(r"^(?:|无|-|—|－|/|／|待补|暂无|【待补[^】]*】)$")
DATE_RE = re.compile(r"(\d{4})\s*[年\-./]\s*(\d{1,2})\s*(?:[月\-./]\s*(\d{1,2})\s*日?|月)?")
YEAR_END_RE = re.compile(r"(\d{4})\s*年\s*(?:底|末|年底|年末)")
REASON_RE = re.compile(r"原因|因|由于|受|待|尚未|未能|新时限|调整")


def parse_tables(lines: Lines) -> List[Tuple[List[str], List[Tuple[int, List[str]]]]]:
    """解析 Markdown 表格，返回 [(表头, [(行号, 单元格)])]。"""
    tables = []
    flat = list(lines)
    i = 0
    while i < len(flat) - 1:
        number, text = flat[i]
        if "|" in text and TABLE_SEP_RE.match(flat[i + 1][1]):
            header = _cells(text)
            rows: List[Tuple[int, List[str]]] = []
            j = i + 2
            while j < len(flat) and "|" in flat[j][1] and flat[j][1].strip():
                rows.append((flat[j][0], _cells(flat[j][1])))
                j += 1
            tables.append((header, rows))
            i = j
        else:
            i += 1
    return tables


def _cells(text: str) -> List[str]:
    stripped = text.strip()
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|"):
        stripped = stripped[:-1]
    return [c.strip().replace("**", "") for c in re.split(r"(?<!\\)\|", stripped)]


def _find_col(header: List[str], pattern: str) -> Optional[int]:
    regex = re.compile(pattern)
    for idx, name in enumerate(header):
        if regex.search(name):
            return idx
    return None


QUARTER_RE = re.compile(r"(?:(\d{4})\s*年\s*)?(?:第)?([一二三四1-4])\s*季度")
HALF_RE = re.compile(r"(?:(\d{4})\s*年\s*)?(上半年|年中|下半年|年底|年末|全年)")
MONTH_DAY_RE = re.compile(r"(?<!\d)(\d{1,2})\s*月\s*(?:(\d{1,2})\s*日|底|末)?")
QUARTER_END = {1: (3, 31), 2: (6, 30), 3: (9, 30), 4: (12, 31)}


def _month_end(year: int, month: int) -> _dt.date:
    return _dt.date(year, month, calendar.monthrange(year, month)[1])


def parse_deadline(cell: str, default_year: Optional[int] = None) -> Optional[_dt.date]:
    """解析完成时限。没写年份的（如“8月31日”“二季度”“年底”）用 default_year 补齐；不给 default_year 时返回 None。"""
    end = YEAR_END_RE.search(cell)
    if end:
        return _dt.date(int(end.group(1)), 12, 31)
    m = DATE_RE.search(cell)
    if m:
        year, month = int(m.group(1)), int(m.group(2))
        if not 1 <= month <= 12:
            return None
        if m.group(3):
            try:
                return _dt.date(year, month, int(m.group(3)))
            except ValueError:
                return None
        return _month_end(year, month)
    m = QUARTER_RE.search(cell)
    if m:
        year = int(m.group(1)) if m.group(1) else default_year
        quarter = to_int(m.group(2))
        if year is None or quarter not in QUARTER_END:
            return None
        month, day = QUARTER_END[quarter]
        return _dt.date(year, month, day)
    m = HALF_RE.search(cell)
    if m:
        year = int(m.group(1)) if m.group(1) else default_year
        if year is None:
            return None
        return _dt.date(year, 6, 30) if m.group(2) in ("上半年", "年中") else _dt.date(year, 12, 31)
    m = MONTH_DAY_RE.search(cell)
    if m and default_year is not None:
        month = int(m.group(1))
        if not 1 <= month <= 12:
            return None
        if m.group(2):
            try:
                return _dt.date(default_year, month, int(m.group(2)))
            except ValueError:
                return None
        return _month_end(default_year, month)
    return None


def check_ledger(lines: Lines, ruleset: RuleSet, as_of: _dt.date) -> List[Dict[str, Any]]:
    rule = ruleset.procedural["SOE-105"]
    issues: List[Dict[str, Any]] = []
    tables = [t for t in parse_tables(lines) if _find_col(t[0], r"状态") is not None]
    if not tables:
        line, title = _first_text_line(lines)
        return [_issue(rule, "P1", line, title, message="没有找到带“状态”列的 Markdown 表格",
                       suggestion="台账用 Markdown 表格，列见 references/templates/gaige-taizhang.md 第二部分")]
    for header, rows in tables:
        c_status = _find_col(header, r"状态")
        c_deadline = _find_col(header, r"完成时限|时限|完成时间|计划完成")
        c_evidence = _find_col(header, r"佐证|证明材料|成果材料|支撑材料")
        c_owner = _find_col(header, r"牵头|责任单位|责任部门")
        c_remark = _find_col(header, r"备注|原因")
        c_progress = _find_col(header, r"进展|完成情况")
        c_rate = _find_col(header, r"完成率|进度")
        for number, cells in rows:
            def cell(idx: Optional[int]) -> str:
                return cells[idx] if idx is not None and idx < len(cells) else ""
            status = cell(c_status)
            label = _clip(" | ".join(c for c in cells[:2] if c), 30) or status
            if not status:
                issues.append(_issue(rule, "P1", number, label, message="状态为空"))
                continue
            if "待核实" in status:
                continue
            if status not in ALLOWED_STATUS:
                issues.append(_issue(rule, "P1", number, status,
                                     message=f"状态“{status}”不在约定口径内（已完成、按计划推进、滞后、未启动、持续推进）",
                                     suggestion="按认定标准改为五种状态之一；完成比例写进“进展情况”"))
            if status == "已完成" and c_evidence is not None and EMPTY_CELL_RE.match(cell(c_evidence)):
                issues.append(_issue(rule, "P1", number, label, message="标为“已完成”，但没有佐证材料",
                                     suggestion="写明成果文件名称、文号或纪要期号；没有佐证的不标已完成"))
            deadline = parse_deadline(cell(c_deadline), as_of.year) if c_deadline is not None else None
            if deadline and deadline <= as_of and status not in ("已完成", "滞后"):
                passed_word = "已到" if deadline == as_of else "已过"
                issues.append(_issue(rule, "P1", number, f"{cell(c_deadline)}｜{status}",
                                     message=f"完成时限 {deadline.isoformat()} {passed_word}（统计截止 {as_of.isoformat()}），状态仍为“{status}”",
                                     suggestion="到期未完成的改为“滞后”，备注写原因和新的时限；确已完成的写“已完成”并附佐证"))
            if status == "滞后" and not REASON_RE.search(cell(c_remark) + cell(c_progress)):
                issues.append(_issue(rule, "P2", number, label, message="滞后任务没有写原因和新的时限"))
            if c_owner is not None and EMPTY_CELL_RE.match(cell(c_owner)):
                issues.append(_issue(rule, "P2", number, label, message="牵头部门为空"))
            if c_rate is not None:
                rate = re.search(r"(\d+(?:\.\d+)?)\s*[%％]", cell(c_rate))
                if rate and float(rate.group(1)) > 100:
                    issues.append(_issue(rule, "P1", number, cell(c_rate), message="完成率超过 100%"))
    return issues


ESG_ELEMENTS = (
    ("报告期", re.compile(r"报告期|报告时间范围|时间范围|报告周期|\d{4}年1月1日至")),
    ("报告范围", re.compile(r"报告范围|报告边界|组织范围|覆盖范围|本报告涵盖|涵盖范围|合并(?:财务)?报表范围")),
    ("编制依据", re.compile(r"编制依据|参照《|依据《|按照《|参考《|根据《")),
    ("数据来源和口径", re.compile(r"数据来源|数据说明|数据口径|统计口径|计算方法")),
)


def check_esg(lines: Lines, ruleset: RuleSet) -> List[Dict[str, Any]]:
    text = "\n".join(t for _n, t in lines)
    line, title = _first_text_line(lines)
    rule = ruleset.procedural["SOE-106"]
    return [_issue(rule, rule.level, line, title, message=f"缺少披露口径说明：{name}")
            for name, regex in ESG_ELEMENTS if not regex.search(text)]


MONEY_RE = re.compile(r"\d+(?:\.\d+)?\s*(亿元|万元|千元|元)")
# 制度标准金额（如“5000万元以上的投资须提交党委会前置研究”）是门槛，不是本事项的金额，不计入金额单位统计。
# 认定条件：金额后接“以上”“以下”“以内”（含“及以上”“（含）以上”），或金额前有“超过”“达到”“不足”；
# 并且所在句有“须”“应”“需”“规定”“标准”等规定性用语（“应收”“应付”“供应”“需求”等不算）。
STANDARD_AFTER_RE = re.compile(r"\s*(?:（含）|\(含\))?\s*及?(?:以上|以下|以内)")
STANDARD_BEFORE_RE = re.compile(r"(?:超过|达到|不足)\s*$")
NORMATIVE_RE = re.compile(r"须|需(?!求)|(?<![供相对反适响回呼答效])应(?![收付交税])|规定|标准")
SENTENCE_STOPS = "。！？!?"


def _sentence_at(text: str, start: int, end: int) -> str:
    """取 text 中包含 [start, end) 的那一句（以句号、问号、叹号为界，分号不断句）。"""
    left = max(text.rfind(ch, 0, start) for ch in SENTENCE_STOPS) + 1
    rights = [i for i in (text.find(ch, end) for ch in SENTENCE_STOPS) if i != -1]
    return text[left:min(rights) if rights else len(text)]


def is_standard_amount(text: str, m: "re.Match[str]") -> bool:
    """金额是否为制度规定的标准（门槛）金额。"""
    if not (STANDARD_AFTER_RE.match(text, m.end()) or STANDARD_BEFORE_RE.search(text, 0, m.start())):
        return False
    return bool(NORMATIVE_RE.search(_sentence_at(text, m.start(), m.end())))


def check_money_units(lines: Lines, ruleset: RuleSet) -> List[Dict[str, Any]]:
    rule = ruleset.procedural["SOE-107"]
    counts: Dict[str, int] = {}
    first_seen: Dict[str, Tuple[int, str]] = {}
    for number, text in lines:
        for m in MONEY_RE.finditer(text):
            if is_standard_amount(text, m):
                continue
            unit = m.group(1)
            counts[unit] = counts.get(unit, 0) + 1
            first_seen.setdefault(unit, (number, m.group(0)))
    if len(counts) < 2:
        return []
    order = sorted(first_seen.items(), key=lambda kv: kv[1][0])
    number, excerpt = order[1][1]
    detail = "，".join(f"{u} {counts[u]} 处" for u, _ in order)
    return [_issue(rule, rule.level, number, excerpt, message=f"金额单位不统一（{detail}）")]


DECISION_LINE_RE = re.compile(r"会议(?:议定|同意|原则同意|决定|要求|明确)")
SKIP_DECISION_RE = re.compile(r"暂缓|不同意|不予同意|如下[：:]?\s*$")
OWNER_RE = re.compile(r"由[^。]{0,20}?(?:牵头|负责|组织|落实|办理)|牵头|负责|责任(?:部门|单位)|要求[^。，；]{0,12}?(?:部|室|中心|公司|单位)|【待补")
TIME_RE = re.compile(r"\d{1,2}月\d{1,2}日|\d{1,2}月(?:底|末|中旬|上旬|下旬)|年底|年内|月底|季度末|前完成|前报|时限|每月|每季度|【待补")
MAJOR_RE = re.compile(r"三重一大|对外投资|投资建设|增资|产权转让|股权转让|大额资金|重大项目|并购|资产重组|对外担保"
                      r"|提供[^。；\n]{0,8}?担保|为[^。；\n]{0,10}?担保|担保\s*\d|借款\s*\d|融资\s*\d")
PREPOSITION_RE = re.compile(r"党委(?:会)?(?:已|已经)?(?:前置)?研究|前置研究|提交(?:公司)?党委")
ITEM_HEAD_RE = re.compile(r"^\s*(?:[一二三四五六七八九十]+、|（[一二三四五六七八九十]+）)")


def _minutes_items(lines: List[Tuple[int, str]]) -> List[List[Tuple[int, str]]]:
    """按“一、”“（一）”分议定事项；没有序号的，每行一项。第一个序号前的开头段不算议定事项。"""
    heads = [i for i, (_n, text) in enumerate(lines) if ITEM_HEAD_RE.match(text)]
    if not heads:
        return [[item] for item in lines if item[1].strip()]
    bounds = heads + [len(lines)]
    return [lines[bounds[k]:bounds[k + 1]] for k in range(len(heads))]


def check_minutes(lines: Lines, ruleset: RuleSet) -> List[Dict[str, Any]]:
    issues: List[Dict[str, Any]] = []
    rule = ruleset.procedural["SOE-108"]
    for number, text in lines:
        m = DECISION_LINE_RE.search(text)
        if not m or SKIP_DECISION_RE.search(text):
            continue
        missing = []
        if not OWNER_RE.search(text):
            missing.append("责任部门")
        if not TIME_RE.search(text):
            missing.append("时限")
        if missing:
            issues.append(_issue(rule, rule.level, number, _clip(text[m.start():], 30),
                                 message=f"议定事项没有写明{'和'.join(missing)}"))
    major_rule = ruleset.procedural["SOE-109"]
    for item in _minutes_items(list(lines)):
        if PREPOSITION_RE.search("\n".join(t for _n, t in item)):
            continue
        for number, line_text in item:
            m = MAJOR_RE.search(line_text)
            if m:
                issues.append(_issue(major_rule, major_rule.level, number, _clip(m.group(0)),
                                     _context(line_text, m.start(), m.end())))
                break
    return issues


PARTY_RE = re.compile(r"党委|党支部|党总支|党员|党组织|党建|支部|突击队|先锋岗")
RESULT_RE = re.compile(r"\d+(?:\.\d+)?\s*(?:%|％|个百分点|万元|亿元|元|吨|天|小时|项|个|件|台|套|人次|次|千瓦时|万千瓦时|公里|米)")


def check_fusion(lines: Lines, ruleset: RuleSet) -> List[Dict[str, Any]]:
    rule = ruleset.procedural["SOE-110"]
    party_lines = [n for n, t in lines if PARTY_RE.search(t)]
    if not party_lines:
        return []
    joined = any(PARTY_RE.search(t) and RESULT_RE.search(t) for _n, t in lines)
    if joined:
        return []
    line, title = _first_text_line(lines)
    has_result = any(RESULT_RE.search(t) for _n, t in lines)
    message = rule.message if has_result else "全文没有结果数据，党建做法的作用无从体现"
    return [_issue(rule, rule.level, line, title, message=message)]


PROCEDURAL_IDS = ("SOE-101", "SOE-102", "SOE-103", "SOE-104", "SOE-105", "SOE-106",
                  "SOE-107", "SOE-108", "SOE-109", "SOE-110")


def check_procedural(lines: Lines, doc_type: Optional[str], ruleset: RuleSet,
                     as_of: _dt.date) -> List[Dict[str, Any]]:
    issues: List[Dict[str, Any]] = []
    if doc_type == "议题材料":
        issues += check_yiti(lines, ruleset)
    if doc_type == "议案":
        issues += check_yian(lines, ruleset)
    if doc_type == "决议":
        issues += check_resolution(lines, ruleset)
    if doc_type == "改革台账":
        issues += check_ledger(lines, ruleset, as_of)
    if doc_type == "ESG报告":
        issues += check_esg(lines, ruleset)
    if doc_type == "纪要":
        issues += check_minutes(lines, ruleset)
    if doc_type == "融合材料":
        issues += check_fusion(lines, ruleset)
    if ruleset.procedural["SOE-107"].applies(doc_type) and doc_type is not None:
        issues += check_money_units(lines, ruleset)
    return issues


# ---------------------------------------------------------------------------
# 汇总
# ---------------------------------------------------------------------------

def check_lines(lines: Lines, doc_type: Optional[str] = None, external: bool = False,
                ruleset: Optional[RuleSet] = None, as_of: Optional[_dt.date] = None) -> List[Dict[str, Any]]:
    """对正文检查，返回 issue 列表（按行号排序）。"""
    ruleset = ruleset or load_rules()
    as_of = as_of or _dt.date.today()
    external = external or doc_type in ruleset.external_types
    issues = check_regex(lines, doc_type, external, ruleset)
    issues += check_procedural(lines, doc_type, ruleset, as_of)
    issues.sort(key=lambda item: (item["line"], LEVELS.index(item["level"]), item["id"]))
    return issues


def build_result(file_name: str, lines: Lines, doc_type: Optional[str], external: bool,
                 ruleset: Optional[RuleSet] = None, as_of: Optional[_dt.date] = None) -> Dict[str, Any]:
    ruleset = ruleset or load_rules()
    if doc_type is not None and doc_type not in ruleset.types:
        raise InputError(f"--type 只能取：{'、'.join(ruleset.types)}")
    as_of = as_of or _dt.date.today()
    is_external = bool(external or doc_type in ruleset.external_types)
    issues = check_lines(lines, doc_type, is_external, ruleset, as_of)
    notes: List[str] = []
    if doc_type is None:
        notes.append("未指定 --type，只执行了不限文书的正则规则；议题要件、决议票数、台账状态等检查没有执行。")
    if not is_external:
        notes.append("未按对外材料检查；对外发布的材料加 --external，查未公开信息、商业秘密和经营数据。")
    notes.append("行业机检只覆盖能写成规则的项目；前置程序是否属实、数据是否与财务快报一致，按 references/review-addon.md 人工核对。")
    result: Dict[str, Any] = {
        "file": file_name,
        "doc_type": doc_type or "未指定",
        "external": is_external,
        "char_count": sum(len(re.sub(r"\s", "", text)) for _n, text in lines),
        "summary": {lv: sum(1 for i in issues if i["level"] == lv) for lv in LEVELS},
        "issues": issues,
        "notes": notes,
        "engine_version": VERSION,
    }
    if doc_type == "改革台账":
        result["as_of"] = as_of.isoformat()
    return result


def check_text(text: str, doc_type: Optional[str] = None, external: bool = False,
               file_name: str = "<text>", ruleset: Optional[RuleSet] = None,
               as_of: Optional[_dt.date] = None) -> Dict[str, Any]:
    """检查一段文本（供测试和其他脚本调用）。"""
    return build_result(file_name, cut_at_note(split_lines(text)), doc_type, external, ruleset, as_of)


def check_file(path: Path, doc_type: Optional[str] = None, external: bool = False,
               ruleset: Optional[RuleSet] = None, as_of: Optional[_dt.date] = None) -> Dict[str, Any]:
    return build_result(path.name, read_lines(path), doc_type, external, ruleset, as_of)


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------

def render_text(result: Dict[str, Any]) -> str:
    unit = "段" if str(result["file"]).lower().endswith(".docx") else "行"
    s = result["summary"]
    scope = "，按对外材料从严" if result.get("external") else ""
    as_of = f"；统计截止 {result['as_of']}" if result.get("as_of") else ""
    out = [
        f"国企行业机检结果：{result['file']}（文书：{result['doc_type']}{scope}{as_of}；正文 {result['char_count']} 字）",
        f"P0 {s['P0']} 处｜P1 {s['P1']} 处｜P2 {s['P2']} 处",
    ]
    if not result["issues"]:
        out += ["", "未发现行业规则问题。"]
    for issue in result["issues"]:
        out.append("")
        out.append(f"[{issue['level']}] 第{issue['line']}{unit}  {issue['id']}  {issue['category']}：{issue['message']}")
        line = f"    原文：{issue['excerpt']}"
        if issue.get("context"):
            line += f"（上下文：{issue['context']}）"
        out.append(line)
        if issue.get("suggestion"):
            out.append(f"    建议：{issue['suggestion']}")
        out.append(f"    依据：{issue['basis']}")
    if result.get("notes"):
        out.append("")
        out.extend(f"提示：{note}" for note in result["notes"])
    return "\n".join(out)


def render_rules(ruleset: RuleSet) -> str:
    out = ["国企行业机检规则清单："]
    for rule in ruleset.rules + list(ruleset.procedural.values()):
        scope = "、".join(rule.types) if rule.types else "全部"
        ext = f"；对外材料 {rule.external_level}" if rule.external_level else ""
        only = "；只查对外材料" if rule.external_only else ""
        out.append(f"{rule.id}  [{rule.level}{ext}{only}]  {rule.category}：{rule.message}（适用：{scope}）")
    return "\n".join(out)


class _ChineseHelpFormatter(argparse.HelpFormatter):
    def add_usage(self, usage, actions, groups, prefix=None):  # type: ignore[override]
        return super().add_usage(usage, actions, groups, prefix="用法：")


def _parse_date(value: str) -> _dt.date:
    try:
        return _dt.date.fromisoformat(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("日期格式应为 YYYY-MM-DD，如 2026-09-30") from exc


def build_parser(types: Sequence[str]) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="soe_check.py",
        description="国有企业专用文书的行业机检：决策程序、表决票数、议题和议案要件、改革台账状态、"
                    "ESG 披露口径、对外表述、未公开信息、数据写法。",
        epilog="退出码：0 正常；--strict 且存在 P0 时为 1；输入错误为 2。docx 输入时“行号”为段落序号。",
        add_help=False,
        formatter_class=_ChineseHelpFormatter,
    )
    parser._positionals.title = "位置参数"
    parser._optionals.title = "选项"
    parser.add_argument("file", nargs="?", help="待检查文件（.txt、.md、.docx），遇到【起草说明】行即停止读取")
    parser.add_argument("--type", dest="doc_type", default=None, metavar="文书",
                        help="文书类型：" + "、".join(types) + "。不填时只执行不限文书的规则")
    parser.add_argument("--external", action="store_true",
                        help="按对外发布材料从严（未公开信息、商业秘密、经营数据、领先类表述）；"
                             "--type ESG报告、对外材料 时自动启用")
    parser.add_argument("--as-of", dest="as_of", type=_parse_date, default=None, metavar="日期",
                        help="改革台账的统计截止日期（YYYY-MM-DD），默认今天")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    parser.add_argument("--strict", action="store_true", help="存在 P0 时返回 1")
    parser.add_argument("--list-rules", action="store_true", help="列出全部规则后退出")
    parser.add_argument("-h", "--help", action="help", help="显示帮助并退出")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
        except Exception:
            pass
    try:
        ruleset = load_rules()
    except InputError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    parser = build_parser(ruleset.types)
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:
        return 0 if exc.code == 0 else 2
    if args.list_rules:
        print(render_rules(ruleset))
        return 0
    if not args.file:
        print("错误：请指定待检查的文件", file=sys.stderr)
        return 2
    try:
        result = check_file(Path(args.file), args.doc_type, args.external, ruleset, args.as_of)
    except InputError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(render_text(result))
    if args.strict and result["summary"]["P0"] > 0:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
