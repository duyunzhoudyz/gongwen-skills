#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""医院专用文书的行业机检（gongwen-hospital）。

在通用机检 gongwen_lint.py 之外，补查医院行业中能写成规则的问题：
  - 疗效保证、治愈率百分之百、最高级和领先表述（HOS-001～003，对外材料升级）；
  - 决策程序：党政联席会议代替党委会议（HOS-004），议题涉及“三重一大”却没有党委会议程序（HOS-005）；
  - 指数写法：CMI 等指数和天数用“百分点”（HOS-006），CMI 带百分号（HOS-007）；
  - 指标名称不规范（HOS-008～015）；
  - 引用已废止的依据（HOS-016、017）；
  - 根因分析把原因落在个人品质、把直接原因当根本原因、写入处理处罚（HOS-018～020）。

用法：
    python scripts/hospital_check.py <文件> [--type 文书] [--external] [--json] [--strict]
    python scripts/hospital_check.py --list-rules

输入 .txt、.md、.docx（docx 需要 python-docx），遇到单独成行的【起草说明】即停止读取。
输出字段与 gongwen_lint.py 一致（file、doc_type、char_count、summary、issues），summary 同样含 markers
（【待补】【待核实】计数，开发规范第 8 节），可以并入核稿意见。
退出码：0 正常；--strict 且存在 P0 时为 1；输入错误为 2。
只依赖标准库；规则在同目录的 hospital_rules.json 中，修改规则不必改代码。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Pattern, Sequence, Tuple

VERSION = "1.0.0"
RULES_PATH = Path(__file__).resolve().parent / "hospital_rules.json"
LEVELS = ("P0", "P1", "P2")
EXTERNAL_TYPE = "对外宣传"

# 【起草说明】分界行，与 gongwen_lint.py、fact_diff.py 的写法一致
DRAFT_NOTE_RE = re.compile(r"^\s*(?:#{1,6}\s*)?(?:(?:\*\*|__)\s*)?【起草说明】")
SENTENCE_END = "。！？；!?;\n"
# 程序性规则的否定判断：只看 satisfied 命中处之前同一分句内的若干字
CLAUSE_END = SENTENCE_END + "，,"
NEGATION_WINDOW = 8
# summary.markers 计数，与 gongwen_lint.py 的 TODO_MARK_RE、VERIFY_MARK_RE 写法一致（全角冒号）
TODO_MARK_RE = re.compile(r"【待补：[^】\n]*】")
VERIFY_MARK_RE = re.compile(r"【待核实：[^】\n]*】")


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
    exempt: Optional[Pattern[str]] = None
    trigger: Optional[Pattern[str]] = None
    satisfied: Optional[Pattern[str]] = None
    negation: Optional[Pattern[str]] = None
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
    rules: List[Rule]
    procedural: List[Rule]


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
        exempt=_compile(item.get("exempt_context"), rid, "exempt_context"),
        hit_examples=list(item.get("hit_examples") or []),
        miss_examples=list(item.get("miss_examples") or []),
    )
    if procedural:
        rule.trigger = _compile(item.get("trigger"), rid, "trigger")
        rule.satisfied = _compile(item.get("satisfied"), rid, "satisfied")
        rule.negation = _compile(item.get("satisfied_negation"), rid, "satisfied_negation")
        if rule.trigger is None or rule.satisfied is None:
            raise InputError(f"程序性规则 {rid} 须同时有 trigger 和 satisfied")
    else:
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
    rules = [_build_rule(item, False) for item in data.get("rules") or []]
    procedural = [_build_rule(item, True) for item in data.get("procedural") or []]
    ids = [r.id for r in rules + procedural]
    duplicated = sorted({i for i in ids if ids.count(i) > 1})
    if duplicated:
        raise InputError(f"规则 id 重复：{'、'.join(duplicated)}")
    for rule in rules + procedural:
        unknown = [t for t in rule.types if t not in types]
        if unknown:
            raise InputError(f"规则 {rule.id} 的 types 含未登记的文书类型：{'、'.join(unknown)}")
    return RuleSet(types=types, rules=rules, procedural=procedural)


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
        lines = list(enumerate(text.replace("\r\n", "\n").replace("\r", "\n").split("\n"), start=1))
    elif suffix == ".docx":
        lines = _read_docx(path)
    else:
        raise InputError(f"不支持的文件类型：{path.suffix or '无扩展名'}。只支持 .txt、.md、.docx。")
    return cut_at_note(lines)


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
            for p_el in child.iter(p_tag):
                number += 1
                lines.append((number, Paragraph(p_el, document).text))
    return lines


def cut_at_note(lines: Sequence[Tuple[int, str]]) -> List[Tuple[int, str]]:
    """只保留【起草说明】之前的正文。"""
    kept: List[Tuple[int, str]] = []
    for number, text in lines:
        if DRAFT_NOTE_RE.match(text):
            break
        kept.append((number, text))
    return kept


# ---------------------------------------------------------------------------
# 检查
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


def _issue(rule: Rule, level: str, line: int, excerpt: str, context: str = "") -> Dict[str, Any]:
    data: Dict[str, Any] = {
        "id": rule.id,
        "level": level,
        "category": rule.category,
        "line": line,
        "excerpt": excerpt,
        "message": rule.message,
        "suggestion": rule.suggestion,
        "basis": rule.basis,
        "check": "hospital",
        "gate": rule.gate,
    }
    if context and context != excerpt:
        data["context"] = context
    return data


def _is_satisfied(rule: Rule, text: str) -> bool:
    """程序性规则：全文有一处未被否定的 satisfied 命中即算满足。

    satisfied_negation 只看命中处之前、同一分句内的 NEGATION_WINDOW 个字，
    如“尚未提交党委会议”“未经党委会议”不算已有党委会议程序。
    """
    if rule.satisfied is None:
        return False
    for match in rule.satisfied.finditer(text):
        if rule.negation is None:
            return True
        before = text[max(0, match.start() - NEGATION_WINDOW):match.start()]
        cut = max(before.rfind(ch) for ch in CLAUSE_END)
        if not rule.negation.search(before[cut + 1:]):
            return True
    return False


def check_lines(lines: Sequence[Tuple[int, str]], doc_type: Optional[str] = None,
                external: bool = False, ruleset: Optional[RuleSet] = None) -> List[Dict[str, Any]]:
    """对正文逐行检查，返回 issue 列表（按行号排序）。"""
    ruleset = ruleset or load_rules()
    external = external or doc_type == EXTERNAL_TYPE
    issues: List[Dict[str, Any]] = []
    for rule in ruleset.rules:
        if not rule.applies(doc_type) or rule.regex is None:
            continue
        level = rule.external_level if (external and rule.external_level) else rule.level
        for number, text in lines:
            for match in rule.regex.finditer(text):
                if match.end() == match.start():
                    continue
                if rule.exempt and rule.exempt.search(_sentence(text, match.start(), match.end())):
                    continue
                issues.append(_issue(rule, level, number, _clip(match.group(0)),
                                     _context(text, match.start(), match.end())))
    full_text = "\n".join(text for _number, text in lines)
    for rule in ruleset.procedural:
        if not rule.applies(doc_type) or rule.trigger is None or rule.satisfied is None:
            continue
        if _is_satisfied(rule, full_text):
            continue
        for number, text in lines:
            match = rule.trigger.search(text)
            if match:
                issues.append(_issue(rule, rule.level, number, match.group(0),
                                     _context(text, match.start(), match.end())))
                break
    issues.sort(key=lambda item: (item["line"], LEVELS.index(item["level"]), item["id"]))
    return issues


def check_text(text: str, doc_type: Optional[str] = None, external: bool = False,
               file_name: str = "<text>", ruleset: Optional[RuleSet] = None) -> Dict[str, Any]:
    """检查一段文本（供测试和其他脚本调用）。"""
    lines = list(enumerate(text.replace("\r\n", "\n").replace("\r", "\n").split("\n"), start=1))
    return build_result(file_name, cut_at_note(lines), doc_type, external, ruleset)


def check_file(path: Path, doc_type: Optional[str] = None, external: bool = False,
               ruleset: Optional[RuleSet] = None) -> Dict[str, Any]:
    return build_result(path.name, read_lines(path), doc_type, external, ruleset)


def build_result(file_name: str, lines: Sequence[Tuple[int, str]], doc_type: Optional[str],
                 external: bool, ruleset: Optional[RuleSet] = None) -> Dict[str, Any]:
    ruleset = ruleset or load_rules()
    if doc_type is not None and doc_type not in ruleset.types:
        raise InputError(f"--type 只能取：{'、'.join(ruleset.types)}")
    issues = check_lines(lines, doc_type, external, ruleset)
    notes: List[str] = []
    if doc_type is None:
        notes.append("未指定 --type，只限某类文书的规则（如根因分析写法、议题程序）没有执行。")
    notes.append("行业机检只覆盖能写成规则的项目；数据溯源、口径、前置程序是否属实，按 references/review-addon.md 人工核对。")
    summary: Dict[str, Any] = {lv: sum(1 for i in issues if i["level"] == lv) for lv in LEVELS}
    summary["markers"] = {
        "待补": sum(len(TODO_MARK_RE.findall(text)) for _n, text in lines),
        "待核实": sum(len(VERIFY_MARK_RE.findall(text)) for _n, text in lines),
    }
    return {
        "file": file_name,
        "doc_type": doc_type or "未指定",
        "external": bool(external or doc_type == EXTERNAL_TYPE),
        "char_count": sum(len(re.sub(r"\s", "", text)) for _n, text in lines),
        "summary": summary,
        "issues": issues,
        "notes": notes,
        "engine_version": VERSION,
    }


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------

def render_text(result: Dict[str, Any]) -> str:
    unit = "段" if str(result["file"]).lower().endswith(".docx") else "行"
    s = result["summary"]
    scope = "，按对外材料从严" if result.get("external") else ""
    out = [
        f"行业机检结果：{result['file']}（文书：{result['doc_type']}{scope}；正文 {result['char_count']} 字）",
        f"P0 {s['P0']} 处｜P1 {s['P1']} 处｜P2 {s['P2']} 处",
    ]
    markers = s.get("markers") or {}
    if markers.get("待补") or markers.get("待核实"):
        out.append(f"待补 {markers.get('待补', 0)} 处、待核实 {markers.get('待核实', 0)} 处，定稿前补齐")
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
    out = ["行业机检规则清单："]
    for rule in ruleset.rules + ruleset.procedural:
        scope = "、".join(rule.types) if rule.types else "全部"
        ext = f"；对外材料 {rule.external_level}" if rule.external_level else ""
        out.append(f"{rule.id}  [{rule.level}{ext}]  {rule.category}：{rule.message}（适用：{scope}）")
    return "\n".join(out)


class _ChineseHelpFormatter(argparse.HelpFormatter):
    def add_usage(self, usage, actions, groups, prefix=None):  # type: ignore[override]
        return super().add_usage(usage, actions, groups, prefix="用法：")


def build_parser(types: Sequence[str]) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="hospital_check.py",
        description="医院专用文书的行业机检：疗效表述、决策程序、指数写法、指标名称、过时依据、根因分析写法。",
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
                        help="按对外发布材料从严（疗效、领先类表述升级）；--type 对外宣传 时自动启用")
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
        result = check_file(Path(args.file), args.doc_type, args.external, ruleset)
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
