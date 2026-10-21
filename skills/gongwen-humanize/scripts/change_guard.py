# -*- coding: utf-8 -*-
"""改动范围守卫（change_guard.py）：改稿相对原稿到底动了什么，每处改动有没有依据。

源文件位于仓库 shared/scripts/，各 skill 中的副本由 tools/sync_shared.py 生成，请在源文件处修改。

用法：
    python scripts/change_guard.py <原稿> <改稿> [--keep-structure] [--doc-type 文书类型] [--json] [--strict]

做法：把原稿和改稿按句对齐（只比对【起草说明】之前的正文，.docx 按接受全部修订后的文字），逐处列出改动，并分三类：
    CHG-001 事实变化（P0）：数字、日期、人名、单位名称、引号内内容、书名号和文件名、发文字号、【待补】标记，
            在这处改动里丢失或新增（以整篇的出现次数为准，句子之间搬动、合并不算丢失）；
            单位名称只是全称与简称互换（一个写法包含另一个）的降为 P1；
    CHG-002 有规则依据的改动（info，不计入问题）：原句命中 gongwen_lint.py 的规则或文风词表，记录规则号和级别；
    CHG-003 无规则依据的改动（P2）：既不涉及事实、原句也没有命中规则，列出供核对是否越界改写。
--keep-structure（用户要求结构不动时使用）：另查结构，下列变化报 CHG-004（P1）：
    文章标题、层次序号（一、（一）1.（1）的序列）、层次标题文字、段落数、段落先后顺序。

对齐方法：先用 difflib 对句子序列做最长匹配，匹配不上的按相似度（不低于 0.35）配对成“改写”，
跨段落搬动并改写的句子再按相似度（不低于 0.6）配对；仍配不上的是“删除”或“新增”。

退出码：0 正常；--strict 且存在 P0 时为 1；输入错误（文件不存在、格式不支持等）为 2。
只依赖标准库；与 gongwen_lint.py、fact_diff.py、lint_rules.json、style_patterns.json 放在同一目录。
"""
from __future__ import annotations

import argparse
import difflib
import json
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import fact_diff as fd  # noqa: E402
import gongwen_lint as gl  # noqa: E402

VERSION = "1.0.0"
LEVELS = ("P0", "P1", "P2")
PAIR_MIN_BLOCK = 0.35     # 同一处差异块内，两句配成“改写”的最低相似度
PAIR_MIN_GLOBAL = 0.6     # 跨块（搬动并改写）配对的最低相似度
PARA_MATCH_MIN = 0.5      # 段落顺序检查：两段视为同一段的最低相似度

# 规则登记（lint_rules.json 的 change_guard 段缺失时的兜底）
DEFAULT_META = {
    "CHG-001": {"level": "P0", "message": "事实变化：数字、日期、人名、单位、引号内容、文件名等在改动中丢失或新增",
                "basis": "references/_shared/guardrails.md 不编造、不改事实；fact_diff.py"},
    "CHG-002": {"level": "info", "message": "有规则依据的改动：原句命中机检规则",
                "basis": "gongwen_lint.py、style_patterns.json 对应规则的依据"},
    "CHG-003": {"level": "P2", "message": "无规则依据的改动：不涉及事实，原句也没有命中规则，请核对是否越界改写",
                "basis": "references/_shared/guardrails.md 最小改动原则"},
    "CHG-004": {"level": "P1", "message": "结构变化：用户要求结构不动时，标题、层次序号、段落数或顺序发生了变化",
                "basis": "references/_shared/guardrails.md 最小改动原则；用户的结构不动要求"},
}


def load_meta() -> Dict[str, Dict[str, str]]:
    meta = {k: dict(v) for k, v in DEFAULT_META.items()}
    try:
        data = json.loads((SCRIPT_DIR / "lint_rules.json").read_text(encoding="utf-8"))
        for item in data.get("change_guard", []):
            if isinstance(item, dict) and item.get("id") in meta:
                meta[item["id"]].update({k: item[k] for k in ("level", "message", "basis") if k in item})
    except (OSError, ValueError):
        pass
    return meta


META = load_meta()


class InputError(Exception):
    """输入错误，退出码 2。"""


# ---------------------------------------------------------------------------
# 切句
# ---------------------------------------------------------------------------

def split_sentences(text: str) -> List[Tuple[int, int]]:
    """按 。！？； 切句，返回 [(起, 止)]；引号内的句末标点不切，句末的右引号、右括号并入本句。"""
    spans: List[Tuple[int, int]] = []
    depth, begin = 0, 0
    for pos, ch in enumerate(text):
        if ch in "“‘「『":
            depth += 1
        elif ch in "”’」』" and depth:
            depth -= 1
        elif ch in "。！？；" and depth == 0:
            end = pos + 1
            while end < len(text) and text[end] in "”’」』）)":
                end += 1
            spans.append((begin, end))
            begin = end
    if begin < len(text):
        spans.append((begin, len(text)))
    return [(s, e) for s, e in spans if text[s:e].strip()]


class Unit:
    """一个句子（层次标题等无句末标点的短行也各算一句）。"""

    __slots__ = ("para", "line", "start", "end", "text", "norm")

    def __init__(self, para: int, line: int, start: int, end: int, text: str) -> None:
        self.para = para
        self.line = line
        self.start = start
        self.end = end
        self.text = text.strip()
        self.norm = re.sub(r"\s+", "", text)


class Doc:
    def __init__(self, name: str, text: str) -> None:
        self.name = name
        self.lines = fd.body_lines(text)
        self.paras: List[Tuple[int, str]] = []     # [(行号, 去掉行尾空白的文字)]
        self.units: List[Unit] = []
        for idx, line in enumerate(self.lines, start=1):
            if not line.strip():
                continue
            text_line = line.rstrip()
            para_no = len(self.paras)
            self.paras.append((idx, text_line))
            for s, e in split_sentences(text_line):
                self.units.append(Unit(para_no, idx, s, e, text_line[s:e]))


def read_doc(path: Path) -> Doc:
    try:
        text, _ = fd.read_document(path)
    except fd.InputError as exc:
        raise InputError(str(exc)) from exc
    return Doc(path.name, text)


# ---------------------------------------------------------------------------
# 事实抽取（fact_diff 的数字、日期、引号等 + 单位、人名、文件名）
# ---------------------------------------------------------------------------

ORG_RE = re.compile(
    r"[一-鿿]{2,14}?(?:人民政府|办公室|办公厅|委员会|党委|党组|[一-鿿]{1,6}局|医院|公司|集团|大学|学院|研究院|研究所|检察院|法院|"
    r"派出所|卫生院|幼儿园|中学|小学|办事处|协会|管委会|卫健委|教委|团委|工会|妇联|政协|人大)")
ORG_LEAD_STOP = "请由向对将把经与和及同在从至为让令使的了并各等其该本我贵报送呈抄交发转按据依联合会同召开组织协调印发"
ORG_NOT_ORG = ("布局", "格局", "大局", "结局", "政局", "战局", "时局", "棋局", "当局", "全局", "开局", "破局", "残局", "危局", "骗局",
               "困局", "邮局", "书局", "药局", "商局", "饭局", "赌局", "牌局")
PERSON_TITLE_RE = re.compile(
    "([" + "".join(sorted(gl.SURNAMES)) + r"][一-鿿某]{1,2})(?=同志|书记|局长|主任|校长|院长|经理|主席|老师|医生|教授|处长|科长|部长|县长|市长|区长|镇长)")
PERSON_LABEL_RE = re.compile(
    r"(?:联系人|负责人|组长|副组长|主持人|汇报人|牵头人|经办人|报告人|主讲人|发言人|签发人|审核人|拟稿人)\s*[：:]?\s*([一-鿿]{2,4})")
FILE_RE = re.compile(r"[^\s，。；：、“”《》（）()\[\]]{1,40}\.(?:docx?|xlsx?|pptx?|pdf|txt|md|wps|csv|zip|rar|jpe?g|png)\b")


def strip_lead(token: str) -> str:
    while len(token) > 3 and token[0] in ORG_LEAD_STOP:
        token = token[1:]
    return token


def org_tokens(text: str) -> List[str]:
    tokens: List[str] = []
    for m in ORG_RE.finditer(text):
        token = strip_lead(m.group(0))
        if token.endswith(ORG_NOT_ORG):
            continue
        tokens.append(token)
    return tokens


def entity_facts(text: str) -> List[Tuple[str, str]]:
    result: List[Tuple[str, str]] = []
    result.extend(("单位", t) for t in org_tokens(text))
    result.extend(("人名", m.group(1)) for m in PERSON_TITLE_RE.finditer(text))
    result.extend(("人名", m.group(1)) for m in PERSON_LABEL_RE.finditer(text))
    result.extend(("文件名", strip_lead(m.group(0))) for m in FILE_RE.finditer(text))
    return result


def unit_facts(unit: Unit) -> List[Tuple[Tuple[str, str], str]]:
    """返回 [((类型, 比对值), 显示值)]；fact_diff 中带 hint（序数、概数）的不比对。"""
    facts: List[Tuple[Tuple[str, str], str]] = []
    for f in fd.extract_line(unit.text, unit.line):
        if f.hint:
            continue
        facts.append(((f.kind, f.norm), f.value))
    for kind, value in entity_facts(unit.text):
        facts.append(((kind, value), value))
    return facts


# ---------------------------------------------------------------------------
# 对齐
# ---------------------------------------------------------------------------

def similarity(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    if sm.real_quick_ratio() < 0.2 or sm.quick_ratio() < 0.2:
        return 0.0
    return sm.ratio()


def greedy_pairs(left: Sequence[int], right: Sequence[int], a_units: Sequence[Unit], b_units: Sequence[Unit],
                 threshold: float) -> Tuple[List[Tuple[int, int]], List[int], List[int]]:
    scored: List[Tuple[float, int, int]] = []
    for i in left:
        for j in right:
            r = similarity(a_units[i].norm, b_units[j].norm)
            if r >= threshold:
                scored.append((r, i, j))
    scored.sort(key=lambda x: (-x[0], x[1], x[2]))
    used_i, used_j = set(), set()
    pairs: List[Tuple[int, int]] = []
    for _, i, j in scored:
        if i in used_i or j in used_j:
            continue
        used_i.add(i)
        used_j.add(j)
        pairs.append((i, j))
    pairs.sort()
    return pairs, [i for i in left if i not in used_i], [j for j in right if j not in used_j]


def align(a: Doc, b: Doc) -> Tuple[List[Tuple[Optional[int], Optional[int]]], int]:
    """返回（改动列表 [(原稿句序号或 None, 改稿句序号或 None)]，搬动未改的句数）。"""
    sm = difflib.SequenceMatcher(None, [u.norm for u in a.units], [u.norm for u in b.units], autojunk=False)
    pairs: List[Tuple[Optional[int], Optional[int]]] = []
    left_over_a: List[int] = []
    left_over_b: List[int] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        got, rest_a, rest_b = greedy_pairs(range(i1, i2), range(j1, j2), a.units, b.units, PAIR_MIN_BLOCK)
        pairs.extend(got)
        left_over_a.extend(rest_a)
        left_over_b.extend(rest_b)
    # 搬动：整句相同只是换了位置的不算改动；搬动并改写的按相似度配对
    moved = 0
    by_norm: Dict[str, List[int]] = {}
    for j in left_over_b:
        by_norm.setdefault(b.units[j].norm, []).append(j)
    still_a: List[int] = []
    for i in left_over_a:
        bucket = by_norm.get(a.units[i].norm)
        if bucket:
            j = bucket.pop(0)
            left_over_b.remove(j)
            moved += 1
        else:
            still_a.append(i)
    got, rest_a, rest_b = greedy_pairs(still_a, left_over_b, a.units, b.units, PAIR_MIN_GLOBAL)
    pairs.extend(got)
    result: List[Tuple[Optional[int], Optional[int]]] = list(pairs)
    result.extend((i, None) for i in rest_a)
    result.extend((None, j) for j in rest_b)
    result.sort(key=lambda p: (p[0] if p[0] is not None else 10 ** 9, p[1] if p[1] is not None else 10 ** 9))
    return result, moved


# ---------------------------------------------------------------------------
# 规则依据
# ---------------------------------------------------------------------------

def lint_original(doc: Doc, doc_type: Optional[str]) -> Tuple[Dict[Tuple[int, int], List[Dict[str, str]]], str]:
    """用机检引擎检查原稿，把命中映射到句子：{(行号, 句首): [{id, level, message}]}。"""
    lines = [(i, line, False) for i, line in enumerate(doc.lines, start=1)]
    gdoc = gl.Document(doc.name, lines)
    dtype, _ = gl.resolve_doc_type(gdoc, doc_type)
    linter = gl.Linter()
    issues = linter.lint(gdoc, dtype, set(gl.CHECK_ORDER))
    by_line: Dict[int, List[Unit]] = {}
    for u in doc.units:
        by_line.setdefault(u.line, []).append(u)
    mapped: Dict[Tuple[int, int], List[Dict[str, str]]] = {}
    for issue in issues:
        if issue.start < 0:
            continue
        for u in by_line.get(issue.line, []):
            if u.start <= issue.start < u.end or (issue.start < u.start < issue.end):
                entry = {"id": issue.id, "level": issue.level, "message": issue.message}
                bucket = mapped.setdefault((u.line, u.start), [])
                if entry not in bucket:
                    bucket.append(entry)
                break
    return mapped, dtype


# ---------------------------------------------------------------------------
# 结构
# ---------------------------------------------------------------------------

def heading_info(doc: Doc) -> List[Tuple[Tuple[int, int], str, int]]:
    result = []
    for idx, (line, text) in enumerate(doc.paras):
        h = gl.heading_of(text)
        if h:
            result.append((h, re.sub(r"\s+", "", text), line))
    return result


def lis_indices(seq: Sequence[int]) -> set:
    """最长递增子序列所在的位置集合（用于找出被搬动的段落）。"""
    n = len(seq)
    if n == 0:
        return set()
    best = [1] * n
    prev = [-1] * n
    for i in range(n):
        for j in range(i):
            if seq[j] < seq[i] and best[j] + 1 > best[i]:
                best[i] = best[j] + 1
                prev[i] = j
    end = max(range(n), key=lambda k: best[k])
    keep = set()
    while end != -1:
        keep.add(end)
        end = prev[end]
    return keep


def structure_issues(a: Doc, b: Doc) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []

    def add(message: str, original: str = "", revised: str = "", line_a: int = 0, line_b: int = 0) -> None:
        out.append({"id": "CHG-004", "level": META["CHG-004"]["level"], "message": message,
                    "original": {"line": line_a, "text": original}, "revised": {"line": line_b, "text": revised}})

    ga = gl.Document(a.name, [(i, l, False) for i, l in enumerate(a.lines, start=1)])
    gb = gl.Document(b.name, [(i, l, False) for i, l in enumerate(b.lines, start=1)])
    ta, tb = re.sub(r"\s+", "", ga.title_text), re.sub(r"\s+", "", gb.title_text)
    if ta != tb:
        add("文章标题变化", ga.title_text, gb.title_text,
            ga.paras[ga.title[0]].line if ga.title else 0, gb.paras[gb.title[0]].line if gb.title else 0)
    ha, hb = heading_info(a), heading_info(b)
    seq_a = [h[0] for h in ha]
    seq_b = [h[0] for h in hb]
    if seq_a != seq_b:
        add(f"层次序号变化（原稿 {len(seq_a)} 个层次标题，改稿 {len(seq_b)} 个，或序号不同）",
            "、".join(f"{lv}级{n}" for lv, n in seq_a[:8]), "、".join(f"{lv}级{n}" for lv, n in seq_b[:8]))
    else:
        for x, y in zip(ha, hb):
            if x[1] != y[1]:
                add("层次标题文字变化", x[1], y[1], x[2], y[2])
    if len(a.paras) != len(b.paras):
        add(f"段落数变化（原稿 {len(a.paras)} 段，改稿 {len(b.paras)} 段）")
    # 段落顺序：每个原稿段落找最相似的改稿段落，按最长递增子序列找出被搬动的
    norms_b = [re.sub(r"\s+", "", t) for _, t in b.paras]
    used: set = set()
    matched: List[Tuple[int, int]] = []
    for i, (line, text) in enumerate(a.paras):
        na = re.sub(r"\s+", "", text)
        best_j, best_r = -1, 0.0
        for j, nb in enumerate(norms_b):
            if j in used:
                continue
            r = 1.0 if na == nb else similarity(na, nb)
            if r > best_r:
                best_j, best_r = j, r
            if r == 1.0:
                break
        if best_j >= 0 and best_r >= PARA_MATCH_MIN:
            used.add(best_j)
            matched.append((i, best_j))
    keep = lis_indices([j for _, j in matched])
    for pos, (i, j) in enumerate(matched):
        if pos not in keep:
            add(f"段落顺序变化（原稿第 {i + 1} 段在改稿中移到第 {j + 1} 段）", a.paras[i][1][:40], b.paras[j][1][:40],
                a.paras[i][0], b.paras[j][0])
    return out


# ---------------------------------------------------------------------------
# 汇总
# ---------------------------------------------------------------------------

def unit_ref(unit: Optional[Unit]) -> Dict[str, Any]:
    if unit is None:
        return {"line": 0, "text": ""}
    return {"line": unit.line, "text": unit.text}


def build_report(original: Doc, revised: Doc, keep_structure: bool = False, doc_type: Optional[str] = None) -> Dict[str, Any]:
    pairs, moved = align(original, revised)
    rule_map, dtype = lint_original(original, doc_type)

    facts_a = [unit_facts(u) for u in original.units]
    facts_b = [unit_facts(u) for u in revised.units]
    count_a = Counter(k for fs in facts_a for k, _ in fs)
    count_b = Counter(k for fs in facts_b for k, _ in fs)
    missing = count_a - count_b
    added = count_b - count_a

    changes: List[Dict[str, Any]] = []
    for i, j in pairs:
        ua = original.units[i] if i is not None else None
        ub = revised.units[j] if j is not None else None
        kind = "modify" if ua and ub else ("delete" if ua else "add")
        lost: List[Dict[str, str]] = []
        gained: List[Dict[str, str]] = []
        if i is not None:
            for key, shown in facts_a[i]:
                if missing[key] > 0:
                    missing[key] -= 1
                    lost.append({"type": key[0], "value": shown})
        if j is not None:
            for key, shown in facts_b[j]:
                if added[key] > 0:
                    added[key] -= 1
                    gained.append({"type": key[0], "value": shown})
        rules = rule_map.get((ua.line, ua.start), []) if ua else []
        item: Dict[str, Any] = {"kind": kind, "original": unit_ref(ua), "revised": unit_ref(ub)}
        if lost or gained:
            only_abbrev = (bool(lost) and bool(gained) and all(x["type"] == "单位" for x in lost + gained)
                           and all(any(g["value"] in l["value"] or l["value"] in g["value"] for g in gained)
                                   for l in lost))
            item.update(id="CHG-001", category="fact", level="P1" if only_abbrev else "P0",
                        message=META["CHG-001"]["message"] if not only_abbrev else "单位名称的全称与简称互换，请核对是否同一单位",
                        facts={"lost": lost, "added": gained}, rules=rules)
        elif rules:
            item.update(id="CHG-002", category="ruled", level="info", message=META["CHG-002"]["message"], rules=rules)
        else:
            msg = {"modify": "改写了这句话，没有事实变化，原句也没有命中规则",
                   "delete": "删除了原句，没有事实丢失，原句也没有命中规则",
                   "add": "新增了这句话，没有新增事实"}[kind]
            item.update(id="CHG-003", category="unruled", level=META["CHG-003"]["level"], message=msg, rules=[])
        changes.append(item)
    for n, item in enumerate(changes, start=1):
        item["no"] = n

    structure: List[Dict[str, Any]] = structure_issues(original, revised) if keep_structure else []
    levels = [c["level"] for c in changes if c["level"] in LEVELS] + [s["level"] for s in structure]
    summary = {lv: levels.count(lv) for lv in LEVELS}
    summary["ruled"] = sum(1 for c in changes if c["category"] == "ruled")
    summary["changes"] = len(changes)
    summary["moved"] = moved
    summary["unchanged"] = len(original.units) - sum(1 for c in changes if c["original"]["line"]) - moved
    notes = []
    if not keep_structure:
        notes.append("未加 --keep-structure，没有检查标题、序号、段落数和顺序。")
    return {
        "original_file": original.name, "revised_file": revised.name, "doc_type": dtype,
        "keep_structure": keep_structure, "summary": summary, "changes": changes, "structure": structure,
        "notes": notes, "engine_version": VERSION,
    }


def render_text(report: Dict[str, Any]) -> str:
    s = report["summary"]
    out = [
        f"改动范围守卫：{report['original_file']} → {report['revised_file']}（文书类型：{report['doc_type']}）",
        f"P0 {s['P0']} 处｜P1 {s['P1']} 处｜P2 {s['P2']} 处｜有规则依据 {s['ruled']} 处｜共 {s['changes']} 处改动，"
        f"{s['unchanged']} 句未动，{s['moved']} 句仅搬动",
    ]
    order = {"P0": 0, "P1": 1, "P2": 2, "info": 3}
    for c in sorted(report["changes"], key=lambda x: (order.get(x["level"], 9), x["no"])):
        tag = c["level"] if c["level"] in LEVELS else "依据"
        out.append("")
        out.append(f"[{tag}] #{c['no']} {c['id']}  {c['message']}")
        if c["original"]["text"]:
            out.append(f"    原稿第{c['original']['line']}行：{c['original']['text']}")
        if c["revised"]["text"]:
            out.append(f"    改稿第{c['revised']['line']}行：{c['revised']['text']}")
        facts = c.get("facts") or {}
        for key, label in (("lost", "丢失"), ("added", "新增")):
            if facts.get(key):
                out.append(f"    {label}：" + "；".join(f"{x['type']}“{x['value']}”" for x in facts[key]))
        if c.get("rules"):
            out.append("    命中规则：" + "；".join(f"{r['id']}（{r['level']}）{r['message']}" for r in c["rules"]))
    for st in report["structure"]:
        out.append("")
        out.append(f"[{st['level']}] {st['id']}  {st['message']}")
        if st["original"]["text"]:
            out.append(f"    原稿：{st['original']['text']}")
        if st["revised"]["text"]:
            out.append(f"    改稿：{st['revised']['text']}")
    if not report["changes"] and not report["structure"]:
        out.append("")
        out.append("改稿与原稿的句子完全一致（或只有搬动）。")
    for n in report["notes"]:
        out.append("")
        out.append(f"提示：{n}")
    return "\n".join(out)


class _ChineseHelpFormatter(argparse.RawDescriptionHelpFormatter):
    def add_usage(self, usage, actions, groups, prefix=None):  # type: ignore[override]
        return super().add_usage(usage, actions, groups, prefix="用法：")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="change_guard.py", add_help=False, formatter_class=_ChineseHelpFormatter,
        description="改动范围守卫：句子级对齐原稿与改稿，列出每处改动并分为事实变化、有规则依据、无规则依据三类。",
        epilog="退出码：0 正常；--strict 且存在 P0 时为 1；输入错误为 2。")
    parser._positionals.title = "位置参数"
    parser._optionals.title = "选项"
    parser.add_argument("original", help="原稿（.txt、.md、.docx）")
    parser.add_argument("revised", help="改稿（.txt、.md、.docx）")
    parser.add_argument("--keep-structure", action="store_true", help="用户要求结构不动：标题、序号、段落数和顺序变化报 P1")
    parser.add_argument("--doc-type", default=None, help="原稿的文书类型（不填时按标题推断）")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    parser.add_argument("--strict", action="store_true", help="存在 P0 时返回 1")
    parser.add_argument("-h", "--help", action="help", help="显示帮助并退出")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
        except Exception:
            pass
    args = build_parser().parse_args(argv)
    try:
        original = read_doc(Path(args.original))
        revised = read_doc(Path(args.revised))
        report = build_report(original, revised, args.keep_structure, args.doc_type)
    except InputError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(render_text(report))
    if args.strict and report["summary"]["P0"] > 0:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
