# -*- coding: utf-8 -*-
"""密级和内部标注识别：style_profile.py、draft_diff.py 共用，不单独运行（gongwen-profile）。

规则来源：
    1. 优先读同目录 lint_rules.json 中 PRI-005（密级标识）、PRI-006（内部资料等标注）、
       PRI-007（未公开、未披露）的 pattern，与共享机检 gongwen_lint.py 保持一致；
    2. 读不到、缺条目或无法编译时，用本文件内置的副本；
    3. 另加本模块的补充规则（"严禁外传""不得翻印""不得复制"等后缀，后缀可带括号；"内部参阅""内部参考"
       "仅供内部参考"；"此件……"；"本报告数据尚未公开披露"一类不带"为、系、属"的说明），始终检查。

判定：
    逐行匹配；命中位置在引号内的视为举例或引述，不算稿件本身的标注（与机检 skip_in_quotes 一致）。
    正文中提到"保密工作""涉密人员管理"之类的词语不算。
    只返回命中的类别，不返回命中内容。

docx：正文（含修订中删除的文字）、页眉、页脚都检查，标注常放在页眉。
只依赖标准库。
"""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Pattern, Sequence, Tuple
from xml.etree import ElementTree as ET

RULES_FILE = Path(__file__).resolve().parent / "lint_rules.json"
RULE_IDS = ("PRI-005", "PRI-006", "PRI-007")
LABELS = {
    "PRI-005": "密级标识",
    "PRI-006": "内部标注",
    "PRI-007": "未公开或未披露信息",
}

# 内置副本：与 shared/scripts/lint_rules.json 同编号规则一致（核对日期 2026-10-16）
BUILTIN: Dict[str, str] = {
    "PRI-005": (r"(?:绝密|机密|秘密)\s*★|^\s*(?:绝密|机密|秘密)\s*$|密级\s*[：:]\s*(?:绝密|机密|秘密)|"
                r"^\s*(?:绝密|机密|秘密)\s*[（(]?\s*\d+\s*年\s*[）)]?\s*$"),
    "PRI-006": (r"^\s*[（(【\[]?\s*(?:内部资料|内部文件|内部材料|内部参阅|工作秘密|内部使用|不得外传|注意保密|"
                r"仅限内部(?:使用|传阅)?|内部传阅)\s*[，,、　\s]*(?:注意保存|妥善保管|不得外传|请勿外传|注意保密)?"
                r"\s*[）)】\]]?\s*$|(?:本|此|该)(?:材料|文件|稿件|稿|资料|报告|件)[^。\n]{0,6}"
                r"(?:内部资料|工作秘密|内部文件|内部材料|内部使用|不得外传|仅供内部)"),
    "PRI-007": (r"(?:上述|以上|本|该|这些|此)[^。；！？\n]{0,12}(?:为|系|属于?|均为|均系)"
                r"(?:未公开|未披露|尚未公开|尚未披露)"),
}

# 补充规则：机检词表之外的常见写法
_HEAD = (r"(?:内部资料|内部文件|内部材料|内部参阅|内部参考|内部传阅|工作秘密|内部使用|不得外传|严禁外传|不得翻印|"
         r"注意保密|仅限内部(?:使用|传阅)?|仅供内部(?:参考|使用|参阅|传阅)?)")
_TAIL = (r"(?:注意保存|妥善保管|注意保管|不得外传|请勿外传|严禁外传|不得翻印|严禁翻印|不得复制|严禁复制|"
         r"不得转发|注意保密)")
_OPEN, _CLOSE = r"[（(【\[]", r"[）)】\]]"
EXTRA: Tuple[Tuple[str, Pattern[str]], ...] = (
    # 整行只有标注，如"内部资料，严禁外传""（内部资料 不得翻印）""内部资料（注意保存）""仅供内部参考"
    ("内部标注", re.compile(r"^\s*" + _OPEN + r"?\s*" + _HEAD + r"\s*" + _CLOSE + r"?"
                           + r"(?:\s*[，,、；;　]?\s*" + _OPEN + r"?\s*" + _TAIL + r"\s*" + _CLOSE + r"?)*"
                           r"\s*[。.]?\s*$")),
    # "本材料属工作秘密""此件为内部参阅"
    ("内部标注", re.compile(r"(?:本|此|该)(?:材料|文件|稿件|稿|资料|报告|件)[^。\n]{0,6}"
                           r"(?:内部资料|工作秘密|内部文件|内部材料|内部使用|内部参阅|不得外传|严禁外传|"
                           r"不得翻印|仅供内部)")),
    # "本报告数据尚未公开披露""以上指标未对外公开"：主语须是材料、数据一类的词，避免"该企业未披露……"误报
    ("未公开或未披露信息", re.compile(r"(?:上述|以上|本|该|这些|此)[^。；！？\n]{0,4}"
                                    r"(?:报告|材料|文件|数据|指标|信息|内容|稿件|资料|情况)[^。；！？\n]{0,6}"
                                    r"(?:尚)?未(?:对外)?(?:公开|披露)")),
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS
DOCX_PART_RE = re.compile(r"^word/(?:document|header\d*|footer\d*)\.xml$")

_cache: Optional[List[Tuple[str, Pattern[str]]]] = None


def load_rules(rules_file: Optional[Path] = None) -> List[Tuple[str, Pattern[str]]]:
    """返回 [(类别, 正则)]：PRI-005/006/007（优先取 lint_rules.json）加补充规则。"""
    patterns = dict(BUILTIN)
    path = rules_file or RULES_FILE
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        for item in data.get("rules", []):
            if not isinstance(item, dict):
                continue
            rid, pattern = item.get("id"), item.get("pattern")
            if rid in RULE_IDS and isinstance(pattern, str):
                try:
                    re.compile(pattern)
                except re.error:
                    continue  # 无法编译的保留内置副本
                patterns[rid] = pattern
    except (OSError, ValueError, AttributeError):
        pass  # 读不到或格式不对：用内置副本
    rules = [(LABELS[rid], re.compile(patterns[rid])) for rid in RULE_IDS]
    rules.extend(EXTRA)
    return rules


def _rules() -> List[Tuple[str, Pattern[str]]]:
    global _cache
    if _cache is None:
        _cache = load_rules()
    return _cache


def _inside_quote(text: str, pos: int) -> bool:
    before = text[:pos]
    return before.count("“") > before.count("”") or before.count("\"") % 2 == 1


def find_marks(lines: Sequence[str]) -> List[str]:
    """逐行检查，返回命中的类别（去重，按规则顺序）；没有命中返回空列表。"""
    found: List[str] = []
    for label, regex in _rules():
        if label in found:
            continue
        for line in lines:
            hit = False
            for match in regex.finditer(line):
                if match.start() == match.end():
                    continue
                if _inside_quote(line, match.start()):
                    continue
                hit = True
                break
            if hit:
                found.append(label)
                break
    return found


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def docx_mark_lines(path: Path) -> List[str]:
    """docx 中用于查标注的文字，每段一行：正文（含修订中删除的文字）、页眉、页脚。

    读取失败返回空列表，由调用方的正常读取流程报错。
    """
    lines: List[str] = []
    try:
        with zipfile.ZipFile(str(path)) as zf:
            parts = sorted(n for n in zf.namelist() if DOCX_PART_RE.match(n))
            for name in parts:
                root = ET.fromstring(zf.read(name))
                for p in root.iter(W + "p"):
                    pieces = []
                    for node in p.iter():
                        tag = _local(node.tag)
                        if tag in ("t", "delText") and node.text:
                            pieces.append(node.text)
                        elif tag == "tab":
                            pieces.append("\t")
                    lines.append("".join(pieces))
    except (zipfile.BadZipFile, KeyError, OSError, ET.ParseError):
        return []
    return lines


def describe(labels: Sequence[str]) -> str:
    """命中类别的说明文字，用于提示用户。"""
    return "、".join(labels) if labels else ""
