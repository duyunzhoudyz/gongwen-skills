#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""校验仓库内所有 skill 是否符合 docs/开发规范.md。

检查项：frontmatter、SKILL.md 行数、manifest.json 与共享资料同步、SKILL.md 中的相对路径、
共享资料引用闭合（SKILL.md 和 references/ 下各 .md 中出现的 references/_shared/ 路径都要存在）、
evals、README.md、.claude-plugin/marketplace.json，以及全仓库的字体文件、大文件和 JSON 合法性。

可选检查 --eval-overlap（默认不运行，结果记为警告）：评测材料（evals/files/、evals.json 的 prompt、
expected_output、expectations，以及 trigger.json 的 query）中带单位的数字和完整日期，是否出现在
本 skill 的 SKILL.md、README.md、references/（不含 _shared）中；全仓库检查时另查 shared/ 源文件
和 examples/，提示示例可能复用了评测材料。

输出"错误 / 警告"清单；存在错误时返回 1，参数错误返回 2。
"""
from __future__ import annotations

import fnmatch
import os
import re
import sys
import zipfile
from pathlib import Path, PurePosixPath
from typing import Dict, List, Optional, Sequence, Set, Tuple

# 允许以 python tools/validate.py 以外的方式加载时也能找到同目录模块
_TOOLS_DIR = str(Path(__file__).resolve().parent)
if _TOOLS_DIR not in sys.path:
    sys.path.insert(0, _TOOLS_DIR)

from sync_shared import (  # noqa: E402
    REPO_ROOT,
    add_root_argument,
    build_plan,
    diff_plan,
    exists_exact,
    list_skill_dirs,
    load_json_file,
    make_parser,
    normalize_rel_path,
    setup_console,
)

NAME_RE = re.compile(r"[a-z0-9]+(?:-[a-z0-9]+)*", re.ASCII)
SEMVER_RE = re.compile(
    r"(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)"
    r"(?:-((?:0|[1-9][0-9]*|[0-9]*[a-zA-Z-][0-9a-zA-Z-]*)"
    r"(?:\.(?:0|[1-9][0-9]*|[0-9]*[a-zA-Z-][0-9a-zA-Z-]*))*))?"
    r"(?:\+([0-9a-zA-Z-]+(?:\.[0-9a-zA-Z-]+)*))?",
    re.ASCII,
)
ALLOWED_FIELDS = ("name", "description", "license")
RESERVED_WORDS = ("claude", "anthropic")
NAME_PREFIX = "gongwen-"
ENTRY_NAME = "gongwen"  # 总入口不带后缀
REQUIRED_SHARED = (
    "guardrails.md",
    "style/output-format.md",
    "style/ai-tone.md",
    "style/defensive.md",
)
VALID_STATUS = ("stable", "beta", "draft")
FONT_SUFFIXES = {".ttf", ".otf", ".ttc", ".woff", ".woff2"}
MAX_FILE_BYTES = 1024 * 1024
# 超过 1MB 也不报警告的文件：
# 1. evals/results/ 下的 *-raw.json（如 v1.0-raw.json）——评测原始逐条结果，aggregate.py 靠它复算通过率。
# 2. evals/results/ 各版本子目录下的 results.json（如 compete-v6/results.json）——版本归档结果。
# 3. .claude/worktrees/ 下任意深度的同类文件——worktree 副本，路径不同但内容同上。
# 体量随评测用例数和 skill 数增长，属于有意保留。放行的文件仍检查 JSON 是否合法。
LARGE_FILE_EXEMPT_GLOBS = (
    "evals/results/*-raw.json",
    "evals/results/*/*-raw.json",
    "evals/results/*/results.json",
    ".claude/worktrees/*/evals/results/*-raw.json",
    ".claude/worktrees/*/evals/results/*/*-raw.json",
)
SKILL_MD_WARN_LINES = 300
SKILL_MD_MAX_LINES = 500
DESCRIPTION_WARN_CHARS = 300
DESCRIPTION_MAX_CHARS = 1024
NAME_MAX_CHARS = 64
MIN_EVALS = 12
TRIGGER_COUNT = 20

# 规范第 6 节统一词表
UNIT_TYPES = ("党政机关", "基层", "社区", "公立医院", "高校", "中小学", "国有企业")
ROLE_MODES = ("教学模式", "高效模式", "转译模式", "审稿模式", "严谨模式")

# SKILL.md 中的相对路径：前面不能紧跟 ASCII 路径字符（避免匹配 shared/scripts/... 的后半段）
REL_PATH_RE = re.compile(
    r"(?<![A-Za-z0-9_./\\-])"
    r"((?:references|scripts|assets|evals)/[^\s`'\"()（）\[\]，。；：、！？|“”‘’]*)"
)
PLACEHOLDER_MARKS = ("<", ">", "*", "{", "}", "…", "...", "＜", "＞", "$")
EXT_CUT_RE = re.compile(r"\.[A-Za-z0-9]{1,8}(?=[^A-Za-z0-9]|$)")
# 共享资料引用闭合检查（D18）：同步进来的共享资料默认放在 references/_shared/ 下
SHARED_PREFIX = "references/_shared/"
MAX_SHOWN_PLACES = 3

# --eval-overlap：评测材料中的"带单位的数字"（整数部分不少于 3 位，或带小数）和完整日期
EVAL_UNITS = (
    "个百分点", "百分点", "万人次", "人次", "例次", "台次", "户次", "万元", "亿元", "万人", "万户", "万亩",
    "平方公里", "平方千米", "平方米", "立方米", "公里", "千米", "公顷", "千克", "公斤", "小时", "分钟", "工作日",
    "%", "‰", "元", "人", "户", "家", "个", "项", "次", "台", "件", "条", "所", "名", "位", "天", "吨", "亩",
    "米", "床", "例", "套", "间", "篇", "份", "场", "批", "辆", "处", "座", "张", "本", "册", "节", "期", "轮",
    "万", "亿", "倍", "岁", "层", "栋", "块", "箱", "分",
)
# 前面紧挨数字、小数点、英文字母（编号、标准号）或"第"（条款号、序号）的不抽取
EVAL_NUMBER_RE = re.compile(
    r"(?<![\d.A-Za-z第])(\d{3,}(?:\.\d+)?|\d+\.\d+)(" + "|".join(re.escape(u) for u in EVAL_UNITS) + r")"
)
EVAL_DATE_RES = (
    re.compile(r"(?<!\d)(\d{4})年(\d{1,2})月(\d{1,2})[日号]"),
    re.compile(r"(?<![\d.])(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?![\d.])"),
)
# ---- --eval-overlap 白名单：这些写法在评测材料和示例中同时出现是合法的，不算复用 ----
# 白名单只放"固定不变、不是编出来的"内容。放行前先问一句：换一个数字会不会让示例失真？
# 会失真的（标准施行日期、法规发布日期、外部来源日期）才放行；示例里随手写的数字和日期不放行。
# 法规规定的截止日期（如年报披露期限）与示例日期恰好相同时，先改示例日期，再把法规日期加入下面第 4 类。
#
# 1) 通用写法。
EVAL_ALLOWED = {"100%"}
# 2) 年初、季末、年末是通用的截止日期和统计时点（1月1日、3月31日、6月30日、9月30日、12月31日）。
EVAL_ALLOWED_MONTH_DAYS = {(1, 1), (3, 31), (6, 30), (9, 30), (12, 31)}
# 3) 评测统一使用的"当前日期"。全部评测用例都设定为这一天，写作日历、示例、格式规则里的"今天"
#    也只能是这一天，换掉会让评测的时间线对不上。
EVAL_CURRENT_DATE = "2026年10月16日"
# 4) 法规、标准中的固定日期。这些日期由文件本身决定，示例引用时不能改，所以放行；
#    取值与 shared/rules/ 各要点文件首部记录的发布、施行日期一致。键为统一写法（年月日不补零），值为出处，
#    仅供维护者查阅。
EVAL_ALLOWED_LEGAL_DATES = {
    "2012年4月16日": "《党政机关公文处理工作条例》印发日期",
    "2012年7月1日": "《党政机关公文处理工作条例》施行日期；GB/T 9704-2012 实施日期",
    "2012年6月29日": "GB/T 9704-2012 发布日期",
    "2011年7月29日": "GB/T 15835-2011 发布日期",
    "2011年11月1日": "GB/T 15835-2011 实施日期",
    "2011年12月30日": "GB/T 15834-2011 发布日期",
    "2012年6月1日": "GB/T 15834-2011 实施日期",
    "2024年7月30日": "《整治形式主义为基层减负若干规定》审议批准日期",
    "2024年8月6日": "《整治形式主义为基层减负若干规定》发布日期",
    "2026年4月30日": "沪深交易所可持续发展报告指引规定的 2025 年度报告披露截止日（与年报同时披露）；"
                     "gongwen-soe 的 quick-cards.md 登记，评测期望要求据此判断",
}
# 5) 规则依据所引外部文章的发布日期。skill 的 references 登记了来源日期，评测期望要求"依据栏注明来源和日期"，
#    二者必然同时出现；日期是外部文章的固定信息，不能为避开重合而改。新增条目须同样满足这两点。
EVAL_ALLOWED_SOURCE_DATES = {
    "2021年4月17日": "共产党员网转载《秘书工作》文章的日期（gongwen-review 的 tifa-risk.md 登记，评测期望要求注明）",
}
# 6) 标准号、文号、条款号不在白名单里，而是由抽取规则结构性排除：数字前紧挨英文字母（GB/T 9704、T/CAS 123）
#    或"第"（第305条）的不抽取；"号"不在单位表内（国办发〔2019〕34号 不会被抽到）；
#    带"-年份"的标准号后缀（9704-2012）后面没有单位，也不会被抽到。见 EVAL_NUMBER_RE。
EVAL_ALLOWED_DATES = {EVAL_CURRENT_DATE, *EVAL_ALLOWED_LEGAL_DATES, *EVAL_ALLOWED_SOURCE_DATES}
EVAL_TEXT_SUFFIXES = {".txt", ".md", ".csv", ".tsv", ".json"}
EVAL_ZIP_PARTS = {".docx": ("word/document.xml",), ".xlsx": ("xl/sharedStrings.xml", "xl/worksheets/")}
FULLWIDTH_DIGITS = str.maketrans("０１２３４５６７８９％．", "0123456789%.")
XML_TAG_RE = re.compile(r"<[^>]+>")
MAX_EVAL_ZIP_PART_BYTES = 20 * 1024 * 1024

# 全仓库扫描时跳过的目录
SKIP_DIR_NAMES = {
    ".git", "__pycache__", ".venv", "venv", "node_modules", ".tox",
    ".mypy_cache", ".pytest_cache", ".gongwen", "公文",
}
SKIP_REL_DIRS = {"dist/offline"}
# 只在 tests/ 下跳过的目录（与 .gitignore 的 tests/**/_out/ 一致，其他位置的 _out 会入库）
SKIP_TEST_DIR_NAMES = {"_out"}

REPO_SCOPE = "仓库"


# ---------------------------------------------------------------------------
# 报告
# ---------------------------------------------------------------------------

class Report:
    def __init__(self) -> None:
        self.items: List[Tuple[str, str, str]] = []  # (范围, 级别, 信息)
        self.scopes: List[str] = []

    def _add(self, scope: str, level: str, message: str) -> None:
        if scope not in self.scopes:
            self.scopes.append(scope)
        self.items.append((scope, level, message))

    def touch(self, scope: str) -> None:
        if scope not in self.scopes:
            self.scopes.append(scope)

    def error(self, scope: str, message: str) -> None:
        self._add(scope, "错误", message)

    def warn(self, scope: str, message: str) -> None:
        self._add(scope, "警告", message)

    @property
    def errors(self) -> List[Tuple[str, str]]:
        return [(s, m) for s, lv, m in self.items if lv == "错误"]

    @property
    def warnings(self) -> List[Tuple[str, str]]:
        return [(s, m) for s, lv, m in self.items if lv == "警告"]

    def render(self, quiet: bool = False) -> str:
        lines: List[str] = []
        for scope in self.scopes:
            entries = [(lv, m) for s, lv, m in self.items if s == scope]
            if quiet:
                entries = [(lv, m) for lv, m in entries if lv == "错误"]
                if not entries:
                    continue
            lines.append(f"== {scope} ==")
            if not entries:
                lines.append("  通过")
            for level, message in entries:
                lines.append(f"  [{level}] {message}")
        return "\n".join(lines)


# ---------------------------------------------------------------------------
# frontmatter 简易解析（只支持单行 key: value，不依赖 PyYAML）
# ---------------------------------------------------------------------------

def _unquote(value: str) -> Tuple[str, Optional[str]]:
    """去掉成对引号，返回 (值, 错误信息)。"""
    if value[:1] in ("'", '"'):
        quote = value[0]
        if len(value) < 2 or value[-1] != quote:
            return value, "引号未闭合（只支持单行值）"
        inner = value[1:-1]
        if quote == '"':
            inner = inner.replace('\\"', '"').replace("\\\\", "\\")
        else:
            inner = inner.replace("''", "'")
        return inner, None
    return value, None


def split_frontmatter(text: str) -> Tuple[Dict[str, str], str, List[str], bool]:
    """拆分 frontmatter 与正文。

    返回 (字段, 正文, 错误列表, 是否存在 frontmatter)。
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if text.startswith("﻿"):
        text = text[1:]
    lines = text.split("\n")
    if not lines or lines[0].rstrip() != "---":
        return {}, text, ["缺少 frontmatter（文件第一行必须是 ---）"], False
    end = None
    for index in range(1, len(lines)):
        if lines[index].rstrip() == "---":
            end = index
            break
    if end is None:
        return {}, text, ["frontmatter 未闭合（缺少结束的 ---）"], False

    fields: Dict[str, str] = {}
    errors: List[str] = []
    for lineno in range(2, end + 1):
        line = lines[lineno - 1]
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        if line[0] in (" ", "\t"):
            errors.append(f"frontmatter 第 {lineno} 行：不支持多行值或缩进（每个字段必须写在一行内）")
            continue
        match = re.match(r"^([A-Za-z0-9_-]+):(?:[ \t]+(.*))?$", line)
        if not match:
            errors.append(f"frontmatter 第 {lineno} 行无法解析（只支持单行 key: value）：{line}")
            continue
        key, raw = match.group(1), (match.group(2) or "").strip()
        if key in fields:
            errors.append(f"frontmatter 字段 {key} 重复")
            continue
        if raw in ("|", ">", "|-", ">-", "|+", ">+"):
            errors.append(f"frontmatter 字段 {key} 使用了多行写法，必须写在一行内")
            fields[key] = ""
            continue
        value, err = _unquote(raw)
        if err:
            errors.append(f"frontmatter 字段 {key}：{err}")
        elif raw[:1] not in ("'", '"'):
            # 未加引号的值在 YAML 中有特殊含义的写法，会导致 Agent 工具解析失败
            if ": " in raw or raw.endswith(":"):
                errors.append(
                    f"frontmatter 字段 {key} 含半角冒号加空格，YAML 会解析出错，"
                    f"请改用全角冒号或给整个值加双引号"
                )
            if " #" in raw:
                errors.append(f"frontmatter 字段 {key} 含\" #\"，YAML 会把其后内容当作注释")
            if raw[:1] in ("[", "{", "&", "*", "!", "%", "@", "`"):
                errors.append(f"frontmatter 字段 {key} 以 YAML 特殊字符 {raw[0]} 开头，请给整个值加双引号")
        fields[key] = value
    body = "\n".join(lines[end + 1:])
    return fields, body, errors, True


# ---------------------------------------------------------------------------
# 单个 skill 的检查
# ---------------------------------------------------------------------------

def check_frontmatter(name: str, fields: Dict[str, str], rep: Report) -> None:
    value = fields.get("name")
    if value is None:
        rep.error(name, "frontmatter 缺少 name")
    else:
        if not NAME_RE.fullmatch(value):
            rep.error(name, f"name \"{value}\" 不合规：只能用小写字母、数字和单个连字符，且不能以连字符开头或结尾")
        if len(value) > NAME_MAX_CHARS:
            rep.error(name, f"name 超过 {NAME_MAX_CHARS} 字符（当前 {len(value)}）")
        if value != ENTRY_NAME and not value.startswith(NAME_PREFIX):
            rep.error(name, f"name \"{value}\" 须以 {NAME_PREFIX} 开头（总入口为 {ENTRY_NAME}）")
        if value != name:
            rep.error(name, f"name \"{value}\" 与目录名 \"{name}\" 不一致")
        for word in RESERVED_WORDS:
            if word in value.lower():
                rep.error(name, f"name 不能包含保留词 {word}")

    desc = fields.get("description")
    if desc is None or not desc.strip():
        rep.error(name, "frontmatter 缺少 description 或为空")
    else:
        if len(desc) > DESCRIPTION_MAX_CHARS:
            rep.error(name, f"description 超过 {DESCRIPTION_MAX_CHARS} 字符（当前 {len(desc)}）")
        elif len(desc) > DESCRIPTION_WARN_CHARS:
            rep.warn(name, f"description 超过 {DESCRIPTION_WARN_CHARS} 字（当前 {len(desc)}），建议精简")

    lic = fields.get("license")
    if lic is None or not lic.strip():
        rep.error(name, "frontmatter 缺少 license")
    elif lic not in ("MIT", "PolyForm-Noncommercial-1.0.0"):
        rep.warn(name, f"license 为 \"{lic}\"，规范要求为 MIT 或 PolyForm-Noncommercial-1.0.0")

    for key in fields:
        if key not in ALLOWED_FIELDS:
            rep.warn(name, f"frontmatter 出现规范外字段 {key}（只允许 name、description、license）")


def _path_ok(skill_dir: Path, rel: str) -> bool:
    if rel.endswith("/"):
        return exists_exact(skill_dir, rel.rstrip("/"), want_dir=True)
    return exists_exact(skill_dir, rel)


def find_missing_paths(skill_dir: Path, text: str, cache: Optional[Dict[str, bool]] = None) -> List[str]:
    """找出文本中引用的 references/、scripts/、assets/、evals/ 路径里不存在的（去重，按出现顺序）。"""
    cache = {} if cache is None else cache

    def ok(rel: str) -> bool:
        if rel not in cache:
            cache[rel] = _path_ok(skill_dir, rel)
        return cache[rel]

    missing: List[str] = []
    for match in REL_PATH_RE.finditer(text):
        raw = match.group(1)
        if any(mark in raw for mark in PLACEHOLDER_MARKS):
            continue  # 示例占位符，如 references/<文种>.md、scripts/*.py
        candidate = raw.rstrip(".,;:!?")
        if not candidate or ok(candidate):
            continue
        # 中文正文紧贴路径时（如 "guardrails.md第一条"），尝试在扩展名处截断
        cuts = [candidate[:m.end()] for m in EXT_CUT_RE.finditer(candidate)]
        if any(ok(cut) for cut in cuts):
            continue
        # 路径后紧跟中文时，只报告到扩展名为止的部分
        shown = candidate if (not cuts or cuts[-1] == candidate) else cuts[0]
        if shown not in missing:
            missing.append(shown)
    return missing


def check_relative_paths(name: str, skill_dir: Path, text: str, rep: Report) -> None:
    """检查 SKILL.md 中引用的 references/、scripts/、assets/、evals/ 路径是否存在。

    references/_shared/ 下的路径由 check_shared_references 统一检查，这里不重复报告。
    """
    for shown in find_missing_paths(skill_dir, text):
        if not shown.startswith(SHARED_PREFIX):
            rep.error(name, f"SKILL.md 引用的路径不存在（注意大小写）：{shown}")


def _join_places(places: List[str], unit: str) -> str:
    text = "、".join(places[:MAX_SHOWN_PLACES])
    if len(places) > MAX_SHOWN_PLACES:
        text += f" 等 {len(places)} {unit}"
    return text


def check_shared_references(repo_root: Path, name: str, skill_dir: Path, rep: Report) -> None:
    """共享资料引用闭合（D18）：SKILL.md 和 references/ 下各 .md（含 _shared）中出现的
    references/_shared/<路径> 都必须存在；缺失时提示在 manifest.json 的 shared 中补上。evals/ 不查。"""
    files: List[Path] = []
    if (skill_dir / "SKILL.md").is_file():
        files.append(skill_dir / "SKILL.md")
    refs_dir = skill_dir / "references"
    if refs_dir.is_dir():
        files += sorted(p for p in refs_dir.rglob("*.md") if p.is_file())
    cache: Dict[str, bool] = {}
    referrers: Dict[str, List[str]] = {}
    for path in files:
        rel = path.relative_to(skill_dir).as_posix()
        try:
            text = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            if rel != "SKILL.md":  # SKILL.md 的编码问题由 check_skill_md 报告
                rep.warn(name, f"{rel} 不是合法的 UTF-8 编码，已跳过共享资料引用检查")
            continue
        except OSError:
            continue
        for shown in find_missing_paths(skill_dir, text, cache):
            if shown.startswith(SHARED_PREFIX):
                referrers.setdefault(shown, []).append(rel)
    for target, sources in referrers.items():
        src = target[len(SHARED_PREFIX):]
        if target.endswith("/"):
            hint = f"请在 manifest.json 的 shared 中补上 {src} 下所需的文件"
        elif exists_exact(repo_root / "shared", src, want_dir=False):
            hint = f"请在 manifest.json 的 shared 中补上 {src}"
        else:
            hint = f"shared/{src} 也不存在，请改正引用或先在 shared/ 中补上该文件"
        rep.error(name, f"共享资料引用不闭合：{_join_places(sources, '个文件')} 引用 {target}，{hint}")


def check_skill_md(name: str, skill_dir: Path, rep: Report) -> None:
    path = skill_dir / "SKILL.md"
    if not exists_exact(skill_dir, "SKILL.md", want_dir=False):
        rep.error(name, "缺少 SKILL.md")
        return
    try:
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        rep.error(name, "SKILL.md 不是合法的 UTF-8 编码")
        return
    fields, _body, errors, has_frontmatter = split_frontmatter(text)
    for err in errors:
        rep.error(name, err)
    if has_frontmatter:
        check_frontmatter(name, fields, rep)

    line_count = len(text.splitlines())
    if line_count > SKILL_MD_MAX_LINES:
        rep.error(name, f"SKILL.md 共 {line_count} 行，超过上限 {SKILL_MD_MAX_LINES} 行，请把细节移到 references/")
    elif line_count > SKILL_MD_WARN_LINES:
        rep.warn(name, f"SKILL.md 共 {line_count} 行，超过建议值 {SKILL_MD_WARN_LINES} 行")

    check_relative_paths(name, skill_dir, text, rep)


def check_manifest(repo_root: Path, name: str, skill_dir: Path, rep: Report) -> Optional[dict]:
    """检查 manifest.json，返回解析后的数据（失败时返回 None）。"""
    path = skill_dir / "manifest.json"
    if not path.is_file():
        rep.error(name, "缺少 manifest.json")
        return None
    data, err = load_json_file(path)
    if err:
        rep.error(name, f"manifest.json：{err}")
        return None
    if not isinstance(data, dict):
        rep.error(name, "manifest.json：顶层必须是 JSON 对象")
        return None

    if data.get("name") != name:
        rep.error(name, f"manifest.json 的 name \"{data.get('name')}\" 与目录名不一致")
    status = data.get("status")
    if status not in VALID_STATUS:
        rep.error(name, f"manifest.json 的 status \"{status}\" 无效，只能是 stable、beta、draft")
    version = data.get("version")
    if not isinstance(version, str) or not SEMVER_RE.fullmatch(version):
        rep.error(name, f"manifest.json 的 version \"{version}\" 不是语义化版本号（如 0.1.0）")

    # 共享资料：src 存在、同步结果逐字节一致
    plan = build_plan(repo_root, skill_dir, data)
    for err in plan.errors:
        rep.error(name, err)
    declared = set()
    shared = data.get("shared", [])
    if isinstance(shared, list):
        for item in shared:
            if isinstance(item, dict):
                src, _ = normalize_rel_path(item.get("src"))
                if src:
                    declared.add(src)
    for required in REQUIRED_SHARED:
        if required not in declared:
            rep.error(name, f"manifest.json 的 shared 必须包含 {required}")
    if not plan.errors:
        diffs = diff_plan(plan)
        for diff in diffs:
            rep.error(name, f"共享资料未同步：{diff.describe()}")
        if diffs:
            rep.error(name, "请运行 python tools/sync_shared.py 后重新提交")

    # 提示词版打包清单
    bundle = data.get("prompt_bundle")
    if bundle is None:
        rep.warn(name, "manifest.json 未设置 prompt_bundle，提示词版将不带附录")
    elif not isinstance(bundle, list):
        rep.error(name, "manifest.json 的 prompt_bundle 必须是数组")
    else:
        for index, item in enumerate(bundle, start=1):
            rel, err = normalize_rel_path(item)
            if err:
                rep.error(name, f"prompt_bundle 第 {index} 项：{err}")
            elif not exists_exact(skill_dir, rel, want_dir=False):
                rep.error(name, f"prompt_bundle 引用的文件不存在（注意大小写）：{rel}")
    return data


def _is_str_list(value: object) -> bool:
    return isinstance(value, list) and all(isinstance(x, str) for x in value)


def check_evals(name: str, skill_dir: Path, status: Optional[str], rep: Report, checked: Set[Path]) -> None:
    is_draft = status == "draft"
    path = skill_dir / "evals" / "evals.json"
    if not path.is_file():
        if is_draft:
            rep.warn(name, "缺少 evals/evals.json")
        else:
            rep.error(name, "缺少 evals/evals.json")
    else:
        checked.add(path.resolve())
        data, err = load_json_file(path)
        if err:
            rep.error(name, f"evals/evals.json：{err}")
        elif not isinstance(data, dict):
            rep.error(name, "evals/evals.json：顶层必须是对象（含 skill_name 和 evals）")
        else:
            _check_evals_data(name, skill_dir, data, is_draft, rep)

    trigger = skill_dir / "evals" / "trigger.json"
    if not trigger.is_file():
        rep.warn(name, "缺少 evals/trigger.json")
        return
    checked.add(trigger.resolve())
    data, err = load_json_file(trigger)
    if err:
        rep.error(name, f"evals/trigger.json：{err}")
        return
    if not isinstance(data, list):
        rep.error(name, "evals/trigger.json：顶层必须是数组")
        return
    positives = negatives = 0
    queries: Set[str] = set()
    for index, item in enumerate(data, start=1):
        where = f"evals/trigger.json 第 {index} 项"
        if not isinstance(item, dict):
            rep.error(name, f"{where}：必须是对象")
            continue
        query = item.get("query")
        if not isinstance(query, str) or not query.strip():
            rep.error(name, f"{where}：query 必须是非空字符串")
        elif query in queries:
            rep.warn(name, f"{where}：query 重复：{query}")
        else:
            queries.add(query)
        flag = item.get("should_trigger")
        if not isinstance(flag, bool):
            rep.error(name, f"{where}：should_trigger 必须是 true 或 false")
        elif flag:
            positives += 1
        else:
            negatives += 1
    if len(data) != TRIGGER_COUNT:
        rep.warn(name, f"evals/trigger.json 共 {len(data)} 条，规范要求 {TRIGGER_COUNT} 条")
    if positives != negatives:
        rep.warn(name, f"evals/trigger.json 正反例不平衡：应触发 {positives} 条，不应触发 {negatives} 条")


def is_handoff_case(item: dict) -> bool:
    """判断一条评测用例是否覆盖交接卡或转交场景。

    现有 22 个 skill 的 evals 用 tags 标记：绝大多数写"交接卡"，个别写"转交""转交direct""转交行业包"等，
    没有更细的统一字段。所以这样判断，满足任一条即算：
    1. tags 中有含"交接卡"或"转交"的标签（子串匹配，兼容"转交行业包"这类写法）；
    2. expectations 中有一条提到"交接卡"（兼容没打标签、但期望里检查交接卡的用例）。
    """
    tags = item.get("tags")
    if _is_str_list(tags) and any("交接卡" in t or "转交" in t for t in tags):
        return True
    expectations = item.get("expectations")
    return _is_str_list(expectations) and any("交接卡" in e for e in expectations)


def _check_evals_data(name: str, skill_dir: Path, data: dict, is_draft: bool, rep: Report) -> None:
    skill_name = data.get("skill_name")
    if skill_name != name:
        rep.error(name, f"evals/evals.json 的 skill_name \"{skill_name}\" 与目录名不一致")
    evals = data.get("evals")
    if not isinstance(evals, list):
        rep.error(name, "evals/evals.json：evals 必须是数组")
        return

    ids: Set[int] = set()
    units: Set[str] = set()
    roles: Set[str] = set()
    anti_fabrication = 0
    secrecy = 0
    handoff = 0
    for index, item in enumerate(evals, start=1):
        where = f"evals/evals.json 第 {index} 条"
        if not isinstance(item, dict):
            rep.error(name, f"{where}：必须是对象")
            continue
        eid = item.get("id")
        if not isinstance(eid, int) or isinstance(eid, bool):
            rep.error(name, f"{where}：id 必须是整数")
        elif eid in ids:
            rep.error(name, f"{where}：id {eid} 重复")
        else:
            ids.add(eid)
        prompt = item.get("prompt")
        if not isinstance(prompt, str) or not prompt.strip():
            rep.error(name, f"{where}：prompt 必须是非空字符串")
        if not isinstance(item.get("expected_output"), str):
            rep.error(name, f"{where}：expected_output 必须是字符串")
        expectations = item.get("expectations")
        if not _is_str_list(expectations) or not expectations:
            rep.error(name, f"{where}：expectations 必须是非空字符串数组")
        files = item.get("files", [])
        if not _is_str_list(files):
            rep.error(name, f"{where}：files 必须是字符串数组")
        else:
            for rel_raw in files:
                rel, err = normalize_rel_path(rel_raw)
                if err:
                    rep.error(name, f"{where}：files {err}")
                elif not exists_exact(skill_dir, rel, want_dir=False):
                    rep.error(name, f"{where}：files 引用的文件不存在（注意大小写）：{rel}")
        unit = item.get("unit_type")
        if unit is not None:
            if not isinstance(unit, str):
                rep.error(name, f"{where}：unit_type 必须是字符串")
            else:
                units.add(unit)
                if unit not in UNIT_TYPES:
                    rep.warn(name, f"{where}：unit_type \"{unit}\" 不在统一词表中")
        role = item.get("role_mode")
        if role is not None:
            if not isinstance(role, str):
                rep.error(name, f"{where}：role_mode 必须是字符串")
            else:
                roles.add(role)
                if role not in ROLE_MODES:
                    rep.warn(name, f"{where}：role_mode \"{role}\" 不在统一词表中")
        tags = item.get("tags", [])
        if not _is_str_list(tags):
            rep.error(name, f"{where}：tags 必须是字符串数组")
        else:
            if "防编造" in tags:
                anti_fabrication += 1
            if "保密" in tags:
                secrecy += 1
        if is_handoff_case(item):
            handoff += 1

    if is_draft:
        return
    if len(evals) < MIN_EVALS:
        rep.warn(name, f"evals/evals.json 只有 {len(evals)} 条，规范要求至少 {MIN_EVALS} 条")
    if len(units) < 4:
        rep.warn(name, f"evals 覆盖的单位类型 {len(units)} 种，规范要求至少 4 种")
    if len(roles) < 3:
        rep.warn(name, f"evals 覆盖的角色模式 {len(roles)} 种，规范要求至少 3 种")
    if anti_fabrication < 2:
        rep.warn(name, f"evals 中标注\"防编造\"的用例 {anti_fabrication} 条，规范要求至少 2 条")
    if secrecy < 1:
        rep.warn(name, "evals 中缺少标注\"保密\"的用例")
    if handoff < 1:
        # 开发规范第 9 节要求至少 1 条交接卡用例；22 个 skill 实际全部满足，所以按错误处理
        rep.error(name, "evals 中缺少交接卡或转交场景的用例（tags 含\"交接卡\"或\"转交\"，或 expectations 提到交接卡）")


def check_skill(repo_root: Path, skill_dir: Path, rep: Report, checked_json: Set[Path],
                eval_overlap: bool = False) -> None:
    name = skill_dir.name
    rep.touch(name)
    check_skill_md(name, skill_dir, rep)
    check_shared_references(repo_root, name, skill_dir, rep)
    if not exists_exact(skill_dir, "README.md", want_dir=False):
        rep.error(name, "缺少 README.md")
    manifest_path = skill_dir / "manifest.json"
    if manifest_path.is_file():
        checked_json.add(manifest_path.resolve())
    manifest = check_manifest(repo_root, name, skill_dir, rep)
    status = manifest.get("status") if isinstance(manifest, dict) else None
    check_evals(name, skill_dir, status, rep, checked_json)
    if eval_overlap:
        check_eval_overlap(name, skill_dir, rep)


# ---------------------------------------------------------------------------
# 可选检查：示例是否复用评测材料（--eval-overlap）
# ---------------------------------------------------------------------------

def _normalize_overlap_text(text: str) -> str:
    return re.sub(r"\s+", "", text.translate(FULLWIDTH_DIGITS))


def overlap_tokens(text: str) -> List[str]:
    """抽取带单位的数字（整数部分不少于 3 位或带小数）和完整日期。

    年份（"2026年"）、条款号（"第305条"）、标准号（"GB/T 9704"）、文号（"〔2019〕34号"）、字数（"1500字"）不抽取；
    "100%"、年初和季末年末日期（1月1日、3月31日、6月30日、9月30日、12月31日）、评测统一使用的当前日期、
    法规标准的固定日期不算重合（见 EVAL_ALLOWED 等白名单常量）。
    """
    norm = _normalize_overlap_text(text)
    tokens: List[str] = []
    for m in EVAL_NUMBER_RE.finditer(norm):
        number, unit = m.group(1), m.group(2)
        if number.startswith("0") and "." not in number:
            continue  # 编号、电话区号
        token = number + unit
        if token not in EVAL_ALLOWED and token not in tokens:
            tokens.append(token)
    for regex in EVAL_DATE_RES:
        for m in regex.finditer(norm):
            year, month, day = (int(g) for g in m.groups())
            if 1 <= month <= 12 and 1 <= day <= 31 and (month, day) not in EVAL_ALLOWED_MONTH_DAYS:
                token = f"{year}年{month}月{day}日"
                if token not in tokens and token not in EVAL_ALLOWED_DATES:
                    tokens.append(token)
    return tokens


def _eval_file_text(path: Path) -> Optional[str]:
    """读取评测材料的文字；docx、xlsx 只取 XML 中的文字，读不了的返回 None。"""
    suffix = path.suffix.lower()
    try:
        if suffix in EVAL_TEXT_SUFFIXES:
            data = path.read_bytes()
            for encoding in ("utf-8-sig", "gb18030"):
                try:
                    return data.decode(encoding)
                except UnicodeDecodeError:
                    continue
            return None
        parts = EVAL_ZIP_PARTS.get(suffix)
        if not parts:
            return None
        texts: List[str] = []
        with zipfile.ZipFile(str(path)) as archive:
            for info in archive.infolist():
                wanted = any(info.filename == p or (p.endswith("/") and info.filename.startswith(p)) for p in parts)
                if wanted and info.file_size <= MAX_EVAL_ZIP_PART_BYTES:
                    xml = archive.read(info).decode("utf-8", errors="replace")
                    xml = re.sub(r"</w:p>|</si>|</row>", "\n", xml)  # 段落、单元格之间断行，避免数字粘连
                    texts.append(XML_TAG_RE.sub("", xml))
        return "\n".join(texts)
    except (zipfile.BadZipFile, OSError, RuntimeError):
        return None


def _eval_sources(skill_dir: Path) -> List[Tuple[str, str]]:
    """评测来源：evals/files/ 下的文件；evals.json 各用例的 prompt、expected_output、expectations；
    evals/trigger.json 的 query。每项为 (来源说明, 文本)。"""
    sources: List[Tuple[str, str]] = []
    files_dir = skill_dir / "evals" / "files"
    if files_dir.is_dir():
        for path in sorted(p for p in files_dir.rglob("*") if p.is_file()):
            text = _eval_file_text(path)
            if text:
                sources.append((path.relative_to(skill_dir).as_posix(), text))
    evals_path = skill_dir / "evals" / "evals.json"
    if evals_path.is_file():
        data, err = load_json_file(evals_path)
        if not err and isinstance(data, dict) and isinstance(data.get("evals"), list):
            for item in data["evals"]:
                if not isinstance(item, dict):
                    continue
                label = f"evals.json 第 {item.get('id')} 条"
                for field in ("prompt", "expected_output"):
                    if isinstance(item.get(field), str):
                        sources.append((f"{label} {field}", item[field]))
                if _is_str_list(item.get("expectations")):
                    # 一条用例的全部期望并成一段，来源只标到条，不再细分第几句
                    sources.append((f"{label} expectations", "\n".join(item["expectations"])))
    trigger_path = skill_dir / "evals" / "trigger.json"
    if trigger_path.is_file():
        data, err = load_json_file(trigger_path)
        if not err and isinstance(data, list):
            for index, item in enumerate(data, start=1):
                if isinstance(item, dict) and isinstance(item.get("query"), str):
                    sources.append((f"trigger.json 第 {index} 项 query", item["query"]))
    return sources


def eval_token_sources(skill_dir: Path) -> Dict[str, str]:
    """一个 skill 的全部评测来源中抽取到的 {token: 首次出现的来源}。"""
    token_source: Dict[str, str] = {}
    for source, text in _eval_sources(skill_dir):
        for token in overlap_tokens(text):
            token_source.setdefault(token, source)
    return token_source


def _scan_overlap_targets(token_source: Dict[str, str], targets: Sequence[Tuple[str, Path]]) -> Dict[str, List[str]]:
    """在被扫描文件中找出与评测 token 相同的数字、日期，返回 {token: [位置, ...]}。

    被扫描文件用与评测来源相同的 overlap_tokens 抽取后再取交集，两边口径一致（含白名单）；
    docx、xlsx 取文字后按行（段）报位置。
    """
    hits: Dict[str, List[str]] = {}
    for rel, path in targets:
        text = _eval_file_text(path)
        if not text:
            continue
        for lineno, line in enumerate(text.splitlines(), start=1):
            if not any(ch.isdigit() or ch in "０１２３４５６７８９" for ch in line):
                continue
            for token in overlap_tokens(line):
                if token in token_source:
                    places = hits.setdefault(token, [])
                    place = f"{rel} 第 {lineno} 行"
                    if place not in places:
                        places.append(place)
    return hits


def _skill_overlap_targets(skill_dir: Path) -> List[Tuple[str, Path]]:
    """skill 内被扫描的文件：SKILL.md、README.md、references/ 下的 .md（不含 _shared，那份由 shared/ 源文件统一报）。"""
    targets: List[Tuple[str, Path]] = [(name, skill_dir / name) for name in ("SKILL.md", "README.md")]
    refs_dir = skill_dir / "references"
    if refs_dir.is_dir():
        targets += [(p.relative_to(skill_dir).as_posix(), p) for p in sorted(refs_dir.rglob("*.md"))
                    if p.is_file() and "_shared" not in p.relative_to(refs_dir).parts]
    return [(rel, path) for rel, path in targets if path.is_file()]


def check_eval_overlap(name: str, skill_dir: Path, rep: Report) -> None:
    """评测来源中的数字、日期出现在本 skill 的 SKILL.md、README.md 或 references/（不含 _shared）中时，报警告。"""
    token_source = eval_token_sources(skill_dir)
    if not token_source:
        return
    hits = _scan_overlap_targets(token_source, _skill_overlap_targets(skill_dir))
    for token, places in hits.items():
        rep.warn(name, f"示例可能复用评测材料：{token}（评测材料：{token_source[token]}；"
                       f"出现在：{_join_places(places, '处')}）")


def _shared_consumers(skill_dirs: Sequence[Path]) -> Dict[str, List[Path]]:
    """{shared 下的源文件相对路径: 把它同步进去的 skill 目录列表}，依据各 skill 的 manifest.json。"""
    consumers: Dict[str, List[Path]] = {}
    for skill_dir in skill_dirs:
        data, err = load_json_file(skill_dir / "manifest.json")
        if err or not isinstance(data, dict) or not isinstance(data.get("shared"), list):
            continue
        for item in data["shared"]:
            if isinstance(item, dict):
                src, _ = normalize_rel_path(item.get("src"))
                if src:
                    consumers.setdefault(src, []).append(skill_dir)
    return consumers


def check_repo_eval_overlap(repo_root: Path, skill_dirs: Sequence[Path], rep: Report) -> None:
    """全仓库的 --eval-overlap：shared/ 下的源文件，以及 examples/。

    shared/ 的 .md 源文件只与把它同步进去的那些 skill 的评测材料比对（别的 skill 读不到它），
    重合报到 shared 源文件上，不对各 skill 的 _shared 副本重复报。
    examples/ 是公开展示的前后对比案例，与全部 skill 的评测材料比对。
    """
    rep.touch(REPO_SCOPE)
    token_cache: Dict[Path, Dict[str, str]] = {}

    def merged(dirs: Sequence[Path]) -> Dict[str, str]:
        result: Dict[str, str] = {}
        for d in dirs:
            if d not in token_cache:
                token_cache[d] = eval_token_sources(d)
            for token, source in token_cache[d].items():
                result.setdefault(token, f"{d.name} 的 {source}")
        return result

    shared_dir = repo_root / "shared"
    if shared_dir.is_dir():
        consumers = _shared_consumers(skill_dirs)
        for path in sorted(p for p in shared_dir.rglob("*.md") if p.is_file()):
            rel = path.relative_to(repo_root).as_posix()
            users = consumers.get(path.relative_to(shared_dir).as_posix(), [])
            token_source = merged(users)
            if not token_source:
                continue
            for token, places in _scan_overlap_targets(token_source, [(rel, path)]).items():
                rep.warn(REPO_SCOPE, f"示例可能复用评测材料：{token}（评测材料：{token_source[token]}；"
                                     f"出现在：{_join_places(places, '处')}；该文件同步进 {len(users)} 个 skill）")

    examples_dir = repo_root / "examples"
    if examples_dir.is_dir():
        token_source = merged(skill_dirs)
        suffixes = EVAL_TEXT_SUFFIXES | set(EVAL_ZIP_PARTS)
        targets = [(p.relative_to(repo_root).as_posix(), p) for p in sorted(examples_dir.rglob("*"))
                   if p.is_file() and p.suffix.lower() in suffixes]
        for token, places in _scan_overlap_targets(token_source, targets).items():
            rep.warn(REPO_SCOPE, f"示例可能复用评测材料：{token}（评测材料：{token_source[token]}；"
                                 f"出现在：{_join_places(places, '处')}）")


# ---------------------------------------------------------------------------
# 插件清单与全仓库检查
# ---------------------------------------------------------------------------

def check_marketplace(repo_root: Path, skill_dirs: List[Path], rep: Report, checked_json: Set[Path]) -> None:
    path = repo_root / ".claude-plugin" / "marketplace.json"
    if not path.is_file():
        rep.warn(REPO_SCOPE, "未找到 .claude-plugin/marketplace.json，已跳过插件清单检查")
        return
    checked_json.add(path.resolve())
    data, err = load_json_file(path)
    if err:
        rep.error(REPO_SCOPE, f".claude-plugin/marketplace.json：{err}")
        return
    if not isinstance(data, dict) or not isinstance(data.get("plugins"), list):
        rep.error(REPO_SCOPE, "marketplace.json：缺少 plugins 数组")
        return

    skills_root = os.path.normcase(os.path.abspath(repo_root / "skills"))
    covered: Set[str] = set()
    for index, plugin in enumerate(data["plugins"], start=1):
        if not isinstance(plugin, dict):
            rep.error(REPO_SCOPE, f"marketplace.json 第 {index} 个 plugin 必须是对象")
            continue
        pname = plugin.get("name")
        label = pname if isinstance(pname, str) else f"第 {index} 个"
        if not isinstance(pname, str) or not pname:
            rep.error(REPO_SCOPE, f"marketplace.json 第 {index} 个 plugin 缺少 name")
        elif pname.lower().startswith(("claude-", "anthropic-")):
            rep.error(REPO_SCOPE, f"plugin 名 {pname} 不能以 claude- 或 anthropic- 开头")
        source = plugin.get("source", "./")
        if not isinstance(source, str) or not source.startswith("."):
            rep.warn(REPO_SCOPE, f"plugin {label} 的 source 不是本地相对路径，已跳过其 skills 路径检查")
            continue
        base = repo_root / source
        skills = plugin.get("skills")
        if skills is None:
            # 未列出 skills 时，按插件目录下的 skills/ 自动发现
            for sub in list_skill_dirs(base):
                covered.add(os.path.normcase(os.path.abspath(sub)))
            continue
        if not _is_str_list(skills):
            rep.error(REPO_SCOPE, f"plugin {label} 的 skills 必须是字符串数组")
            continue
        for rel in skills:
            target = Path(os.path.abspath(base / rel))
            if not (target.is_dir() and (target / "SKILL.md").is_file()):
                rep.error(REPO_SCOPE, f"plugin {label} 收录的 skill 路径不存在或缺少 SKILL.md：{rel}")
                continue
            covered.add(os.path.normcase(str(target)))

    for skill_dir in skill_dirs:
        key = os.path.normcase(os.path.abspath(skill_dir))
        if os.path.dirname(key) == skills_root and key not in covered:
            rep.error(REPO_SCOPE, f"skills/{skill_dir.name} 未被 marketplace.json 中任何 plugin 收录")


def _iter_repo_files(repo_root: Path):
    for dirpath, dirnames, filenames in os.walk(repo_root):
        current = Path(dirpath)
        rel_dir = current.relative_to(repo_root).as_posix()
        kept = []
        for d in sorted(dirnames):
            rel = d if rel_dir == "." else f"{rel_dir}/{d}"
            if d in SKIP_DIR_NAMES or rel in SKIP_REL_DIRS:
                continue
            if d in SKIP_TEST_DIR_NAMES and rel.startswith("tests/"):
                continue
            if (current / d / "pyvenv.cfg").is_file():  # 虚拟环境
                continue
            kept.append(d)
        dirnames[:] = kept
        for name in sorted(filenames):
            yield current / name


def is_large_file_exempt(rel: str) -> bool:
    """rel 为相对仓库根的 POSIX 路径；匹配 LARGE_FILE_EXEMPT_GLOBS 的文件不做 1MB 体量警告。"""
    return any(fnmatch.fnmatchcase(rel, pattern) for pattern in LARGE_FILE_EXEMPT_GLOBS)


def check_repository(repo_root: Path, rep: Report, checked_json: Set[Path]) -> None:
    rep.touch(REPO_SCOPE)
    for path in _iter_repo_files(repo_root):
        rel = path.relative_to(repo_root).as_posix()
        suffix = path.suffix.lower()
        if suffix in FONT_SUFFIXES:
            rep.error(REPO_SCOPE, f"不得打包字体文件：{rel}")
        try:
            size = path.stat().st_size
        except OSError:
            continue
        if size > MAX_FILE_BYTES and not is_large_file_exempt(rel):
            rep.warn(REPO_SCOPE, f"文件超过 1MB（{size / 1024 / 1024:.1f}MB）：{rel}")
        if suffix == ".json" and path.resolve() not in checked_json:
            _data, err = load_json_file(path)
            if err:
                rep.error(REPO_SCOPE, f"{rel}：{err}")


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

def run_validation(repo_root: Path, skill: Optional[str] = None,
                   eval_overlap: bool = False) -> Tuple[Report, int]:
    """执行校验，返回 (报告, 已检查的 skill 数)。skill 不存在时抛出 LookupError。

    eval_overlap：同时运行可选的"示例是否复用评测材料"检查（结果记为警告）。
    """
    rep = Report()
    checked_json: Set[Path] = set()
    skill_dirs = list_skill_dirs(repo_root)
    if skill:
        selected = [d for d in skill_dirs if d.name == skill]
        if not selected:
            raise LookupError(skill)
        for skill_dir in selected:
            check_skill(repo_root, skill_dir, rep, checked_json, eval_overlap)
        return rep, len(selected)

    if not skill_dirs:
        rep.warn(REPO_SCOPE, "skills/ 下还没有任何 skill，已跳过 skill 检查")
    for skill_dir in skill_dirs:
        check_skill(repo_root, skill_dir, rep, checked_json, eval_overlap)
    check_marketplace(repo_root, skill_dirs, rep, checked_json)
    check_repository(repo_root, rep, checked_json)
    if eval_overlap:
        check_repo_eval_overlap(repo_root, skill_dirs, rep)
    return rep, len(skill_dirs)


def main(argv: Optional[Sequence[str]] = None, repo_root: Optional[Path] = None) -> int:
    setup_console()
    parser = make_parser(
        "validate.py",
        "校验各 skill 的 frontmatter、manifest、共享资料同步、相对路径、evals，以及插件清单和全仓库合规项。",
        "退出码：0 无错误；1 存在错误；2 参数错误。",
    )
    parser.add_argument("--skill", metavar="NAME", help="只检查指定 skill（同时跳过插件清单和全仓库检查）")
    parser.add_argument("--quiet", action="store_true", help="只输出错误，不输出警告和通过信息")
    parser.add_argument("--eval-overlap", action="store_true",
                        help="另查示例是否复用评测材料：evals/files/、evals.json 的 prompt、expected_output、"
                             "expectations 和 trigger.json 的 query 中带单位的数字（3 位以上或带小数）、完整日期，"
                             "出现在各 skill 的 SKILL.md、README.md、references/（不含 _shared）中的记为警告；"
                             "全仓库检查时另查 shared/ 源文件和 examples/（用 --skill 时不查这两处）。"
                             "评测当前日期、法规标准的固定日期等在白名单内；默认不运行")
    add_root_argument(parser)
    args = parser.parse_args(argv)

    root = (args.root or repo_root or REPO_ROOT).resolve()
    try:
        rep, count = run_validation(root, args.skill, eval_overlap=args.eval_overlap)
    except LookupError:
        print(f"错误：skills/ 下找不到名为 {args.skill} 的 skill")
        return 2

    output = rep.render(quiet=args.quiet)
    if output:
        print(output)
    errors, warnings = len(rep.errors), len(rep.warnings)
    print(f"校验结束：检查 {count} 个 skill，错误 {errors} 个，警告 {warnings} 个。")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
