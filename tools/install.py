#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 skills/ 下的 skill 复制到各 Agent 工具的 skill 目录；也可探测本机工具、卸载本项目安装的 skill。

示例：
  python tools/install.py --list-targets
  python tools/install.py --detect
  python tools/install.py --target auto
  python tools/install.py --target qwen --pack hospital
  python tools/install.py --target claude,cursor --skills gongwen,gongwen-report
  python tools/install.py --target trae --scope project --dry-run
  python tools/install.py --target auto --uninstall          # 只列出将要删除的目录
  python tools/install.py --target auto --uninstall --yes    # 确认后删除

安装哪些 skill：默认只装核心插件 gongwen 的 skill，行业包用 --pack 选装，清单都读自
.claude-plugin/marketplace.json；仓库中没有该文件时，安装 skills/ 下的全部 skill。

目标目录已存在同名 skill 时：默认跳过并提示；加 --force 先把旧目录改名为
<名>.bak-<时间戳> 备份，再复制新版本。安装过程不删除任何文件。
--uninstall 只删除"目录名是 marketplace.json 中的 skill 名、且目录内 SKILL.md 的
name 字段与目录名一致"的目录；不加 --yes 或 --force 时只列出，不删除。
--detect 和 --target auto 只检查用户目录和命令是否存在，不联网、不运行这些工具。
退出码：0 成功；1 部分 skill 安装或删除失败；2 参数或配置错误；3 --target auto 未探测到任何工具。
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence, Tuple

REPO_ROOT = Path(__file__).resolve().parent.parent
UNVERIFIED_NOTE = "该路径来自社区资料，如不生效请查阅该工具文档"
# .gongwen（单位画像）和 公文（用户产出）可能含内部信息，不复制到安装目录
COPY_IGNORE = shutil.ignore_patterns(
    "__pycache__", "*.pyc", "*.pyo", ".DS_Store", "Thumbs.db", ".gongwen", "公文",
)
# 卸载时，skill 目录内出现这些子目录即视为含用户资料，整个目录保留不删
USER_DATA_DIRS = (".gongwen", "公文")
AUTO_TARGET = "auto"
MARKETPLACE_REL = ".claude-plugin/marketplace.json"
CORE_PLUGIN = "gongwen"
PACK_PREFIX = "gongwen-"
ALL_PACKS = "all"
SKILL_NAME_RE = re.compile(r"^[a-z0-9][a-z0-9-]{0,63}$")
# 排版功能的 Python 依赖：(pip 包名, 导入名)
FORMAT_DEPS: Tuple[Tuple[str, str], ...] = (("python-docx", "docx"), ("openpyxl", "openpyxl"))
FORMAT_DEPS_CMD = "pip install python-docx openpyxl"
# 排版 skill 的字体检查脚本：公文排版要求标准字体，缺字体时不会改用其他字体顶替，安装后顺带检查一次
FORMAT_SKILL = "gongwen-format"
FONT_CHECK_REL = "skills/gongwen-format/scripts/gw_fonts.py"
NO_TOOL_HINT = (
    "未探测到已安装的 Agent 工具。可以用 --target 手动指定（--list-targets 查看可选目标），"
    "或者使用提示词版：打开 dist/prompts/，把对应文件复制到网页大模型的对话中。"
)

WhichFunc = Callable[[str], Optional[str]]


@dataclass(frozen=True)
class Target:
    key: str
    label: str
    user_path: str        # 以 ~/ 开头
    project_path: str     # 相对当前目录
    user_verified: bool
    project_verified: bool

    def path_for(self, scope: str) -> str:
        return self.user_path if scope == "user" else self.project_path

    def verified(self, scope: str) -> bool:
        return self.user_verified if scope == "user" else self.project_verified


TARGETS: Tuple[Target, ...] = (
    Target("claude", "Claude Code", "~/.claude/skills", ".claude/skills", True, True),
    Target("codex", "OpenAI Codex", "~/.agents/skills", ".agents/skills", True, True),
    Target("agents", "通用 Agent Skills 目录", "~/.agents/skills", ".agents/skills", True, True),
    Target("cursor", "Cursor", "~/.cursor/skills", ".cursor/skills", True, True),
    Target("gemini", "Gemini CLI", "~/.gemini/skills", ".gemini/skills", True, True),
    Target("copilot", "GitHub Copilot", "~/.copilot/skills", ".github/skills", True, True),
    Target("qwen", "Qwen Code", "~/.qwen/skills", ".qwen/skills", True, True),
    Target("kimi", "Kimi Code CLI", "~/.kimi-code/skills", ".agents/skills", True, True),
    Target("codebuddy", "CodeBuddy", "~/.codebuddy/skills", ".codebuddy/skills", False, True),
    Target("trae", "Trae", "~/.trae/skills", ".trae/skills", False, False),
)
TARGET_MAP: Dict[str, Target] = {t.key: t for t in TARGETS}

# 探测依据：(用户目录下的配置目录, 命令名)。只判断是否存在，不运行这些命令。
DETECT_HINTS: Dict[str, Tuple[Tuple[str, ...], Tuple[str, ...]]] = {
    "claude": ((".claude",), ("claude",)),
    "codex": ((".codex",), ("codex",)),
    "agents": ((".agents",), ()),
    "cursor": ((".cursor",), ("cursor", "cursor-agent")),
    "gemini": ((".gemini",), ("gemini",)),
    "copilot": ((".copilot",), ("copilot",)),
    "qwen": ((".qwen",), ("qwen",)),
    "kimi": ((".kimi-code",), ("kimi",)),
    "codebuddy": ((".codebuddy",), ("codebuddy",)),
    "trae": ((".trae",), ("trae",)),
}


class UsageError(Exception):
    """参数或仓库配置错误，main 捕获后以退出码 2 结束。"""


@dataclass
class Catalog:
    """marketplace.json 中的 skill 清单。"""
    core: List[str]                 # 核心插件 gongwen 的 skill
    packs: Dict[str, List[str]]     # 行业包短名（如 hospital）-> skill 名
    all_skills: List[str]           # 全部插件的 skill，去重保序


def setup_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


class ChineseHelpFormatter(argparse.RawDescriptionHelpFormatter):
    def add_usage(self, usage, actions, groups, prefix=None):
        if prefix is None:
            prefix = "用法："
        return super().add_usage(usage, actions, groups, prefix)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="install.py",
        description=(
            "把 gongwen skill 复制到各 Agent 工具的 skill 目录。安装不删除任何文件；"
            "--uninstall 只删除本项目安装的 skill 目录，且需 --yes 确认。"
        ),
        epilog="退出码：0 成功；1 部分 skill 安装或删除失败；2 参数或配置错误；3 --target auto 未探测到任何工具。",
        formatter_class=ChineseHelpFormatter,
        add_help=False,
    )
    parser._optionals.title = "选项"
    parser.add_argument("-h", "--help", action="help", help="显示本帮助信息并退出")
    parser.add_argument("--list-targets", action="store_true", help="列出支持的目标及其路径")
    parser.add_argument(
        "--detect",
        action="store_true",
        help="探测本机已安装的 Agent 工具（只看用户目录和命令是否存在，不联网、不运行这些工具）",
    )
    parser.add_argument(
        "--target",
        action="append",
        metavar="NAME",
        help="安装目标，可多选：--target claude --target qwen，或 --target claude,qwen；"
             "auto 表示 --detect 探测到的全部工具",
    )
    parser.add_argument(
        "--scope",
        choices=("user", "project"),
        default="user",
        help="user 安装到用户目录（默认）；project 安装到当前目录下的对应目录",
    )
    parser.add_argument(
        "--pack",
        action="append",
        metavar="NAME",
        help="选装行业包，可多选：hospital、university、k12、soe、grassroots，或 all（全部）；"
             "如 --pack hospital --pack soe 或 --pack hospital,soe",
    )
    parser.add_argument(
        "--skills",
        metavar="A,B",
        help="只安装指定 skill，逗号分隔（默认：核心插件 gongwen 的全部 skill，再加 --pack 选中的行业包）",
    )
    parser.add_argument("--dry-run", action="store_true", help="只打印将要执行的操作，不做任何改动")
    parser.add_argument(
        "--force",
        action="store_true",
        help="安装时：目标已存在同名 skill，先改名备份为 <名>.bak-<时间戳> 再复制；"
             "卸载时：不再要求 --yes，直接删除",
    )
    parser.add_argument(
        "--uninstall",
        action="store_true",
        help="卸载本项目安装的 skill（只删目录名与 SKILL.md 的 name 都匹配 marketplace.json 的目录）；"
             "可配合 --pack、--skills 只卸载一部分",
    )
    parser.add_argument("--yes", action="store_true", help="确认执行 --uninstall 列出的删除操作")
    parser.add_argument("--root", type=Path, metavar="DIR", help="仓库根目录（默认为本脚本所在仓库，主要供测试使用）")
    return parser


def _split_csv(values: Optional[Sequence[str]]) -> List[str]:
    result: List[str] = []
    for value in values or []:
        for item in value.split(","):
            item = item.strip()
            if item and item not in result:
                result.append(item)
    return result


def list_installable_skills(repo_root: Path) -> List[str]:
    skills_dir = repo_root / "skills"
    if not skills_dir.is_dir():
        return []
    return sorted(
        p.name for p in skills_dir.iterdir()
        if p.is_dir() and not p.name.startswith((".", "_")) and (p / "SKILL.md").is_file()
    )


def resolve_base(target: Target, scope: str, home: Path, cwd: Path) -> Path:
    raw = target.path_for(scope)
    if scope == "user":
        return home / raw[2:] if raw.startswith("~/") else home / raw
    return cwd / raw


def _dedup(items: Sequence[str]) -> List[str]:
    result: List[str] = []
    for item in items:
        if item not in result:
            result.append(item)
    return result


# ---------- marketplace.json 清单 ----------

def _skill_name_from_path(raw: object) -> Optional[str]:
    """把 "./skills/gongwen" 解析为 "gongwen"；不是 skills/<小写名> 形式的返回 None。"""
    if not isinstance(raw, str):
        return None
    parts = [p for p in raw.replace("\\", "/").split("/") if p not in ("", ".")]
    if len(parts) != 2 or parts[0] != "skills" or not SKILL_NAME_RE.match(parts[1]):
        return None
    return parts[1]


def load_catalog(repo_root: Path) -> Optional[Catalog]:
    """读取 marketplace.json 的插件与 skill 清单。文件不存在返回 None；格式不合规抛 UsageError。"""
    path = repo_root / MARKETPLACE_REL
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeDecodeError, ValueError) as exc:
        raise UsageError(f"无法读取 {MARKETPLACE_REL}：{exc}") from exc
    plugins = data.get("plugins") if isinstance(data, dict) else None
    if not isinstance(plugins, list):
        raise UsageError(f"{MARKETPLACE_REL} 缺少 plugins 列表")

    core: List[str] = []
    packs: Dict[str, List[str]] = {}
    all_skills: List[str] = []
    for plugin in plugins:
        if not isinstance(plugin, dict) or not isinstance(plugin.get("name"), str):
            continue
        plugin_name = plugin["name"]
        raw_skills = plugin.get("skills")
        if not isinstance(raw_skills, list):
            raw_skills = []
        names: List[str] = []
        for raw in raw_skills:
            name = _skill_name_from_path(raw)
            if name is None:
                raise UsageError(f"{MARKETPLACE_REL} 中插件 {plugin_name} 的 skill 路径不合规：{raw!r}（应为 ./skills/<小写名>）")
            names.append(name)
        names = _dedup(names)
        if plugin_name == CORE_PLUGIN:
            core = names
        else:
            key = plugin_name[len(PACK_PREFIX):] if plugin_name.startswith(PACK_PREFIX) else plugin_name
            packs[key] = names
        all_skills = _dedup(all_skills + names)
    return Catalog(core=core, packs=packs, all_skills=all_skills)


def resolve_packs(raw: Optional[Sequence[str]], catalog: Optional[Catalog]) -> List[str]:
    """把 --pack 的取值解析为行业包短名列表。"""
    packs = _split_csv(raw)
    if not packs:
        return []
    if catalog is None:
        raise UsageError(f"找不到 {MARKETPLACE_REL}，无法按 --pack 选装行业包")
    if ALL_PACKS in packs:
        return list(catalog.packs)
    unknown = [p for p in packs if p not in catalog.packs]
    if unknown:
        choices = "、".join(list(catalog.packs) + [ALL_PACKS])
        raise UsageError(f"不支持的行业包 {', '.join(unknown)}。可选：{choices}")
    return packs


def resolve_install_skills(
    skills_arg: Optional[str], packs: List[str], catalog: Optional[Catalog], available: List[str]
) -> Tuple[List[str], List[str]]:
    """确定要安装的 skill。返回 (要安装的, 清单中列出但 skills/ 下缺失而跳过的)。"""
    explicit = _split_csv([skills_arg]) if skills_arg else []
    missing = [s for s in explicit if s not in available]
    if missing:
        raise UsageError(f"skills/ 下找不到 {', '.join(missing)}（或缺少 SKILL.md）")
    if catalog is None:
        # 兼容旧仓库：没有 marketplace.json 时安装 skills/ 下的全部 skill
        return (explicit or list(available)), []
    listed = list(explicit) if explicit else list(catalog.core)
    for pack in packs:
        listed += catalog.packs[pack]
    listed = _dedup(listed)
    return [s for s in listed if s in available], [s for s in listed if s not in available]


# ---------- 探测本机工具 ----------

def detect_targets(home: Path, which: Optional[WhichFunc] = None) -> Dict[str, List[str]]:
    """探测已安装的工具，返回 {目标名: [依据, ...]}，顺序同 TARGETS。

    只检查用户目录下的配置目录、PATH 中的命令是否存在，不联网、不运行这些命令。
    """
    which = which or shutil.which
    found: Dict[str, List[str]] = {}
    for target in TARGETS:
        dirs, commands = DETECT_HINTS.get(target.key, ((), ()))
        evidence: List[str] = []
        for rel in dirs:
            try:
                if (home / rel).is_dir():
                    evidence.append(f"~/{rel} 目录")
            except OSError:
                pass
        for command in commands:
            try:
                located = which(command)
            except Exception:  # 探测失败按未发现处理，不中断安装
                located = None
            if located:
                evidence.append(f"{command} 命令")
        if evidence:
            found[target.key] = evidence
    return found


def print_detection(found: Dict[str, List[str]], scope: str) -> None:
    print("本机 Agent 工具探测结果（只检查用户目录和命令是否存在，不联网、不运行这些工具）：")
    for t in TARGETS:
        if t.key in found:
            print(f"  [已发现] {t.key:<10} {t.label}（依据：{'；'.join(found[t.key])}；安装到 {t.path_for(scope)}）")
        else:
            print(f"  [未发现] {t.key:<10} {t.label}")
    if found:
        print(f"共发现 {len(found)} 个：{'、'.join(found)}。")
        print("安装到以上全部工具：python tools/install.py --target auto")
    else:
        print(NO_TOOL_HINT)


# ---------- 卸载 ----------

def read_skill_name(skill_md: Path) -> Optional[str]:
    """读取 SKILL.md 开头 frontmatter 中的 name 字段；读不到返回 None。兼容 BOM 和 CRLF 换行。"""
    try:
        text = skill_md.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError):
        return None
    lines = text.splitlines()
    if not lines or lines[0].strip() != "---":
        return None
    for line in lines[1:]:
        if line.strip() == "---":
            break
        match = re.match(r"^name\s*:\s*(.*?)\s*$", line)
        if match:
            value = match.group(1)
            if len(value) >= 2 and value[0] == value[-1] and value[0] in "'\"":
                value = value[1:-1].strip()
            return value
    return None


def _is_link_like(path: Path) -> bool:
    """符号链接或 Windows 目录联接（junction）。这类目录卸载时一律不删，避免误删链接指向的内容。"""
    try:
        if path.is_symlink():
            return True
        isjunction = getattr(os.path, "isjunction", None)  # Python 3.12+
        if isjunction is not None and isjunction(path):
            return True
        attrs = getattr(os.lstat(path), "st_file_attributes", 0)
        return bool(attrs & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0))
    except OSError:
        return False


def _count_files(path: Path) -> int:
    count = 0
    for _root, _dirs, files in os.walk(path):  # 不跟随符号链接
        count += len(files)
    return count


def check_removable(dest: Path, name: str) -> Tuple[str, str]:
    """判断已安装的 skill 目录能否删除。返回 (状态, 说明)，状态为 remove / absent / keep。"""
    if _is_link_like(dest):
        return "keep", "是符号链接或目录联接，未删除，请手动处理"
    if not dest.exists():
        return "absent", ""
    if not dest.is_dir():
        return "keep", "不是目录，未删除"
    skill_md = dest / "SKILL.md"
    if not skill_md.is_file():
        return "keep", "缺少 SKILL.md，无法确认是本项目安装的，未删除"
    found = read_skill_name(skill_md)
    if found != name:
        return "keep", f"SKILL.md 的 name 字段为 {found or '（缺失）'}，与目录名不一致，未删除"
    user_data = [d for d in USER_DATA_DIRS if (dest / d).exists()]
    if user_data:
        return "keep", f"目录内有用户资料（{'、'.join(user_data)}），未删除，请先移走"
    return "remove", f"{_count_files(dest)} 个文件"


def _retry_writable(func: Callable[[str], object], path: str, _exc: object) -> None:
    """Windows 上只读文件无法删除时，去掉只读属性后重试一次；仍失败则抛出。"""
    os.chmod(path, stat.S_IWRITE)
    func(path)


def remove_tree(path: Path) -> None:
    if _is_link_like(path):  # 删除前再确认一次
        raise OSError(f"{path} 是链接，拒绝删除")
    if sys.version_info >= (3, 12):
        shutil.rmtree(path, onexc=_retry_writable)
    else:
        shutil.rmtree(path, onerror=_retry_writable)


def resolve_uninstall_skills(skills_arg: Optional[str], packs: List[str], catalog: Optional[Catalog]) -> List[str]:
    """确定卸载范围：默认 marketplace.json 中的全部 skill；给了 --skills 或 --pack 时只取其中一部分。"""
    if catalog is None:
        raise UsageError(f"找不到 {MARKETPLACE_REL}，无法确认哪些目录由本项目安装，已停止卸载")
    explicit = _split_csv([skills_arg]) if skills_arg else []
    unknown = [s for s in explicit if s not in catalog.all_skills]
    if unknown:
        raise UsageError(f"{', '.join(unknown)} 不在 {MARKETPLACE_REL} 的 skill 清单中，不能卸载")
    if not explicit and not packs:
        return list(catalog.all_skills)
    names = list(explicit)
    for pack in packs:
        names += catalog.packs[pack]
    return _dedup(names)


def run_uninstall(
    names: List[str], targets: List[str], scope: str, home: Path, cwd: Path, confirmed: bool, dry_run: bool
) -> int:
    plan: List[Path] = []
    done: Dict[str, str] = {}
    for key in targets:
        target = TARGET_MAP[key]
        base = resolve_base(target, scope, home, cwd)
        print(f"目标 {key}（{target.label}，{scope} 级）：{target.path_for(scope)} → {base}")
        norm = os.path.normcase(os.path.abspath(base))
        if norm in done:
            print(f"  与目标 {done[norm]} 的目录相同，已在上一步处理，跳过。")
            continue
        done[norm] = key
        shown = 0
        for name in names:
            status, note = check_removable(base / name, name)
            if status == "remove":
                print(f"  将删除  {name}（{note}）")
                plan.append(base / name)
                shown += 1
            elif status == "keep":
                print(f"  保留    {name}：{note}")
                shown += 1
        backups: List[str] = []
        if base.is_dir():
            backups = sorted(
                p.name for p in base.iterdir() if p.name.split(".bak-", 1)[0] in names and ".bak-" in p.name
            )
        if backups:
            print(f"  另有备份目录（不会删除，如不需要请手动处理）：{'、'.join(backups)}")
        if not shown:
            print("  未发现本项目安装的 skill。")

    if not plan:
        print("没有需要删除的目录。")
        return 0
    if dry_run:
        print(f"[dry-run] 结束：以上 {len(plan)} 个目录未删除。")
        return 0
    if not confirmed:
        print(f"以上 {len(plan)} 个目录将被删除，其他文件不受影响。确认无误后，加 --yes 重新运行。")
        return 0

    failures = 0
    for dest in plan:
        try:
            remove_tree(dest)
        except OSError as exc:
            failures += 1
            print(f"  失败  {dest}：{exc}")
            continue
        print(f"  已删除  {dest}")
    if failures:
        print(f"完成，但有 {failures} 个目录删除失败。")
        return 1
    print(f"卸载完成，共删除 {len(plan)} 个目录。单位画像（.gongwen）和已写好的公文不受影响。")
    return 0


# ---------- 安装后提示 ----------

def module_available(module: str) -> bool:
    """只查找模块，不导入。"""
    try:
        return importlib.util.find_spec(module) is not None
    except (ImportError, ValueError):
        return False


def print_next_steps(targets: List[str], wanted: List[str], catalog: Optional[Catalog], packs: List[str]) -> None:
    # "agents" 是通用目录而非具体工具，提示语中不单列
    labels = _dedup([TARGET_MAP[k].label for k in targets if k != "agents"]) or ["你的 Agent 工具"]
    first = "Claude Code" if "claude" in targets else labels[0]
    print()
    print("下一步怎么用：")
    print(f"  1. 重启 {'、'.join(labels)}（或新开一个会话），让工具加载新安装的 skill。")
    print(f"  2. 直接用中文说需求，例如在 {first} 中说：\"用公文 skill 帮我写个通知\"。")
    if "gongwen-profile" in wanted:
        print("  3. 第一次用，可以先说\"用公文 skill 帮我建个单位画像\"，之后的文稿会更贴合本单位。")
    if catalog is not None and catalog.packs and not packs:
        print(f"  行业包按需选装：加 --pack {'、'.join(catalog.packs)} 之一，或 --pack all。")
    print("排版功能（Word 国标排版、表格转 Excel）需要 python-docx 和 openpyxl：")
    missing = []
    for pip_name, module in FORMAT_DEPS:
        ok = module_available(module)
        print(f"  {pip_name:<12} {'已安装' if ok else '未安装'}")
        if not ok:
            missing.append(pip_name)
    if missing:
        print(f"  安装命令：{FORMAT_DEPS_CMD}")
        print("  不装也能用：排版会改为输出 WPS/Word 手动设置清单。")
    else:
        print("  依赖齐全，可以直接使用排版功能。")


def print_font_check(root: Path, wanted: List[str]) -> None:
    """安装了排版 skill 的，调用其字体检查，打印本机缺哪些标准字体以及获取、安装方法。

    只读检查，不联网、不改动任何文件；检查失败只提示，不影响安装结果。
    """
    script = root / FONT_CHECK_REL
    if FORMAT_SKILL not in wanted or not script.is_file():
        return
    print()
    print("排版字体检查（公文排版要求标准字体，缺字体时不会改用其他字体顶替，须自行下载安装）：")
    try:
        result = subprocess.run(
            [sys.executable, "-X", "utf8", str(script)],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=60,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        print(f"  字体检查没能运行（{exc}），不影响安装；之后可在排版 skill 目录运行 python scripts/gw_fonts.py 自行检查。")
        return
    output = (result.stdout or "").strip()
    if not output:
        print("  字体检查没有输出，不影响安装；之后可在排版 skill 目录运行 python scripts/gw_fonts.py 自行检查。")
        return
    print(output)


def print_targets() -> None:
    print("支持的安装目标：")
    print(f"  {'target':<10} {'user 级':<22} {'project 级':<20} 状态      说明")
    for t in TARGETS:
        if t.user_verified and t.project_verified:
            status = "已核实"
        elif t.project_verified:
            status = "user 级未核实"
        else:
            status = "未核实"
        print(f"  {t.key:<10} {t.user_path:<22} {t.project_path:<20} {status:<8} {t.label}")
    print("说明：codex、agents 与 kimi（project 级）共用 .agents/skills，同时选择时只复制一次。")
    print("另可用 --target auto：安装到 --detect 探测到的全部工具。")


def _unique_backup_path(base: Path, name: str, stamp: str) -> Path:
    candidate = base / f"{name}.bak-{stamp}"
    counter = 1
    while candidate.exists() or candidate.is_symlink():
        candidate = base / f"{name}.bak-{stamp}-{counter}"
        counter += 1
    return candidate


def install_one(src: Path, base: Path, name: str, force: bool, dry_run: bool, stamp: str) -> Tuple[str, str]:
    """安装单个 skill，返回 (描述, 结果)。结果为 installed / skipped / backed_up / failed。"""
    dest = base / name
    prefix = "[dry-run] " if dry_run else ""
    if dest.exists() or dest.is_symlink():
        if not force:
            return f"  {prefix}跳过  {name}（已存在，加 --force 可备份后覆盖）", "skipped"
        backup = _unique_backup_path(base, name, stamp)
        if dry_run:
            return f"  {prefix}备份  {name} → {backup.name}，再复制新版本", "backed_up"
        try:
            dest.rename(backup)
        except OSError as exc:
            return f"  失败  {name}：备份旧目录失败（{exc}）", "failed"
        try:
            shutil.copytree(src, dest, ignore=COPY_IGNORE)
        except OSError as exc:
            return f"  失败  {name}：复制失败（{exc}），旧版本已备份在 {backup.name}", "failed"
        return f"  覆盖  {name}（旧版本已备份为 {backup.name}）", "backed_up"

    if dry_run:
        return f"  {prefix}安装  {name}", "installed"
    try:
        base.mkdir(parents=True, exist_ok=True)
        shutil.copytree(src, dest, ignore=COPY_IGNORE)
    except OSError as exc:
        return f"  失败  {name}：复制失败（{exc}）", "failed"
    return f"  安装  {name}", "installed"


def main(
    argv: Optional[Sequence[str]] = None,
    repo_root: Optional[Path] = None,
    home: Optional[Path] = None,
    cwd: Optional[Path] = None,
    stamp: Optional[str] = None,
    which: Optional[WhichFunc] = None,
) -> int:
    setup_console()
    args = build_parser().parse_args(argv)
    home = Path(home) if home else Path.home()
    cwd = Path(cwd) if cwd else Path.cwd()

    if args.list_targets:
        print_targets()
        return 0
    if args.detect:
        print_detection(detect_targets(home, which), args.scope)
        return 0

    targets = _split_csv(args.target)
    if not targets:
        print("错误：请用 --target 指定安装目标（或 --target auto 自动探测）；可先运行 --list-targets 查看支持的目标。")
        return 2
    unknown = [t for t in targets if t != AUTO_TARGET and t not in TARGET_MAP]
    if unknown:
        print(f"错误：不支持的目标 {', '.join(unknown)}。可选：{AUTO_TARGET}, {', '.join(TARGET_MAP)}")
        return 2

    root = (args.root or repo_root or REPO_ROOT).resolve()
    names: List[str] = []
    wanted: List[str] = []
    absent: List[str] = []
    try:
        catalog = load_catalog(root)
        packs = resolve_packs(args.pack, catalog)
        if args.uninstall:
            names = resolve_uninstall_skills(args.skills, packs, catalog)
        else:
            available = list_installable_skills(root)
            wanted, absent = resolve_install_skills(args.skills, packs, catalog, available)
    except UsageError as exc:
        print(f"错误：{exc}")
        return 2

    if args.dry_run:
        print("[dry-run] 只打印操作，不做任何改动。")
    if AUTO_TARGET in targets:
        found = detect_targets(home, which)
        explicit = [t for t in targets if t != AUTO_TARGET]
        targets = _dedup(explicit + list(found))
        if found:
            shown = "、".join(f"{k}（{TARGET_MAP[k].label}）" for k in found)
            print(f"自动探测到 {len(found)} 个工具：{shown}")
        elif not targets:
            print(NO_TOOL_HINT)
            return 3

    if args.uninstall:
        return run_uninstall(names, targets, args.scope, home, cwd, args.yes or args.force, args.dry_run)

    if absent:
        print(f"提醒：以下 skill 列在 {MARKETPLACE_REL} 中，但 skills/ 下还没有，已跳过：{', '.join(absent)}")
    if not wanted:
        print("提示：skills/ 下还没有可安装的 skill。")
        return 0

    stamp = stamp or time.strftime("%Y%m%d-%H%M%S")
    failures = 0
    backups = 0
    done: Dict[str, str] = {}  # 已处理的目标目录 -> 目标名，避免重复复制
    for key in targets:
        target = TARGET_MAP[key]
        base = resolve_base(target, args.scope, home, cwd)
        shown = target.path_for(args.scope)
        print(f"目标 {key}（{target.label}，{args.scope} 级）：{shown} → {base}")
        norm = os.path.normcase(os.path.abspath(base))
        if norm in done:
            print(f"  与目标 {done[norm]} 的目录相同，已在上一步处理，跳过。")
            continue
        done[norm] = key
        if not target.verified(args.scope):
            print(f"  提醒：{UNVERIFIED_NOTE}")
        for name in wanted:
            message, outcome = install_one(root / "skills" / name, base, name, args.force, args.dry_run, stamp)
            print(message)
            if outcome == "failed":
                failures += 1
            elif outcome == "backed_up":
                backups += 1

    if failures:
        print(f"完成，但有 {failures} 个 skill 安装失败。")
        return 1
    if backups and not args.dry_run:
        print("提示：备份目录与新版本位于同一 skills 目录，部分工具可能把备份识别为重复 skill，确认无误后请手动移走。")
    print("完成。" if not args.dry_run else "[dry-run] 结束，未做任何改动。")
    print_next_steps(targets, wanted, catalog, packs)
    if not args.dry_run:
        print_font_check(root, wanted)
    return 0


if __name__ == "__main__":
    sys.exit(main())
