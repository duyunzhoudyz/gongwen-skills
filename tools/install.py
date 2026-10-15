#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 skills/ 下的 skill 复制到各 Agent 工具的 skill 目录。

示例：
  python tools/install.py --list-targets
  python tools/install.py --target qwen
  python tools/install.py --target claude,cursor --skills gongwen,gongwen-report
  python tools/install.py --target trae --scope project --dry-run

目标目录已存在同名 skill 时：默认跳过并提示；加 --force 先把旧目录改名为
<名>.bak-<时间戳> 备份，再复制新版本。本脚本不删除任何文件。
退出码：0 成功；1 部分 skill 安装失败；2 参数错误。
"""
from __future__ import annotations

import argparse
import os
import shutil
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

REPO_ROOT = Path(__file__).resolve().parent.parent
UNVERIFIED_NOTE = "该路径来自社区资料，如不生效请查阅该工具文档"
# .gongwen（单位画像）和 公文（用户产出）可能含内部信息，不复制到安装目录
COPY_IGNORE = shutil.ignore_patterns(
    "__pycache__", "*.pyc", "*.pyo", ".DS_Store", "Thumbs.db", ".gongwen", "公文",
)


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
        description="把 gongwen skill 复制到各 Agent 工具的 skill 目录。不删除任何文件。",
        epilog="退出码：0 成功；1 部分 skill 安装失败；2 参数错误。",
        formatter_class=ChineseHelpFormatter,
        add_help=False,
    )
    parser._optionals.title = "选项"
    parser.add_argument("-h", "--help", action="help", help="显示本帮助信息并退出")
    parser.add_argument("--list-targets", action="store_true", help="列出支持的目标及其路径")
    parser.add_argument(
        "--target",
        action="append",
        metavar="NAME",
        help="安装目标，可多选：--target claude --target qwen，或 --target claude,qwen",
    )
    parser.add_argument(
        "--scope",
        choices=("user", "project"),
        default="user",
        help="user 安装到用户目录（默认）；project 安装到当前目录下的对应目录",
    )
    parser.add_argument("--skills", metavar="A,B", help="只安装指定 skill，逗号分隔（默认全部）")
    parser.add_argument("--dry-run", action="store_true", help="只打印将要执行的操作，不做任何改动")
    parser.add_argument(
        "--force",
        action="store_true",
        help="目标已存在同名 skill 时，先改名备份为 <名>.bak-<时间戳> 再复制",
    )
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
) -> int:
    setup_console()
    args = build_parser().parse_args(argv)

    if args.list_targets:
        print_targets()
        return 0

    targets = _split_csv(args.target)
    if not targets:
        print("错误：请用 --target 指定安装目标；可先运行 --list-targets 查看支持的目标。")
        return 2
    unknown = [t for t in targets if t not in TARGET_MAP]
    if unknown:
        print(f"错误：不支持的目标 {', '.join(unknown)}。可选：{', '.join(TARGET_MAP)}")
        return 2

    root = (args.root or repo_root or REPO_ROOT).resolve()
    available = list_installable_skills(root)
    wanted = _split_csv([args.skills]) if args.skills else list(available)
    missing = [s for s in wanted if s not in available]
    if missing:
        print(f"错误：skills/ 下找不到 {', '.join(missing)}（或缺少 SKILL.md）")
        return 2
    if not wanted:
        print("提示：skills/ 下还没有可安装的 skill。")
        return 0

    home = Path(home) if home else Path.home()
    cwd = Path(cwd) if cwd else Path.cwd()
    stamp = stamp or time.strftime("%Y%m%d-%H%M%S")
    if args.dry_run:
        print("[dry-run] 只打印操作，不做任何改动。")

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
    return 0


if __name__ == "__main__":
    sys.exit(main())
