#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按各 skill 的 manifest.json，把 shared/ 下的共享资料同步进 skill 目录。

shared/ 是唯一事实源。每个 skill 在 manifest.json 的 shared 列表中声明需要的文件：
  {"src": "guardrails.md"}                                   -> references/_shared/guardrails.md
  {"src": "scripts/gongwen_lint.py", "dest": "scripts/gongwen_lint.py"}

行为：
  1. 逐字节复制 shared/<src> 到 skill 内的 dest，自动创建目录，内容相同时不重写；
  2. 在 references/_shared/ 下生成 README.md，提示该目录禁止手改；
  3. 清理 references/_shared/ 下已不在清单中的旧文件（只清理该目录，不碰其他位置）。

退出码：0 成功或无差异；1 存在差异（--check）；2 配置错误（manifest 无效、src 不存在等）。

本模块同时提供 validate.py、build_prompts.py 复用的公共函数。
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Dict, List, Optional, Sequence, Tuple

# 仓库根目录：tools/ 的上一级
REPO_ROOT = Path(__file__).resolve().parent.parent

SHARED_DIR_REL = "references/_shared"
README_REL = SHARED_DIR_REL + "/README.md"
README_TEXT = (
    "# 自动生成目录\n"
    "\n"
    "本目录由 tools/sync_shared.py 自动生成，请勿手改，修改请改 shared/。\n"
)
# dest 不允许覆盖的 skill 关键文件
PROTECTED_DESTS = ("SKILL.md", "README.md", "manifest.json")


# ---------------------------------------------------------------------------
# 公共工具函数（validate.py、build_prompts.py 共用）
# ---------------------------------------------------------------------------

def setup_console() -> None:
    """Windows 控制台默认编码可能不是 UTF-8，输出中文前统一切换。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:  # 被重定向为 StringIO 等对象时没有 reconfigure
            pass


class ChineseHelpFormatter(argparse.RawDescriptionHelpFormatter):
    """把 argparse 的 usage 前缀改为中文。"""

    def add_usage(self, usage, actions, groups, prefix=None):
        if prefix is None:
            prefix = "用法："
        return super().add_usage(usage, actions, groups, prefix)


def make_parser(prog: str, description: str, epilog: Optional[str] = None) -> argparse.ArgumentParser:
    """创建中文帮助信息的参数解析器。"""
    parser = argparse.ArgumentParser(
        prog=prog,
        description=description,
        epilog=epilog,
        formatter_class=ChineseHelpFormatter,
        add_help=False,
    )
    parser._optionals.title = "选项"  # argparse 未提供公开接口修改分组标题
    parser.add_argument("-h", "--help", action="help", help="显示本帮助信息并退出")
    return parser


def add_root_argument(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--root",
        type=Path,
        metavar="DIR",
        help="仓库根目录（默认为本脚本所在仓库，主要供测试使用）",
    )


def list_skill_dirs(repo_root: Path) -> List[Path]:
    """列出 skills/ 下的 skill 目录（按名称排序，忽略以 . 或 _ 开头的目录）。"""
    skills_dir = repo_root / "skills"
    if not skills_dir.is_dir():
        return []
    dirs = [
        p for p in skills_dir.iterdir()
        if p.is_dir() and not p.name.startswith((".", "_"))
    ]
    return sorted(dirs, key=lambda p: p.name)


def load_json_file(path: Path) -> Tuple[object, Optional[str]]:
    """读取 JSON 文件，返回 (数据, 错误信息)。"""
    try:
        text = path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return None, "文件不存在"
    except UnicodeDecodeError as exc:
        return None, f"不是合法的 UTF-8 编码（{exc.reason}）"
    except OSError as exc:
        return None, f"读取失败（{exc}）"
    try:
        return json.loads(text), None
    except json.JSONDecodeError as exc:
        return None, f"JSON 解析失败（第 {exc.lineno} 行第 {exc.colno} 列：{exc.msg}）"


def normalize_rel_path(value: object) -> Tuple[Optional[str], Optional[str]]:
    """校验并规范化相对路径（POSIX 风格），返回 (规范路径, 错误信息)。"""
    if not isinstance(value, str) or not value.strip():
        return None, "路径必须是非空字符串"
    if "\\" in value:
        return None, f"路径请使用正斜杠 /：{value}"
    pure = PurePosixPath(value)
    if pure.is_absolute() or (pure.parts and ":" in pure.parts[0]):
        return None, f"路径必须是相对路径：{value}"
    parts = [p for p in pure.parts if p != "."]
    if not parts:
        return None, f"路径无效：{value}"
    if ".." in parts:
        return None, f"路径不能包含 ..：{value}"
    return "/".join(parts), None


def exists_exact(base: Path, rel: str, want_dir: Optional[bool] = None) -> bool:
    """大小写敏感地判断 base/rel 是否存在。

    Windows、macOS 的文件系统默认不区分大小写，而 CI 所用的 Linux 区分；
    逐级比对目录项可以在本地提前发现大小写不一致的问题。
    want_dir：True 要求是目录，False 要求是文件，None 不限。
    """
    current = base
    for part in PurePosixPath(rel).parts:
        try:
            names = os.listdir(current)
        except OSError:
            return False
        if part not in names:
            return False
        current = current / part
    if want_dir is True:
        return current.is_dir()
    if want_dir is False:
        return current.is_file()
    return current.exists()


def iter_files(base: Path) -> List[str]:
    """递归列出 base 下的文件（相对路径，POSIX 风格，排序），不跟随目录符号链接。"""
    result: List[str] = []
    if not base.is_dir():
        return result
    for dirpath, dirnames, filenames in os.walk(base):
        dirnames.sort()
        rel_dir = Path(dirpath).relative_to(base)
        for name in sorted(filenames):
            result.append((rel_dir / name).as_posix())
    return sorted(result)


def same_text_ignoring_eol(data: bytes, text: str) -> bool:
    """比较生成文本时忽略换行符差异（兼容 Windows 下 git 的 autocrlf）。"""
    return data.replace(b"\r\n", b"\n") == text.encode("utf-8")


# ---------------------------------------------------------------------------
# 同步计划与比较
# ---------------------------------------------------------------------------

@dataclass
class SyncEntry:
    src_rel: str      # 相对 shared/
    dest_rel: str     # 相对 skill 根目录
    src_path: Path
    dest_path: Path


@dataclass
class SyncPlan:
    name: str
    skill_dir: Path
    entries: List[SyncEntry] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    @property
    def shared_dir(self) -> Path:
        return self.skill_dir / SHARED_DIR_REL

    def src_set(self) -> List[str]:
        return [e.src_rel for e in self.entries]


@dataclass
class Diff:
    kind: str                  # 缺失 / 不一致 / 多余
    path: str                  # 相对 skill 根目录
    src: Optional[str] = None  # 相对 shared/

    def describe(self) -> str:
        if self.src:
            return f"{self.kind}  {self.path}  ← shared/{self.src}"
        return f"{self.kind}  {self.path}"


def build_plan(repo_root: Path, skill_dir: Path, manifest: object = None) -> SyncPlan:
    """根据 manifest.json 生成同步计划；manifest 为 None 时自动读取。"""
    plan = SyncPlan(name=skill_dir.name, skill_dir=skill_dir)
    if manifest is None:
        manifest, err = load_json_file(skill_dir / "manifest.json")
        if err:
            plan.errors.append(f"manifest.json：{err}")
            return plan
    if not isinstance(manifest, dict):
        plan.errors.append("manifest.json：顶层必须是 JSON 对象")
        return plan

    shared = manifest.get("shared", [])
    if not isinstance(shared, list):
        plan.errors.append("manifest.json：shared 必须是数组")
        return plan

    shared_root = repo_root / "shared"
    seen_dest: Dict[str, int] = {}
    for index, item in enumerate(shared, start=1):
        where = f"manifest.json 的 shared 第 {index} 项"
        if not isinstance(item, dict) or "src" not in item:
            plan.errors.append(f"{where}：必须是包含 src 字段的对象")
            continue
        src, err = normalize_rel_path(item.get("src"))
        if err:
            plan.errors.append(f"{where}：src {err}")
            continue
        dest_raw = item.get("dest")
        if dest_raw is None:
            dest = f"{SHARED_DIR_REL}/{src}"
        else:
            dest, err = normalize_rel_path(dest_raw)
            if err:
                plan.errors.append(f"{where}：dest {err}")
                continue
        # Windows、macOS 文件系统不区分大小写，skill.md 也会覆盖 SKILL.md，故不区分大小写比较
        if dest.lower() in {p.lower() for p in PROTECTED_DESTS + (README_REL,)}:
            plan.errors.append(f"{where}：dest 不能是 {dest}")
            continue
        if dest in seen_dest:
            plan.errors.append(f"{where}：dest {dest} 与第 {seen_dest[dest]} 项重复")
            continue
        seen_dest[dest] = index
        if not exists_exact(shared_root, src, want_dir=False):
            plan.errors.append(
                f"{where}：源文件 shared/{src} 不存在（注意大小写），"
                f"请检查 src 拼写，或先在 shared/ 中创建该文件"
            )
            continue
        plan.entries.append(SyncEntry(src, dest, shared_root / src, skill_dir / dest))
    return plan


def diff_plan(plan: SyncPlan) -> List[Diff]:
    """比较同步计划与 skill 目录现状，返回差异清单（不写入）。"""
    diffs: List[Diff] = []
    for entry in plan.entries:
        if not exists_exact(plan.skill_dir, entry.dest_rel, want_dir=False):
            diffs.append(Diff("缺失", entry.dest_rel, entry.src_rel))
        elif entry.dest_path.read_bytes() != entry.src_path.read_bytes():
            diffs.append(Diff("不一致", entry.dest_rel, entry.src_rel))

    readme = plan.skill_dir / README_REL
    if not exists_exact(plan.skill_dir, README_REL, want_dir=False):
        diffs.append(Diff("缺失", README_REL))
    elif not same_text_ignoring_eol(readme.read_bytes(), README_TEXT):
        diffs.append(Diff("不一致", README_REL))

    expected = {e.dest_rel for e in plan.entries}
    expected.add(README_REL)
    for rel in iter_files(plan.shared_dir):
        full = f"{SHARED_DIR_REL}/{rel}"
        if full not in expected:
            diffs.append(Diff("多余", full))
    return sorted(diffs, key=lambda d: (d.path, d.kind))


def _is_inside(path: Path, base: Path) -> bool:
    """不解析符号链接，判断 path 是否位于 base 之内。"""
    try:
        Path(os.path.abspath(path)).relative_to(Path(os.path.abspath(base)))
        return True
    except ValueError:
        return False


def _real(path: Path) -> str:
    """解析符号链接和 Windows 目录联接（junction）后的真实路径，已统一大小写。"""
    return os.path.normcase(os.path.realpath(path))


def _is_inside_real(path: Path, base: Path) -> bool:
    """解析链接后判断 path 是否仍位于 base 之内。

    Windows 目录联接不被 is_symlink() 识别，os.walk 也会进入其中，
    只按字面路径判断会删到联接目标里的文件。
    """
    real, real_base = _real(path), _real(base)
    return real == real_base or real.startswith(real_base.rstrip(os.sep) + os.sep)


def _write_bytes_if_changed(path: Path, data: bytes) -> bool:
    if path.is_file() and path.read_bytes() == data:
        return False
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    return True


def apply_plan(plan: SyncPlan) -> List[str]:
    """执行同步，返回动作描述列表。"""
    actions: List[str] = []
    shared_dir = plan.shared_dir
    for guard in (plan.skill_dir / "references", shared_dir):
        if guard.is_symlink():
            raise RuntimeError(f"{guard} 是符号链接，为避免误删源文件，拒绝处理")
    # 目录联接（junction）不是符号链接，再按真实路径核对一次
    expected = os.path.join(os.path.realpath(plan.skill_dir), *SHARED_DIR_REL.split("/"))
    if _real(shared_dir) != os.path.normcase(expected):
        raise RuntimeError(f"{shared_dir} 指向别处（符号链接或目录联接），为避免误删源文件，拒绝处理")

    # 1. 先删除 references/_shared/ 下的多余文件（大小写不敏感的文件系统上，
    #    改名后的旧文件必须先删，再写新文件）
    stale: List[Tuple[Diff, Path]] = []
    for diff in diff_plan(plan):
        if diff.kind != "多余":
            continue
        target = plan.skill_dir / diff.path
        if not _is_inside(target, shared_dir):  # 防御：绝不删除该目录以外的文件
            continue
        if not _is_inside_real(target.parent, shared_dir):  # 子目录是目录联接或符号链接
            raise RuntimeError(f"{diff.path} 位于指向别处的目录联接或符号链接中，为避免误删源文件，拒绝处理")
        stale.append((diff, target))
    for diff, target in stale:  # 全部核对通过后再删除
        target.unlink()
        actions.append(f"删除  {diff.path}（已不在清单中）")

    # 2. 写入缺失或不一致的文件
    for entry in plan.entries:
        data = entry.src_path.read_bytes()
        existed = exists_exact(plan.skill_dir, entry.dest_rel, want_dir=False)
        if _write_bytes_if_changed(entry.dest_path, data) or not existed:
            verb = "更新" if existed else "新增"
            actions.append(f"{verb}  {entry.dest_rel}  ← shared/{entry.src_rel}")

    readme = plan.skill_dir / README_REL
    if not readme.is_file() or not same_text_ignoring_eol(readme.read_bytes(), README_TEXT):
        _write_bytes_if_changed(readme, README_TEXT.encode("utf-8"))
        actions.append(f"生成  {README_REL}")

    # 3. 清理 references/_shared/ 下的空目录（不删除该目录本身）
    if shared_dir.is_dir():
        for dirpath, _dirnames, _filenames in os.walk(shared_dir, topdown=False):
            current = Path(dirpath)
            if current == shared_dir or not _is_inside(current, shared_dir):
                continue
            if not _is_inside_real(current, shared_dir):  # 目录联接本身及其内部：不删除
                continue
            try:
                if not any(current.iterdir()):
                    current.rmdir()
            except OSError:
                pass
    return actions


# ---------------------------------------------------------------------------
# 命令行入口
# ---------------------------------------------------------------------------

def main(argv: Optional[Sequence[str]] = None, repo_root: Optional[Path] = None) -> int:
    setup_console()
    parser = make_parser(
        "sync_shared.py",
        "按各 skill 的 manifest.json，把 shared/ 下的共享资料同步进 skill 目录。",
        "退出码：0 成功或无差异；1 存在差异（--check）；2 配置错误。",
    )
    parser.add_argument("--check", action="store_true", help="只比较不写入；有差异时列出并返回 1")
    parser.add_argument("--skill", metavar="NAME", help="只处理指定 skill")
    add_root_argument(parser)
    args = parser.parse_args(argv)

    root = (args.root or repo_root or REPO_ROOT).resolve()
    skill_dirs = list_skill_dirs(root)
    if args.skill:
        skill_dirs = [d for d in skill_dirs if d.name == args.skill]
        if not skill_dirs:
            print(f"错误：skills/ 下找不到名为 {args.skill} 的 skill")
            return 2
    if not skill_dirs:
        print("提示：skills/ 下还没有任何 skill，无需同步。")
        return 0

    error_count = 0
    diff_count = 0
    for skill_dir in skill_dirs:
        name = skill_dir.name
        if not (skill_dir / "manifest.json").is_file():
            print(f"[{name}] 警告：缺少 manifest.json，已跳过")
            continue
        plan = build_plan(root, skill_dir)
        if plan.errors:
            error_count += len(plan.errors)
            for err in plan.errors:
                print(f"[{name}] 错误：{err}")
            print(f"[{name}] 存在配置错误，本 skill 未做任何改动")
            continue

        if args.check:
            diffs = diff_plan(plan)
            if diffs:
                diff_count += len(diffs)
                print(f"[{name}] 发现 {len(diffs)} 处差异：")
                for diff in diffs:
                    print(f"  {diff.describe()}")
            else:
                print(f"[{name}] 一致")
        else:
            try:
                actions = apply_plan(plan)
            except (OSError, RuntimeError) as exc:
                error_count += 1
                print(f"[{name}] 错误：同步失败（{exc}）")
                continue
            if actions:
                print(f"[{name}] 已同步：")
                for action in actions:
                    print(f"  {action}")
            else:
                print(f"[{name}] 已是最新")

    if error_count:
        print(f"同步未完成：{error_count} 个配置错误。")
        return 2
    if args.check:
        if diff_count:
            print(f"共 {diff_count} 处差异。请运行 python tools/sync_shared.py 同步后再提交。")
            return 1
        print("检查通过：所有 skill 的共享资料与 shared/ 一致。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
