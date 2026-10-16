#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把每个 skill 编译成可直接粘贴到网页大模型（DeepSeek、豆包、Kimi 等）的"提示词版"。

产物：
  dist/prompts/<skill>.md   信息块 + SKILL.md 正文（去掉 frontmatter）+ 按 prompt_bundle 顺序的附录
  dist/prompts/README.md    索引表
  dist/offline/gongwen-skills-offline-<版本>.zip（--offline 时生成）

输出是确定性的：同样的输入，生成的文件逐字节相同（统一 LF 换行，不含时间戳）。
退出码：0 成功；1 存在错误，或 --check 发现不一致。
"""
from __future__ import annotations

import os
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Dict, List, Optional, Sequence, Set, Tuple

_TOOLS_DIR = str(Path(__file__).resolve().parent)
if _TOOLS_DIR not in sys.path:
    sys.path.insert(0, _TOOLS_DIR)

from sync_shared import (  # noqa: E402
    REPO_ROOT,
    add_root_argument,
    exists_exact,
    list_skill_dirs,
    load_json_file,
    make_parser,
    normalize_rel_path,
    same_text_ignoring_eol,
    setup_console,
)
from validate import split_frontmatter  # noqa: E402

PROMPTS_REL = "dist/prompts"
OFFLINE_REL = "dist/offline"
INDEX_NAME = "README.md"
LONG_PROMPT_CHARS = 30000
STATUS_LABELS = {
    "stable": "Stable（稳定版）",
    "beta": "Beta（测试版）",
    "draft": "Draft（草稿）",
}
CN_DIGITS = "零一二三四五六七八九"
FENCE_RE = re.compile(r"^ {0,3}(`{3,}|~{3,})")
CODE_LANGS = {".json": "json", ".py": "python", ".yaml": "yaml", ".yml": "yaml", ".csv": "csv"}

# 离线包中跳过的缓存和系统文件；.gongwen（单位画像）和 公文（用户产出）可能含内部信息，也不得打包
JUNK_DIRS = {
    "__pycache__", ".git", ".pytest_cache", ".mypy_cache", ".tox",
    ".venv", "venv", "node_modules", ".gongwen", "公文",
}
JUNK_FILES = {".DS_Store", "Thumbs.db"}
JUNK_SUFFIXES = {".pyc", ".pyo"}
ZIP_DATE = (1980, 1, 1, 0, 0, 0)


# ---------------------------------------------------------------------------
# 文本处理
# ---------------------------------------------------------------------------

def normalize_newlines(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


def cn_number(n: int) -> str:
    """1 -> 一，12 -> 十二，35 -> 三十五（附录编号用）。"""
    if n <= 0 or n >= 100:
        return str(n)
    if n < 10:
        return CN_DIGITS[n]
    tens, ones = divmod(n, 10)
    head = "十" if tens == 1 else CN_DIGITS[tens] + "十"
    return head + (CN_DIGITS[ones] if ones else "")


def find_first_h1(text: str) -> Tuple[Optional[str], Optional[int]]:
    """返回第一个一级标题（跳过代码块）及其行号（从 0 开始）。"""
    fence: Optional[Tuple[str, int]] = None
    for index, line in enumerate(text.split("\n")):
        match = FENCE_RE.match(line)
        if match:
            marker = match.group(1)
            if fence is None:
                fence = (marker[0], len(marker))
            elif marker[0] == fence[0] and len(marker) >= fence[1]:
                fence = None
            continue
        if fence is None and line.startswith("# "):
            title = re.sub(r"\s+#+\s*$", "", line[2:]).strip()
            if title:
                return title, index
    return None, None


def _is_cjk(ch: str) -> bool:
    """中日韩文字或全角标点。"""
    return bool(ch) and ("\u3000" <= ch <= "\u9fff" or "\uff00" <= ch <= "\uffef")


def replace_bundle_paths(text: str, mapping: Dict[str, str]) -> str:
    """把已打包文件的路径替换为"（见附录X）"。

    路径可带反引号；与中文相邻时去掉多余空格，如"按 `references/x.md` 第一条"->"按（见附录一）第一条"。
    """
    if not mapping:
        return text
    alternatives = "|".join(re.escape(p) for p in sorted(mapping, key=len, reverse=True))
    pattern = re.compile(
        r"([ \t]?)(`?)(?<![A-Za-z0-9_./\\-])(" + alternatives + r")(?!\.?[A-Za-z0-9_/\\-])\2([ \t]?)"
    )

    def repl(match: "re.Match[str]") -> str:
        lead, trail = match.group(1), match.group(4)
        prev_ch = text[match.start() - 1] if match.start() > 0 else ""
        next_ch = text[match.end()] if match.end() < len(text) else ""
        if lead and _is_cjk(prev_ch):
            lead = ""
        if trail and _is_cjk(next_ch):
            trail = ""
        return lead + mapping[match.group(3)] + trail

    return pattern.sub(repl, text)


def short_title(title: str, name: str) -> str:
    """"综合文稿（gongwen-report）" -> "综合文稿"。"""
    match = re.match(r"^(.*?)\s*[（(]\s*" + re.escape(name) + r"\s*[)）]\s*$", title)
    if match and match.group(1).strip():
        return match.group(1).strip()
    return title


def first_sentence(text: str) -> str:
    match = re.search(r"[。！？!?]", text)
    return text[:match.end()] if match else text


def _escape_cell(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", " ")


# ---------------------------------------------------------------------------
# 单个 skill 的渲染
# ---------------------------------------------------------------------------

@dataclass
class PromptDoc:
    name: str
    title: str
    short_title: str
    summary: str
    text: str
    char_count: int
    is_long: bool


@dataclass
class RenderResult:
    doc: Optional[PromptDoc] = None
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


def render_appendix(index: int, rel: str, content: str, mapping: Dict[str, str]) -> str:
    pure = PurePosixPath(rel)
    suffix = pure.suffix.lower()
    label = f"附录{cn_number(index)}"
    if suffix in (".md", ".markdown"):
        if content.startswith("---\n"):
            _fields, body, _errors, has_fm = split_frontmatter(content)
            if has_fm:
                content = body
        title, line_index = find_first_h1(content)
        lines = content.split("\n")
        first_text = next((i for i, line in enumerate(lines) if line.strip()), None)
        if title is not None and line_index == first_text:
            del lines[line_index]  # 文件标题已用作附录小标题
        body = replace_bundle_paths("\n".join(lines).strip("\n"), mapping)
        heading = title or pure.name
    elif suffix == ".txt":
        body = replace_bundle_paths(content.strip("\n"), mapping)
        heading = pure.name
    else:
        longest = max((len(m.group(0)) for m in re.finditer(r"`+", content)), default=0)
        fence = "`" * max(3, longest + 1)
        body = f"{fence}{CODE_LANGS.get(suffix, '')}\n{content.rstrip(chr(10))}\n{fence}"
        heading = pure.name
    return f"# {label}　{heading}\n\n{body}\n"


def render_header(name: str, title: str, version: str, status: str, char_count: int, is_long: bool) -> str:
    lines = [
        f"> **{title}**（提示词版）",
        ">",
        f"> 版本：{version}　成熟度：{status}　全文约 {char_count} 字符",
        ">",
        "> 使用方法：",
        "> 1. 复制全文，粘贴到对话开头；",
        "> 2. 再说出你的需求；",
        "> 3. 网页版无法运行脚本，自检清单需要人工逐项执行；",
        "> 4. 请勿输入涉密和内部敏感信息。",
    ]
    if is_long:
        lines += [
            ">",
            f"> 注意：全文超过 {LONG_PROMPT_CHARS} 字符，部分网页大模型可能截断或遗忘后文，建议选用支持长文本的模型。",
        ]
    lines += [">", f"> 本文件由 tools/build_prompts.py 根据 skills/{name} 自动生成，请勿手改。"]
    return "\n".join(lines) + "\n\n---\n\n"


HANDOFF_SOURCE = "references/_shared/style/output-format.md"
HANDOFF_HEADING = "## 四、交接卡与转交"


def extract_handoff(text: str) -> str:
    """从 output-format.md 中摘出第四部分"交接卡与转交"，改写为可作附录的独立小节；找不到时返回空串。"""
    lines = text.split("\n")
    start = next((i for i, line in enumerate(lines) if line.strip().startswith(HANDOFF_HEADING)), None)
    if start is None:
        return ""
    end = next((i for i in range(start + 1, len(lines)) if lines[i].startswith("## ")), len(lines))
    section = lines[start + 1:end]
    # 三级标题提为二级，与其他附录的层级一致
    section = ["## " + line[4:] if line.startswith("### ") else line for line in section]
    body = "\n".join(section).strip("\n")
    return "# 交接卡与转交（摘自输出格式规范第四部分）\n\n" + body + "\n"


def render_skill(skill_dir: Path) -> RenderResult:
    result = RenderResult()
    name = skill_dir.name
    if not exists_exact(skill_dir, "SKILL.md", want_dir=False):
        result.errors.append("缺少 SKILL.md")
        return result
    try:
        skill_text = normalize_newlines((skill_dir / "SKILL.md").read_text(encoding="utf-8"))
    except UnicodeDecodeError:
        result.errors.append("SKILL.md 不是合法的 UTF-8 编码")
        return result
    fields, body, _fm_errors, has_fm = split_frontmatter(skill_text)
    if not has_fm:
        result.warnings.append("SKILL.md 缺少 frontmatter，已按全文处理")
        body = skill_text

    manifest, err = load_json_file(skill_dir / "manifest.json")
    if err:
        result.errors.append(f"manifest.json：{err}")
        return result
    if not isinstance(manifest, dict):
        result.errors.append("manifest.json：顶层必须是 JSON 对象")
        return result
    version = manifest.get("version") if isinstance(manifest.get("version"), str) else "未标注"
    status = STATUS_LABELS.get(manifest.get("status"), "未标注")
    bundle = manifest.get("prompt_bundle", [])
    if not isinstance(bundle, list):
        result.errors.append("manifest.json 的 prompt_bundle 必须是数组")
        return result

    mapping: Dict[str, str] = {}
    files: List[Tuple[str, str]] = []
    for index, item in enumerate(bundle, start=1):
        rel, err = normalize_rel_path(item)
        if err:
            result.errors.append(f"prompt_bundle 第 {index} 项：{err}")
            continue
        if rel in mapping:
            result.warnings.append(f"prompt_bundle 重复列出 {rel}，已忽略重复项")
            continue
        if not exists_exact(skill_dir, rel, want_dir=False):
            result.errors.append(f"prompt_bundle 引用的文件不存在（注意大小写）：{rel}")
            continue
        try:
            # utf-8-sig：去掉记事本等工具写入的 BOM，避免混入提示词、影响标题识别
            content = normalize_newlines((skill_dir / rel).read_text(encoding="utf-8-sig"))
        except UnicodeDecodeError:
            result.errors.append(f"prompt_bundle 文件不是 UTF-8 文本：{rel}")
            continue
        mapping[rel] = f"（见附录{cn_number(len(files) + 1)}）"
        files.append((rel, content))
    if result.errors:
        return result

    # 网页版跨 skill 转交需要交接卡和安装命令（开发规范 D14）。没有整份打包 output-format.md 的，
    # 自动附上它的第四部分"交接卡与转交"，避免各 skill 为此超出字数目标。
    if HANDOFF_SOURCE not in mapping and exists_exact(skill_dir, HANDOFF_SOURCE, want_dir=False):
        handoff = extract_handoff(normalize_newlines((skill_dir / HANDOFF_SOURCE).read_text(encoding="utf-8-sig")))
        if handoff:
            files.append((HANDOFF_SOURCE[:-3] + "-第四部分.md", handoff))

    title = find_first_h1(body)[0] or name
    content = replace_bundle_paths(body.strip("\n"), mapping) + "\n"
    for index, (rel, text) in enumerate(files, start=1):
        content += "\n---\n\n" + render_appendix(index, rel, text, mapping)

    # 信息块中的字数包含信息块本身，迭代到位数稳定为止
    is_long = len(render_header(name, title, version, status, len(content), False)) + len(content) > LONG_PROMPT_CHARS
    count = len(content)
    full = ""
    for _ in range(10):
        full = render_header(name, title, version, status, count, is_long) + content
        if len(full) == count:
            break
        count = len(full)
    if is_long:
        result.warnings.append(f"提示词版共 {count} 字符，超过 {LONG_PROMPT_CHARS}，部分网页大模型可能截断")

    description = fields.get("description", "") if has_fm else ""
    result.doc = PromptDoc(
        name=name,
        title=title,
        short_title=short_title(title, name),
        summary=first_sentence(description.strip()),
        text=full,
        char_count=count,
        is_long=is_long,
    )
    return result


def render_index(docs: List[PromptDoc]) -> str:
    lines = [
        "# 提示词版索引",
        "",
        "本目录由 tools/build_prompts.py 自动生成，请勿手改。"
        "每个文件是一份可直接粘贴到网页大模型（DeepSeek、豆包、Kimi 等）的提示词。",
        "",
        "使用方法：打开对应文件，复制全文粘贴到对话开头，再说出你的需求。"
        "网页版无法运行脚本，自检清单需要人工逐项执行。请勿输入涉密和内部敏感信息。",
        "",
        "| 名称 | 中文名 | 说明 | 字数 | 文件 |",
        "|---|---|---|---:|---|",
    ]
    for doc in docs:
        lines.append(
            f"| {doc.name} | {_escape_cell(doc.short_title)} | {_escape_cell(doc.summary)} "
            f"| {doc.char_count} | [{doc.name}.md]({doc.name}.md) |"
        )
    return "\n".join(lines) + "\n"


# ---------------------------------------------------------------------------
# 整体构建、比较与写入
# ---------------------------------------------------------------------------

@dataclass
class BuildResult:
    outputs: Dict[str, str] = field(default_factory=dict)  # 文件名 -> 内容
    keep: Set[str] = field(default_factory=set)            # 出错未重建、需保留的旧文件
    docs: List[PromptDoc] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    skill_count: int = 0


def build_all(repo_root: Path) -> BuildResult:
    build = BuildResult()
    skill_dirs = list_skill_dirs(repo_root)
    build.skill_count = len(skill_dirs)
    for skill_dir in skill_dirs:
        name = skill_dir.name
        result = render_skill(skill_dir)
        build.warnings.extend(f"[{name}] {w}" for w in result.warnings)
        if result.doc is None:
            build.errors.extend(f"[{name}] {e}" for e in result.errors)
            build.keep.add(f"{name}.md")
            continue
        build.docs.append(result.doc)
        build.outputs[f"{name}.md"] = result.doc.text
    if build.docs:
        build.outputs[INDEX_NAME] = render_index(build.docs)
    if build.errors:
        build.keep.add(INDEX_NAME)
    return build


def _existing_outputs(out_dir: Path) -> List[str]:
    if not out_dir.is_dir():
        return []
    return sorted(p.name for p in out_dir.iterdir() if p.is_file() and p.suffix == ".md")


def check_outputs(repo_root: Path, build: BuildResult) -> List[str]:
    """比较 dist/prompts 与当前源文件，返回差异描述（不写入）。"""
    out_dir = repo_root / PROMPTS_REL
    diffs: List[str] = []
    for fname in sorted(build.outputs):
        path = out_dir / fname
        if not exists_exact(out_dir, fname, want_dir=False):
            diffs.append(f"缺失  {PROMPTS_REL}/{fname}")
        elif not same_text_ignoring_eol(path.read_bytes(), build.outputs[fname]):
            diffs.append(f"不一致  {PROMPTS_REL}/{fname}")
    for fname in _existing_outputs(out_dir):
        if fname not in build.outputs and fname not in build.keep:
            diffs.append(f"多余  {PROMPTS_REL}/{fname}")
    return diffs


def write_outputs(repo_root: Path, build: BuildResult) -> List[str]:
    out_dir = repo_root / PROMPTS_REL
    actions: List[str] = []
    for fname in _existing_outputs(out_dir):  # 只删除 dist/prompts 下的旧 .md
        if fname not in build.outputs and fname not in build.keep:
            (out_dir / fname).unlink()
            actions.append(f"删除  {PROMPTS_REL}/{fname}")
    for fname in sorted(build.outputs):
        path = out_dir / fname
        text = build.outputs[fname]
        if path.is_file() and same_text_ignoring_eol(path.read_bytes(), text):
            continue
        out_dir.mkdir(parents=True, exist_ok=True)
        path.write_bytes(text.encode("utf-8"))
        actions.append(f"写入  {PROMPTS_REL}/{fname}")
    return actions


# ---------------------------------------------------------------------------
# 离线包
# ---------------------------------------------------------------------------

def offline_version(repo_root: Path) -> Tuple[str, Optional[str]]:
    """从 marketplace.json 的 metadata.version 取版本号，返回 (版本, 警告)。"""
    path = repo_root / ".claude-plugin" / "marketplace.json"
    if not path.is_file():
        return "0.0.0", "未找到 .claude-plugin/marketplace.json，离线包版本号使用 0.0.0"
    data, err = load_json_file(path)
    if err:
        return "0.0.0", f"marketplace.json {err}，离线包版本号使用 0.0.0"
    metadata = data.get("metadata") if isinstance(data, dict) else None
    version = metadata.get("version") if isinstance(metadata, dict) else None
    if not isinstance(version, str) or not re.fullmatch(r"[0-9A-Za-z.+-]+", version):
        return "0.0.0", "marketplace.json 未设置有效的 metadata.version，离线包版本号使用 0.0.0"
    return version, None


def collect_offline_files(repo_root: Path) -> List[Tuple[str, Path]]:
    files: List[Tuple[str, Path]] = []
    for top in ("skills", PROMPTS_REL):
        base = repo_root / top
        if not base.is_dir():
            continue
        for dirpath, dirnames, filenames in os.walk(base):
            dirnames[:] = sorted(d for d in dirnames if d not in JUNK_DIRS)
            current = Path(dirpath)
            for name in sorted(filenames):
                if name in JUNK_FILES or Path(name).suffix in JUNK_SUFFIXES:
                    continue
                path = current / name
                files.append((f"{top}/{path.relative_to(base).as_posix()}", path))
    for name in ("README.md", "LICENSE", "使用须知.md", "THIRD_PARTY_NOTICES.md", "CHANGELOG.md"):
        path = repo_root / name
        if path.is_file():
            files.append((name, path))
    return sorted(files, key=lambda item: item[0])


def write_offline_zip(repo_root: Path, version: str) -> Tuple[Path, int]:
    """生成离线包；固定时间戳与权限，保证同样输入得到同样内容。"""
    out_dir = repo_root / OFFLINE_REL
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / f"gongwen-skills-offline-{version}.zip"
    tmp = target.with_name(target.name + ".tmp")
    files = collect_offline_files(repo_root)
    with zipfile.ZipFile(tmp, "w") as zf:
        for arcname, path in files:
            info = zipfile.ZipInfo(arcname, date_time=ZIP_DATE)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = 0o644 << 16
            zf.writestr(info, path.read_bytes(), compresslevel=9)
    os.replace(tmp, target)
    return target, len(files)


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

def main(argv: Optional[Sequence[str]] = None, repo_root: Optional[Path] = None) -> int:
    setup_console()
    parser = make_parser(
        "build_prompts.py",
        "为每个 skill 生成给网页大模型使用的提示词版（dist/prompts/），可选生成内网离线包。",
        "退出码：0 成功；1 存在错误，或 --check 发现不一致。",
    )
    parser.add_argument("--check", action="store_true", help="只比较不写入；dist/prompts 与源文件不一致时返回 1（供 CI 使用）")
    parser.add_argument(
        "--offline",
        action="store_true",
        help="额外生成 dist/offline/gongwen-skills-offline-<版本>.zip（含 skills/、dist/prompts/、README.md、LICENSE、使用须知.md、THIRD_PARTY_NOTICES.md、CHANGELOG.md）",
    )
    add_root_argument(parser)
    args = parser.parse_args(argv)
    root = (args.root or repo_root or REPO_ROOT).resolve()

    build = build_all(root)
    if build.skill_count == 0:
        print("提示：skills/ 下还没有任何 skill，未生成提示词版。")
    for warning in build.warnings:
        print(f"警告：{warning}")
    for error in build.errors:
        print(f"错误：{error}")

    if args.check:
        diffs = check_outputs(root, build)
        for diff in diffs:
            print(f"  {diff}")
        if args.offline:
            print("提示：--check 模式下不生成离线包。")
        if diffs:
            print(f"dist/prompts 与源文件不一致（{len(diffs)} 处）。请运行 python tools/build_prompts.py 后提交。")
            return 1
        if build.errors:
            return 1
        print(f"检查通过：{len(build.docs)} 份提示词版与源文件一致。")
        return 0

    try:
        actions = write_outputs(root, build)
    except OSError as exc:
        print(f"错误：写入 {PROMPTS_REL} 失败（{exc}）")
        return 1
    for action in actions:
        print(f"  {action}")
    for doc in build.docs:
        flag = "（超长）" if doc.is_long else ""
        print(f"[{doc.name}] {doc.char_count} 字符{flag}")
    if build.docs:
        print(f"已生成 {len(build.docs)} 份提示词版：{PROMPTS_REL}/")

    if args.offline:
        if build.errors:
            print("存在错误，未生成离线包。")
        else:
            version, warning = offline_version(root)
            if warning:
                print(f"警告：{warning}")
            try:
                target, count = write_offline_zip(root, version)
            except OSError as exc:
                print(f"错误：生成离线包失败（{exc}）")
                return 1
            print(f"已生成离线包：{target.relative_to(root).as_posix()}（{count} 个文件）")
    return 1 if build.errors else 0


if __name__ == "__main__":
    sys.exit(main())
