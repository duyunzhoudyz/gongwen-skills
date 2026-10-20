# -*- coding: utf-8 -*-
"""公文标准字体的检测、名称比对与缺字体引导（只依赖标准库）。

字体严格模式：
- 文档中一律写入标准字体名（assets/gbt9704_params.json 的 fonts.standard），没有降级、替代的路径。
- 本机缺少任一标准字体时，只做两件事：明确指出缺哪些字体，引导用户自行下载安装（fonts.install_guide）。
  文件仍写入标准字体名，但报告必须提示“在缺字体的电脑上看到的不是正式效果”。
- 本工具不提供、不打包任何字体文件。
- 单位制度另有规定字体（如正文用“方正仿宋_GBK”）时，用 --unit-fonts 按角色指定，该角色写入的字体名改为单位指定的字体；
  严格模式不变：本机没装单位指定字体同样报“字体缺失”并引导安装，不回退到国标字体或其他字体。

检测方式：
- Windows：读取 C:\\Windows\\Fonts 和用户字体目录（%LOCALAPPDATA%\\Microsoft\\Windows\\Fonts）。
- macOS：读取 /Library/Fonts、~/Library/Fonts，以及系统字体目录。
- Linux：调用 fc-list。
字体文件只读取 name 表中的字体族名称，不复制、不修改任何字体文件。

命令行：python scripts/gw_fonts.py [--json] [--unit-fonts "仿宋=方正仿宋_GBK,小标宋=方正小标宋_GBK"]
  输出各标准字体是否已安装；缺失（或无法检测）时输出每款字体的获取途径和通用安装方法。
  退出码：0 标准字体齐全；3 缺字体或无法检测（preview_docx.py 拒绝渲染时也返回 3）；2 参数或参数文件错误。
"""
from __future__ import annotations

import argparse
import json
import os
import platform
import re
import struct
import subprocess
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

FONT_SUFFIXES = {".ttf", ".otf", ".ttc"}
_CACHE: Optional[Tuple[Set[str], str]] = None
PARAMS_PATH = Path(__file__).resolve().parent.parent / "assets" / "gbt9704_params.json"
EXIT_FONTS_INCOMPLETE = 3  # 缺标准字体或无法检测：gw_fonts.py 命令行与 preview_docx.py 共用

# 缺字体的后果和态度，报告、命令行、预览拒绝时共用同一段文字
CONSEQUENCE = ("文件中写入的仍是标准字体名，在装有这些字体的电脑上会正常显示；但在本机用 Word/WPS 打开时，"
               "缺失的字体会被自动替换成其他字体显示，看到的不是正式效果，不能据此核对版式或打印。")
NO_SUBSTITUTE = "本工具不改用、也不建议任何替代字体（仿宋、楷体、华文中宋、宋体加粗等都不是标准字体）。"

# 字体类别关键词（小写比较）。类别用于检查"是不是仿宋体、楷体、黑体、小标宋体"
CLASS_KEYWORDS = {
    "小标宋": ("小标宋", "xiaobiaosong", "fzxbs"),
    "仿宋": ("仿宋", "fangsong", "fzfs", "stfangsong"),
    "楷体": ("楷体", "楷", "kaiti", "stkaiti", "fzkai"),
    "黑体": ("黑体", "simhei", "heiti", "fzhei", "stheiti"),
    "宋体": ("宋体", "中宋", "simsun", "songti", "stsong", "stzhongsong", "nsimsun", "宋"),
}
# 同一款字体的不同名称（中文名、英文名），检测和比对时互相等同。键、值均为 norm() 后的写法。
# 只收录“确实是同一款字体”的别名；名称相近但不是同一款的（如“仿宋”与“仿宋_GB2312”、
# “方正小标宋_GBK”与“方正小标宋简体”）不是别名，不得互相代替。
ALIASES = {
    "仿宋_gb2312": ("fangsong_gb2312",),
    "楷体_gb2312": ("kaiti_gb2312",),
    "方正小标宋简体": ("fzxiaobiaosong-b05s",),
    "黑体": ("simhei",),
    "宋体": ("simsun",),
}


def norm(name: str) -> str:
    return (name or "").strip().lower().replace(" ", "")


def font_class(name: Optional[str]) -> str:
    """返回字体类别：小标宋、仿宋、楷体、黑体、宋体、其他。"""
    n = norm(name or "")
    if not n:
        return "其他"
    for cls in ("小标宋", "仿宋", "楷体", "黑体", "宋体"):
        if any(k in n for k in CLASS_KEYWORDS[cls]):
            return cls
    return "其他"


def _same_font_names(name: str) -> Set[str]:
    """与 name 是同一款字体的全部名称（含自身），均为 norm() 后的写法。"""
    key = norm(name)
    for std, aliases in ALIASES.items():
        group = {std, *aliases}
        if key in group:
            return group
    return {key} if key else set()


def same_font(a: Optional[str], b: Optional[str]) -> bool:
    """a、b 是否同一款字体：名称相同，或互为别名（如 仿宋_GB2312 与 FangSong_GB2312）。"""
    names = _same_font_names(a or "")
    return bool(names) and norm(b or "") in names


# ---------------------------------------------------------------------------
# 读取字体文件 name 表
# ---------------------------------------------------------------------------

def _decode_name(platform_id: int, encoding_id: int, raw: bytes) -> Optional[str]:
    try:
        if platform_id in (0, 3):
            return raw.decode("utf-16-be", errors="ignore")
        if platform_id == 1:
            if encoding_id == 25:  # Mac 简体中文
                return raw.decode("gb2312", errors="ignore")
            return raw.decode("mac_roman", errors="ignore")
    except LookupError:
        return None
    return None


def _names_from_sfnt(fh, base: int) -> List[str]:
    fh.seek(base)
    header = fh.read(12)
    if len(header) < 12:
        return []
    num_tables = struct.unpack(">H", header[4:6])[0]
    records = fh.read(16 * num_tables)
    name_off = None
    for i in range(num_tables):
        tag, _chk, off, _length = struct.unpack(">4sIII", records[16 * i:16 * i + 16])
        if tag == b"name":
            name_off = off
            break
    if name_off is None:
        return []
    fh.seek(name_off)
    head = fh.read(6)
    if len(head) < 6:
        return []
    _fmt, count, str_off = struct.unpack(">HHH", head)
    recs = fh.read(12 * count)
    names: List[str] = []
    for i in range(count):
        if 12 * i + 12 > len(recs):
            break
        pid, eid, _lid, nid, length, off = struct.unpack(">HHHHHH", recs[12 * i:12 * i + 12])
        if nid not in (1, 4, 16):
            continue
        fh.seek(name_off + str_off + off)
        value = _decode_name(pid, eid, fh.read(length))
        if value:
            value = value.replace("\x00", "").strip()
            if value and value not in names:
                names.append(value)
    return names


def read_font_names(path: Path) -> List[str]:
    """读取 .ttf/.otf/.ttc 中的字体族名称和全名，失败时返回空列表。"""
    try:
        with open(path, "rb") as fh:
            tag = fh.read(4)
            if tag == b"ttcf":
                fh.read(4)
                num = struct.unpack(">I", fh.read(4))[0]
                offsets = struct.unpack(f">{num}I", fh.read(4 * num))
                out: List[str] = []
                for off in offsets:
                    for name in _names_from_sfnt(fh, off):
                        if name not in out:
                            out.append(name)
                return out
            return _names_from_sfnt(fh, 0)
    except (OSError, struct.error, ValueError):
        return []


def font_dirs() -> List[Path]:
    system = platform.system()
    dirs: List[Path] = []
    if system == "Windows":
        windir = os.environ.get("WINDIR", r"C:\Windows")
        dirs.append(Path(windir) / "Fonts")
        local = os.environ.get("LOCALAPPDATA")
        if local:
            dirs.append(Path(local) / "Microsoft" / "Windows" / "Fonts")
    elif system == "Darwin":
        dirs += [Path("/Library/Fonts"), Path.home() / "Library" / "Fonts",
                 Path("/System/Library/Fonts"), Path("/System/Library/Fonts/Supplemental")]
    return dirs


def _scan_dirs(dirs: Iterable[Path]) -> Set[str]:
    found: Set[str] = set()
    for d in dirs:
        if not d.is_dir():
            continue
        try:
            entries = list(d.iterdir())
        except OSError:
            continue
        for f in entries:
            if f.suffix.lower() in FONT_SUFFIXES:
                found.add(norm(f.stem))
                for name in read_font_names(f):
                    found.add(norm(name))
    return found


def _scan_fc_list() -> Tuple[Set[str], bool]:
    try:
        out = subprocess.run(["fc-list", ":", "family"], capture_output=True, text=True,
                             encoding="utf-8", errors="ignore", timeout=30)
    except (OSError, subprocess.SubprocessError):
        return set(), False
    found: Set[str] = set()
    for line in out.stdout.splitlines():
        # fc-list 输出中族名之间用逗号分隔，族名内的 \ - : , 前加反斜杠转义（如 FZXiaoBiaoSong\-B05S）
        for part in re.split(r"(?<!\\),", line):
            part = re.sub(r"\\(.)", r"\1", part)
            if part.strip():
                found.add(norm(part))
    return found, True


def installed_fonts(refresh: bool = False) -> Tuple[Set[str], str]:
    """返回 (已安装字体名集合（规范化小写）, 检测方式说明)。检测失败时集合为空。"""
    global _CACHE
    if _CACHE is not None and not refresh:
        return _CACHE
    system = platform.system()
    if system in ("Windows", "Darwin"):
        dirs = font_dirs()
        names = _scan_dirs(dirs)
        method = "扫描字体目录：" + "、".join(str(d) for d in dirs if d.is_dir())
    else:
        names, ok = _scan_fc_list()
        method = "fc-list" if ok else "未找到 fc-list，无法检测本机字体"
    _CACHE = (names, method)
    return _CACHE


def is_installed(font: str, installed: Set[str]) -> bool:
    """font 这一款字体（含同款别名）是否已安装；名称相近的其他字体不算。"""
    return bool(_same_font_names(font) & installed)


# ---------------------------------------------------------------------------
# 参数与字体核对（严格模式：只检测、提示，不替代）
# ---------------------------------------------------------------------------

def load_params(path: Optional[Path] = None) -> dict:
    with open(path or PARAMS_PATH, encoding="utf-8") as fh:
        return json.load(fh)

UNIT_FONTS_HELP = ('按单位规定指定标准字体：角色=字体名，逗号分隔，如 "仿宋=方正仿宋_GBK,小标宋=方正小标宋_GBK"；'
                   "角色限 仿宋、楷体、黑体、小标宋、宋体。仅限本单位制度明文规定了字体时使用；没装国标字体想换个字体不属于单位规定")


def parse_unit_fonts(value, params: dict) -> Dict[str, str]:
    """解析“按单位规定指定标准字体”：字符串“仿宋=方正仿宋_GBK,小标宋=…”或 {角色: 字体名} 字典，返回 {角色: 字体名}。

    角色只允许 params 中 fonts.standard 的键；未知角色、空字体名、格式错误、重复角色一律抛 ValueError（中文说明）。
    value 为 None 或空字典返回空字典；给出了字符串参数却没有任何内容（如空字符串）则报错。
    """
    if value is None or (isinstance(value, dict) and not value):
        return {}
    roles = list(params["fonts"]["standard"])
    pairs: List[Tuple[str, str]] = []
    if isinstance(value, dict):
        pairs = [(str(k).strip(), "" if v is None else str(v).strip()) for k, v in value.items()]
    elif isinstance(value, str):
        for item in re.split(r"[,，;；、\n]+", value):
            if not item.strip():
                continue
            parts = re.split(r"[=＝:：]", item, maxsplit=1)
            if len(parts) != 2:
                raise ValueError(f"单位规定字体格式错误：“{item.strip()}”，应写作 角色=字体名，如 仿宋=方正仿宋_GBK")
            pairs.append((parts[0].strip(), parts[1].strip()))
    else:
        raise ValueError("单位规定字体应为“角色=字体名”的逗号分隔字符串，或 {角色: 字体名} 对象")
    if not pairs:
        raise ValueError("单位规定字体没有给出任何内容，应写作 仿宋=方正仿宋_GBK,小标宋=方正小标宋_GBK")
    out: Dict[str, str] = {}
    for role, font in pairs:
        if role not in roles:
            raise ValueError(f"单位规定字体中的角色“{role}”无法识别，只能是：{'、'.join(roles)}")
        if not font:
            raise ValueError(f"单位规定字体中，角色“{role}”的字体名为空")
        if role in out:
            raise ValueError(f"单位规定字体中，角色“{role}”重复出现")
        out[role] = font
    return out


def apply_unit_fonts(params: dict, unit_fonts: Optional[Dict[str, str]]) -> dict:
    """返回把单位规定字体写入 fonts.standard 的参数副本（不改动传入的 params）。

    副本的 fonts.unit_override 记录 {角色: {unit_font, standard_font}}，供字体核对、报告和安装引导使用。
    没有单位规定字体时原样返回 params。
    """
    if not unit_fonts:
        return params
    out = json.loads(json.dumps(params))
    override = {}
    for role, font in unit_fonts.items():
        std = params["fonts"]["standard"][role]
        out["fonts"]["standard"][role] = font
        override[role] = {"unit_font": font, "standard_font": std}
    out["fonts"]["unit_override"] = override
    return out


def unit_notes(report: dict) -> List[str]:
    """报告中“按单位规定使用 X（国标为 Y）”的说明行；没有单位规定字体时为空列表。"""
    return [f"{role}：按单位规定使用 {v['unit_font']}（国标为 {v['standard_font']}）"
            for role, v in (report.get("unit_fonts") or {}).items()]


def font_guide(params: dict, role: str) -> dict:
    """某个字体角色的获取与安装说明：usage（用途）、source（来源）、obtain（获取途径）、note（注意事项）。

    单位规定字体不是预置字体，给通用说明：向单位信息中心或办公室索取，核对字体名完全一致，按通用方法安装。
    """
    entry = params["fonts"].get("install_guide", {}).get("fonts", {}).get(role, {})
    guide = {k: v for k, v in entry.items() if k != "font"}
    over = params["fonts"].get("unit_override", {}).get(role)
    if over:
        guide.update({
            "source": f"本单位制度规定的字体（国标为“{over['standard_font']}”），本工具不预置其获取途径。",
            "obtain": f"请向本单位信息中心或办公室索取“{over['unit_font']}”字体文件（须为单位有权使用的正版授权字体），再按下方通用方法安装。",
            "note": f"安装后在 Word/WPS 字体列表中核对：名称必须与“{over['unit_font']}”完全一致，名称相近的其他字体不能代替；"
                    "不要改用国标字体或其他字体。"})
    return guide


def resolve_fonts(params: dict, installed: Optional[Set[str]], method: str = "") -> Tuple[Dict[str, str], dict]:
    """核对本机标准字体，返回 (角色 -> 写入文档的字体名, 字体核对结果)。

    严格模式：写入文档的字体名一律是标准字体名，缺字体时也不降级、不替代，只在核对结果中指出缺哪些。
    角色为 params["fonts"]["standard"] 的键：仿宋、楷体、黑体、小标宋、宋体。
    installed 为空集合表示无法检测。核对结果的键与报告字段同名，可直接并入报告：
    - font_status：complete（齐全）、missing（有缺失）、undetected（无法检测，需人工核对）；
    - font_method：检测方式说明；
    - missing_fonts：缺失字体 [{font, role, guide}]，仅 missing 时非空；
    - unverified_fonts：无法确认的字体（全部标准字体，同样结构），仅 undetected 时非空；
    - font_notice：缺失或无法检测时的醒目提示（headline、summary、consequence、no_substitute、
      install_methods、statement），齐全时为 None。
    """
    standard: Dict[str, str] = params["fonts"]["standard"]
    fonts = dict(standard)
    entries = [{"font": std, "role": role, "guide": font_guide(params, role)} for role, std in standard.items()]
    detected = bool(installed)
    missing = [e for e in entries if detected and not is_installed(e["font"], installed or set())]
    status = "complete" if detected and not missing else "missing" if missing else "undetected"
    method = method or "未执行检测"
    notice = None
    if status != "complete":
        install = params["fonts"].get("install_guide", {})
        if status == "missing":
            headline = "字体缺失"
            summary = f"本机缺少 {len(missing)} 款标准字体：{'、'.join(e['font'] for e in missing)}。"
        else:
            headline = "无法检测本机字体，需人工核对"
            summary = (f"无法检测本机已安装的字体（{method}），不能确认是否装有以下标准字体："
                       f"{'、'.join(e['font'] for e in entries)}。请在打开文档的电脑上人工核对：字体列表中必须有名称完全一致的这些字体，"
                       "缺少任何一款，Word/WPS 都会自动替换显示。")
        notice = {"headline": headline, "summary": summary, "consequence": CONSEQUENCE, "no_substitute": NO_SUBSTITUTE,
                  "install_methods": dict(install.get("install_methods", {})), "statement": install.get("statement", "")}
    check = {"font_status": status, "font_method": method, "missing_fonts": missing if status == "missing" else [],
             "unverified_fonts": entries if status == "undetected" else [], "font_notice": notice,
             "unit_fonts": dict(params["fonts"].get("unit_override", {}))}
    return fonts, check


def status_line(report: dict, where: str = "开头") -> str:
    """一句话说明本机字体检测结论（where：醒目提示区块相对本句的位置）；报告里没有字体核对结果时返回空串。"""
    status = report.get("font_status")
    method = report.get("font_method") or "未执行检测"
    if status == "complete":
        return f"标准字体均已安装（{method}）"
    if status == "missing":
        names = "、".join(e["font"] for e in report.get("missing_fonts", []))
        return f"缺少 {len(report.get('missing_fonts', []))} 款标准字体（{names}），获取与安装方法见{where}的“字体缺失”提示"
    if status == "undetected":
        return f"无法检测（{method}），需在打开文档的电脑上人工核对，说明见{where}的提示"
    return ""


def render_notice(report: dict, markdown: bool = False) -> str:
    """把字体核对结果排成醒目的提示区块（多行文本）；字体齐全或没有核对结果时返回空串。

    report 可以是 resolve_fonts 返回的核对结果，也可以是已并入核对结果的 build、fix 报告。
    markdown=True 时排成 Markdown 引用块（写进修复报告 .md 用），否则为屏幕上的纯文本区块。
    """
    notice = report.get("font_notice")
    if not notice:
        return ""
    missing = report.get("font_status") == "missing"
    entries = (report.get("missing_fonts") if missing else report.get("unverified_fonts")) or []
    ask = ("请自行下载安装缺失的字体，装好后重新打开文档核对：" if missing
           else "如果缺少，请自行下载安装（逐款说明如下），装好后重新打开文档核对：")
    method = f"检测方式：{report.get('font_method') or '未执行检测'}"
    items = []  # [(字体及用途, [(标签, 说明)])]
    for entry in entries:
        guide = entry.get("guide", {})
        usage = f"（用于{guide['usage']}）" if guide.get("usage") else ""
        details = [(label, guide[key]) for label, key in (("来源", "source"), ("获取", "obtain"), ("注意", "note"))
                   if guide.get(key)]
        items.append((entry["font"], usage, details))
    methods = list(notice["install_methods"].items())
    if markdown:  # 段落之间空一行，字体和安装方法用列表，避免被 Markdown 合并成一个段落
        out = [f"**【{notice['headline']}】请先看这里**", "", notice["summary"], "", "**影响**：" + notice["consequence"], "",
               notice["no_substitute"], "", ask, ""]
        for i, (font, usage, details) in enumerate(items, start=1):
            out.append(f"{i}. **{font}**{usage}")
            out += [f"   - {label}：{text}" for label, text in details]
        out += ["", "通用安装方法：", ""] + [f"- {k}：{v}" for k, v in methods]
        out += ["", "声明：" + notice["statement"], "", method]
        return "\n".join(("> " + line).rstrip() for line in out)
    bar = "=" * 18
    out = [f"{bar} 【{notice['headline']}】请先看这里 {bar}", notice["summary"], "影响：" + notice["consequence"],
           notice["no_substitute"], ask]
    for i, (font, usage, details) in enumerate(items, start=1):
        out.append(f" {i}. {font}{usage}")
        out += [f"    {label}：{text}" for label, text in details]
    out.append("通用安装方法：")
    out += [f"  - {k}：{v}" for k, v in methods]
    out += ["声明：" + notice["statement"], method, "=" * (2 * len(bar) + 2)]
    return "\n".join(out)


# ---------------------------------------------------------------------------
# 命令行
# ---------------------------------------------------------------------------

def main(argv: Optional[List[str]] = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    parser = argparse.ArgumentParser(
        description="检查本机是否装有公文标准字体（GB/T 9704 字体名称表）；缺失时输出获取途径和安装方法。"
                    "本工具不提供、不打包任何字体文件。",
        epilog="退出码：0 标准字体齐全；3 缺字体或无法检测；2 参数文件读取失败。")
    parser.add_argument("--json", action="store_true", help="以 JSON 输出（各字体是否已安装、缺失字体及其获取与安装引导）")
    parser.add_argument("--unit-fonts", metavar="角色=字体名,…", help=UNIT_FONTS_HELP)
    args = parser.parse_args(argv)
    try:
        params = load_params()
    except (OSError, ValueError) as exc:
        print(f"错误：读取字体参数失败：{exc}")
        return 2
    try:
        params = apply_unit_fonts(params, parse_unit_fonts(args.unit_fonts, params))
    except ValueError as exc:
        print(f"错误：{exc}")
        return 2
    installed, method = installed_fonts()
    fonts, check = resolve_fonts(params, installed, method)
    state = {std: (None if check["font_status"] == "undetected" else is_installed(std, installed))
             for std in fonts.values()}
    code = 0 if check["font_status"] == "complete" else EXIT_FONTS_INCOMPLETE
    if args.json:
        print(json.dumps({"fonts": fonts, "installed": state, **check}, ensure_ascii=False, indent=2))
        return code
    print("公文标准字体检查（GB/T 9704-2012 字体名称表）")
    print(f"检测方式：{check['font_method']}")
    for role, std in fonts.items():
        mark = {True: "已安装", False: "未安装", None: "未检测"}[state[std]]
        usage = font_guide(params, role).get("usage", "")
        print(f"  [{mark}] {std}（{role}：{usage}）")
    for line in unit_notes(check):
        print(f"  单位规定：{line}")
    print(f"结论：{status_line(check, '下方')}")
    notice = render_notice(check)
    if notice:
        print()
        print(notice)
    return code


if __name__ == "__main__":
    sys.exit(main())
