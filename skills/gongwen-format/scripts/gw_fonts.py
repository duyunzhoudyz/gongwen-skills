# -*- coding: utf-8 -*-
"""本机字体检测与字体类别判断（只依赖标准库）。

- Windows：读取 C:\\Windows\\Fonts 和用户字体目录（%LOCALAPPDATA%\\Microsoft\\Windows\\Fonts）。
- macOS：读取 /Library/Fonts、~/Library/Fonts，以及系统字体目录。
- Linux：调用 fc-list。

字体文件只读取 name 表中的字体族名称，不复制、不修改任何字体文件。
"""
from __future__ import annotations

import os
import platform
import re
import struct
import subprocess
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

FONT_SUFFIXES = {".ttf", ".otf", ".ttc"}
_CACHE: Optional[Tuple[Set[str], str]] = None

# 字体类别关键词（小写比较）。类别用于检查"是不是仿宋体、楷体、黑体、小标宋体"
CLASS_KEYWORDS = {
    "小标宋": ("小标宋", "xiaobiaosong", "fzxbs"),
    "仿宋": ("仿宋", "fangsong", "fzfs", "stfangsong"),
    "楷体": ("楷体", "楷", "kaiti", "stkaiti", "fzkai"),
    "黑体": ("黑体", "simhei", "heiti", "fzhei", "stheiti"),
    "宋体": ("宋体", "中宋", "simsun", "songti", "stsong", "stzhongsong", "nsimsun", "宋"),
}
# 常见英文名与中文名对照，检测时两种写法都算已安装
ALIASES = {
    "仿宋_gb2312": ("fangsong_gb2312",),
    "楷体_gb2312": ("kaiti_gb2312",),
    "方正小标宋简体": ("fzxiaobiaosong-b05s", "方正小标宋_gbk", "fzxiaobiaosong-b05"),
    "仿宋": ("fangsong", "simfang"),
    "楷体": ("kaiti", "simkai"),
    "黑体": ("simhei",),
    "宋体": ("simsun",),
    "华文中宋": ("stzhongsong",),
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
    key = norm(font)
    if key in installed:
        return True
    return any(norm(a) in installed for a in ALIASES.get(key, ()))


# ---------------------------------------------------------------------------
# 字体方案
# ---------------------------------------------------------------------------

def resolve_fonts(params: dict, fallback: bool, installed: Optional[Set[str]],
                  method: str = "") -> Tuple[Dict[str, Tuple[str, bool]], List[str], List[str]]:
    """确定各字体角色实际写入的字体名。

    返回 (角色 -> (字体名, 是否额外加粗), 警告列表, 说明列表)。
    角色为 params["fonts"]["standard"] 的键：仿宋、楷体、黑体、小标宋、宋体。
    - 默认写入标准字体名；本机缺字体时只警告。
    - fallback=True 时，缺失的标准字体改用本机已安装的第一个替代字体。
    """
    fonts_cfg = params["fonts"]
    standard: Dict[str, str] = fonts_cfg["standard"]
    alternatives: Dict[str, List[str]] = fonts_cfg["fallback"]
    bold_when: Dict[str, List[str]] = fonts_cfg.get("fallback_bold", {})
    result: Dict[str, Tuple[str, bool]] = {}
    warnings: List[str] = []
    notes: List[str] = []
    can_detect = bool(installed)
    if not can_detect:
        warnings.append(f"无法检测本机字体（{method or '未检测'}），已写入标准字体名；请在打开文档的电脑上核对字体")
    for role, std in standard.items():
        if not can_detect:
            result[role] = (std, False)
            continue
        if is_installed(std, installed):
            result[role] = (std, False)
            continue
        if fallback:
            alt = next((a for a in alternatives.get(role, []) if is_installed(a, installed)), None)
            if alt:
                result[role] = (alt, alt in bold_when.get(role, []))
                extra = "，并加粗" if result[role][1] else ""
                notes.append(f"{role}：本机未安装“{std}”，已改用“{alt}”{extra}")
                continue
            warnings.append(f"本机未安装“{std}”，也没有可用的替代字体（{'、'.join(alternatives.get(role, []))}），仍写入“{std}”")
            result[role] = (std, False)
        else:
            alts = "、".join(alternatives.get(role, []))
            warnings.append(f"本机未安装“{std}”，Word/WPS 会用其他字体显示。可安装已授权的该字体，或加 --font-fallback 改用本机替代字体（{alts}）")
            result[role] = (std, False)
    return result, warnings, notes
