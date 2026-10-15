#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 docx 转成 PDF 和逐页 PNG，供目检排版效果（有条件时使用）。

- 需要本机装有 LibreOffice（soffice）；找不到时给出提示后退出，不报错。
- 装有 pymupdf 时再把 PDF 转成逐页 PNG；没有时只输出 PDF。
- 只在本机运行，不联网。LibreOffice 的排版与 Word、WPS 略有差异，正式印发前仍以打印核对为准。

退出码：0 完成或环境不具备（已提示）；2 输入错误或转换失败。
"""
from __future__ import annotations

import argparse
import platform
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import List, Optional, Sequence

WINDOWS_CANDIDATES = [r"C:\Program Files\LibreOffice\program\soffice.exe",
                      r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]
MAC_CANDIDATES = ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]


def find_soffice() -> Optional[str]:
    """按 PATH、常见安装位置查找 soffice。"""
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    system = platform.system()
    candidates = WINDOWS_CANDIDATES if system == "Windows" else MAC_CANDIDATES if system == "Darwin" else []
    for cand in candidates:
        if Path(cand).is_file():
            return cand
    return None


def unique_path(path: Path) -> Path:
    """输出路径已存在时依次加 -v2、-v3，不覆盖已有文件（与 gw_docx.unique_path 相同；本脚本不依赖 python-docx）。"""
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    m = re.match(r"^(.*)-v(\d+)$", stem)
    base, k = (m.group(1), int(m.group(2)) + 1) if m else (stem, 2)
    while path.with_name(f"{base}-v{k}{suffix}").exists():
        k += 1
    return path.with_name(f"{base}-v{k}{suffix}")


def to_pdf(soffice: str, src: Path, out_dir: Path, timeout: int = 180) -> Path:
    """调用 LibreOffice 无界面转换。用临时用户配置，避免与已打开的 LibreOffice 冲突。

    LibreOffice 会直接覆盖输出目录中的同名 PDF，因此先转到临时目录，再按不覆盖的文件名移入 out_dir。
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    profile = tempfile.mkdtemp(prefix="gw-lo-")
    work = tempfile.mkdtemp(prefix="gw-pdf-")
    try:
        cmd = [soffice, f"-env:UserInstallation={Path(profile).as_uri()}", "--headless", "--convert-to", "pdf",
               "--outdir", work, str(src)]
        subprocess.run(cmd, capture_output=True, timeout=timeout, check=False)
        made = Path(work) / (src.stem + ".pdf")
        if not made.is_file():
            raise RuntimeError("LibreOffice 没有生成 PDF，请确认文件能正常打开")
        pdf = unique_path(out_dir / made.name)
        shutil.move(str(made), str(pdf))
    finally:
        shutil.rmtree(profile, ignore_errors=True)
        shutil.rmtree(work, ignore_errors=True)
    return pdf


def to_png(pdf: Path, out_dir: Path, dpi: int = 80) -> List[Path]:
    """用 pymupdf 把 PDF 逐页转成 PNG；没有 pymupdf 时返回空列表。"""
    try:
        import pymupdf  # noqa: F401
    except ImportError:
        try:
            import fitz as pymupdf  # 旧版 PyMuPDF 的模块名
        except ImportError:
            return []
    pages: List[Path] = []
    doc = pymupdf.open(str(pdf))
    try:
        for i, page in enumerate(doc, start=1):
            target = unique_path(out_dir / f"{pdf.stem}-第{i}页.png")
            page.get_pixmap(dpi=dpi).save(str(target))
            pages.append(target)
    finally:
        doc.close()
    return pages


def main(argv: Optional[Sequence[str]] = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    parser = argparse.ArgumentParser(description="把 docx 转成 PDF 和逐页 PNG，用于目检排版效果（需要 LibreOffice）。",
                                     epilog="示例：python scripts/preview_docx.py 成稿.docx --dpi 80")
    parser.add_argument("file", help="要预览的 .docx")
    parser.add_argument("-o", "--outdir", help="输出目录；默认在 docx 同目录下建“<文件名>-预览”")
    parser.add_argument("--dpi", type=int, default=80, help="PNG 分辨率，默认 80")
    parser.add_argument("--soffice", help="soffice 可执行文件路径（自动查找失败时指定）")
    args = parser.parse_args(argv)

    src = Path(args.file)
    if not src.is_file() or src.suffix.lower() != ".docx":
        print(f"错误：找不到 docx 文件 {src}")
        return 2
    soffice = args.soffice or find_soffice()
    if not soffice or not (Path(soffice).is_file() or shutil.which(soffice)):
        print("未找到 LibreOffice（soffice），跳过预览。需要目检时可安装 LibreOffice，或直接用 Word、WPS 打开并打印核对。")
        return 0
    out_dir = Path(args.outdir) if args.outdir else src.with_name(src.stem + "-预览")
    try:
        pdf = to_pdf(soffice, src.resolve(), out_dir.resolve())
    except (OSError, subprocess.SubprocessError, RuntimeError) as exc:
        print(f"错误：转换失败：{exc}")
        return 2
    pages = to_png(pdf, out_dir, dpi=max(30, min(args.dpi, 300)))
    print(f"PDF：{pdf}")
    if pages:
        print(f"逐页图片 {len(pages)} 张：{out_dir}")
    else:
        print("未安装 pymupdf，只输出了 PDF。需要逐页图片时运行：pip install pymupdf")
    print("提示：LibreOffice 与 Word、WPS 的排版略有差异，字体缺失时会用其他字体代替；红头和版记位置以打印核对为准。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
