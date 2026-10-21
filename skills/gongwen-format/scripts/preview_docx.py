#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 docx 转成 PDF 和逐页 PNG\uff0c供目检排版效果\uff08有条件时使用\uff09\u3002

- 字体严格模式\uff1a本机缺任一标准字体\uff08或无法检测\uff09时拒绝渲染\uff0c输出缺失字体清单和安装引导\uff0c没有绕过选项\u3002
  原因\uff1aLibreOffice 遇到缺失字体会静默换用其他字体\uff0c渲染出的不是正式效果\uff0c据此目检版式会得出错误结论\u3002
- 需要本机装有 LibreOffice\uff08soffice\uff09\uff1b找不到时给出提示后退出\uff0c不报错\u3002
- 装有 pymupdf 时再把 PDF 转成逐页 PNG\uff1b没有时只输出 PDF\u3002
- 只在本机运行\uff0c不联网\u3002LibreOffice 的排版与 Word\u3001WPS 略有差异\uff0c正式印发前仍以打印核对为准\u3002

退出码\uff1a0 完成或环境不具备\uff08已提示\uff09\uff1b2 输入错误或转换失败\uff1b3 缺标准字体或无法检测字体\uff0c拒绝渲染\u3002
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

sys.path.insert(0, str(Path(__file__).resolve().parent))

import gw_fonts as F  # noqa: E402  只依赖标准库，本脚本仍不依赖 python-docx

WINDOWS_CANDIDATES = [r"C:\Program Files\LibreOffice\program\soffice.exe",
                      r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]
MAC_CANDIDATES = ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]


def find_soffice() -> Optional[str]:
    """按 PATH\u3001常见安装位置查找 soffice\u3002"""
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
    """输出路径已存在时依次加 -v2\u3001-v3\uff0c不覆盖已有文件\uff08与 gw_docx.unique_path 相同\uff1b本脚本不依赖 python-docx\uff09\u3002"""
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    m = re.match(r"^(.*)-v(\d+)$", stem)
    base, k = (m.group(1), int(m.group(2)) + 1) if m else (stem, 2)
    while path.with_name(f"{base}-v{k}{suffix}").exists():
        k += 1
    return path.with_name(f"{base}-v{k}{suffix}")


def to_pdf(soffice: str, src: Path, out_dir: Path, timeout: int = 180) -> Path:
    """调用 LibreOffice 无界面转换\u3002用临时用户配置\uff0c避免与已打开的 LibreOffice 冲突\u3002

    LibreOffice 会直接覆盖输出目录中的同名 PDF\uff0c因此先转到临时目录\uff0c再按不覆盖的文件名移入 out_dir\u3002
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
            raise RuntimeError("LibreOffice 没有生成 PDF\uff0c请确认文件能正常打开")
        pdf = unique_path(out_dir / made.name)
        shutil.move(str(made), str(pdf))
    finally:
        shutil.rmtree(profile, ignore_errors=True)
        shutil.rmtree(work, ignore_errors=True)
    return pdf


def to_png(pdf: Path, out_dir: Path, dpi: int = 80) -> List[Path]:
    """用 pymupdf 把 PDF 逐页转成 PNG\uff1b没有 pymupdf 时返回空列表\u3002"""
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


PAGE_NO_RE = re.compile(r"^[\u2014-]\s*[0-9]+\s*[\u2014-]$")
DATE_LINE_RE = re.compile(r"^\d{4}\s*年\s*\d{1,2}\s*月\s*(?:\d{1,2}\s*)?日$")
TAIL_LINE_RE = re.compile(r"^(抄送|附件)\s*[\uff1a:]|^\d{1,2}\s*[.．\u3001]\s*\S|^[\uff08(].*[\uff09)]$|印发$")


def analyze_pages(pdf: Path, bottom_mm: float = 35.0) -> List[str]:
    """读 PDF 逐页文字\uff0c给出版面事实\uff08交付说明只能写这里核对到的\uff09\uff1a落款是否与正文同页\u3001首页下方留白\u3002没有 pymupdf 时返回空列表\u3002

    落款孤页\uff1a末页\uff08超过一页时\uff09去掉页码\u3001成文日期\u3001单位名称短行\u3001附注\u3001抄送\u3001印发行\u3001附件说明后\uff0c没有正文行了\u3002
    首页留白\uff1a多页文档的首页\uff0c正文最后一行到版心下边缘\uff08页下边距 bottom_mm\uff09的距离大于 60 mm\u3002"""
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf
        except ImportError:
            return []
    facts: List[str] = []
    doc = pymupdf.open(str(pdf))
    try:
        pages = []
        for page in doc:
            limit = page.rect.height - bottom_mm * 72 / 25.4
            blocks = [b for b in page.get_text("blocks") if b[3] <= limit + 2 and b[4].strip()]
            lines = [ln.strip() for b in blocks for ln in b[4].split("\n") if ln.strip()]
            pages.append((lines, max([b[3] for b in blocks], default=0.0), limit))
    finally:
        doc.close()
    if not pages:
        return facts
    facts.append(f"共 {len(pages)} 页")
    if len(pages) > 1:
        lines = [ln for ln in pages[-1][0] if not PAGE_NO_RE.match(ln)]
        body = [ln for ln in lines if not (DATE_LINE_RE.match(ln.replace(" ", "")) or TAIL_LINE_RE.match(ln)
                                           or (len(ln) <= 20 and not any(ch in ln for ch in "\u3002\uff0c\uff1b\uff1a\uff01\uff1f\u3001")))]
        has_date = any(DATE_LINE_RE.match(ln.replace(" ", "")) for ln in lines)
        if has_date and not body:
            facts.append("末页只有落款\uff08署名\u3001成文日期\uff0c可能连同附件说明\u3001版记\uff09\uff0c落款与正文最后一段不在同一页\uff0c需要调整")
        elif has_date:
            facts.append("落款与正文最后一段在同一页\uff08末页\uff09")
        first_bottom, limit = pages[0][1], pages[0][2]
        gap_mm = (limit - first_bottom) * 25.4 / 72
        if gap_mm > 60:
            facts.append(f"首页正文下方留白约 {gap_mm:.0f} mm（页面其余部分被后移的段落占用），请核对是否合适")
    return facts


def check_fonts(unit_fonts=None) -> int:
    """字体严格模式\uff1a本机标准字体齐全时返回 0\uff1b缺字体或无法检测时输出缺失清单和安装引导\uff0c返回 3\uff08拒绝渲染\uff09\u3002

    LibreOffice 遇到缺失字体会静默换用其他字体渲染\uff0c据此目检版式会得出错误结论\uff0c所以不提供绕过选项\u3002
    """
    try:
        params = F.load_params()
    except (OSError, ValueError) as exc:
        print(f"错误：读取字体参数失败：{exc}")
        return 2
    try:
        params = F.apply_unit_fonts(params, F.parse_unit_fonts(unit_fonts, params))
    except ValueError as exc:
        print(f"错误：{exc}")
        return 2
    installed, method = F.installed_fonts()
    _fonts, check = F.resolve_fonts(params, installed, method)
    if check["font_status"] == "complete":
        return 0
    print(F.render_notice(check))
    print()
    print("预览已拒绝\uff1a本机缺少标准字体\uff08或无法确认字体齐全\uff09时\uff0cLibreOffice 会静默换用其他字体渲染\uff0c"
          "预览看到的不是正式效果\uff0c不能据此核对版式\u3002请先按上面的提示安装字体\uff0c再重新运行本命令\uff1b本命令没有绕过选项\u3002")
    return F.EXIT_FONTS_INCOMPLETE


def main(argv: Optional[Sequence[str]] = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    parser = argparse.ArgumentParser(description="把 docx 转成 PDF 和逐页 PNG\uff0c用于目检排版效果\uff08需要 LibreOffice\uff09\u3002"
                                                 "本机缺任一标准字体\uff08或无法检测\uff09时拒绝渲染\uff0c没有绕过选项\u3002",
                                     epilog="示例\uff1apython scripts/preview_docx.py 成稿.docx --dpi 80\n"
                                            "退出码\uff1a0 完成或找不到 LibreOffice\uff08只提示\uff09\uff1b2 输入错误或转换失败\uff1b3 缺标准字体\uff0c拒绝渲染\u3002",
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("file", help="要预览的 .docx")
    parser.add_argument("-o", "--outdir", help="输出目录\uff1b默认在 docx 同目录下建\u201c<文件名>-预览\u201d")
    parser.add_argument("--dpi", type=int, default=80, help="PNG 分辨率\uff0c默认 80")
    parser.add_argument("--soffice", help="soffice 可执行文件路径\uff08自动查找失败时指定\uff09")
    parser.add_argument("--unit-fonts", metavar="角色=字体名,…", help=F.UNIT_FONTS_HELP + "\uff1b字体齐全检查按单位指定字体核对")
    args = parser.parse_args(argv)

    src = Path(args.file)
    if not src.is_file() or src.suffix.lower() != ".docx":
        print(f"错误：找不到 docx 文件 {src}")
        return 2
    code = check_fonts(args.unit_fonts)
    if code:
        return code
    soffice = args.soffice or find_soffice()
    if not soffice or not (Path(soffice).is_file() or shutil.which(soffice)):
        print("未找到 LibreOffice\uff08soffice\uff09\uff0c跳过预览\u3002需要目检时可安装 LibreOffice\uff0c或直接用 Word\u3001WPS 打开并打印核对\u3002")
        return 0
    out_dir = Path(args.outdir) if args.outdir else src.with_name(src.stem + "-预览")
    try:
        pdf = to_pdf(soffice, src.resolve(), out_dir.resolve())
    except (OSError, subprocess.SubprocessError, RuntimeError) as exc:
        print(f"错误：转换失败：{exc}")
        return 2
    pages = to_png(pdf, out_dir, dpi=max(30, min(args.dpi, 300)))
    print(f"PDF：{pdf}")
    try:
        facts = analyze_pages(pdf)
    except Exception:  # noqa: BLE001 版面核对只是附加信息，PDF 读不了时不影响预览
        facts = []
    for fact in facts:
        print(f"版面核对：{fact}")
    if pages:
        print(f"逐页图片 {len(pages)} 张：{out_dir}")
    else:
        print("未安装 pymupdf\uff0c只输出了 PDF\u3002需要逐页图片时运行\uff1apip install pymupdf")
    print("提示\uff1aLibreOffice 与 Word\u3001WPS 的排版略有差异\uff1b红头和版记位置以打印核对为准\u3002")
    return 0


if __name__ == "__main__":
    sys.exit(main())
