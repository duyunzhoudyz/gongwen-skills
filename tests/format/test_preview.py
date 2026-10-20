# -*- coding: utf-8 -*-
"""preview_docx：缺标准字体时拒绝渲染（严格模式）；字体齐全时，找不到 LibreOffice 提示后正常退出，有 LibreOffice 时转出 PDF 和 PNG。"""
import io
import sys
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest import mock

from tests.format.helpers import FULL_FONTS, NO_FONTS, OUT, WIN_FONTS, build_sample, fonts_installed

import preview_docx  # noqa: E402  helpers 已把 scripts 加入 sys.path


class PreviewTest(unittest.TestCase):
    """字体齐全（用 fonts_installed 模拟）时的常规行为。"""

    def setUp(self):
        patcher = fonts_installed(FULL_FONTS)
        patcher.__enter__()
        self.addCleanup(patcher.__exit__, None, None, None)

    def test_without_soffice(self):
        src, _ = build_sample("白头", "预览-无软件.docx")
        buf = io.StringIO()
        with mock.patch.object(preview_docx, "find_soffice", return_value=None), redirect_stdout(buf):
            code = preview_docx.main([str(src)])
        self.assertEqual(code, 0)
        self.assertIn("未找到 LibreOffice", buf.getvalue())

    def test_bad_input(self):
        with redirect_stdout(io.StringIO()):
            self.assertEqual(preview_docx.main([str(OUT / "不存在.docx")]), 2)

    def test_existing_pdf_not_overwritten(self):
        """输出目录里已有同名 PDF（如 -o 指向 docx 所在目录、已有正式导出的“成稿.pdf”）时不覆盖，另存为 -v2。"""
        import shutil
        folder = OUT / "预览 不覆盖"
        if folder.exists():
            shutil.rmtree(folder)
        folder.mkdir(parents=True)
        src, _ = build_sample("白头", "预览-不覆盖.docx")
        shutil.copy(src, folder / "成稿.docx")
        (folder / "成稿.pdf").write_bytes(b"USER PDF")

        def fake_run(cmd, **_kw):
            # 模拟 LibreOffice：在 --outdir 指定的目录写出 <文件名>.pdf
            out_dir = Path(cmd[cmd.index("--outdir") + 1])
            (out_dir / (Path(cmd[-1]).stem + ".pdf")).write_bytes(b"NEW PDF")

        buf = io.StringIO()
        with mock.patch.object(preview_docx.subprocess, "run", side_effect=fake_run), \
                mock.patch.object(preview_docx, "to_png", return_value=[]), redirect_stdout(buf):
            code = preview_docx.main([str(folder / "成稿.docx"), "-o", str(folder), "--soffice", sys.executable])
        self.assertEqual(code, 0, buf.getvalue())
        self.assertEqual((folder / "成稿.pdf").read_bytes(), b"USER PDF")
        self.assertEqual((folder / "成稿-v2.pdf").read_bytes(), b"NEW PDF")

    def test_existing_png_not_overwritten(self):
        try:
            import pymupdf  # noqa: F401
        except ImportError:
            self.skipTest("没有 pymupdf")
        import shutil
        folder = OUT / "预览 图片不覆盖"
        if folder.exists():
            shutil.rmtree(folder)
        folder.mkdir(parents=True)
        pdf = folder / "成稿.pdf"
        doc = pymupdf.open()
        doc.new_page()
        doc.save(str(pdf))
        doc.close()
        (folder / "成稿-第1页.png").write_bytes(b"USER PNG")
        pages = preview_docx.to_png(pdf, folder, dpi=30)
        self.assertEqual((folder / "成稿-第1页.png").read_bytes(), b"USER PNG")
        self.assertEqual([p.name for p in pages], ["成稿-第1页-v2.png"])

    @unittest.skipUnless(preview_docx.find_soffice(), "本机没有 LibreOffice")
    def test_convert(self):
        src, _ = build_sample("材料", "预览-年终总结.docx")
        out = OUT / "预览-年终总结-预览"
        for old in out.glob("*") if out.exists() else []:
            old.unlink()
        with redirect_stdout(io.StringIO()):
            code = preview_docx.main([str(src), "-o", str(out), "--dpi", "40"])
        self.assertEqual(code, 0)
        self.assertTrue((out / "预览-年终总结.pdf").exists())


class PreviewRefusesWithoutFontsTest(unittest.TestCase):
    """严格模式：缺任一标准字体（或无法检测）时拒绝渲染，退出码 3，输出缺失清单和安装引导，不调用 LibreOffice，没有绕过选项。"""

    def run_preview(self, installed, extra=()):
        import shutil
        src, _ = build_sample("白头", "预览-缺字体.docx")
        out = OUT / "预览-缺字体-预览"
        shutil.rmtree(out, ignore_errors=True)
        buf = io.StringIO()
        with fonts_installed(installed), redirect_stdout(buf), \
                mock.patch.object(preview_docx.subprocess, "run") as run, \
                mock.patch.object(preview_docx, "to_pdf") as to_pdf:
            code = preview_docx.main([str(src), "-o", str(out), "--soffice", sys.executable, *extra])
        return code, buf.getvalue(), run, to_pdf, out

    def test_refuses_when_fonts_missing(self):
        code, out, run, to_pdf, outdir = self.run_preview(WIN_FONTS)
        self.assertEqual(code, 3)
        run.assert_not_called()
        to_pdf.assert_not_called()
        self.assertFalse(outdir.exists())
        for text in ("【字体缺失】", "仿宋_GB2312", "楷体_GB2312", "方正小标宋简体", "https://www.foundertype.com",
                     "为所有用户安装", "不提供、不打包任何字体文件", "预览已拒绝", "没有绕过选项"):
            self.assertIn(text, out)
        self.assertNotIn("（用于一级标题", out)  # 黑体已装，不在缺失清单中

    def test_refuses_when_fonts_undetected(self):
        code, out, run, to_pdf, _outdir = self.run_preview(NO_FONTS)
        self.assertEqual(code, 3)
        run.assert_not_called()
        to_pdf.assert_not_called()
        self.assertIn("无法检测本机字体，需人工核对", out)
        self.assertIn("预览已拒绝", out)

    def test_refuses_even_without_soffice(self):
        """缺字体的拒绝先于查找 LibreOffice：本机没有 LibreOffice 时也报缺字体，不是“跳过预览”。"""
        src, _ = build_sample("白头", "预览-缺字体无软件.docx")
        buf = io.StringIO()
        with fonts_installed(WIN_FONTS), mock.patch.object(preview_docx, "find_soffice", return_value=None), \
                redirect_stdout(buf):
            code = preview_docx.main([str(src)])
        self.assertEqual(code, 3)
        self.assertIn("【字体缺失】", buf.getvalue())
        self.assertNotIn("未找到 LibreOffice", buf.getvalue())

    def test_no_bypass_option(self):
        """没有 --force、--allow-missing-fonts 之类的绕过参数。"""
        src, _ = build_sample("白头", "预览-无绕过.docx")
        for flag in ("--force", "--allow-missing-fonts", "--ignore-fonts", "--font-fallback"):
            buf = io.StringIO()
            with fonts_installed(WIN_FONTS), redirect_stderr(buf), self.assertRaises(SystemExit) as ctx:
                preview_docx.main([str(src), flag])
            self.assertEqual(ctx.exception.code, 2, flag)
        buf = io.StringIO()
        with redirect_stdout(buf), self.assertRaises(SystemExit):
            preview_docx.main(["--help"])
        self.assertNotIn("--force", buf.getvalue())
        # 参数表里没有绕过字体检查的开关（--unit-fonts 只指定单位规定的字体，同样要求装齐，不是绕过）
        self.assertNotIn("font", buf.getvalue().lower().replace("--unit-fonts", ""))

    def test_renders_when_fonts_complete(self):
        src, _ = build_sample("白头", "预览-齐全.docx")
        out = OUT / "预览-齐全-预览"

        def fake_to_pdf(_soffice, _src, out_dir, timeout=180):
            out_dir.mkdir(parents=True, exist_ok=True)
            pdf = out_dir / "预览-齐全.pdf"
            pdf.write_bytes(b"PDF")
            return pdf

        buf = io.StringIO()
        with fonts_installed(FULL_FONTS), redirect_stdout(buf), mock.patch.object(preview_docx, "to_pdf", side_effect=fake_to_pdf), \
                mock.patch.object(preview_docx, "to_png", return_value=[]):
            code = preview_docx.main([str(src), "-o", str(out), "--soffice", sys.executable])
        self.assertEqual(code, 0, buf.getvalue())
        self.assertNotIn("字体缺失", buf.getvalue())
        self.assertIn("PDF：", buf.getvalue())

    @unittest.skipUnless(preview_docx.find_soffice(), "本机没有 LibreOffice")
    def test_real_check_on_this_machine(self):
        """不模拟字体：按本机真实情况，字体齐全才渲染，缺字体就拒绝（退出码 3），两种结果都合规。"""
        import gw_fonts
        src, _ = build_sample("白头", "预览-本机.docx")
        installed, method = gw_fonts.installed_fonts()
        _fonts, check = gw_fonts.resolve_fonts(gw_fonts.load_params(), installed, method)
        out = OUT / "预览-本机-预览"
        for old in out.glob("*") if out.exists() else []:
            old.unlink()
        buf = io.StringIO()
        with redirect_stdout(buf):
            code = preview_docx.main([str(src), "-o", str(out), "--dpi", "30"])
        if check["font_status"] == "complete":
            self.assertEqual(code, 0, buf.getvalue())
        else:
            self.assertEqual(code, 3, buf.getvalue())
            self.assertFalse(out.exists() and any(out.glob("*.pdf")))


if __name__ == "__main__":
    unittest.main()
