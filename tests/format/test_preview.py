# -*- coding: utf-8 -*-
"""preview_docx：找不到 LibreOffice 时提示后正常退出；本机有 LibreOffice 时转出 PDF 和 PNG。"""
import io
import sys
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

from tests.format.helpers import OUT, build_sample

import preview_docx  # noqa: E402  helpers 已把 scripts 加入 sys.path


class PreviewTest(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
