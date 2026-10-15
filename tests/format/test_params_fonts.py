# -*- coding: utf-8 -*-
"""参数文件与 gbt9704.md 一致、每项注明依据；字体检测与类别判断。"""
import re
import unittest
from pathlib import Path

from tests.format.helpers import SKILL, gw_docx

import gw_fonts  # noqa: E402

RULES = SKILL / "references" / "_shared" / "rules" / "gbt9704.md"


class ParamsTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.params = gw_docx.load_params()
        cls.rules = RULES.read_text(encoding="utf-8") if RULES.exists() else ""

    def test_every_value_has_basis(self):
        def walk(node, path):
            if isinstance(node, dict) and "value" in node:
                self.assertTrue(node.get("basis"), f"{path} 缺少 basis")
                return
            if isinstance(node, dict):
                for k, v in node.items():
                    if k not in ("basis", "note", "source", "version"):
                        walk(v, f"{path}.{k}")
        for group in ("page", "grid", "sizes_pt", "colors"):
            walk(self.params[group], group)
        for name, el in self.params["elements"].items():
            basis = el.get("basis", {})
            for key in el:
                if key != "basis":
                    self.assertIn(key, basis, f"elements.{name}.{key} 缺少 basis")
                    text = basis[key]
                    self.assertTrue(re.search(r"\d+\.\d|实践值|推算", text), f"elements.{name}.{key} 依据不明：{text}")
        self.assertTrue(self.params["fonts"]["basis"])
        self.assertTrue(self.params["seal_default"]["basis"])

    def test_values_match_gbt9704_md(self):
        if not self.rules:
            self.skipTest("references/_shared 尚未同步")
        pv = lambda k: gw_docx.pv(self.params, k)  # noqa: E731
        self.assertIn("37 mm±1 mm", self.rules)
        self.assertIn("28 mm±1 mm", self.rules)
        self.assertIn("156 mm×225 mm", self.rules)
        self.assertEqual((pv("page.margin_top_mm"), pv("page.margin_bottom_mm"), pv("page.margin_left_mm"),
                          pv("page.margin_right_mm")), (37, 35, 28, 26))
        self.assertIn("上 3.7 cm，下 3.5 cm，左 2.8 cm，右 2.6 cm", self.rules)
        self.assertEqual((pv("grid.lines_per_page"), pv("grid.chars_per_line")), (22, 28))
        self.assertIn("每面 22 行，每行 28 字", self.rules)
        self.assertEqual(pv("grid.line_spacing_pt"), 28.9)
        self.assertIn("28.9 磅", self.rules)
        self.assertEqual((pv("sizes_pt.二号"), pv("sizes_pt.三号"), pv("sizes_pt.四号")), (22, 16, 14))
        self.assertIn("二号 22 磅，三号 16 磅，四号 14 磅", self.rules)
        el = self.params["elements"]
        self.assertEqual(el["issuer"]["top_from_body_mm"], 35)
        self.assertEqual(el["header_rule"]["gap_mm"], 4)
        self.assertEqual((el["letter"]["issuer_top_mm"], el["letter"]["rule_length_mm"],
                          el["letter"]["bottom_rule_from_page_mm"]), (30, 170, 20))
        self.assertIn("线长均 170 mm", self.rules)
        self.assertEqual((el["h1"]["font"], el["h2"]["font"], el["h3"]["font"], el["h4"]["font"]),
                         ("黑体", "楷体", "仿宋", "仿宋"))
        std = self.params["fonts"]["standard"]
        for name in ("仿宋_GB2312", "楷体_GB2312", "方正小标宋简体"):
            self.assertIn(name, std.values())
            self.assertIn(name, self.rules)

    def test_grid_char_space_gives_28_chars(self):
        p = self.params
        width = (gw_docx.mm_to_twips(210) - gw_docx.mm_to_twips(gw_docx.pv(p, "page.margin_left_mm"))
                 - gw_docx.mm_to_twips(gw_docx.pv(p, "page.margin_right_mm"))) / 20
        pitch = 16 + gw_docx.pv(p, "grid.doc_grid_char_space") / 4096
        self.assertEqual(int(width // pitch), 28)
        self.assertEqual(int(gw_docx.pv(p, "page.body_height_mm") * 72 / 25.4 // gw_docx.pv(p, "grid.line_spacing_pt")), 22)


class FontsTest(unittest.TestCase):
    def test_font_class(self):
        cases = {"仿宋_GB2312": "仿宋", "FangSong": "仿宋", "方正仿宋_GBK": "仿宋", "楷体_GB2312": "楷体",
                 "KaiTi": "楷体", "黑体": "黑体", "SimHei": "黑体", "方正小标宋简体": "小标宋", "华文中宋": "宋体",
                 "宋体": "宋体", "微软雅黑": "其他", "等线": "其他", None: "其他"}
        for name, cls in cases.items():
            self.assertEqual(gw_fonts.font_class(name), cls, name)

    def test_detection_never_raises(self):
        names, method = gw_fonts.installed_fonts(refresh=True)
        self.assertIsInstance(names, set)
        self.assertTrue(method)

    def test_resolve_without_detection_keeps_standard(self):
        params = gw_docx.load_params()
        fonts, warnings, _ = gw_fonts.resolve_fonts(params, True, set(), "测试")
        self.assertEqual(fonts["仿宋"], ("仿宋_GB2312", False))
        self.assertTrue(warnings)

    def test_alias_detection(self):
        params = gw_docx.load_params()
        fonts, warnings, notes = gw_fonts.resolve_fonts(params, True, {"fangsong_gb2312", "stzhongsong", "kaiti",
                                                                       "simhei", "simsun"}, "测试")
        self.assertEqual(fonts["仿宋"][0], "仿宋_GB2312")  # 英文名也算已安装
        self.assertEqual(fonts["小标宋"], ("华文中宋", False))  # 优先华文中宋，不加粗
        self.assertEqual(fonts["楷体"][0], "楷体")

    def test_fc_list_escapes(self):
        """fc-list 输出的族名对 - : , \\ 加反斜杠转义（如 FZXiaoBiaoSong\\-B05S），要还原后再比对。"""
        from unittest import mock

        class Result:
            stdout = "FZXiaoBiaoSong\\-B05S\nFangSong_GB2312,仿宋_GB2312\nFoo\\, Bar,Baz\n"

        with mock.patch.object(gw_fonts.subprocess, "run", return_value=Result()):
            names, ok = gw_fonts._scan_fc_list()
        self.assertTrue(ok)
        self.assertIn("fzxiaobiaosong-b05s", names)
        self.assertIn("foo,bar", names)
        self.assertIn("baz", names)
        self.assertTrue(gw_fonts.is_installed("方正小标宋简体", names))
        self.assertTrue(gw_fonts.is_installed("仿宋_GB2312", names))

    def test_read_font_file_names(self):
        path = Path(r"C:\Windows\Fonts\simfang.ttf")
        if not path.exists():
            self.skipTest("本机没有 simfang.ttf")
        names = gw_fonts.read_font_names(path)
        self.assertTrue(any("仿宋" in n or "FangSong" in n for n in names), names)


if __name__ == "__main__":
    unittest.main()
