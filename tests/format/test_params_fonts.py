# -*- coding: utf-8 -*-
"""参数文件与 gbt9704.md 一致、每项注明依据；字体检测与类别判断。"""
import inspect
import io
import json
import re
import unittest
from contextlib import redirect_stdout
from pathlib import Path

from tests.format.helpers import FULL_FONTS, SKILL, STANDARD_FONTS, WIN_FONTS, fonts_installed, gw_docx

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

    def test_resolve_never_substitutes(self):
        """严格模式：不管本机装了什么，写入文档的字体一律是标准字体名；没有降级、替代的分支。"""
        params = gw_docx.load_params()
        standard = params["fonts"]["standard"]
        sets = {"无法检测": set(), "齐全": FULL_FONTS, "只有近似字体": {"仿宋", "楷体", "fangsong", "kaiti", "simfang", "simkai",
                                                                   "华文中宋", "stzhongsong", "方正小标宋_gbk", "fzxiaobiaosong-b05",
                                                                   "方正仿宋_gbk", "华文仿宋", "宋体", "黑体"}}
        for label, installed in sets.items():
            fonts, check = gw_fonts.resolve_fonts(params, installed, "测试")
            self.assertEqual(fonts, standard, label)
        self.assertEqual(list(inspect.signature(gw_fonts.resolve_fonts).parameters), ["params", "installed", "method"])

    def test_resolve_status_and_missing(self):
        params = gw_docx.load_params()
        _fonts, check = gw_fonts.resolve_fonts(params, FULL_FONTS, "测试")
        self.assertEqual((check["font_status"], check["missing_fonts"], check["unverified_fonts"], check["font_notice"]),
                         ("complete", [], [], None))
        # 近似字体不算数：仿宋、楷体、华文中宋、方正小标宋_GBK 都不能顶替标准字体
        near = {"仿宋", "楷体", "fangsong", "kaiti", "华文中宋", "stzhongsong", "方正小标宋_gbk", "fzxiaobiaosong-b05", "黑体",
                "宋体"}
        _fonts, check = gw_fonts.resolve_fonts(params, near, "测试")
        self.assertEqual(check["font_status"], "missing")
        self.assertEqual([m["font"] for m in check["missing_fonts"]], ["仿宋_GB2312", "楷体_GB2312", "方正小标宋简体"])
        self.assertEqual(check["font_method"], "测试")
        self.assertEqual(check["unverified_fonts"], [])
        _fonts, check = gw_fonts.resolve_fonts(params, set(), "")
        self.assertEqual(check["font_status"], "undetected")
        self.assertEqual(check["missing_fonts"], [])
        self.assertEqual(len(check["unverified_fonts"]), 5)
        self.assertIn("人工核对", check["font_notice"]["summary"])

    def test_alias_means_same_font_only(self):
        """别名只认同一款字体的不同名称；名称相近的另一款字体（仿宋、方正小标宋_GBK 等）不算。"""
        same = [("仿宋_GB2312", "FangSong_GB2312"), ("楷体_GB2312", "KaiTi_GB2312"), ("方正小标宋简体", "FZXiaoBiaoSong-B05S"),
                ("黑体", "SimHei"), ("宋体", "SimSun"), ("仿宋_GB2312", "仿宋_gb2312")]
        for std, other in same:
            self.assertTrue(gw_fonts.same_font(std, other), (std, other))
            self.assertTrue(gw_fonts.same_font(other, std), (other, std))
            self.assertTrue(gw_fonts.is_installed(std, {gw_fonts.norm(other)}), (std, other))
        different = [("仿宋_GB2312", "仿宋"), ("仿宋_GB2312", "FangSong"), ("仿宋_GB2312", "SimFang"), ("仿宋_GB2312", "方正仿宋_GBK"),
                     ("楷体_GB2312", "楷体"), ("楷体_GB2312", "KaiTi"), ("方正小标宋简体", "方正小标宋_GBK"),
                     ("方正小标宋简体", "FZXiaoBiaoSong-B05"), ("方正小标宋简体", "华文中宋"), ("方正小标宋简体", "宋体"),
                     ("黑体", "方正黑体_GBK"), ("宋体", "华文宋体"), ("宋体", "新宋体"), ("宋体", "NSimSun")]
        for std, other in different:
            self.assertFalse(gw_fonts.same_font(std, other), (std, other))
            self.assertFalse(gw_fonts.same_font(other, std), (other, std))
            self.assertFalse(gw_fonts.is_installed(std, {gw_fonts.norm(other)}), (std, other))
        self.assertFalse(gw_fonts.same_font(None, "宋体"))
        self.assertFalse(gw_fonts.same_font("", ""))
        # ALIASES 里只能是标准字体的别名，不留“近似字体”的条目
        standard = {gw_fonts.norm(v) for v in gw_docx.load_params()["fonts"]["standard"].values()}
        self.assertTrue(set(gw_fonts.ALIASES) <= standard, set(gw_fonts.ALIASES) - standard)
        for aliases in gw_fonts.ALIASES.values():
            for alias in aliases:
                self.assertFalse(alias.endswith(("gbk", "b05")), alias)

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


class InstallGuideTest(unittest.TestCase):
    """fonts.install_guide：每款标准字体都有来源、获取途径、注意事项；参数和文档里没有降级、替代的数据或建议。"""

    @classmethod
    def setUpClass(cls):
        cls.fonts = gw_docx.load_params()["fonts"]
        cls.standard = cls.fonts["standard"]
        cls.guide = cls.fonts["install_guide"]

    def test_no_fallback_data(self):
        for key in ("fallback", "fallback_bold"):
            self.assertNotIn(key, self.fonts)
            self.assertNotIn(key, self.fonts["basis"])
        for key in ("policy", "install_guide", "standard"):
            self.assertTrue(self.fonts["basis"][key])

    def test_every_standard_font_has_guide(self):
        self.assertEqual(set(self.guide["fonts"]), set(self.standard))
        for role, entry in self.guide["fonts"].items():
            self.assertEqual(entry["font"], self.standard[role], role)
            for key in ("usage", "source", "obtain", "note"):
                self.assertTrue(entry[key].strip(), (role, key))
        self.assertIn("不提供、不打包任何字体文件", self.guide["statement"])
        methods = self.guide["install_methods"]
        self.assertEqual(set(methods), {"Windows", "macOS", "Linux", "安装后"})
        self.assertIn("为所有用户安装", methods["Windows"])
        self.assertIn("字体册", methods["macOS"])
        self.assertIn("~/.local/share/fonts", methods["Linux"])
        self.assertIn("fc-cache -f", methods["Linux"])
        self.assertIn("重启 Word/WPS", methods["安装后"])

    def test_facts_and_no_invented_urls(self):
        fonts = self.guide["fonts"]
        song = fonts["小标宋"]
        self.assertIn("北京北大方正电子有限公司", song["source"])
        self.assertIn("https://www.foundertype.com", song["obtain"])
        self.assertIn("商业授权", song["obtain"])
        for role in ("仿宋", "楷体"):
            self.assertIn("Windows XP", fonts[role]["source"])
            self.assertIn("单位信息中心", fonts[role]["obtain"])
        for role in ("黑体", "宋体"):
            self.assertIn("简体中文版 Windows 预装", fonts[role]["source"])
        # 只允许出现方正字库官网这一个网址，不得编造下载链接
        urls = set(re.findall(r"https?://[A-Za-z0-9./_-]+", json.dumps(self.fonts, ensure_ascii=False)))
        self.assertEqual(urls, {"https://www.foundertype.com"})

    def test_docs_consistent_with_guide(self):
        manual = (SKILL / "references" / "manual-setup.md").read_text(encoding="utf-8")
        for name in self.standard.values():
            self.assertIn(name, manual)
            if RULES.exists():
                self.assertIn(name, RULES.read_text(encoding="utf-8"))
        self.assertIn("https://www.foundertype.com", manual)
        for method in ("为所有用户安装", "字体册", "fc-cache -f", "重启 Word/WPS"):
            self.assertIn(method, manual)

    def test_no_substitute_advice_left(self):
        """skill 文档、共享规则、README、评测用例里不再有 --font-fallback 和替代字体的建议。"""
        repo = SKILL.parents[1]
        files = [SKILL / "SKILL.md", SKILL / "README.md", SKILL / "references" / "manual-setup.md",
                 SKILL / "references" / "check-items.md", SKILL / "evals" / "evals.json",
                 repo / "shared" / "rules" / "gbt9704.md", repo / "README.md", repo / "README_EN.md", repo / "使用须知.md",
                 repo / "docs" / "开发规范.md"] + sorted((SKILL / "scripts").glob("*.py"))
        banned = ("--font-fallback", "font_fallback", "缺字体用", "应急用宋体加粗", "缺字体时的常见替代", "缺字体时用",
                  "改用本机替代字体", "使用替代字体")
        for path in files:
            if not path.exists():
                continue
            text = path.read_text(encoding="utf-8")
            for word in banned:
                self.assertNotIn(word, text, f"{path.relative_to(repo)} 仍有“{word}”")


class FontsCliTest(unittest.TestCase):
    """python scripts/gw_fonts.py：各标准字体是否已安装，缺失或无法检测时输出获取与安装引导。"""

    def run_cli(self, args, installed):
        buf = io.StringIO()
        with fonts_installed(installed), redirect_stdout(buf):
            code = gw_fonts.main(args)
        return code, buf.getvalue()

    def test_missing_fonts(self):
        code, out = self.run_cli([], WIN_FONTS)
        self.assertEqual(code, gw_fonts.EXIT_FONTS_INCOMPLETE)
        for line in ("[未安装] 仿宋_GB2312", "[未安装] 楷体_GB2312", "[未安装] 方正小标宋简体", "[已安装] 黑体", "[已安装] 宋体"):
            self.assertIn(line, out)
        for text in ("【字体缺失】", "https://www.foundertype.com", "单位信息中心", "为所有用户安装", "字体册", "fc-cache -f",
                     "不提供、不打包任何字体文件", "不是正式效果"):
            self.assertIn(text, out)
        self.assertNotIn("（用于一级标题", out)  # 黑体已装，只对缺失的字体给获取引导

    def test_complete_fonts(self):
        code, out = self.run_cli([], FULL_FONTS)
        self.assertEqual(code, 0)
        self.assertEqual(out.count("[已安装]"), 5)
        self.assertNotIn("字体缺失", out)

    def test_undetected(self):
        code, out = self.run_cli([], set())
        self.assertEqual(code, gw_fonts.EXIT_FONTS_INCOMPLETE)
        self.assertEqual(out.count("[未检测]"), 5)
        self.assertIn("人工核对", out)
        self.assertIn("https://www.foundertype.com", out)

    def test_json(self):
        code, out = self.run_cli(["--json"], WIN_FONTS)
        data = json.loads(out)
        self.assertEqual(code, gw_fonts.EXIT_FONTS_INCOMPLETE)
        self.assertEqual(data["fonts"], STANDARD_FONTS)
        self.assertEqual(data["installed"], {"仿宋_GB2312": False, "楷体_GB2312": False, "黑体": True,
                                             "方正小标宋简体": False, "宋体": True})
        self.assertEqual(data["font_status"], "missing")
        self.assertEqual([m["font"] for m in data["missing_fonts"]], ["仿宋_GB2312", "楷体_GB2312", "方正小标宋简体"])
        self.assertIn("obtain", data["missing_fonts"][0]["guide"])
        self.assertIn("install_methods", data["font_notice"])
        code, out = self.run_cli(["--json"], FULL_FONTS)
        data = json.loads(out)
        self.assertEqual((code, data["font_status"], data["missing_fonts"], data["font_notice"]), (0, "complete", [], None))

    def test_runs_as_script(self):
        """作为脚本直接运行（只依赖标准库）：退出码只能是 0 或 3，--json 输出可解析。"""
        import subprocess
        import sys
        from tests.format.helpers import SCRIPTS
        r = subprocess.run([sys.executable, "-X", "utf8", str(SCRIPTS / "gw_fonts.py"), "--json"],
                           capture_output=True, text=True, encoding="utf-8")
        self.assertIn(r.returncode, (0, gw_fonts.EXIT_FONTS_INCOMPLETE), r.stderr)
        data = json.loads(r.stdout)
        self.assertIn(data["font_status"], ("complete", "missing", "undetected"))
        self.assertEqual(r.returncode == 0, data["font_status"] == "complete")


if __name__ == "__main__":
    unittest.main()
