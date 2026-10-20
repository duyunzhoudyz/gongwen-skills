# -*- coding: utf-8 -*-
"""“按单位规定指定标准字体”（--unit-fonts）：参数解析与校验、写入字体名、缺单位字体不回退、机检判定、meta 与命令行优先级。"""
import io
import json
import unittest
from contextlib import redirect_stdout

from tests.format.helpers import (FULL_FONTS, NO_FONTS, OUT, build_sample, build_text, fonts_installed,
                                  make_messy_docx, read_lines, written_fonts)

import build_docx  # noqa: E402  helpers 已把 scripts 加入 sys.path
import check_docx  # noqa: E402
import fix_docx  # noqa: E402
import gw_fonts  # noqa: E402
import preview_docx  # noqa: E402

UNIT = {"仿宋": "方正仿宋_GBK", "小标宋": "方正小标宋_GBK"}
UNIT_TEXT = "仿宋=方正仿宋_GBK,小标宋=方正小标宋_GBK"
UNIT_INSTALLED = FULL_FONTS | {"方正仿宋_gbk", "方正小标宋_gbk"}


def ids(result):
    return {i["id"] for i in result["issues"]}


def run_main(module, argv):
    buf = io.StringIO()
    with redirect_stdout(buf):
        code = module.main(argv)
    return code, buf.getvalue()


class ParseTest(unittest.TestCase):
    def setUp(self):
        self.params = gw_fonts.load_params()

    def test_parse_string_and_dict(self):
        self.assertEqual(gw_fonts.parse_unit_fonts(UNIT_TEXT, self.params), UNIT)
        self.assertEqual(gw_fonts.parse_unit_fonts(" 仿宋 ＝ 方正仿宋_GBK ； 小标宋：方正小标宋_GBK ", self.params), UNIT)
        self.assertEqual(gw_fonts.parse_unit_fonts(UNIT, self.params), UNIT)
        self.assertEqual(gw_fonts.parse_unit_fonts(None, self.params), {})

    def test_invalid(self):
        for bad in ("宋=方正仿宋_GBK", "仿宋=", "仿宋", "", "  ", "仿宋=a,仿宋=b", {"仿宋": ""}, {"仿宋": None}, ["仿宋=a"]):
            with self.assertRaises(ValueError, msg=repr(bad)):
                gw_fonts.parse_unit_fonts(bad, self.params)

    def test_apply_does_not_mutate(self):
        out = gw_fonts.apply_unit_fonts(self.params, UNIT)
        self.assertEqual(out["fonts"]["standard"]["仿宋"], "方正仿宋_GBK")
        self.assertEqual(self.params["fonts"]["standard"]["仿宋"], "仿宋_GB2312")
        self.assertEqual(out["fonts"]["unit_override"]["仿宋"], {"unit_font": "方正仿宋_GBK", "standard_font": "仿宋_GB2312"})
        self.assertIs(gw_fonts.apply_unit_fonts(self.params, {}), self.params)

    def test_missing_guide_is_generic_and_no_fallback(self):
        params = gw_fonts.apply_unit_fonts(self.params, UNIT)
        _fonts, check = gw_fonts.resolve_fonts(params, FULL_FONTS, "测试")
        self.assertEqual(check["font_status"], "missing")
        self.assertEqual([e["font"] for e in check["missing_fonts"]], ["方正仿宋_GBK", "方正小标宋_GBK"])
        text = gw_fonts.render_notice(check)
        self.assertIn("信息中心", text)
        self.assertIn("完全一致", text)
        self.assertNotIn("foundertype", text)  # 不套用国标字体的预置获取途径
        self.assertEqual(check["unit_fonts"]["仿宋"]["standard_font"], "仿宋_GB2312")
        self.assertIn("按单位规定使用 方正仿宋_GBK（国标为 仿宋_GB2312）", "".join(gw_fonts.unit_notes(check)))

    def test_gw_fonts_cli(self):
        with fonts_installed(UNIT_INSTALLED):
            code, out = run_main(gw_fonts, ["--unit-fonts", UNIT_TEXT, "--json"])
        self.assertEqual(code, 0)
        data = json.loads(out)
        self.assertEqual(data["fonts"]["仿宋"], "方正仿宋_GBK")
        self.assertIn("仿宋", data["unit_fonts"])
        with fonts_installed(FULL_FONTS):
            self.assertEqual(run_main(gw_fonts, ["--unit-fonts", UNIT_TEXT])[0], gw_fonts.EXIT_FONTS_INCOMPLETE)
        code, out = run_main(gw_fonts, ["--unit-fonts", "宋=x"])
        self.assertEqual(code, 2)
        self.assertIn("无法识别", out)


class BuildTest(unittest.TestCase):
    def test_written_names_and_report(self):
        path, report = build_sample("红头", "单位字体-红头.docx", installed=UNIT_INSTALLED, meta_update={"unit_fonts": UNIT})
        fonts = written_fonts(path)
        self.assertIn("方正仿宋_GBK", fonts)
        self.assertIn("方正小标宋_GBK", fonts)
        self.assertNotIn("仿宋_GB2312", fonts)
        self.assertNotIn("方正小标宋简体", fonts)
        self.assertEqual(report["font_status"], "complete")
        self.assertEqual(report["fonts"]["仿宋"], "方正仿宋_GBK")
        self.assertEqual(report["unit_fonts"]["仿宋"], {"unit_font": "方正仿宋_GBK", "standard_font": "仿宋_GB2312"})
        self.assertIn("按单位规定使用 方正仿宋_GBK（国标为 仿宋_GB2312）", build_docx.render_text(report))
        self.assertFalse(report["warnings"] and any("未识别的键" in w for w in report["warnings"]))

    def test_no_unit_fonts_field_empty(self):
        _path, report = build_sample("白头", "单位字体-无.docx", installed=FULL_FONTS)
        self.assertEqual(report["unit_fonts"], {})

    def test_missing_unit_font_reports_and_does_not_fall_back(self):
        # 国标字体齐全，但没装单位指定字体：照样报缺失，文件写入单位字体名，不回退到国标字体
        path, report = build_sample("红头", "单位字体-缺失.docx", installed=FULL_FONTS, meta_update={"unit_fonts": UNIT})
        self.assertEqual(report["font_status"], "missing")
        self.assertEqual({e["font"] for e in report["missing_fonts"]}, {"方正仿宋_GBK", "方正小标宋_GBK"})
        self.assertIn("信息中心", F_notice(report))
        fonts = written_fonts(path)
        self.assertIn("方正仿宋_GBK", fonts)
        self.assertNotIn("仿宋_GB2312", fonts)

    def test_undetected_lists_unit_fonts(self):
        _path, report = build_sample("白头", "单位字体-未检测.docx", installed=NO_FONTS, meta_update={"unit_fonts": {"仿宋": "方正仿宋_GBK"}})
        self.assertEqual(report["font_status"], "undetected")
        self.assertIn("方正仿宋_GBK", [e["font"] for e in report["unverified_fonts"]])
        self.assertNotIn("仿宋_GB2312", [e["font"] for e in report["unverified_fonts"]])

    def test_meta_alias_key_and_cli_priority(self):
        lines = read_lines("白头请示-正文.txt")
        meta = {"unit_fonts": {"仿宋": "方正仿宋_GBK", "小标宋": "方正小标宋_GBK"}}
        doc, report = build_docx.build(lines, meta, installed=NO_FONTS, unit_fonts="仿宋=华康仿宋体")
        self.assertEqual(report["fonts"]["仿宋"], "华康仿宋体")  # 命令行优先
        self.assertEqual(report["fonts"]["小标宋"], "方正小标宋_GBK")  # 其余角色沿用版式信息
        doc, report = build_docx.build(lines, {"单位规定字体": "仿宋=方正仿宋_GBK"}, installed=NO_FONTS)
        self.assertEqual(report["fonts"]["仿宋"], "方正仿宋_GBK")
        _doc, report = build_docx.build(lines, {"unit_fonts": ""}, installed=NO_FONTS)
        self.assertEqual(report["unit_fonts"], {})

    def test_invalid_meta_raises_build_error(self):
        lines = read_lines("白头请示-正文.txt")
        for bad in ({"宋": "x"}, {"仿宋": " "}, "仿宋"):
            with self.assertRaises(build_docx.BuildError):
                build_docx.build(lines, {"unit_fonts": bad}, installed=NO_FONTS)

    def test_cli_json_and_readback_check(self):
        out = OUT / "单位字体-命令行.docx"
        if out.exists():
            out.unlink()
        src = OUT / "单位字体-正文.txt"
        src.write_text("\n".join(read_lines("白头请示-正文.txt")), encoding="utf-8")
        with fonts_installed(UNIT_INSTALLED):
            code, text = run_main(build_docx, [str(src), "-o", str(out), "--unit-fonts", UNIT_TEXT, "--json"])
        self.assertEqual(code, 0)
        data = json.loads(text)
        self.assertEqual(data["unit_fonts"]["小标宋"]["unit_font"], "方正小标宋_GBK")
        self.assertEqual(data["check"]["unit_fonts"], data["unit_fonts"])  # 回读检查收到同样的参数
        self.assertNotIn("DOCX-016", ids(data["check"]))
        self.assertNotIn("DOCX-010", ids(data["check"]))

    def test_cli_bad_args_exit_2(self):
        src = OUT / "单位字体-正文2.txt"
        src.write_text("\n".join(read_lines("白头请示-正文.txt")), encoding="utf-8")
        for bad in ("宋=x", "仿宋=", "仿宋"):
            with fonts_installed(FULL_FONTS):
                code, text = run_main(build_docx, [str(src), "--unit-fonts", bad, "-o", str(OUT / "单位字体-错误.docx")])
            self.assertEqual(code, 2, bad)
            self.assertIn("错误", text)


def F_notice(report):
    return gw_fonts.render_notice(report)


class CheckTest(unittest.TestCase):
    def test_unit_font_is_compliant_and_standard_font_is_not(self):
        unit_doc, _ = build_sample("红头", "单位字体-机检-单位.docx", installed=NO_FONTS, meta_update={"unit_fonts": UNIT})
        std_doc, _ = build_sample("红头", "单位字体-机检-国标.docx", installed=NO_FONTS)
        ok = check_docx.check_document(unit_doc, layout="红头", unit_fonts=UNIT_TEXT)
        self.assertNotIn("DOCX-016", ids(ok))
        self.assertNotIn("DOCX-010", ids(ok))
        self.assertNotIn("DOCX-020", ids(ok))
        self.assertEqual(ok["unit_fonts"]["仿宋"]["unit_font"], "方正仿宋_GBK")
        self.assertTrue(any("按单位规定使用字体" in n for n in ok["notes"]))
        # 不传参数：单位字体按非标准字体报 DOCX-016
        self.assertIn("DOCX-016", ids(check_docx.check_document(unit_doc, layout="红头")))
        # 传参数但文档用国标字体：同样报 DOCX-016，且消息指向单位字体
        bad = check_docx.check_document(std_doc, layout="红头", unit_fonts=UNIT)
        hits = [i for i in bad["issues"] if i["id"] == "DOCX-016"]
        self.assertTrue(hits)
        self.assertTrue(any("方正仿宋_GBK" in h["message"] for h in hits))
        self.assertTrue(any("方正小标宋_GBK" in h["message"] for h in hits))
        self.assertEqual(check_docx.check_document(std_doc, layout="红头")["unit_fonts"], {})

    def test_unit_font_with_unrecognised_class_still_compliant(self):
        # 单位字体名里没有“仿宋”等关键词时，按单位规定使用同样视为合规
        path, _ = build_text("关于某事的通知\n\n各单位：\n\n现通知如下。\n\n某市某局\n\n2026年3月5日", "单位字体-自定义名.docx",
                             meta={"版式": "白头", "unit_fonts": {"仿宋": "单位专用体", "小标宋": "单位标题体"}})
        res = check_docx.check_document(path, unit_fonts={"仿宋": "单位专用体", "小标宋": "单位标题体"})
        self.assertFalse({"DOCX-010", "DOCX-016", "DOCX-020"} & ids(res))

    def test_cli_exit_codes(self):
        path, _ = build_sample("白头", "单位字体-机检命令行.docx", installed=NO_FONTS, meta_update={"unit_fonts": UNIT})
        code, text = run_main(check_docx, [str(path), "--unit-fonts", UNIT_TEXT, "--json"])
        self.assertEqual(code, 0)
        self.assertNotIn("DOCX-016", ids(json.loads(text)))
        for bad in ("宋=x", "仿宋="):
            code, text = run_main(check_docx, [str(path), "--unit-fonts", bad])
            self.assertEqual(code, 2)
            self.assertIn("错误", text)
        with self.assertRaises(ValueError):
            check_docx.check_document(path, unit_fonts="宋=x")


class FixTest(unittest.TestCase):
    def test_fix_writes_unit_fonts_and_report(self):
        src = make_messy_docx(OUT / "单位字体-混乱.docx")
        dst = OUT / "单位字体-混乱-修复.docx"
        if dst.exists():
            dst.unlink()
        report = fix_docx.fix_document(src, dst, installed=FULL_FONTS, unit_fonts=UNIT)
        self.assertEqual(report["font_status"], "missing")  # 没装单位字体：报缺失，不回退
        self.assertIn("方正仿宋_GBK", {e["font"] for e in report["missing_fonts"]})
        fonts = written_fonts(dst)
        self.assertIn("方正仿宋_GBK", fonts)
        self.assertNotIn("仿宋_GB2312", fonts)
        self.assertIn("按单位规定使用", fix_docx.render_report(report))
        res = check_docx.check_document(dst, unit_fonts=UNIT)
        self.assertNotIn("DOCX-016", ids(res))

    def test_fix_cli(self):
        src = make_messy_docx(OUT / "单位字体-混乱2.docx")
        dst = OUT / "单位字体-混乱2-修复.docx"
        rep = OUT / "单位字体-混乱2-修复报告.md"
        for f in (dst, rep):
            if f.exists():
                f.unlink()
        with fonts_installed(UNIT_INSTALLED):
            code, text = run_main(fix_docx, [str(src), "-o", str(dst), "--unit-fonts", UNIT_TEXT, "--json"])
        self.assertEqual(code, 0)
        data = json.loads(text)
        self.assertEqual(data["unit_fonts"]["仿宋"]["unit_font"], "方正仿宋_GBK")
        self.assertEqual(data["font_status"], "complete")
        self.assertNotIn("DOCX-016", ids(data["check"]))
        code, text = run_main(fix_docx, [str(src), "--unit-fonts", "宋=x"])
        self.assertEqual(code, 2)
        self.assertIn("错误", text)


class PreviewTest(unittest.TestCase):
    def test_font_check_uses_unit_fonts(self):
        src, _ = build_sample("白头", "单位字体-预览.docx", installed=NO_FONTS)
        with fonts_installed(FULL_FONTS):  # 国标字体齐全，但缺单位字体：拒绝渲染
            code, text = run_main(preview_docx, [str(src), "--unit-fonts", UNIT_TEXT])
        self.assertEqual(code, gw_fonts.EXIT_FONTS_INCOMPLETE)
        self.assertIn("方正仿宋_GBK", text)
        with fonts_installed(UNIT_INSTALLED):
            self.assertEqual(preview_docx.check_fonts(UNIT_TEXT), 0)
        with fonts_installed(FULL_FONTS):
            code, text = run_main(preview_docx, [str(src), "--unit-fonts", "宋=x"])
        self.assertEqual(code, 2)


if __name__ == "__main__":
    unittest.main()
