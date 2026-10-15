# -*- coding: utf-8 -*-
"""shared/scripts/fact_diff.py 的单元测试。

运行：python -m unittest discover -s tests/fact_diff -t .
"""
from __future__ import annotations

import io
import json
import shutil
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
SCRIPTS_DIR = REPO_ROOT / "shared" / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import fact_diff  # noqa: E402

try:
    import docx  # type: ignore  # noqa: F401
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


class FactDiffTestCase(unittest.TestCase):

    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="fact_diff_test_"))

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp, ignore_errors=True)

    def write(self, name: str, text: str, encoding: str = "utf-8") -> Path:
        path = self.tmp / name
        path.write_bytes(text.encode(encoding))
        return path

    def run_main(self, *args: str):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = fact_diff.main(list(args))
        return code, out.getvalue(), err.getvalue()

    def diff(self, original: str, revised: str) -> dict:
        a = self.write("原稿.txt", original)
        b = self.write("改稿.txt", revised)
        code, out, _err = self.run_main(str(a), str(b), "--json")
        self.assertEqual(code, 0)
        return json.loads(out)

    @staticmethod
    def values(entries):
        return [(item["type"], item["value"]) for item in entries]


class NumberTest(FactDiffTestCase):

    def test_number_change_reported_in_both_lists(self):
        report = self.diff(
            "全年平均办件时长由3.5天压缩至1.2天，满意度达到98.6%。",
            "全年平均办件时长由3.5天压缩至1天，满意度达到98.6%。",
        )
        self.assertEqual(self.values(report["missing"]), [("数字", "1.2天")])
        self.assertEqual(self.values(report["added"]), [("数字", "1天")])
        self.assertFalse(report["consistent"])

    def test_unit_change_is_a_fact_change(self):
        report = self.diff("同比提高2.3个百分点。", "同比提高2.3%。")
        self.assertEqual(self.values(report["missing"]), [("数字", "2.3个百分点")])
        self.assertEqual(self.values(report["added"]), [("数字", "2.3%")])

    def test_amount_and_date_change(self):
        report = self.diff(
            "投入资金66万元，于2026年10月16日前完成。",
            "投入资金60万元，于2026年10月26日前完成。",
        )
        self.assertIn(("数字", "66万元"), self.values(report["missing"]))
        self.assertIn(("日期", "2026年10月16日"), self.values(report["missing"]))
        self.assertIn(("数字", "60万元"), self.values(report["added"]))
        self.assertIn(("日期", "2026年10月26日"), self.values(report["added"]))

    def test_occurrence_count_is_compared(self):
        report = self.diff(
            "一季度新增12家。三季度又新增12家。",
            "一季度和三季度各新增家数相同。",
        )
        self.assertEqual(report["missing"][0]["value"], "12家")
        self.assertEqual(report["missing"][0]["original_count"], 2)
        self.assertEqual(report["missing"][0]["revised_count"], 0)

    def test_partial_occurrence_loss(self):
        report = self.diff("新增12家，其中12家已投产。", "新增12家，均已投产。")
        self.assertEqual(self.values(report["missing"]), [("数字", "12家")])
        self.assertEqual(report["missing"][0]["revised_count"], 1)

    def test_equivalent_forms_not_reported(self):
        report = self.diff(
            "覆盖十二个乡镇、三百五十项事项、1,200余家企业，百分之三十的事项实现网办，"
            "二〇二六年十一月完成，10月16号前上报，会议9：30召开，完成第三季度目标。",
            "覆盖12个乡镇、350项事项、1200余家企业，30%的事项实现网办，"
            "2026年11月完成，10月16日前上报，会议9:30召开，完成第3季度目标。",
        )
        self.assertTrue(report["consistent"], report)

    def test_percent_with_space_equivalent(self):
        report = self.diff("增长3.5 %。", "增长3.5%。")
        self.assertTrue(report["consistent"], report)

    def test_generic_chinese_numerals_ignored(self):
        report = self.diff(
            "统一部署，进一步推进，建立一个窗口，第一时间报告。",
            "统筹部署，持续推进，设立综合窗口，立即报告。",
        )
        self.assertTrue(report["consistent"], report)


class MarkerTest(FactDiffTestCase):

    def test_marker_lost(self):
        report = self.diff(
            "开展专项检查【待补：次数】次，整改率【待核实：以台账为准】。",
            "开展专项检查，整改率【待核实：以台账为准】。",
        )
        self.assertEqual(self.values(report["missing"]), [("标记", "【待补：次数】")])
        self.assertEqual(report["added"], [])

    def test_marker_text_changed(self):
        report = self.diff("全年接诊【待补：人次】人次。", "全年接诊【待补】人次。")
        self.assertEqual(self.values(report["missing"]), [("标记", "【待补：人次】")])
        self.assertEqual(self.values(report["added"]), [("标记", "【待补】")])


class TitleQuoteDocnoTest(FactDiffTestCase):

    def test_title_content_changed(self):
        report = self.diff(
            "根据《关于深化放管服改革的实施意见》要求，制定本方案。",
            "根据《关于深化放管服改革的意见》要求，制定本方案。",
        )
        self.assertEqual(self.values(report["missing"]), [("书名号", "《关于深化放管服改革的实施意见》")])
        self.assertEqual(self.values(report["added"]), [("书名号", "《关于深化放管服改革的意见》")])

    def test_numbers_inside_title_not_double_reported(self):
        report = self.diff("印发《2026年工作要点》。", "印发《2025年工作要点》。")
        kinds = {item["type"] for item in report["missing"] + report["added"]}
        self.assertEqual(kinds, {"书名号"})

    def test_quote_removed(self):
        report = self.diff("推行“最多跑一次”改革。", "推行最多跑一次改革。")
        self.assertEqual(self.values(report["missing"]), [("引号", "“最多跑一次”")])

    def test_docno_change_and_context_insensitive(self):
        report = self.diff(
            "根据某政发〔2026〕12号文件要求。",
            "按照某政发〔2026〕21号文件要求。",
        )
        self.assertEqual(self.values(report["missing"]), [("发文字号", "某政发〔2026〕12号")])
        self.assertEqual(self.values(report["added"]), [("发文字号", "某政发〔2026〕21号")])
        same = self.diff("根据某政发〔2026〕12号文件要求。", "按照某政发〔2026〕12号文件要求。")
        self.assertTrue(same["consistent"], same)

    def test_docno_without_agency_code(self):
        report = self.diff("根据〔2026〕12号文件要求。", "按照〔2026〕12号文件要求。")
        self.assertTrue(report["consistent"], report)


class SequenceTest(FactDiffTestCase):

    def test_sequence_renumbering_not_reported(self):
        original = (
            "一、主要工作\n\n（一）政务服务\n\n1. 建立清单，共12项。\n\n（1）完成培训3次。\n\n"
            "主要做法：1.建立台账；2.完善制度；（3）加强考核。\n\n"
            "第一，压实责任；第二，强化督导。"
        )
        revised = (
            "二、主要工作\n\n（三）政务服务\n\n3. 建立清单，共12项。\n\n（2）完成培训3次。\n\n"
            "主要做法：（1）建立台账；（2）完善制度；3.加强考核。\n\n"
            "一是压实责任；二是强化督导。"
        )
        report = self.diff(original, revised)
        self.assertTrue(report["consistent"], report)

    def test_markdown_heading_sequence_not_reported(self):
        report = self.diff("## 1. 工作成效\n\n**（二）**保障措施", "一、工作成效\n\n（一）保障措施")
        self.assertTrue(report["consistent"], report)


class BoundaryTest(FactDiffTestCase):

    def test_text_after_drafting_note_ignored(self):
        original = "全年完成12项。\n【起草说明】\n一、修改说明\n1. 原稿3.5天。"
        revised = "全年完成12项。\n\n【起草说明】\n一、修改说明\n1. 将4天改为5天，删除【待补：次数】。"
        report = self.diff(original, revised)
        self.assertTrue(report["consistent"], report)
        self.assertEqual(report["original_lines"], 1)

    def test_text_before_drafting_note_still_compared(self):
        report = self.diff("完成12项。\n【起草说明】\n无", "完成13项。\n【起草说明】\n无")
        self.assertEqual(self.values(report["missing"]), [("数字", "12项")])


class RealisticRewriteTest(FactDiffTestCase):

    def test_style_only_rewrite_is_consistent(self):
        original = (
            "值得注意的是，通过一年努力，全院门急诊量达到45.6万人次，同比增长8.2%，"
            "这不仅仅是数字的提升，更是服务能力的飞跃——四级手术占比提高至21.3%。"
        )
        revised = "全院门急诊量45.6万人次，同比增长8.2%；四级手术占比提高至21.3%。"
        report = self.diff(original, revised)
        self.assertTrue(report["consistent"], report)


class CliTest(FactDiffTestCase):

    def test_strict_exit_codes(self):
        a = self.write("a.txt", "完成12项。")
        b = self.write("b.txt", "完成13项。")
        c = self.write("c.txt", "共完成12项。")
        self.assertEqual(self.run_main(str(a), str(b), "--strict")[0], 1)
        self.assertEqual(self.run_main(str(a), str(b))[0], 0)
        self.assertEqual(self.run_main(str(a), str(c), "--strict")[0], 0)

    def test_text_output_has_two_lists(self):
        a = self.write("a.txt", "完成12项。")
        b = self.write("b.txt", "完成13项。")
        code, out, _ = self.run_main(str(a), str(b))
        self.assertEqual(code, 0)
        self.assertIn("一、原稿有而改稿没有（1 项）", out)
        self.assertIn("二、改稿新增（1 项）", out)
        self.assertIn("12项", out)

    def test_input_errors_return_2(self):
        a = self.write("a.txt", "完成12项。")
        code, _out, err = self.run_main(str(a), str(self.tmp / "不存在.txt"))
        self.assertEqual(code, 2)
        self.assertIn("文件不存在", err)
        pdf = self.write("b.pdf", "x")
        self.assertEqual(self.run_main(str(a), str(pdf))[0], 2)
        self.assertEqual(self.run_main(str(a))[0], 2)

    def test_md_and_gb18030_input(self):
        a = self.write("a.md", "完成12项，投入5万元。")
        b = self.write("b.txt", "完成12项，投入5万元。", encoding="gb18030")
        code, out, err = self.run_main(str(a), str(b), "--json")
        self.assertEqual(code, 0)
        self.assertTrue(json.loads(out)["consistent"])
        self.assertIn("GB18030", err)

    @unittest.skipUnless(HAS_DOCX, "未安装 python-docx")
    def test_docx_input(self):
        import docx as docx_module  # type: ignore
        document = docx_module.Document()
        document.add_paragraph("全年完成12项，投入66万元。")
        table = document.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "检查次数"
        table.cell(0, 1).text = "8次"
        document.add_paragraph("【起草说明】")
        document.add_paragraph("原稿为13项。")
        path = self.tmp / "改稿.docx"
        document.save(str(path))
        a = self.write("原稿.txt", "全年完成12项，投入66万元。\n检查次数\t9次")
        code, out, _ = self.run_main(str(a), str(path), "--json")
        self.assertEqual(code, 0)
        report = json.loads(out)
        self.assertEqual(self.values(report["missing"]), [("数字", "9次")])
        self.assertEqual(self.values(report["added"]), [("数字", "8次")])


class ReviewRegressionTest(FactDiffTestCase):
    """代码审查发现的缺陷回归测试。"""

    def test_docno_context_words_not_reported(self):
        # 原问题：发文字号前的“学习”“对照”“按”等上下文被并入代字，改稿只改了上下文也报“发文字号”增删
        pairs = (
            ("认真学习某政发〔2026〕12号文件精神。", "深入学习某政发〔2026〕12号文件精神。"),
            ("请各单位对照某政发〔2026〕12号文件开展自查。", "各单位要对照某政发〔2026〕12号文件开展自查。"),
            ("按某政发〔2026〕12号文件要求执行。", "依照某政发〔2026〕12号文件要求执行。"),
            ("传达某政发〔2026〕12号文件。", "以某政发〔2026〕12号文件为依据。"),
        )
        for original, revised in pairs:
            with self.subTest(original=original):
                report = self.diff(original, revised)
                self.assertTrue(report["consistent"], report)

    def test_docno_agency_change_still_reported(self):
        report = self.diff("认真学习京政发〔2026〕12号文件精神。", "深入学习沪政发〔2026〕12号文件精神。")
        self.assertEqual(self.values(report["missing"]), [("发文字号", "京政发〔2026〕12号")])
        self.assertEqual(self.values(report["added"]), [("发文字号", "沪政发〔2026〕12号")])

    def test_chinese_fraction_compared(self):
        # 原问题：“三分之一”改成“四分之一”不报差异
        report = self.diff("超过三分之一的村完成改造，千分之五的抽检不合格。", "超过四分之一的村完成改造，千分之五的抽检不合格。")
        self.assertEqual(self.values(report["missing"]), [("数字", "1/3")])
        self.assertEqual(self.values(report["added"]), [("数字", "1/4")])
        self.assertTrue(self.diff("三分之二以上", "三分之二以上")["consistent"])

    def test_markdown_drafting_note_boundary(self):
        # 原问题：“## 【起草说明】”“**【起草说明】**”不截断，起草说明里的数字参与比对（gongwen-format 已截断）
        report = self.diff("全年完成12项。\n## 【起草说明】\n原稿为13项。", "全年完成12项。\n**【起草说明】**\n改为14项。")
        self.assertTrue(report["consistent"], report)
        self.assertEqual(report["original_lines"], 1)

    @unittest.skipUnless(HAS_DOCX, "未安装 python-docx")
    def test_docx_merged_cells_counted_once(self):
        # 原问题：python-docx 对合并单元格逐格重复返回，横向合并 3 格的“100万元”被计为 3 处
        import docx as docx_module  # type: ignore
        document = docx_module.Document()
        table = document.add_table(rows=3, cols=3)
        table.cell(0, 0).merge(table.cell(0, 2)).text = "全年预算100万元"
        table.cell(1, 0).merge(table.cell(2, 0)).text = "参与单位3家"
        table.cell(1, 1).text = "一季度"
        table.cell(1, 2).text = "25万元"
        table.cell(2, 1).text = "二季度"
        table.cell(2, 2).text = "30万元"
        path = self.tmp / "原稿.docx"
        document.save(str(path))
        txt = self.write("改稿.txt", "全年预算100万元\n参与单位3家\t一季度\t25万元\n二季度\t30万元")
        code, out, _ = self.run_main(str(path), str(txt), "--json")
        self.assertEqual(code, 0)
        report = json.loads(out)
        self.assertTrue(report["consistent"], report)


class NormalizationTest(FactDiffTestCase):
    """只改写法、数值不变的，不报差异（端到端实测：review 直接改后的复检）。"""

    def test_date_zero_padding(self):
        report = self.diff("各学校于2026年11月05日前完成自查。", "各学校于2026年11月5日前完成自查。")
        self.assertTrue(report["consistent"], report)
        self.assertTrue(self.diff("截至2026-01-05。", "截至2026-1-5。")["consistent"])

    def test_range_unit_completion(self):
        for original, revised in (("可安排经费3～5万元。", "可安排经费3万～5万元。"),
                                  ("增幅控制在15～30%。", "增幅控制在15%～30%。"),
                                  ("2026—2028年实现全覆盖。", "2026年—2028年实现全覆盖。")):
            with self.subTest(original=original):
                self.assertTrue(self.diff(original, revised)["consistent"])

    def test_range_value_change_still_reported(self):
        report = self.diff("可安排经费3～5万元。", "可安排经费3万～6万元。")
        self.assertEqual(self.values(report["missing"]), [("数字", "5万元")])
        self.assertEqual(self.values(report["added"]), [("数字", "6万元")])

    def test_non_range_numbers_untouched(self):
        report = self.diff("联系电话：0571-87654321。", "联系电话：0571-87654321。")
        self.assertTrue(report["consistent"])
        self.assertIn(("数字", "571"), [(f.kind, f.value) for f in fact_diff.extract_line("0571-87654321", 1)])


LEDGER = """# 素材台账：某科室2026年工作总结

## 二、台账表

| 序号 | 类别 | 内容 | 来源 | 状态 |
|---|---|---|---|---|
| 1 | 数据 | 2026年1—11月完成PCI 612台 | 材料1第5段 | 已确认 |
| 2 | 数据 | 出院随访430人次 | 材料1第6段 | 已确认 |
| 3 | 数据 | 现有医生18人、护士36人 | 材料2第3行 | 已确认 |
"""


class LedgerModeTest(FactDiffTestCase):
    """--ledger：素材台账.md 对正文.txt，只列正文中有、台账中没有的事实，按集合比较。"""

    def ledger(self, body: str, *extra: str):
        a = self.write("素材台账.md", LEDGER)
        b = self.write("正文.txt", body)
        return self.run_main(str(a), str(b), "--ledger", *extra)

    def test_only_added_by_set(self):
        body = ("全年完成PCI 612台，出院随访430人次。PCI 612台中急诊占一半。\n"
                "现有医生18人、护士36人，新增设备2台。")
        code, out, _ = self.ledger(body, "--json")
        self.assertEqual(code, 0)
        report = json.loads(out)
        self.assertEqual(report["mode"], "ledger")
        self.assertNotIn("missing", report)
        self.assertEqual(self.values(report["added"]), [("数字", "2台")])
        self.assertFalse(report["consistent"])

    def test_ignore_seq_and_source_columns(self):
        lines = fact_diff.ledger_lines(LEDGER.split("\n"))
        joined = "\n".join(lines)
        self.assertNotIn("材料1第5段", joined)
        self.assertIn("类别\t内容\t状态", lines)
        self.assertIn("数据\t2026年1—11月完成PCI 612台\t已确认", lines)
        self.assertEqual(len(lines), len(LEDGER.split("\n")), "行号保持不变")
        # 正文中的“第5段”“第3行”式写法不能被台账的来源列覆盖
        code, out, _ = self.ledger("详见材料1第5段。", "--json")
        self.assertEqual(self.values(json.loads(out)["added"]), [("数字", "1"), ("数字", "5段")])

    def test_strict_and_text_output(self):
        code, out, _ = self.ledger("全年完成PCI 612台。", "--strict")
        self.assertEqual(code, 0)
        self.assertIn("正文中有、台账中没有（0 项）", out)
        self.assertNotIn("原稿有而改稿没有", out)
        code, out, _ = self.ledger("全年完成PCI 613台。", "--strict")
        self.assertEqual(code, 1)
        self.assertIn("613台", out)


class LongPathTest(unittest.TestCase):

    @unittest.skipUnless(sys.platform.startswith("win"), "只在 Windows 下加长路径前缀")
    def test_long_path_prefixed_on_windows(self):
        long_path = Path("C:/" + "a" * 250 + "/原稿.txt")
        self.assertTrue(str(fact_diff._long_path(long_path)).startswith("\\\\?\\"))
        short_path = Path("C:/temp/原稿.txt")
        self.assertEqual(fact_diff._long_path(short_path), short_path)


class CnToIntTest(unittest.TestCase):

    def test_values(self):
        cases = {
            "十": 10, "十五": 15, "二十": 20, "三百零五": 305, "三百五": 350, "一千二百": 1200,
            "一万二千": 12000, "两千": 2000, "二〇二六": 2026,
        }
        for text, value in cases.items():
            self.assertEqual(fact_diff.cn_to_int(text), value, text)

    def test_invalid(self):
        for text in ("十四五", "三四", "七一", "一一", "百"):
            self.assertIsNone(fact_diff.cn_to_int(text), text)


if __name__ == "__main__":
    unittest.main()
