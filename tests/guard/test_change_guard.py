# -*- coding: utf-8 -*-
"""shared/scripts/change_guard.py 的单元测试：事实变化、有规则依据、无规则依据、结构变化。

运行：python -m unittest discover -s tests/guard -t .
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

import change_guard as cg  # noqa: E402

ORIG = """某区教育局关于做好2026年寒假安全工作的通知

各学校：

一、排查隐患

各校要在1月10日前完成校舍排查，建立整改台账。校舍排查涉及12栋教学楼。

二、落实值班

寒假期间实行24小时值班制度。联系人：王某某，电话0571-86660000。

三、报送材料

请于1月20日前将《寒假安全检查表》报区教育局办公室。

某区教育局
2026年12月20日
"""


def guard(original: str, revised: str, keep_structure: bool = False, doc_type=None):
    return cg.build_report(cg.Doc("原稿.txt", original), cg.Doc("改稿.txt", revised), keep_structure, doc_type)


def by_id(report, rid):
    return [c for c in report["changes"] if c["id"] == rid]


def lost_values(change):
    return [x["value"] for x in change["facts"]["lost"]]


def added_values(change):
    return [x["value"] for x in change["facts"]["added"]]


class IdenticalTest(unittest.TestCase):
    def test_identical(self):
        report = guard(ORIG, ORIG)
        self.assertEqual(report["changes"], [])
        self.assertEqual(report["structure"], [])
        self.assertEqual((report["summary"]["P0"], report["summary"]["P1"], report["summary"]["P2"]), (0, 0, 0))

    def test_whitespace_only_is_not_a_change(self):
        report = guard(ORIG, ORIG.replace("各校要在", "各校要在 ").replace("\n\n", "\n\n\n"))
        self.assertEqual(report["changes"], [])

    def test_draft_note_ignored(self):
        report = guard(ORIG + "\n【起草说明】\n依据2025年文件。\n", ORIG + "\n【起草说明】\n依据2024年文件。\n")
        self.assertEqual(report["changes"], [])


class FactChangeTest(unittest.TestCase):
    """必错标本：事实变化报 P0 CHG-001。"""

    def assert_fact(self, revised, kind=None, lost=None, added=None):
        report = guard(ORIG, revised)
        found = by_id(report, "CHG-001")
        self.assertTrue(found, report["changes"])
        self.assertEqual(found[0]["level"], "P0")
        self.assertEqual(report["summary"]["P0"], len(found))
        if kind:
            self.assertEqual(found[0]["kind"], kind)
        if lost:
            self.assertTrue(set(lost) <= set(lost_values(found[0])), found[0]["facts"])
        if added:
            self.assertTrue(set(added) <= set(added_values(found[0])), found[0]["facts"])
        return found

    def test_date_changed(self):
        self.assert_fact(ORIG.replace("1月10日前", "1月15日前"), "modify", ["1月10日"], ["1月15日"])

    def test_number_changed(self):
        self.assert_fact(ORIG.replace("涉及12栋", "涉及13栋"), "modify", ["12栋"], ["13栋"])

    def test_number_lost_by_deleting_sentence(self):
        found = self.assert_fact(ORIG.replace("校舍排查涉及12栋教学楼。", ""), "delete", ["12栋"])
        self.assertEqual(found[0]["revised"]["text"], "")

    def test_number_added(self):
        self.assert_fact(ORIG.replace("建立整改台账。", "建立整改台账，台账保存3年。"), "modify", None, ["3年"])

    def test_new_sentence_with_fact(self):
        found = self.assert_fact(ORIG.replace("校舍排查涉及12栋教学楼。", "校舍排查涉及12栋教学楼。全区共有85所学校。"),
                                 "add", None, ["85所"])
        self.assertEqual(found[0]["original"]["text"], "")

    def test_organization_changed(self):
        self.assert_fact(ORIG.replace("报区教育局办公室", "报区卫健局办公室"), "modify", ["区教育局"], ["区卫健局"])

    def test_person_changed(self):
        found = self.assert_fact(ORIG.replace("王某某", "李某某"), "modify")
        self.assertIn("王某某", lost_values(found[0]))

    def test_phone_changed(self):
        self.assert_fact(ORIG.replace("86660000", "86660001"), "modify")

    def test_book_title_changed(self):
        self.assert_fact(ORIG.replace("《寒假安全检查表》", "《安全检查表》"), "modify", ["《寒假安全检查表》"])

    def test_quote_changed(self):
        original = "会议强调要“守住安全底线”。\n"
        revised = "会议强调要“守牢安全底线”。\n"
        report = guard(original, revised)
        self.assertEqual([c["id"] for c in report["changes"]], ["CHG-001"])

    def test_file_name_changed(self):
        report = guard("详见附件《报表.xlsx》和台账表.xlsx。\n", "详见附件《报表.xlsx》和汇总表.xlsx。\n")
        found = by_id(report, "CHG-001")
        self.assertTrue(found)
        self.assertIn("台账表.xlsx", lost_values(found[0]))

    def test_marker_lost(self):
        report = guard("共投入资金【待补：金额】用于改造。\n", "共投入资金用于改造。\n")
        self.assertEqual([c["id"] for c in report["changes"]], ["CHG-001"])

    def test_docno_changed(self):
        report = guard("依据某政发〔2026〕12号文件执行。\n", "依据某政发〔2026〕21号文件执行。\n")
        self.assertEqual([c["id"] for c in report["changes"]], ["CHG-001"])


class FactPassTest(unittest.TestCase):
    """必过标本：事实不变的改写、搬动、合并不报 P0。"""

    def test_rewrite_keeps_facts(self):
        report = guard(ORIG, ORIG.replace("各校要在1月10日前完成校舍排查，建立整改台账。", "请各学校于1月10日前完成校舍排查并建立整改台账。"))
        self.assertEqual(report["summary"]["P0"], 0)
        self.assertEqual([c["id"] for c in report["changes"]], ["CHG-003"])

    def test_chinese_numeral_same_value(self):
        report = guard("共检查三家单位。\n", "累计检查3家单位。\n")
        self.assertEqual(report["summary"]["P0"], 0)

    def test_sentence_moved_with_facts(self):
        revised = ORIG.replace("各校要在1月10日前完成校舍排查，建立整改台账。校舍排查涉及12栋教学楼。",
                               "校舍排查涉及12栋教学楼。各校要在1月10日前完成校舍排查，建立整改台账。")
        report = guard(ORIG, revised)
        self.assertEqual(report["summary"]["P0"], 0)
        self.assertEqual(report["changes"], [])
        self.assertGreaterEqual(report["summary"]["moved"], 1)

    def test_fact_moved_to_another_paragraph(self):
        original = "一、情况\n\n全区共有85所学校。\n\n二、要求\n\n各校做好排查。\n"
        revised = "一、情况\n\n全区学校做好排查。\n\n二、要求\n\n全区共有85所学校，各校做好排查。\n"
        report = guard(original, revised)
        self.assertEqual(report["summary"]["P0"], 0, report["changes"])

    def test_sentences_merged(self):
        report = guard("共检查12家单位。发现问题5个。\n", "共检查12家单位，发现问题5个。\n")
        self.assertEqual(report["summary"]["P0"], 0)

    def test_abbreviation_is_p1_not_p0(self):
        report = guard("请报某区教育局办公室。\n", "请报区教育局办公室。\n")
        found = by_id(report, "CHG-001")
        self.assertEqual([c["level"] for c in found], ["P1"])
        self.assertEqual(report["summary"]["P0"], 0)
        self.assertEqual(report["summary"]["P1"], 1)

    def test_ordinal_words_not_facts(self):
        report = guard("一是加强领导。二是狠抓落实。\n", "一是强化领导。二是狠抓落实。\n")
        self.assertEqual(report["summary"]["P0"], 0)


class RuleBasisTest(unittest.TestCase):
    def test_residue_markdown_removed(self):
        original = "**强化组织领导**。成立工作专班，由局长任组长。\n"
        revised = "强化组织领导。成立工作专班，由局长任组长。\n"
        report = guard(original, revised)
        found = by_id(report, "CHG-002")
        self.assertEqual(len(found), 1, report["changes"])
        self.assertIn("RES-002", [r["id"] for r in found[0]["rules"]])
        self.assertEqual(found[0]["level"], "info")
        self.assertEqual(report["summary"]["ruled"], 1)
        self.assertEqual(report["summary"]["P0"] + report["summary"]["P1"] + report["summary"]["P2"], 0)

    def test_style_pattern_rewrite_is_ruled(self):
        original = "总的来说，我局全年工作稳步推进，取得了明显成效。\n"
        revised = "我局全年工作稳步推进。\n"
        report = guard(original, revised)
        self.assertEqual(report["summary"]["P0"], 0)
        self.assertTrue(by_id(report, "CHG-002"), report["changes"])

    def test_fact_change_also_records_rules(self):
        original = "**各校要在1月10日前完成排查**。\n"
        revised = "各校要在1月15日前完成排查。\n"
        report = guard(original, revised)
        found = by_id(report, "CHG-001")
        self.assertEqual(len(found), 1)
        self.assertIn("RES-002", [r["id"] for r in found[0]["rules"]])

    def test_unruled_rewrite(self):
        report = guard(ORIG, ORIG.replace("寒假期间实行24小时值班制度。", "寒假期间要实行24小时值班制度。"))
        found = by_id(report, "CHG-003")
        self.assertEqual([c["level"] for c in found], ["P2"])
        self.assertEqual(report["summary"]["P2"], 1)
        self.assertEqual(found[0]["rules"], [])

    def test_unruled_deletion(self):
        report = guard(ORIG, ORIG.replace("，建立整改台账。", "。"))
        found = by_id(report, "CHG-003")
        self.assertEqual(len(found), 1)
        self.assertEqual(found[0]["kind"], "modify")  # 同一句里删去一小段，仍是改写

    def test_unruled_whole_sentence_deleted(self):
        original = "一、情况\n\n今年工作平稳。要继续保持。\n"
        report = guard(original, "一、情况\n\n今年工作平稳。\n")
        found = by_id(report, "CHG-003")
        self.assertEqual([(c["kind"], c["original"]["text"]) for c in found], [("delete", "要继续保持。")])

    def test_unruled_addition(self):
        report = guard("今年工作平稳。\n", "今年工作平稳。明年继续保持。\n")
        found = by_id(report, "CHG-003")
        self.assertEqual([c["kind"] for c in found], ["add"])


class StructureTest(unittest.TestCase):
    def structure(self, revised):
        return guard(ORIG, revised, keep_structure=True)["structure"]

    def test_unchanged_structure(self):
        revised = ORIG.replace("请于1月20日前将", "请于1月20日前把")
        self.assertEqual(self.structure(revised), [])

    def test_not_checked_without_flag(self):
        revised = ORIG.replace("一、排查隐患", "（一）排查隐患")
        self.assertEqual(guard(ORIG, revised)["structure"], [])
        self.assertTrue(guard(ORIG, revised, keep_structure=True)["structure"])

    def test_title_changed(self):
        issues = self.structure(ORIG.replace("关于做好2026年寒假安全工作的通知", "关于做好寒假安全工作的通知"))
        self.assertEqual([i["message"] for i in issues], ["文章标题变化"])
        self.assertEqual(issues[0]["level"], "P1")

    def test_heading_number_changed(self):
        issues = self.structure(ORIG.replace("二、落实值班", "（二）落实值班"))
        self.assertTrue(any("层次序号变化" in i["message"] for i in issues), issues)

    def test_heading_text_changed(self):
        issues = self.structure(ORIG.replace("二、落实值班", "二、强化值班"))
        self.assertEqual([i["message"] for i in issues], ["层次标题文字变化"])

    def test_heading_removed(self):
        issues = self.structure(ORIG.replace("三、报送材料\n\n", ""))
        messages = "；".join(i["message"] for i in issues)
        self.assertIn("层次序号变化", messages)
        self.assertIn("段落数变化", messages)

    def test_paragraph_count_changed(self):
        issues = self.structure(ORIG.replace("各校要在1月10日前完成校舍排查，建立整改台账。校舍排查涉及12栋教学楼。",
                                             "各校要在1月10日前完成校舍排查，建立整改台账。\n\n校舍排查涉及12栋教学楼。"))
        self.assertEqual([i["message"] for i in issues], ["段落数变化（原稿 10 段，改稿 11 段）"])

    def test_paragraph_order_changed(self):
        revised = ORIG.replace(
            "寒假期间实行24小时值班制度。联系人：王某某，电话0571-86660000。\n\n三、报送材料\n\n请于1月20日前将《寒假安全检查表》报区教育局办公室。",
            "请于1月20日前将《寒假安全检查表》报区教育局办公室。\n\n三、报送材料\n\n寒假期间实行24小时值班制度。联系人：王某某，电话0571-86660000。")
        issues = self.structure(revised)
        self.assertTrue(any("段落顺序变化" in i["message"] for i in issues), issues)

    def test_structure_issue_counts_in_summary(self):
        report = guard(ORIG, ORIG.replace("二、落实值班", "二、强化值班"), keep_structure=True)
        self.assertEqual(report["summary"]["P1"], 1)


class CliTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="change_guard_test_"))
        (self.tmp / "原稿.txt").write_text(ORIG, encoding="utf-8")
        (self.tmp / "改稿.txt").write_text(ORIG.replace("1月10日前", "1月15日前"), encoding="utf-8")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def run_cli(self, *args):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = cg.main([str(a) for a in args])
        return code, out.getvalue(), err.getvalue()

    def test_text_output(self):
        code, out, _ = self.run_cli(self.tmp / "原稿.txt", self.tmp / "改稿.txt")
        self.assertEqual(code, 0)
        self.assertIn("CHG-001", out)
        self.assertIn("1月10日", out)
        self.assertIn("1月15日", out)

    def test_json_output(self):
        code, out, _ = self.run_cli(self.tmp / "原稿.txt", self.tmp / "改稿.txt", "--json")
        self.assertEqual(code, 0)
        data = json.loads(out)
        self.assertEqual(data["summary"]["P0"], 1)
        self.assertEqual(data["changes"][0]["id"], "CHG-001")
        self.assertEqual(data["changes"][0]["category"], "fact")
        for key in ("original_file", "revised_file", "doc_type", "summary", "changes", "structure", "notes"):
            self.assertIn(key, data)

    def test_strict_exit_code(self):
        self.assertEqual(self.run_cli(self.tmp / "原稿.txt", self.tmp / "改稿.txt", "--strict")[0], 1)
        self.assertEqual(self.run_cli(self.tmp / "原稿.txt", self.tmp / "原稿.txt", "--strict")[0], 0)

    def test_keep_structure_flag(self):
        (self.tmp / "改稿2.txt").write_text(ORIG.replace("二、落实值班", "二、强化值班"), encoding="utf-8")
        code, out, _ = self.run_cli(self.tmp / "原稿.txt", self.tmp / "改稿2.txt", "--keep-structure", "--json")
        data = json.loads(out)
        self.assertEqual(data["summary"]["P1"], 1)
        self.assertEqual(data["structure"][0]["id"], "CHG-004")
        _, out2, _ = self.run_cli(self.tmp / "原稿.txt", self.tmp / "改稿2.txt", "--json")
        self.assertEqual(json.loads(out2)["structure"], [])

    def test_missing_file(self):
        code, _, err = self.run_cli(self.tmp / "没有.txt", self.tmp / "改稿.txt")
        self.assertEqual(code, 2)
        self.assertIn("文件不存在", err)

    def test_unsupported_suffix(self):
        (self.tmp / "改稿.pdf").write_text("x", encoding="utf-8")
        code, _, err = self.run_cli(self.tmp / "原稿.txt", self.tmp / "改稿.pdf")
        self.assertEqual(code, 2)

    def test_registered_in_rules_json(self):
        data = json.loads((SCRIPTS_DIR / "lint_rules.json").read_text(encoding="utf-8"))
        self.assertEqual([r["id"] for r in data["change_guard"]], ["CHG-001", "CHG-002", "CHG-003", "CHG-004"])
        for item in data["change_guard"]:
            self.assertTrue(item["basis"])
            self.assertIn(item["level"], ("P0", "P1", "P2", "info"))
            self.assertEqual(cg.META[item["id"]]["level"], item["level"])


class SentenceSplitTest(unittest.TestCase):
    def test_quote_not_split(self):
        text = "他说：“要守住底线。要抓好落实。”会议结束。"
        spans = cg.split_sentences(text)
        self.assertEqual([text[s:e] for s, e in spans], ["他说：“要守住底线。要抓好落实。”会议结束。"])

    def test_basic(self):
        text = "第一句。第二句；第三句"
        self.assertEqual([text[s:e] for s, e in cg.split_sentences(text)], ["第一句。", "第二句；", "第三句"])


if __name__ == "__main__":
    unittest.main()
