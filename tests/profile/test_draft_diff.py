# -*- coding: utf-8 -*-
"""skills/gongwen-profile/scripts/draft_diff.py 的单元测试。

运行：python -m unittest discover -s tests/profile -t .

必过标本：两稿相同、只差空白、只差【起草说明】之后的内容、没有修订痕迹的 docx，必须报"无差异"。
必错标本：换词、删评价词、数字更正、删铺垫、调整顺序、拆句、合句、语气改为要求、改为百分点、
数据前移、修订痕迹和批注，必须逐项报出；修订人全名不得出现在输出中。
"""
from __future__ import annotations

import json
import unittest

from tests.profile.helpers import ScriptTestCase, make_docx

import draft_diff  # noqa: E402  （helpers 已把脚本目录加入 sys.path）

ORIGINAL = """某公司2026年三季度经营分析汇报

集团公司：

今年以来，在集团党委的坚强领导下，公司上下攻坚克难，各项工作取得了显著成效。现将三季度经营情况汇报如下。

一、主要指标

前三季度，公司实现营业收入12.3亿元，同比增长8.5%。毛利率18.6%，同比提高2.1%。

二、主要工作

（一）着力推进市场开拓。公司大力开拓省外市场，新签合同额4.2亿元。

（二）着力抓好降本增效。全面推行集中采购，采购成本下降3.2%。

（三）着力防范经营风险。全面开展应收账款清理。

三、下一步工作

我们将坚决贯彻集团部署，全力以赴完成全年目标任务。建议各部门可以考虑加强协同，尽量在11月底前完成专项清理。
"""

REVISED = """某公司2026年三季度经营分析汇报

集团公司：

现将三季度经营情况汇报如下。

一、主要指标

前三季度，公司实现营业收入12.8亿元，同比增长8.5%。毛利率18.6%，同比提高2.1个百分点。

二、主要工作

（一）重点推进市场开拓。新签合同额4.2亿元，主要来自省外市场。

（二）重点抓好降本增效。推行集中采购，采购成本下降3.2%。

（三）重点防范经营风险。开展应收账款清理。

三、下一步工作

各部门要加强协同，于11月底前完成专项清理。
"""

AUTHOR = ("王建军", "2026-10-14T08:30:00Z")
OTHER = ("刘敏", "2026-10-13T16:00:00Z")


def ins(text, who=AUTHOR):
    return ("ins", text, who[0], who[1])


def dele(text, who=AUTHOR):
    return ("del", text, who[0], who[1])


class DraftDiffBase(ScriptTestCase):
    module = draft_diff

    def diff_texts(self, a: str, b: str, *extra):
        pa = self.write("原稿.txt", a)
        pb = self.write("改稿.txt", b)
        code, out, err = self.run_main(str(pa), str(pb), "--json", *extra)
        self.assertEqual(code, 0, err)
        return json.loads(out)

    @staticmethod
    def items_of(rep, kind):
        return [x for x in rep["items"] if x["type"] == kind]

    @staticmethod
    def has_tag(rep, tag):
        return any(tag in x["tags"] for x in rep["items"])


class MustPassSpecimens(DraftDiffBase):
    """必过标本：没有实质改动的，必须报无差异。"""

    def test_identical(self):
        rep = self.diff_texts(ORIGINAL, ORIGINAL)
        self.assertTrue(rep["summary"]["identical"])
        self.assertEqual(rep["items"], [])

    def test_whitespace_and_line_endings_ignored(self):
        noisy = ORIGINAL.replace("\n\n", "\r\n\r\n\r\n").replace("公司", "公 司", 2)
        rep = self.diff_texts(ORIGINAL, noisy)
        self.assertTrue(rep["summary"]["identical"])

    def test_draft_note_boundary(self):
        a = ORIGINAL + "【起草说明】\n一、所作假设\n1. 按上行文处理。\n"
        b = ORIGINAL + "【起草说明】\n一、所作假设\n1. 改成别的说明。\n"
        rep = self.diff_texts(a, b)
        self.assertTrue(rep["summary"]["identical"])

    def test_docx_without_revisions(self):
        path = make_docx(self.tmp / "无修订.docx", [[("t", "某汇报")], [("t", "全年办件1.2万件。")]])
        code, out, _ = self.run_main(str(path), "--json")
        rep = json.loads(out)
        self.assertEqual(code, 0)
        self.assertTrue(rep["summary"]["identical"])
        self.assertTrue(any("未发现修订痕迹" in n for n in rep["notes"]))

    def test_docx_views_rebuild_both_drafts(self):
        path = make_docx(self.tmp / "修订.docx", [
            [("t", "今年以来，"), dele("在院党委的正确领导下，"), ("t", "门诊量"), dele("大幅"), ("t", "增长"),
             ins("12.3%"), ("t", "。")],
            [("moveFrom", "设备更新已完成。", *AUTHOR)],
            [("t", "下一步要加强管理。"), ("moveTo", "设备更新已完成。", *AUTHOR)],
        ])
        rev = draft_diff.DocxRevisions(path)
        original = [x for x in rev.lines("original") if x]
        revised = [x for x in rev.lines("revised") if x]
        self.assertEqual(original, ["今年以来，在院党委的正确领导下，门诊量大幅增长。", "设备更新已完成。", "下一步要加强管理。"])
        self.assertEqual(revised, ["今年以来，门诊量增长12.3%。", "下一步要加强管理。设备更新已完成。"])

    def test_mask_name(self):
        self.assertEqual(draft_diff.mask_name("王建军"), "王**")
        self.assertEqual(draft_diff.mask_name("刘敏"), "刘*")
        self.assertEqual(draft_diff.mask_name(""), "未署名")

    def test_two_files_with_unaccepted_revisions_noted(self):
        a = self.write("原稿.txt", "全年办件1.2万件。\n")
        b = make_docx(self.tmp / "改稿.docx", [[("t", "全年办件"), ins("1.3万件"), dele("1.2万件"), ("t", "。")]])
        code, out, _ = self.run_main(str(a), str(b), "--json")
        rep = json.loads(out)
        self.assertTrue(any("接受全部修订" in n for n in rep["notes"]))
        self.assertTrue(self.has_tag(rep, "数字改动"))

    def test_text_render(self):
        pa = self.write("原稿.txt", ORIGINAL)
        pb = self.write("改稿.txt", REVISED)
        code, out, _ = self.run_main(str(pa), str(pb))
        self.assertEqual(code, 0)
        self.assertIn("换词：着力→重点（3）", out)
        self.assertIn("不存入 .gongwen/", out)

    def test_identical_text_render(self):
        pa = self.write("原稿.txt", ORIGINAL)
        code, out, _ = self.run_main(str(pa), str(pa))
        self.assertIn("两稿正文没有差异", out)


class MustFlagSpecimens(DraftDiffBase):
    """必错标本：每类改动都要报出。"""

    def setUp(self):
        super().setUp()
        self.rep = self.diff_texts(ORIGINAL, REVISED)

    def test_word_swap_counted(self):
        swaps = {x["pair"]: x["count"] for x in self.rep["summary"]["word_swaps"]}
        self.assertEqual(swaps.get("着力→重点"), 3)

    def test_eval_words_removed(self):
        words = {x["word"]: x["count"] for x in self.rep["summary"]["removed_eval_words"]}
        for w in ("显著", "大力", "全面"):
            self.assertIn(w, words)
        self.assertEqual(words["全面"], 2)

    def test_number_correction_flagged(self):
        hits = [x for x in self.rep["items"] if "12.3" in x["before"] and "12.8" in x["after"]]
        self.assertEqual(len(hits), 1)
        self.assertIn("数字改动", hits[0]["tags"])

    def test_preamble_deleted(self):
        deleted = self.items_of(self.rep, "删除句")
        self.assertTrue(any("删铺垫" in x["tags"] and "坚强领导" in x["before"] for x in deleted))
        self.assertTrue(any("删表态" in x["tags"] for x in deleted))

    def test_tone_firmed_up(self):
        self.assertTrue(self.has_tag(self.rep, "语气改为要求"))

    def test_percent_to_percentage_point(self):
        self.assertTrue(self.has_tag(self.rep, "改为百分点"))

    def test_data_moved_forward(self):
        self.assertTrue(self.has_tag(self.rep, "数据前移"))

    def test_summary_numbers(self):
        s = self.rep["summary"]
        self.assertGreater(s["chars_before"], s["chars_after"])
        self.assertLess(s["change_ratio"], 0)
        self.assertFalse(s["identical"])

    def test_moved_sentence(self):
        a = "甲事项已经完成部署。\n乙事项正在推进之中。\n丙事项列入明年计划。\n"
        b = "丙事项列入明年计划。\n甲事项已经完成部署。\n乙事项正在推进之中。\n"
        rep = self.diff_texts(a, b)
        moved = self.items_of(rep, "调整顺序")
        self.assertEqual(len(moved), 1)
        self.assertEqual(self.items_of(rep, "删除句"), [])
        self.assertEqual(self.items_of(rep, "新增句"), [])

    def test_split_and_merge(self):
        a = "全年新建养老服务站12个并且全部投入运营，服务老人3600人次，群众满意度明显提高。\n第二件事。\n"
        b = "全年新建养老服务站12个，全部投入运营。服务老人3600人次。\n第二件事。\n"
        rep = self.diff_texts(a, b)
        self.assertEqual(len(self.items_of(rep, "拆句")), 1)
        rep2 = self.diff_texts(b, a)
        self.assertEqual(len(self.items_of(rep2, "合句")), 1)

    def test_rewritten_short_sentence_paired(self):
        a = "建议各部门可以考虑加强协同。\n"
        b = "各部门要加强协同，于11月底前完成年度目标任务。\n"
        rep = self.diff_texts(a, b)
        self.assertEqual(len(self.items_of(rep, "改写")), 1)

    def test_min_sim_controls_pairing(self):
        a = "积极参加各类行业展会，品牌影响力明显提升。\n"
        b = "参加行业展会5次。\n"
        loose = self.diff_texts(a, b, "--min-sim", "0.3")
        strict = self.diff_texts(a, b, "--min-sim", "0.95")
        self.assertEqual(len(self.items_of(loose, "改写")), 1)
        self.assertEqual(len(self.items_of(strict, "删除句")), 1)
        self.assertEqual(len(self.items_of(strict, "新增句")), 1)

    def test_docx_revisions_and_comments(self):
        path = make_docx(self.tmp / "院长修订稿.docx", [
            [("t", "护理部三季度工作汇报")],
            [("cstart", "1"), dele("今年以来，护理部高度重视质量管理，各项工作取得明显成效。"), ("cend", "1"),
             ins("三季度护理质量考核平均得分96.2分。"), ("t", "现将工作汇报如下。")],
            [("t", "（一）"), dele("着力"), ins("重点"), ("t", "加强质量管理。")],
            [("t", "二是"), dele("部分", OTHER), ins("3个", OTHER), ("t", "病区文书不规范。")],
        ], comments=[("1", AUTHOR[0], "2026-10-14T08:40:00Z", "开头太长，直接说结果。")])
        code, out, _ = self.run_main(str(path), "--json")
        self.assertEqual(code, 0)
        rep = json.loads(out)
        self.assertEqual([a["alias"] for a in rep["authors"]], ["修订人1", "修订人2"])
        self.assertEqual(rep["authors"][0]["masked_name"], "王**")
        self.assertEqual(rep["authors"][0]["dates"], ["2026-10-14", "2026-10-14"])
        self.assertEqual(rep["comments"][0]["text"], "开头太长，直接说结果。")
        self.assertIn("高度重视质量管理", rep["comments"][0]["anchor"])
        self.assertTrue(self.has_tag(rep, "删铺垫"))
        self.assertIn({"pair": "着力→重点", "count": 1}, rep["summary"]["word_swaps"])
        # 修订人全名不得出现在输出中
        self.assertNotIn("王建军", out)
        self.assertNotIn("刘敏", out)
        code, text_out, _ = self.run_main(str(path))
        self.assertNotIn("王建军", text_out)

    def test_author_filter(self):
        path = make_docx(self.tmp / "两人修订.docx", [
            [("t", "（一）"), dele("着力"), ins("重点"), ("t", "加强质量管理。")],
            [("t", "二是"), dele("部分", OTHER), ins("3个", OTHER), ("t", "病区文书不规范。")],
        ])
        code, out, _ = self.run_main(str(path), "--author", "1", "--json")
        rep = json.loads(out)
        self.assertEqual(rep["source"]["author"], "修订人1")
        befores = " ".join(x["before"] for x in rep["items"])
        self.assertIn("着力", befores)
        self.assertNotIn("部分病区", befores)
        code, out, _ = self.run_main(str(path), "--author", "2", "--json")
        rep = json.loads(out)
        befores = " ".join(x["before"] for x in rep["items"])
        self.assertIn("部分病区", befores)
        self.assertNotIn("着力", befores)

    def test_large_block_falls_back_without_error(self):
        a = "\n".join(f"第{i}项工作按计划推进，完成率达到百分之{i % 90 + 10}。" for i in range(520))
        b = "\n".join(f"第{i}项任务已经启动，目前进展顺利，预计年底完成。" for i in range(520))
        rep = self.diff_texts(a, b)
        self.assertFalse(rep["summary"]["identical"])

    def test_input_errors_return_2(self):
        txt = self.write("单个.txt", "内容。")
        code, _, err = self.run_main(str(txt))
        self.assertEqual(code, 2)
        self.assertIn("必须是带修订痕迹的 .docx", err)
        code, _, err = self.run_main(str(self.tmp / "不存在.docx"))
        self.assertEqual(code, 2)
        bad = self.write("坏.docx", "不是压缩包")
        code, _, err = self.run_main(str(bad))
        self.assertEqual(code, 2)
        a = self.write("a.txt", "甲。")
        b = self.write("b.txt", "乙。")
        code, _, err = self.run_main(str(a), str(b), "--author", "1")
        self.assertEqual(code, 2)
        code, _, err = self.run_main(str(a), str(b), "--min-sim", "1.5")
        self.assertEqual(code, 2)
        path = make_docx(self.tmp / "一人.docx", [[("t", "甲"), ins("乙"), ("t", "。")]])
        code, _, err = self.run_main(str(path), "--author", "3")
        self.assertEqual(code, 2)
        self.assertIn("超出范围", err)
        pdf = self.write("稿.pdf", "x")
        code, _, err = self.run_main(str(a), str(pdf))
        self.assertEqual(code, 2)

    def test_output_file_not_overwritten(self):
        pa = self.write("原稿.txt", ORIGINAL)
        pb = self.write("改稿.txt", REVISED)
        target = self.write("对比.md", "已有内容")
        code, out, _ = self.run_main(str(pa), str(pb), "--output", str(target))
        self.assertEqual(code, 0)
        self.assertEqual(target.read_text(encoding="utf-8"), "已有内容")
        self.assertTrue((self.tmp / "对比-v2.md").exists())


class NumberAndOrdinalTags(DraftDiffBase):
    """评审补充：数字类标签区分事实更正和概括改数字；层次序号变化不算换词，整节挪动判为调整顺序。"""

    def test_deleted_sentence_with_number_not_number_change(self):
        a = "全年完成检查120次。\n下一步要加强管理。\n"
        b = "下一步要加强管理。\n"
        rep = self.diff_texts(a, b)
        deleted = self.items_of(rep, "删除句")
        self.assertEqual(len(deleted), 1)
        self.assertIn("删除含数据句", deleted[0]["tags"])
        self.assertNotIn("数字改动", deleted[0]["tags"])

    def test_vague_word_to_number_not_number_change(self):
        a = "二是部分病区护理文书书写不规范。\n覆盖病区大幅增加。\n培训护士近千人。\n营业收入12.3亿元。\n"
        b = "二是3个病区护理文书书写不规范。\n覆盖病区由18个增至24个。\n培训护士932人次。\n营业收入12.8亿元。\n"
        rep = self.diff_texts(a, b)
        tags = {x["before"]: x["tags"] for x in rep["items"]}
        for before in ("二是部分病区护理文书书写不规范。", "覆盖病区大幅增加。", "培训护士近千人。"):
            with self.subTest(before=before):
                self.assertIn("改为具体数字", tags[before])
                self.assertNotIn("数字改动", tags[before])
        self.assertIn("数字改动", tags["营业收入12.3亿元。"])
        self.assertNotIn("改为具体数字", tags["营业收入12.3亿元。"])
        # 只把"%"改成"个百分点"的，数值没变，不算数字改动
        rep = self.diff_texts("毛利率同比提高2.1%。\n", "毛利率同比提高2.1个百分点。\n")
        self.assertEqual(rep["items"][0]["tags"], ["改为百分点"])

    def test_section_reorder_is_move_not_word_swap(self):
        a = ("一、主要成绩\n全年办结群众诉求1200件。\n新建便民服务点8个。\n"
             "二、存在问题\n部分诉求办理时限偏长。\n基层人手不足。\n")
        b = ("一、存在问题\n部分诉求办理时限偏长。\n基层人手不足。\n"
             "二、主要成绩\n全年办结群众诉求1200件。\n新建便民服务点8个。\n")
        rep = self.diff_texts(a, b)
        self.assertEqual(rep["summary"]["word_swaps"], [])
        self.assertFalse(any("换词" in x["tags"] for x in rep["items"]))
        moved = self.items_of(rep, "调整顺序")
        self.assertTrue(any(x["before"] == "二、存在问题" and x["after"] == "一、存在问题" for x in moved))
        self.assertEqual(self.items_of(rep, "删除句"), [])
        self.assertEqual(self.items_of(rep, "新增句"), [])

    def test_renumbering_only_not_counted(self):
        a = "一、总体情况\n全年办件1.2万件。\n二、下一步打算\n（一）加强管理。\n（二）完善制度。\n"
        b = "一、总体情况\n全年办件1.2万件。\n二、存在问题\n人手不足。\n三、下一步打算\n1.加强管理。\n2.完善制度。\n"
        rep = self.diff_texts(a, b)
        self.assertEqual(rep["summary"]["word_swaps"], [])
        self.assertEqual(rep["summary"]["renumbered"], 3)
        self.assertEqual(sorted(x["after"] for x in self.items_of(rep, "新增句")), ["二、存在问题", "人手不足。"])
        self.assertEqual(self.items_of(rep, "改写"), [])
        self.assertFalse(rep["summary"]["identical"])

    def test_decimal_number_not_treated_as_ordinal(self):
        self.assertEqual(draft_diff.strip_ordinal("1.2万件已办结。"), "1.2万件已办结。")
        self.assertEqual(draft_diff.strip_ordinal("（三）提升能力。"), "提升能力。")
        self.assertEqual(draft_diff.strip_ordinal("12.加强管理。"), "加强管理。")
        self.assertEqual(draft_diff.strip_ordinal("（一）"), "（一）")


class AuthorCommentFilter(DraftDiffBase):
    """评审补充：--author 时批注只列该修订人的；只写批注的人也编入修订人清单。"""

    def make(self):
        return make_docx(self.tmp / "两人批注.docx", [
            [("cstart", "1"), ("t", "（一）"), dele("着力"), ins("重点"), ("t", "加强质量管理。"), ("cend", "1")],
            [("cstart", "2"), ("t", "二是"), dele("部分", OTHER), ins("3个", OTHER), ("t", "病区文书不规范。"),
             ("cend", "2")],
            [("cstart", "3"), ("t", "下一步要加强管理。"), ("cend", "3")],
        ], comments=[("1", AUTHOR[0], "2026-10-14T08:40:00Z", "小标题再短一些。"),
                     ("2", OTHER[0], "2026-10-13T16:10:00Z", "核对数字。"),
                     ("3", "张三", "2026-10-15T09:00:00Z", "写具体时间。")])

    def test_comment_only_author_listed(self):
        code, out, _ = self.run_main(str(self.make()), "--json")
        rep = json.loads(out)
        rows = {r["alias"]: r for r in rep["authors"]}
        self.assertEqual(list(rows), ["修订人1", "修订人2", "修订人3"])
        self.assertEqual(rows["修订人3"]["revisions"], 0)
        self.assertEqual(rows["修订人3"]["comments"], 1)
        self.assertEqual(rows["修订人3"]["masked_name"], "张*")
        self.assertEqual([c["author"] for c in rep["comments"]], ["修订人1", "修订人2", "修订人3"])
        self.assertNotIn("张三", out)

    def test_author_filter_limits_comments(self):
        path = self.make()
        code, out, _ = self.run_main(str(path), "--author", "2", "--json")
        rep = json.loads(out)
        self.assertEqual([c["text"] for c in rep["comments"]], ["核对数字。"])
        self.assertTrue(any("其他修订人的批注 2 条未列出" in n for n in rep["notes"]))
        code, text, _ = self.run_main(str(path), "--author", "2")
        self.assertIn("只采用：修订人2的修订和批注", text)
        self.assertNotIn("小标题再短一些", text)
        code, out, _ = self.run_main(str(path), "--author", "3", "--json")
        rep = json.loads(out)
        self.assertEqual(code, 0)
        self.assertEqual([c["text"] for c in rep["comments"]], ["写具体时间。"])
        self.assertEqual(rep["items"], [])


class NestedRevisions(DraftDiffBase):
    """多人修订的嵌套结构：甲插入、乙又删掉的文字，原稿和改稿里都不应出现。"""

    def test_inserted_then_deleted_text_absent_from_both_views(self):
        path = make_docx(self.tmp / "嵌套.docx", [
            [("t", "全年营业收入12.8亿元"), ("insdel", "，创历史新高", AUTHOR[0], OTHER[0], AUTHOR[1]), ("t", "。")],
        ])
        rev = draft_diff.DocxRevisions(path)
        self.assertEqual([x for x in rev.lines("original") if x], ["全年营业收入12.8亿元。"])
        self.assertEqual([x for x in rev.lines("revised") if x], ["全年营业收入12.8亿元。"])
        code, out, _ = self.run_main(str(path), "--json")
        rep = json.loads(out)
        self.assertEqual(code, 0)
        self.assertTrue(rep["summary"]["identical"])
        self.assertNotIn("删评价词", out)

    def test_author_filter_on_nested_revision(self):
        path = make_docx(self.tmp / "嵌套筛选.docx", [
            [("t", "全年营业收入12.8亿元"), ("insdel", "，创历史新高", AUTHOR[0], OTHER[0], AUTHOR[1]), ("t", "。")],
        ])
        rev = draft_diff.DocxRevisions(path)
        # 只接受插入人（修订人1）的修订：插入保留，删除人的删除按拒绝处理
        self.assertEqual([x for x in rev.lines("revised", AUTHOR[0]) if x], ["全年营业收入12.8亿元，创历史新高。"])
        # 只接受删除人的修订：插入人的插入被拒绝，文字不出现
        self.assertEqual([x for x in rev.lines("revised", OTHER[0]) if x], ["全年营业收入12.8亿元。"])


class SecrecyMarks(DraftDiffBase):
    """必错标本：含密级或内部标注的，不输出对比清单，退出码 1；必过标本：正文提到保密工作的照常比对。"""

    def assert_blocked(self, code, out, *secrets):
        self.assertEqual(code, 1)
        for s in secrets:
            self.assertNotIn(s, out)

    def test_revision_docx_with_secret_mark(self):
        path = make_docx(self.tmp / "修订稿.docx", [
            [("t", "秘密★1年")],
            [("t", "某项目推进情况汇报")],
            [("t", "项目总投资"), dele("8.4亿元"), ins("8.6亿元"), ("t", "。")],
        ])
        code, out, _ = self.run_main(str(path), "--json")
        rep = json.loads(out)
        self.assert_blocked(code, out, "8.4亿元", "8.6亿元", "王建军")
        self.assertEqual(rep["blocked"], [{"file": "修订稿.docx", "marks": ["密级标识"]}])
        self.assertEqual(rep["items"], [])
        self.assertEqual(rep["comments"], [])
        code, text, _ = self.run_main(str(path))
        self.assert_blocked(code, text, "8.4亿元", "8.6亿元")
        self.assertIn("未比对", text)

    def test_internal_marks_in_text_files(self):
        for mark in ("内部资料，严禁外传", "（内部资料 不得翻印）", "此件为工作秘密，注意保管。", "机密（5年）",
                     "内部参阅", "以上经营数据均为未公开的财务数据。"):
            with self.subTest(mark=mark):
                pa = self.write("原稿.txt", mark + "\n营业收入23.6亿元。\n")
                pb = self.write("改稿.txt", "营业收入23.8亿元。\n")
                code, out, _ = self.run_main(str(pa), str(pb), "--json")
                rep = json.loads(out)
                self.assert_blocked(code, out, "23.6", "23.8")
                self.assertEqual(rep["blocked"][0]["file"], "原稿.txt")

    def test_mark_only_in_revised_file(self):
        pa = self.write("原稿.txt", "营业收入23.6亿元。\n")
        pb = self.write("改稿.txt", "内部资料\n营业收入23.8亿元。\n")
        code, out, _ = self.run_main(str(pa), str(pb), "--json")
        self.assert_blocked(code, out, "23.6", "23.8")
        self.assertEqual(json.loads(out)["blocked"][0]["file"], "改稿.txt")

    def test_mark_in_docx_header(self):
        path = make_docx(self.tmp / "页眉标注.docx", [
            [("t", "全年营业收入"), dele("23.6"), ins("23.8"), ("t", "亿元。")],
        ], header="内部资料　注意保存")
        code, out, _ = self.run_main(str(path), "--json")
        self.assert_blocked(code, out, "23.6", "23.8")

    def test_mark_in_docx_footer_two_file_mode(self):
        a = self.write("原稿.txt", "全年营业收入23.6亿元。\n")
        b = make_docx(self.tmp / "改稿.docx", [[("t", "全年营业收入23.8亿元。")]], footer="此件为内部资料，不得外传")
        code, out, _ = self.run_main(str(a), str(b), "--json")
        self.assert_blocked(code, out, "23.6", "23.8")

    def test_mark_inside_deleted_text(self):
        path = make_docx(self.tmp / "删去标注.docx", [
            [dele("内部资料")],
            [("t", "全年营业收入"), dele("23.6"), ins("23.8"), ("t", "亿元。")],
        ])
        code, out, _ = self.run_main(str(path), "--json")
        self.assert_blocked(code, out, "23.6", "23.8")

    def test_secrecy_words_in_body_not_blocked(self):
        a = "某局保密办工作总结\n全年开展保密工作检查4次，加强涉密人员管理。严守工作秘密。\n规范标注“秘密★1年”等密级标识。\n"
        b = "某局保密办工作总结\n全年开展保密检查4次，加强涉密人员管理。严守工作秘密。\n规范标注“秘密★1年”等密级标识。\n"
        rep = self.diff_texts(a, b)
        self.assertNotIn("blocked", rep)
        self.assertFalse(rep["summary"]["identical"])


if __name__ == "__main__":
    unittest.main()
