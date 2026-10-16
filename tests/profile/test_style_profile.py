# -*- coding: utf-8 -*-
"""skills/gongwen-profile/scripts/style_profile.py 的单元测试。

运行：python -m unittest discover -s tests/profile -t .

必过标本：规范、已脱敏的历史稿，统计值与手算一致，不报隐私、不跳过。
必错标本：含身份证号、手机号、姓名年龄组合、密级标识、"内部资料"标注的历史稿，必须报出或跳过，且输出中不得出现原文号码。
"""
from __future__ import annotations

import json
import unittest

from tests.profile.helpers import ScriptTestCase, make_docx

import style_profile  # noqa: E402  （helpers 已把脚本目录加入 sys.path）

CLEAN_REPORT = """关于某项工作情况的报告

某市人民政府：

全年完成检查12次，整改问题30个。现将有关情况报告如下。

一、主要做法

（一）完善制度。制定办法3项。

一是建立台账；二是每月调度。

特此报告。

某局
2026年10月16日
"""

SUMMARY_A = """某局2025年工作总结

2025年，在市委、市政府的坚强领导下，我局统筹推进各项工作。现将有关情况总结如下。

一、聚焦主责主业　推动提质增效

（一）优化审批。统筹推进审批改革，办理时长压缩至3.1个工作日。要统筹推进窗口整合，扎实开展减材料行动。

（二）强化监管。扎实开展专项检查，统筹推进信用监管。

二、存在的问题

一是部分事项办理时限偏长；二是基层执法力量不足。
"""

SUMMARY_B = """某局2026年上半年工作总结

今年以来，我局统筹推进营商环境优化。上半年网办事项1.12万件，网办率83.6%。

一、聚焦改革攻坚　推动服务提质

（一）深化改革。扎实开展系统整合，平均用时2.5个工作日。

（二）守牢底线。扎实开展食品安全检查，统筹推进隐患排查。

二、下半年工作打算

一是持续深化改革；二是扎实开展专项整治。
"""


class StyleProfileBase(ScriptTestCase):
    module = style_profile

    def report(self, *names_texts, extra_args=()):
        paths = [str(self.write(name, text)) for name, text in names_texts]
        code, out, err = self.run_main(*paths, "--json", *extra_args)
        return code, json.loads(out), out


class MustPassSpecimens(StyleProfileBase):
    """必过标本：规范稿件，统计与手算一致。"""

    def test_structure_counts_match_hand_count(self):
        code, rep, _ = self.report(("报告.txt", CLEAN_REPORT))
        self.assertEqual(code, 0)
        doc = rep["documents"][0]
        self.assertTrue(doc["title"])
        # 段落：首段、"制定办法3项。"、"一是……"、"特此报告。"；主送、落款、日期不计
        self.assertEqual(doc["paragraphs"], 4)
        # 句子：首段 2 句，其余各 1 句
        self.assertEqual(doc["sentences"], 5)
        self.assertEqual(doc["headings"]["1"], 1)
        self.assertEqual(doc["headings"]["2"], 1)
        self.assertEqual(doc["yishi"]["count"], 2)
        self.assertEqual(doc["yishi"]["max_enum"], 2)
        self.assertEqual(doc["opening"], "数据式")
        self.assertEqual(doc["ending"], "公文结束语")
        self.assertEqual(rep["privacy"]["by_file"], [])
        self.assertEqual(rep["sample"]["skipped"], [])

    def test_sentence_length_excludes_terminal_punctuation(self):
        code, rep, _ = self.report(("短.txt", "一、标题\n\n甲乙丙丁。戊己。\n"))
        self.assertEqual(code, 0)
        stats = rep["documents"][0]["sentence_len"]
        self.assertEqual(stats["min"], 2)
        self.assertEqual(stats["max"], 4)

    def test_secrecy_words_in_body_not_skipped(self):
        text = "某局保密办2026年工作总结\n\n全年开展保密工作检查4次，加强涉密人员管理。严守工作秘密，不在非涉密计算机上处理工作秘密。\n"
        code, rep, _ = self.report(("保密办总结.txt", text))
        self.assertEqual(code, 0)
        self.assertEqual(rep["sample"]["counted"], 1)
        self.assertEqual(rep["sample"]["skipped"], [])

    def test_masked_personal_info_not_flagged(self):
        text = ("情况说明\n\n患者张某，男，56岁，住院号已隐去。联系电话138****0000。"
                "身份证号990102********0011。王某同志参加会议。\n")
        code, rep, _ = self.report(("已脱敏.txt", text))
        self.assertEqual(code, 0)
        self.assertEqual(rep["privacy"]["by_file"], [])

    def test_draft_note_boundary(self):
        text = "某总结\n\n一是加强学习。\n【起草说明】\n一是这里不统计。二是也不统计。\n"
        code, rep, _ = self.report(("带说明.txt", text))
        self.assertEqual(rep["documents"][0]["yishi"]["count"], 1)

    def test_four_char_phrases_cross_document_and_filtered(self):
        code, rep, _ = self.report(("a.txt", SUMMARY_A), ("b.txt", SUMMARY_B))
        phrases = {x["phrase"]: x for x in rep["summary"]["four_char_phrases"]}
        self.assertIn("统筹推进", phrases)
        self.assertIn("扎实开展", phrases)
        self.assertEqual(phrases["统筹推进"]["docs"], 2)
        # 数字后的计量单位不是四字格；长短语截出来的片段不应出现
        self.assertNotIn("个工作日", phrases)
        self.assertNotIn("筹推进各", phrases)

    def test_preamble_opening_detected(self):
        code, rep, _ = self.report(("a.txt", SUMMARY_A))
        self.assertEqual(rep["documents"][0]["opening"], "铺垫式")

    def test_parallel_heading_style(self):
        code, rep, _ = self.report(("a.txt", SUMMARY_A), ("b.txt", SUMMARY_B))
        styles = rep["summary"]["heading"]["styles_level1_2"]
        self.assertEqual(styles.get("对仗式"), 2)

    def test_no_original_sentences_in_output(self):
        code, _rep, raw = self.report(("a.txt", SUMMARY_A), ("b.txt", SUMMARY_B))
        self.assertNotIn("办理时长压缩至3.1个工作日", raw)
        self.assertNotIn("聚焦主责主业", raw)  # 标题原文默认不输出
        code, out, _ = self.run_main(str(self.tmp / "a.txt"), str(self.tmp / "b.txt"))
        self.assertNotIn("办理时长压缩至3.1个工作日", out)

    def test_examples_flag_outputs_headings_only_on_request(self):
        a = self.write("a.txt", SUMMARY_A)
        code, out, _ = self.run_main(str(a), "--json", "--examples")
        rep = json.loads(out)
        self.assertIn("examples", rep["summary"]["heading"])
        self.assertIn("聚焦主责主业　推动提质增效", rep["summary"]["heading"]["examples"]["1"])

    def test_docx_input_and_table_excluded(self):
        path = make_docx(self.tmp / "稿.docx", [
            [("t", "某局工作汇报")],
            [("t", "全年办件1.2万件。"), ("del", "取得显著成效。", "某人", "2026-10-01T00:00:00Z")],
            [("t", "一是优化流程；二是压缩时限。")],
        ], table_rows=[["指标", "数值"], ["办件量", "1.2万件"]])
        code, out, _ = self.run_main(str(path), "--json")
        rep = json.loads(out)
        self.assertEqual(code, 0)
        doc = rep["documents"][0]
        self.assertEqual(doc["paragraphs"], 2)  # 表格行不计入；删除的修订内容不计入
        self.assertEqual(doc["sentences"], 2)

    def test_gb18030_text_is_read(self):
        path = self.write("国标编码.txt", "某总结\n\n一是加强学习；二是改进作风。\n", encoding="gb18030")
        code, out, _ = self.run_main(str(path), "--json")
        rep = json.loads(out)
        self.assertEqual(code, 0)
        self.assertTrue(any("GB18030" in n for n in rep["notes"]))

    def test_directory_input_and_dedup(self):
        sub = self.tmp / "历史稿"
        sub.mkdir()
        (sub / "a.txt").write_text(SUMMARY_A, encoding="utf-8")
        (sub / "b.txt").write_text(SUMMARY_B, encoding="utf-8")
        (sub / "说明.pdf").write_text("x", encoding="utf-8")
        code, out, _ = self.run_main(str(sub), str(sub / "a.txt"), "--json")
        rep = json.loads(out)
        self.assertEqual(rep["sample"]["files"], 2)

    def test_output_file_not_overwritten(self):
        a = self.write("a.txt", SUMMARY_A)
        target = self.write("文风统计.md", "已有内容")
        code, out, _ = self.run_main(str(a), "--output", str(target))
        self.assertEqual(code, 0)
        self.assertEqual(target.read_text(encoding="utf-8"), "已有内容")
        self.assertTrue((self.tmp / "文风统计-v2.md").exists())

    def test_text_render_has_sections(self):
        a = self.write("a.txt", SUMMARY_A)
        b = self.write("b.txt", SUMMARY_B)
        code, out, _ = self.run_main(str(a), str(b))
        self.assertEqual(code, 0)
        for heading in ("一、每篇概况", "二、量化参数", "三、高频表述"):
            self.assertIn(heading, out)


class MustFlagSpecimens(StyleProfileBase):
    """必错标本：隐私和密级必须报出，且不泄露原文号码。"""

    PRIVATE = ("关于高龄津贴发放对象的公示\n\n"
               "一、李秀兰，女，82岁，身份证号990102194403150028，住址某小区5幢1单元302室。\n"
               "二、陈德明，男，86岁。\n"
               "联系电话：13800002222。经研究，推荐刘建华同志参加评选。\n")

    def test_personal_info_counted_not_echoed(self):
        code, rep, raw = self.report(("公示.txt", self.PRIVATE))
        self.assertEqual(code, 0)
        hits = rep["privacy"]["by_file"][0]["hits"]
        self.assertEqual(hits.get("身份证号"), 1)
        self.assertEqual(hits.get("手机号"), 1)
        self.assertEqual(hits.get("姓名＋性别＋年龄"), 2)
        self.assertEqual(hits.get("家庭住址"), 1)
        self.assertEqual(hits.get("疑似姓名（××同志）"), 1)
        for secret in ("990102194403150028", "13800002222", "李秀兰", "陈德明", "刘建华"):
            self.assertNotIn(secret, raw)

    def test_personal_info_not_in_text_output(self):
        path = self.write("公示.txt", self.PRIVATE)
        code, out, _ = self.run_main(str(path))
        self.assertIn("隐私提示", out)
        self.assertNotIn("13800002222", out)

    def test_repeated_name_not_extracted_as_phrase(self):
        text = "".join(f"第{i}段，患者王建国，男，56岁，复查正常。\n" for i in range(5))
        code, rep, raw = self.report(("a.txt", text), ("b.txt", text))
        self.assertNotIn("王建国", raw)

    def test_secret_mark_skips_document(self):
        secret = "秘密★1年\n\n某项目情况报告\n\n项目投资8.4亿元。\n"
        code, rep, raw = self.report(("涉密.txt", secret), ("a.txt", SUMMARY_A))
        self.assertEqual(code, 0)
        self.assertEqual(rep["sample"]["counted"], 1)
        self.assertEqual(rep["sample"]["skipped"][0]["file"], "涉密.txt")
        self.assertNotIn("8.4亿元", raw)

    def test_internal_material_mark_skips_document(self):
        internal = "内部资料　注意保存\n\n某公司经营情况报告\n\n营业收入23.6亿元。\n"
        code, rep, raw = self.report(("内部.txt", internal))
        self.assertEqual(code, 1)  # 没有可统计的文件
        self.assertEqual(rep["sample"]["counted"], 0)
        self.assertNotIn("23.6", raw)

    def test_more_internal_mark_variants_skip_document(self):
        for mark in ("内部资料，严禁外传", "此件为工作秘密，注意保管。", "（内部资料 不得翻印）", "机密（5年）",
                     "内部参阅", "内部传阅", "上述经营数据为未公开的财务数据，请注意保密。"):
            with self.subTest(mark=mark):
                text = mark + "\n\n某公司经营情况报告\n\n营业收入23.6亿元，二期融资已完成。\n"
                code, rep, raw = self.report(("标注.txt", text))
                self.assertEqual(code, 1)
                self.assertEqual(rep["sample"]["skipped"][0]["file"], "标注.txt")
                self.assertTrue(rep["sample"]["skipped"][0]["marks"])
                self.assertNotIn("23.6", raw)
                self.assertNotIn("二期融资", raw)

    def test_docx_header_mark_skips_document(self):
        path = make_docx(self.tmp / "页眉标注.docx", [[("t", "某公司经营情况报告")], [("t", "营业收入23.6亿元。")]],
                         header="内部资料　注意保存")
        code, out, _ = self.run_main(str(path), "--json")
        rep = json.loads(out)
        self.assertEqual(code, 1)
        self.assertEqual(rep["sample"]["skipped"][0]["marks"], ["内部标注"])
        self.assertNotIn("23.6", out)

    def test_docx_footer_secret_mark_skips_document(self):
        path = make_docx(self.tmp / "页脚密级.docx", [[("t", "某项目情况报告")], [("t", "项目投资8.4亿元。")]],
                         footer="秘密★1年")
        code, out, _ = self.run_main(str(path), "--json")
        self.assertEqual(code, 1)
        self.assertNotIn("8.4", out)

    def test_small_sample_note(self):
        code, rep, _ = self.report(("a.txt", SUMMARY_A))
        self.assertTrue(any("少于 3 篇" in n for n in rep["notes"]))

    def test_large_sample_note(self):
        items = [(f"{i}.txt", SUMMARY_A) for i in range(11)]
        code, rep, _ = self.report(*items)
        self.assertTrue(any("多于 10 篇" in n for n in rep["notes"]))

    def test_input_errors_return_2(self):
        code, _, err = self.run_main(str(self.tmp / "不存在.txt"))
        self.assertEqual(code, 2)
        self.assertIn("文件不存在", err)
        pdf = self.write("稿.pdf", "x")
        code, _, err = self.run_main(str(pdf))
        self.assertEqual(code, 2)
        bad = self.write("坏.docx", "不是压缩包")
        code, _, err = self.run_main(str(bad))
        self.assertEqual(code, 2)
        code, _, _ = self.run_main(str(pdf), "--top", "0")
        self.assertEqual(code, 2)


if __name__ == "__main__":
    unittest.main()
