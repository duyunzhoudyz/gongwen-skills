# -*- coding: utf-8 -*-
"""函数类检查逐项测试（结构、日期、主送、隐私、篇幅、标记等）。"""
from __future__ import annotations

import unittest

from tests.lint.helpers import gl, ids, issues_of, run

NOTICE_HEAD = "某市某局关于做好某项工作的通知\n\n各区某局：\n\n"


class QuoteAndColonTest(unittest.TestCase):
    def test_pun009_nested_double_quotes(self):
        r = run("会议提出“落实“三个区分开来”要求”。", checks="punct")
        self.assertIn("PUN-009", ids(r))

    def test_pun009_single_quote_inside_ok(self):
        r = run("会议提出“落实‘三个区分开来’要求”。", checks="punct")
        self.assertNotIn("PUN-009", ids(r))

    def test_pun014_nested_colon(self):
        r = run("会议要求：各单位要做到：按时报送。", checks="punct")
        self.assertIn("PUN-014", ids(r))

    def test_pun014_field_labels_and_time_ok(self):
        r = run("联系人：张某，联系电话：0571-88880000。会议于9：30开始。", checks="punct")
        self.assertNotIn("PUN-014", ids(r))

    def test_pun014_marker_colon_ok(self):
        r = run("全年发表论文【待补：篇数】篇，其中核心期刊【待补：篇数】篇。", checks="punct,style")
        self.assertNotIn("PUN-014", ids(r))
        self.assertNotIn("AI-013", ids(r))


class DateTest(unittest.TestCase):
    def test_num010_mixed_styles(self):
        r = run("于2026年10月16日启动，2026-11-30前完成。", checks="number")
        self.assertIn("NUM-010", ids(r))

    def test_num010_dotted(self):
        r = run("检查于2026.10.16启动。", checks="number")
        self.assertIn("NUM-010", ids(r))

    def test_num010_uniform_ok(self):
        r = run("于2026年10月16日启动，2026年11月30日前完成。", checks="number")
        self.assertNotIn("NUM-010", ids(r))

    def test_fmt005_padding_signoff_is_p0(self):
        text = NOTICE_HEAD + "请按时报送。\n\n某市某局\n2026年01月05日\n"
        found = issues_of(run(text, "通知", "format"), "FMT-005")
        self.assertEqual([i["level"] for i in found], ["P0"])
        self.assertIn("2026年1月5日", found[0]["suggestion"])

    def test_fmt005_padding_body_is_p1(self):
        text = NOTICE_HEAD + "请于2026年01月05日前报送。\n\n某市某局\n2026年1月5日\n"
        found = issues_of(run(text, "通知", "format"), "FMT-005")
        self.assertEqual([i["level"] for i in found], ["P1"])

    def test_fmt005_month_day_only(self):
        found = issues_of(run("请于10月05日前报送。", checks="format"), "FMT-005")
        self.assertEqual(len(found), 1)

    def test_fmt006_chinese_signoff_date(self):
        text = NOTICE_HEAD + "请按时报送。\n\n某市某局\n二〇二六年十月十六日\n"
        self.assertIn("FMT-006", ids(run(text, "通知", "format")))

    def test_fmt006_arabic_ok(self):
        text = NOTICE_HEAD + "请按时报送。\n\n某市某局\n2026年10月16日\n"
        self.assertNotIn("FMT-006", ids(run(text, "通知", "format")))


class ZhusongTest(unittest.TestCase):
    def test_fmt007_missing_colon(self):
        text = "某市某局关于做好某项工作的通知\n\n各区某局\n\n请按时报送。\n"
        self.assertIn("FMT-007", ids(run(text, "通知", "format")))

    def test_fmt007_halfwidth_colon(self):
        text = "某市某局关于做好某项工作的通知\n\n各区某局:\n\n请按时报送。\n"
        r = run(text, "通知", "format,punct")
        self.assertIn("FMT-007", ids(r))
        self.assertNotIn("PUN-012", ids(r), "与 FMT-007 重叠时 PUN-012 应让位")

    def test_fmt007_ok(self):
        self.assertNotIn("FMT-007", ids(run(NOTICE_HEAD + "请按时报送。\n", "通知", "format")))

    def test_fmt007_not_for_speech(self):
        text = "在全市安全生产工作会议上的讲话\n\n同志们\n\n今天召开会议。\n"
        self.assertNotIn("FMT-007", ids(run(text, "讲话稿", "format")))


class HeadingSequenceTest(unittest.TestCase):
    def test_skip_number(self):
        text = NOTICE_HEAD + "一、总体要求\n\n内容。\n\n三、工作要求\n\n内容。\n"
        found = issues_of(run(text, "通知", "structure"), "STR-001")
        self.assertTrue(any(i["level"] == "P1" and "跳号" in i["message"] for i in found))

    def test_repeat_number(self):
        text = NOTICE_HEAD + "一、总体要求\n\n（一）甲。\n\n（一）乙。\n"
        found = issues_of(run(text, "通知", "structure"), "STR-001")
        self.assertTrue(any("重复" in i["message"] for i in found))

    def test_level_jump_is_p2(self):
        text = NOTICE_HEAD + "一、总体要求\n\n1. 甲。\n\n2. 乙。\n"
        found = issues_of(run(text, "通知", "structure"), "STR-001")
        self.assertEqual([i["level"] for i in found], ["P2"])

    def test_sequence_ok(self):
        text = NOTICE_HEAD + "一、总体要求\n\n（一）甲。\n\n1. 子项。\n\n2. 子项。\n\n（二）乙。\n\n二、工作要求\n\n（一）丙。\n"
        self.assertNotIn("STR-001", ids(run(text, "通知", "structure")))

    def test_restart_after_colon_intro(self):
        text = "某学院公示\n\n一等助学金：\n\n1. 甲\n\n2. 乙\n\n二等助学金：\n\n1. 丙\n"
        self.assertNotIn("STR-001", ids(run(text, "其他", "structure")))

    def test_continue_numbering_across_colon_groups(self):
        text = "某中学工作记录\n\n德育处：\n\n1. 甲。\n\n2. 乙。\n\n教务处：\n\n3. 丙。\n\n4. 丁。\n"
        self.assertNotIn("STR-001", ids(run(text, "其他", "structure")))

    def test_attachment_list_ignored(self):
        text = NOTICE_HEAD + "一、总体要求\n\n内容。\n\n附件：1. 检查表\n      2. 汇总表\n\n某市某局\n2026年10月16日\n"
        self.assertNotIn("STR-001", ids(run(text, "通知", "structure")))

    def test_heading_of(self):
        self.assertEqual(gl.heading_of("一、总体要求"), (1, 1))
        self.assertEqual(gl.heading_of("（十二）其他"), (2, 12))
        self.assertEqual(gl.heading_of("3. 子项"), (3, 3))
        self.assertEqual(gl.heading_of("（4）细项"), (4, 4))
        self.assertIsNone(gl.heading_of("1.5亿元已拨付"))
        self.assertIsNone(gl.heading_of("一、二年级学生"))


class CountSummaryTest(unittest.TestCase):
    def test_mismatch_with_headings(self):
        text = NOTICE_HEAD + "一、主要任务\n\n重点抓好三个方面工作：\n\n（一）甲。\n\n（二）乙。\n\n（三）丙。\n\n（四）丁。\n\n二、工作要求\n\n内容。\n"
        found = issues_of(run(text, "通知", "structure"), "STR-002")
        self.assertEqual(len(found), 1)
        self.assertEqual(found[0]["level"], "P2")
        self.assertIn("4", found[0]["message"])

    def test_mismatch_inline(self):
        r = run("存在三个方面问题：一是台账不全；二是数据滞后。", checks="structure")
        self.assertIn("STR-002", ids(r))

    def test_match_ok(self):
        r = run("存在三个方面问题：一是台账不全；二是数据滞后；三是培训不足。", checks="structure")
        self.assertNotIn("STR-002", ids(r))

    def test_heading_summary_ok(self):
        text = NOTICE_HEAD + "一、着力抓好两项重点工作\n\n（一）甲。\n\n（二）乙。\n\n二、工作要求\n\n内容。\n"
        self.assertNotIn("STR-002", ids(run(text, "通知", "structure")))

    def test_quoted_term_ignored(self):
        r = run("深入开展“三个方面”专项行动：一是甲；二是乙。", checks="structure")
        self.assertNotIn("STR-002", ids(r))


class DocTypeRulesTest(unittest.TestCase):
    QINGSHI_OK = "某区教育局关于申请校舍维修经费的请示\n\n区人民政府：\n\n恳请安排维修经费386万元。\n\n妥否，请批示。\n\n某区教育局\n2026年10月16日\n"

    def test_str003_missing_ending(self):
        text = self.QINGSHI_OK.replace("妥否，请批示。", "以上为本局意见。")
        self.assertIn("STR-003", ids(run(text, "请示", "structure")))

    def test_str003_ok(self):
        self.assertNotIn("STR-003", ids(run(self.QINGSHI_OK, "请示", "structure")))

    def test_str006_multiple_recipients(self):
        text = self.QINGSHI_OK.replace("区人民政府：", "区人民政府、区财政局：")
        self.assertIn("STR-006", ids(run(text, "请示", "structure")))

    def test_str006_leader_personal(self):
        text = self.QINGSHI_OK.replace("区人民政府：", "王区长：")
        self.assertIn("STR-006", ids(run(text, "请示", "structure")))

    def test_str007_qingshi_with_tecibaogao(self):
        text = self.QINGSHI_OK.replace("妥否，请批示。", "妥否，请批示。特此报告。")
        self.assertIn("STR-007", ids(run(text, "请示", "structure")))

    def test_str004_report_with_request(self):
        text = "某镇关于防汛工作情况的报告\n\n县人民政府：\n\n汛期工作开展顺利。恳请县政府拨付补助资金200万元。\n\n妥否，请批示。\n"
        found = issues_of(run(text, "报告", "structure"), "STR-004")
        self.assertEqual(sorted(i["level"] for i in found), ["P0", "P1"])

    def test_str004_past_action_ok(self):
        text = "某镇关于防汛工作情况的报告\n\n县人民政府：\n\n该项目已报请批准实施。\n\n特此报告。\n"
        self.assertNotIn("STR-004", ids(run(text, "报告", "structure")))

    def test_str004_only_for_report(self):
        self.assertNotIn("STR-004", ids(run(self.QINGSHI_OK, "请示", "structure")))

    def test_str008_letter(self):
        text = "某市某局关于商请支持某项工作的函\n\n某市某局：\n\n恳请贵局给予支持。请批示。\n"
        self.assertIn("STR-008", ids(run(text, "函", "structure")))

    def test_str009_upward_without_signer(self):
        text = "某县某局文件\n\n某发〔2026〕5号\n\n某县某局关于某事项的请示\n\n县人民政府：\n\n内容。\n\n妥否，请批示。\n"
        self.assertIn("STR-009", ids(run(text, "请示", "structure")))
        with_signer = text.replace("某发〔2026〕5号", "某发〔2026〕5号　　签发人：张某")
        self.assertNotIn("STR-009", ids(run(with_signer, "请示", "structure")))

    def test_str010_qingshi_baogao(self):
        text = "某局关于申请经费的请示报告\n\n县人民政府：\n\n内容。妥否，请批示。\n"
        self.assertIn("STR-010", ids(run(text, checks="structure")))

    def test_str010_type_mismatch_p2(self):
        found = issues_of(run(self.QINGSHI_OK, "报告", "structure"), "STR-010")
        self.assertEqual([i["level"] for i in found], ["P2"])

    def test_str011_downward_with_request(self):
        text = NOTICE_HEAD + "请按时报送。妥否，请批示。\n"
        self.assertIn("STR-011", ids(run(text, "通知", "structure")))

    def test_str005_title_punct(self):
        text = "某市某局关于做好某项工作的通知。\n\n各区某局：\n\n内容。\n"
        found = issues_of(run(text, "通知", "structure"), "STR-005")
        self.assertEqual(len(found), 1)
        self.assertEqual(found[0]["line"], 1)

    def test_str005_body_sentence_not_title(self):
        r = run("全年共受理群众来信来访326件，办结率98%。", checks="structure")
        self.assertNotIn("STR-005", ids(r))


class DocTypeInferenceTest(unittest.TestCase):
    def test_infer(self):
        cases = {
            "某局关于申请经费的请示": "请示",
            "某局2026年工作总结": "总结",
            "关于某县产业发展情况的调研报告": "调研报告",
            "某局关于印发《某某实施方案》的通知": "通知",
            "某局关于商请支持的函": "函",
        }
        for title, expected in cases.items():
            with self.subTest(title=title):
                r = run(title + "\n\n内容。\n")
                self.assertEqual(r["doc_type"], expected)
                self.assertEqual(r["doc_type_source"], "标题推断")

    def test_speech_subtitle(self):
        r = run("坚定信心　奋勇争先\n——在全市安全生产工作会议上的讲话\n\n同志们：\n\n内容。\n")
        self.assertEqual(r["doc_type"], "讲话稿")

    def test_default(self):
        r = run("全年共受理群众来信来访326件。\n")
        self.assertEqual((r["doc_type"], r["doc_type_source"]), ("其他", "默认"))

    def test_wenhao_reference_in_body_not_header(self):
        r = run("某局关于做好某项工作的通知\n\n各区某局：\n\n根据某发〔2026〕5号文件要求，现将有关事项通知如下：\n")
        self.assertEqual(r["doc_type"], "通知")
        self.assertEqual(ids(r), [])

    def test_qianbao_with_qingshi_title_no_mismatch(self):
        r = run("关于采购办公设备的请示\n\n院领导：\n\n内容。妥否，请批示。\n", "签报")
        self.assertNotIn("STR-010", ids(r))

    def test_param_wins(self):
        r = run("某局关于申请经费的请示\n\n内容。\n", "报告")
        self.assertEqual((r["doc_type"], r["doc_type_source"]), ("报告", "参数"))


class PrivacyTest(unittest.TestCase):
    # 990102 不是真实的行政区划代码，号码均为虚构
    VALID_ID = "990102195803150011"

    def test_pri001_valid_id_p0_and_masked(self):
        r = run(f"申请人身份证号{self.VALID_ID}。", checks="privacy")
        found = issues_of(r, "PRI-001")
        self.assertEqual(found[0]["level"], "P0")
        self.assertNotIn(self.VALID_ID, found[0]["excerpt"])
        self.assertNotIn(self.VALID_ID, found[0].get("context", ""))
        self.assertEqual(found[0]["excerpt"], "990102********0011")

    def test_pri001_bad_checksum_p1(self):
        found = issues_of(run("身份证号990102195803150012。", checks="privacy"), "PRI-001")
        self.assertEqual(found[0]["level"], "P1")

    def test_pri001_masked_id_ignored(self):
        self.assertNotIn("PRI-001", ids(run("身份证号990102********0011。", checks="privacy")))

    def test_pri002_contact_phone_p2(self):
        found = issues_of(run("联系人：王某，电话：13912345678。", checks="privacy"), "PRI-002")
        self.assertEqual(found[0]["level"], "P2")
        self.assertEqual(found[0]["excerpt"], "139****5678")

    def test_pri002_bare_phone_p0(self):
        found = issues_of(run("张某 13912345678 已核实。", checks="privacy"), "PRI-002")
        self.assertEqual(found[0]["level"], "P0")

    def test_pri002_with_identity_p0(self):
        text = f"刘某，身份证号{self.VALID_ID}，联系电话13812345678。"
        found = issues_of(run(text, checks="privacy"), "PRI-002")
        self.assertEqual(found[0]["level"], "P0")

    def test_pri002_amount_and_masked_ignored(self):
        r = run("总投资15000000000元。联系电话138****5678。", checks="privacy")
        self.assertNotIn("PRI-002", ids(r))

    def test_pri003_bank_card_labeled(self):
        found = issues_of(run("银行卡号：6222 0212 3456 7890 123。", checks="privacy"), "PRI-003")
        self.assertEqual(found[0]["level"], "P0")

    def test_pri003_account_p2(self):
        found = issues_of(run("户名：某市某局，账号：1202 0212 0900 1234 567。", checks="privacy"), "PRI-003")
        self.assertEqual(found[0]["level"], "P2")

    def test_pri003_unlabeled_luhn(self):
        # 4111111111111111 为公开的测试卡号
        found = issues_of(run("卡4111111111111111已冻结。", checks="privacy"), "PRI-003")
        self.assertEqual(found[0]["level"], "P1")

    def test_pri003_random_long_number_ignored(self):
        self.assertNotIn("PRI-003", ids(run("合同编号1234567890123456。", checks="privacy")))

    def test_pri004_labeled_numbers(self):
        r = run("住院号2026091234，学号20240310215，工号A1023。", checks="privacy")
        levels = sorted(i["level"] for i in issues_of(r, "PRI-004"))
        self.assertEqual(levels, ["P0", "P0", "P1"])

    def test_pri004_masked_ignored(self):
        self.assertNotIn("PRI-004", ids(run("学号2024****15。", checks="privacy")))

    def test_pri008_address(self):
        self.assertIn("PRI-008", ids(run("家庭住址：某小区3幢2单元501室。", checks="privacy")))
        self.assertNotIn("PRI-008", ids(run("住址变更的，及时到社区登记。", checks="privacy")))

    def test_pri009_name_gender_age(self):
        found = issues_of(run("患者王建国，男，72岁，术后恢复良好。", checks="privacy"), "PRI-009")
        self.assertEqual(found[0]["excerpt"], "王**，男，72岁")
        self.assertNotIn("PRI-009", ids(run("患者张某，男，72岁，术后恢复良好。", checks="privacy")))
        self.assertNotIn("PRI-009", ids(run("患者，男，72岁，术后恢复良好。", checks="privacy")))

    def test_pri010_bed_and_name(self):
        self.assertIn("PRI-010", ids(run("12床李桂兰已康复出院。", checks="privacy")))
        self.assertNotIn("PRI-010", ids(run("12床患者已康复出院。全院床位紧张。", checks="privacy")))

    def test_pri005_secret_mark_and_pri006(self):
        r = run("机密★1年\n\n某局关于某事项的报告\n\n内容。\n", checks="privacy")
        self.assertIn("PRI-005", ids(r))
        r = run("内部资料　注意保存\n\n某局关于某事项的报告\n\n内容。\n", checks="privacy")
        self.assertIn("PRI-006", ids(r))
        self.assertNotIn("PRI-007", ids(r), "PRI-007 应让位于 PRI-006")
        self.assertEqual(r["doc_type"], "报告", "内部资料标识行不应被当作标题")


class LegalTest(unittest.TestCase):
    """legal（法规关）：通知、通告、公告和面向群众的文书才查；函、报告、总结不报。"""

    PUNISH = "对不配合的住户处以200元罚款，情节严重的取消其低保资格。"

    def test_notice_reported_p1(self):
        found = issues_of(run(NOTICE_HEAD + self.PUNISH + "\n", "通知", "legal"), "LEG-001")
        self.assertEqual([i["level"] for i in found], ["P1"])
        self.assertEqual(found[0]["gate"], "法规关")
        self.assertIn("核对法律依据和发文权限", found[0]["message"])

    def test_exempt_types(self):
        for doc_type in ("函", "报告", "总结"):
            with self.subTest(doc_type=doc_type):
                r = run(NOTICE_HEAD + self.PUNISH + "\n", doc_type, "legal")
                self.assertEqual(r["issues"], [])

    def test_public_facing_by_greeting(self):
        text = "关于规范小区停车管理的告知\n\n各位业主：\n\n" + self.PUNISH + "\n"
        r = run(text, "其他", "legal")
        self.assertTrue({"LEG-001", "LEG-003"} <= set(ids(r)))

    def test_public_facing_by_title(self):
        r = run("某小区物业管理须知\n\n" + self.PUNISH + "\n", None, "legal")
        self.assertIn("LEG-001", ids(r))

    def test_internal_material_not_checked(self):
        self.assertEqual(run("某局2026年执法工作情况\n\n全年处以罚款326起。\n", "其他", "legal")["issues"], [])

    def test_basis_written_skips(self):
        text = NOTICE_HEAD + "依据《中华人民共和国道路交通安全法》，对违法停车行为处以罚款。\n"
        self.assertNotIn("LEG-001", ids(run(text, "通知", "legal")))

    def test_leg005_public_only(self):
        to_units = NOTICE_HEAD + "对工作不力的，严肃追究相关人员责任。\n"
        self.assertNotIn("LEG-005", ids(run(to_units, "通知", "legal")))
        to_residents = "关于垃圾分类的通知\n\n各位居民：\n\n不配合的，将追究法律责任。\n"
        self.assertIn("LEG-005", ids(run(to_residents, "通知", "legal")))


class JianfuTest(unittest.TestCase):
    def test_jfu001_first_body_only(self):
        first = NOTICE_HEAD + "做好秋冬季防火工作具有十分重要的意义，现将有关事项通知如下：\n\n一、工作要求\n\n内容。\n"
        found = issues_of(run(first, "通知", "jianfu"), "JFU-001")
        self.assertEqual([i["level"] for i in found], ["P1"])
        later = NOTICE_HEAD + "现将有关事项通知如下：\n\n一、工作要求\n\n这项工作具有重要意义，各单位要抓好落实。\n"
        self.assertNotIn("JFU-001", ids(run(later, "通知", "jianfu")))

    def test_jfu001_skips_zhusong_without_colon(self):
        text = "某局关于做好某项工作的通知\n\n各区某局\n\n当前，防汛形势严峻，现将有关事项通知如下。\n"
        self.assertIn("JFU-001", ids(run(text, "通知", "jianfu")))

    def test_jfu004_doc_types(self):
        text = "该项工作纳入年度绩效考核，市里将开展专项督查。"
        self.assertEqual(len(issues_of(run(text, "讲话稿", "jianfu"), "JFU-004")), 2)
        self.assertNotIn("JFU-004", ids(run(text, "总结", "jianfu")))

    def test_jfu003_exempt_reporting(self):
        self.assertIn("JFU-003", ids(run("实行一票否决。", "通知", "jianfu")))
        self.assertNotIn("JFU-003", ids(run("实行一票否决。", "总结", "jianfu")))

    def test_length_stays_in_length_class(self):
        """篇幅上限保留在 length 类，--checks jianfu 不含 LEN-001。"""
        para = "各单位要按照职责分工抓好落实，确保各项措施落到实处。"
        text = NOTICE_HEAD + "\n\n".join([para] * 200) + "\n"
        self.assertNotIn("LEN-001", ids(run(text, "通知", "jianfu")))
        found = issues_of(run(text, "通知", "length"), "LEN-001")
        self.assertEqual(found[0]["check"], "length")


class PrivacyMarkingScopeTest(unittest.TestCase):
    """红线一只拦规范密级标识和标注性字样，正文中的普通词语不报。"""

    def test_ordinary_words_not_reported(self):
        text = ("某局保密委员会办公室2026年工作总结\n\n全年开展保密工作检查4次，加强涉密人员管理，"
                "严守工作秘密，规范标注“秘密★1年”“内部资料”等标识，编印《保密手册》（内部资料）500册，"
                "加强内部资料管理，防止内幕信息泄露。\n")
        self.assertEqual(run(text, "总结", "privacy")["issues"], [])

    def test_header_markings_reported_p0(self):
        for mark in ("秘密★1年", "机密★", "绝密", "内部资料", "工作秘密　不得外传", "内部使用"):
            with self.subTest(mark=mark):
                found = [i for i in run(mark + "\n\n某局关于某事项的报告\n\n内容。\n", checks="privacy")["issues"]
                         if i["id"] in ("PRI-005", "PRI-006")]
                self.assertEqual([i["level"] for i in found], ["P0"])


class MarkerAndLengthTest(unittest.TestCase):
    def test_marker_counts_and_note(self):
        r = run("检查【待补：次数】次，完工于【待核实：完工时间】，地点【待定】。", checks="marker")
        self.assertEqual(sorted(ids(r)), ["MRK-001", "MRK-002", "MRK-003"])
        # 规范的待补、待核实标记报 P2（开发规范第 8 节）；写法不规范的标记仍报 P1
        self.assertEqual({i["id"]: i["level"] for i in r["issues"]},
                         {"MRK-001": "P2", "MRK-002": "P2", "MRK-003": "P1"})
        self.assertTrue(any("【待补】1 处" in n and "【待核实】1 处" in n for n in r["notes"]))

    def _long_notice(self, n_chars: int) -> str:
        para = "各单位要按照职责分工抓好落实，确保各项措施落到实处。"
        body = "\n\n".join([para] * (n_chars // len(para) + 1))
        return NOTICE_HEAD + body + "\n"

    def test_len001_thresholds(self):
        self.assertNotIn("LEN-001", ids(run(self._long_notice(3800), "通知", "length")))
        found = issues_of(run(self._long_notice(4300), "通知", "length"), "LEN-001")
        self.assertEqual(found[0]["level"], "P2")
        self.assertIn("4000", found[0]["message"])
        found = issues_of(run(self._long_notice(5300), "通知", "length"), "LEN-001")
        self.assertIn("5000", found[0]["message"])

    def test_len001_not_for_summary(self):
        self.assertNotIn("LEN-001", ids(run(self._long_notice(6000), "总结", "length")))

    def test_len002_speech(self):
        self.assertIn("LEN-002", ids(run(self._long_notice(12500), "讲话稿", "length")))


class EngineBehaviourTest(unittest.TestCase):
    def test_once_rule_aggregates(self):
        found = issues_of(run("XXX同志、XXX同志、XXX同志参加。", checks="residue"), "RES-015")
        self.assertEqual(len(found), 1)
        self.assertEqual(found[0]["count"], 3)

    def test_doc_types_positive_filter(self):
        self.assertIn("STR-004", ids(run("情况如下。妥否，请批示。", "报告", "structure")))
        self.assertNotIn("STR-004", ids(run("情况如下。妥否，请批示。", "总结", "structure")))

    def test_exempt_doc_types(self):
        text = "基层减负不是减责任，而是减形式。"
        self.assertIn("AI-006", ids(run(text, "通知", "style")))
        self.assertNotIn("AI-006", ids(run(text, "讲话稿", "style")))

    def test_whitelist_removes_overlap(self):
        r = run("存在的主要问题是：台账不全。", checks="style")
        self.assertNotIn("AI-011", ids(r))

    def test_density_needs_two_hits(self):
        one = run("抓落实的关键——压实责任。", checks="style")
        self.assertNotIn("AI-010", ids(one))
        two = run("抓落实的关键——压实责任。管理的要害——落实到人。", checks="style")
        found = issues_of(two, "AI-010")
        self.assertEqual(found[0]["count"], 2)

    def test_checks_filter(self):
        r = run("**加粗**。某发[2026]5号。", checks="residue")
        self.assertEqual({i["check"] for i in r["issues"]}, {"residue"})
        self.assertEqual(r["checks"], ["residue"])

    def test_parse_checks_rejects_unknown(self):
        with self.assertRaises(gl.InputError):
            gl.parse_checks("residue,foo")

    def test_fix_suggestion(self):
        found = issues_of(run("某农发[2026]第05号", checks="format"), "FMT-001")
        self.assertEqual(found[0]["suggestion"], "改为“某农发〔2026〕5号”")
        self.assertEqual(found[0]["excerpt"], "某农发[2026]第05号")


if __name__ == "__main__":
    unittest.main()
