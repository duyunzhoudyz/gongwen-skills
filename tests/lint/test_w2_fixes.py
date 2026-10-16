# -*- coding: utf-8 -*-
"""v1.0 第四阶段（W2）机检修复的回归测试：文书类型词表、数据表述、法规、减负、文风例外。

每个测试类对应 tests/format/_out/w2_tasks.json 中 lint 组的一条任务，类说明里写任务编号。
"""
from __future__ import annotations

import unittest

from tests.lint.helpers import LINTER, RULES_PATH, gl, ids, issues_of, load_json, run

NEW_TYPES = ("对照检查材料", "党课讲稿", "心得体会", "研讨发言", "会议记录", "议题材料", "简报", "新闻稿",
             "整改方案", "整改报告", "自查报告", "贯彻落实方案", "自评报告", "公众号推文")


class DocTypeVocabularyTest(unittest.TestCase):
    """L7、D10：文书类型词表扩充为开发规范第 6 节的全部类型，新类型套用合适的规则范围。"""

    def test_new_types_in_vocabulary(self):
        for doc_type in NEW_TYPES:
            with self.subTest(doc_type=doc_type):
                self.assertIn(doc_type, gl.DOC_TYPES)
                notes = run("内容。", doc_type)["notes"]
                self.assertFalse(any("不在统一词表" in n for n in notes), notes)
        self.assertTrue(any("不在统一词表" in n for n in run("内容。", "通稿")["notes"]))

    def test_infer_from_title(self):
        cases = {
            "某某同志2026年度民主生活会个人对照检查材料": "对照检查材料",
            "2026年第三季度专题党课": "党课讲稿",
            "学习贯彻全会精神心得体会": "心得体会",
            "理论学习中心组专题研讨发言提纲": "研讨发言",
            "某镇党委会会议记录": "会议记录",
            "关于审议某项目的议题材料": "议题材料",
            "迎评工作简报": "简报",
            "某医院义诊活动新闻稿": "新闻稿",
            "关于巡察反馈问题整改工作方案": "整改方案",
            "关于巡察整改进展情况的报告": "整改报告",
            "某学校安全工作自查报告": "自查报告",
            "某医院等级评审自评报告": "自评报告",
            "关于贯彻落实全省安全生产会议精神的工作方案": "贯彻落实方案",
            "某局关于整改工作的请示报告": "请示",          # “请示报告”仍留给 STR-010
            "关于印发《某局整改方案》的通知": "通知",
            "某局2026年工作总结": "总结",
        }
        for title, expected in cases.items():
            with self.subTest(title=title):
                self.assertEqual(run(title + "\n\n内容。\n")["doc_type"], expected)

    def test_infer_from_subtitle(self):
        self.assertEqual(run("坚定理想信念\n——在党支部专题党课上的讲稿\n\n同志们：\n\n内容。\n")["doc_type"], "党课讲稿")
        self.assertEqual(run("以学促干\n——在理论学习中心组专题研讨会上的发言\n\n内容。\n")["doc_type"], "研讨发言")
        self.assertEqual(run("坚定信心\n——在全市安全生产工作会议上的讲话\n\n内容。\n")["doc_type"], "讲话稿")

    def test_title_with_author_line(self):
        # “心得体会”标题以文种结尾，下一行的署名不并入标题
        r = run("学习贯彻全会精神心得体会\n某局办公室　王某\n\n通过学习，我认识到……\n")
        self.assertEqual(r["doc_type"], "心得体会")

    def test_speech_family_relaxed(self):
        text = "基层减负不是减责任，而是减形式。说白了，就是要把精力用在干事上。"
        self.assertTrue({"AI-005", "AI-006"} <= set(ids(run(text, "通知", "style"))))
        for doc_type in ("讲话稿", "党课讲稿", "研讨发言"):
            with self.subTest(doc_type=doc_type):
                self.assertFalse({"AI-005", "AI-006"} & set(ids(run(text, doc_type, "style"))))

    def test_plan_family_length_limit(self):
        body = "各单位要按照分工抓好落实，确保按期完成各项任务。" * 180   # 约 4300 字
        for doc_type in ("实施方案", "贯彻落实方案", "整改方案"):
            with self.subTest(doc_type=doc_type):
                self.assertIn("LEN-001", ids(run("某局某项工作方案\n\n" + body, doc_type, "length")))
        for doc_type in ("整改报告", "对照检查材料", "心得体会"):
            with self.subTest(doc_type=doc_type):
                self.assertNotIn("LEN-001", ids(run("某局某项工作材料\n\n" + body, doc_type, "length")))

    def test_plan_family_deploy_rules(self):
        text = "建议各单位于月底前完成自查。"
        for doc_type in ("贯彻落实方案", "整改方案"):
            with self.subTest(doc_type=doc_type):
                self.assertIn("DEF-015", ids(run(text, doc_type, "defensive")))
        # 整改方案的报送、考核要求来自巡视巡察整改规定，不查 JFU-002、JFU-004
        text = "请各单位于每月5日前报送整改落实情况，整改情况纳入年度考核。"
        self.assertTrue({"JFU-002", "JFU-004"} <= set(ids(run(text, "贯彻落实方案", "jianfu"))))
        self.assertFalse({"JFU-002", "JFU-004"} & set(ids(run(text, "整改方案", "jianfu"))))

    def test_field_heavy_types_skip_colon_density(self):
        text = "核心问题：机制不顺。主要原因：协调不够。"
        self.assertIn("AI-013", ids(run(text, "通知", "style")))
        for doc_type in ("议题材料", "会议记录", "整改方案", "整改报告", "自查报告", "贯彻落实方案"):
            with self.subTest(doc_type=doc_type):
                self.assertNotIn("AI-013", ids(run(text, doc_type, "style")))

    def test_info_family_like_xinxi(self):
        text = "该镇与各村签订目标责任状，推行“一票否决”。活动充分彰显了基层干部的担当。"
        self.assertTrue({"JFU-003", "JFU-009"} <= set(ids(run(text, "通知", "jianfu"))))
        for doc_type in ("信息", "简报", "新闻稿", "公众号推文"):
            with self.subTest(doc_type=doc_type):
                self.assertFalse({"JFU-003", "JFU-009"} & set(ids(run(text, doc_type, "jianfu"))))
                self.assertNotIn("AI-047", ids(run(text, doc_type, "style")))

    def test_report_family_structure(self):
        text = ("关于巡察整改进展情况的报告\n\n区委巡察办：\n\n已完成整改12项。恳请拨付整改资金20万元。\n\n"
                "妥否，请批示。\n")
        found = issues_of(run(text, None, "structure"), "STR-004")
        self.assertEqual(sorted(i["level"] for i in found), ["P0", "P1"])
        text = text.replace("区委巡察办：", "王书记：")
        self.assertIn("STR-006", ids(run(text, "整改报告", "structure")))

    def test_inspection_material_not_qingshi(self):
        text = "某局领导班子对照检查材料\n\n按照通知要求，对照检查如下。\n\n一、存在的问题\n\n学习不够深入。\n"
        result = run(text, None, "structure,style")
        self.assertEqual(result["doc_type"], "对照检查材料")
        self.assertFalse({"STR-003", "STR-004", "STR-006"} & set(ids(result)))
        self.assertIn("AI-053", ids(result))   # 空泛问题表述照查


class DataExpressionTest(unittest.TestCase):
    """W1-DT-2：data-expression.md 第十三部分 DE-01～DE-10。正则规则的必错、必过标本由 test_rule_examples 覆盖。"""

    MAPPING = {"NUM-013": "DE-01", "NUM-014": "DE-02", "NUM-015": "DE-03", "NUM-016": "DE-04", "NUM-017": "DE-05",
               "NUM-018": "DE-06", "NUM-019": "DE-07", "NUM-020": "DE-08", "NUM-021": "DE-09", "NUM-022": "DE-10"}

    def test_rules_registered_with_levels(self):
        data = load_json(RULES_PATH)
        items = {r["id"]: r for r in data["rules"] + data["procedural"]}
        for rid, de in self.MAPPING.items():
            with self.subTest(rule=rid):
                item = items[rid]
                self.assertEqual(item["check"], "number")
                self.assertIn(de, item["basis"])
                self.assertIn("data-expression.md", item["basis"])
                expected = "P2" if de in ("DE-08", "DE-09", "DE-10") else "P1"
                self.assertEqual(item["level"], expected)
        self.assertIn("NUM-020", {p["id"] for p in data["procedural"]})

    def test_de08_quantity_with_points(self):
        for text in ("门诊人次同比增长3个百分点。", "营业收入增长2个百分点。", "门诊量较上年增加1.5个百分点。"):
            with self.subTest(text=text):
                found = issues_of(run(text, checks="number"), "NUM-020")
                self.assertEqual(len(found), 1)
                self.assertEqual(found[0]["level"], "P2")
                self.assertIn(found[0]["excerpt"], text)
        for text in ("出院6171人次，四级手术占比提高2.5个百分点。", "接诊3000人次，满意度提高2个百分点。",
                     "床位使用比例提高2个百分点。", "门诊人次增速提高3个百分点。", "完成率比时间进度慢3.3个百分点。"):
            with self.subTest(text=text):
                self.assertNotIn("NUM-020", ids(run(text, checks="number")))

    def test_no_duplicate_with_num011_num012(self):
        # 与已有规则重叠时只报已有规则
        r = run("床位使用率增速比上年回落2.1%。", checks="number")
        self.assertIn("NUM-012", ids(r))
        self.assertNotIn("NUM-018", ids(r))
        r = run("约300多人左右参加培训。", checks="number")
        self.assertIn("NUM-011", ids(r))
        self.assertNotIn("NUM-022", ids(r))

    def test_de06_fix_suggestion(self):
        found = issues_of(run("营业收入增速比上年回落2.1%。", checks="number"), "NUM-018")
        self.assertEqual(found[0]["suggestion"], "改为“增速比上年回落2.1个百分点”")

    def test_de01_table_cells_skipped(self):
        lines = [(1, "利润总额增长-3.5%", True)]
        result = gl.lint_document(gl.Document("x.docx", lines, "段"), "其他", "number", LINTER)
        self.assertNotIn("NUM-013", ids(result))


class LegalWaterSupplyTest(unittest.TestCase):
    """W1-GR-3：LEG-003 漏报分红资格；以停水停电作为处罚手段（并入 LEG-002）。"""

    def test_hits(self):
        self.assertEqual(ids(run("拒不签订的，取消其年度村集体分红资格。", "通知", "legal")), ["LEG-003"])
        self.assertEqual(ids(run("拒不整改的，对其家庭停止供水。", "通知", "legal")), ["LEG-002"])
        self.assertEqual(ids(run("违规私接管道的，一律停水处理。", "通告", "legal")), ["LEG-002"])

    def test_outage_notice_not_reported(self):
        for text in ("因管道检修，10月20日8时至12时停水。", "我公司将于10月20日对城东片区停水，请提前储水。",
                     "取消原定于周五的会议。"):
            with self.subTest(text=text):
                self.assertEqual(ids(run(text, "通知", "legal")), [])

    def test_basis_written_skips(self):
        text = "依据《某市城市供水条例》，对逾期不缴纳水费的用户停止供水。"
        self.assertNotIn("LEG-002", ids(run(text, "通知", "legal")))


class Pun007LevelTest(unittest.TestCase):
    """RV-3：PUN-007 由 P0 降为 P1，与 gongwen-review 定级表一致（标点问题为 P1）。"""

    def test_level(self):
        found = issues_of(run("医院、学校、社区……等", checks="punct"), "PUN-007")
        self.assertEqual([i["level"] for i in found], ["P1"])

    def test_p0_rules_match_review_table(self):
        data = load_json(RULES_PATH)
        p0 = {r["id"] for r in data["rules"] if r["level"] == "P0"}
        self.assertEqual(p0, {"RES-010", "RES-011", "RES-012", "FMT-001", "FMT-002", "FMT-003", "FMT-004",
                              "PRI-005", "PRI-006"})


class LiteralMeaningTest(unittest.TestCase):
    """L4：“护城河”“拉通”在城市规划、水利、道路语境中的字面义不报。"""

    def test_literal_not_reported(self):
        for text in ("推进护城河清淤整治。", "年内拉通三条断头路。", "打通断头路、拉通主干道。",
                     "老城护城河水质明显改善。", "两条断头路已全线拉通。", "环城护城河沿岸绿化提升。"):
            with self.subTest(text=text):
                self.assertNotIn("AI-025", ids(run(text, checks="style")))

    def test_jargon_still_reported(self):
        for text in ("构建核心竞争护城河。", "拉通各部门数据。", "形成技术护城河，拉通上下游。"):
            with self.subTest(text=text):
                self.assertIn("AI-025", ids(run(text, checks="style")))

    def test_exclusion_is_rule_specific(self):
        # AI-025 的字面义例外只作用于 AI-025，同一位置的其他规则照常报告
        text = "赋能护城河整治，打造断头路拉通工程，助力水系连通。"
        found = set(ids(run(text, checks="style")))
        self.assertNotIn("AI-025", found)
        self.assertIn("AI-026", found)


class FieldColonTest(unittest.TestCase):
    """W1-RC-1：字段标签、称呼行、编号标签不计入冒号密度；空格隔开的字段行不报 PUN-014。"""

    SAMPLE = ("尊敬的家长：\n\n各位老师、同学们：\n\nS01：教学楼外墙瓷砖脱落，已修复。\n\n责任领导：王某\n\n"
              "整改措施：建立台账，每月通报。\n\n提交部门：教务处　汇报人：李某\n")

    def test_sample_not_reported(self):
        found = set(ids(run(self.SAMPLE, "通知", "style,punct")))
        self.assertNotIn("AI-013", found)
        self.assertNotIn("PUN-014", found)

    def test_label_style_prose_still_reported(self):
        self.assertIn("AI-013", ids(run("目标：三年见效。路径：先试点再推广。", "通知", "style")))
        self.assertIn("PUN-014", ids(run("会议要求：各单位要做到：按时报送。", checks="punct")))

    def test_exclusion_does_not_hide_other_rules(self):
        # 字段标签只对 AI-013 排除，“关键是：”仍按 AI-011 报
        self.assertIn("AI-011", ids(run("关键是：抓落实。", checks="style")))

    def test_field_line_helper(self):
        self.assertTrue(gl.Linter._is_field_line("提交部门：财务科　汇报人：李某"))
        self.assertTrue(gl.Linter._is_field_line("时间：9:30　地点：三楼会议室"))
        self.assertFalse(gl.Linter._is_field_line("会议要求：各单位要做到：按时报送。"))
        self.assertFalse(gl.Linter._is_field_line("目标：三年见效。路径：先试点再推广。"))


class JianfuNegationTest(unittest.TestCase):
    """W1-SP-2：减负类规则在否定语境（不要求、严禁、不再……）中不报。"""

    def test_negated_not_reported(self):
        cases = {
            "不要求基层拍照留痕。": "JFU-005",
            "严禁以签订责任状方式转嫁责任。": "JFU-009",
            "不得随意设置“一票否决”。": "JFU-003",
            "此项工作不再纳入年度考核。": "JFU-004",
            "考核不搞“千分制”。": "JFU-008",
            "取消按月排名通报。": "JFU-007",
        }
        for text, rid in cases.items():
            with self.subTest(text=text):
                self.assertNotIn(rid, ids(run(text, "通知", "jianfu")))

    def test_affirmative_still_reported(self):
        cases = {
            "各村要拍照上传。": "JFU-005",
            "与各村签订目标责任状。": "JFU-009",
            "实行安全生产“一票否决”。": "JFU-003",
            "此项工作纳入年度考核。": "JFU-004",
        }
        for text, rid in cases.items():
            with self.subTest(text=text):
                self.assertIn(rid, ids(run(text, "通知", "jianfu")))

    def test_negation_in_other_sentence_does_not_count(self):
        self.assertIn("JFU-005", ids(run("不搞形式。各村要拍照上传。", "通知", "jianfu")))


class OpeningPaddingTest(unittest.TestCase):
    """W1-PF-2：首段“在……坚强领导下”铺垫，P2，只查正文第一段。"""

    def test_first_paragraph(self):
        found = issues_of(run("在县委县政府的坚强领导下，我们取得成绩。", checks="jianfu"), "JFU-010")
        self.assertEqual([i["level"] for i in found], ["P2"])
        text = "某局2026年工作总结\n\n2026年，我局在县委县政府的坚强领导下，完成各项任务。\n\n一、主要工作\n\n内容。\n"
        self.assertEqual(len(issues_of(run(text, None, "jianfu"), "JFU-010")), 1)

    def test_middle_paragraph_not_reported(self):
        text = ("某局2026年工作总结\n\n2026年，我局完成检查12次，整改隐患36处。\n\n一、主要工作\n\n"
                "在县委县政府的坚强领导下，我们取得成绩。\n")
        self.assertNotIn("JFU-010", ids(run(text, None, "jianfu")))


if __name__ == "__main__":
    unittest.main()
