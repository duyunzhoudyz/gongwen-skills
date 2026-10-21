# -*- coding: utf-8 -*-
"""要素完整性（ELM-001～ELM-007）：每条规则的必错标本、必过标本。"""
from __future__ import annotations

import unittest

from tests.lint.helpers import LINTER, gl, ids, issues_of, run

TONGZHI_OK = """某区教育局关于做好2026年寒假安全工作的通知

各学校、各幼儿园：

为切实做好寒假期间安全工作，现将有关事项通知如下。

一、排查隐患

各校要在1月10日前完成校舍、消防设施和食堂的安全排查，建立整改台账。

二、落实值班

寒假期间实行领导带班和24小时值班制度，发现问题立即报告。详见附件。

特此通知。

附件：寒假安全检查表

某区教育局
2026年12月20日
"""

QINGSHI_OK = """某区教育局关于申请某中学教学楼加固维修经费的请示

区人民政府：

某中学教学楼建于1998年，经鉴定为C级危房。为保障师生安全，拟于2027年寒假期间实施加固维修，现将有关事项请示如下：

一、项目内容

主要包括结构加固、屋面防水和消防设施改造，工期60天，总投资386万元。

二、资金需求

我局已统筹年度预算86万元，尚缺300万元，建议从区级教育费附加中安排。

妥否，请批示。

附件：某中学教学楼结构安全鉴定报告

某区教育局
2026年10月16日
"""

BAOGAO_OK = """某镇人民政府关于2026年度防汛救灾工作情况的报告

县人民政府：

今年汛期，我镇共经历强降雨过程6次，最大日降雨量186毫米，未发生人员伤亡。现将有关情况报告如下：

一、主要工作

全镇划分防汛责任网格42个，明确三级责任人356名，汛前开展应急演练3次。

二、存在问题

部分村组排水设施老化，转移避险路线标识不清，群众自救互救能力有待提高，我们将在明年汛前集中整改。

特此报告。

某镇人民政府
2026年11月2日
"""

HAN_OK = """某区教育局关于商请协助开展校园周边环境整治的函

某区城市管理局：

我区部分学校周边存在占道经营问题，影响师生出行安全。为保障上下学秩序，拟于11月开展联合整治，整治范围涉及12所中小学周边300米范围内的流动摊点和占道停车，现商请贵局派员参加，并安排执法车辆配合。

请予支持。

某区教育局
2026年10月28日
"""

JIYAO_OK = """某区教育局党委会会议纪要

2026年10月16日上午，区教育局党委召开会议，研究2027年度工作安排。

出席：张某某、李某某、王某某。
列席：办公室主任赵某某。

会议议定：一、成立工作专班，由办公室牵头，各科室各派一名联络员；二、11月底前完成方案并报局党委审定；三、12月起分阶段推进落实，每月底通报进展。

某区教育局
2026年10月16日
"""


def elm(text, doc_type):
    return [i for i in run(text, doc_type)["issues"] if i["id"].startswith("ELM-")]


def elm_ids(text, doc_type):
    return [i["id"] for i in elm(text, doc_type)]


class ElementPassTest(unittest.TestCase):
    """必过标本：规范的完整公文不报任何 ELM。"""

    def test_complete_documents(self):
        for name, text, dt in (("通知", TONGZHI_OK, "通知"), ("请示", QINGSHI_OK, "请示"), ("报告", BAOGAO_OK, "报告"),
                               ("函", HAN_OK, "函"), ("纪要", JIYAO_OK, "纪要")):
            with self.subTest(name=name):
                self.assertEqual(elm_ids(text, dt), [])

    def test_fragment_not_checked(self):
        """没有标题又很短的片段（如只贴了一段正文）不查要素。"""
        self.assertEqual(elm_ids("现将有关事项通知如下：\n\n一、做好值班。\n\n二、报送材料。", "通知"), [])

    def test_unknown_doc_type_not_checked(self):
        self.assertEqual(elm_ids(TONGZHI_OK.replace("某区教育局\n2026年12月20日\n", ""), "其他"), [])

    def test_title_without_org_is_allowed(self):
        """标题省略发文机关名称（实务惯例）不报。"""
        text = TONGZHI_OK.replace("某区教育局关于做好", "关于做好", 1)
        self.assertEqual([i for i in elm_ids(text, "通知") if i == "ELM-001"], [])

    def test_public_notice_without_zhusong(self):
        text = """关于暂停办理部分业务的通知

各位居民：

因系统升级，社区服务中心11月3日至5日暂停办理证明开具业务，给您带来不便，敬请谅解。

某社区居委会
2026年10月30日
"""
        self.assertNotIn("ELM-002", elm_ids(text, "通知"))


class ElementTitleTest(unittest.TestCase):
    def test_title_without_kind(self):
        text = TONGZHI_OK.replace("某区教育局关于做好2026年寒假安全工作的通知", "某区教育局做好2026年寒假安全工作", 1)
        self.assertIn("ELM-001", elm_ids(text, "通知"))
        self.assertEqual(issues_of(run(text, "通知"), "ELM-001")[0]["level"], "P1")

    def test_title_without_reason(self):
        text = TONGZHI_OK.replace("某区教育局关于做好2026年寒假安全工作的通知", "某区教育局通知", 1)
        found = issues_of(run(text, "通知"), "ELM-001")
        self.assertTrue(found)
        self.assertIn("缺少事由", found[0]["message"])

    def test_title_guanyu_empty(self):
        text = TONGZHI_OK.replace("某区教育局关于做好2026年寒假安全工作的通知", "关于通知", 1)
        self.assertIn("ELM-001", elm_ids(text, "通知"))

    def test_missing_title(self):
        body = TONGZHI_OK.split("\n", 2)[2]
        found = issues_of(run(body, "通知"), "ELM-001")
        self.assertTrue(found)
        self.assertIn("缺少标题", found[0]["message"])

    def test_title_placeholder_is_p2(self):
        text = TONGZHI_OK.replace("某区教育局关于做好2026年寒假安全工作的通知", "关于【待补：事由】的通知", 1)
        found = issues_of(run(text, "通知"), "ELM-001")
        self.assertEqual([i["level"] for i in found], ["P2"])

    def test_missing_title_with_placeholder_is_p2(self):
        body = "【待补：标题】\n\n" + TONGZHI_OK.split("\n", 2)[2]
        found = issues_of(run(body, "通知"), "ELM-001")
        self.assertEqual([i["level"] for i in found], ["P2"])


class ElementZhusongTest(unittest.TestCase):
    def test_missing_zhusong(self):
        text = TONGZHI_OK.replace("各学校、各幼儿园：\n\n", "", 1)
        found = issues_of(run(text, "通知"), "ELM-002")
        self.assertEqual([i["level"] for i in found], ["P1"])

    def test_lead_in_line_is_not_zhusong(self):
        """“现将……通知如下：”不是主送机关。"""
        text = TONGZHI_OK.replace("各学校、各幼儿园：\n\n", "现将有关事项通知如下：\n\n", 1)
        self.assertIn("ELM-002", elm_ids(text, "通知"))

    def test_zhusong_placeholder(self):
        text = TONGZHI_OK.replace("各学校、各幼儿园：", "【待补：主送机关】：", 1)
        found = issues_of(run(text, "通知"), "ELM-002")
        self.assertEqual([i["level"] for i in found], ["P2"])

    def test_zhusong_in_banji(self):
        text = TONGZHI_OK.replace("各学校、各幼儿园：\n\n", "", 1) + "\n主送：各学校、各幼儿园\n"
        self.assertNotIn("ELM-002", elm_ids(text, "通知"))

    def test_not_for_gonggao(self):
        text = """某市人民政府关于调整城区限行措施的公告

自2026年11月1日起，城区限行范围调整为二环路以内，具体路段见附件。

特此公告。

某市人民政府
2026年10月20日

附件：限行路段清单
"""
        self.assertNotIn("ELM-002", elm_ids(text, "公告"))


class ElementClosingTest(unittest.TestCase):
    def test_baogao_without_closing(self):
        text = BAOGAO_OK.replace("特此报告。\n\n", "", 1)
        found = issues_of(run(text, "报告"), "ELM-003")
        self.assertEqual([i["level"] for i in found], ["P2"])

    def test_han_without_closing(self):
        text = HAN_OK.replace("请予支持。\n\n", "", 1)
        self.assertIn("ELM-003", elm_ids(text, "函"))

    def test_han_alternative_closings(self):
        for ending in ("特此函告。", "盼复。", "请予函复。", "专此函达。"):
            with self.subTest(ending=ending):
                text = HAN_OK.replace("请予支持。", ending, 1)
                self.assertNotIn("ELM-003", elm_ids(text, "函"))

    def test_qingshi_closing_left_to_str003(self):
        text = QINGSHI_OK.replace("妥否，请批示。\n\n", "", 1)
        self.assertNotIn("ELM-003", elm_ids(text, "请示"))
        self.assertIn("STR-003", ids(run(text, "请示")))

    def test_blank_day_date_not_missing(self):
        """起草时成文日期日留空（“2026年10月　日”，签发时填）：不报缺少成文日期。"""
        text = TONGZHI_OK.replace("2026年12月20日\n", "2026年12月　日\n", 1)
        self.assertNotIn("ELM-005", elm_ids(text, "通知"))

    def test_tongzhi_not_required(self):
        text = TONGZHI_OK.replace("特此通知。\n\n", "", 1)
        self.assertNotIn("ELM-003", elm_ids(text, "通知"))


class ElementSignoffTest(unittest.TestCase):
    def test_missing_org_signature(self):
        text = TONGZHI_OK.replace("某区教育局\n2026年12月20日", "2026年12月20日", 1)
        found = issues_of(run(text, "通知"), "ELM-004")
        self.assertEqual([i["level"] for i in found], ["P1"])
        self.assertNotIn("ELM-005", elm_ids(text, "通知"))

    def test_missing_date(self):
        text = TONGZHI_OK.replace("\n2026年12月20日", "", 1)
        found = issues_of(run(text, "通知"), "ELM-005")
        self.assertEqual([i["level"] for i in found], ["P1"])

    def test_missing_both(self):
        text = TONGZHI_OK.replace("某区教育局\n2026年12月20日\n", "", 1)
        got = elm_ids(text, "通知")
        self.assertIn("ELM-004", got)
        self.assertIn("ELM-005", got)

    def test_org_and_date_on_one_line(self):
        text = TONGZHI_OK.replace("某区教育局\n2026年12月20日", "某区教育局　2026年12月20日", 1)
        got = elm_ids(text, "通知")
        self.assertNotIn("ELM-004", got)
        self.assertNotIn("ELM-005", got)

    def test_org_with_seal_mark(self):
        text = TONGZHI_OK.replace("某区教育局\n2026年12月20日", "某区教育局（盖章）\n2026年12月20日", 1)
        self.assertNotIn("ELM-004", elm_ids(text, "通知"))

    def test_placeholders_are_p2(self):
        text = TONGZHI_OK.replace("某区教育局\n2026年12月20日\n", "【待补：落款单位】\n【待补：成文日期】\n", 1)
        found = [i for i in elm(text, "通知") if i["id"] in ("ELM-004", "ELM-005")]
        self.assertTrue(all(i["level"] == "P2" for i in found), found)

    def test_jiyao_not_required_signoff(self):
        text = JIYAO_OK.replace("某区教育局\n2026年10月16日\n", "")
        got = elm_ids(text, "纪要")
        self.assertNotIn("ELM-004", got)
        self.assertNotIn("ELM-005", got)

    def test_attachment_after_signoff_is_not_org(self):
        """落款在附件说明之上也能识别；附件说明行不当作落款。"""
        text = TONGZHI_OK.replace("附件：寒假安全检查表\n\n某区教育局\n2026年12月20日", "某区教育局\n2026年12月20日\n\n附件：寒假安全检查表", 1)
        self.assertNotIn("ELM-004", elm_ids(text, "通知"))


class ElementAttachmentTest(unittest.TestCase):
    def test_reference_without_description(self):
        text = TONGZHI_OK.replace("附件：寒假安全检查表\n\n", "", 1)
        found = issues_of(run(text, "通知"), "ELM-006")
        self.assertEqual([i["level"] for i in found], ["P1"])
        self.assertIn("没有附件说明", found[0]["message"])

    def test_number_beyond_list(self):
        text = TONGZHI_OK.replace("详见附件。", "详见附件3。", 1)
        found = issues_of(run(text, "通知"), "ELM-006")
        self.assertEqual([i["level"] for i in found], ["P1"])
        self.assertIn("附件3", found[0]["message"])

    def test_multi_attachments_numbered(self):
        text = TONGZHI_OK.replace("详见附件。", "检查表见附件1，台账格式见附件2。", 1).replace(
            "附件：寒假安全检查表", "附件：1. 寒假安全检查表\n      2. 整改台账模板", 1)
        self.assertNotIn("ELM-006", elm_ids(text, "通知"))

    def test_uncited_attachment_is_p2(self):
        text = TONGZHI_OK.replace("详见附件。", "检查表见附件1。", 1).replace(
            "附件：寒假安全检查表", "附件：1. 寒假安全检查表\n      2. 整改台账模板", 1)
        found = issues_of(run(text, "通知"), "ELM-006")
        self.assertEqual([i["level"] for i in found], ["P2"])
        self.assertIn("附件2", found[0]["message"])

    def test_description_without_any_reference_is_ok(self):
        text = TONGZHI_OK.replace("详见附件。", "", 1)
        self.assertNotIn("ELM-006", elm_ids(text, "通知"))

    def test_reference_with_placeholder_is_p2(self):
        text = TONGZHI_OK.replace("附件：寒假安全检查表\n\n", "【待补：附件名称】\n\n", 1)
        found = issues_of(run(text, "通知"), "ELM-006")
        self.assertEqual([i["level"] for i in found], ["P2"])


class ElementJiyaoTest(unittest.TestCase):
    def test_missing_attendance(self):
        text = JIYAO_OK.replace("出席：张某某、李某某、王某某。\n列席：办公室主任赵某某。\n\n", "", 1)
        found = issues_of(run(text, "纪要"), "ELM-007")
        self.assertEqual([i["level"] for i in found], ["P1"])

    def test_missing_attendee_only(self):
        text = JIYAO_OK.replace("列席：办公室主任赵某某。\n", "", 1)
        found = issues_of(run(text, "纪要"), "ELM-007")
        self.assertEqual([i["level"] for i in found], ["P2"])

    def test_leave_counts_as_listed(self):
        text = JIYAO_OK.replace("列席：办公室主任赵某某。", "请假：赵某某。", 1)
        self.assertNotIn("ELM-007", elm_ids(text, "纪要"))

    def test_only_for_jiyao(self):
        self.assertNotIn("ELM-007", elm_ids(TONGZHI_OK, "通知"))


class ElementRegistryTest(unittest.TestCase):
    def test_registered(self):
        for rid in ("ELM-001", "ELM-002", "ELM-003", "ELM-004", "ELM-005", "ELM-006", "ELM-007"):
            with self.subTest(rule=rid):
                meta = LINTER.proc_meta[rid]
                self.assertIn(meta["level"], ("P1", "P2"))
                self.assertTrue(meta["basis"])
                self.assertEqual(meta["check"], "format")

    def test_checks_format_only(self):
        text = TONGZHI_OK.replace("\n2026年12月20日", "", 1)
        self.assertIn("ELM-005", ids(run(text, "通知", checks="format")))
        self.assertNotIn("ELM-005", ids(run(text, "通知", checks="style")))

    def test_allow_downgrades(self):
        text = TONGZHI_OK.replace("\n2026年12月20日", "", 1)
        result = gl.lint_text(text, "通知", linter=LINTER, allow=["ELM-005"])
        self.assertEqual([i["level"] for i in issues_of(result, "ELM-005")], ["P2"])


if __name__ == "__main__":
    unittest.main()
