# -*- coding: utf-8 -*-
"""误报测试：规范的虚构公文跑一遍，P0、P1 都应为 0；问题密集稿件的主要问题都要报出。"""
from __future__ import annotations

import unittest

from tests.lint.helpers import EVAL_FILES, gl, ids, run

QINGSHI = """某区教育局文件

某教〔2026〕21号　　　　签发人：李某

某区教育局关于申请某中学教学楼加固维修经费的请示

区人民政府：

某中学教学楼建于1998年，建筑面积4600平方米。经某检测机构2026年9月鉴定，该楼为C级危房，需要加固维修。为保障师生安全，我局拟于2027年寒假期间实施加固维修，现将有关事项请示如下：

一、项目内容

主要包括结构加固、屋面防水和消防设施改造，工期60天。

二、资金需求

经某造价咨询公司测算，项目总投资386万元，其中结构加固246万元、屋面防水58万元、消防设施改造82万元。我局已统筹年度预算安排86万元，尚缺300万元。

三、我局意见

建议从区级教育费附加中安排300万元。项目由我局作为建设单位组织实施，严格执行政府采购规定。

妥否，请批示。

附件：某中学教学楼结构安全鉴定报告

某区教育局
2026年10月16日

（联系人：王某，电话：0571-86660000）
"""

BAOGAO = """某镇人民政府关于2026年度防汛救灾工作情况的报告

县人民政府：

今年汛期，我镇共经历强降雨过程6次，最大日降雨量186毫米，未发生人员伤亡。现将有关情况报告如下：

一、主要工作

（一）落实防汛责任。全镇划分防汛责任网格42个，明确镇、村、组三级责任人356名，汛前开展应急演练3次。

（二）排查整治隐患。排查山塘水库、地质灾害点、危旧房屋等隐患点127处，已整改119处，其余8处落实专人监测。

（三）转移安置群众。汛期累计转移群众1260人次，设置集中安置点5个。

二、存在的问题

一是部分山塘年久失修，泄洪能力不足；二是应急物资储备偏少；三是村级防汛队伍以老年人为主，抢险能力有限。

三、下一步打算

我镇将对5座病险山塘实施除险加固，新购冲锋舟2艘，每村组建不少于15人的应急抢险队。

特此报告。

某镇人民政府
2026年10月16日
"""

SPEECH = """坚持问题导向　抓好安全生产
——在全市安全生产工作会议上的讲话

同志们：

今天召开全市安全生产工作会议，主要任务是分析前三季度安全生产形势，部署第四季度工作。下面，我讲三点意见。

一、清醒认识当前形势

前三季度，全市发生生产安全事故42起、死亡17人，同比分别下降12.5%和15%。同时也要看到，建筑施工和道路运输领域事故仍然多发，占事故总数的六成以上。

二、突出抓好重点领域

（一）建筑施工。对在建工地开展一次全覆盖检查，重点查高处作业和起重机械。

（二）道路运输。对“两客一危”车辆实行动态监控，严查超速、疲劳驾驶。

（三）消防安全。对人员密集场所开展夜查，坚决消除重大火灾隐患。

三、压紧压实工作责任

各区、各部门主要负责同志要亲自研究部署，分管负责同志要靠前指挥。市安委办要加强督导，对工作不力的单位进行约谈。

同志们，四季度时间紧、任务重，希望大家坚持问题导向，扎实做好各项工作，确保全市安全生产形势持续稳定。
"""

ZONGJIE = """某医院护理部2026年工作总结

2026年，护理部围绕护理质量、人员培训和优质护理三项重点开展工作，现将全年工作总结如下。

一、护理质量

全年护理不良事件上报86例，均按规定完成根因分析。压力性损伤发生率为0.12%，较去年下降0.03个百分点。

二、人员培训

组织全院护士分层培训24场，参训2160人次；选派12名护士参加省级专科护士培训。

三、优质护理

优质护理病区覆盖率达到100%，患者满意度为96.8%。

四、存在问题

一是夜班护士配置偏紧；二是部分新入职护士专科能力不足。

五、明年打算

2027年，护理部将重点推进护理人力弹性调配，完善新护士规范化培训方案。
"""


BAOMI_ZONGJIE = """某市国家保密局2026年保密工作总结

2026年，市保密局围绕保密宣传教育、涉密人员管理和保密检查三项重点开展工作，现将全年工作总结如下。

一、保密宣传教育

组织全市机关干部保密知识培训12场，参训1860人次；编印《保密知识读本》（内部资料）2000册，发放到各单位。在保密宣传月期间，指导各单位规范标注“秘密★1年”等密级标识，严守工作秘密，不在非涉密计算机上处理工作秘密。

二、涉密人员管理

全市在册涉密人员3260名，全部完成年度保密审查和保密承诺书签订，涉密人员出国（境）审批办结率100%。

三、保密检查

开展保密工作检查4次，覆盖市直单位86家，发现问题23个，已全部整改。加强内部资料管理，防止内幕信息泄露。

四、存在问题

一是部分单位保密自查流于形式，23个问题中有9个属于自查未发现；二是新入职人员保密培训覆盖率为76%。

五、明年打算

2027年，市保密局将把新入职人员保密培训覆盖率提高到100%，每半年开展一次涉密载体专项清查。
"""

class FalsePositiveTest(unittest.TestCase):
    def assert_clean(self, result):
        bad = [(i["level"], i["id"], i["line"], i["excerpt"]) for i in result["issues"] if i["level"] in ("P0", "P1")]
        self.assertEqual(bad, [], "规范稿件不应出现 P0、P1")

    def test_normal_notice_file(self):
        result = gl.lint_file(EVAL_FILES / "规范通知.txt")
        self.assertEqual(result["doc_type"], "通知")
        self.assert_clean(result)

    def test_normal_qingshi(self):
        result = run(QINGSHI)
        self.assertEqual(result["doc_type"], "请示")
        self.assert_clean(result)

    def test_normal_baogao(self):
        result = run(BAOGAO)
        self.assertEqual(result["doc_type"], "报告")
        self.assert_clean(result)

    def test_normal_speech(self):
        result = run(SPEECH)
        self.assertEqual(result["doc_type"], "讲话稿")
        self.assert_clean(result)

    def test_normal_summary(self):
        result = run(ZONGJIE)
        self.assertEqual(result["doc_type"], "总结")
        self.assert_clean(result)

    def test_secrecy_office_summary(self):
        """保密办自己的总结正常提到“保密工作”“涉密人员管理”“工作秘密”，不应被拦截。"""
        result = run(BAOMI_ZONGJIE)
        self.assertEqual(result["doc_type"], "总结")
        self.assert_clean(result)
        self.assertEqual([i for i in result["issues"] if i["check"] == "privacy"], [])


class ProblemDenseTest(unittest.TestCase):
    """问题密集稿件：主要问题都要报出。"""

    def test_problem_notice(self):
        result = gl.lint_file(EVAL_FILES / "问题密集通知.txt", "通知")
        found = set(ids(result))
        expected = {
            "FMT-001", "FMT-003", "FMT-004",   # 发文字号
            "STR-005",                         # 标题末尾标点
            "FMT-007",                         # 主送缺冒号
            "PUN-012",                         # 半角冒号
            "PUN-002", "FMT-012", "PUN-001", "PUN-003",  # 序数写法
            "PUN-007",                         # 省略号与“等”
            "NUM-001", "NUM-002", "NUM-003",   # 年份、范围
            "STR-001", "STR-002",              # 序号跳号、数字概括
            "PUN-005",                         # 并列书名号
            "STR-011",                         # 下行文用请示结语
            "DEF-001", "DEF-002",              # 免责声明
            "FMT-008", "FMT-006", "FMT-009",   # 附件说明、成文日期、抄送
            "LEG-001",                         # 罚款、拘留
            "JFU-001", "JFU-003", "JFU-006", "JFU-007",  # 首段铺垫、一票否决、纸质电子双报、按月排名
        }
        self.assertEqual(expected - found, set())
        self.assertGreaterEqual(result["summary"]["P0"], 5)

    def test_report_with_request(self):
        result = gl.lint_file(EVAL_FILES / "报告夹带请示.txt")
        levels = sorted(i["level"] for i in result["issues"] if i["id"] == "STR-004")
        self.assertEqual(levels, ["P0", "P1"])

    def test_privacy_material(self):
        result = gl.lint_file(EVAL_FILES / "社区救助对象申报材料.txt")
        found = set(ids(result))
        self.assertTrue({"PRI-006", "PRI-001", "PRI-008", "PRI-009", "PRI-002"} <= found)
        self.assertTrue(all(i["level"] == "P0" for i in result["issues"] if i["check"] == "privacy"))

    def test_resident_notice_legal(self):
        result = gl.lint_file(EVAL_FILES / "居民垃圾分类通知.txt")
        found = set(ids(result))
        self.assertTrue({"LEG-001", "LEG-003", "LEG-004", "JFU-005"} <= found)
        self.assertTrue(all(i["level"] == "P1" for i in result["issues"] if i["check"] == "legal"))

    def test_hospital_summary(self):
        result = gl.lint_file(EVAL_FILES / "骨科年度工作总结.txt")
        found = set(ids(result))
        self.assertTrue({"PRI-004", "PRI-009", "PRI-010", "NUM-012", "MRK-001"} <= found)


if __name__ == "__main__":
    unittest.main()
