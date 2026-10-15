# -*- coding: utf-8 -*-
"""代码审查发现的缺陷回归测试：每个用例对应一处已修复的误报、漏报或性能问题。"""
from __future__ import annotations

import importlib.util
import sys
import time
import unittest

from tests.lint.helpers import REPO_ROOT, SCRIPTS_DIR, gl, ids, issues_of, run

FORMAT_SCRIPTS = REPO_ROOT / "skills" / "gongwen-format" / "scripts"


def _load(name: str, path):
    """按文件路径加载模块（不污染 sys.path），同一进程只加载一次。"""
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)  # type: ignore[union-attr]
    return module


class DraftNoteConsistencyTest(unittest.TestCase):
    """【起草说明】分界行：lint、fact_diff、gongwen-format 三处判定必须一致。

    原问题：gongwen-format 已把“## 【起草说明】”“**【起草说明】**”当作分界行，lint 和 fact_diff 不认，
    同一份 .md 稿件，format 排版时去掉了起草说明，lint 和 fact_diff 却把起草说明当正文检查。
    """

    VARIANTS = ("【起草说明】", "　【起草说明】", " 【起草说明】 ", "\t【起草说明】", "【起草说明】：",
                "## 【起草说明】", "**【起草说明】**", "### **【起草说明】**", "__【起草说明】__")
    NOT_BOUNDARY = ("正文提到【起草说明】的写法。", "起草说明", "#######【起草说明】", "- 【起草说明】")

    def test_three_scripts_agree(self):
        fact_diff = _load("fact_diff_under_review", SCRIPTS_DIR / "fact_diff.py")
        gw_layout = _load("gw_layout_under_review", FORMAT_SCRIPTS / "gw_layout.py")
        for line in self.VARIANTS + self.NOT_BOUNDARY:
            with self.subTest(line=line):
                expected = gw_layout.is_draft_note_line(line)
                self.assertEqual(bool(gl.DRAFT_NOTE_RE.match(line)), expected)
                self.assertEqual(fact_diff.body_lines("正文。\n" + line + "\n说明。") == ["正文。"], expected)
        for line in self.VARIANTS:
            self.assertTrue(gw_layout.is_draft_note_line(line), line)

    def test_markdown_heading_stops_lint(self):
        for boundary in ("## 【起草说明】", "**【起草说明】**"):
            with self.subTest(boundary=boundary):
                result = run("全年完成检查12次。\n" + boundary + "\n以上内容仅供参考。\n")
                self.assertNotIn("DEF-001", ids(result))
                self.assertTrue(any("起草说明" in n for n in result["notes"]))


class YieldPerformanceTest(unittest.TestCase):
    """_resolve_yields 原为二次方：每条让位规则的命中都要遍历同段全部问题。

    原问题：一段 4 万字的“、、、……”产生 2 万条 PUN-020，约 40 秒；10 万字约 160 秒。
    """

    def test_many_issues_in_one_paragraph_is_fast(self):
        text = "、" * 40000
        started = time.perf_counter()
        result = run(text, checks="punct")
        elapsed = time.perf_counter() - started
        self.assertEqual(len(issues_of(result, "PUN-020")), 20000)
        self.assertLess(elapsed, 8.0)

    def test_yield_semantics_kept(self):
        # PUN-020 与 PUN-008 重叠时让位；不重叠的照常报告
        self.assertNotIn("PUN-020", ids(run("医院、学校。。。", checks="punct")))
        self.assertIn("PUN-020", ids(run("医院、学校。。。另外，，加强领导。", checks="punct")))


class QingshiEndingTest(unittest.TestCase):
    """“请予以批准”是请示最常用的结束语之一。

    原问题：请示以“特此请示，请予以批准。”结尾时误报 STR-003“缺少请求批复的结束语”；
    报告中写“请予以批准”时漏报 STR-004（报告夹带请示事项）。
    """

    QINGSHI = ("关于购置CT设备的请示\n\n市卫生健康委：\n\n我院现有CT设备已使用12年，故障频发。"
               "拟购置64排CT一台，预算约800万元，资金由医院自筹解决。\n\n{}\n\n某医院\n\n2026年10月16日")

    def test_qingshi_with_qing_yu_yi(self):
        for ending in ("特此请示，请予以批准。", "请予以审批。", "特此请示，请予以批复。"):
            with self.subTest(ending=ending):
                self.assertNotIn("STR-003", ids(run(self.QINGSHI.format(ending), "请示", "structure")))

    def test_qingshi_without_ending_still_reported(self):
        self.assertIn("STR-003", ids(run(self.QINGSHI.format("特此说明。"), "请示", "structure")))

    def test_baogao_with_qing_yu_yi_reported(self):
        text = "关于防汛工作的报告\n\n县人民政府：\n\n今年汛期我镇转移群众1260人次。以上报告，请予以批准。"
        self.assertIn("STR-004", ids(run(text, "报告", "structure")))


class RegexFalsePositiveTest(unittest.TestCase):
    """lint_rules.json 正则规则的误报。"""

    def test_pun017_standard_number(self):
        # 原问题：标准编号的年代号被当成数值范围，建议改成“9704～2012”，照改会改坏标准号
        for text in ("按照GB/T 9704-2012要求排版。", "执行GB 50016-2014标准。", "参照DB33/T 1234-2020规定。"):
            with self.subTest(text=text):
                self.assertNotIn("PUN-017", ids(run(text, checks="punct")))
        self.assertIn("PUN-017", ids(run("每批次抽检25-30份。", checks="punct")))
        self.assertIn("PUN-017", ids(run("预计 3-5 年完成。", checks="punct")))

    def test_leg001_mei_shou_dao(self):
        # 原问题：“没收到通知”“没收齐”里的“没收”被当成行政处罚
        for text in ("对没收到通知的住户，由网格员上门告知。", "材料没收齐的，暂缓办理。"):
            with self.subTest(text=text):
                self.assertNotIn("LEG-001", ids(run(text, "通知", "legal")))
        self.assertIn("LEG-001", ids(run("逾期不改的，没收全部违规物品。", "通知", "legal")))

    def test_leg003_zan_ting_fa_che(self):
        # 原问题：“暂停发车”“暂停发布”“暂停发售”里的“停发”被当成停发待遇
        for text in ("受道路施工影响，12路公交暂停发车。", "系统升级期间暂停发布信息。", "演出门票暂停发售。"):
            with self.subTest(text=text):
                self.assertNotIn("LEG-003", ids(run(text, "通告", "legal")))
        self.assertIn("LEG-003", ids(run("停发当月生活补贴。", "通告", "legal")))
        self.assertIn("LEG-003", ids(run("暂停发放生活补贴。", "通告", "legal")))

    def test_res012_lian_xi_wo_men(self):
        # 原问题：公示、告居民书常用的“请随时联系我们”被当成 AI 对话结束语，报 P0
        for text in ("如需帮助，请随时联系我们。", "欢迎随时联系我们社区居委会。",
                     "居民有任何问题可随时联系我们。", "有任何疑问请随时联系我们网格员。"):
            with self.subTest(text=text):
                self.assertNotIn("RES-012", ids(run(text, checks="residue")))
        for text in ("如需进一步修改，请随时告诉我。", "欢迎随时联系我。", "有任何问题随时问我。"):
            with self.subTest(text=text):
                self.assertIn("RES-012", ids(run(text, checks="residue")))


class EndToEndFindingsTest(unittest.TestCase):
    """examples/ 端到端实测发现的缺陷（复现场景见各目录的过程记录.md）。"""

    def test_year_after_dun_hao_is_level1_heading(self):
        # 原问题：“三、2027年工作打算”因“、”后是数字不算一级标题，其后的“（一）”被判乱序
        text = ("某科室2026年工作总结\n\n一、主要工作\n\n（一）医疗质量。内容。\n\n（二）学科建设。内容。\n\n"
                "二、存在的问题\n\n人员不足。\n\n三、2027年工作打算\n\n（一）补充人员。内容。\n\n（二）加强培训。内容。\n")
        self.assertEqual(gl.heading_of("三、2027年工作打算"), (1, 3))
        self.assertNotIn("STR-001", ids(run(text, "总结", "structure")))

    def test_markdown_bold_heading(self):
        # 原问题：“**二、时间安排**”不算一级标题，STR-001 连带误报；加粗符号仍按残留报告
        text = ("关于开展检查的通知\n\n各学校：\n\n现将有关事项通知如下：\n\n**一、检查内容**\n\n（一）校舍安全。\n\n"
                "（二）消防安全。\n\n**二、时间安排**\n\n（一）学校自查。\n\n（二）区级抽查。\n\n## 三、工作要求\n\n（一）加强领导。\n")
        result = run(text, "通知", "structure,residue")
        self.assertNotIn("STR-001", ids(result))
        self.assertEqual(len(issues_of(result, "RES-002")), 2)
        self.assertIn("RES-001", ids(result))

    def test_wenhao_between_title_and_zhusong(self):
        # 原问题：AI 稿把发文字号写在标题下，标题被当成发文机关标志，主送行被当成正文第一段，JFU-001 漏报
        text = ("某区教育局关于开展2026年冬季校园安全检查的通知\n\n某区教发[2026]第12号\n\n各中小学、幼儿园：\n\n"
                "随着冬季的到来，做好冬季校园安全检查工作具有十分重要的意义。现将有关事项通知如下：\n\n一、检查范围\n\n全区中小学。\n")
        lines = [(i, line, False) for i, line in enumerate(text.split("\n"), start=1)]
        doc = gl.Document("x", lines)
        self.assertEqual(doc.paras[doc.title[0]].text, "某区教育局关于开展2026年冬季校园安全检查的通知")
        self.assertEqual(doc.paras[doc.zhusong].text, "各中小学、幼儿园：")
        self.assertTrue(doc.paras[doc.first_body].text.startswith("随着冬季的到来"))
        self.assertTrue(doc.wenhao_present)
        result = run(text, None, "all")
        self.assertEqual(result["doc_type"], "通知")
        self.assertIn("JFU-001", ids(result))
        self.assertIn("FMT-001", ids(result))
        # 真正的版头（发文机关标志 + 发文字号）照常识别
        head = gl.Document("x", [(1, "某市人民政府办公室", False), (2, "某政办发〔2026〕5号", False),
                                 (3, "关于做好防汛工作的通知", False), (4, "各区县人民政府：", False), (5, "现通知如下。", False)])
        self.assertEqual(head.head, {0, 1})
        self.assertEqual(head.title, [2])
        self.assertEqual(head.zhusong, 3)

    def test_def002_with_subject(self):
        # 原问题：“请各学校结合实际情况调整”因“请”与“结合”之间有主语而漏报
        for text in ("以上内容仅供参考，请各学校结合实际情况调整。", "请你单位根据实际情况修改后使用。",
                     "请各地结合实际情况参考。"):
            with self.subTest(text=text):
                self.assertIn("DEF-002", ids(run(text, "通知", "defensive")))
        for text in ("请各学校结合实际制定实施方案。", "请各区结合实际情况调整实施细则。",
                     "申请人可根据实际情况调整。", "请各学校结合实际情况，制定具体检查方案。"):
            with self.subTest(text=text):
                self.assertNotIn("DEF-002", ids(run(text, "通知", "defensive")))

    def test_ai013_short_letter(self):
        # 原问题：350 字的函中“需要立即维修：一是……”和附注“（联系人：……）”两处冒号即超过每千字 3 处
        text = ("某区教育局关于商请追加校舍维修经费的函\n\n区财政局：\n\n"
                "9月开学安全检查中，又发现3所学校校舍存在安全隐患，经第三方检测机构鉴定，需要立即维修：一是区第二小学"
                "教学楼外墙瓷砖脱落，预算35万元；二是区实验中学学生宿舍楼电线老化，预算28万元；三是区第五小学一段围墙倾斜，"
                "预算17万元。以上3个项目共需经费80万元。为保障师生安全，恳请贵局从预备费中追加安排经费80万元。\n\n"
                "请予支持，并请函复。\n\n某区教育局\n\n2026年10月16日\n\n"
                "（联系人：区教育局计划财务科李某，联系电话：0571-86660000）\n")
        self.assertNotIn("AI-013", ids(run(text, "函", "style")))
        self.assertIn("AI-013", ids(run("目标：三年见效。路径：先试点再推广。", "函", "style")))

    def test_marker_level_and_summary(self):
        # 主 Agent 决定：【待补】【待核实】是起草阶段的正当产物，每处报 P2，summary 单独计数
        result = run("检查【待补：次数】次，完工于【待核实：完工时间】，整改【待补：个数】个。", checks="all")
        markers = [i for i in result["issues"] if i["id"] in ("MRK-001", "MRK-002")]
        self.assertEqual(len(markers), 3)
        self.assertTrue(all(i["level"] == "P2" for i in markers))
        self.assertEqual(result["summary"]["markers"], {"待补": 2, "待核实": 1})
        self.assertEqual(result["summary"]["P1"], 0)
        self.assertIn("待补 2 处、待核实 1 处，定稿前补齐", gl.render_text(result))
        # 不查 marker 类时也给出计数
        self.assertEqual(run("检查【待补：次数】次。", checks="punct")["summary"]["markers"], {"待补": 1, "待核实": 0})
        self.assertNotIn("待补 0 处", gl.render_text(run("检查12次。")))


if __name__ == "__main__":
    unittest.main()
