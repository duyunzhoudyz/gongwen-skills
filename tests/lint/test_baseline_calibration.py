# -*- coding: utf-8 -*-
"""依据公开公文语料基线 shared/style/baseline.md 对词表的校准：误报标本、文种豁免、阈值，以及 --baseline 选项。"""
from __future__ import annotations

import io
import json
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

from tests.lint.helpers import gl, ids, run

FILLER = "按时完成日常检查记录。"
CN_NUM = [*"一二三四五六七八九", "十"] + ["十" + c for c in "一二三四五六七八九"] + ["二十"]


def _padded(core: str, total: int) -> str:
    """在 core 之外补不含顿号、冒号和词表词的填充句，使全文非空白字数约为 total。"""
    need = max(total - len(core), 0)
    return FILLER * (need // len(FILLER) + 1) + core


class FalsePositiveSamplesTest(unittest.TestCase):
    """真实公文里出现过的正常写法，以及修正后仍应报告的写法。"""

    def test_ai004_planning_key_projects(self):
        for text in ("支持符合条件的机场项目纳入“十五五”民航发展规划重点项目。",
                     "加快推动就业促进规划重点任务落地落实。",
                     "依据国土空间规划协同开展科学选址，明确国土空间规划重点项目清单。"):
            with self.subTest(text=text):
                self.assertNotIn("AI-004", ids(run(text, "通知", "style")))
        self.assertIn("AI-004", ids(run("划重点，三项改革都要在年底前落地。", "通知", "style")))

    def test_ai018_fixed_expressions(self):
        for text in ("未达到100%的，每低一个百分点扣0.1分；比例低于50%（不含）的不得分。",
                     "于每季度第一个月15日前报送上一季度发生的违背市场准入负面清单相关规定的案例。",
                     "属于当事人一个违法行为同时违反两个以上法律条款的情形。"):
            with self.subTest(text=text):
                self.assertNotIn("AI-018", ids(run(text, "通知", "style")))
        self.assertIn("AI-018", ids(run("这是一个能够让基层在不增加人手的情况下显著提升办事效率的工具。", "通知", "style")))

    def test_ai022_lighthouse(self):
        for text in ("积极培育灯塔工厂和专精特新企业。", "灯塔市柳河子镇入选国家农业产业强镇。"):
            with self.subTest(text=text):
                self.assertNotIn("AI-022", ids(run(text, "通知", "style")))
        self.assertIn("AI-022", ids(run("这项政策是照亮改革的灯塔。", "通知", "style")))

    def test_def012_def022_directive_degree(self):
        text = "口岸城市要加强组织，尽可能提高新冠病毒疫苗全程接种覆盖率。"
        self.assertNotIn("DEF-012", ids(run(text, "通知", "defensive")))
        self.assertNotIn("DEF-022", ids(run(text, "通知", "defensive")))
        for text in ("尽量依托高校现有场地建站，避免重复建设。", "帮助劳动者尽可能稳在企业。"):
            with self.subTest(text=text):
                self.assertNotIn("DEF-022", ids(run(text, "实施方案", "defensive")))
        # 回避表态、放松时限的写法仍要报
        self.assertIn("DEF-012", ids(run("该措施实施后可能会提升审批效率。", "通知", "defensive")))
        self.assertIn("DEF-022", ids(run("请各单位尽量按时报送。", "通知", "defensive")))
        self.assertIn("DEF-022", ids(run("尽量在本月内完成整改。", "通知", "defensive")))


class DocTypeExemptionTest(unittest.TestCase):
    CASES = {
        "AI-023": "全年工作取得明显成效。",
        "AI-028": "这是一个具有里程碑意义的重大决策。",
        "AI-029": "人才队伍建设至关重要，制度保障不可或缺。",
    }

    def test_speech_and_report_exempt(self):
        for rid, text in self.CASES.items():
            for doc_type in ("讲话稿", "党课讲稿", "报告"):
                with self.subTest(rule=rid, doc_type=doc_type):
                    self.assertNotIn(rid, ids(run(text, doc_type, "style")))

    def test_other_types_still_reported(self):
        for rid, text in self.CASES.items():
            for doc_type in ("总结", "汇报材料", "信息", "其他"):
                with self.subTest(rule=rid, doc_type=doc_type):
                    self.assertIn(rid, ids(run(text, doc_type, "style")))


class ThresholdTest(unittest.TestCase):
    def test_ai051_dunhao_45(self):
        # 每千字约 40 个顿号：真实通报、函里常见，不报；约 60 个：报
        low = _padded("甲、乙。" * 40, 1000)
        high = _padded("甲、乙。" * 60, 1000)
        self.assertNotIn("AI-051", ids(run(low, "通报", "style")))
        self.assertIn("AI-051", ids(run(high, "通报", "style")))
        self.assertIn("AI-051", ids(run("教育、卫生、民政、人社、住建、交通、水利、应急、市场监管等部门要按职责分工抓好落实。"
                                        "要统筹推进教育、医疗、养老、托育、就业、住房、交通、文化、体育等民生事项。", "通知", "style")))

    def test_ai026_threshold_2_5(self):
        words = "以数字赋能基层治理，形成服务闭环。"
        base = _padded("", 850 - len(words))
        two = base + words          # 2 处、约 2.4 次/千字：真实意见、方案里常见，不报
        self.assertNotIn("AI-026", ids(run(two, "通知", "style")))
        three = base + words + "同时助力基层，打造样板。"   # 4 处、约 4.7 次/千字，报
        self.assertIn("AI-026", ids(run(three, "通知", "style")))

    def test_ai013_colon_threshold_4(self):
        three = _padded("核心问题：机制不顺。主要原因：协调不够。工作目标：三年见效。", 800)   # 约 3.8 次/千字
        six = _padded("核心问题：机制不顺。主要原因：协调不够。工作目标：三年见效。"
                      "保障措施：专班推进。责任单位：各区。完成时限：年底。", 800)           # 约 7.5 次/千字
        self.assertNotIn("AI-013", ids(run(three, "通报", "style")))
        self.assertIn("AI-013", ids(run(six, "通报", "style")))

    def test_threshold_basis_recorded(self):
        data = json.loads((Path(gl.__file__).parent / "style_patterns.json").read_text(encoding="utf-8"))
        by_id = {p["id"]: p for p in data["patterns"]}
        for rid, value in (("AI-051", 45), ("AI-026", 2.5), ("AI-013", 4)):
            with self.subTest(rule=rid):
                self.assertEqual(by_id[rid]["density_per_1000"], value)
                self.assertIn("shared/style/baseline.md", by_id[rid]["source"])
        for rid in ("AI-023", "AI-028", "AI-029"):
            with self.subTest(rule=rid):
                self.assertIn("shared/style/baseline.md", by_id[rid]["source"])


def _with_dunhao(total: int) -> str:
    """每 46 字左右带 1 个顿号（约 22 个/千字，接近真实通知的中位数）。"""
    unit = "检查教育、卫生等重点场所。" + FILLER * 3
    return unit * (total // len(unit) + 1)


def _notice(chars: int, h1: int = 3) -> str:
    """造一篇通知：标题、主送、h1 个一级标题、补足到约 chars 字。"""
    lines = ["关于做好安全检查工作的通知", "各区县人民政府，市政府各部门："]
    for i in range(h1):
        lines.append(f"{CN_NUM[i]}、工作要求")
        lines.append(_with_dunhao(max(chars // max(h1, 1) - 40, 20)))
    return "\n".join(lines)


class BaselineOptionTest(unittest.TestCase):
    def setUp(self):
        self.ref = gl.load_baseline()

    def test_repo_baseline_found(self):
        self.assertIsNotNone(self.ref.data, self.ref.note)
        self.assertIn("p10", self.ref.data["groups"]["通知"]["chars"])

    def test_within_range_no_hint(self):
        result = gl.lint_text(_notice(1800), "通知", "length", baseline=self.ref, linter=None)
        self.assertEqual([i for i in result["issues"] if i["id"].startswith("BSL-")], [])
        self.assertEqual(result["baseline"]["group"], "通知")

    def test_too_long_notice_hint_p2(self):
        result = gl.lint_text(_notice(9000), "通知", "length", baseline=self.ref)
        found = {i["id"]: i for i in result["issues"] if i["id"].startswith("BSL-")}
        self.assertIn("BSL-001", found)
        self.assertEqual(found["BSL-001"]["level"], "P2")
        self.assertIn("P90", found["BSL-001"]["message"])
        self.assertEqual(result["summary"]["P0"], 0)

    def test_yishi_and_headings_in_notice(self):
        text = _notice(1800, h1=3) + "\n一是加强领导。二是压实责任。\n一是强化督导。二是严格考核。\n一是完善机制。二是提升效能。"
        result = gl.lint_text(text, "通知", "length", baseline=self.ref)
        self.assertIn("BSL-003", ids(result))
        many = gl.lint_text(_notice(1800, h1=20), "通知", "length", baseline=self.ref)
        self.assertIn("BSL-002", ids(many))

    def test_print_forward_notice_uses_own_group(self):
        text = "关于印发《安全生产工作方案》的通知\n各区县人民政府：\n现将《安全生产工作方案》印发给你们，请认真贯彻落实。"
        result = gl.lint_text(text, "通知", "length", baseline=self.ref)
        self.assertEqual(result["baseline"]["group"], "通知（印发转发）")

    def test_unrepresentative_types_skipped_with_note(self):
        for doc_type in ("总结", "其他"):
            with self.subTest(doc_type=doc_type):
                result = gl.lint_text(_notice(3000), doc_type, "length", baseline=self.ref)
                self.assertNotIn("baseline", result)
                self.assertEqual([i for i in result["issues"] if i["id"].startswith("BSL-")], [])
                self.assertTrue(any("未比较" in n for n in result["notes"]))

    def test_without_option_nothing_added(self):
        result = gl.lint_text(_notice(9000), "通知", "length")
        self.assertNotIn("baseline", result)
        self.assertEqual([i for i in result["issues"] if i["id"].startswith("BSL-")], [])

    def test_missing_baseline_is_silent_skip(self):
        ref = gl.load_baseline("不存在的路径/baseline.json")
        self.assertIsNone(ref.data)
        result = gl.lint_text(_notice(9000), "通知", "length", baseline=ref)
        self.assertEqual([i for i in result["issues"] if i["id"].startswith("BSL-")], [])
        self.assertTrue(any("跳过" in n for n in result["notes"]))

    def test_cli_option(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "通知.txt"
            path.write_text(_notice(9000), encoding="utf-8")
            buf = io.StringIO()
            with redirect_stdout(buf):
                code = gl.main([str(path), "--doc-type", "通知", "--checks", "length", "--baseline", "--strict"])
            self.assertEqual(code, 0)  # P2 提示不影响退出码
            self.assertIn("BSL-001", buf.getvalue())
            buf = io.StringIO()
            with redirect_stdout(buf):
                code = gl.main([str(path), "--doc-type", "通知", "--checks", "length",
                                "--baseline", str(Path(tmp) / "无.json"), "--json"])
            self.assertEqual(code, 0)
            self.assertNotIn("BSL-001", buf.getvalue())


if __name__ == "__main__":
    unittest.main()
