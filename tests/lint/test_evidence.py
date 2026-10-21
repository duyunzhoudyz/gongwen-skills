# -*- coding: utf-8 -*-
"""评价词与证据强度匹配（EVD-001～EVD-003）：必错标本、必过标本和白名单机制。"""
from __future__ import annotations

import unittest

from tests.lint.helpers import LINTER, gl, issues_of, run


def evd(text, doc_type="总结", **kw):
    result = gl.lint_text(text, doc_type, "style", linter=LINTER, **kw)
    return [i for i in result["issues"] if i["id"].startswith("EVD-")]


def evd_ids(text, doc_type="总结", **kw):
    return [i["id"] for i in evd(text, doc_type, **kw)]


class StrongEvaluationTest(unittest.TestCase):
    """强评价（EVD-001，P1）：无数字、无对比、无事例。"""

    HITS = (
        "全年工作成效显著，群众反响良好。",
        "我院医疗质量大幅提升，患者就医体验全面改善。",
        "经过努力，信访问题得到根本性解决。",
        "各项工作实现历史性突破。",
        "今年各项重点任务全面完成。",
        "改革举措广受好评。",
        "校园安全形势显著好转。",
        "营商环境发生翻天覆地的变化。",
    )

    def test_hits(self):
        for text in self.HITS:
            with self.subTest(text=text):
                found = evd(text)
                self.assertEqual([i["id"] for i in found], ["EVD-001"], found)
                self.assertEqual(found[0]["level"], "P1")

    def test_number_in_same_sentence(self):
        self.assertEqual(evd_ids("全年投诉量显著下降，由326件降至118件。"), [])
        self.assertEqual(evd_ids("门诊量大幅提升，增长23.5%。"), [])

    def test_number_in_next_sentence(self):
        self.assertEqual(evd_ids("取得明显成效。全年开展检查36次，整改隐患128处。"), [])

    def test_number_in_previous_sentence(self):
        self.assertEqual(evd_ids("全年检查36次，整改隐患128处。安全形势显著好转。"), [])

    def test_comparison_without_number(self):
        self.assertEqual(evd_ids("办事效率较去年显著提高。"), [])
        self.assertEqual(evd_ids("今年办件时限大幅压缩，与往年相比明显缩短。"), [])

    def test_evidence_after_heading(self):
        text = "（一）成效显著\n\n全年累计培训1200人次，覆盖全部乡镇。"
        self.assertEqual(evd_ids(text), [])

    def test_evidence_after_colon(self):
        text = "工作成效显著：\n\n一是接诉即办率达98%；二是满意率同比提高6个百分点。"
        self.assertEqual(evd_ids(text), [])

    def test_date_is_not_evidence(self):
        """年月日不算证据：2026年10月只是时间，不能支撑“显著”。"""
        self.assertEqual(evd_ids("2026年10月以来，工作成效显著。"), ["EVD-001"])

    def test_far_evidence_does_not_count(self):
        text2 = "全年检查36次。开展了专项行动。落实了责任。部署了新的任务。工作成效显著。"
        self.assertEqual(evd_ids(text2), ["EVD-001"])

    def test_case_only_is_p2(self):
        found = evd("服务质量显著提升。例如，某社区通过上门服务解决了独居老人的用药难题。")
        self.assertEqual([(i["id"], i["level"]) for i in found], [("EVD-003", "P2")])

    def test_one_issue_per_sentence(self):
        found = evd("成效显著，群众广受好评，工作大幅提升。")
        self.assertEqual(len(found), 1)

    def test_excerpt_and_context(self):
        found = evd("全年工作成效显著，群众反响良好。")[0]
        self.assertEqual(found["excerpt"], "成效显著")
        self.assertIn("全年工作", found["context"])


class MediumEvaluationTest(unittest.TestCase):
    """中评价（EVD-002，P2）：没有任何事实。"""

    HITS = (
        "校园环境明显改善。",
        "服务质量稳步提升。",
        "风险得到有效化解。",
        "取得了良好成效。",
        "群众获得感较大提升。",
    )

    def test_hits(self):
        for text in self.HITS:
            with self.subTest(text=text):
                found = evd(text)
                self.assertEqual([(i["id"], i["level"]) for i in found], [("EVD-002", "P2")], found)

    def test_any_fact_passes(self):
        self.assertEqual(evd_ids("校园环境明显改善，新增绿化面积1200平方米。"), [])
        self.assertEqual(evd_ids("校园环境明显改善。例如，东区操场完成了塑胶跑道翻新。"), [])
        self.assertEqual(evd_ids("今年的投诉比去年有明显改善。"), [])

    def test_strong_takes_precedence(self):
        found = evd("成效显著，校园环境明显改善。")
        self.assertEqual([i["id"] for i in found], ["EVD-001"])


class GoodUsageTest(unittest.TestCase):
    """公文常见的合理用法不误报。"""

    def test_obtained_effect_followed_by_data(self):
        text = "专项行动取得明显成效。累计排查隐患点位412处，整改完成率100%。"
        self.assertEqual(evd_ids(text), [])

    def test_goal_and_requirement_context(self):
        for text in (
            "力争今年大幅提升办件效率。",
            "确保各项指标显著改善。",
            "各单位要切实加强管理，大幅压缩办理时限。",
            "我们将大幅提升服务质量。",
            "为了显著改善办事环境，设置了咨询窗口。",
            "计划明年全面完成改造。",
            "目标是明显改善校园环境。",
        ):
            with self.subTest(text=text):
                self.assertEqual(evd_ids(text), [])

    def test_hedged_by_weak_words(self):
        for text in ("一定程度上明显改善了环境。", "初步显著改善。", "工作有所明显改善。"):
            with self.subTest(text=text):
                self.assertEqual(evd_ids(text), [])

    def test_quoted_evaluation(self):
        self.assertEqual(evd_ids("群众称赞“成效显著、广受好评”。"), [])
        self.assertEqual(evd_ids("评语为“大幅提升”，经办人签字确认。"), [])

    def test_whitelist_phrases(self):
        for text in (
            "公告张贴在显著位置。",
            "两地存在显著差异。",
            "该项目面临巨大压力。",
            "证件在有效期内使用。",
            "标识应设置在明显位置。",
            "项目存在较大风险。",
            "建立有效机制和有效措施。",
        ):
            with self.subTest(text=text):
                self.assertEqual(evd_ids(text), [])

    def test_placeholder_counts_as_noticed(self):
        self.assertEqual(evd_ids("工作成效显著，【待补：同比增幅】。"), [])
        self.assertEqual(evd_ids("工作成效显著。\n\n【待补：具体数据】"), [])

    def test_prose_without_evaluation(self):
        self.assertEqual(evd_ids("全年开展检查36次，整改隐患128处，约谈责任人5人。"), [])


class ScopeTest(unittest.TestCase):
    def test_not_checked_in_plan_and_speech(self):
        text = "全面完成改造任务，成效显著。"
        for dt in ("计划", "实施方案", "讲话稿", "工作要点", "纪要", "通知"):
            with self.subTest(doc_type=dt):
                self.assertEqual(evd_ids(text, dt), [])

    def test_checked_in_retrospective_documents(self):
        text = "全年工作成效显著。"
        for dt in ("总结", "汇报材料", "报告", "述职报告", "通报", "简报", "信息", "其他"):
            with self.subTest(doc_type=dt):
                self.assertEqual(evd_ids(text, dt), ["EVD-001"])

    def test_title_not_checked(self):
        text = "某区教育局安全工作成效显著情况总结\n\n全年检查36次。"
        self.assertEqual(evd_ids(text), [])

    def test_tables_skipped(self):
        text = "| 指标 | 评价 |\n|---|---|\n| 满意度 | 显著提升 |\n"
        self.assertEqual(evd_ids(text), [])


class AllowlistTest(unittest.TestCase):
    def test_evidence_allow_words(self):
        text = "全年工作成效显著。"
        self.assertEqual(evd_ids(text, evidence_allow="成效显著"), [])
        self.assertEqual(evd_ids(text, evidence_allow=["显著"]), [])
        self.assertEqual(evd_ids(text, evidence_allow="大幅，广受好评"), ["EVD-001"])

    def test_allow_does_not_leak_between_calls(self):
        text = "全年工作成效显著。"
        evd_ids(text, evidence_allow="显著")
        self.assertEqual(evd_ids(text), ["EVD-001"])

    def test_rule_level_allow_downgrades(self):
        result = gl.lint_text("全年工作成效显著。", "总结", linter=LINTER, allow=["EVD-001"])
        self.assertEqual([i["level"] for i in issues_of(result, "EVD-001")], ["P2"])

    def test_cli_option(self):
        import io
        import json
        import tempfile
        from contextlib import redirect_stdout
        from pathlib import Path
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "总结.txt"
            path.write_text("全年工作成效显著。\n", encoding="utf-8")
            for args, expected in (([], 1), (["--evidence-allow", "显著"], 0)):
                buf = io.StringIO()
                with redirect_stdout(buf):
                    gl.main([str(path), "--doc-type", "总结", "--checks", "style", "--json"] + args)
                data = json.loads(buf.getvalue())
                self.assertEqual(sum(1 for i in data["issues"] if i["id"] == "EVD-001"), expected)


class RegistryTest(unittest.TestCase):
    def test_registered(self):
        for rid, level in (("EVD-001", "P1"), ("EVD-002", "P2"), ("EVD-003", "P2")):
            with self.subTest(rule=rid):
                meta = LINTER.proc_meta[rid]
                self.assertEqual(meta["level"], level)
                self.assertEqual(meta["check"], "style")
                self.assertTrue(meta["basis"])

    def test_word_tiers_loaded(self):
        cfg = LINTER.evidence
        for key in ("strong", "medium", "weak"):
            self.assertTrue(cfg[key].pattern)
        self.assertTrue(cfg["whitelist"])
        self.assertEqual(LINTER.warnings, [])

    def test_not_run_when_style_not_selected(self):
        result = gl.lint_text("全年工作成效显著。", "总结", "format", linter=LINTER)
        self.assertEqual(issues_of(result, "EVD-001"), [])


if __name__ == "__main__":
    unittest.main()
