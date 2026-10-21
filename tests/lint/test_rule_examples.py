# -*- coding: utf-8 -*-
"""标本测试：lint_rules.json 和 style_patterns.json 中每条规则的必错、必过标本。"""
from __future__ import annotations

import re
import unittest

from tests.lint.helpers import LINT_PATH, LINTER, PATTERNS_PATH, RULES_PATH, gl, ids, load_json, run

RULES = load_json(RULES_PATH)
PATTERNS = load_json(PATTERNS_PATH)
VALID_CHECKS = set(gl.CHECK_ORDER)
VALID_GATES = {"政治关", "政策关", "权限关", "法规关", "文风关", "文字关", "格式关", "程序关"}


def _doc_type_for(item) -> str:
    """有正向限定的取第一个文种，否则按“其他”。"""
    types = item.get("doc_types") or []
    return types[0] if types else "其他"


class LintRulesSchemaTest(unittest.TestCase):
    def test_ids_unique(self):
        all_ids = [r["id"] for r in RULES["rules"]] + [p["id"] for p in RULES["procedural"]]
        self.assertEqual(len(all_ids), len(set(all_ids)), "规则 ID 重复")

    def test_required_fields(self):
        for rule in RULES["rules"]:
            with self.subTest(rule=rule["id"]):
                for key in ("id", "check", "category", "gate", "level", "pattern", "message", "suggestion", "basis"):
                    self.assertTrue(rule.get(key), f"缺少字段 {key}")
                self.assertIn(rule["check"], VALID_CHECKS - {"style", "defensive"})
                self.assertIn(rule["level"], ("P0", "P1", "P2"))
                self.assertIn(rule["gate"], VALID_GATES)
                self.assertEqual(rule["category"], gl.CHECK_NAMES[rule["check"]])
                self.assertIsInstance(rule["doc_types"], list)
                self.assertIsInstance(rule["exempt_doc_types"], list)
                re.compile(rule["pattern"])
                self.assertGreaterEqual(len(rule.get("hit_examples", [])), 1, "至少一个必错标本")
                self.assertGreaterEqual(len(rule.get("miss_examples", [])), 1, "至少一个必过标本")
        for item in RULES["procedural"]:
            with self.subTest(rule=item["id"]):
                for key in ("id", "check", "category", "gate", "level", "description", "basis"):
                    self.assertTrue(item.get(key), f"缺少字段 {key}")
                self.assertIn(item["gate"], VALID_GATES)

    def test_procedural_ids_registered(self):
        """引擎中用到的函数规则 ID 都必须在 procedural 中登记，反之亦然。"""
        source = LINT_PATH.read_text(encoding="utf-8")
        used = set(re.findall(r'_proc_issue\(\s*"([A-Z]{3}-\d{3})"', source))
        used |= set(re.findall(r'add\("([A-Z]{3}-\d{3})"', source))
        used |= set(re.findall(r'(?:emit|marker_or)\(\s*"([A-Z]{3}-\d{3})"', source))  # 要素完整性、评价证据检查
        registered = {p["id"] for p in RULES["procedural"]}
        self.assertEqual(used - registered, set(), "引擎使用了未登记的规则")
        self.assertEqual(registered - used, set(), "登记了引擎中没有实现的规则")

    def test_engine_loads_without_warning(self):
        self.assertEqual(LINTER.warnings, [])
        self.assertEqual(len(LINTER.rules), len(RULES["rules"]))
        self.assertEqual(len(LINTER.patterns), len(PATTERNS["patterns"]))


class LintRuleExamplesTest(unittest.TestCase):
    def test_hit_examples(self):
        for rule in RULES["rules"]:
            for example in rule["hit_examples"]:
                with self.subTest(rule=rule["id"], example=example):
                    result = run(example, _doc_type_for(rule), rule["check"])
                    self.assertIn(rule["id"], ids(result))

    def test_miss_examples(self):
        for rule in RULES["rules"]:
            for example in rule["miss_examples"]:
                with self.subTest(rule=rule["id"], example=example):
                    result = run(example, _doc_type_for(rule), rule["check"])
                    self.assertNotIn(rule["id"], ids(result))


class StylePatternExamplesTest(unittest.TestCase):
    """style_patterns.json 标本。

    所有规则（含密度规则）：必错标本单独成文时引擎必须报告，必过标本不得报告。
    密度规则的必错标本按约定写成至少命中 2 次；必过标本允许含 1 处命中（密度不足而不报告）。
    """

    def _raw_hits(self, pattern_obj, example: str) -> int:
        doc = gl.Document("<example>", [(1, example, False)])
        return len(LINTER.style_hits(doc, pattern_obj))

    def test_examples(self):
        by_id = {p.id: p for p in LINTER.patterns}
        for item in PATTERNS["patterns"]:
            pat = by_id[item["id"]]
            doc_type = _doc_type_for(item)
            for example in item.get("hit_examples", []):
                with self.subTest(rule=item["id"], example=example, kind="hit"):
                    self.assertIn(item["id"], ids(run(example, doc_type, item["group"])))
                    if item.get("density_per_1000") is not None:
                        self.assertGreaterEqual(self._raw_hits(pat, example), 2, "密度规则的必错标本至少命中 2 次")
            for example in item.get("miss_examples", []):
                with self.subTest(rule=item["id"], example=example, kind="miss"):
                    self.assertNotIn(item["id"], ids(run(example, doc_type, item["group"])))
                    if item.get("density_per_1000") is None:
                        self.assertEqual(self._raw_hits(pat, example), 0)

    def test_issue_category_is_rule_category(self):
        """style、defensive 的 issue 用规则自身的细类名称，不统一写成“AI味”“防御性”。"""
        cats = {p["id"]: p["category"] for p in PATTERNS["patterns"]}
        result = run("值得注意的是，台账更新不及时。以上内容仅供参考。", checks="style,defensive")
        found = {i["id"]: i for i in result["issues"]}
        self.assertEqual(found["AI-001"]["category"], cats["AI-001"])
        self.assertEqual(found["DEF-001"]["category"], cats["DEF-001"])
        self.assertEqual(found["AI-001"]["check"], "style")

    def test_whitelist_ids_unique_and_compile(self):
        wl_ids = [w["id"] for w in PATTERNS["whitelist"]]
        self.assertEqual(len(wl_ids), len(set(wl_ids)))
        for item in PATTERNS["whitelist"]:
            re.compile(item["pattern"])

    def test_new_whitelist_entries(self):
        self.assertNotIn("AI-030", ids(run("经认定为家庭经济特别困难，发生特别重大事故的，特别严重的要上报。", checks="style")))
        self.assertNotIn("AI-013", ids(run("开展检查【待补：次数】次，整改【待补：个数】个，完工于【待核实：时间】。", checks="style")))

    def test_ai052_not_exempt_for_speech(self):
        self.assertIn("AI-052", ids(run("兄弟们，四季度任务很重。", "讲话稿", "style")))
        self.assertNotIn("AI-042", ids(run("亲爱的朋友们，欢迎大家。", "讲话稿", "style")))

    def test_deploy_only_rules(self):
        text = "建议各单位于月底前完成自查。希望各单位能够高度重视。"
        self.assertTrue({"DEF-015", "DEF-021"} <= set(ids(run(text, "通知", "defensive"))))
        self.assertFalse({"DEF-015", "DEF-021"} & set(ids(run(text, "调研报告", "defensive"))))
        self.assertFalse({"DEF-015", "DEF-021"} & set(ids(run(text, "其他", "defensive"))))

    def test_def006_exempt_for_letter(self):
        self.assertIn("DEF-006", ids(run("具体以官方发布信息为准。", "通知", "defensive")))
        self.assertNotIn("DEF-006", ids(run("具体以官方发布信息为准。", "函", "defensive")))


if __name__ == "__main__":
    unittest.main()
