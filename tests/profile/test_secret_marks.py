# -*- coding: utf-8 -*-
"""skills/gongwen-profile/scripts/secret_marks.py 的单元测试。

运行：python -m unittest discover -s tests/profile -t .

必错标本：密级标识、各种"内部资料"写法、"此件……"、"未公开""未披露"，以及 docx 页眉页脚中的标注，必须报出。
必过标本：正文中提到保密工作、引号内举例的标注，不得报出。
规则来源：lint_rules.json 读不到或损坏时，用内置副本，结果不变；内置副本与共享机检规则保持一致。
"""
from __future__ import annotations

import json
import unittest
from pathlib import Path

from tests.profile.helpers import REPO_ROOT, ScriptTestCase, make_docx

import secret_marks  # noqa: E402  （helpers 已把脚本目录加入 sys.path）

MUST_FLAG = {
    "秘密★1年": "密级标识",
    "绝密★启用前": "密级标识",
    "密级：机密": "密级标识",
    "机密（5年）": "密级标识",
    "内部资料": "内部标注",
    "内部资料　注意保存": "内部标注",
    "内部资料，严禁外传": "内部标注",
    "（内部资料 不得翻印）": "内部标注",
    "【内部资料　不得复制】": "内部标注",
    "此件为工作秘密，注意保管。": "内部标注",
    "本材料属工作秘密，不得外传。": "内部标注",
    "内部参阅": "内部标注",
    "内部传阅": "内部标注",
    "本报告仅供内部参阅": "内部标注",
    "上述经营数据为未公开的财务数据，请注意保密。": "未公开或未披露信息",
    "以上指标均系尚未披露的季度数据。": "未公开或未披露信息",
    # 评审补充的漏报写法
    "内部资料（注意保存）": "内部标注",
    "内部资料(注意保存)": "内部标注",
    "【内部资料】（不得外传）": "内部标注",
    "仅供内部参考": "内部标注",
    "（仅供内部使用）": "内部标注",
    "内部参考": "内部标注",
    "本报告数据尚未公开披露": "未公开或未披露信息",
    "以上指标未对外公开，请勿转发。": "未公开或未披露信息",
}

MUST_PASS = (
    "严守工作秘密，不在非涉密计算机上处理工作秘密。",
    "编印《保密知识读本》（内部资料）2000册，发放到各科室。",
    "加强内部资料管理，规范标注“内部资料”字样。",
    "规范标注“秘密★1年”等密级标识。",
    "全年开展保密工作检查4次，加强涉密人员管理。",
    "严格遵守保守国家秘密法。",
    "不得泄露商业秘密。",
    "坚持“内部资料，严禁外传”的管理要求。",
    "公司严禁外传客户信息。",
    "该企业未披露环境信息，已依法立案调查。",
    "对3家未公开招标的项目开展专项检查。",
    "以上意见仅供参考。",
    "建立内部参考资料借阅登记制度。",
    "规范标注“仅供内部参考”字样。",
)


class MustFlagSpecimens(unittest.TestCase):
    def test_marks_found_with_category(self):
        for line, label in MUST_FLAG.items():
            with self.subTest(line=line):
                self.assertIn(label, secret_marks.find_marks(["某公司经营情况报告", line, "营业收入23.6亿元。"]))

    def test_only_categories_returned(self):
        found = secret_marks.find_marks(["秘密★1年", "内部资料"])
        self.assertEqual(found, ["密级标识", "内部标注"])


class MustPassSpecimens(unittest.TestCase):
    def test_body_mentions_not_flagged(self):
        for line in MUST_PASS:
            with self.subTest(line=line):
                self.assertEqual(secret_marks.find_marks([line]), [])


class RuleSource(ScriptTestCase):
    def test_builtin_copy_matches_shared_lint_rules(self):
        data = json.loads((REPO_ROOT / "shared" / "scripts" / "lint_rules.json").read_text(encoding="utf-8"))
        rules = {r["id"]: r for r in data["rules"] if r.get("id") in secret_marks.RULE_IDS}
        self.assertEqual(set(rules), set(secret_marks.RULE_IDS))
        for rid, rule in rules.items():
            with self.subTest(rule=rid):
                self.assertEqual(secret_marks.BUILTIN[rid], rule["pattern"])

    def test_shared_lint_examples_respected(self):
        data = json.loads((REPO_ROOT / "shared" / "scripts" / "lint_rules.json").read_text(encoding="utf-8"))
        for rule in data["rules"]:
            if rule.get("id") not in secret_marks.RULE_IDS:
                continue
            for ex in rule.get("hit_examples", []):
                with self.subTest(rule=rule["id"], hit=ex):
                    self.assertTrue(secret_marks.find_marks([ex]))
            for ex in rule.get("miss_examples", []):
                with self.subTest(rule=rule["id"], miss=ex):
                    self.assertEqual(secret_marks.find_marks([ex]), [])

    def test_fallback_when_rules_file_missing_or_broken(self):
        missing = secret_marks.load_rules(self.tmp / "不存在.json")
        broken_path = self.write("坏.json", "{不是 JSON")
        broken = secret_marks.load_rules(broken_path)
        bad_pattern = self.write("坏规则.json", json.dumps(
            {"rules": [{"id": "PRI-005", "pattern": "(未闭合"}]}, ensure_ascii=False))
        partial = secret_marks.load_rules(bad_pattern)
        normal = secret_marks.load_rules()
        for rules in (missing, broken, partial):
            self.assertEqual([r.pattern for _, r in rules], [r.pattern for _, r in normal])

    def test_rules_file_takes_priority(self):
        custom = self.write("规则.json", json.dumps(
            {"rules": [{"id": "PRI-005", "pattern": "^测试密级$"}]}, ensure_ascii=False))
        rules = secret_marks.load_rules(custom)
        self.assertEqual(rules[0][1].pattern, "^测试密级$")


class DocxParts(ScriptTestCase):
    def test_header_footer_and_deleted_text_read(self):
        path = make_docx(self.tmp / "a.docx", [[("t", "正文"), ("del", "删去的字", "甲", "2026-10-14T08:00:00Z")]],
                         header="内部资料", footer="第1页")
        lines = secret_marks.docx_mark_lines(path)
        self.assertIn("内部资料", lines)
        self.assertIn("第1页", lines)
        self.assertIn("正文删去的字", lines)

    def test_unreadable_docx_returns_empty(self):
        bad = self.write("坏.docx", "不是压缩包")
        self.assertEqual(secret_marks.docx_mark_lines(bad), [])
        self.assertEqual(secret_marks.docx_mark_lines(Path(self.tmp / "不存在.docx")), [])


if __name__ == "__main__":
    unittest.main()
