# -*- coding: utf-8 -*-
"""gongwen-humanize 改写力度相关评测与规则文本的一致性检查。

运行：python -m unittest discover -s tests/humanize -t .

检查三件事：
1. “该改到位”用例的材料文件存在，且期望中点名保留的数字确实出现在材料里；
2. 材料确有待处理的宏大开头和空泛总括句（用例不是空转）；
3. SKILL.md 与 rewrite-guide.md 保留了“改写强度”和（十七）至（十九）类的规则文本。
"""
from __future__ import annotations

import json
import re
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
SKILL_DIR = REPO_ROOT / "skills" / "gongwen-humanize"
TAG = "该改到位"

# 各用例材料中必须仍有的 AI 腔标志句，用于确认用例没有被误改成“已经很干净”的稿子
MUST_CONTAIN = {
    17: ["统筹推进老旧小区改造", "为基层治理注入了新的活力"],
    18: ["值得注意的是", "崭新篇章"],
    19: ["近年来，随着国家对粮食安全的重视程度不断提高", "令人瞩目的成绩"],
    20: ["高度重视，迅速召开专题会议", "筑牢了食品安全防线"],
    21: ["科学化、精细化、规范化", "极大增强了群众的参与感和认同感"],
    22: ["织密网络，阅读服务提质增效", "共赴书香之约", "扬帆起航"],
    23: ["加大监督检查力度", "总体稳定向好"],
}


def load_cases():
    data = json.loads((SKILL_DIR / "evals" / "evals.json").read_text(encoding="utf-8"))
    return data["evals"]


class RewriteStrengthEvalTest(unittest.TestCase):
    def setUp(self):
        self.cases = [c for c in load_cases() if TAG in c.get("tags", [])]

    def test_enough_cases(self):
        self.assertGreaterEqual(len(self.cases), 3)
        self.assertEqual({c["id"] for c in self.cases}, set(MUST_CONTAIN))

    def test_files_exist_and_numbers_kept(self):
        for case in self.cases:
            text = ""
            for rel in case["files"]:
                path = SKILL_DIR / rel
                self.assertTrue(path.is_file(), f"用例 {case['id']} 缺少材料 {rel}")
                text += path.read_text(encoding="utf-8")
            keep = [x for x in case["expectations"] if "不变" in x or "全部保留" in x]
            self.assertTrue(keep, f"用例 {case['id']} 缺少数字保留期望")
            for sentence in keep:
                for token in re.findall(r"\d+(?:\.\d+)?(?:万|亿)?(?:件|户|个|份|家|项|名|元|%|公里|公斤|平方米|次|起|年|月|日|小时|天)?", sentence):
                    number = re.match(r"\d+(?:\.\d+)?", token).group(0)
                    if len(number) >= 3 or "." in number:
                        self.assertIn(number, text, f"用例 {case['id']} 期望中的 {token} 不在材料里")

    def test_materials_still_have_ai_tone(self):
        for case in self.cases:
            text = "".join((SKILL_DIR / rel).read_text(encoding="utf-8") for rel in case["files"])
            for marker in MUST_CONTAIN[case["id"]]:
                self.assertIn(marker, text, f"用例 {case['id']} 材料缺少 AI 腔标志句：{marker}")

    def test_reply_structure_expected(self):
        for case in self.cases:
            joined = "".join(case["expectations"])
            self.assertIn("回复开头直接", joined, f"用例 {case['id']} 未要求开头直接给全文")


class RuleTextTest(unittest.TestCase):
    def test_skill_has_strength_section(self):
        text = (SKILL_DIR / "SKILL.md").read_text(encoding="utf-8")
        for needle in ["### 改写强度", "事实冻结、表达放开", "结构别动", "回复开头直接给改后全文", "不超过 5 条"]:
            self.assertIn(needle, text)

    def test_guide_has_new_categories_and_examples(self):
        text = (SKILL_DIR / "references" / "rewrite-guide.md").read_text(encoding="utf-8")
        for needle in ["### （十七）空泛总括句与段尾拔高句", "### （十八）抽象名词堆叠与副词密集",
                       "### （十九）空壳排比与机械总分", "## 六、改写力度：误正例"]:
            self.assertIn(needle, text)

    def test_skill_has_round3_rules(self):
        text = (SKILL_DIR / "SKILL.md").read_text(encoding="utf-8")
        for needle in ["改写优先于删除", "85%～100%", "标题也要改", "不在正文插【待补】", "拉开帷幕", "每条一行"]:
            self.assertIn(needle, text)

    def test_guide_has_round3_examples(self):
        text = (SKILL_DIR / "references" / "rewrite-guide.md").read_text(encoding="utf-8")
        for needle in ["5. 删除过度。", "6. 标题未改。", "7. 插入待补。"]:
            self.assertIn(needle, text)

    def test_round3_cases_expectations(self):
        cases = {c["id"]: c for c in load_cases()}
        joined22 = "".join(cases[22]["expectations"])
        self.assertIn("按写实改写", joined22)
        joined23 = "".join(cases[23]["expectations"])
        self.assertIn("没有被当作空话删掉", joined23)
        for cid in (5, 8, 9):
            self.assertNotIn("标【待补】", "".join(cases[cid]["expectations"]) + cases[cid]["expected_output"])

    def test_guide_no_longer_over_conservative(self):
        text = (SKILL_DIR / "references" / "rewrite-guide.md").read_text(encoding="utf-8")
        self.assertNotIn("改动最小。能删的不换", text)
        self.assertNotIn("拿不准就保留原文。", text)


if __name__ == "__main__":
    unittest.main()
