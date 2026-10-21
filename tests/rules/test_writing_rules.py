# -*- coding: utf-8 -*-
"""写作赛道第四轮修正：关键规则文本存在性检查。"""
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def read(rel: str) -> str:
    return (ROOT / rel).read_text(encoding="utf-8")


class WritingRulesTest(unittest.TestCase):
    def test_speech_numbers_not_rounded(self):
        s = read("skills/gongwen-speech/SKILL.md")
        self.assertIn("原值照写，不取整", s)
        self.assertNotIn("大数只往下取整", s)
        self.assertIn("预告几点，正文就是几点", s)

    def test_minutes_roster_at_head(self):
        s = read("skills/gongwen-minutes/SKILL.md")
        self.assertIn("出席：×××、×××；列席：×××；请假：×××", s)
        self.assertIn("正文后不再重复", s)

    def test_humanize_draft_mark_and_polish_mode(self):
        s = read("skills/gongwen-humanize/SKILL.md")
        self.assertIn("（初稿）", s)
        self.assertIn("85% 下限不适用", s)

    def test_date_signing_day(self):
        s = read("shared/guardrails.md")
        self.assertIn("成文日期按签发日填写", s)
        self.assertIn("年份类相对表述照抄", s)

    def test_notice_contact_and_han_closing(self):
        t = read("skills/gongwen-draft/references/doc-types/tongzhi.md")
        self.assertIn("联系人：×××，电话：×××", t)
        self.assertIn("回执表", t)
        h = read("skills/gongwen-draft/references/doc-types/han.md")
        self.assertIn("盼复", h)


if __name__ == "__main__":
    unittest.main()
