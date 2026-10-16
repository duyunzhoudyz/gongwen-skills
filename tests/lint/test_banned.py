# -*- coding: utf-8 -*-
"""Q10：--banned 本单位禁用词表。格式见 shared/profiles/profile-template.md 第三部分。"""
from __future__ import annotations

import json
import unittest

from tests.lint.helpers import LINTER, REPO_ROOT, gl
from tests.lint.test_cli import TempDirMixin, cli

TEMPLATE = """# 禁用词

| 词语或表述 | 级别 | 原因 | 替换建议 | 适用范围 |
|---|---|---|---|---|
| 赋能 | 慎用 | 领导多次删改 | 推动、支持、提升 | 全部 |
| 全国领先 | 禁用 | 没有权威认定依据 | 写具体排名并注明来源 | 全部 |
| 高度重视 | 慎用 | 段首套话 | 直接写采取的措施 | 讲话、汇报 |
| 治愈率 100% | 禁用 | 不得作疗效保证 | 写具体病例数和随访结果 | 新闻稿、公众号 |
| 抓手、着力点 | 慎用 | 院长改稿删除 3 次 | 措施、重点 | 全部 |

## 维护记录
- 2026-05-20：建表 5 条（gongwen-profile）
"""

LIST_STYLE = """# 禁用词

## 禁用
- 拉满
- 遥遥领先：没有依据

## 慎用
- 打造（领导不喜欢）
- “闭环”

## 说明
- 本表由办公室维护
"""

TEXT = ("某局2026年工作汇报材料\n\n今年我局高度重视营商环境，排名全国领先，以数字化赋能政务服务。"
        "下一步要找准抓手，形成闭环。\n\n【起草说明】\n全国领先的说法已删去。\n")


def ban_ids(result):
    return [(i["id"], i["excerpt"]) for i in result["issues"] if i["id"].startswith("BAN-")]


class LoadBannedTest(TempDirMixin, unittest.TestCase):
    def test_template_table(self):
        terms = gl.load_banned(self.write("禁用词.md", TEMPLATE))
        by_term = {t.term: t for t in terms}
        self.assertEqual(set(by_term), {"赋能", "全国领先", "高度重视", "治愈率 100%", "抓手", "着力点"})
        self.assertTrue(by_term["全国领先"].banned)
        self.assertFalse(by_term["赋能"].banned)
        self.assertEqual(by_term["高度重视"].scope, "讲话、汇报")
        self.assertEqual(by_term["全国领先"].suggestion, "写具体排名并注明来源")
        self.assertEqual(by_term["全国领先"].source, "禁用词.md")

    def test_list_style_with_sections(self):
        terms = {t.term: t.banned for t in gl.load_banned(self.write("禁用词.md", LIST_STYLE))}
        self.assertEqual(terms, {"拉满": True, "遥遥领先": True, "打造": False, "闭环": False})

    def test_profile_eval_sample(self):
        path = REPO_ROOT / "skills" / "gongwen-profile" / "evals" / "files" / "已有禁用词（某局）.md"
        terms = {t.term: t.banned for t in gl.load_banned(path)}
        self.assertEqual(terms, {"全国领先": False, "高度重视": False, "拉满": True})

    def test_gbk_file(self):
        terms = gl.load_banned(self.write("禁用词.md", LIST_STYLE, encoding="gbk"))
        self.assertIn("拉满", {t.term for t in terms})

    def test_missing_file(self):
        with self.assertRaises(gl.InputError):
            gl.load_banned(self.tmp / "不存在.md")

    def test_term_regex(self):
        self.assertTrue(gl._term_regex("治愈率 100%").search("治愈率100%"))
        self.assertTrue(gl._term_regex("在……领导下").search("在县委的坚强领导下"))


class BannedCheckTest(TempDirMixin, unittest.TestCase):
    def test_levels_and_scope(self):
        terms = gl.load_banned(self.write("禁用词.md", TEMPLATE))
        result = gl.lint_text(TEXT, "汇报材料", "style", linter=LINTER, banned=terms)
        found = {(i["id"], i["excerpt"]): i for i in result["issues"] if i["id"].startswith("BAN-")}
        self.assertEqual(set(found), {("BAN-002", "高度重视"), ("BAN-001", "全国领先"), ("BAN-002", "赋能"),
                                      ("BAN-002", "抓手")})
        self.assertEqual(found[("BAN-001", "全国领先")]["level"], "P1")
        self.assertEqual(found[("BAN-002", "赋能")]["level"], "P2")
        self.assertEqual(found[("BAN-001", "全国领先")]["check"], "style")
        self.assertEqual(found[("BAN-001", "全国领先")]["basis"], "本单位禁用词表（禁用词.md）")
        self.assertEqual(found[("BAN-002", "赋能")]["suggestion"], "推动、支持、提升")
        # 适用范围“讲话、汇报”：通知中不比对“高度重视”
        result = gl.lint_text(TEXT, "通知", "style", linter=LINTER, banned=terms)
        self.assertNotIn(("BAN-002", "高度重视"), ban_ids(result))
        self.assertTrue(any("适用范围" in n for n in result["notes"]))

    def test_not_run_without_style(self):
        terms = gl.load_banned(self.write("禁用词.md", TEMPLATE))
        result = gl.lint_text(TEXT, "汇报材料", "punct", linter=LINTER, banned=terms)
        self.assertEqual(ban_ids(result), [])
        self.assertTrue(any("禁用词表没有比对" in n for n in result["notes"]))

    def test_without_banned_option(self):
        self.assertEqual(ban_ids(gl.lint_text(TEXT, "汇报材料", "style", linter=LINTER)), [])

    def test_marker_and_drafting_note_ignored(self):
        terms = gl.load_banned(self.write("禁用词.md", LIST_STYLE))
        result = gl.lint_text("检查【待补：是否形成闭环】。\n\n【起草说明】\n拉满。\n", "其他", "style",
                              linter=LINTER, banned=terms)
        self.assertEqual(ban_ids(result), [])

    def test_cli(self):
        banned = self.write("禁用词.md", TEMPLATE)
        doc = self.write("正文.txt", TEXT)
        code, out, _err = cli(str(doc), "--doc-type", "汇报材料", "--banned", str(banned), "--json", "--strict")
        self.assertEqual(code, 0)
        data = json.loads(out)
        self.assertIn("BAN-001", [i["id"] for i in data["issues"]])
        self.assertTrue(any("已比对本单位禁用词表" in n for n in data["notes"]))

    def test_cli_missing_banned_file_returns_2(self):
        doc = self.write("正文.txt", TEXT)
        code, _out, err = cli(str(doc), "--banned", str(self.tmp / "不存在.md"))
        self.assertEqual(code, 2)
        self.assertIn("禁用词文件不存在", err)


if __name__ == "__main__":
    unittest.main()
