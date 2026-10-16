# -*- coding: utf-8 -*-
"""tools/validate.py 的单元测试。"""
from __future__ import annotations

import io
import json
import unittest
import zipfile
from pathlib import Path

from tests.tools.fixtures import TempRepoTestCase, make_trigger, run_quiet, skill_md, write_json, write_text

import validate


class ValidateTestBase(TempRepoTestCase):

    def run_checks(self, skill=None):
        rep, _count = validate.run_validation(self.root, skill)
        return rep

    def assert_has(self, items, fragment):
        messages = [m for _scope, m in items]
        self.assertTrue(any(fragment in m for m in messages), f"未找到 \"{fragment}\"：{messages}")

    def assert_not_has(self, items, fragment):
        messages = [m for _scope, m in items]
        self.assertFalse(any(fragment in m for m in messages), f"不应出现 \"{fragment}\"：{messages}")

    def edit_manifest(self, changes, skill="gongwen-demo"):
        path = self.repo.skill_dir(skill) / "manifest.json"
        data = json.loads(path.read_text(encoding="utf-8"))
        data.update(changes)
        write_json(path, data)


class GoodRepoTest(ValidateTestBase):

    def test_empty_repo(self):
        code, out = run_quiet(validate.main, ["--root", str(self.root)])
        self.assertEqual(code, 0, out)
        self.assertIn("还没有任何 skill", out)
        self.assertIn("marketplace.json", out)

    def test_good_skill_passes_without_warnings(self):
        self.repo.add_skill()
        self.repo.add_marketplace()
        rep = self.run_checks()
        self.assertEqual(rep.errors, [])
        self.assertEqual(rep.warnings, [])
        code, out = run_quiet(validate.main, ["--root", str(self.root)])
        self.assertEqual(code, 0, out)

    def test_skill_option_and_quiet(self):
        self.repo.add_skill()
        code, out = run_quiet(validate.main, ["--root", str(self.root), "--skill", "gongwen-demo", "--quiet"])
        self.assertEqual(code, 0, out)
        self.assertNotIn("警告", out.replace("警告 0 个", ""))
        code, out = run_quiet(validate.main, ["--root", str(self.root), "--skill", "nope"])
        self.assertEqual(code, 2)


class FrontmatterTest(ValidateTestBase):

    def test_name_rules(self):
        self.repo.add_skill("gongwen-demo", skill_text=skill_md("Gongwen_Demo"))
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "不合规")
        self.assert_has(rep.errors, "与目录名")

    def test_reserved_word(self):
        self.repo.add_skill("claude-helper", skill_text=skill_md("claude-helper"))
        rep = self.run_checks("claude-helper")
        self.assert_has(rep.errors, "保留词 claude")

    def test_name_prefix(self):
        # 规范第 2 节：统一 gongwen- 前缀，总入口为 gongwen
        self.repo.add_skill("report-helper", skill_text=skill_md("report-helper"))
        rep = self.run_checks("report-helper")
        self.assert_has(rep.errors, "gongwen- 开头")
        self.repo.add_skill("gongwenx", skill_text=skill_md("gongwenx"))
        rep = self.run_checks("gongwenx")
        self.assert_has(rep.errors, "gongwen- 开头")
        for ok in ("gongwen", "gongwen-report"):
            self.repo.add_skill(ok, skill_text=skill_md(ok))
            rep = self.run_checks(ok)
            self.assertEqual(rep.errors, [], ok)

    def test_name_too_long(self):
        name = "g" + "-x" * 40
        self.repo.add_skill(name, skill_text=skill_md(name))
        rep = self.run_checks(name)
        self.assert_has(rep.errors, "超过 64 字符")

    def test_multiline_description(self):
        text = skill_md("gongwen-demo").replace(
            "description: 演示用", "description: >\n  演示用"
        )
        self.repo.add_skill(skill_text=text)
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "一行内")

    def test_description_length(self):
        self.repo.add_skill(skill_text=skill_md("gongwen-demo", description="很" * 301))
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.warnings, "超过 300 字")
        self.assertEqual(rep.errors, [])

        self.repo.add_skill(skill_text=skill_md("gongwen-demo", description="很" * 1025))
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "超过 1024 字符")

    def test_yaml_colon_in_description(self):
        self.repo.add_skill(skill_text=skill_md("gongwen-demo", description="用法: 写材料。"))
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "半角冒号")

    def test_extra_field_and_missing_license(self):
        text = skill_md("gongwen-demo", extra_fields="version: 1.0\n").replace("license: MIT\n", "")
        self.repo.add_skill(skill_text=text)
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.warnings, "规范外字段 version")
        self.assert_has(rep.errors, "缺少 license")

    def test_missing_frontmatter(self):
        self.repo.add_skill(skill_text="# 没有 frontmatter\n")
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "缺少 frontmatter")

    def test_split_frontmatter_quotes(self):
        fields, body, errors, has_fm = validate.split_frontmatter(
            "---\nname: a\ndescription: \"含: 冒号\"\n---\n正文\n"
        )
        self.assertTrue(has_fm)
        self.assertEqual(errors, [])
        self.assertEqual(fields["description"], "含: 冒号")
        self.assertEqual(body, "正文\n")


class SkillMdTest(ValidateTestBase):

    def test_line_limits(self):
        self.repo.add_skill(skill_text=skill_md("gongwen-demo", extra_body="行\n" * 290))
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.warnings, "超过建议值 300 行")
        self.repo.add_skill(skill_text=skill_md("gongwen-demo", extra_body="行\n" * 500))
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "超过上限 500 行")

    def test_missing_relative_path(self):
        body = (
            "\n见 `references/missing.md`，再看 references/_shared/guardrails.md第一条。\n"
            "```text\nreferences/<文种>.md\nscripts/*.py\nassets/{name}.yaml\n```\n"
            "目录 `references/` 和 `scripts/nothing/` 。\n"
        )
        self.repo.add_skill(skill_text=skill_md("gongwen-demo", extra_body=body))
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "references/missing.md")
        self.assert_has(rep.errors, "scripts/nothing/")
        # 紧贴中文的已存在路径、占位符、已存在目录都不报
        self.assert_not_has(rep.errors, "guardrails.md")
        self.assert_not_has(rep.errors, "<")
        self.assert_not_has(rep.errors, "*")
        self.assert_not_has(rep.errors, "{")
        self.assertEqual(len(rep.errors), 2, rep.errors)

    def test_case_sensitive_path(self):
        body = "\n见 `references/Checklist.md`。\n"
        self.repo.add_skill(skill_text=skill_md("gongwen-demo", extra_body=body))
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "references/Checklist.md")

    def test_missing_readme(self):
        d = self.repo.add_skill()
        (d / "README.md").unlink()
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "缺少 README.md")


class ManifestTest(ValidateTestBase):

    def test_bad_fields(self):
        self.repo.add_skill()
        self.edit_manifest({"status": "alpha", "version": "v1", "name": "other"})
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "status")
        self.assert_has(rep.errors, "语义化版本号")
        self.assert_has(rep.errors, "manifest.json 的 name")

    def test_invalid_json(self):
        d = self.repo.add_skill()
        write_text(d / "manifest.json", "{\"name\": ")
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "JSON 解析失败")

    def test_required_shared(self):
        self.repo.add_skill(shared=[{"src": "guardrails.md"}])
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "必须包含 style/ai-tone.md")
        self.assert_has(rep.errors, "必须包含 style/defensive.md")

    def test_missing_src(self):
        shared = [{"src": s} for s in ("guardrails.md", "style/output-format.md",
                                        "style/ai-tone.md", "style/defensive.md", "rules/nope.md")]
        self.repo.add_skill(shared=shared, sync=False)
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "shared/rules/nope.md 不存在")

    def test_sync_drift(self):
        self.repo.add_skill()
        write_text(self.root / "shared/style/ai-tone.md", "# 改过了\n")
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "共享资料未同步")
        self.assert_has(rep.errors, "references/_shared/style/ai-tone.md")

    def test_prompt_bundle_missing(self):
        self.repo.add_skill(bundle=["references/_shared/guardrails.md", "references/nope.md"])
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "prompt_bundle 引用的文件不存在")


class EvalsTest(ValidateTestBase):

    def test_few_evals_warning_and_draft(self):
        self.repo.add_skill(evals_count=3)
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.warnings, "只有 3 条")
        self.repo.add_skill(evals_count=3, status="draft")
        rep = self.run_checks("gongwen-demo")
        self.assert_not_has(rep.warnings, "只有 3 条")

    def test_eval_structure_and_files(self):
        d = self.repo.add_skill()
        path = d / "evals/evals.json"
        data = json.loads(path.read_text(encoding="utf-8"))
        data["evals"][0]["files"] = ["evals/files/不存在.txt"]
        data["evals"][1]["id"] = data["evals"][0]["id"]
        del data["evals"][2]["prompt"]
        write_json(path, data)
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "evals/files/不存在.txt")
        self.assert_has(rep.errors, "重复")
        self.assert_has(rep.errors, "prompt 必须是非空字符串")

    def test_invalid_evals_json(self):
        d = self.repo.add_skill()
        write_text(d / "evals/evals.json", "[")
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "evals/evals.json：JSON 解析失败")

    def test_trigger_checks(self):
        d = self.repo.add_skill()
        write_json(d / "evals/trigger.json", make_trigger(12, 6))
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.warnings, "共 18 条")
        self.assert_has(rep.warnings, "不平衡")
        write_json(d / "evals/trigger.json", [{"query": "缺字段"}])
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "should_trigger")
        write_json(d / "evals/trigger.json", {"query": "不是数组"})
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "顶层必须是数组")


REQUIRED = [{"src": s} for s in ("guardrails.md", "style/output-format.md", "style/ai-tone.md", "style/defensive.md")]


class SharedClosureTest(ValidateTestBase):
    """D18 / CM-1：references/_shared 中各文件互相引用的共享资料必须同步进来。"""

    def setUp(self) -> None:
        super().setUp()
        # 复现 v0.1 CM-1：guardrails.md 第四条写“详见 references/_shared/rules/jianfu.md”，但 skill 没有同步 jianfu.md
        write_text(self.root / "shared/guardrails.md",
                   "# 六条写作红线\n\n## 四、篇幅\n\n详见 references/_shared/rules/jianfu.md第三条。\n")
        write_text(self.root / "shared/rules/jianfu.md",
                   "# 减负规定\n\n格式见 `references/_shared/rules/gbt9704.md`。\n")

    def closure_errors(self, rep):
        return [m for _s, m in rep.errors if m.startswith("共享资料引用不闭合")]

    def test_missing_shared_reference_reported(self):
        self.repo.add_skill()
        rep = self.run_checks("gongwen-demo")
        self.assertEqual(self.closure_errors(rep), [
            "共享资料引用不闭合：references/_shared/guardrails.md 引用 references/_shared/rules/jianfu.md，"
            "请在 manifest.json 的 shared 中补上 rules/jianfu.md"])
        code, out = run_quiet(validate.main, ["--root", str(self.root), "--skill", "gongwen-demo"])
        self.assertEqual(code, 1, out)

    def test_closed_after_adding_to_manifest(self):
        # 补上 jianfu.md 后，jianfu.md 又引用 gbt9704.md，继续报；两个都补上后通过
        self.repo.add_skill(shared=REQUIRED + [{"src": "rules/jianfu.md"}])
        rep = self.run_checks("gongwen-demo")
        self.assertEqual(len(self.closure_errors(rep)), 1)
        self.assert_has(rep.errors, "references/_shared/rules/jianfu.md 引用 references/_shared/rules/gbt9704.md")
        self.repo.add_skill(shared=REQUIRED + [{"src": "rules/jianfu.md"}, {"src": "rules/gbt9704.md"}])
        rep = self.run_checks("gongwen-demo")
        self.assertEqual(rep.errors, [])
        self.assertEqual(rep.warnings, [])

    def test_referrers_grouped_and_skill_md_included(self):
        body = "\n格式要求见 references/_shared/rules/jianfu.md。\n"
        d = self.repo.add_skill(skill_text=skill_md("gongwen-demo", extra_body=body))
        write_text(d / "references/checklist.md", "# 自检清单\n\n- [ ] 篇幅符合 `references/_shared/rules/jianfu.md`\n")
        rep = self.run_checks("gongwen-demo")
        errors = self.closure_errors(rep)
        self.assertEqual(len(errors), 1, errors)
        for who in ("SKILL.md", "references/_shared/guardrails.md", "references/checklist.md"):
            self.assertIn(who, errors[0])
        # SKILL.md 中的同一路径不再按“路径不存在”重复报告
        self.assert_not_has(rep.errors, "SKILL.md 引用的路径不存在")

    def test_source_missing_in_shared(self):
        write_text(self.root / "shared/guardrails.md", "# 红线\n\n见 references/_shared/rules/nope.md。\n")
        self.repo.add_skill()
        rep = self.run_checks("gongwen-demo")
        self.assert_has(rep.errors, "shared/rules/nope.md 也不存在")

    def test_directory_placeholder_and_evals(self):
        write_text(self.root / "shared/guardrails.md",
                   "# 红线\n\n画像见 references/_shared/profiles/units/ 下对应文件，"
                   "如 references/_shared/profiles/units/<文件名>.md。\n")
        d = self.repo.add_skill()
        # evals/ 不查
        write_text(d / "evals/files/说明.md", "见 references/_shared/rules/missing.md\n")
        rep = self.run_checks("gongwen-demo")
        errors = self.closure_errors(rep)
        self.assertEqual(len(errors), 1, errors)
        self.assertIn("references/_shared/profiles/units/", errors[0])
        self.assertIn("下所需的文件", errors[0])

    def test_real_repository_reports_d18_gaps(self):
        """对真实仓库运行：各 skill 补齐 manifest 前，缺口应与 D18 清单一致（用于核对 v0.1 CM-1）。"""
        repo_root = Path(__file__).resolve().parents[2]
        if not (repo_root / "skills/gongwen-title/manifest.json").is_file():
            self.skipTest("不在完整仓库中")
        rep, _ = validate.run_validation(repo_root, "gongwen-title")
        missing = {m.split(" 引用 ")[1].split("，")[0] for m in self.closure_errors(rep)}
        # 补齐后本用例自然不再命中，只要求报出的缺口都是 references/_shared 下的路径
        self.assertTrue(all(p.startswith("references/_shared/") for p in missing), missing)


class EvalOverlapTest(ValidateTestBase):
    """W1-SO-6：--eval-overlap 检查示例是否复用评测材料。"""

    MATERIAL = ("某医院2026年工作记录：全年门急诊3216人次，满意度98.6%，投入资金1200万元。\n"
                "2026年9月26日召开推进会。依据第305条和GB/T 9704-2012，全文不超过1500字，覆盖率100%。\n"
                "统计截至2026年12月31日，参与1320人。\n")

    def make(self, reference_text: str):
        d = self.repo.add_skill()
        write_text(d / "evals/files/素材.txt", self.MATERIAL)
        write_text(d / "references/checklist.md", reference_text)
        return d

    def overlap_warnings(self, rep):
        return [m for _s, m in rep.warnings if m.startswith("示例可能复用评测材料")]

    def test_overlap_reported(self):
        self.make("# 示例\n\n- 正：全年门急诊3216 人次，满意度98.6％。\n- 推进会于2026年9月26日召开。\n")
        rep, _ = validate.run_validation(self.root, "gongwen-demo", eval_overlap=True)
        warnings = self.overlap_warnings(rep)
        self.assertEqual(len(warnings), 3, warnings)
        joined = "\n".join(warnings)
        for token in ("3216人次", "98.6%", "2026年9月26日"):
            self.assertIn(token, joined)
        self.assertIn("evals/files/素材.txt", joined)
        self.assertIn("references/checklist.md 第 3 行", joined)
        self.assertEqual(rep.errors, [])

    def test_not_run_by_default(self):
        self.make("全年门急诊3216人次。\n")
        rep = self.run_checks("gongwen-demo")
        self.assertEqual(self.overlap_warnings(rep), [])
        code, out = run_quiet(validate.main, ["--root", str(self.root), "--skill", "gongwen-demo", "--eval-overlap"])
        self.assertEqual(code, 0, out)
        self.assertIn("示例可能复用评测材料：3216人次", out)

    def test_excluded_tokens_not_reported(self):
        # 年份、条款号、标准号、字数、100%、年末日期、更长数字的一部分、_shared 中的内容都不报
        d = self.make("2026年工作要点；第305条；GB/T 9704-2012；不超过1500字；覆盖率100%；"
                      "截至2026年12月31日；参与320人；参与11320人。\n")
        write_text(d / "references/_shared/rules/demo.md", "全年门急诊3216人次。\n")
        rep, _ = validate.run_validation(self.root, "gongwen-demo", eval_overlap=True)
        self.assertEqual(self.overlap_warnings(rep), [])

    def test_prompt_and_docx_sources(self):
        d = self.make("# 示例\n\n培训覆盖456名职工，经费88.5万元。\n")
        path = d / "evals/evals.json"
        data = json.loads(path.read_text(encoding="utf-8"))
        data["evals"][0]["prompt"] = "帮我写总结，今年培训覆盖456名职工。"
        write_json(path, data)
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("word/document.xml", "<w:document><w:body><w:p><w:r><w:t>经费88.5万元</w:t></w:r></w:p>"
                                                  "</w:body></w:document>")
        (d / "evals/files/材料.docx").write_bytes(buffer.getvalue())
        rep, _ = validate.run_validation(self.root, "gongwen-demo", eval_overlap=True)
        joined = "\n".join(self.overlap_warnings(rep))
        self.assertIn("456名（评测材料：evals.json 第 1 条 prompt", joined)
        self.assertIn("88.5万元（评测材料：evals/files/材料.docx", joined)

    def test_overlap_tokens(self):
        tokens = validate.overlap_tokens("门急诊3216 人次；2026-09-26；２０２６年１０月１６日；12人；0571人；第123条；2026年")
        self.assertEqual(tokens, ["3216人次", "2026年10月16日", "2026年9月26日"])


class RepoChecksTest(ValidateTestBase):

    def test_marketplace_checks(self):
        self.repo.add_skill("gongwen-a")
        self.repo.add_skill("gongwen-b")
        self.repo.add_marketplace(skills=["gongwen-a", "gongwen-missing"], plugin_name="claude-gongwen")
        rep = self.run_checks()
        self.assert_has(rep.errors, "不能以 claude- 或 anthropic- 开头")
        self.assert_has(rep.errors, "./skills/gongwen-missing")
        self.assert_has(rep.errors, "skills/gongwen-b 未被")
        self.assert_not_has(rep.errors, "skills/gongwen-a 未被")

    def test_font_large_and_bad_json(self):
        self.repo.add_skill()
        self.repo.add_marketplace()
        (self.root / "skills/gongwen-demo/assets").mkdir(parents=True)
        (self.root / "skills/gongwen-demo/assets/仿宋.TTF").write_bytes(b"\x00")
        (self.root / "docs").mkdir()
        (self.root / "docs/big.bin").write_bytes(b"0" * (1024 * 1024 + 1))
        write_text(self.root / "shared/scripts/broken.json", "{oops}")
        # 跳过目录中的问题文件不应报告
        write_text(self.root / ".venv/lib/broken.json", "{oops}")
        (self.root / "dist/offline").mkdir(parents=True)
        (self.root / "dist/offline/f.woff2").write_bytes(b"\x00")
        rep = self.run_checks()
        self.assert_has(rep.errors, "字体文件：skills/gongwen-demo/assets/仿宋.TTF")
        self.assert_has(rep.warnings, "超过 1MB")
        self.assert_has(rep.errors, "shared/scripts/broken.json")
        self.assert_not_has(rep.errors, ".venv")
        self.assert_not_has(rep.errors, "f.woff2")
        code, _out = run_quiet(validate.main, ["--root", str(self.root)])
        self.assertEqual(code, 1)

    def test_out_dir_skipped_only_under_tests(self):
        # .gitignore 只忽略 tests/**/_out/，其他位置的 _out 目录会入库，必须照常检查
        self.repo.add_skill()
        self.repo.add_marketplace()
        (self.root / "skills/gongwen-demo/assets/_out").mkdir(parents=True)
        (self.root / "skills/gongwen-demo/assets/_out/仿宋.ttf").write_bytes(b"\x00")
        (self.root / "tests/format/_out").mkdir(parents=True)
        (self.root / "tests/format/_out/测试.ttf").write_bytes(b"\x00")
        rep = self.run_checks()
        self.assert_has(rep.errors, "字体文件：skills/gongwen-demo/assets/_out/仿宋.ttf")
        self.assert_not_has(rep.errors, "tests/format/_out")


if __name__ == "__main__":
    unittest.main()
