# -*- coding: utf-8 -*-
"""tools/build_prompts.py 的单元测试。"""
from __future__ import annotations

import shutil
import unittest
import zipfile

from tests.tools.fixtures import TempRepoTestCase, run_quiet, skill_md, write_text

import build_prompts


class BuildPromptsTest(TempRepoTestCase):

    def build(self, *args: str, root=None):
        return run_quiet(build_prompts.main, ["--root", str(root or self.root), *args])

    def read_prompt(self, name="gongwen-demo", root=None) -> str:
        return ((root or self.root) / "dist/prompts" / f"{name}.md").read_text(encoding="utf-8")

    def test_empty_repo(self):
        code, out = self.build()
        self.assertEqual(code, 0, out)
        self.assertIn("还没有任何 skill", out)
        self.assertFalse((self.root / "dist").exists())
        code, out = self.build("--check")
        self.assertEqual(code, 0, out)

    def test_content_structure(self):
        self.repo.add_skill()
        code, out = self.build()
        self.assertEqual(code, 0, out)
        text = self.read_prompt()
        # 信息块
        self.assertTrue(text.startswith("> **演示技能（gongwen-demo）**（提示词版）"))
        self.assertIn("版本：0.1.0", text)
        self.assertIn("成熟度：Beta", text)
        for item in ("1. 复制全文，粘贴到对话开头；", "2. 再说出你的需求；",
                     "3. 网页版无法运行脚本，自检清单需要人工逐项执行；", "4. 请勿输入涉密和内部敏感信息。"):
            self.assertIn(item, text)
        self.assertIn(f"全文约 {len(text)} 字符", text)  # 字数包含信息块本身
        # 正文去掉 frontmatter
        self.assertNotIn("license: MIT", text)
        self.assertIn("# 演示技能（gongwen-demo）", text)
        # 附录顺序与标题
        i1 = text.index("# 附录一　六条写作红线")
        i2 = text.index("# 附录二　输出格式规范")
        i3 = text.index("# 附录三　自检清单")
        self.assertLess(i1, i2)
        self.assertLess(i2, i3)
        # 路径替换
        body = text[:i1]
        self.assertIn("保密：按（见附录一）第一条执行。", body)
        self.assertIn("输出前读（见附录二）。", body)
        self.assertIn("| （见附录三） | 自检时 |", body)
        self.assertNotIn("references/_shared/guardrails.md", body)
        self.assertNotIn("references/checklist.md", body)
        # 未打包的路径保持原样
        self.assertIn("`evals/evals.json`", body)
        # 文件原标题不重复出现
        self.assertEqual(text.count("六条写作红线"), 1)
        self.assertNotIn("\r", text)

    def test_index_readme(self):
        self.repo.add_skill("gongwen-a")
        self.repo.add_skill("gongwen")
        code, out = self.build()
        self.assertEqual(code, 0, out)
        index = (self.root / "dist/prompts/README.md").read_text(encoding="utf-8")
        self.assertIn("| 名称 | 中文名 | 说明 | 字数 | 文件 |", index)
        size = len(self.read_prompt("gongwen-a"))
        self.assertIn(
            f"| gongwen-a | 演示技能 | 演示用的公文 skill，用于测试工具链。 | {size} | [gongwen-a.md](gongwen-a.md) |",
            index,
        )
        self.assertLess(index.index("| gongwen |"), index.index("| gongwen-a |"))

    def test_deterministic(self):
        self.repo.add_skill("gongwen-a")
        self.repo.add_skill("gongwen-b", status="draft")
        self.assertEqual(self.build()[0], 0)
        first = {p.name: p.read_bytes() for p in (self.root / "dist/prompts").iterdir()}
        # 同一仓库重复构建
        self.assertEqual(self.build()[0], 0)
        second = {p.name: p.read_bytes() for p in (self.root / "dist/prompts").iterdir()}
        self.assertEqual(first, second)
        # 另一份拷贝独立构建
        other = self.tmp / "copy"
        shutil.copytree(self.root, other, ignore=shutil.ignore_patterns("dist"))
        self.assertEqual(self.build(root=other)[0], 0)
        third = {p.name: p.read_bytes() for p in (other / "dist/prompts").iterdir()}
        self.assertEqual(first, third)

    def test_check_mode(self):
        d = self.repo.add_skill()
        code, out = self.build("--check")
        self.assertEqual(code, 1)
        self.assertIn("缺失", out)
        self.assertFalse((self.root / "dist").exists())

        self.assertEqual(self.build()[0], 0)
        code, out = self.build("--check")
        self.assertEqual(code, 0, out)

        # CRLF 换行不算差异（兼容 Windows 下 git 的 autocrlf）
        prompt = self.root / "dist/prompts/gongwen-demo.md"
        prompt.write_bytes(prompt.read_bytes().replace(b"\n", b"\r\n"))
        self.assertEqual(self.build("--check")[0], 0)

        write_text(d / "SKILL.md", skill_md("gongwen-demo", extra_body="\n新增一行。\n"))
        code, out = self.build("--check")
        self.assertEqual(code, 1)
        self.assertIn("不一致  dist/prompts/gongwen-demo.md", out)

        self.assertEqual(self.build()[0], 0)
        write_text(self.root / "dist/prompts/gongwen-old.md", "旧文件\n")
        code, out = self.build("--check")
        self.assertEqual(code, 1)
        self.assertIn("多余", out)
        self.assertEqual(self.build()[0], 0)
        self.assertFalse((self.root / "dist/prompts/gongwen-old.md").exists())

    def test_long_prompt_warning(self):
        d = self.repo.add_skill()
        write_text(d / "references/checklist.md", "# 自检清单\n\n" + "长" * 31000 + "\n")
        code, out = self.build()
        self.assertEqual(code, 0, out)
        self.assertIn("警告", out)
        self.assertIn("超过 30000", out)
        text = self.read_prompt()
        self.assertIn("注意：全文超过 30000 字符", text)
        self.assertIn(f"全文约 {len(text)} 字符", text)

    def test_missing_bundle_file_is_error(self):
        self.repo.add_skill(bundle=["references/nope.md"])
        code, out = self.build()
        self.assertEqual(code, 1)
        self.assertIn("references/nope.md", out)
        self.assertFalse((self.root / "dist/prompts/gongwen-demo.md").exists())

    def test_non_markdown_bundle(self):
        self.repo.add_skill(bundle=["evals/trigger.json"])
        self.assertEqual(self.build()[0], 0)
        text = self.read_prompt()
        self.assertIn("# 附录一　trigger.json", text)
        self.assertIn("```json\n[", text)

    def test_offline_zip(self):
        self.repo.add_skill()
        self.repo.add_marketplace(version="1.2.3")
        write_text(self.root / "README.md", "# 主页\n")
        write_text(self.root / "skills/gongwen-demo/scripts/__pycache__/x.pyc", "缓存")
        code, out = self.build("--offline")
        self.assertEqual(code, 0, out)
        target = self.root / "dist/offline/gongwen-skills-offline-1.2.3.zip"
        self.assertTrue(target.is_file())
        with zipfile.ZipFile(target) as zf:
            names = zf.namelist()
        self.assertIn("README.md", names)
        self.assertNotIn("LICENSE", names)  # 不存在的不打包
        self.assertIn("skills/gongwen-demo/SKILL.md", names)
        self.assertIn("dist/prompts/gongwen-demo.md", names)
        self.assertIn("dist/prompts/README.md", names)
        self.assertFalse(any("__pycache__" in n for n in names))
        first = target.read_bytes()
        self.assertEqual(self.build("--offline")[0], 0)
        self.assertEqual(target.read_bytes(), first)

    def test_offline_zip_excludes_user_dirs(self):
        # 单位画像、用户产出等目录可能含内部信息，不得打进离线包
        d = self.repo.add_skill()
        write_text(d / ".gongwen/单位画像.md", "内部信息\n")
        write_text(d / "公文/20261016-测试/正文.txt", "用户草稿\n")
        write_text(d / ".venv/lib/x.py", "print(1)\n")
        write_text(d / "references/公文样例.md", "# 样例\n")  # 同名前缀的普通文件照常打包
        code, out = self.build("--offline")
        self.assertEqual(code, 0, out)
        target = self.root / "dist/offline/gongwen-skills-offline-0.0.0.zip"
        with zipfile.ZipFile(target) as zf:
            names = zf.namelist()
        self.assertFalse(any(".gongwen" in n for n in names), names)
        self.assertFalse(any("/公文/" in n for n in names), names)
        self.assertFalse(any(".venv" in n for n in names), names)
        self.assertIn("skills/gongwen-demo/references/公文样例.md", names)

    def test_bundle_file_with_bom(self):
        # 用记事本等工具保存的 UTF-8 BOM 文件：BOM 不得混入提示词，标题照常识别
        d = self.repo.add_skill()
        (d / "references/checklist.md").write_bytes("﻿# 自检清单\n\n- [ ] 标题规范\n".encode("utf-8"))
        code, out = self.build()
        self.assertEqual(code, 0, out)
        text = self.read_prompt()
        self.assertNotIn("﻿", text)
        self.assertIn("# 附录三　自检清单\n\n- [ ] 标题规范", text)

    def test_offline_default_version(self):
        self.repo.add_skill()
        code, out = self.build("--offline")
        self.assertEqual(code, 0, out)
        self.assertTrue((self.root / "dist/offline/gongwen-skills-offline-0.0.0.zip").is_file())
        self.assertIn("0.0.0", out)

    def test_helpers(self):
        self.assertEqual(build_prompts.cn_number(1), "一")
        self.assertEqual(build_prompts.cn_number(10), "十")
        self.assertEqual(build_prompts.cn_number(12), "十二")
        self.assertEqual(build_prompts.cn_number(20), "二十")
        self.assertEqual(build_prompts.cn_number(35), "三十五")
        mapping = {"references/a.md": "（见附录一）"}
        replace = build_prompts.replace_bundle_paths
        self.assertEqual(replace("按 `references/a.md` 执行", mapping), "按（见附录一）执行")
        self.assertEqual(replace("see references/a.md.", mapping), "see （见附录一）.")
        self.assertEqual(replace("references/a.md.bak", mapping), "references/a.md.bak")
        self.assertEqual(replace("shared/references/a.md", mapping), "shared/references/a.md")
        self.assertEqual(build_prompts.short_title("综合文稿（gongwen-report）", "gongwen-report"), "综合文稿")
        self.assertEqual(build_prompts.first_sentence("第一句。第二句。"), "第一句。")


if __name__ == "__main__":
    unittest.main()
