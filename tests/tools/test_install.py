# -*- coding: utf-8 -*-
"""tools/install.py 的单元测试（用临时 home 和 cwd，不写入真实用户目录）。"""
from __future__ import annotations

import unittest
from pathlib import Path

from tests.tools.fixtures import TempRepoTestCase, run_quiet, write_text

import install


class InstallTest(TempRepoTestCase):

    def setUp(self):
        super().setUp()
        self.home = self.tmp / "home"
        self.cwd = self.tmp / "project"
        self.home.mkdir()
        self.cwd.mkdir()

    def install(self, *args: str, stamp: str = "20261016-120000"):
        return run_quiet(
            install.main,
            ["--root", str(self.root), *args],
            home=self.home,
            cwd=self.cwd,
            stamp=stamp,
        )

    def all_files(self, base: Path):
        return sorted(p.relative_to(base).as_posix() for p in base.rglob("*"))

    def test_list_targets(self):
        code, out = self.install("--list-targets")
        self.assertEqual(code, 0)
        for key in ("claude", "codex", "agents", "cursor", "gemini", "copilot",
                    "qwen", "kimi", "codebuddy", "trae"):
            self.assertIn(key, out)
        self.assertIn("~/.kimi-code/skills", out)
        self.assertIn(".github/skills", out)

    def test_requires_target(self):
        self.repo.add_skill()
        self.assertEqual(self.install()[0], 2)
        self.assertEqual(self.install("--target", "nope")[0], 2)
        self.assertEqual(self.install("--target", "claude", "--skills", "gongwen-none")[0], 2)

    def test_empty_repo(self):
        code, out = self.install("--target", "claude")
        self.assertEqual(code, 0)
        self.assertIn("还没有可安装的 skill", out)

    def test_dry_run_changes_nothing(self):
        self.repo.add_skill()
        code, out = self.install("--target", "claude,trae", "--dry-run")
        self.assertEqual(code, 0, out)
        self.assertIn("[dry-run]", out)
        self.assertEqual(self.all_files(self.home), [])
        self.assertEqual(self.all_files(self.cwd), [])

    def test_install_user_scope(self):
        d = self.repo.add_skill()
        write_text(d / "scripts/__pycache__/x.pyc", "缓存")
        code, out = self.install("--target", "claude", "--target", "qwen")
        self.assertEqual(code, 0, out)
        for rel in (".claude/skills", ".qwen/skills"):
            dest = self.home / rel / "gongwen-demo"
            self.assertEqual(
                (dest / "SKILL.md").read_bytes(), (d / "SKILL.md").read_bytes()
            )
            self.assertTrue((dest / "references/_shared/guardrails.md").is_file())
            self.assertFalse((dest / "scripts/__pycache__").exists())

    def test_user_dirs_not_copied(self):
        # 仓库内 skill 目录下的单位画像、用户产出不得复制到安装目录
        d = self.repo.add_skill()
        write_text(d / ".gongwen/单位画像.md", "内部信息\n")
        write_text(d / "公文/20261016-测试/正文.txt", "用户草稿\n")
        code, out = self.install("--target", "claude", "--scope", "project")
        self.assertEqual(code, 0, out)
        dest = self.cwd / ".claude/skills/gongwen-demo"
        self.assertTrue((dest / "SKILL.md").is_file())
        self.assertFalse((dest / ".gongwen").exists())
        self.assertFalse((dest / "公文").exists())

    def test_project_scope(self):
        self.repo.add_skill()
        code, out = self.install("--target", "copilot", "--scope", "project")
        self.assertEqual(code, 0, out)
        self.assertTrue((self.cwd / ".github/skills/gongwen-demo/SKILL.md").is_file())
        self.assertEqual(self.all_files(self.home), [])

    def test_skills_filter(self):
        self.repo.add_skill("gongwen-a")
        self.repo.add_skill("gongwen-b")
        code, out = self.install("--target", "cursor", "--skills", "gongwen-b")
        self.assertEqual(code, 0, out)
        base = self.home / ".cursor/skills"
        self.assertEqual(sorted(p.name for p in base.iterdir()), ["gongwen-b"])

    def test_existing_is_skipped_without_force(self):
        self.repo.add_skill()
        dest = self.home / ".claude/skills/gongwen-demo"
        write_text(dest / "mine.txt", "用户自己的修改")
        code, out = self.install("--target", "claude")
        self.assertEqual(code, 0, out)
        self.assertIn("跳过", out)
        self.assertEqual(self.all_files(dest), ["mine.txt"])

    def test_force_backs_up_without_deleting(self):
        self.repo.add_skill()
        base = self.home / ".claude/skills"
        write_text(base / "gongwen-demo/mine.txt", "用户自己的修改")
        # 同一时间戳的备份已存在时，追加序号避免冲突
        write_text(base / "gongwen-demo.bak-20261016-120000/old.txt", "更早的备份")

        code, out = self.install("--target", "claude", "--force", "--dry-run")
        self.assertEqual(code, 0, out)
        self.assertIn("备份", out)
        self.assertEqual(sorted(p.name for p in base.iterdir()),
                         ["gongwen-demo", "gongwen-demo.bak-20261016-120000"])

        code, out = self.install("--target", "claude", "--force")
        self.assertEqual(code, 0, out)
        backup = base / "gongwen-demo.bak-20261016-120000-1"
        self.assertEqual((backup / "mine.txt").read_text(encoding="utf-8"), "用户自己的修改")
        self.assertTrue((base / "gongwen-demo.bak-20261016-120000/old.txt").is_file())
        self.assertTrue((base / "gongwen-demo/SKILL.md").is_file())
        self.assertFalse((base / "gongwen-demo/mine.txt").exists())

    def test_unverified_reminder(self):
        self.repo.add_skill()
        note = install.UNVERIFIED_NOTE
        code, out = self.install("--target", "trae", "--dry-run")
        self.assertIn(note, out)
        code, out = self.install("--target", "codebuddy", "--dry-run")
        self.assertIn(note, out)
        code, out = self.install("--target", "codebuddy", "--scope", "project", "--dry-run")
        self.assertNotIn(note, out)
        code, out = self.install("--target", "claude", "--dry-run")
        self.assertNotIn(note, out)

    def test_shared_directory_installed_once(self):
        self.repo.add_skill()
        code, out = self.install("--target", "codex,agents")
        self.assertEqual(code, 0, out)
        self.assertIn("目录相同", out)
        self.assertNotIn("跳过  gongwen-demo", out)
        self.assertTrue((self.home / ".agents/skills/gongwen-demo/SKILL.md").is_file())


if __name__ == "__main__":
    unittest.main()
