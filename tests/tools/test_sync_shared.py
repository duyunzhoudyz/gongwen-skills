# -*- coding: utf-8 -*-
"""tools/sync_shared.py 的单元测试。"""
from __future__ import annotations

import os
import subprocess
import unittest
from pathlib import Path

from tests.tools.fixtures import REQUIRED_SHARED, SHARED_FILES, TempRepoTestCase, run_quiet, write_json, write_text

import sync_shared


def make_junction(link: Path, target: Path) -> bool:
    """在 Windows 上创建目录联接（junction），失败时返回 False。"""
    if os.name != "nt":
        return False
    link.parent.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        ["cmd", "/c", "mklink", "/J", str(link), str(target)],
        capture_output=True,
    )
    return result.returncode == 0 and os.path.lexists(link)


class SyncSharedTest(TempRepoTestCase):

    def sync(self, *args: str):
        return run_quiet(sync_shared.main, ["--root", str(self.root), *args])

    def test_empty_repo_ok(self):
        code, out = self.sync()
        self.assertEqual(code, 0)
        self.assertIn("还没有任何 skill", out)
        code, _ = self.sync("--check")
        self.assertEqual(code, 0)

    def test_copies_files_bytewise_and_writes_readme(self):
        # 含 CRLF 和非 UTF-8 字节，验证逐字节复制
        raw = b"# \xe7\xba\xa2\xe7\xba\xbf\r\n\xff\xfe binary\r\n"
        (self.root / "shared" / "guardrails.md").write_bytes(raw)
        d = self.repo.add_skill(sync=False)
        code, out = self.sync()
        self.assertEqual(code, 0, out)
        self.assertEqual((d / "references/_shared/guardrails.md").read_bytes(), raw)
        self.assertEqual(
            (d / "references/_shared/style/ai-tone.md").read_text(encoding="utf-8"),
            SHARED_FILES["style/ai-tone.md"],
        )
        readme = (d / "references/_shared/README.md").read_text(encoding="utf-8")
        self.assertIn("本目录由 tools/sync_shared.py 自动生成，请勿手改，修改请改 shared/", readme)
        # 再次同步无改动
        code, out = self.sync()
        self.assertEqual(code, 0)
        self.assertIn("已是最新", out)

    def test_custom_dest(self):
        write_text(self.root / "shared/scripts/gongwen_lint.py", "print('lint')\n")
        shared = [{"src": s} for s in ("guardrails.md", "style/output-format.md",
                                        "style/ai-tone.md", "style/defensive.md")]
        shared.append({"src": "scripts/gongwen_lint.py", "dest": "scripts/gongwen_lint.py"})
        d = self.repo.add_skill(shared=shared)
        self.assertEqual((d / "scripts/gongwen_lint.py").read_text(encoding="utf-8"), "print('lint')\n")
        self.assertFalse((d / "references/_shared/scripts/gongwen_lint.py").exists())

    def test_check_detects_drift_without_writing(self):
        d = self.repo.add_skill()
        code, out = self.sync("--check")
        self.assertEqual(code, 0, out)

        write_text(self.root / "shared/guardrails.md", "# 新版红线\n")
        code, out = self.sync("--check")
        self.assertEqual(code, 1)
        self.assertIn("不一致", out)
        self.assertIn("references/_shared/guardrails.md", out)
        # --check 不写入
        self.assertEqual(
            (d / "references/_shared/guardrails.md").read_text(encoding="utf-8"),
            SHARED_FILES["guardrails.md"],
        )

    def test_check_detects_missing_and_stale(self):
        d = self.repo.add_skill()
        (d / "references/_shared/style/defensive.md").unlink()
        write_text(d / "references/_shared/old.md", "旧文件\n")
        code, out = self.sync("--check")
        self.assertEqual(code, 1)
        self.assertIn("缺失", out)
        self.assertIn("多余", out)
        self.assertTrue((d / "references/_shared/old.md").exists())

    def test_cleanup_only_inside_shared_dir(self):
        d = self.repo.add_skill()
        write_text(d / "references/_shared/old.md", "旧文件\n")
        write_text(d / "references/_shared/legacy/deep.md", "旧文件\n")
        write_text(d / "references/own.md", "skill 自有资料\n")
        write_text(d / "scripts/tool.py", "print(1)\n")
        code, out = self.sync()
        self.assertEqual(code, 0, out)
        self.assertFalse((d / "references/_shared/old.md").exists())
        self.assertFalse((d / "references/_shared/legacy").exists())  # 空目录被清理
        self.assertTrue((d / "references/own.md").exists())
        self.assertTrue((d / "scripts/tool.py").exists())
        self.assertTrue((d / "references/_shared/README.md").exists())
        self.assertTrue((self.root / "shared/guardrails.md").exists())

    def test_removed_entry_is_cleaned(self):
        d = self.repo.add_skill(shared=[{"src": s} for s in (
            "guardrails.md", "style/output-format.md", "style/ai-tone.md",
            "style/defensive.md", "rules/gbt9704.md")])
        self.assertTrue((d / "references/_shared/rules/gbt9704.md").exists())
        manifest = (d / "manifest.json")
        import json
        data = json.loads(manifest.read_text(encoding="utf-8"))
        data["shared"] = data["shared"][:4]
        write_json(manifest, data)
        code, out = self.sync()
        self.assertEqual(code, 0, out)
        self.assertFalse((d / "references/_shared/rules").exists())

    def test_missing_src_reports_clear_error(self):
        shared = [{"src": "guardrails.md"}, {"src": "style/nope.md"}]
        d = self.repo.add_skill(shared=shared, sync=False)
        code, out = self.sync()
        self.assertEqual(code, 2)
        self.assertIn("shared/style/nope.md 不存在", out)
        self.assertIn("gongwen-demo", out)
        # 出错的 skill 不做任何改动
        self.assertFalse((d / "references/_shared").exists())

    def test_rejects_path_traversal(self):
        shared = [{"src": "guardrails.md", "dest": "../escape.md"}]
        self.repo.add_skill(shared=shared, sync=False)
        code, out = self.sync()
        self.assertEqual(code, 2)
        self.assertIn("..", out)
        self.assertFalse((self.root / "skills/escape.md").exists())

    def test_invalid_manifest(self):
        d = self.repo.add_skill(sync=False)
        write_text(d / "manifest.json", "{不是 JSON")
        code, out = self.sync()
        self.assertEqual(code, 2)
        self.assertIn("JSON 解析失败", out)

    def test_skill_filter(self):
        a = self.repo.add_skill("gongwen-a", sync=False)
        b = self.repo.add_skill("gongwen-b", sync=False)
        code, out = self.sync("--skill", "gongwen-a")
        self.assertEqual(code, 0, out)
        self.assertTrue((a / "references/_shared/guardrails.md").exists())
        self.assertFalse((b / "references/_shared").exists())
        code, out = self.sync("--skill", "gongwen-none")
        self.assertEqual(code, 2)

    def test_protected_dest_is_case_insensitive(self):
        # Windows、macOS 文件系统不区分大小写，dest 写成 skill.md 会覆盖 SKILL.md
        for dest in ("skill.md", "Manifest.json", "readme.md", "references/_shared/readme.md"):
            with self.subTest(dest=dest):
                shared = [{"src": s} for s in REQUIRED_SHARED]
                shared.append({"src": "rules/gbt9704.md", "dest": dest})
                d = self.repo.add_skill(shared=shared, sync=False)
                originals = {n: (d / n).read_bytes() for n in ("SKILL.md", "README.md", "manifest.json")}
                code, out = self.sync()
                self.assertEqual(code, 2, out)
                self.assertIn(f"dest 不能是 {dest}", out)
                for n, data in originals.items():
                    self.assertEqual((d / n).read_bytes(), data, f"{n} 被覆盖")

    def test_junction_shared_dir_is_not_followed(self):
        # references/_shared 是指向别处的目录联接时，不得删除联接目标中的文件
        outside = self.tmp / "outside"
        write_text(outside / "victim.md", "外部文件\n")
        d = self.repo.add_skill(sync=False)
        if not make_junction(d / "references/_shared", outside):
            self.skipTest("当前系统无法创建目录联接")
        code, out = self.sync()
        self.assertEqual(code, 2, out)
        self.assertTrue((outside / "victim.md").is_file())

    def test_nested_junction_is_not_followed(self):
        # references/_shared 下的子目录是目录联接时，不得删除联接目标中的文件
        outside = self.tmp / "outside"
        write_text(outside / "victim.md", "外部文件\n")
        d = self.repo.add_skill()
        if not make_junction(d / "references/_shared/legacy", outside):
            self.skipTest("当前系统无法创建目录联接")
        code, out = self.sync()
        self.assertEqual(code, 2, out)
        self.assertTrue((outside / "victim.md").is_file())

    def test_empty_dir_cleanup_does_not_enter_junction(self):
        # 清理空目录时不得进入目录联接，删除联接目标中的空目录
        outside = self.tmp / "outside"
        (outside / "empty").mkdir(parents=True)
        d = self.repo.add_skill()
        if not make_junction(d / "references/_shared/legacy", outside):
            self.skipTest("当前系统无法创建目录联接")
        code, out = self.sync()
        self.assertEqual(code, 0, out)
        self.assertTrue((outside / "empty").is_dir())


if __name__ == "__main__":
    unittest.main()
