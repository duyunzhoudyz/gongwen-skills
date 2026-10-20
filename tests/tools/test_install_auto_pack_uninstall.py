# -*- coding: utf-8 -*-
"""tools/install.py 新功能的单元测试：--detect、--target auto、--pack、--uninstall、安装后提示。

全部使用临时目录构造的假 HOME、假工作目录和假 which，不读写真实用户目录，也不查找真实命令。
"""
from __future__ import annotations

import os
import stat
import unittest
from pathlib import Path
from typing import Dict, List, Optional
from unittest import mock

from tests.tools.fixtures import TempRepoTestCase, run_quiet, skill_md, write_json, write_text

import install

CORE = ["gongwen", "gongwen-report", "gongwen-profile"]
PACKS = {
    "grassroots": ["gongwen-grassroots"],
    "hospital": ["gongwen-hospital"],
    "university": ["gongwen-university"],
    "k12": ["gongwen-k12"],
    "soe": ["gongwen-soe"],
}


def fake_which(found: Optional[Dict[str, str]] = None):
    """假的 shutil.which：只认 found 中列出的命令，并记录被查询过的命令名。"""
    found = found or {}
    calls: List[str] = []

    def which(name: str) -> Optional[str]:
        calls.append(name)
        return found.get(name)

    which.calls = calls  # type: ignore[attr-defined]
    return which


class InstallTestBase(TempRepoTestCase):

    def setUp(self):
        super().setUp()
        self.home = self.tmp / "home"
        self.cwd = self.tmp / "project"
        self.home.mkdir()
        self.cwd.mkdir()
        self.which = fake_which()

    def make_marketplace(self, core=None, packs=None) -> Path:
        core = CORE if core is None else core
        packs = PACKS if packs is None else packs
        plugins = [{"name": "gongwen", "source": "./", "strict": False,
                    "skills": [f"./skills/{s}" for s in core]}]
        for key, skills in packs.items():
            plugins.append({"name": f"gongwen-{key}", "source": "./", "strict": False,
                            "skills": [f"./skills/{s}" for s in skills]})
        data = {"name": "gongwen-skills", "owner": {"name": "test"},
                "metadata": {"version": "1.0.0"}, "plugins": plugins}
        return write_json(self.root / ".claude-plugin" / "marketplace.json", data)

    def add_skills(self, *names: str) -> None:
        for name in names:
            self.repo.add_skill(name, sync=False)

    def run_install(self, *args: str, stamp: str = "20261016-120000"):
        return run_quiet(
            install.main,
            ["--root", str(self.root), *args],
            home=self.home,
            cwd=self.cwd,
            stamp=stamp,
            which=self.which,
        )

    def installed(self, rel: str) -> List[str]:
        base = self.home / rel
        if not base.is_dir():
            return []
        return sorted(p.name for p in base.iterdir())

    def snapshot(self, base: Path) -> List[str]:
        return sorted(p.relative_to(base).as_posix() for p in base.rglob("*"))


class DetectTest(InstallTestBase):

    def test_detect_by_user_dirs_and_commands(self):
        (self.home / ".claude").mkdir()
        (self.home / ".qwen").mkdir()
        self.which = fake_which({"cursor": "/fake/bin/cursor"})
        before = self.snapshot(self.home)
        code, out = self.run_install("--detect")
        self.assertEqual(code, 0, out)
        self.assertIn("[已发现] claude", out)
        self.assertIn("[已发现] qwen", out)
        self.assertIn("[已发现] cursor", out)
        self.assertIn("cursor 命令", out)
        self.assertIn("[未发现] gemini", out)
        self.assertIn("共发现 3 个", out)
        # 探测不得写入任何文件
        self.assertEqual(self.snapshot(self.home), before)

    def test_detect_nothing(self):
        code, out = self.run_install("--detect")
        self.assertEqual(code, 0, out)
        self.assertIn("未探测到已安装的 Agent 工具", out)
        self.assertNotIn("[已发现]", out)

    def test_detect_ignores_plain_file(self):
        # 同名普通文件不算配置目录
        write_text(self.home / ".gemini", "不是目录")
        code, out = self.run_install("--detect")
        self.assertIn("[未发现] gemini", out)

    def test_detect_never_executes_tools(self):
        (self.home / ".claude").mkdir()
        self.which = fake_which({"claude": "/fake/bin/claude"})
        with mock.patch("subprocess.Popen", side_effect=AssertionError("不应运行任何命令")), \
                mock.patch("os.system", side_effect=AssertionError("不应运行任何命令")):
            code, out = self.run_install("--detect")
        self.assertEqual(code, 0, out)
        # 只查询命令是否存在，查询的都是预设的命令名
        expected = {c for _dirs, cmds in install.DETECT_HINTS.values() for c in cmds}
        self.assertTrue(set(self.which.calls) <= expected)

    def test_detect_which_failure_is_tolerated(self):
        def broken(_name):
            raise OSError("PATH 读取失败")
        self.which = broken
        (self.home / ".claude").mkdir()
        code, out = self.run_install("--detect")
        self.assertEqual(code, 0, out)
        self.assertIn("[已发现] claude", out)

    def test_every_target_has_detect_hints(self):
        self.assertEqual(set(install.DETECT_HINTS), set(install.TARGET_MAP))


class AutoTargetTest(InstallTestBase):

    def setUp(self):
        super().setUp()
        self.add_skills("gongwen", "gongwen-report")
        self.make_marketplace(core=["gongwen", "gongwen-report"], packs={})

    def test_auto_installs_to_detected_only(self):
        (self.home / ".claude").mkdir()
        (self.home / ".gemini").mkdir()
        code, out = self.run_install("--target", "auto")
        self.assertEqual(code, 0, out)
        self.assertIn("自动探测到 2 个工具", out)
        self.assertEqual(self.installed(".claude/skills"), ["gongwen", "gongwen-report"])
        self.assertEqual(self.installed(".gemini/skills"), ["gongwen", "gongwen-report"])
        self.assertFalse((self.home / ".cursor").exists())
        self.assertFalse((self.home / ".qwen").exists())

    def test_auto_detects_by_command(self):
        self.which = fake_which({"qwen": "C:\\fake\\qwen.cmd"})
        code, out = self.run_install("--target", "auto")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".qwen/skills"), ["gongwen", "gongwen-report"])

    def test_auto_nothing_detected(self):
        code, out = self.run_install("--target", "auto")
        self.assertEqual(code, 3, out)
        self.assertIn("未探测到", out)
        self.assertEqual(self.snapshot(self.home), [])

    def test_auto_combined_with_explicit(self):
        (self.home / ".claude").mkdir()
        code, out = self.run_install("--target", "qwen,auto")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".claude/skills"), ["gongwen", "gongwen-report"])
        self.assertEqual(self.installed(".qwen/skills"), ["gongwen", "gongwen-report"])

    def test_auto_explicit_target_still_works_when_nothing_detected(self):
        code, out = self.run_install("--target", "auto", "--target", "qwen")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".qwen/skills"), ["gongwen", "gongwen-report"])

    def test_auto_shared_dir_installed_once(self):
        (self.home / ".codex").mkdir()
        (self.home / ".agents").mkdir()
        code, out = self.run_install("--target", "auto")
        self.assertEqual(code, 0, out)
        self.assertIn("目录相同", out)
        self.assertEqual(self.installed(".agents/skills"), ["gongwen", "gongwen-report"])

    def test_auto_dry_run_changes_nothing(self):
        (self.home / ".claude").mkdir()
        before = self.snapshot(self.home)
        code, out = self.run_install("--target", "auto", "--dry-run")
        self.assertEqual(code, 0, out)
        self.assertIn("[dry-run]", out)
        self.assertEqual(self.snapshot(self.home), before)
        self.assertEqual(self.snapshot(self.cwd), [])

    def test_unknown_target_lists_auto(self):
        code, out = self.run_install("--target", "nope")
        self.assertEqual(code, 2)
        self.assertIn("auto", out)


class PackTest(InstallTestBase):

    def setUp(self):
        super().setUp()
        self.add_skills("gongwen", "gongwen-report", "gongwen-hospital", "gongwen-soe", "gongwen-university")
        self.make_marketplace()

    def test_default_installs_core_only(self):
        code, out = self.run_install("--target", "claude")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".claude/skills"), ["gongwen", "gongwen-report"])
        # 清单中列出但仓库还没有的 skill 给出提醒，不算失败
        self.assertIn("gongwen-profile", out)
        self.assertIn("已跳过", out)

    def test_single_pack(self):
        code, out = self.run_install("--target", "claude", "--pack", "hospital")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".claude/skills"),
                         ["gongwen", "gongwen-hospital", "gongwen-report"])

    def test_multiple_packs_repeat_and_comma(self):
        code, out = self.run_install("--target", "claude", "--pack", "hospital", "--pack", "soe")
        self.assertEqual(code, 0, out)
        expected = ["gongwen", "gongwen-hospital", "gongwen-report", "gongwen-soe"]
        self.assertEqual(self.installed(".claude/skills"), expected)
        code, out = self.run_install("--target", "qwen", "--pack", "hospital,soe")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".qwen/skills"), expected)

    def test_pack_all(self):
        code, out = self.run_install("--target", "claude", "--pack", "all")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".claude/skills"), [
            "gongwen", "gongwen-hospital", "gongwen-report", "gongwen-soe", "gongwen-university",
        ])
        self.assertIn("gongwen-k12", out)  # 缺失的行业包给出提醒

    def test_unknown_pack(self):
        code, out = self.run_install("--target", "claude", "--pack", "bank")
        self.assertEqual(code, 2, out)
        self.assertIn("不支持的行业包", out)
        self.assertEqual(self.snapshot(self.home), [])

    def test_skills_with_pack_takes_union_without_core(self):
        code, out = self.run_install("--target", "claude", "--skills", "gongwen-report", "--pack", "soe")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".claude/skills"), ["gongwen-report", "gongwen-soe"])

    def test_skill_list_comes_from_marketplace(self):
        # 改动清单后安装范围随之改变，说明没有写死
        self.make_marketplace(core=["gongwen-report"], packs={"hospital": ["gongwen-university"]})
        code, out = self.run_install("--target", "claude", "--pack", "hospital")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".claude/skills"), ["gongwen-report", "gongwen-university"])

    def test_marketplace_with_bom_and_crlf(self):
        path = self.root / ".claude-plugin" / "marketplace.json"
        text = path.read_bytes().decode("utf-8").replace("\n", "\r\n")
        path.write_bytes(b"\xef\xbb\xbf" + text.encode("utf-8"))
        code, out = self.run_install("--target", "claude", "--pack", "soe")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".claude/skills"), ["gongwen", "gongwen-report", "gongwen-soe"])

    def test_invalid_marketplace(self):
        write_text(self.root / ".claude-plugin" / "marketplace.json", "{ 不是 JSON")
        code, out = self.run_install("--target", "claude")
        self.assertEqual(code, 2, out)
        self.assertIn("无法读取", out)

    def test_marketplace_rejects_path_traversal(self):
        self.make_marketplace(core=["gongwen"], packs={})
        path = self.root / ".claude-plugin" / "marketplace.json"
        write_text(path, path.read_text(encoding="utf-8").replace("./skills/gongwen", "./skills/../../evil"))
        code, out = self.run_install("--target", "claude")
        self.assertEqual(code, 2, out)
        self.assertIn("不合规", out)

    def test_pack_requires_marketplace(self):
        (self.root / ".claude-plugin" / "marketplace.json").unlink()
        code, out = self.run_install("--target", "claude", "--pack", "hospital")
        self.assertEqual(code, 2, out)
        # 没有 marketplace.json 且不选行业包时，兼容旧行为：安装 skills/ 下全部 skill
        code, out = self.run_install("--target", "claude")
        self.assertEqual(code, 0, out)
        self.assertEqual(len(self.installed(".claude/skills")), 5)


class UninstallTest(InstallTestBase):

    def setUp(self):
        super().setUp()
        self.add_skills("gongwen", "gongwen-report", "gongwen-hospital")
        self.make_marketplace()
        self.base = self.home / ".claude/skills"
        code, out = self.run_install("--target", "claude", "--pack", "hospital")
        self.assertEqual(code, 0, out)

    def test_lists_without_deleting_by_default(self):
        before = self.snapshot(self.home)
        code, out = self.run_install("--target", "claude", "--uninstall")
        self.assertEqual(code, 0, out)
        self.assertIn("将删除  gongwen", out)
        self.assertIn("--yes", out)
        self.assertEqual(self.snapshot(self.home), before)

    def test_dry_run_with_yes_deletes_nothing(self):
        before = self.snapshot(self.home)
        code, out = self.run_install("--target", "claude", "--uninstall", "--yes", "--dry-run")
        self.assertEqual(code, 0, out)
        self.assertIn("[dry-run]", out)
        self.assertEqual(self.snapshot(self.home), before)

    def test_yes_deletes_only_matching_skill_dirs(self):
        # 用户自己的 skill、skills 根目录下的文件、备份目录、name 不一致的同名目录都必须保留
        write_text(self.base / "my-skill/SKILL.md", skill_md("my-skill"))
        write_text(self.base / "说明.txt", "用户自己的文件\n")
        write_text(self.base / "gongwen.bak-20260101-000000/SKILL.md", skill_md("gongwen"))
        write_text(self.base / "gongwen-report/SKILL.md", skill_md("someone-else"))
        code, out = self.run_install("--target", "claude", "--uninstall", "--yes")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".claude/skills"), [
            "gongwen-report", "gongwen.bak-20260101-000000", "my-skill", "说明.txt",
        ])
        self.assertIn("与目录名不一致", out)
        self.assertIn("备份目录", out)
        self.assertTrue((self.base / "my-skill/SKILL.md").is_file())

    def test_force_skips_confirmation(self):
        code, out = self.run_install("--target", "claude", "--uninstall", "--force")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".claude/skills"), [])
        self.assertIn("卸载完成", out)

    def test_pack_only(self):
        code, out = self.run_install("--target", "claude", "--uninstall", "--pack", "hospital", "--yes")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".claude/skills"), ["gongwen", "gongwen-report"])

    def test_skills_filter(self):
        code, out = self.run_install("--target", "claude", "--uninstall", "--skills", "gongwen-report", "--yes")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".claude/skills"), ["gongwen", "gongwen-hospital"])

    def test_skill_not_in_marketplace_is_refused(self):
        write_text(self.base / "my-skill/SKILL.md", skill_md("my-skill"))
        code, out = self.run_install("--target", "claude", "--uninstall", "--skills", "my-skill", "--yes")
        self.assertEqual(code, 2, out)
        self.assertTrue((self.base / "my-skill/SKILL.md").is_file())

    def test_requires_marketplace(self):
        (self.root / ".claude-plugin" / "marketplace.json").unlink()
        code, out = self.run_install("--target", "claude", "--uninstall", "--yes")
        self.assertEqual(code, 2, out)
        self.assertEqual(self.installed(".claude/skills"), ["gongwen", "gongwen-hospital", "gongwen-report"])

    def test_missing_skill_md_or_name_is_kept(self):
        (self.base / "gongwen/SKILL.md").unlink()
        write_text(self.base / "gongwen-report/SKILL.md", "# 没有 frontmatter\n")
        code, out = self.run_install("--target", "claude", "--uninstall", "--yes")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".claude/skills"), ["gongwen", "gongwen-report"])
        self.assertIn("缺少 SKILL.md", out)

    def test_user_data_inside_skill_dir_is_kept(self):
        write_text(self.base / "gongwen/.gongwen/单位画像.md", "内部信息\n")
        code, out = self.run_install("--target", "claude", "--uninstall", "--yes")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".claude/skills"), ["gongwen"])
        self.assertTrue((self.base / "gongwen/.gongwen/单位画像.md").is_file())
        self.assertIn("用户资料", out)

    def test_symlinked_skill_dir_is_kept(self):
        outside = self.tmp / "outside" / "gongwen-hospital"
        write_text(outside / "SKILL.md", skill_md("gongwen-hospital"))
        link = self.base / "gongwen-hospital"
        install.remove_tree(link)
        try:
            os.symlink(str(outside), str(link), target_is_directory=True)
        except (OSError, NotImplementedError):
            self.skipTest("当前系统不允许创建符号链接")
        code, out = self.run_install("--target", "claude", "--uninstall", "--yes")
        self.assertEqual(code, 0, out)
        self.assertIn("符号链接", out)
        self.assertTrue(link.is_symlink())
        self.assertTrue((outside / "SKILL.md").is_file())

    @unittest.skipUnless(os.name == "nt", "目录联接（junction）只在 Windows 上存在")
    def test_junction_skill_dir_is_kept(self):
        import _winapi  # CPython 在 Windows 上自带
        outside = self.tmp / "outside" / "gongwen-hospital"
        write_text(outside / "SKILL.md", skill_md("gongwen-hospital"))
        link = self.base / "gongwen-hospital"
        install.remove_tree(link)
        _winapi.CreateJunction(str(outside), str(link))
        code, out = self.run_install("--target", "claude", "--uninstall", "--yes")
        self.assertEqual(code, 0, out)
        self.assertIn("目录联接", out)
        self.assertTrue((outside / "SKILL.md").is_file())
        self.assertTrue(link.exists())
        os.rmdir(str(link))  # 只移除联接本身

    def test_read_only_files_are_removed(self):
        ro = self.base / "gongwen/SKILL.md"
        os.chmod(ro, stat.S_IREAD)
        code, out = self.run_install("--target", "claude", "--uninstall", "--yes")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".claude/skills"), [])

    def test_auto_uninstall(self):
        (self.home / ".claude").mkdir(exist_ok=True)
        code, out = self.run_install("--target", "auto", "--uninstall", "--yes")
        self.assertEqual(code, 0, out)
        self.assertEqual(self.installed(".claude/skills"), [])

    def test_nothing_installed(self):
        code, out = self.run_install("--target", "qwen", "--uninstall", "--yes")
        self.assertEqual(code, 0, out)
        self.assertIn("没有需要删除的目录", out)

    def test_project_scope(self):
        code, out = self.run_install("--target", "cursor", "--scope", "project")
        self.assertEqual(code, 0, out)
        code, out = self.run_install("--target", "cursor", "--scope", "project", "--uninstall", "--yes")
        self.assertEqual(code, 0, out)
        self.assertEqual(sorted(p.name for p in (self.cwd / ".cursor/skills").iterdir()), [])
        # 用户级安装不受影响
        self.assertEqual(self.installed(".claude/skills"), ["gongwen", "gongwen-hospital", "gongwen-report"])


class ReadSkillNameTest(TempRepoTestCase):

    def read(self, raw: bytes) -> Optional[str]:
        path = self.tmp / "SKILL.md"
        path.write_bytes(raw)
        return install.read_skill_name(path)

    def test_plain(self):
        self.assertEqual(self.read(skill_md("gongwen").encode("utf-8")), "gongwen")

    def test_crlf_and_bom(self):
        text = skill_md("gongwen-report").replace("\n", "\r\n")
        self.assertEqual(self.read(b"\xef\xbb\xbf" + text.encode("utf-8")), "gongwen-report")

    def test_quoted(self):
        self.assertEqual(self.read('---\nname: "gongwen"\n---\n'.encode("utf-8")), "gongwen")
        self.assertEqual(self.read("---\nname: 'gongwen'\n---\n".encode("utf-8")), "gongwen")

    def test_missing(self):
        self.assertIsNone(self.read("# 无 frontmatter\nname: gongwen\n".encode("utf-8")))
        self.assertIsNone(self.read("---\ndescription: x\n---\nname: gongwen\n".encode("utf-8")))
        self.assertIsNone(self.read(b"\xff\xfe\x00"))


class NextStepsTest(InstallTestBase):

    def setUp(self):
        super().setUp()
        self.add_skills("gongwen", "gongwen-profile", "gongwen-hospital")
        self.make_marketplace(core=["gongwen", "gongwen-profile"], packs={"hospital": ["gongwen-hospital"]})

    def test_usage_hint(self):
        code, out = self.run_install("--target", "claude")
        self.assertEqual(code, 0, out)
        self.assertIn("下一步怎么用", out)
        self.assertIn("用公文 skill 帮我写个通知", out)
        self.assertIn("单位画像", out)
        self.assertIn("--pack", out)  # 未选行业包时提示可选装

    def test_no_pack_hint_when_pack_chosen(self):
        code, out = self.run_install("--target", "qwen", "--pack", "hospital")
        self.assertEqual(code, 0, out)
        self.assertNotIn("行业包按需选装", out)
        self.assertIn("Qwen Code", out)

    def test_missing_format_deps(self):
        with mock.patch.object(install, "module_available", return_value=False):
            code, out = self.run_install("--target", "claude", "--dry-run")
        self.assertEqual(code, 0, out)
        self.assertIn("未安装", out)
        self.assertIn("pip install python-docx openpyxl", out)

    def test_format_deps_present(self):
        with mock.patch.object(install, "module_available", return_value=True):
            code, out = self.run_install("--target", "claude", "--dry-run")
        self.assertEqual(code, 0, out)
        self.assertIn("已安装", out)
        self.assertNotIn("pip install python-docx openpyxl", out)

    def test_module_available_never_raises(self):
        self.assertTrue(install.module_available("json"))
        self.assertFalse(install.module_available("gongwen_no_such_module_xyz"))
        self.assertFalse(install.module_available(""))


class FontCheckTest(InstallTestBase):
    """安装排版 skill 后，调用其 gw_fonts.py 检查标准字体并打印缺失字体与引导；检查失败不影响安装结果。"""

    STUB = "import sys\nprint('假的字体检查报告：缺方正小标宋简体')\nsys.exit(3)\n"

    def setUp(self):
        super().setUp()
        self.add_skills("gongwen", "gongwen-format")
        self.make_marketplace(core=["gongwen", "gongwen-format"], packs={})
        self.script = write_text(self.root / install.FONT_CHECK_REL, self.STUB)

    def test_font_report_printed_after_install(self):
        code, out = self.run_install("--target", "claude")
        self.assertEqual(code, 0, out)  # 字体检查返回 3 也不影响安装结果
        self.assertIn("排版字体检查", out)
        self.assertIn("假的字体检查报告：缺方正小标宋简体", out)
        self.assertIn("不会改用其他字体顶替", out)
        self.assertEqual(self.installed(".claude/skills"), ["gongwen", "gongwen-format"])

    def test_skipped_in_dry_run(self):
        code, out = self.run_install("--target", "claude", "--dry-run")
        self.assertEqual(code, 0, out)
        self.assertNotIn("排版字体检查", out)

    def test_skipped_when_format_not_installed(self):
        code, out = self.run_install("--target", "claude", "--skills", "gongwen")
        self.assertEqual(code, 0, out)
        self.assertNotIn("排版字体检查", out)

    def test_skipped_when_script_missing(self):
        self.script.unlink()
        code, out = self.run_install("--target", "claude")
        self.assertEqual(code, 0, out)
        self.assertNotIn("排版字体检查", out)

    def test_failure_to_run_does_not_fail_install(self):
        with mock.patch.object(install.subprocess, "run", side_effect=OSError("找不到 Python")):
            code, out = self.run_install("--target", "claude")
        self.assertEqual(code, 0, out)
        self.assertIn("字体检查没能运行", out)
        self.assertIn("python scripts/gw_fonts.py", out)
        self.assertEqual(self.installed(".claude/skills"), ["gongwen", "gongwen-format"])

    def test_empty_output_does_not_fail_install(self):
        write_text(self.script, "pass\n")
        code, out = self.run_install("--target", "claude")
        self.assertEqual(code, 0, out)
        self.assertIn("字体检查没有输出", out)


if __name__ == "__main__":
    unittest.main()
