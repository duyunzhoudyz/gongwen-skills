# -*- coding: utf-8 -*-
"""release.yml 的静态检查（不联网、不解析 YAML，只核对关键行）。

发布流程出错只有打标签那一刻才会暴露，所以把几条不能丢的约束固定下来：触发条件、测试门槛、构建离线包、
上传方式，以及只用官方 actions。
"""
from __future__ import annotations

import re
import unittest

from tests.tools.fixtures import REPO_ROOT

WORKFLOWS = REPO_ROOT / ".github" / "workflows"
ALLOWED_ACTIONS = {"actions/checkout", "actions/setup-python"}


class ReleaseWorkflowTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls) -> None:
        cls.text = (WORKFLOWS / "release.yml").read_text(encoding="utf-8")

    def test_triggered_only_by_version_tags(self):
        self.assertRegex(self.text, r'(?m)^on:\n  push:\n    tags:\n      - "v\*"\n')
        self.assertNotIn("pull_request", self.text)
        self.assertNotIn("branches:", self.text)

    def test_only_official_actions(self):
        used = re.findall(r"(?m)^\s*(?:-\s*)?uses:\s*([^\s@]+)@", self.text)
        self.assertTrue(used)
        self.assertLessEqual(set(used), ALLOWED_ACTIONS, used)
        self.assertIn("actions/checkout", used)
        self.assertIn("actions/setup-python", used)

    def test_runs_tests_before_building(self):
        steps = ["python -m unittest discover -s tests -t .", "python tools/build_prompts.py --check",
                 "python tools/build_prompts.py --offline"]
        positions = [self.text.index(s) for s in steps]
        self.assertEqual(positions, sorted(positions))
        self.assertIn("python tools/validate.py", self.text)

    def test_uploads_with_gh_and_creates_draft_when_missing(self):
        self.assertIn("gh release view", self.text)
        self.assertRegex(self.text, r"gh release create .*--draft")
        self.assertIn("gh release upload", self.text)
        self.assertIn("dist/offline/*.zip", self.text)
        self.assertRegex(self.text, r"(?m)^permissions:\n  contents: write\n")

    def test_checks_tag_against_marketplace_version(self):
        self.assertIn("marketplace.json", self.text)
        self.assertIn("GITHUB_REF_NAME", self.text)

    def test_ci_still_has_its_own_workflow(self):
        # release.yml 是新增的，不得取代 ci.yml
        self.assertTrue((WORKFLOWS / "ci.yml").is_file())


if __name__ == "__main__":
    unittest.main()
