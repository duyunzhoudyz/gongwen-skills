# -*- coding: utf-8 -*-
"""命令行接口：JSON 结构、退出码、【起草说明】截断、编码处理。"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from tests.lint.helpers import EVAL_FILES, LINT_PATH, gl

ISSUE_KEYS = {"id", "level", "category", "line", "excerpt", "message", "suggestion", "basis", "check", "gate"}


def cli(*args: str):
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    proc = subprocess.run([sys.executable, str(LINT_PATH), *args], capture_output=True, env=env)
    return proc.returncode, proc.stdout.decode("utf-8"), proc.stderr.decode("utf-8")


class TempDirMixin:
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

    def write(self, name: str, text: str, encoding: str = "utf-8") -> Path:
        path = self.tmp / name
        path.write_bytes(text.encode(encoding))
        return path


class JsonOutputTest(TempDirMixin, unittest.TestCase):
    def test_json_structure(self):
        code, out, _err = cli(str(EVAL_FILES / "问题密集通知.txt"), "--doc-type", "通知", "--json")
        self.assertEqual(code, 0)
        data = json.loads(out)
        for key in ("file", "doc_type", "char_count", "summary", "issues"):
            self.assertIn(key, data)
        self.assertEqual(data["file"], "问题密集通知.txt")
        self.assertEqual(data["doc_type"], "通知")
        self.assertIsInstance(data["char_count"], int)
        self.assertEqual(set(data["summary"]), {"P0", "P1", "P2", "markers"})
        self.assertEqual(set(data["summary"]["markers"]), {"待补", "待核实"})
        for level in ("P0", "P1", "P2"):
            self.assertEqual(data["summary"][level], sum(1 for i in data["issues"] if i["level"] == level))
        self.assertGreater(len(data["issues"]), 0)
        for issue in data["issues"]:
            self.assertTrue(ISSUE_KEYS <= set(issue), issue)
            self.assertIn(issue["level"], ("P0", "P1", "P2"))
            self.assertIsInstance(issue["line"], int)
            self.assertTrue(issue["excerpt"])
            self.assertIn(issue["check"], gl.CHECK_ORDER)

    def test_excerpt_is_verbatim(self):
        """非隐私类问题的 excerpt 必须是原文中的连续片段（核稿意见 JSON 要逐字定位）。"""
        path = EVAL_FILES / "问题密集通知.txt"
        lines = path.read_text(encoding="utf-8").split("\n")
        _code, out, _err = cli(str(path), "--doc-type", "通知", "--json")
        for issue in json.loads(out)["issues"]:
            if issue["check"] == "privacy":
                continue
            with self.subTest(issue=issue["id"], line=issue["line"]):
                self.assertIn(issue["excerpt"], lines[issue["line"] - 1])

    def test_privacy_excerpt_masked(self):
        _code, out, _err = cli(str(EVAL_FILES / "社区救助对象申报材料.txt"), "--json")
        text = out
        self.assertNotIn("990102195803150011", text)
        self.assertNotIn("13812345678", text)

    def test_text_output(self):
        code, out, _err = cli(str(EVAL_FILES / "问题密集通知.txt"), "--doc-type", "通知")
        self.assertEqual(code, 0)
        self.assertIn("机检结果", out)
        self.assertIn("FMT-001", out)


class ExitCodeTest(TempDirMixin, unittest.TestCase):
    def test_strict_with_p0_returns_1(self):
        code, _out, _err = cli(str(EVAL_FILES / "问题密集通知.txt"), "--strict")
        self.assertEqual(code, 1)

    def test_strict_without_p0_returns_0(self):
        code, _out, _err = cli(str(EVAL_FILES / "规范通知.txt"), "--strict")
        self.assertEqual(code, 0)

    def test_missing_file_returns_2(self):
        code, _out, err = cli(str(self.tmp / "不存在.txt"))
        self.assertEqual(code, 2)
        self.assertIn("文件不存在", err)

    def test_bad_checks_returns_2(self):
        path = self.write("a.txt", "内容。")
        code, _out, err = cli(str(path), "--checks", "residue,foo")
        self.assertEqual(code, 2)
        self.assertIn("foo", err)

    def test_unsupported_suffix_returns_2(self):
        path = self.write("a.pdf", "内容。")
        code, _out, _err = cli(str(path))
        self.assertEqual(code, 2)

    def test_checks_filter(self):
        code, out, _err = cli(str(EVAL_FILES / "问题密集通知.txt"), "--checks", "punct,number", "--json")
        self.assertEqual(code, 0)
        data = json.loads(out)
        self.assertTrue(data["issues"])
        self.assertTrue({i["check"] for i in data["issues"]} <= {"punct", "number"})


class TruncationAndEncodingTest(TempDirMixin, unittest.TestCase):
    def test_stop_at_drafting_note(self):
        text = ("某市某局关于做好某项工作的通知\n\n各区某局：\n\n请按时报送。\n\n"
                "【起草说明】\n一、所作假设\n1. 以上内容仅供参考，**待定**，某发[2026]第5号。\n")
        path = self.write("正文.txt", text)
        code, out, _err = cli(str(path), "--json")
        self.assertEqual(code, 0)
        data = json.loads(out)
        self.assertEqual(data["issues"], [])
        self.assertTrue(any("起草说明" in n for n in data["notes"]))
        self.assertEqual(data["char_count"], gl.nonspace_len("某市某局关于做好某项工作的通知各区某局：请按时报送。"))

    def test_drafting_note_must_start_line(self):
        text = "正文提到【起草说明】的写法。以上内容仅供参考。\n"
        path = self.write("正文.txt", text)
        _code, out, _err = cli(str(path), "--json")
        self.assertIn("DEF-001", [i["id"] for i in json.loads(out)["issues"]])

    def test_utf8_bom_and_crlf(self):
        path = self.tmp / "bom.txt"
        path.write_bytes("﻿某发[2026]5号\r\n\r\n内容。\r\n".encode("utf-8"))
        _code, out, _err = cli(str(path), "--json")
        data = json.loads(out)
        self.assertIn("FMT-001", [i["id"] for i in data["issues"]])
        self.assertEqual([i["line"] for i in data["issues"] if i["id"] == "FMT-001"], [1])

    def test_gbk_fallback(self):
        path = self.write("gbk.txt", "某发[2026]5号\n", encoding="gbk")
        code, out, _err = cli(str(path), "--json")
        self.assertEqual(code, 0)
        data = json.loads(out)
        self.assertIn("FMT-001", [i["id"] for i in data["issues"]])
        self.assertTrue(any("GB18030" in n for n in data["notes"]))

    def test_markdown_input(self):
        path = self.write("正文.md", "## 一、总体要求\n\n**强化组织领导**。\n")
        _code, out, _err = cli(str(path), "--json")
        found = {i["id"] for i in json.loads(out)["issues"]}
        self.assertTrue({"RES-001", "RES-002"} <= found)


class BodyCharCountTest(TempDirMixin, unittest.TestCase):
    """W1-PL-4：JSON 顶层给出正文字数 body_char_count（标题之后、不含空白和版记），char_count 原义不变。"""

    def test_body_char_count(self):
        text = ("某市某局关于做好某项工作的通知\n\n各区某局：\n\n请按时报送。\n\n某市某局\n2026年10月16日\n\n"
                "抄送：市政府办公室。\n\n【起草说明】\n说明。\n")
        path = self.write("正文.txt", text)
        _code, out, _err = cli(str(path), "--json")
        data = json.loads(out)
        expected = gl.nonspace_len("各区某局：请按时报送。某市某局2026年10月16日")
        self.assertEqual(data["body_char_count"], expected)
        self.assertEqual(data["char_count"], gl.nonspace_len(
            "某市某局关于做好某项工作的通知各区某局：请按时报送。某市某局2026年10月16日抄送：市政府办公室。"))
        note = next(n for n in data["notes"] if n.startswith("篇幅口径"))
        self.assertIn(f"约 {data['body_char_count']} 字", note)


class AllowTest(TempDirMixin, unittest.TestCase):
    """FM-6-L：--allow 列出的规则命中时降为 P2、注明按本单位制度保留，不计入 --strict。"""

    def test_allow_downgrades_and_strict_passes(self):
        path = self.write("正文.txt", "某能源发［2026］47号\n\n关于做好安全生产工作的通知\n\n各子公司：\n\n请按时报送。\n")
        code, out, _err = cli(str(path), "--json", "--strict")
        self.assertEqual(code, 1)
        self.assertIn("P0", [i["level"] for i in json.loads(out)["issues"] if i["id"] == "FMT-001"])
        code, out, _err = cli(str(path), "--json", "--strict", "--allow", "FMT-001")
        self.assertEqual(code, 0)
        data = json.loads(out)
        found = [i for i in data["issues"] if i["id"] == "FMT-001"]
        self.assertTrue(found)
        self.assertTrue(all(i["level"] == "P2" and i["message"].startswith("按本单位制度保留：") for i in found))
        self.assertEqual(data["summary"]["P0"], 0)
        self.assertTrue(any("--allow" in n for n in data["notes"]))

    def test_unknown_rule_noted(self):
        path = self.write("正文.txt", "内容。\n")
        code, out, _err = cli(str(path), "--json", "--allow", "FMT-001,ABC-999")
        self.assertEqual(code, 0)
        self.assertTrue(any("ABC-999" in n for n in json.loads(out)["notes"]))

    def test_help_mentions_options(self):
        _code, out, _err = cli("--help")
        for word in ("--allow", "按本单位制度保留", "--banned", "BAN-001"):
            self.assertIn(word, out)


@unittest.skipUnless(os.name == "nt", "长路径前缀只在 Windows 下需要")
class LongPathTest(unittest.TestCase):
    """W1-IM-3：Windows 下超过 260 字符的路径能正常检查。"""

    PREFIX = "\\\\?\\"

    def test_long_path(self):
        base = Path(tempfile.mkdtemp()).resolve()
        try:
            folder = base
            while len(str(folder)) < 300:
                folder = folder / ("用于构造超长路径的目录名称" * 3)
                os.mkdir(self.PREFIX + str(folder))
            target = folder / "正文.txt"
            with open(self.PREFIX + str(target), "w", encoding="utf-8") as fh:
                fh.write("某发[2026]5号\n\n内容。\n")
            self.assertGreater(len(str(target)), 260)
            code, out, err = cli(str(target), "--json")
            self.assertEqual(code, 0, err)
            self.assertIn("FMT-001", [i["id"] for i in json.loads(out)["issues"]])
            self.assertEqual(gl.long_path(Path("相对路径.txt")), Path("相对路径.txt"))
        finally:
            import shutil
            shutil.rmtree(self.PREFIX + str(base), ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
