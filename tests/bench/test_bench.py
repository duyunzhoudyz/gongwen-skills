# -*- coding: utf-8 -*-
"""评测框架 evals/bench 的最小单测（audit G17）。

覆盖：
    aggregate.py   用 evals/results/v0.1-raw.json 复算出 README 里的 v0.1 数字；另用手工小标本核对统计口径；
    trigger_score  用临时小标本核对正例、反例、缺答案、混淆的统计；
    prep.py        --groups A 只重建 A 组，B 组的工作目录、输出和哈希原样保留（评测目录放在仓库之外）；
    grade_prep.py  write_ref_list 能从 SKILL.md 和 references 摘出含书名号、文号、标准号的行。

运行：python -m unittest discover -s tests/bench -t .
bench 脚本不是包，也不改动它们，这里按文件路径加载；需要仓库根目录的脚本，用 mock 把模块里的 REPO 指到临时假仓库。
"""
from __future__ import annotations

import contextlib
import importlib.util
import io
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

REPO_ROOT = Path(__file__).resolve().parents[2]
BENCH_DIR = REPO_ROOT / "evals" / "bench"
V01_RAW = REPO_ROOT / "evals" / "results" / "v0.1-raw.json"


def load_bench(name: str):
    """按文件路径加载 evals/bench/<name>.py（目录不是包，脚本之间靠 sys.path 互相导入）。"""
    if str(BENCH_DIR) not in sys.path:
        sys.path.insert(0, str(BENCH_DIR))
    spec = importlib.util.spec_from_file_location(f"bench_{name}", BENCH_DIR / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_main(module, argv):
    """用给定命令行参数运行脚本的 main()，返回 (退出码, 标准输出, 标准错误)。"""
    out, err = io.StringIO(), io.StringIO()
    with mock.patch.object(sys, "argv", [module.__file__, *argv]), \
            contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        try:
            code = module.main()
        except SystemExit as exc:
            code = exc.code if isinstance(exc.code, int) else 1
    return code, out.getvalue(), err.getvalue()


def write(path: Path, text: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(text.encode("utf-8"))
    return path


def write_json(path: Path, data) -> Path:
    return write(path, json.dumps(data, ensure_ascii=False, indent=1))


class TempDirTestCase(unittest.TestCase):
    def setUp(self) -> None:
        # resolve：Windows 的临时目录可能是 8.3 短路径，脚本内部用 resolve() 比较是否在仓库内
        self.tmp = Path(tempfile.mkdtemp(prefix="gongwen-bench-")).resolve()
        self.addCleanup(shutil.rmtree, str(self.tmp), ignore_errors=True)


# ---------------------------------------------------------------------------
# aggregate.py
# ---------------------------------------------------------------------------

def make_run(passes, fabrication_numeric=0, fabrication_other=0, lint=None):
    """一组运行结果：passes 为每条期望是否通过。"""
    return {
        "expectations": [{"idx": i, "text": f"期望{i}", "class": "通用" if i == 0 else "约定", "pass": p, "evidence": ""}
                         for i, p in enumerate(passes)],
        "numeric_fabrications": [{"verdict": "编造"}] * fabrication_numeric,
        "other_fabrications": ["x"] * fabrication_other,
        "lint": lint or [],
        "check_docx": [],
        "inputs_changed": [],
        "repo_mentions": [],
    }


class AggregateTest(TempDirTestCase):

    @classmethod
    def setUpClass(cls) -> None:
        cls.aggregate = load_bench("aggregate")

    def test_from_raw_reproduces_v01_numbers(self):
        # README 里的 v0.1 数字：全部期望 A 589/621、B 371/621；通用期望 A 401/422、B 315/422
        code, out, err = run_main(self.aggregate, ["--from-raw", str(V01_RAW)])
        self.assertEqual(code, 0, err)
        self.assertNotIn("不一致", err)  # 复算结果与文件保存的 summary 逐项相同
        self.assertIn("用例 104 条；缺失 0 项", out)
        self.assertIn("| 全部期望 | 589/621（94.8%） | 371/621（59.7%） |", out)
        self.assertIn("| 通用期望 | 401/422（95.0%） | 315/422（74.6%） |", out)
        self.assertIn("配对：A 过 B 不过 232 条；B 过 A 不过 14 条", out)

    def test_from_raw_summary_matches_saved_summary(self):
        raw = json.loads(V01_RAW.read_text(encoding="utf-8"))
        self.aggregate.SKILLS[:] = self.aggregate.ordered_unique([c["skill"] for c in raw["cases"]], REPO_ROOT)
        summary = self.aggregate.summarize(raw["cases"], raw["summary"].get("model_lines", {}))
        self.assertEqual(summary, raw["summary"])
        self.assertEqual(summary["overall"], {"A": [589, 621], "B": [371, 621]})
        self.assertEqual(summary["general_only"], {"A": [401, 422], "B": [315, 422]})
        self.assertEqual(summary["cases_all_pass"], {"A": [84, 104], "B": [8, 104]})

    def test_from_raw_writes_tables_file(self):
        tables = self.tmp / "表.md"
        code, out, err = run_main(self.aggregate, ["--from-raw", str(V01_RAW), "--tables", str(tables)])
        self.assertEqual(code, 0, err)
        self.assertEqual(tables.read_text(encoding="utf-8"), out.rstrip("\n") + "\n")

    def test_summarize_hand_made_cases(self):
        # 两条用例：甲含"防编造"标签。A 组 3+2 条期望全过，B 组各漏过一条；B 组在防编造用例里编造 1 个数字
        cases = [
            {"skill": "gongwen-x", "id": 1, "tags": ["防编造"],
             "runs": {"A": make_run([True, True, True]), "B": make_run([True, False, True], fabrication_numeric=1)}},
            {"skill": "gongwen-x", "id": 2, "tags": ["保密"],
             "runs": {"A": make_run([True, True]), "B": make_run([False, True])}},
        ]
        self.aggregate.SKILLS[:] = ["gongwen-x"]
        summary = self.aggregate.summarize(cases, {})
        self.assertEqual(summary["overall"], {"A": [5, 5], "B": [3, 5]})
        self.assertEqual(summary["general_only"], {"A": [2, 2], "B": [1, 2]})  # 每条用例的第 0 条期望算"通用"
        self.assertEqual(summary["convention_only"], {"A": [3, 3], "B": [2, 3]})
        self.assertEqual(summary["tag_防编造"], {"A": [3, 3], "B": [2, 3]})
        self.assertEqual(summary["tag_保密"], {"A": [2, 2], "B": [1, 2]})
        self.assertEqual(summary["cases_all_pass"], {"A": [2, 2], "B": [0, 2]})
        self.assertEqual(summary["paired"], {"A_only": 2, "B_only": 0})
        anti = summary["fabrication_anti_cases"]
        self.assertEqual((anti["A"]["numeric"], anti["B"]["numeric"]), (0, 1))
        self.assertEqual((anti["A"]["runs"], anti["B"]["runs_with_fabrication"]), (1, 1))
        self.assertEqual(summary["by_skill"]["gongwen-x"], {"A": [5, 5], "B": [3, 5]})

    def test_pct_handles_zero(self):
        self.assertEqual(self.aggregate.pct(1, 3), "33.3%")
        self.assertEqual(self.aggregate.pct(0, 0), "—")


# ---------------------------------------------------------------------------
# trigger_score.py
# ---------------------------------------------------------------------------

class TriggerScoreTest(TempDirTestCase):

    @classmethod
    def setUpClass(cls) -> None:
        cls.score = load_bench("trigger_score")

    def make_bench(self):
        bench = self.tmp / "bench"
        queries = [
            {"qid": "q1", "owner": "s1", "should_trigger": True, "query": "正例，选对"},
            {"qid": "q2", "owner": "s1", "should_trigger": True, "query": "正例，被选到别的 skill"},
            {"qid": "q3", "owner": "s1", "should_trigger": False, "query": "反例，选了别的 skill（对）"},
            {"qid": "q4", "owner": "s1", "should_trigger": False, "query": "反例，误选来源 skill（错）"},
            {"qid": "q5", "owner": "s2", "should_trigger": True, "query": "正例，没选任何 skill（错）"},
            {"qid": "q6", "owner": "s2", "should_trigger": False, "query": "反例，没选任何 skill（对）"},
            {"qid": "q7", "owner": "s2", "should_trigger": True, "query": "没有答案"},
        ]
        write_json(bench / "trigger" / "queries.json", queries)
        write_json(bench / "trigger" / "answers" / "batch-1.json",
                   {"answers": [{"id": "q1", "skill": "s1"}, {"id": "q2", "skill": "s2"}, {"id": "q3", "skill": "s2"}]})
        write_json(bench / "trigger" / "answers" / "batch-2.json",
                   {"answers": [{"id": "q4", "skill": "s1"}, {"id": "q5", "skill": "无"}, {"id": "q6", "skill": "无"}]})
        return bench

    def test_counts_positive_negative_missing_and_confusion(self):
        bench = self.make_bench()
        out_json, out_md = self.tmp / "trigger.json", self.tmp / "trigger.md"
        code, out, err = run_main(self.score, ["--bench", str(bench), "--out", str(out_json), "--tables", str(out_md)])
        self.assertEqual(code, 0, err)
        result = json.loads(out_json.read_text(encoding="utf-8"))
        summary = result["summary"]
        self.assertEqual(summary["total"], 6)  # 缺答案的 q7 不计入
        self.assertEqual(summary["correct"], 3)
        self.assertEqual(summary["positive"], [1, 3])  # 正例 q1 对，q2、q5 错
        self.assertEqual(summary["negative"], [2, 3])  # 反例 q3、q6 对，q4 错
        self.assertEqual(summary["missing"], ["q7"])
        self.assertEqual(summary["by_skill"], {"s1": {"pos": [1, 2], "neg": [1, 2]},
                                               "s2": {"pos": [0, 1], "neg": [1, 1]}})
        confusions = {(c["from"], c["chosen"]): c["count"] for c in summary["confusions"]}
        self.assertEqual(confusions, {("s1", "s2"): 1, ("s1（反例）", "s1"): 1, ("s2", "无"): 1})
        correct_by_id = {r["qid"]: r["correct"] for r in result["rows"]}
        self.assertEqual(correct_by_id, {"q1": True, "q2": False, "q3": True, "q4": False, "q5": False, "q6": True})
        text = out_md.read_text(encoding="utf-8")
        self.assertIn("| 全部说法 | 3/6 | 50.0% |", text)
        self.assertIn("| 正例（应触发本 skill） | 1/3 | 33.3% |", text)
        self.assertIn("| 反例（不应触发本 skill） | 2/3 | 66.7% |", text)
        self.assertIn("缺少答案：1 条", text)
        self.assertEqual(text, out[:-1])  # 屏幕输出是同一份文本，print 多一个换行

    def test_runs_without_output_files(self):
        code, out, err = run_main(self.score, ["--bench", str(self.make_bench())])
        self.assertEqual(code, 0, err)
        self.assertIn("| 全部说法 | 3/6 | 50.0% |", out)

    def test_perfect_run(self):
        bench = self.tmp / "bench"
        write_json(bench / "trigger" / "queries.json", [
            {"qid": "a", "owner": "s1", "should_trigger": True},
            {"qid": "b", "owner": "s1", "should_trigger": False},
        ])
        write_json(bench / "trigger" / "answers" / "batch-1.json",
                   {"answers": [{"id": "a", "skill": "s1"}, {"id": "b", "skill": "无"}]})
        out_json = self.tmp / "r.json"
        code, _out, err = run_main(self.score, ["--bench", str(bench), "--out", str(out_json)])
        self.assertEqual(code, 0, err)
        summary = json.loads(out_json.read_text(encoding="utf-8"))["summary"]
        self.assertEqual((summary["correct"], summary["total"], summary["missing"], summary["confusions"]),
                         (2, 2, [], []))


# ---------------------------------------------------------------------------
# prep.py
# ---------------------------------------------------------------------------

def make_fake_repo(root: Path, skill: str = "gongwen-fake") -> Path:
    """只含 1 个 skill、2 条评测用例的假仓库：用例 1 带一份材料，用例 2 不带。"""
    write_json(root / ".claude-plugin" / "marketplace.json",
               {"plugins": [{"name": "gongwen", "skills": [f"./skills/{skill}"]}]})
    write_json(root / "skills" / skill / "evals" / "evals.json", {
        "skill_name": skill,
        "evals": [
            {"id": 1, "prompt": "帮我写总结，素材在附件", "expected_output": "总结",
             "files": ["evals/files/素材.txt"], "expectations": ["不编造"],
             "unit_type": "公立医院", "role_mode": "高效模式", "tags": ["防编造"]},
            {"id": 2, "prompt": "你们能做什么", "expected_output": "介绍", "files": [],
             "expectations": ["说明能力"], "unit_type": "党政机关", "role_mode": "教学模式", "tags": []},
        ],
    })
    write(root / "skills" / skill / "evals" / "files" / "素材.txt", "某科室全年门诊 1000 人次。\n")
    return root


class PrepTest(TempDirTestCase):

    @classmethod
    def setUpClass(cls) -> None:
        cls.prep = load_bench("prep")

    def setUp(self) -> None:
        super().setUp()
        self.repo = make_fake_repo(self.tmp / "repo")
        patcher = mock.patch.object(self.prep, "REPO", self.repo)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.out = self.tmp / "bench"  # 评测目录在假仓库之外

    def run_prep(self, *extra):
        return run_main(self.prep, ["--out", str(self.out), "--skills", "gongwen-fake", "--tag", "t1", *extra])

    def workdir(self, eid, group):
        return self.out / "w" / self.prep.token("gongwen-fake", eid, group, "t1")

    def test_token_is_deterministic_and_opaque(self):
        token = self.prep.token
        self.assertEqual(token("s", 1, "A", "v1"), token("s", 1, "A", "v1"))
        self.assertEqual(len(token("s", 1, "A", "v1")), 10)
        self.assertEqual(len({token("s", 1, "A", "v1"), token("s", 1, "B", "v1"),
                              token("s", 2, "A", "v1"), token("s", 1, "A", "v2")}), 4)
        self.assertRegex(token("s", 1, "A", "v1"), r"^[0-9a-f]{10}$")  # 只含十六进制字符，看不出分组

    def test_full_build_creates_both_groups(self):
        code, out, err = self.run_prep()
        self.assertEqual(code, 0, err)
        index = json.loads((self.out / "index.json").read_text(encoding="utf-8"))
        self.assertEqual(sorted((r["id"], r["group"]) for r in index), [(1, "A"), (1, "B"), (2, "A"), (2, "B")])
        hashes = json.loads((self.out / "hashes.json").read_text(encoding="utf-8"))
        for group in "AB":
            copied = self.workdir(1, group) / "素材.txt"
            self.assertEqual(copied.read_text(encoding="utf-8"), "某科室全年门诊 1000 人次。\n")
            self.assertEqual(hashes[str(copied)], self.prep.sha256(copied))
            self.assertTrue(self.workdir(2, group).is_dir())
            self.assertTrue((self.out / "tasks" / f"gongwen-fake-1-{group}.md").is_file())
        self.assertEqual(len(hashes), 2)  # 只有用例 1 有材料，A、B 各一份
        self.assertEqual(json.loads((self.out / "meta.json").read_text(encoding="utf-8"))["tag"], "t1")
        inputs = (self.out / "inputs" / "gongwen-fake-1.txt").read_text(encoding="utf-8")
        self.assertIn("【用户消息】\n帮我写总结，素材在附件", inputs)
        self.assertIn("门诊 1000 人次", inputs)
        self.assertIn("用例 2 条", out)

    def test_groups_a_rebuilds_only_group_a(self):
        self.assertEqual(self.run_prep()[0], 0)
        hashes_before = json.loads((self.out / "hashes.json").read_text(encoding="utf-8"))
        # 模拟运行结果：A、B 两组的工作目录里各有一份产出，任务说明里各留一个标记
        for group in "AB":
            write(self.workdir(1, group) / "正文.txt", f"{group}组产出")
            write(self.out / "tasks" / f"gongwen-fake-1-{group}.md", f"旧任务说明-{group}")

        code, out, err = self.run_prep("--groups", "A")
        self.assertEqual(code, 0, err)
        self.assertIn("本次重建 A 组", out)
        self.assertNotIn("一并重建", out)

        a_dir, b_dir = self.workdir(1, "A"), self.workdir(1, "B")
        self.assertFalse((a_dir / "正文.txt").exists())  # A 组目录被清空重建
        self.assertEqual((b_dir / "正文.txt").read_text(encoding="utf-8"), "B组产出")  # B 组产出保留

        hashes = json.loads((self.out / "hashes.json").read_text(encoding="utf-8"))
        self.assertEqual(hashes, hashes_before)
        self.assertEqual((self.out / "tasks" / "gongwen-fake-1-B.md").read_text(encoding="utf-8"), "旧任务说明-B")
        self.assertNotEqual((self.out / "tasks" / "gongwen-fake-1-A.md").read_text(encoding="utf-8"), "旧任务说明-A")
        written = json.loads((self.out / "tasks_written.json").read_text(encoding="utf-8"))
        self.assertEqual(written, ["gongwen-fake-1-A"])

        index = json.loads((self.out / "index.json").read_text(encoding="utf-8"))
        self.assertEqual(len(index), 4)  # index 仍含两组，对应关系不丢

    def test_groups_a_rebuilds_b_when_material_changed(self):
        # 材料改了：B 组如果沿用旧输出，两组输入就不一致，必须一并重建
        self.assertEqual(self.run_prep()[0], 0)
        write(self.workdir(1, "B") / "正文.txt", "B组产出")
        write(self.repo / "skills/gongwen-fake/evals/files/素材.txt", "新素材。\n")
        code, out, err = self.run_prep("--groups", "A")
        self.assertEqual(code, 0, err)
        self.assertIn("一并重建：gongwen-fake-1-B", out)
        b_dir = self.workdir(1, "B")
        self.assertFalse((b_dir / "正文.txt").exists())
        self.assertEqual((b_dir / "素材.txt").read_text(encoding="utf-8"), "新素材。\n")
        written = json.loads((self.out / "tasks_written.json").read_text(encoding="utf-8"))
        self.assertEqual(written, ["gongwen-fake-1-A", "gongwen-fake-1-B"])
        b_task = (self.out / "tasks" / "gongwen-fake-1-B.md").read_text(encoding="utf-8")
        self.assertIn(str(b_dir), b_task)
        self.assertNotIn(str(self.workdir(2, "B")), b_task)  # 用例 2 的 B 组照旧保留，不进任务说明

    def test_groups_a_builds_b_for_new_case(self):
        self.assertEqual(self.run_prep()[0], 0)
        write(self.workdir(1, "B") / "正文.txt", "B组产出")
        data = json.loads((self.repo / "skills/gongwen-fake/evals/evals.json").read_text(encoding="utf-8"))
        data["evals"].append({"id": 3, "prompt": "新用例", "expected_output": "x", "files": [],
                              "expectations": ["y"], "unit_type": "高校", "role_mode": "高效模式", "tags": []})
        write_json(self.repo / "skills/gongwen-fake/evals/evals.json", data)
        code, out, err = self.run_prep("--groups", "A")
        self.assertEqual(code, 0, err)
        self.assertIn("一并重建：gongwen-fake-3-B", out)
        self.assertTrue(self.workdir(3, "B").is_dir())
        self.assertEqual((self.workdir(1, "B") / "正文.txt").read_text(encoding="utf-8"), "B组产出")
        b_task = (self.out / "tasks" / "gongwen-fake-1-B.md").read_text(encoding="utf-8")
        self.assertIn("## 请求 1", b_task)
        self.assertIn(str(self.workdir(3, "B")), b_task)
        self.assertNotIn(str(self.workdir(1, "B")), b_task)

    def test_groups_a_requires_existing_hashes(self):
        code, _out, err = self.run_prep("--groups", "A")
        self.assertEqual(code, 2)
        self.assertIn("hashes.json", err)

    def test_groups_a_rebuilds_b_when_hash_missing(self):
        # hashes.json 里没有 B 组材料的记录：无法确认 B 组输入，一并重建
        self.assertEqual(self.run_prep()[0], 0)
        hashes = json.loads((self.out / "hashes.json").read_text(encoding="utf-8"))
        hashes.pop(str(self.workdir(1, "B") / "素材.txt"))
        write_json(self.out / "hashes.json", hashes)
        code, out, err = self.run_prep("--groups", "A")
        self.assertEqual(code, 0, err)
        self.assertIn("一并重建：gongwen-fake-1-B", out)

    def test_out_inside_repo_is_refused(self):
        code, _out, err = run_main(self.prep, ["--out", str(self.repo / "bench"), "--skills", "gongwen-fake"])
        self.assertEqual(code, 2)
        self.assertIn("不要放在仓库内", err)
        self.assertFalse((self.repo / "bench").exists())


# ---------------------------------------------------------------------------
# grade_prep.py
# ---------------------------------------------------------------------------

class WriteRefListTest(TempDirTestCase):

    @classmethod
    def setUpClass(cls) -> None:
        cls.grade_prep = load_bench("grade_prep")

    def setUp(self) -> None:
        super().setUp()
        self.repo = self.tmp / "repo"
        self.out_dir = self.tmp / "refs"
        self.out_dir.mkdir()
        patcher = mock.patch.object(self.grade_prep, "REPO", self.repo)
        patcher.start()
        self.addCleanup(patcher.stop)
        skill = self.repo / "skills" / "gongwen-fake"
        write(skill / "SKILL.md", "# 演示\n\n依据《党政机关公文处理工作条例》第九条办理。\n没有任何引用的一行。\n")
        write(skill / "references" / "a.md",
              "- 文号示例：国办发〔2019〕 34 号\n"
              "  GB/T 9704-2012 规定版心尺寸。\n"
              "依据《党政机关公文处理工作条例》第九条办理。\n"  # 与 SKILL.md 重复，只留一条
              "单字书名号《短》不算。\n"
              "\n"
              "纯说明文字，不含文件名称。\n")
        write(skill / "references" / "sub" / "b.md", "子目录里的《关于加强基层减负工作的若干规定》也要摘。\n")
        write(skill / "references" / "notes.txt", "这行有《不会被读取的文件名称》，但不是 md。\n")

    def test_extracts_titles_document_numbers_and_standards(self):
        path = self.grade_prep.write_ref_list("gongwen-fake", self.out_dir)
        self.assertEqual(path, self.out_dir / "gongwen-fake.txt")
        lines = path.read_text(encoding="utf-8").splitlines()
        self.assertEqual(lines, [
            "依据《党政机关公文处理工作条例》第九条办理。",
            "- 文号示例：国办发〔2019〕 34 号",
            "GB/T 9704-2012 规定版心尺寸。",
            "子目录里的《关于加强基层减负工作的若干规定》也要摘。",
        ])

    def test_skill_without_references_dir(self):
        skill = self.repo / "skills" / "gongwen-bare"
        write(skill / "SKILL.md", "引用《整治形式主义为基层减负若干规定》。\n")
        path = self.grade_prep.write_ref_list("gongwen-bare", self.out_dir)
        self.assertEqual(path.read_text(encoding="utf-8"), "引用《整治形式主义为基层减负若干规定》。\n")

    def test_ref_line_pattern(self):
        pattern = self.grade_prep.REF_LINE_RE
        for text in ("《党政机关公文处理工作条例》", "中办发〔2012〕14号", "国发〔2018〕 12 号", "GB/T 15834-2011", "GB/T9704"):
            self.assertTrue(pattern.search(text), text)
        for text in ("《短》", "没有引用", "〔2019〕号", "文号 [2019] 34号"):
            self.assertFalse(pattern.search(text), text)


if __name__ == "__main__":
    unittest.main()
