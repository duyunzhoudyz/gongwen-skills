# -*- coding: utf-8 -*-
"""skills/gongwen-speech/scripts/speech_check.py 的单元测试。

运行：python -m unittest discover -s tests/speech -t .

“必过标本”是一篇规范的表态发言，在目标时长内，不应报任何问题；
“必错标本”逐项构造超时、长句、数字过密、段落过长、缺称谓、来源不明的引文等毛病，必须报出。
标本均为虚构，场景和数字与 skills/gongwen-speech/evals 中的评测材料无关。
"""
from __future__ import annotations

import importlib.util
import io
import json
import shutil
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from typing import Any, Dict, List

REPO_ROOT = Path(__file__).resolve().parents[2]
SCRIPT_PATH = REPO_ROOT / "skills" / "gongwen-speech" / "scripts" / "speech_check.py"


def _load():
    name = "speech_check_under_test"
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, SCRIPT_PATH)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)  # type: ignore[union-attr]
    return module


sc = _load()

try:
    import docx  # type: ignore  # noqa: F401
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# 必过标本：规范的表态发言（虚构）
PASS_SAMPLE = """在全县森林防火工作会议上的表态发言
某乡乡长　张某某
（2027年3月6日）
各位领导，同志们：
刚才，县领导对今年森林防火工作作了部署。我代表某乡表个态。
今年我乡的目标是，不发生较大森林火灾。围绕这个目标，我们抓好三件事。
一是管住火源。清明前，把进山的7个路口全部设卡。每个卡口两人值守，进山人员逐一登记。
二是清理隐患。4月底前，把林区坟头周边的可燃物清理一遍。每个村明确一名清理责任人。
三是守好值班。防火期实行乡村干部轮流值班。发现火情后，扑火队半小时内赶到现场。
同志们，森林防火没有旁观者。我们一定按会议要求，把每一项措施落到实处。
我的发言完毕，谢谢大家。
"""

SHORT_PARA = "我们要把每一项任务落实到人，按时完成。"


def body_chars(text: str) -> int:
    """独立计算：从称谓行起，不含空白。"""
    lines = [line for line in text.splitlines() if line.strip()]
    start = next(i for i, line in enumerate(lines) if line.endswith("："))
    return sum(len("".join(line.split())) for line in lines[start:])


class SpeechCheckTestCase(unittest.TestCase):

    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="speech_check_test_"))

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp, ignore_errors=True)

    def write(self, text: str, name: str = "正文.txt", encoding: str = "utf-8") -> Path:
        path = self.tmp / name
        path.write_bytes(text.encode(encoding))
        return path

    def run_main(self, *args: str):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = sc.main(list(args))
        return code, out.getvalue(), err.getvalue()

    def check(self, text: str, *args: str) -> Dict[str, Any]:
        path = self.write(text)
        code, out, err = self.run_main(str(path), "--json", *args)
        self.assertEqual(code, 0, err)
        return json.loads(out)

    @staticmethod
    def ids(result: Dict[str, Any]) -> List[str]:
        return [issue["id"] for issue in result["issues"]]


class PassSampleTests(SpeechCheckTestCase):

    def test_pass_sample_has_no_issue(self):
        expected = body_chars(PASS_SAMPLE)
        minutes = round(expected / 200, 1)
        result = self.check(PASS_SAMPLE, "--minutes", str(minutes))
        self.assertEqual(result["issues"], [])
        self.assertEqual(result["char_count"], expected)
        self.assertTrue(result["within_target"])
        self.assertEqual(result["salutation"], "各位领导，同志们：")

    def test_header_lines_not_counted(self):
        result = self.check(PASS_SAMPLE)
        self.assertEqual(result["header"], ["在全县森林防火工作会议上的表态发言", "某乡乡长　张某某", "（2027年3月6日）"])
        self.assertEqual(result["char_count"], body_chars(PASS_SAMPLE))

    def test_pass_sample_strict_returns_zero(self):
        path = self.write(PASS_SAMPLE)
        code, _out, _err = self.run_main(str(path), "--strict")
        self.assertEqual(code, 0)

    def test_dates_and_list_numbers_not_counted(self):
        text = "同志们：\n1. 2027年3月1日至4月28日，开展第二轮巡查，巡查林区8片。\n"
        result = self.check(text)
        self.assertNotIn("ORAL-002", self.ids(result))

    def test_text_output_readable(self):
        path = self.write(PASS_SAMPLE)
        code, out, _err = self.run_main(str(path), "--minutes", "3")
        self.assertEqual(code, 0)
        self.assertIn("讲话稿检查", out)
        self.assertIn("预计", out)

    def test_date_range_not_counted(self):
        # “1—6月”中的“1”属于日期范围，不算数字；本段只有 3200 一个数
        text = "同志们：\n今年1—6月，全市新增就业岗位3200个，完成全年任务的一半。\n"
        self.assertNotIn("ORAL-002", self.ids(self.check(text, "--max-numbers", "1")))
        more = "同志们：\n今年1—6月，全市新增就业岗位3200个，其中高校毕业生1150人。\n"
        self.assertIn("ORAL-002", self.ids(self.check(more, "--max-numbers", "1")))

    def test_long_salutation_not_reported_as_sentence(self):
        salutation = "尊敬的省某厅某某副厅长，尊敬的市政府某某副市长，尊敬的各位来宾，各位领导、同志们："
        self.assertGreater(len(salutation), 40)
        result = self.check(salutation + "\n今天开会，我讲三点。\n")
        self.assertEqual(result["salutation"], salutation)
        self.assertNotIn("ORAL-001", self.ids(result))

    def test_very_long_salutation_still_detected(self):
        names = "，".join(f"尊敬的某某单位某某职务第{i}位来宾" for i in range(1, 6))
        salutation = names + "，各位领导、同志们："
        self.assertGreater(len(salutation), 60)
        result = self.check("在某某会议上的致辞\n" + salutation + "\n今天开会，我讲三点。\n")
        self.assertEqual(result["salutation"], salutation)
        self.assertNotIn("ORAL-004", self.ids(result))
        self.assertNotIn("ORAL-001", self.ids(result))


class TimeTests(SpeechCheckTestCase):

    def test_too_short_for_target(self):
        result = self.check(PASS_SAMPLE, "--minutes", "10")
        self.assertIn("TIME-001", self.ids(result))
        self.assertFalse(result["within_target"])
        self.assertEqual(result["target_range"], [1800, 2200])

    def test_too_long_for_target(self):
        result = self.check(PASS_SAMPLE, "--minutes", "1")
        issue = next(i for i in result["issues"] if i["id"] == "TIME-001")
        self.assertEqual(issue["level"], "P1")
        self.assertIn("超出", issue["message"])

    def test_text_verdict_wording(self):
        path = self.write(PASS_SAMPLE)
        self.assertIn("篇幅不足", self.run_main(str(path), "--minutes", "10")[1])
        self.assertIn("篇幅超出", self.run_main(str(path), "--minutes", "1")[1])
        minutes = str(round(body_chars(PASS_SAMPLE) / 200, 1))
        out = self.run_main(str(path), "--minutes", minutes)[1]
        self.assertIn("在范围内", out)
        self.assertNotIn("超出范围", out)

    def test_rate_changes_estimate(self):
        chars = body_chars(PASS_SAMPLE)
        result = self.check(PASS_SAMPLE, "--rate", "160")
        self.assertEqual(result["est_minutes"], round(chars / 160, 1))

    def test_over_one_hour(self):
        text = "同志们：\n" + "\n".join([SHORT_PARA] * 700)
        result = self.check(text)
        self.assertGreater(result["char_count"], 12000)
        self.assertIn("TIME-002", self.ids(result))

    def test_strict_returns_one_on_p1(self):
        path = self.write(PASS_SAMPLE)
        code, _out, _err = self.run_main(str(path), "--minutes", "10", "--strict")
        self.assertEqual(code, 1)

    def test_markers_excluded_from_count(self):
        plain = "同志们：\n今年全镇排查隐患处。\n"
        marked = "同志们：\n今年全镇排查隐患【待补：数量】处。【待核实：统计口径】\n"
        a = self.check(plain)
        b = self.check(marked)
        self.assertEqual(a["char_count"], b["char_count"])
        self.assertEqual(b["markers"], {"待补": 1, "待核实": 1})

    def test_draft_note_truncates(self):
        text = PASS_SAMPLE + "【起草说明】\n" + "\n".join([SHORT_PARA] * 50)
        result = self.check(text)
        self.assertEqual(result["char_count"], body_chars(PASS_SAMPLE))
        self.assertIsNotNone(result["truncated_line"])

    def test_sections(self):
        text = "同志们：\n今天开会，我讲两点。\n一、把隐患查清楚\n全镇一处一处查。\n二、把责任压下去\n每村明确一名责任人。\n"
        result = self.check(text)
        headings = [s["heading"] for s in result["sections"]]
        self.assertEqual(headings, ["开场", "一、把隐患查清楚", "二、把责任压下去"])
        self.assertEqual(sum(s["chars"] for s in result["sections"]), result["char_count"])


class OralTests(SpeechCheckTestCase):

    def test_long_sentence(self):
        long_sentence = "通过全面构建多层次、全方位、立体化的风险防控体系，实现了全市安全生产形势的持续稳定向好和根本好转。"
        result = self.check("同志们：\n" + long_sentence + "\n")
        self.assertIn("ORAL-001", self.ids(result))

    def test_style_nian_relaxes_sentence_limit(self):
        sentence = "全市各级各部门要按照会议部署，把排查整治的各项任务分解到具体单位、具体岗位和具体责任人。"
        self.assertIn("ORAL-001", self.ids(self.check("同志们：\n" + sentence + "\n")))
        self.assertNotIn("ORAL-001", self.ids(self.check("同志们：\n" + sentence + "\n", "--style", "念")))

    def test_number_dense_paragraph(self):
        text = "同志们：\n今年检修换热站386座，更换阀门2140个，改造管网47公里，涉及小区53个，投入资金960万元。\n"
        result = self.check(text)
        issue = next(i for i in result["issues"] if i["id"] == "ORAL-002")
        self.assertIn("5 个数字", issue["message"])

    def test_custom_max_numbers(self):
        text = "同志们：\n今年检修换热站386座，更换阀门2140个。\n"
        self.assertNotIn("ORAL-002", self.ids(self.check(text)))
        self.assertIn("ORAL-002", self.ids(self.check(text, "--max-numbers", "1")))

    def test_long_paragraph(self):
        text = "同志们：\n" + "".join([SHORT_PARA] * 15) + "\n"
        result = self.check(text)
        self.assertIn("ORAL-003", self.ids(result))

    def test_missing_salutation(self):
        text = "今天召开动员部署会。我讲三点。\n一、把隐患查清楚\n"
        result = self.check(text)
        self.assertIn("ORAL-004", self.ids(result))
        self.assertIsNone(result["salutation"])

    def test_single_salutation_in_long_speech(self):
        text = "同志们：\n" + "\n".join([SHORT_PARA] * 120)
        self.assertIn("ORAL-005", self.ids(self.check(text)))
        self.assertNotIn("ORAL-005", self.ids(self.check(text, "--style", "念")))
        with_more = "同志们：\n" + "\n".join([SHORT_PARA] * 60) + "\n同志们，最后再强调一点。\n" + "\n".join([SHORT_PARA] * 60)
        self.assertNotIn("ORAL-005", self.ids(self.check(with_more)))


class QuoteTests(SpeechCheckTestCase):

    def test_unsourced_quote(self):
        text = "同志们：\n总书记指出：“安全生产是民生大事，一丝一毫不能放松。”\n"
        result = self.check(text)
        self.assertIn("QUOTE-001", self.ids(result))
        self.assertEqual(len(result["quotes"]), 1)
        self.assertFalse(result["quotes"][0]["has_source"])

    def test_sourced_quote_passes(self):
        text = "同志们：\n上级在《关于开展安全生产大检查的通知》中要求：“全面排查，不留死角。”\n"
        result = self.check(text)
        self.assertNotIn("QUOTE-001", self.ids(result))
        self.assertTrue(result["quotes"][0]["has_source"])

    def test_marked_quote_passes(self):
        text = "同志们：\n总书记指出：“安全生产是民生大事。”【待核实：出处】\n"
        result = self.check(text)
        self.assertNotIn("QUOTE-001", self.ids(result))
        self.assertTrue(result["quotes"][0]["marked"])

    def test_suoshuo_de_quote_found(self):
        text = "同志们：\n正如总书记所指出的：“保护生态环境就是保护生产力。”\n"
        result = self.check(text)
        self.assertEqual(len(result["quotes"]), 1)
        self.assertIn("QUOTE-001", self.ids(result))

    def test_quote_with_meeting_source(self):
        text = "同志们：\n总书记在全国生态环境保护大会上所指出的：“要坚持把绿色发展摆在突出位置。”\n"
        result = self.check(text)
        self.assertEqual(len(result["quotes"]), 1)
        self.assertTrue(result["quotes"][0]["has_source"])
        self.assertNotIn("QUOTE-001", self.ids(result))

    def test_nayang_and_lunshu_quotes_found(self):
        text = ("同志们：\n正如总书记强调的那样，“发展是第一要务。”\n"
                "我们要深入学习总书记关于生态文明建设的重要论述：“生态兴则文明兴。”\n")
        result = self.check(text)
        self.assertEqual(len(result["quotes"]), 2)
        self.assertEqual(self.ids(result).count("QUOTE-001"), 2)

    def test_partial_quote_after_comma_and_yao(self):
        # 必错标本：“强调，要“……””式的部分引用，没有出处
        text = "同志们：\n总书记强调，要“坚持以人民为中心的发展思想”。\n"
        result = self.check(text)
        self.assertEqual(len(result["quotes"]), 1)
        self.assertIn("QUOTE-001", self.ids(result))

    def test_partial_quote_after_comma_and_bixu_with_source(self):
        # 必过标本：“指出，必须“……””式的部分引用，注明了会议出处
        text = "同志们：\n总书记在全国生态环境保护大会上指出，必须“把绿色发展摆在突出位置”。\n"
        result = self.check(text)
        self.assertEqual(len(result["quotes"]), 1)
        self.assertTrue(result["quotes"][0]["has_source"])
        self.assertNotIn("QUOTE-001", self.ids(result))

    def test_term_in_quotes_is_not_quote(self):
        text = "同志们：\n各单位要落实“三重一大”决策制度，按规定上会研究。\n"
        result = self.check(text)
        self.assertEqual(result["quotes"], [])

    def test_vague_source_p1(self):
        text = "同志们：\n古人云：“千里之堤，溃于蚁穴。”小隐患不处理，就会酿成大事故。\n"
        result = self.check(text)
        issue = next(i for i in result["issues"] if i["id"] == "QUOTE-002")
        self.assertEqual(issue["level"], "P1")
        self.assertNotIn("QUOTE-001", self.ids(result))

    def test_vague_source_p2(self):
        text = "同志们：\n俗话说，磨刀不误砍柴工。正如刚才张局长所说，准备工作要做足。\n"
        result = self.check(text)
        levels = [i["level"] for i in result["issues"] if i["id"] == "QUOTE-003"]
        self.assertEqual(levels, ["P2", "P2"])

    def test_vague_source_strict(self):
        path = self.write("同志们：\n古人云：“千里之堤，溃于蚁穴。”\n")
        code, _out, _err = self.run_main(str(path), "--strict")
        self.assertEqual(code, 1)


class InputTests(SpeechCheckTestCase):

    def test_missing_file(self):
        code, _out, err = self.run_main(str(self.tmp / "不存在.txt"))
        self.assertEqual(code, 2)
        self.assertIn("输入错误", err)

    def test_unsupported_suffix(self):
        path = self.write(PASS_SAMPLE, name="正文.pdf")
        code, _out, _err = self.run_main(str(path))
        self.assertEqual(code, 2)

    def test_invalid_minutes_and_rate(self):
        path = self.write(PASS_SAMPLE)
        self.assertEqual(self.run_main(str(path), "--minutes", "0")[0], 2)
        self.assertEqual(self.run_main(str(path), "--rate", "-5")[0], 2)
        self.assertEqual(self.run_main(str(path), "--max-sentence", "0")[0], 2)

    def test_bad_argument_returns_two(self):
        path = self.write(PASS_SAMPLE)
        self.assertEqual(self.run_main(str(path), "--style", "唱")[0], 2)

    def test_help_returns_zero(self):
        code, out, _err = self.run_main("--help")
        self.assertEqual(code, 0)
        self.assertIn("用法：", out)

    def test_empty_body(self):
        path = self.write("\n\n【起草说明】\n说明\n")
        self.assertEqual(self.run_main(str(path))[0], 2)

    def test_gb18030_file(self):
        path = self.write(PASS_SAMPLE, encoding="gb18030")
        code, out, _err = self.run_main(str(path), "--json")
        self.assertEqual(code, 0)
        self.assertEqual(json.loads(out)["char_count"], body_chars(PASS_SAMPLE))

    @unittest.skipUnless(HAS_DOCX, "未安装 python-docx")
    def test_docx_input(self):
        import docx  # type: ignore
        document = docx.Document()
        for line in PASS_SAMPLE.strip().splitlines():
            document.add_paragraph(line)
        path = self.tmp / "正文.docx"
        document.save(str(path))
        code, out, _err = self.run_main(str(path), "--json")
        self.assertEqual(code, 0)
        result = json.loads(out)
        self.assertEqual(result["char_count"], body_chars(PASS_SAMPLE))
        self.assertEqual(result["issues"], [])

    def test_script_writes_no_file(self):
        path = self.write(PASS_SAMPLE)
        before = sorted(p.name for p in self.tmp.iterdir())
        self.run_main(str(path), "--json")
        self.assertEqual(sorted(p.name for p in self.tmp.iterdir()), before)


if __name__ == "__main__":
    unittest.main()
