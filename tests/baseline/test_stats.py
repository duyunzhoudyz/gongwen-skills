# -*- coding: utf-8 -*-
"""tools/build_baseline.py 的单元测试。

只用自拟的小样本，不依赖真实语料，也不读取仓库里的 baseline.json。
"""
from __future__ import annotations

import contextlib
import io
import json
import re
import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
TOOLS_DIR = REPO_ROOT / "tools"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import build_baseline as bb  # noqa: E402


def write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(text)


class PercentileTests(unittest.TestCase):
    def test_basic_quartiles(self):
        values = [1, 2, 3, 4, 5]
        self.assertEqual(bb.percentile(values, 50), 3)
        self.assertEqual(bb.percentile(values, 25), 2)
        self.assertEqual(bb.percentile(values, 75), 4)

    def test_interpolation(self):
        self.assertAlmostEqual(bb.percentile([10, 20], 50), 15.0)
        self.assertAlmostEqual(bb.percentile([0, 10, 20, 30], 90), 27.0)

    def test_single_and_empty(self):
        self.assertEqual(bb.percentile([7], 90), 7.0)
        self.assertIsNone(bb.percentile([], 50))

    def test_summarize(self):
        s = bb.summarize([1, 2, 3, 4, 100], 0)
        self.assertEqual(s["n"], 5)
        self.assertEqual(s["median"], 3)
        self.assertEqual(s["q1"], 2)
        self.assertEqual(s["q3"], 4)
        self.assertEqual(s["max"], 100)
        self.assertEqual(bb.summarize([], 1), {"n": 0})


class TextTests(unittest.TestCase):
    def test_clean_paragraphs_removes_wenhao_and_blank(self):
        lines = ["关于做好某项工作的通知 甲办发〔2026〕12号", "", "  各单位：  ", "参见《某办法》（乙发〔2025〕3号）。"]
        paras = bb.clean_paragraphs(lines)
        self.assertEqual(paras[0], "关于做好某项工作的通知")
        self.assertEqual(paras[1], "各单位：")
        self.assertEqual(paras[2], "参见《某办法》。")  # 文号及其留下的空括号一并去掉
        self.assertEqual(len(paras), 3)

    def test_count_chars_ignores_whitespace(self):
        self.assertEqual(bb.count_chars(["甲 乙　丙", "丁，戊。"]), 7)

    def test_split_sentences(self):
        paras = ["这是第一句话。这是第二句话；第三句话怎么样？", "标题没有句末标点", "最后一句没有标点结尾的残句不足"]
        lengths = bb.split_sentences(paras)
        # 第一段 3 句；第二段无句末标点不计；第三段无句末标点不计
        self.assertEqual(lengths, [6, 6, 7])

    def test_split_sentences_trailing_fragment(self):
        # 含句末标点的自然段，末尾无标点的残句 >=8 字才计
        self.assertEqual(bb.split_sentences(["第一句话在这里。后面又来了一段较长的无标点残句"]), [7, 15])
        self.assertEqual(bb.split_sentences(["第一句话在这里。残句"]), [7])


class HeadingTests(unittest.TestCase):
    def test_levels(self):
        paras = [
            "一、总体要求", "（一）指导思想。坚持底线思维。", "1. 明确目标", "（1）细化任务",
            "二、重点任务", "（二）强化保障", "1.5倍的增长不是标题", "2025年的工作", "三、其他",
        ]
        counts = bb.count_headings(paras)
        self.assertEqual(counts, {1: 3, 2: 2, 3: 1, 4: 1})

    def test_heading_level_none(self):
        self.assertEqual(bb.heading_level("一是加强领导。"), 0)
        self.assertEqual(bb.heading_level("正文第一段。"), 0)


class YishiTests(unittest.TestCase):
    def test_sets_and_items(self):
        paras = ["要抓好三件事。一是加强领导，二是完善机制，三是狠抓落实。", "另外，一是资金，二是人员。"]
        self.assertEqual(bb.count_yishi(paras), (2, 5))

    def test_not_inside_words(self):
        paras = ["统一是前提，唯一是目标，之一是关键。"]
        self.assertEqual(bb.count_yishi(paras), (0, 0))

    def test_paragraph_start(self):
        self.assertEqual(bb.count_yishi(["一是做好准备。", "二是加强协调。"]), (1, 2))


class FourCharTests(unittest.TestCase):
    def test_ratio(self):
        # 片段：加强领导(4) 完善机制(4) 狠抓落实(4) 各单位(3) -> 3/4
        ratio = bb.four_char_ratio(["加强领导，完善机制，狠抓落实，各单位。"])
        self.assertAlmostEqual(ratio, 0.75)

    def test_digits_and_latin_split_segments(self):
        # 数字和字母切断片段：“2025”“AI”不算汉字片段
        ratio = bb.four_char_ratio(["推进工作2025年落实AI应用"])
        # 片段：推进工作(4) 年落实(3) 应用(2) -> 1/3
        self.assertAlmostEqual(ratio, 1 / 3)

    def test_empty(self):
        self.assertIsNone(bb.four_char_ratio(["2025，123。"]))


class OpeningClosingTests(unittest.TestCase):
    def test_opening_purpose(self):
        paras = ["关于某事的通知", "各有关单位：", "为做好某项工作，现就有关事项通知如下。", "一、总体要求"]
        self.assertEqual(bb.classify_opening(paras), "目的起笔式")
        self.assertTrue(bb.has_address_line(paras))

    def test_opening_reply(self):
        paras = ["关于同意某事的批复", "某某单位：", "你单位关于申请设立某机构的请示收悉。经研究，同意。"]
        self.assertEqual(bb.classify_opening(paras), "来文答复式")

    def test_opening_basis_and_report(self):
        self.assertEqual(bb.classify_opening(["根据有关规定，现通知如下。"]), "依据起笔式")
        self.assertEqual(bb.classify_opening(["各位代表：", "现在，我代表市人民政府向大会报告工作，请予审议。"]), "代表政府报告式")

    def test_opening_meeting(self):
        self.assertEqual(bb.classify_opening(["同志们：", "刚才，几位同志作了发言，讲得都很好。"]), "会议起笔式")

    def test_opening_numbered(self):
        self.assertEqual(bb.classify_opening(["（一）明确责任。各地要明确责任。"]), "分条直陈式")

    def test_closing_tezhi(self):
        paras = ["正文第一段。", "请遵照执行。特此通知。", "某某办公室", "2026年1月5日"]
        self.assertEqual(bb.classify_closing(paras), "特此X式")

    def test_closing_qingshi(self):
        paras = ["情况如上。", "妥否，请批示。", "某某局", "2026年3月1日"]
        self.assertEqual(bb.classify_closing(paras), "请批示审议式")

    def test_closing_cifu(self):
        self.assertEqual(bb.classify_closing(["经研究，同意你单位的申请。此复。", "某某委", "2026年3月1日"]), "此复式")

    def test_closing_substantive(self):
        paras = ["各地要加强统筹协调，确保各项措施落地见效。", "某某部", "2026年3月1日"]
        self.assertEqual(bb.classify_closing(paras), "无固定结语（以实质内容收束）")

    def test_closing_report_to_superior_matters(self):
        paras = ["要加强统筹协调。重大事项及时按程序请示报告。"]
        self.assertEqual(bb.classify_closing(paras), "重大事项请示报告式")

    def test_closing_speech(self):
        self.assertEqual(bb.classify_closing(["让我们团结奋斗，开创新局面！谢谢大家！"]), "号召致谢式")

    def test_strip_signoff_glued_date(self):
        paras = ["正文结束。", "某某部某某局2026年5月22日"]
        self.assertEqual(bb.strip_signoff(paras), ["正文结束。"])
        # 以句号结尾的正文即使带日期也不能被当作落款
        paras = ["本通知自2026年5月22日起施行，请遵照执行。2026年5月22日"]
        self.assertEqual(bb.strip_signoff(paras), paras)


class StylePatternTests(unittest.TestCase):
    def make(self, **kw):
        item = {"id": "T-001", "pattern": "赋能", "density_per_1000": None}
        item.update(kw)
        return bb.StylePat(item)

    def test_hits_and_whitelist(self):
        pat = self.make()
        wl = [re.compile("数字赋能行动")]
        paras = ["要数字赋能行动。", "以科技赋能治理，以数据赋能决策。"]
        self.assertEqual(bb.count_pattern_hits(paras, pat, wl), 2)  # 白名单覆盖的一处不计
        self.assertEqual(bb.count_pattern_hits(paras, pat, []), 3)

    def test_exclude(self):
        pat = self.make(exclude=["赋能中心"])
        self.assertEqual(bb.count_pattern_hits(["赋能中心成立，赋能基层。"], pat, []), 1)

    def test_applies(self):
        pat = self.make(exempt_doc_types=["讲话稿"])
        self.assertFalse(pat.applies("讲话稿"))
        self.assertTrue(pat.applies("通知"))
        pat2 = self.make(doc_types=["通知"])
        self.assertTrue(pat2.applies("通知"))
        self.assertFalse(pat2.applies("函"))

    def test_load_style_patterns(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "p.json"
            write(path, json.dumps({
                "version": "9.9.9",
                "patterns": [{"id": "T-1", "pattern": "甲", "density_per_1000": 2}],
                "whitelist": [{"id": "WL-1", "pattern": "乙"}],
            }, ensure_ascii=False))
            pats, wl, version = bb.load_style_patterns(path)
            self.assertEqual(version, "9.9.9")
            self.assertEqual(len(pats), 1)
            self.assertEqual(pats[0].density, 2.0)
            self.assertEqual(len(wl), 1)


class MiscTests(unittest.TestCase):
    def test_org_key(self):
        self.assertEqual(bb.org_key("国家发展改革委 民政部"), "发展改革委")
        self.assertEqual(bb.org_key("发展改革委"), "发展改革委")
        self.assertEqual(bb.org_key("国务院办公厅"), "国务院办公厅")

    def test_pct(self):
        self.assertEqual(bb._pct(0), "0%")
        self.assertEqual(bb._pct(0.003), "<1%")
        self.assertEqual(bb._pct(0.073), "7%")
        self.assertEqual(bb._pct(1.0), "100%")

    def test_phrase_counts(self):
        counts = bb.phrase_counts("进一步加强，进一步推进，打造高地。")
        self.assertEqual(counts["进一步"], 2)
        self.assertEqual(counts["打造"], 1)

    def test_dump_json_roundtrip_and_stable(self):
        data = {"a": 1, "b": {"x": [1, 2, 3], "y": {"p": 1.5, "q": "中文"}}, "c": [{"k": 1}, {"k": 2}], "d": {}}
        text = bb.dump_json(data)
        self.assertEqual(json.loads(text), data)
        self.assertEqual(bb.dump_json(data), text)
        self.assertTrue(text.endswith("\n"))
        self.assertNotIn("\\u", text)  # 中文不转义


def make_corpus(root: Path) -> None:
    """造一个 3 文种、共 13 篇的小语料（全部自拟）。"""
    docs = []

    def add(doc_type: str, name: str, org: str, text: str, subtype: str = "") -> None:
        write(root / doc_type / f"{name}.txt", text)
        docs.append({
            "doc_type": doc_type, "file": f"{doc_type}/{name}.txt", "title": f"自拟样本{name}",
            "org": org, "date": "2026-01-02", "url": f"https://example.gov.cn/{name}", "segment": "main",
            "subtype": subtype,
        })

    notice = ("关于做好某项工作的通知\n各有关单位：\n为做好某项工作，现就有关事项通知如下。\n"
              "一、加强组织领导\n各单位要进一步加强组织领导，切实落实责任，确保工作取得实效。\n"
              "二、狠抓工作落实\n一是明确任务，二是压实责任，三是强化督促。\n特此通知。\n某某办公室\n2026年1月2日\n")
    for i in range(5):
        add("通知", f"n{i}", f"某部门{i}", notice)
    reply = "关于同意设立某机构的批复\n某某单位：\n你单位关于设立某机构的请示收悉。经研究，同意设立。\n请按有关规定抓好落实。\n某某委\n2026年2月3日\n"
    for i in range(4):
        add("批复", f"r{i}", f"某委{i % 2}", reply)
    speech = ("在某某工作会议上的讲话\n同志们：\n刚才，几位同志作了发言。下面，我讲三点意见。\n"
              "一、提高认识。这项工作是当前的重要任务，要打造一流队伍。\n二、狠抓落实。要进一步加强统筹，赋能基层。\n"
              "让我们团结奋斗，谢谢大家！\n")
    for i in range(3):
        add("讲话", f"s{i}", f"某机关{i}", speech, subtype="甲类" if i < 2 else "乙类")
    add("讲话", "s3", "某机关3", speech, subtype="乙类")
    write(root / "sources.json", json.dumps({"meta": {"fetched": "2026-10-21"}, "docs": docs}, ensure_ascii=False))


class BuildTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.corpus = self.root / "corpus"
        make_corpus(self.corpus)
        self.patterns = self.root / "patterns.json"
        write(self.patterns, json.dumps({
            "version": "0.0.1",
            "patterns": [
                {"id": "T-1", "group": "style", "category": "新八股词", "level": "P2", "pattern": "打造|赋能", "density_per_1000": 1},
                {"id": "T-2", "group": "style", "category": "不存在", "level": "P1", "pattern": "绝不出现的词"},
            ],
            "whitelist": [],
        }, ensure_ascii=False))

    def tearDown(self):
        self._tmp.cleanup()

    def build(self):
        docs, meta = bb.load_corpus(self.corpus)
        return bb.build_baseline(docs, meta, self.patterns)

    def test_groups_and_counts(self):
        data = self.build()
        groups = data["groups"]
        self.assertEqual(list(groups)[:3], ["通知", "批复", "讲话"])  # 按文种顺序
        self.assertEqual(groups["通知"]["n"], 5)
        self.assertEqual(groups["通知"]["org_count"], 5)
        self.assertEqual(groups["批复"]["org_count"], 2)
        self.assertTrue(groups["批复"]["low_sample"])      # 4 篇 < 10
        self.assertTrue(groups["批复"]["low_orgs"])        # 2 个机关 < 3

    def test_subgroup_threshold(self):
        data = self.build()
        # 甲类 2 篇、乙类 2 篇，都少于 MIN_SUBGROUP(5)，不应单独成组
        self.assertNotIn("讲话/甲类", data["groups"])

    def test_metrics_values(self):
        data = self.build()
        notice = data["groups"]["通知"]
        self.assertEqual(notice["headings"]["h1_per_doc"]["median"], 2.0)
        self.assertEqual(notice["yishi"]["sets_per_doc"]["median"], 1.0)
        self.assertEqual(notice["openings"]["目的起笔式"]["share"], 1.0)
        self.assertEqual(notice["closings"]["特此X式"]["n"], 5)
        self.assertEqual(notice["address_line_share"], 1.0)
        speech = data["groups"]["讲话"]
        self.assertEqual(speech["openings"]["会议起笔式"]["n"], 4)
        self.assertEqual(speech["closings"]["号召致谢式"]["share"], 1.0)

    def test_style_pattern_comparison(self):
        data = self.build()
        speech = data["groups"]["讲话"]["style_patterns"]["T-1"]
        self.assertEqual(speech["docs_with_hit"], 4)
        self.assertEqual(speech["doc_share"], 1.0)
        self.assertEqual(speech["threshold"], 1.0)
        # 每篇命中 2 处，字数不足 1000，千字密度远大于阈值 1
        self.assertEqual(speech["over_threshold_share"], 1.0)
        self.assertNotIn("T-1", data["groups"]["通知"]["style_patterns"])
        self.assertEqual(data["style_patterns_zero_hit"], ["T-2"])
        self.assertIn("T-1", data["style_patterns_overview"])

    def test_sources_no_fulltext(self):
        data = self.build()
        self.assertEqual(len(data["sources"]), 13)
        first = data["sources"][0]
        self.assertEqual(set(first), {"doc_type", "subtype", "title", "org", "date", "url", "segment", "chars", "sha256_16"})
        self.assertNotIn("某某办公室", bb.dump_json(data))  # 正文内容不入库

    def test_deterministic(self):
        self.assertEqual(bb.dump_json(self.build()), bb.dump_json(self.build()))

    def test_main_write_and_check(self):
        out = self.root / "out" / "baseline.json"
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            code = bb.main(["--corpus", str(self.corpus), "--out", str(out), "--patterns", str(self.patterns)])
        self.assertEqual(code, 0)
        self.assertTrue(out.is_file())
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(bb.main(["--corpus", str(self.corpus), "--out", str(out), "--patterns", str(self.patterns), "--check"]), 0)
        # 篡改输出后 --check 应报不一致
        write(out, out.read_text(encoding="utf-8").replace('"n": 5', '"n": 6', 1))
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(bb.main(["--corpus", str(self.corpus), "--out", str(out), "--patterns", str(self.patterns), "--check"]), 1)

    def test_main_bad_corpus(self):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            code = bb.main(["--corpus", str(self.root / "nope"), "--out", str(self.root / "x.json")])
        self.assertEqual(code, 2)
        self.assertIn("sources.json", buf.getvalue())

    def test_missing_file_reported(self):
        (self.corpus / "通知" / "n0.txt").unlink()
        with self.assertRaises(bb.CorpusError):
            bb.load_corpus(self.corpus)

    def test_missing_field_reported(self):
        path = self.corpus / "sources.json"
        data = json.loads(path.read_text(encoding="utf-8"))
        del data["docs"][0]["org"]
        write(path, json.dumps(data, ensure_ascii=False))
        with self.assertRaises(bb.CorpusError):
            bb.load_corpus(self.corpus)

    def test_render_tables(self):
        text = bb.render_tables(self.build())
        for name in ("通知", "批复", "讲话"):
            self.assertIn(f"| {name}", text)
        self.assertIn("样本偏少", text)


if __name__ == "__main__":
    unittest.main()
