# -*- coding: utf-8 -*-
"""正文结构自动识别：标题、主送、层次标题、附件说明、署名、成文日期、附注、版记。"""
import re
import time
import unittest

from tests.format.helpers import FILES, OUT, gw_layout


def parse(text: str):
    return gw_layout.parse_lines(text.split("\n"))


class TitleRecipientTest(unittest.TestCase):
    def test_guanyu_title_and_recipient(self):
        r = parse("关于做好2026年防汛工作的通知\n\n各乡镇人民政府：\n\n现将有关事项通知如下。")
        self.assertEqual(r.fields["标题"], "关于做好2026年防汛工作的通知")
        self.assertEqual(r.fields["主送机关"], "各乡镇人民政府")
        self.assertEqual(r.roles(), ["title", "recipient", "body"])
        self.assertEqual(r.doc_type, "通知")
        self.assertEqual(r.kind, "法定")

    def test_title_ending_with_doc_type(self):
        r = parse("某医院医务部2026年工作总结\n\n2026年，医务部完成了三项工作。")
        self.assertEqual(r.fields["标题"], "某医院医务部2026年工作总结")
        self.assertEqual(r.kind, "材料")

    def test_multiline_title_merged(self):
        r = parse("关于印发某区2026年基层减负工作方案的\n通知\n\n各街道办事处：\n\n现印发给你们。")
        self.assertEqual(r.fields["标题"], "关于印发某区2026年基层减负工作方案的通知")
        self.assertEqual([it.role for it in r.items][:2], ["title", "title"])
        self.assertEqual(r.fields["主送机关"], "各街道办事处")

    def test_inferred_title_warns(self):
        r = parse("某社区2026年第三季度工作情况\n\n本季度，社区开展了五项活动。")
        self.assertEqual(r.fields["标题"], "某社区2026年第三季度工作情况")
        self.assertTrue(any("推断" in w for w in r.warnings))

    def test_no_title(self):
        r = parse("各科室：\n\n请于周五前报送材料。")
        self.assertNotIn("标题", r.fields)
        self.assertEqual(r.fields["主送机关"], "各科室")
        self.assertTrue(any("未识别到标题" in w for w in r.warnings))

    def test_half_width_colon_recipient(self):
        r = parse("关于调整作息时间的通知\n\n各部门:\n\n自11月1日起执行冬季作息时间。")
        self.assertEqual(r.fields["主送机关"], "各部门")
        self.assertTrue(any("半角冒号" in w for w in r.warnings))

    def test_speech_subtitle_and_address(self):
        r = parse("扎实推进安全生产工作\n——在全区安全生产工作会议上的讲话\n\n同志们：\n\n今天召开会议。")
        self.assertEqual(r.roles(), ["title", "subtitle", "recipient", "body"])


class TailTest(unittest.TestCase):
    def test_signature_date_note_banji(self):
        text = (FILES / "红头通知-正文.txt").read_text(encoding="utf-8")
        lines, _ = gw_layout.read_body_text(FILES / "红头通知-正文.txt")
        r = gw_layout.parse_lines(lines)
        self.assertEqual(r.fields["发文机关署名"], ["某市应急管理局"])
        self.assertEqual(r.fields["成文日期"], "2026年10月28日")
        self.assertEqual(r.fields["附注"], "（此件公开发布）")
        self.assertEqual(len(r.fields["附件"]), 2)
        self.assertIn("【起草说明】", text)
        self.assertFalse(any("所作假设" in ln for ln in lines))  # 【起草说明】之后不读

    def test_cc_and_print_lines(self):
        r = parse("关于某事的通知\n\n各单位：\n\n现通知如下。\n\n某局\n\n2026年3月5日\n\n"
                  "抄送：市委办公室，市政府办公室。\n\n某局办公室　2026年3月5日印发")
        self.assertEqual(r.fields["抄送"], "市委办公室，市政府办公室。")
        self.assertEqual(r.fields["印发机关"], "某局办公室")
        self.assertEqual(r.fields["印发日期"], "2026年3月5日")
        self.assertEqual(r.roles()[-4:], ["signature", "date", "cc", "print"])

    def test_joint_signatures(self):
        r = parse("关于联合开展检查的通知\n\n各区：\n\n现通知如下。\n\n某市教育局\n\n某市卫生健康委员会\n\n2026年9月1日")
        self.assertEqual(r.fields["发文机关署名"], ["某市教育局", "某市卫生健康委员会"])

    def test_closing_phrase_not_signature(self):
        r = parse("关于某事的报告\n\n市政府：\n\n现报告如下。\n\n特此报告\n\n某局\n\n2026年3月5日")
        self.assertEqual(r.fields["发文机关署名"], ["某局"])
        self.assertIn("body", [it.role for it in r.items if it.text == "特此报告"])

    def test_chinese_numeral_date_warns(self):
        r = parse("关于某事的通知\n\n各单位：\n\n现通知如下。\n\n某局\n\n二〇二六年十月十六日")
        self.assertEqual(r.fields["成文日期"], "二〇二六年十月十六日")
        self.assertTrue(any("阿拉伯数字" in w for w in r.warnings))

    def test_single_attachment(self):
        r = parse("关于某事的通知\n\n各单位：\n\n现通知如下。\n\n附件：报名表\n\n某局\n\n2026年3月5日")
        self.assertEqual(r.fields["附件"], ["报名表"])


class HeadingTest(unittest.TestCase):
    def test_levels_and_split(self):
        r = parse("关于某事的通知\n\n各单位：\n\n一、总体要求\n\n（一）强化组织领导。成立工作专班。\n\n"
                  "1.制定方案。各单位制定方案。\n\n（1）明确分工\n\n一、二季度共完成检查20次。")
        roles = [(it.role, it.split) for it in r.items[2:]]
        self.assertEqual(roles[0], ("h1", len("一、总体要求")))
        self.assertEqual(roles[1], ("h2", len("（一）强化组织领导。")))
        self.assertEqual(roles[2][0], "h3")
        self.assertEqual(roles[3][0], "h4")
        self.assertEqual(roles[4][0], "body")  # “一、二季度……。”按正文处理

    def test_attendance_lines(self):
        r = parse("某局局长办公会议纪要\n\n会议研究了两项议题。\n\n出席：张某、李某\n\n列席：办公室负责人")
        self.assertEqual([it.role for it in r.items][-2:], ["attendance", "attendance"])

    def test_secret_mark_warning(self):
        lines, _ = gw_layout.read_body_text(FILES / "涉密材料-正文.txt")
        r = gw_layout.parse_lines(lines)
        self.assertTrue(any("保密提示" in w for w in r.warnings))
        self.assertEqual(r.items[0].role, "secret")

    def test_markers_warning(self):
        r = parse("关于某事的通知\n\n各单位：\n\n全年检查【待补：次数】次。")
        self.assertTrue(any("待补" in w for w in r.warnings))


class TitleBreakTest(unittest.TestCase):
    def test_breaks(self):
        cases = {
            "某能源集团有限公司人力资源部2026年工作总结": ["某能源集团有限公司人力资源部", "2026年工作总结"],
            "关于开展2026年冬季安全生产大检查的通知": ["关于开展2026年冬季安全生产大检查的通知"],
        }
        for text, expect in cases.items():
            self.assertEqual(gw_layout.break_title(text, 20), expect)

    def test_rules(self):
        for text in ("关于印发《某市2026年基层减负工作实施方案（试行）》的通知",
                     "中共某市第一医院委员会关于进一步加强2026年度党风廉政建设和反腐败工作的实施意见",
                     "关于申请增加2027年硕士研究生招生计划的请示"):
            lines = gw_layout.break_title(text, 20)
            self.assertEqual("".join(lines), text)  # 不增删文字
            widths = [gw_layout.text_width_chars(x) for x in lines]
            self.assertTrue(all(3 <= w <= 20 for w in widths), lines)
            self.assertGreaterEqual(widths[0], widths[-1], lines)  # 上行不短于下行
            for a, b in zip(lines, lines[1:]):
                self.assertFalse(a[-1:].isdigit() and b[:1] in "年月日", lines)  # “2026年”不拆开


class HelperTest(unittest.TestCase):
    def test_normalize_date(self):
        self.assertEqual(gw_layout.normalize_date("2026-01-05"), ("2026年1月5日", True))
        self.assertEqual(gw_layout.normalize_date("2026年01月05日"), ("2026年01月05日", False))

    def test_width(self):
        self.assertEqual(gw_layout.text_width_chars("2026年10月16日"), 7.0)

    def test_width_chinese_quotes_full(self):
        """中文引号按 eastAsia 字体排，占一个字位置；按半角估算会让 21 字标题不分行，Word 自动折出单字行。"""
        self.assertEqual(gw_layout.text_width_chars("“清廉”‘机关’"), 8.0)
        title = "关于开展“清廉机关”创建活动的工作安排通知"  # 21 字（含 2 个引号）
        lines = gw_layout.break_title(title, 20)
        self.assertGreater(len(lines), 1, lines)
        self.assertEqual("".join(lines), title)


LINT_DRAFT_NOTE_RE = re.compile(r"^\s*【起草说明】")  # 与 shared/scripts/gongwen_lint.py 的 DRAFT_NOTE_RE 一致


class ReadBodyTextTest(unittest.TestCase):
    def write(self, name, data: bytes):
        path = OUT / name
        path.write_bytes(data)
        return path

    def test_gbk_file_falls_back(self):
        """GBK（GB18030）编码的正文不能抛 UnicodeDecodeError，按 GB18030 读取并提示（与 gongwen_lint 一致）。"""
        path = self.write("编码-GBK.txt", "关于某事的通知\r\n\r\n各单位：\r\n\r\n现通知如下。\r\n".encode("gbk"))
        lines, notes = gw_layout.read_body_text(path)
        self.assertEqual(lines[0], "关于某事的通知")
        self.assertIn("现通知如下。", lines)
        self.assertTrue(any("GB18030" in n for n in notes), notes)

    def test_undecodable_file_raises_value_error(self):
        path = self.write("编码-UTF16.txt", "关于某事的通知".encode("utf-16"))
        with self.assertRaises(ValueError):
            gw_layout.read_body_text(path)

    def test_bom_and_crlf(self):
        path = self.write("编码-BOM.txt", "﻿关于某事的通知\r\n各单位：\r\n现通知如下。\r\n【起草说明】\r\n说明".encode("utf-8"))
        lines, _ = gw_layout.read_body_text(path)
        self.assertEqual(lines, ["关于某事的通知", "各单位：", "现通知如下。"])

    def test_draft_note_variants(self):
        """lint 截断的写法这里都要截断；Markdown 标题、加粗包裹的【起草说明】也要截断，不能排进正文。"""
        variants = ["【起草说明】", "　【起草说明】", " 【起草说明】 ", "【起草说明】：", "\t【起草说明】",
                    "## 【起草说明】", "**【起草说明】**", "### **【起草说明】**"]
        for suffix in (".md", ".txt"):
            for k, variant in enumerate(variants):
                path = self.write(f"起草说明-{k}{suffix}",
                                  ("关于某事的通知\n\n现通知如下。\n" + variant + "\n所作假设：某某\n").encode("utf-8"))
                lines, _ = gw_layout.read_body_text(path)
                self.assertNotIn("所作假设：某某", lines, (suffix, variant))
                self.assertFalse(any("起草说明" in ln for ln in lines), (suffix, variant))
                if LINT_DRAFT_NOTE_RE.match(variant):
                    self.assertTrue(gw_layout.is_draft_note_line(variant), variant)
        # 正文中提到“起草说明”的句子不截断
        self.assertFalse(gw_layout.is_draft_note_line("请在起草说明中写明【起草说明】的内容。"))
        self.assertEqual(gw_layout.draft_note_index(["标题", "正文", "【起草说明】", "说明"]), 2)
        self.assertIsNone(gw_layout.draft_note_index(["标题", "正文"]))


class DocnoUnderTitleParseTest(unittest.TestCase):
    def test_docno_line_after_title(self):
        """标题与主送机关之间的发文字号行认作发文字号（原来被当成题注排在标题下）。"""
        r = parse("关于开展冬季校园安全检查的通知\n\n某区教发[2026]第12号\n\n各中小学、幼儿园：\n\n现通知如下。")
        self.assertEqual(r.roles(), ["title", "docno", "recipient", "body"])
        self.assertEqual(r.fields["发文字号"], "某区教发[2026]第12号")
        self.assertEqual(r.fields["主送机关"], "各中小学、幼儿园")
        r2 = parse("关于某事的请示\n某发〔2026〕5号　　签发人：张某\n市政府：\n现请示如下。")
        self.assertEqual(r2.fields["发文字号"], "某发〔2026〕5号")
        self.assertEqual(r2.fields["签发人"], "张某")
        # 正文中引用文号的句子不受影响
        r3 = parse("关于某事的通知\n\n各单位：\n\n根据某发〔2026〕5号文件精神，现通知如下。")
        self.assertNotIn("发文字号", r3.fields)


class RobustParseTest(unittest.TestCase):
    def test_opening_sentence_not_recipient(self):
        """通告、公告没有主送机关，标题后“……现将有关事项通告如下：”是正文首段，不能认成主送机关。"""
        r = parse("某市人民政府关于某事的通告\n\n为加强管理，现将有关事项通告如下：\n\n一、自即日起执行。\n\n"
                  "某市人民政府\n\n2026年3月5日")
        self.assertNotIn("主送机关", r.fields)
        self.assertEqual(r.items[1].role, "body")
        r2 = parse("关于某事的通知\n\n各省、自治区、直辖市人民政府，国务院各部委：\n\n现通知如下。")
        self.assertEqual(r2.fields["主送机关"], "各省、自治区、直辖市人民政府，国务院各部委")

    def test_print_line_regex_is_linear(self):
        """末行含超长空白时，印发行正则不能出现二次方回溯（原写法 2 万个空格约 5 秒）。"""
        for line in ("某" + "　" * 50000 + "x", "a" + " " * 50000 + "x", "某局 " * 20000 + "印发"):
            start = time.perf_counter()
            gw_layout.parse_lines(["关于某事的通知", "正文。", line])
            self.assertLess(time.perf_counter() - start, 1.0, line[:10])
        r = parse("关于某事的通知\n\n各单位：\n\n现通知如下。\n\n某局\n\n2026年3月5日\n\n某局 办公室　　2026年3月5日印发")
        self.assertEqual(r.fields["印发机关"], "某局 办公室")
        self.assertEqual(r.fields["印发日期"], "2026年3月5日")


if __name__ == "__main__":
    unittest.main()
