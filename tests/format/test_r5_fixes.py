# -*- coding: utf-8 -*-
"""新留出题评测发现的排版缺陷（第五轮）的回归测试：
1. 排版不补原稿没有的内容（不补版记、不插【待补】），缺的要素只在报告中列出；
2. 文字不重复，--keep-text 双向、按顺序核对输入全文与文档全文；
3. fix 和重排清除原页眉、页脚中的旧页码；
4. 非标准序号（1.、1.1、（1））的层次标题按层级套用标题字体；
5. 落款、成文日期与正文最后一段同页，附件说明不被分页割裂；
6. 信函双线与版心等宽；
7. 讲话稿称呼行顶格；
8. 报告如实：文字核对结论、版面核对事实。"""
import io
import shutil
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from tests.format.helpers import NO_FONTS, OUT, build_text, check_docx, gw_docx, gw_layout

import build_docx  # noqa: E402
import extract_docx  # noqa: E402
import fix_docx  # noqa: E402
import gw_verify  # noqa: E402
import preview_docx  # noqa: E402

REQUEST = ("关于申请更新实验室消防设施的请示\n\n市应急管理局：\n\n我院实验楼消防设施已使用多年，需要更新。\n\n"
           "妥否，请批示。\n\n某职业技术学院\n\n2026年4月8日\n\n（联系人：周某，联系电话：0000-55551234）")
RED_UP = {"版式": "红头", "发文机关标志": "某职业技术学院文件", "发文字号": "某职院发〔2026〕12号", "签发人": "钱某"}


def texts(doc):
    return [p.text for p in doc.paragraphs]


def ids(result):
    return {(i["id"], i["level"]) for i in result["issues"]}


def add_page_field(paragraph, prefix="第 ", suffix=" 页"):
    """在段落里写“第 {PAGE} 页”（模拟原稿把页码放在页眉）。"""
    paragraph.add_run(prefix)
    for kind, text in (("begin", None), ("instr", " PAGE "), ("separate", None), ("text", "1"), ("end", None)):
        run = paragraph.add_run()
        if kind == "instr":
            el = OxmlElement("w:instrText")
            el.text = text
            run._r.append(el)
        elif kind == "text":
            run.text = text
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
            run._r.append(el)
    paragraph.add_run(suffix)


def build_file(text, name, meta=None, **kw):
    doc, report = build_docx.build(text.split("\n"), meta or {}, installed=NO_FONTS, **kw)
    path = OUT / name
    if path.exists():
        path.unlink()
    doc.save(str(path))
    return path, report


# ---------------------------------------------------------------------------
# 1. 排版不补内容
# ---------------------------------------------------------------------------

class NoInventedContentTest(unittest.TestCase):
    def test_upward_red_without_banji_info_adds_nothing(self):
        """上行文（请示）红头：原稿和版式信息没有印发机关、印发日期时，不补版记、不插【待补】。"""
        path, report = build_file(REQUEST, "r5-请示-不补版记.docx", RED_UP)
        doc = Document(str(path))
        joined = "\n".join(texts(doc))
        self.assertNotIn("待补", joined)
        self.assertNotIn("印发", joined)
        self.assertEqual(len(doc.tables), 0)  # 没有版记表格
        self.assertEqual(report["placeholders"], [])
        self.assertTrue(any("印发机关" in m for m in report["missing"]), report["missing"])
        self.assertTrue(any("印发日期" in m for m in report["missing"]), report["missing"])

    def test_missing_head_and_signoff_not_filled(self):
        """红头缺标志、字号、署名、日期时一律不写，只在“缺失要素”列出；留占位模式才写【待补】。"""
        text = "关于某事的请示\n\n县政府：\n\n现请示如下。"
        path, report = build_file(text, "r5-缺要素.docx", {"版式": "红头"})
        doc = Document(str(path))
        self.assertNotIn("待补", "".join(texts(doc)))
        names = "".join(report["missing"])
        for key in ("发文机关标志", "发文字号", "发文机关署名", "成文日期", "印发机关"):
            self.assertIn(key, names)
        self.assertIn("缺失要素", build_docx.render_text(dict(report, output="x.docx", layout="红头", summary=[])))
        # 用户明确要求留位置时才写占位
        path2, report2 = build_file(text, "r5-缺要素-留占位.docx", {"版式": "红头"}, placeholders=True)
        self.assertIn("【待补：发文机关标志】", "".join(texts(Document(str(path2)))))
        self.assertEqual(report2["missing"], [])
        self.assertTrue(report2["placeholders"])

    def test_cli_placeholders_flag(self):
        src = OUT / "r5-正文.txt"
        src.write_text("关于某事的请示\n\n县政府：\n\n现请示如下。\n", encoding="utf-8")
        for flag, expect in ((None, False), ("--placeholders", True)):
            out = OUT / f"r5-cli-{'p' if flag else 'n'}.docx"
            if out.exists():
                out.unlink()
            args = [str(src), "-o", str(out), "--layout", "红头", "--no-check"] + ([flag] if flag else [])
            with redirect_stdout(io.StringIO()):
                self.assertEqual(build_docx.main(args), 0)
            self.assertEqual("待补" in "".join(texts(Document(str(out)))), expect)

    def test_attachment_note_not_generated(self):
        """正文没有“附件：”说明时不按附件页标题补写，只列入缺失要素。"""
        text = REQUEST + "\n\n【附件1】设施清单\n\n一、清单\n\n灭火器 20 具。"
        path, report = build_file(text, "r5-附件说明不补.docx", RED_UP)
        self.assertNotIn("附件：", "".join(texts(Document(str(path)))))
        self.assertTrue(any("附件说明" in m for m in report["missing"]), report["missing"])

    def test_rebuild_after_extract_adds_nothing(self):
        """已有 docx 抽取—重排：原稿没有的要素不补（版记、发文机关标志等）。"""
        src = OUT / "r5-重排-原稿.docx"
        doc = Document()
        for line in ("关于某事的请示", "县政府：", "现请示如下。", "某局", "2026年4月8日"):
            doc.add_paragraph(line)
        doc.save(str(src))
        res = extract_docx.extract(src, "红头")
        built, report = build_docx.build(res["lines"], res["meta"], keep_text=True, installed=NO_FONTS)
        self.assertNotIn("待补", "".join(texts(built)))
        self.assertEqual(len(built.tables), 0)
        self.assertTrue(any("发文机关标志" in m for m in report["missing"]))
        self.assertTrue(report["text_check"]["ok"])


# ---------------------------------------------------------------------------
# 2. 文字不重复；--keep-text 双向、按顺序核对
# ---------------------------------------------------------------------------

class TextCheckTest(unittest.TestCase):
    def test_compare_text_detects_all_kinds(self):
        inp = ["标题", "第一段文字。", "第二段文字。", "某局", "2026年4月8日"]
        self.assertTrue(gw_verify.compare_text(inp, inp).ok)
        # 少字
        res = gw_verify.compare_text(["标题", "第一段文字。", "某局", "2026年4月8日"], inp)
        self.assertFalse(res.ok)
        self.assertTrue(any(p.startswith("少字") for p in res.problems), res.problems)
        # 多字（含重复的落款日期）
        res = gw_verify.compare_text(inp + ["2026年4月8日"], inp)
        self.assertTrue(any(p.startswith("重复") for p in res.problems), res.problems)
        res = gw_verify.compare_text(inp + ["【待补：印发机关】"], inp)
        self.assertTrue(any(p.startswith("多字") for p in res.problems), res.problems)
        # 乱序
        res = gw_verify.compare_text(["标题", "第二段文字。", "第一段文字。", "某局", "2026年4月8日"], inp)
        self.assertTrue(any(p.startswith("乱序") for p in res.problems), res.problems)
        # 空白差异不算
        self.assertTrue(gw_verify.compare_text(["标 题", "第一段　文字。", "第二段文字。", "某局", "2026年4月8日"], inp).ok)
        # 版式信息写出的文字允许；版头要素可以换位置
        meta_texts = ["某局文件", "某发〔2026〕1号"]
        self.assertTrue(gw_verify.compare_text(meta_texts + inp, inp, allowed=meta_texts).ok)
        moved = gw_verify.compare_text(["某局"] + inp[:3] + inp[4:], inp, movable=["某局"])
        self.assertTrue(moved.ok, moved.problems)

    def test_keep_text_aborts_on_duplicate_date(self):
        """重排后文末多出一行重复的落款日期：--keep-text 必须中止，不得保存。"""
        real = build_docx.Builder.para

        def twice(self, role, text, *a, **kw):
            p = real(self, role, text, *a, **kw)
            if role == "date":
                real(self, role, text)
            return p
        with mock.patch.object(build_docx.Builder, "para", twice):
            with self.assertRaises(build_docx.BuildError) as cm:
                build_file(REQUEST, "r5-重复日期.docx", RED_UP, keep_text=True)
        self.assertIn("重复", str(cm.exception))
        self.assertIn("2026年4月8日", str(cm.exception))

    def test_keep_text_aborts_on_extra_and_missing(self):
        real = build_docx.Builder.para

        def extra(self, role, text, *a, **kw):
            p = real(self, role, text, *a, **kw)
            if role == "signature":
                real(self, "note", "【待补：联系人】")
            return p

        def lose(self, role, text, *a, **kw):
            return real(self, role, "" if text == "妥否，请批示。" else text, *a, **kw)
        with mock.patch.object(build_docx.Builder, "para", extra):
            with self.assertRaises(build_docx.BuildError) as cm:
                build_file(REQUEST, "r5-多字.docx", RED_UP, keep_text=True)
        self.assertIn("多字", str(cm.exception))
        with mock.patch.object(build_docx.Builder, "para", lose):
            with self.assertRaises(build_docx.BuildError) as cm:
                build_file(REQUEST, "r5-少字.docx", RED_UP, keep_text=True)
        self.assertIn("少字", str(cm.exception))

    def test_keep_text_passes_and_reports_verified(self):
        path, report = build_file(REQUEST, "r5-核对通过.docx", dict(RED_UP, 印发机关="某学院办公室", 印发日期="2026年4月9日"),
                                  keep_text=True)
        tc = report["text_check"]
        self.assertTrue(tc["ok"] and tc["keep_text"], tc)
        text = build_docx.render_text(dict(report, output=str(path), summary=[]))
        self.assertIn("逐字一致", text)

    def test_normal_mode_does_not_claim_verified(self):
        """未加 --keep-text 时，报告不声称逐字一致；有差异就列出。"""
        _, report = build_file("关于某事的通知\n\n各单位:\n\n现通知如下。", "r5-普通模式.docx", {"版式": "白头"},
                               auto_fix_text=True)
        text = build_docx.render_text(dict(report, output="x", layout="白头", summary=[]))
        self.assertNotIn("已双向、按顺序规范化比对", text)
        self.assertIn("不得声称逐字一致", text)  # 补了全角冒号，如实写出有差异
        self.assertFalse(report["text_check"]["keep_text"])

    def test_same_line_in_body_and_meta_rendered_once(self):
        """版式信息给了成文日期、署名，原稿同一行又留在正文里（如日期后还跟着联系人行）：只排一次。"""
        text = ("关于某事的请示\n\n县政府：\n\n现请示如下。\n\n某局\n\n2026年4月8日\n\n联系人：周某，电话：0000-5555")
        path, report = build_file(text, "r5-正文与版式信息重复.docx",
                                  {"版式": "白头", "发文机关署名": "某局", "成文日期": "2026年4月8日"})
        got = texts(Document(str(path)))
        self.assertEqual(got.count("2026年4月8日"), 1, got)
        self.assertEqual(got.count("某局"), 1, got)
        # 联系人行（没加括号）按附注识别，排在日期下一行，不再被当成正文
        self.assertEqual(got[-1], "联系人：周某，电话：0000-5555")
        self.assertTrue(report["text_check"]["ok"] or not report["text_check"]["keep_text"])
        # 版式信息同样给了日期且该行留在正文中间：只排一次，并说明
        text2 = text.replace("联系人：周某，电话：0000-5555", "特此说明。")
        path2, report2 = build_file(text2, "r5-正文与版式信息重复2.docx",
                                    {"版式": "白头", "发文机关署名": "某局", "成文日期": "2026年4月8日"}, keep_text=True)
        got2 = texts(Document(str(path2)))
        self.assertEqual(got2.count("2026年4月8日"), 1, got2)
        self.assertTrue(report2["text_check"]["ok"], report2["text_check"])

    def test_blank_day_date_recognized(self):
        """成文日期日留空（“2026年10月　日”）仍识别为成文日期并排在署名之下，回读检查无 P0、P1。"""
        text = "关于某事的通知\n\n各单位:\n\n现通知如下。\n\n特此通知。\n\n某局\n2026年10月　日"
        path, report = build_file(text, "r5-日留空.docx", {"版式": "白头"}, auto_fix_text=True)
        got = texts(Document(str(path)))
        self.assertEqual(got[-2:], ["某局", "2026年10月　日"], got)
        bad = {x for x in ids(check_docx.check_document(path)) if x[1] in ("P0", "P1")}
        self.assertFalse(bad, bad)

    def test_check_reports_duplicate_date_line(self):
        """回读检查：同一行成文日期在正文中出现两次（如重排时多排了一行）报 P1。"""
        src = OUT / "r5-重复日期-检查.docx"
        doc = Document()
        for line in ("关于某事的通知", "各单位：", "现通知如下。", "某局", "2026年4月8日", "2026年4月8日"):
            doc.add_paragraph(line)
        doc.save(str(src))
        self.assertIn(("DOCX-092", "P1"), ids(check_docx.check_document(src)))
        ok = OUT / "r5-重复日期-正常.docx"
        doc = Document()
        for line in ("关于某事的通知", "各单位：", "现通知如下。", "某局", "2026年4月8日"):
            doc.add_paragraph(line)
        doc.save(str(ok))
        self.assertNotIn("DOCX-092", {i["id"] for i in check_docx.check_document(ok)["issues"]})

    def test_extract_checks_against_original(self):
        src = OUT / "r5-抽取核对.docx"
        doc = Document()
        for line in ("某学院文件", "某院发〔2026〕3号", "关于某事的通知", "各单位：", "现通知如下。", "某学院", "2026年4月8日"):
            doc.add_paragraph(line)
        doc.save(str(src))
        res = extract_docx.extract(src)
        self.assertTrue(res["text_check"]["ok"], res["text_check"])
        # 版式信息里把日期又写一遍（常见的手工补写），抽取核对之外由 build 的核对兜底
        meta = dict(res["meta"], 成文日期="2026年4月8日")
        built, report = build_docx.build(res["lines"], meta, keep_text=True, installed=NO_FONTS)
        self.assertEqual(texts(built).count("2026年4月8日"), 1)


# ---------------------------------------------------------------------------
# 3. 页眉残留旧页码
# ---------------------------------------------------------------------------

def messy_with_header(path: Path, *, letter: bool = False) -> Path:
    doc = Document()
    sec = doc.sections[0]
    add_page_field(sec.header.paragraphs[0])
    sec.footer.paragraphs[0].text = "某单位内部材料"
    for line in ("关于某事的通知", "各单位：", "现通知如下。", "某局", "2026年4月8日"):
        doc.add_paragraph(line)
    doc.save(str(path))
    return path


class HeaderPageNumberTest(unittest.TestCase):
    def setUp(self):
        self.src = messy_with_header(OUT / "r5-页眉页码.docx")

    def test_check_reports_header_page_number(self):
        result = check_docx.check_document(self.src)
        self.assertIn(("DOCX-075", "P1"), ids(result))

    def test_fix_clears_header_and_resets_footer(self):
        dst = OUT / "r5-页眉页码-修复.docx"
        if dst.exists():
            dst.unlink()
        report = fix_docx.fix_document(self.src, dst, installed=NO_FONTS, keep_text=True)
        doc = Document(str(dst))
        sec = doc.sections[0]
        self.assertEqual(gw_docx.part_summary(sec.header._element), "")
        self.assertEqual(gw_docx.part_summary(sec.even_page_header._element), "")
        footer = gw_docx.part_summary(sec.footer._element)
        self.assertIn("页码域", footer)
        self.assertNotIn("内部材料", footer)
        self.assertTrue(any("已清除原页眉、页脚" in n and "页码域" in n for n in report["notes"]), report["notes"])
        result = check_docx.check_document(dst)
        self.assertNotIn("DOCX-075", {i["id"] for i in result["issues"]})
        self.assertEqual(result["summary"]["P0"], 0, result["issues"])
        # 正文文字不变
        self.assertEqual(texts(doc), texts(Document(str(self.src))))

    def test_first_page_header_cleared_for_non_letter(self):
        doc = Document(str(self.src))
        sec = doc.sections[0]
        sec.different_first_page_header_footer = True
        add_page_field(sec.first_page_header.paragraphs[0])
        sec.first_page_footer.paragraphs[0].text = "1"
        src = OUT / "r5-首页页眉页码.docx"
        doc.save(str(src))
        dst = OUT / "r5-首页页眉页码-修复.docx"
        if dst.exists():
            dst.unlink()
        fix_docx.fix_document(src, dst, installed=NO_FONTS)
        sec = Document(str(dst)).sections[0]
        self.assertFalse(sec.different_first_page_header_footer)
        self.assertEqual(gw_docx.part_summary(sec.first_page_header._element), "")
        self.assertEqual(gw_docx.part_summary(sec.first_page_footer._element), "")

    def test_letter_keeps_first_page_issuer_and_rules(self):
        path, _ = build_file("关于商请某事的函\n\n县局：\n\n请予研究函复。\n\n某院\n\n2026年4月8日", "r5-信函-页眉.docx",
                             {"版式": "信函", "发文机关标志": "某院", "发文字号": "某院函〔2026〕1号"})
        doc = Document(str(path))
        add_page_field(doc.sections[0].header.paragraphs[0])  # 非首页页眉里的旧页码
        doc.sections[0].first_page_footer.paragraphs[0].add_run("1")  # 首页页脚里的页码
        src = OUT / "r5-信函-页眉2.docx"
        doc.save(str(src))
        dst = OUT / "r5-信函-页眉2-修复.docx"
        if dst.exists():
            dst.unlink()
        fix_docx.fix_document(src, dst, installed=NO_FONTS)
        sec = Document(str(dst)).sections[0]
        self.assertEqual(gw_docx.part_summary(sec.header._element), "")
        self.assertEqual(gw_docx.part_summary(sec.first_page_header._element), "某院")  # 发文机关标志保留
        self.assertEqual(gw_docx.part_summary(sec.first_page_footer._element), "")
        self.assertEqual(len(sec.first_page_footer._element.findall(qn("w:p"))), 2)  # 红色双线（两个带边框的段落）保留
        self.assertNotIn("DOCX-075", {i["id"] for i in check_docx.check_document(dst, layout="信函")["issues"]})

    def test_extract_reports_dropped_header_footer(self):
        res = extract_docx.extract(self.src)
        note = "".join(res["notes"])
        self.assertIn("原页眉、页脚中的内容不带入重排稿", note)
        self.assertIn("页码域", note)
        built, _ = build_docx.build(res["lines"], res["meta"], keep_text=True, installed=NO_FONTS)
        out = OUT / "r5-重排-页眉.docx"
        built.save(str(out))
        self.assertEqual(gw_docx.part_summary(Document(str(out)).sections[0].header._element), "")


# ---------------------------------------------------------------------------
# 4. 非标准序号的层次标题按层级套用字体
# ---------------------------------------------------------------------------

NUMERIC_TEXT = ("关于加强安全生产工作的通知\n\n各单位：\n\n现将有关事项通知如下。\n\n1. 总体要求\n\n1.1 指导思想\n\n坚持安全第一、预防为主。\n\n"
                "1.2 基本原则\n\n（1）坚持预防为主。\n\n（2）坚持综合治理。\n\n2. 主要任务\n\n2.1 压实责任\n\n逐级落实安全生产责任制。\n\n"
                "3.5万元用于设备更新，另有说明。\n\n某局\n\n2026年4月8日")


def east_asia_font(run):
    rf = run._r.rPr.find(qn("w:rFonts")) if run._r.rPr is not None else None
    return rf.get(qn("w:eastAsia")) if rf is not None else None


class HeadingLevelTest(unittest.TestCase):
    def test_parse_nests_by_depth(self):
        parsed = gw_layout.parse_lines(NUMERIC_TEXT.split("\n"))
        roles = {it.text: it.role for it in parsed.items}
        self.assertEqual(roles["1. 总体要求"], "h1")
        self.assertEqual(roles["1.1 指导思想"], "h2")
        self.assertEqual(roles["（1）坚持预防为主。"], "h3")
        self.assertEqual(roles["2. 主要任务"], "h1")
        self.assertEqual(roles["3.5万元用于设备更新，另有说明。"], "body")  # 小数开头的正文不是标题
        self.assertTrue(any("序号不是国标写法" in w for w in parsed.warnings), parsed.warnings)

    def test_standard_numbering_unchanged(self):
        parsed = gw_layout.parse_lines("关于某事的通知\n各单位：\n一、总体要求\n（一）基本原则\n1. 坚持预防\n（1）细则".split("\n"))
        self.assertEqual([it.role for it in parsed.items][-4:], ["h1", "h2", "h3", "h4"])
        self.assertFalse(any(it.nonstd for it in parsed.items))

    def test_build_fonts_by_level(self):
        path, report = build_file(NUMERIC_TEXT, "r5-数字序号-排版.docx", {"版式": "白头", "发文机关署名": "某局"})
        doc = Document(str(path))
        want = {"1. 总体要求": "黑体", "1.1 指导思想": "楷体_GB2312", "（1）坚持预防为主。": "仿宋_GB2312"}
        for text, font in want.items():
            p = [q for q in doc.paragraphs if q.text == text][0]
            self.assertEqual(east_asia_font(p.runs[0]), font, text)

    def test_fix_keep_text_only_changes_fonts(self):
        src = OUT / "r5-数字序号.docx"
        doc = Document()
        for line in NUMERIC_TEXT.split("\n"):
            if line:
                doc.add_paragraph(line)
        doc.save(str(src))
        dst = OUT / "r5-数字序号-修复.docx"
        if dst.exists():
            dst.unlink()
        report = fix_docx.fix_document(src, dst, installed=NO_FONTS, keep_text=True)
        fixed = Document(str(dst))
        self.assertEqual(texts(fixed), texts(Document(str(src))))  # 序号文字不改
        want = {"1. 总体要求": "黑体", "1.1 指导思想": "楷体_GB2312", "（1）坚持预防为主。": "仿宋_GB2312", "2. 主要任务": "黑体"}
        for text, font in want.items():
            p = [q for q in fixed.paragraphs if q.text == text][0]
            self.assertEqual(east_asia_font(p.runs[0]), font, text)
        self.assertTrue(any("序号不是国标写法" in w for w in report["warnings"]), report["warnings"])
        # 报告“文字类问题”里建议改成标准序号
        result = check_docx.check_document(dst)
        msgs = [i for i in result["issues"] if i["id"] == "DOCX-033"]
        self.assertTrue(any("建议改为“一、”" in i["message"] for i in msgs), msgs)
        self.assertTrue(any("建议改为“（一）”" in i["message"] for i in msgs), msgs)


# ---------------------------------------------------------------------------
# 5. 落款与正文末段同页，附件说明不被分页割裂
# ---------------------------------------------------------------------------

def keep_next(p) -> bool:
    ppr = p._p.pPr
    return ppr is not None and ppr.find(qn("w:keepNext")) is not None


ATTACH_NOTICE = ("关于某事的函\n\n县局：\n\n请予研究函复。\n\n附件：1. 试点方案\n2. 清单\n\n某市第一医院\n\n2026年10月22日")


class SignoffChainTest(unittest.TestCase):
    def _chain_ok(self, doc):
        paras = doc.paragraphs
        k = [i for i, p in enumerate(paras) if p.text == "请予研究函复。"][0]
        j = [i for i, p in enumerate(paras) if p.text == "2026年10月22日"][0]
        for i in range(k, j):  # 正文末段、空行、附件说明、空行、署名都与下段同页
            self.assertTrue(keep_next(paras[i]), (i, paras[i].text))

    def test_build_chain_covers_attachment_note(self):
        path, _ = build_file(ATTACH_NOTICE, "r5-落款链-build.docx", {"版式": "白头"})
        self._chain_ok(Document(str(path)))

    def test_fix_chain_covers_attachment_note(self):
        path, _ = build_file(ATTACH_NOTICE, "r5-落款链-原.docx", {"版式": "白头"})
        doc = Document(str(path))
        for p in doc.paragraphs:  # 抹掉分页控制，模拟原稿
            ppr = p._p.pPr
            if ppr is not None and ppr.find(qn("w:keepNext")) is not None:
                ppr.remove(ppr.find(qn("w:keepNext")))
        src = OUT / "r5-落款链-无控制.docx"
        doc.save(str(src))
        dst = OUT / "r5-落款链-修复.docx"
        if dst.exists():
            dst.unlink()
        fix_docx.fix_document(src, dst, installed=NO_FONTS)
        self._chain_ok(Document(str(dst)))

    @unittest.skipUnless(preview_docx.find_soffice(), "需要 LibreOffice")
    def test_render_signoff_not_orphaned(self):
        """用 LibreOffice 实际渲染：落款不单独成页（字体缺失时用替代字体渲染，只核对分页结构）。"""
        try:
            import pymupdf  # noqa: F401
        except ImportError:
            try:
                import fitz  # noqa: F401
            except ImportError:
                self.skipTest("需要 pymupdf")
        body = "\n\n".join("为方便参保患者就近办理门诊慢特病认定，我院拟开展联合审核试点，由专科医师初审、医保部门复核。第%d段。" % i
                           for i in range(1, 11))
        text = "关于某事的函\n\n县局：\n\n" + body + "\n\n请予研究函复。\n\n附件：1. 试点方案\n2. 清单\n\n某市第一医院\n\n2026年10月22日"
        path, _ = build_file(text, "r5-落款链-渲染.docx", {"版式": "白头"})
        out_dir = OUT / "r5-渲染"
        if out_dir.exists():
            shutil.rmtree(out_dir)
        pdf = preview_docx.to_pdf(preview_docx.find_soffice(), path, out_dir)
        facts = preview_docx.analyze_pages(pdf)
        self.assertTrue(any("落款与正文最后一段在同一页" in f for f in facts), facts)
        self.assertFalse(any("不在同一页" in f for f in facts), facts)


class AnalyzePagesTest(unittest.TestCase):
    """版面核对只写核对到的事实：用最小 PDF 验证孤页识别。"""

    def _pdf(self, pages, name):
        try:
            import pymupdf
        except ImportError:
            try:
                import fitz as pymupdf
            except ImportError:
                self.skipTest("需要 pymupdf")
        doc = pymupdf.open()
        for lines in pages:
            page = doc.new_page(width=595, height=842)
            y = 100
            for ln in lines:
                page.insert_text((80, y), ln, fontname="china-s")
                y += 30
        path = OUT / name
        doc.save(str(path))
        doc.close()
        return path

    def test_orphan_signoff_flagged(self):
        pdf = self._pdf([["这是正文第一页内容需要很长一段文字，用来占版面。"], ["某市第一医院", "2026年10月22日"]], "r5-孤页.pdf")
        facts = preview_docx.analyze_pages(pdf)
        self.assertTrue(any("不在同一页" in f for f in facts), facts)

    def test_signoff_with_body_ok(self):
        pdf = self._pdf([["第一页的正文内容，占一些版面。"], ["请予研究函复，这一行是末页正文。", "某市第一医院", "2026年10月22日"]],
                        "r5-同页.pdf")
        facts = preview_docx.analyze_pages(pdf)
        self.assertTrue(any("在同一页" in f and "不在" not in f for f in facts), facts)


# ---------------------------------------------------------------------------
# 6. 信函双线与版心等宽
# ---------------------------------------------------------------------------

class LetterRuleWidthTest(unittest.TestCase):
    def test_double_rules_170mm_centered(self):
        path, _ = build_file("关于商请某事的函\n\n县局：\n\n请予研究函复。\n\n某院\n\n2026年4月8日", "r5-信函-线长.docx",
                             {"版式": "信函", "发文机关标志": "某院", "发文字号": "某院函〔2026〕1号"})
        sec = Document(str(path)).sections[0]
        parts = list(sec.first_page_header._element.findall(qn("w:p"))) + list(sec.first_page_footer._element.findall(qn("w:p")))
        ruled = [p for p in parts if p.find(qn("w:pPr") + "/" + qn("w:pBdr")) is not None]
        self.assertEqual(len(ruled), 4)  # 页眉两条、页脚两条
        for p in ruled:
            ind = p.find(qn("w:pPr") + "/" + qn("w:ind"))
            # 10.1：线长 170 mm、居中，比版心（156 mm）两侧各宽 7 mm（397 twips）
            self.assertAlmostEqual(int(ind.get(qn("w:left")) or 0), -397, delta=3)
            self.assertAlmostEqual(int(ind.get(qn("w:right")) or 0), -397, delta=3)
        width_mm = gw_docx.twips_to_mm(int(sec._sectPr.find(qn("w:pgSz")).get(qn("w:w")))
                                       - int(sec._sectPr.find(qn("w:pgMar")).get(qn("w:left")))
                                       - int(sec._sectPr.find(qn("w:pgMar")).get(qn("w:right"))))
        ind = ruled[0].find(qn("w:pPr") + "/" + qn("w:ind"))
        self.assertAlmostEqual(width_mm - 2 * gw_docx.twips_to_mm(int(ind.get(qn("w:left")))), 170, delta=1)  # 10.1 线长 170 mm


# ---------------------------------------------------------------------------
# 7. 讲话稿称呼顶格
# ---------------------------------------------------------------------------

SPEECH = ("在全市安全生产推进会上的讲话\n\n市委常委、副市长　张某某\n\n（2026年10月21日）\n\n同志们：\n\n今天开会，主要任务是安排部署。\n\n"
          "一、看清形势\n\n前三季度形势总体平稳。")


class SalutationTest(unittest.TestCase):
    def test_parse_recognizes_salutation(self):
        parsed = gw_layout.parse_lines(SPEECH.split("\n"))
        roles = {it.text: it.role for it in parsed.items}
        self.assertEqual(roles["同志们："], "recipient")
        self.assertEqual(roles["市委常委、副市长　张某某"], "title_note")
        # 称呼不紧跟标题（中间隔着正文）时也顶格
        parsed = gw_layout.parse_lines("某讲话稿\n会议背景说明\n。\n同志们：\n今天开会。".split("\n"))
        self.assertIn("salutation", parsed.roles())

    def test_build_flush_left(self):
        path, _ = build_file(SPEECH, "r5-讲话稿.docx")
        doc = Document(str(path))
        p = [q for q in doc.paragraphs if q.text == "同志们："][0]
        self.assertEqual(gw_docx.leading_offset_chars(p, doc), 0)
        body = [q for q in doc.paragraphs if q.text.startswith("今天开会")][0]
        self.assertGreater(gw_docx.leading_offset_chars(body, doc), 1.5)

    def test_keep_text_space_indented_salutation_stays_flush(self):
        """原稿用空格缩进称呼：--keep-text 不删空格，改用等宽负缩进抵消，视觉上仍顶格；回读不报缩进。"""
        text = SPEECH.replace("同志们：", "　　同志们：")
        path, report = build_file(text, "r5-讲话稿-空格.docx", keep_text=True)
        doc = Document(str(path))
        p = [q for q in doc.paragraphs if "同志们：" in q.text][0]
        self.assertTrue(p.text.startswith("　　"))  # 文字不改
        self.assertEqual(gw_docx.leading_offset_chars(p, doc), 0)
        self.assertNotIn("DOCX-040", {i["id"] for i in check_docx.check_document(path)["issues"]})
        # fix --keep-text 同样
        src = OUT / "r5-讲话稿-空格-原.docx"
        d = Document()
        for line in text.split("\n"):
            if line:
                d.add_paragraph(line)
        d.save(str(src))
        dst = OUT / "r5-讲话稿-空格-修复.docx"
        if dst.exists():
            dst.unlink()
        fix_docx.fix_document(src, dst, installed=NO_FONTS, keep_text=True)
        fixed = Document(str(dst))
        q = [x for x in fixed.paragraphs if "同志们：" in x.text][0]
        self.assertTrue(q.text.startswith("　　"))
        self.assertEqual(gw_docx.leading_offset_chars(q, fixed), 0)

    def test_fix_default_strips_salutation_indent(self):
        src = OUT / "r5-讲话稿-缩进.docx"
        d = Document()
        for line in SPEECH.replace("同志们：", "　　同志们：").split("\n"):
            if line:
                d.add_paragraph(line)
        d.save(str(src))
        dst = OUT / "r5-讲话稿-缩进-修复.docx"
        if dst.exists():
            dst.unlink()
        fix_docx.fix_document(src, dst, installed=NO_FONTS)
        fixed = Document(str(dst))
        q = [x for x in fixed.paragraphs if "同志们：" in x.text][0]
        self.assertEqual(q.text, "同志们：")
        self.assertEqual(gw_docx.leading_offset_chars(q, fixed), 0)


if __name__ == "__main__":
    unittest.main()
