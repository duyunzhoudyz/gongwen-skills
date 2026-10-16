#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把公文正文（.txt/.md）排成符合 GB/T 9704-2012 的 docx。

版式：红头（文件格式）、信函、纪要、白头（无版头）、材料（只有标题和正文）、简报（报头、各篇、报尾）。
未指定版式时，法定公文用白头，常用材料用材料版式；版式信息或正文有简报报头的用简报版式。
生成后自动回读检查（check_docx.py），报告识别结果、字体情况和检查结论。

退出码：0 成功；2 输入错误或缺少依赖。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parent))

from gw_layout import (BRIEF_META_KEYS, UPWARD_TYPES, Parsed, break_title, doc_kind,  # noqa: E402
                       extract_pipe_tables, has_brief_header, norm_issue, norm_total, normalize_date, parse_brief,
                       parse_lines, read_body_text, summarize, text_width_chars)

LAYOUTS = {"红头": "红头", "文件": "红头", "文件格式": "红头", "红头文件": "红头",
           "信函": "信函", "信函格式": "信函", "函件": "信函",
           "纪要": "纪要", "纪要格式": "纪要",
           "白头": "白头", "白头文件": "白头",
           "材料": "材料", "普通材料": "材料",
           "简报": "简报", "工作简报": "简报", "简报格式": "简报"}
META_KEYS = ("版式", "份号", "密级", "保密期限", "紧急程度", "发文机关标志", "发文字号", "签发人", "标题",
             "主送机关", "附件", "发文机关署名", "成文日期", "印章", "附注", "抄送", "印发机关", "印发日期",
             "三级标题加粗", "表格字体", "表格字号") + BRIEF_META_KEYS
BRIEF_IGNORED = ("发文机关标志", "发文字号", "签发人", "主送机关", "附件", "发文机关署名", "成文日期", "印章", "附注",
                 "印发机关", "份号", "密级", "保密期限", "紧急程度")
TABLE_FONTS = ("仿宋", "宋体")
TABLE_SIZES = ("四号", "小四")
HEAD_ROLES = ("copy_no", "secret", "urgency", "issuer", "docno", "signer")
TITLE_ROLES = ("title", "subtitle", "title_note")
TAIL_ROLES = ("attachment", "signature", "date", "note", "cc", "print", "banji_main")
LEGAL_LAYOUTS = ("红头", "白头", "信函")


def _as_list(value) -> List[str]:
    if value is None or value == "":
        return []
    if isinstance(value, (list, tuple)):
        return [str(v).strip() for v in value if str(v).strip()]
    return [s.strip() for s in re.split(r"[\n]", str(value)) if s.strip()]


def _truthy(value) -> Optional[bool]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return value
    return str(value).strip() in ("是", "加盖", "盖章", "true", "True", "yes", "1")


def _same(a: str, b: str) -> bool:
    return re.sub(r"[\s　]", "", a or "") == re.sub(r"[\s　]", "", b or "")


class BuildError(Exception):
    pass


class Builder:
    """按版式逐段生成文档。"""

    def __init__(self, ctx, layout: str):
        import gw_docx as D
        from docx import Document
        self.D = D
        self.ctx = ctx
        self.layout = layout
        self.doc = Document()
        D.setup_document(self.doc, ctx)
        self.section = self.doc.sections[0]
        D.setup_section(self.section, ctx, letter=(layout == "信函"))

    # 基本段落 ------------------------------------------------------------
    def blank(self, n: int = 1) -> None:
        for _ in range(n):
            p = self.doc.add_paragraph()
            self.D.set_para(p, line_exact=self.ctx.line_pt, before_pt=0, after_pt=0)

    def para(self, role: str, text: str, split: int = -1):
        from gw_format import apply_role
        p = self.doc.add_paragraph()
        p.add_run(text)
        apply_role(p, role, self.ctx, split=split)
        return p

    def title(self, lines: List[str]):
        from gw_format import apply_role
        p = self.doc.add_paragraph()
        for i, line in enumerate(lines):
            run = p.add_run(line)
            if i < len(lines) - 1:
                self.D.add_break(run)
        apply_role(p, "title", self.ctx)
        return p

    # 版头 ----------------------------------------------------------------
    def issuer_size(self, text: str) -> float:
        el = self.ctx.el("issuer")
        width_pt = self.ctx.text_width_twips / 20
        chars = max(1.0, text_width_chars(text))
        size = min(float(el["size_max_pt"]), width_pt * float(el["fill_ratio"]) / chars)
        size = max(float(el["size_min_pt"]), size)
        return int(size * 2) / 2

    def issuer(self, text: str, top_from_body_mm: float, used_lines: int):
        D, ctx = self.D, self.ctx
        el = ctx.el("issuer")
        size = self.issuer_size(text)
        p = self.doc.add_paragraph()
        run = p.add_run(text)
        name, bold = ctx.font(el["font"])
        D.format_run(run, font=name, size=size, bold=True if bold else None, color=D.color_hex(ctx.params, el["color"]))
        before = max(0.0, top_from_body_mm * 72 / 25.4 - used_lines * ctx.line_pt)
        D.set_para(p, align="center", line_exact=round(size * float(el["line_ratio"]), 1), before_pt=round(before, 1),
                   after_pt=0, snap=False)
        return p

    def red_rule_under(self, p) -> None:
        D, ctx = self.D, self.ctx
        rule = ctx.el("header_rule")
        D.set_border(p, "bottom", sz_eighths=int(round(rule["width_pt"] * 8)),
                     color=D.color_hex(ctx.params, rule["color"]), space_pt=int(round(rule["gap_mm"] * 72 / 25.4)))

    def docno(self, docno: str, signers: List[str]):
        D, ctx = self.D, self.ctx
        el = ctx.el("docno")
        sg = ctx.el("signer")
        p = self.doc.add_paragraph()
        if signers:
            width = ctx.text_width_twips
            one = int(round(ctx.size(el["size"]) * 20))
            # 多个签发人：前面的行只排姓名，最后一行与发文字号同行（7.2.6）
            rows = [signers[i:i + 2] for i in range(0, len(signers), 2)]
            for row in rows[:-1]:
                extra = self.doc.add_paragraph()
                p._p.addprevious(extra._p)
                D.set_para(extra, align="right", right=sg["right_chars"], line_exact=ctx.line_pt, before_pt=0, after_pt=0)
                add = extra.add_run("　".join(row))
                D.format_run(add, font=ctx.font(sg["name_font"])[0], size=ctx.size(sg["size"]))
            D.set_para(p, align="left", line_exact=ctx.line_pt, before_pt=0, after_pt=0)
            D.set_tabs(p, [(one * int(el["left_chars_upward"]), "left"), (width - one * int(sg["right_chars"]), "right")])
            r1 = p.add_run("\t" + docno + "\t签发人：")
            D.format_run(r1, font=ctx.font(sg["label_font"])[0], size=ctx.size(sg["size"]))
            r2 = p.add_run("　".join(rows[-1]))
            D.format_run(r2, font=ctx.font(sg["name_font"])[0], size=ctx.size(sg["size"]))
        else:
            D.set_para(p, align=el["align"], line_exact=ctx.line_pt, before_pt=0, after_pt=0)
            run = p.add_run(docno)
            D.format_run(run, font=ctx.font(el["font"])[0], size=ctx.size(el["size"]))
        return p

    def letter_head(self, issuer: str) -> None:
        """信函格式：首页页眉放发文机关标志和红色双线（上粗下细），首页页脚放红色双线（上细下粗）。"""
        D, ctx = self.D, self.ctx
        le = ctx.el("letter")
        el = ctx.el("issuer")
        red = D.color_hex(ctx.params, "red")
        ext = (le["rule_length_mm"] - D.pv(ctx.params, "page.body_width_mm")) / 2 * 72 / 25.4 / 16
        thick = int(round(le["thick_pt"] * 8))
        thin = int(round(le["thin_pt"] * 8))
        section = self.section
        section.different_first_page_header_footer = True
        header = section.first_page_header
        header.is_linked_to_previous = False
        D.clear_part(header)
        p1 = header.paragraphs[0]
        size = self.issuer_size(issuer)
        run = p1.add_run(issuer)
        name, bold = ctx.font(el["font"])
        D.format_run(run, font=name, size=size, bold=True if bold else None, color=red)
        D.set_para(p1, align="center", left=-ext, right=-ext, line_exact=round(size * float(el["line_ratio"]), 1),
                   before_pt=0, after_pt=0, snap=False, chars=False)
        D.set_border(p1, "bottom", sz_eighths=thick, color=red, space_pt=int(round(le["rule_gap_mm"] * 72 / 25.4)))
        p2 = header.add_paragraph()
        D.set_para(p2, left=-ext, right=-ext, line_exact=le["double_gap_pt"], before_pt=0,
                   after_pt=round(ctx.size("三号") * 7 / 8, 1), snap=False, chars=False)
        D.set_border(p2, "bottom", sz_eighths=thin, color=red, space_pt=0)
        footer = section.first_page_footer
        footer.is_linked_to_previous = False
        D.clear_part(footer)
        f1 = footer.paragraphs[0]
        D.set_para(f1, left=-ext, right=-ext, line_exact=le["double_gap_pt"], before_pt=0, after_pt=0, snap=False,
                   chars=False)
        D.set_border(f1, "bottom", sz_eighths=thin, color=red, space_pt=0)
        f2 = footer.add_paragraph()
        D.set_para(f2, left=-ext, right=-ext, line_exact=le["double_gap_pt"], before_pt=0, after_pt=0, snap=False,
                   chars=False)
        D.set_border(f2, "bottom", sz_eighths=thick, color=red, space_pt=0)


def merge_fields(parsed: Parsed, meta: dict, warnings: List[str]) -> Dict[str, object]:
    """版式信息优先，正文识别结果补缺。两者不一致时给出提示。"""
    fields: Dict[str, object] = dict(parsed.fields)
    for key in META_KEYS:
        if key in meta and meta[key] not in (None, "", []):
            old = parsed.fields.get(key)
            new = meta[key]
            if key == "期号" and old and norm_issue(old) != norm_issue(new):
                warnings.append(f"版式信息中的“期号”与正文识别结果不一致，已采用版式信息：{new}（正文中为：{old}）")
            elif old and key in ("标题", "主送机关", "成文日期", "简报名称", "编印单位") and not _same(str(old), str(new)):
                warnings.append(f"版式信息中的“{key}”与正文识别结果不一致，已采用版式信息：{new}（正文中为：{old}）")
            fields[key] = new
    return fields


def decide_layout(cli_layout: Optional[str], meta: dict, parsed: Parsed, report: dict) -> str:
    raw = cli_layout or meta.get("版式")
    if raw:
        layout = LAYOUTS.get(str(raw).strip())
        if not layout:
            raise BuildError(f"无法识别的版式：{raw}。可选：红头、信函、纪要、白头、材料、简报")
        report["layout_reason"] = "命令行指定" if cli_layout else "版式信息指定"
        return layout
    if parsed.kind == "简报":
        report["layout_reason"] = "未指定版式，版式信息或正文中有简报报头，按简报排版"
        return "简报"
    issuers = parsed.find("issuer")
    if issuers:
        layout = "纪要" if issuers[0].text.endswith("纪要") else "红头"
        report["layout_reason"] = "未指定版式，正文中有发文机关标志，按" + layout + "排版"
        return layout
    kind = parsed.kind
    title = meta.get("标题") or parsed.fields.get("标题")
    if meta.get("标题"):
        kind = doc_kind(str(meta["标题"]))[0]
    if kind in ("法定", "内部"):
        report["layout_reason"] = "未指定版式，按法定公文默认白头"
        layout = "白头"
    else:
        report["layout_reason"] = "未指定版式，按常用材料默认材料版式" if title else "未识别到标题，按材料版式处理"
        layout = "材料"
    if meta.get("发文机关标志"):
        report["warnings"].append("提供了发文机关标志但未指定版式，已按" + layout + "排版；需要红头请在版式信息中写 \"版式\": \"红头\"")
    if kind == "内部":
        report["warnings"].append("签报一般用本单位统一的签报单格式，本次按白头排版，请对照单位模板调整")
    return layout


def build(lines: Sequence[str], meta: dict, *, layout: Optional[str] = None, font_fallback: bool = False,
          installed: Optional[set] = None, params: Optional[dict] = None):
    """生成文档，返回 (Document, report)。"""
    import gw_docx as D
    import gw_fonts as F
    from gw_format import keep_signoff_together, layout_signature_block

    params = params or D.load_params()
    report: dict = {"warnings": [], "placeholders": [], "detected": [], "notes": []}
    # 正文中的 Markdown 管道表格先换成占位行，识别结构后在原位置生成 Word 表格
    lines, tables, table_notes = extract_pipe_tables(lines)
    report["warnings"] += table_notes
    explicit = LAYOUTS.get(str(layout or meta.get("版式") or "").strip())
    brief_mode = explicit == "简报" or (not explicit and (has_brief_header(lines) or any(
        meta.get(k) not in (None, "", []) for k in BRIEF_META_KEYS)))
    parsed = parse_brief(lines, force=explicit == "简报") if brief_mode else parse_lines(lines)
    report["warnings"] += parsed.warnings
    report["detected"] = parsed.detected
    lay = decide_layout(layout, meta, parsed, report)
    report["layout"] = lay

    if _truthy(meta.get("三级标题加粗")):
        params = json.loads(json.dumps(params))
        params["elements"]["h3"]["bold"] = True
        report["notes"].append("按版式信息，三级标题加粗")

    method = ""
    if installed is None:
        installed, method = F.installed_fonts()
    fonts, fwarn, fnotes = F.resolve_fonts(params, font_fallback, installed, method)
    report["fonts"] = {k: v[0] + ("（加粗）" if v[1] else "") for k, v in fonts.items()}
    report["font_warnings"] = fwarn
    report["notes"] += fnotes
    ctx = D.Ctx(params, fonts)

    fields = merge_fields(parsed, meta, report["warnings"])
    title_text = str(fields.get("标题") or "")
    kind, doc_type = doc_kind(title_text)
    report["doc_type"] = doc_type
    seal_meta = _truthy(meta.get("印章"))
    seal = seal_meta if seal_meta is not None else bool(params["seal_default"][lay])
    report["seal"] = seal

    # 标题下的发文字号行（识别时已认作发文字号，不进正文）：说明去向；版式信息中有发文字号时以版式信息为准
    title_src = [it.src for it in parsed.items if it.role == "title"]
    for it in parsed.items:
        if it.role != "docno" or not title_src or it.src < title_src[0]:
            continue
        if lay not in ("红头", "信函", "纪要"):
            report["notes"].append(f"正文标题下的发文字号“{it.text}”已从正文移出；{lay}版式不排发文字号，需要时请改用红头或信函版式")
        elif meta.get("发文字号") not in (None, "", []):
            report["notes"].append(f"正文标题下的文号行“{it.text}”已从正文删除，版头采用版式信息中的发文字号“{meta['发文字号']}”")
        else:
            report["notes"].append(f"正文标题下的发文字号“{it.text}”已移到版头发文字号位置，正文中不再保留")

    def need(key: str, label: Optional[str] = None) -> str:
        val = fields.get(key)
        if isinstance(val, list):
            val = "".join(val)
        if val:
            return str(val)
        mark = f"【待补：{label or key}】"
        report["placeholders"].append(f"{label or key}：版式信息和正文中都没有，已写入 {mark}，请补齐")
        return mark

    # 日期写法：版式信息中的 2026-10-16 写法转为“2026年10月16日”
    for key in ("成文日期", "印发日期"):
        if key in meta and isinstance(fields.get(key), str):
            new, changed = normalize_date(str(fields[key]))
            if changed:
                fields[key] = new
                report["notes"].append(f"{key}已按 GB/T 9704-2012 写成“{new}”")

    b = Builder(ctx, lay)
    if lay == "简报":
        _build_brief(b, parsed, fields, meta, tables, report, need)
        parsed.fields.update(fields)  # 识别结果报告含版式信息中的报头、报尾要素
        D.setup_page_numbers(b.section, ctx, first_page_blank=False)
        report["fields"] = {k: v for k, v in fields.items() if v not in (None, "", [])}
        report["fields"]["版式"] = lay
        report["summary"] = summarize(parsed)
        return b.doc, report
    used_lines = 0
    head_lines: List[Tuple[str, str]] = []
    if fields.get("份号") and lay != "信函":
        head_lines.append(("copy_no", str(fields["份号"])))
    if fields.get("密级"):
        secret = str(fields["密级"])
        if fields.get("保密期限"):
            secret += "★" + str(fields["保密期限"])
        head_lines.append(("secret", secret))
        report["warnings"].append("保密提示：版式信息含密级。涉密文件只能在涉密环境中处理，不得输入联网 AI 工具")
    if fields.get("紧急程度"):
        head_lines.append(("urgency", str(fields["紧急程度"])))

    signers = _as_list(fields.get("签发人"))
    issuer_text = str(fields.get("发文机关标志") or "")
    body_items = [it for it in parsed.items if it.role not in HEAD_ROLES + TITLE_ROLES + TAIL_ROLES + ("recipient",)]
    title_lines = [it.text for it in parsed.items if it.role == "title"]
    if meta.get("标题"):
        title_lines = [s for s in str(meta["标题"]).split("\n") if s.strip()]
        if not parsed.find("title") and body_items and _same(body_items[0].text, str(meta["标题"])):
            body_items = body_items[1:]
    sub_items = [it for it in parsed.items if it.role in ("subtitle", "title_note")]

    # ---- 版头 ----
    if lay == "红头":
        for role, text in head_lines:
            b.para(role, text)
            used_lines += 1
        if not issuer_text:
            issuer_text = need("发文机关标志")
        b.issuer(issuer_text, ctx.el("issuer")["top_from_body_mm"], used_lines)
        b.blank(ctx.el("docno")["blank_lines_before"])
        if doc_type in UPWARD_TYPES and not signers:
            signers = [need("签发人", "签发人姓名")]
            report["warnings"].append(f"{doc_type}属上行文，红头版式须标注签发人（GB/T 9704-2012 7.2.6）")
        docno_text = need("发文字号")
        if signers:
            _warn_docno_width(docno_text, signers, ctx, report)
        p = b.docno(docno_text, signers)
        b.red_rule_under(p)
        b.blank(ctx.el("title")["blank_lines_after_rule"])
    elif lay == "信函":
        if not issuer_text:
            issuer_text = need("发文机关标志")
        b.letter_head(issuer_text)
        docno = need("发文字号")
        p = b.doc.add_paragraph()
        copy_no = str(fields.get("份号") or "")
        D.set_para(p, align="left" if copy_no else "right", line_exact=ctx.line_pt, before_pt=0, after_pt=0)
        if copy_no:
            D.set_tabs(p, [(ctx.text_width_twips, "right")])
            r = p.add_run(copy_no)
            D.format_run(r, font=ctx.font("仿宋")[0], size=ctx.size("三号"))
        r = p.add_run(("\t" if copy_no else "") + docno)
        D.format_run(r, font=ctx.font("仿宋")[0], size=ctx.size("三号"))
        for role, text in head_lines:
            b.para(role, text)
        b.blank(ctx.el("letter")["blank_lines_before_title"])
    elif lay == "纪要":
        for role, text in head_lines:
            b.para(role, text)
            used_lines += 1
        if not issuer_text and title_text.endswith("纪要"):
            issuer_text = title_text
            title_lines = []
            report["notes"].append("未提供纪要标志，已用标题“" + title_text + "”作纪要标志，不再另排标题")
        if not issuer_text:
            issuer_text = need("发文机关标志", "纪要标志")
        p = b.issuer(issuer_text, ctx.el("minutes")["top_from_body_mm"], used_lines)
        if fields.get("发文字号"):
            b.blank(ctx.el("minutes")["blank_lines_before_number"])
            p = b.docno(str(fields["发文字号"]), [])
        b.red_rule_under(p)
        b.blank(ctx.el("title")["blank_lines_after_rule"])
    else:
        for role, text in head_lines:
            b.para(role, text)
        if head_lines:
            b.blank(1)
        if lay == "白头" and (fields.get("发文字号") or issuer_text):
            report["warnings"].append("白头版式不排发文机关标志和发文字号；需要时请改用红头或信函版式")

    # ---- 标题 ----
    if len(title_lines) == 1:
        # 超长标题按词意在中点附近回行，上行不短于下行（7.3.1）；用户自己分了行的不动
        max_chars = int(ctx.text_width_twips / 20 // ctx.size(ctx.el("title")["size"]))
        lines_auto = break_title(title_lines[0], max_chars)
        if len(lines_auto) > 1:
            title_lines = lines_auto
            report["notes"].append("标题较长，已分行排列：" + " / ".join(lines_auto) +
                                   "（GB/T 9704-2012 7.3.1）；要改断行位置，在版式信息“标题”中用换行符 \\n 分行")
    if title_lines:
        b.title(title_lines)
        for it in sub_items:
            b.para(it.role, it.text)
    elif (lay != "纪要" or not issuer_text) and not any("未识别到标题" in w for w in report["warnings"]):
        report["warnings"].append("没有标题。如正文首行是标题但未被识别，请在版式信息中填写“标题”")
    if title_lines or sub_items:
        b.blank(ctx.el("title")["blank_lines_after"])

    # ---- 主送机关 ----
    recipient = fields.get("主送机关")
    if recipient:
        rec = "、".join(_as_list(recipient)) if isinstance(recipient, list) else str(recipient)
        if body_items and _same(body_items[0].text.rstrip("：:"), rec.rstrip("：:")):
            body_items = body_items[1:]
        if not rec.endswith(("：", ":")):
            rec += "："
        b.para("recipient", rec)

    # ---- 正文（含表格、表题）----
    table_queue = list(tables)
    table_font, table_size = _table_style(meta, report)
    prev_role = None
    for k, it in enumerate(body_items):
        if it.role == "table":
            if prev_role == "table":
                b.blank(1)  # 两张表格紧挨着会被 Word 合成一张，中间留一个空段
            tb = table_queue.pop(0) if table_queue else None
            if tb is not None:
                D.add_body_table(b.doc, ctx, tb.header, tb.rows, tb.aligns, font_role=table_font, size_name=table_size)
        else:
            if it.role == "attendance" and k > 0 and prev_role != "attendance":
                b.blank(1)  # 纪要出席名单在正文下空一行（GB/T 9704-2012 10.3）
            b.para(it.role, it.text, split=it.split)
        prev_role = it.role
    if tables:
        report["notes"].append(f"正文中的 {len(tables)} 张 Markdown 表格已转为 Word 表格（全框线、表头加粗、"
                               f"{table_font}{table_size}，与版心同宽）；不支持合并单元格，需要合并的请在 Word 中处理")

    # ---- 附件说明 ----
    attachments = _as_list(fields.get("附件"))
    if attachments:
        b.blank(ctx.el("attachment")["blank_lines_before"])
        names = [re.sub(r"^\d{1,2}\s*[.．、]\s*", "", a) for a in attachments]
        names = [re.sub(r"^附\s*件\s*[：:]\s*", "", a) for a in names]
        if len(names) == 1:
            att_ps = [b.para("attachment", "附件：" + names[0])]
        else:
            att_ps = [b.para("attachment", "附件：1. " + names[0])]
            for i, name in enumerate(names[1:], start=2):
                att_ps.append(b.para("attachment", f"{i}. {name}"))
        for ap in att_ps[:-1]:
            D.set_para(ap, reset_indent=False, keep_next=True)  # 附件说明各行不分页
        for name in names:
            if name and name[-1] in "。；，、.;,":
                report["warnings"].append(f"附件名称“{name}”末尾有标点，应删去（GB/T 9704-2012 7.3.4）")

    # ---- 署名、成文日期 ----
    sigs = _as_list(fields.get("发文机关署名"))
    date_text = str(fields.get("成文日期") or "")
    if lay in LEGAL_LAYOUTS:
        if not sigs:
            sigs = [need("发文机关署名")]
        if not date_text:
            date_text = need("成文日期")
    if sigs or date_text:
        el = ctx.el("signature_seal" if seal else "signature_noseal")
        b.blank(el["blank_lines_before"])
        sig_ps = [b.para("signature", s) for s in sigs]
        date_p = b.para("date", date_text) if date_text else None
        layout_signature_block(sig_ps, date_p, ctx, seal=seal)
        signoff_first = (sig_ps or [date_p])[0]
        signoff_last = date_p if date_p is not None else sig_ps[-1]
    else:
        signoff_first = signoff_last = None
        if seal and len(sigs) > 1:
            report["warnings"].append("联合行文加盖印章时，署名和印章须一一对应排列（7.3.5.1），本工具只做近似排列，请人工调整")
    note = fields.get("附注")
    if note:
        note = str(note)
        if "附注" in meta and not (note.startswith(("（", "(")) and note.endswith(("）", ")"))):
            note = f"（{note}）"
        note_p = b.para("note", note)
        if signoff_last is not None:
            signoff_last = note_p
    if signoff_first is not None:
        # 落款不单独成页：正文最后一段、空行、署名、日期（附注）设“与下段同页”
        elements = [p._p for p in b.doc.paragraphs]
        all_paras = b.doc.paragraphs
        keep_signoff_together(all_paras, elements.index(signoff_first._p), elements.index(signoff_last._p))

    # ---- 版记 ----
    cc = fields.get("抄送")
    cc_text = ""
    if cc:
        if isinstance(cc, list):
            cc_text = "抄送：" + "，".join(_as_list(cc)) + "。"
        else:
            cc_text = str(cc) if str(cc).startswith("抄送") else "抄送：" + str(cc)
            if not cc_text.endswith("。"):
                report["warnings"].append("抄送最后一个机关名称后应标句号（GB/T 9704-2012 7.4.2）")
    rows: List[List[Tuple[str, str]]] = []
    if cc_text:
        rows.append([("cc", cc_text)])
    has_print = bool(fields.get("印发机关") or fields.get("印发日期"))
    if lay == "信函":
        if has_print:
            report["warnings"].append("信函格式的版记不加印发机关和印发日期（GB/T 9704-2012 10.1），已省略")
        if rows:
            D.add_banji_table(b.doc, ctx, rows, with_lines=False)
    elif lay == "材料":
        if rows or has_print:
            report["warnings"].append("材料版式不排版记，已省略抄送和印发信息")
    elif lay == "红头" or rows or has_print:
        org = need("印发机关")
        date = need("印发日期")
        rows.append([("print", f"{org}\t{date}印发")])
        D.add_banji_table(b.doc, ctx, rows, with_lines=True)

    # 文档以表格结尾时补一个 1 磅高的空段，Word 要求表格后有段落
    last = [el for el in b.doc.element.body.iterchildren() if not el.tag.endswith("}sectPr")]
    if last and last[-1].tag.endswith("}tbl"):
        anchor = b.doc.add_paragraph()
        D.set_para(anchor, line_exact=1, before_pt=0, after_pt=0, snap=False)

    # ---- 页码 ----
    D.setup_page_numbers(b.section, ctx, first_page_blank=(lay == "信函"))
    if lay == "信函":
        # setup_page_numbers 会清空首页页脚，重画下双线
        b.letter_head(issuer_text)

    report["fields"] = {k: v for k, v in fields.items() if v not in (None, "", [])}
    report["fields"]["版式"] = lay
    report["summary"] = summarize(parsed)
    return b.doc, report


def _build_brief(b: Builder, parsed: Parsed, fields: Dict[str, object], meta: dict, tables, report: dict, need) -> None:
    """简报：报头（简报名称、期号行、编印单位和日期行、红色分隔线）、各篇标题与正文、报尾。

    简报无国家标准格式，参数取 gbt9704_params.json 的 elements.brief（常见做法，各单位以往期样式为准）。
    报头要素缺失写【待补】；报尾（报、送、发、抄送、印数）有就排，没有就不排。
    """
    D, ctx = b.D, b.ctx
    el = ctx.el("brief")
    red = D.color_hex(ctx.params, el["name_color"])
    ignored = [k for k in BRIEF_IGNORED if meta.get(k) not in (None, "", [])]
    if ignored:
        report["warnings"].append("简报版式不排" + "、".join(ignored) + "，已忽略；需要时请在 Word 中手工加注")

    # 简报名称：红色居中，字号随名称长短取小初到二号
    name = need("简报名称")
    width_pt = ctx.text_width_twips / 20
    size = max(float(el["name_size_min_pt"]), min(float(el["name_size_max_pt"]),
               width_pt * float(el["name_fill_ratio"]) / max(1.0, text_width_chars(name))))
    size = int(size * 2) / 2
    if text_width_chars(name) * size > width_pt:
        report["warnings"].append(f"简报名称“{name}”较长，一行排不下，会回行；请核对名称或缩短")
    p = b.doc.add_paragraph()
    font, bold = ctx.font(el["name_font"])
    D.format_run(p.add_run(name), font=font, size=size, bold=True if bold else None, color=red)
    D.set_para(p, align="center", line_exact=round(size * float(el["name_line_ratio"]), 1), before_pt=0, after_pt=0,
               snap=False, keep_next=True)

    # 期号行：第12期，有总期号时写成“第12期（总第68期）”
    issue = norm_issue(fields.get("期号")) if fields.get("期号") else need("期号")
    fields["期号"] = issue
    if fields.get("总期号"):
        fields["总期号"] = norm_total(fields["总期号"])
        issue += "（" + fields["总期号"] + "）"
    p = b.doc.add_paragraph()
    D.format_run(p.add_run(issue), font=ctx.font(el["issue_font"])[0], size=ctx.size(el["issue_size"]))
    D.set_para(p, align="center", line_exact=ctx.line_pt, before_pt=float(el["issue_before_pt"]), after_pt=0,
               snap=False, keep_next=True)

    # 编印行：左编印单位、右日期，同一行；行下红色分隔线
    editor = str(fields["编印单位"]) if fields.get("编印单位") else need("编印单位")
    if not editor.startswith("【待补") and not editor.endswith(("编", "编印", "编发")):
        editor += "编"
    date = str(fields["印发日期"]) if fields.get("印发日期") else need("印发日期")
    if text_width_chars(editor) + text_width_chars(date) + 1 > float(D.pv(ctx.params, "grid.chars_per_line")):
        report["warnings"].append("编印单位较长，与日期排不进同一行，日期会被挤到下一行；请缩短或改用简称")
    p = b.doc.add_paragraph()
    D.format_run(p.add_run(editor + "\t" + date), font=ctx.font(el["editor_font"])[0], size=ctx.size(el["editor_size"]))
    D.set_para(p, align="left", line_exact=ctx.line_pt, before_pt=float(el["editor_before_pt"]), after_pt=0,
               snap=False, keep_next=True)
    D.set_tabs(p, [(ctx.text_width_twips, "right")])
    D.set_border(p, "bottom", sz_eighths=int(round(el["rule_width_pt"] * 8)), color=red,
                 space_pt=int(round(el["rule_gap_mm"] * 72 / 25.4)))
    b.blank(el["blank_lines_after_rule"])

    # 各篇：标题（连续几行都算标题）、题注、层次标题、正文、表格
    table_queue = list(tables)
    table_font, table_size = _table_style(meta, report)
    items = [it for it in parsed.items if not it.role.startswith("brief_")]
    max_chars = int(ctx.text_width_twips / 20 // ctx.size(ctx.el("title")["size"]))
    prev = None
    for it in items:
        role = it.role
        if role == "title":
            if prev is not None and prev != "title":
                b.blank(1)  # 两篇之间空一行
            parts = [t for t in re.split(r"[\s　]+", it.text) if t]  # 标题中间的空格当作换行处
            if len(parts) > 1 and all(text_width_chars(t) <= max_chars for t in parts):
                lines = parts
            else:
                lines = break_title("".join(parts), max_chars)
            if len(lines) > 1 and len(parts) == 1:
                report["notes"].append("标题较长，已分行排列：" + " / ".join(lines) + "（要改断行位置，请把标题拆成两行写）")
            D.set_para(b.title(lines), reset_indent=False, keep_next=True)
        else:
            if prev in ("title", "title_note") and role != "title_note":
                b.blank(1)  # 标题下空一行，与标题、正文连在一起
                D.set_para(b.doc.paragraphs[-1], reset_indent=False, keep_next=True)
            if role == "table":
                if prev == "table":
                    b.blank(1)
                tb = table_queue.pop(0) if table_queue else None
                if tb is not None:
                    D.add_body_table(b.doc, ctx, tb.header, tb.rows, tb.aligns, font_role=table_font, size_name=table_size)
            else:
                para = b.para(role, it.text, split=it.split)
                if role == "title_note":
                    D.set_para(para, reset_indent=False, keep_next=True)
        prev = role
    if tables:
        report["notes"].append(f"正文中的 {len(tables)} 张 Markdown 表格已转为 Word 表格（全框线、表头加粗、"
                               f"{table_font}{table_size}，与版心同宽）；不支持合并单元格，需要合并的请在 Word 中处理")
    if not any(it.role == "title" for it in items):
        report["warnings"].append("没有识别到篇目标题；每篇标题应是不带句末标点的独立短行")

    # 报尾：版心底部一条分隔线，其下排报、送、发、抄送和印数；有就排，没有就不排
    foot: List[Tuple[str, str]] = []
    for key in ("报", "送", "发", "抄送"):
        val = fields.get(key)
        if val in (None, "", []):
            continue
        text = "，".join(_as_list(val)) if isinstance(val, list) else str(val).strip()
        foot.append(("cc", text if re.match(rf"^{key}\s*[：:]", text) else f"{key}：{text}"))
    if fields.get("印数"):
        n = str(fields["印数"]).strip()
        m = re.search(r"\d+", n)
        foot.append(("print", n if n.startswith(("（", "(")) else f"（共印{m.group(0) if m else n}份）"))
    if foot:
        D.add_banji_table(b.doc, ctx, [foot], with_lines="top")
    else:
        report["notes"].append("没有报送范围（报、送、发、抄送）和印数，未排报尾；需要时在版式信息中补“报”“送”“发”“印数”")


def _warn_docno_width(docno: str, signers: List[str], ctx, report: dict) -> None:
    """上行文发文字号与签发人同一行：估算宽度，超出版心时提示（Word 会把签发人挤到下一行）。"""
    import gw_docx as D
    per_line = int(D.pv(ctx.params, "grid.chars_per_line"))
    avail = per_line - int(ctx.el("docno")["left_chars_upward"]) - int(ctx.el("signer")["right_chars"])
    last_row = signers[-2:] if len(signers) % 2 == 0 else signers[-1:]
    need = text_width_chars(docno) + text_width_chars("签发人：" + "　".join(last_row)) + 1
    if need > avail:
        report["warnings"].append(
            f"发文字号“{docno}”与签发人同一行约需 {need:g} 字，超过可排的 {avail} 字，签发人会被挤到下一行；"
            "序号未定的请写成“代字〔年份〕【待补：序号】号”，不要用整段长占位（GB/T 9704-2012 7.2.5、7.2.6）")


def _table_style(meta: dict, report: dict) -> Tuple[str, str]:
    """表格字体字号：默认仿宋四号；版式信息可写“表格字体”（仿宋、宋体）、“表格字号”（四号、小四）。"""
    font, size = "仿宋", "四号"
    for key, allowed in (("表格字体", TABLE_FONTS), ("表格字号", TABLE_SIZES)):
        val = str(meta.get(key) or "").strip()
        if not val:
            continue
        if val in allowed:
            if key == "表格字体":
                font = val
            else:
                size = val
        else:
            report["warnings"].append(f"版式信息“{key}”只能写{'、'.join(allowed)}，“{val}”未采用，按默认处理")
    return font, size


def render_text(report: dict) -> str:
    out: List[str] = []
    out.append(f"输出文件：{report.get('output', '')}")
    out.append(f"版式：{report['layout']}（{report.get('layout_reason', '')}）；印章：{'加盖' if report.get('seal') else '不加盖'}")
    nums = iter("一二三四五六七八九十")  # 各部分按实际出现的顺序连续编号
    out.append(f"{next(nums)}、识别结果（请核对）")
    for i, line in enumerate(report.get("summary", []), start=1):
        out.append(f"{i}. {line}")
    for line in report.get("detected", []):
        out.append(f"   依据：{line}")
    out.append(f"{next(nums)}、字体")
    out.append("   " + "；".join(f"{k}→{v}" for k, v in report.get("fonts", {}).items()))
    for w in report.get("font_warnings", []):
        out.append(f"   警告：{w}")
    sections = [("待补事项", report.get("placeholders", [])), ("提示", report.get("warnings", []) + report.get("notes", []))]
    for head, items in sections:
        if items:
            out.append(f"{next(nums)}、{head}")
            for i, w in enumerate(items, start=1):
                out.append(f"{i}. {w}")
    chk = report.get("check")
    if chk:
        s = chk["summary"]
        out.append(f"{next(nums)}、回读检查：P0 {s['P0']} 项，P1 {s['P1']} 项，P2 {s['P2']} 项")
        for issue in chk["issues"]:
            if issue["level"] in ("P0", "P1"):
                out.append(f"   [{issue['level']}] {issue['message']}：{issue['excerpt']}")
    return "\n".join(out)


def main(argv: Optional[Sequence[str]] = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    parser = argparse.ArgumentParser(description="把公文正文排成符合 GB/T 9704-2012 的 docx。",
                                     epilog="示例：python scripts/build_docx.py 正文.txt --meta 版式信息.json -o 成稿.docx")
    parser.add_argument("input", help="正文文件（.txt 或 .md），遇到【起草说明】行停止读取")
    parser.add_argument("--meta", help="版式信息 JSON（键用中文：版式、发文字号、标题、主送机关等）")
    parser.add_argument("-o", "--output", help="输出 docx 路径；默认输出到正文同目录：正文.txt 对应“成稿.docx”，正文-××.txt "
                                               "对应“成稿-××.docx”，其他文件用原文件名（如 送审稿.txt 对应“送审稿.docx”）；"
                                               "已存在时加 -v2、-v3")
    parser.add_argument("--layout", help="版式：红头、信函、纪要、白头、材料、简报（优先于版式信息）")
    parser.add_argument("--font-fallback", action="store_true", help="本机缺标准字体时改用已安装的替代字体")
    parser.add_argument("--json", action="store_true", help="以 JSON 输出报告")
    parser.add_argument("--no-check", action="store_true", help="生成后不自动回读检查")
    parser.add_argument("--allow", metavar="DOCX-050", help="回读检查时按本单位制度保留的检查项（逗号分隔），命中降为 P2")
    args = parser.parse_args(argv)

    try:
        import docx  # noqa: F401
    except ImportError:
        print("缺少 python-docx。请先安装：pip install -U python-docx（需要 1.2.0 及以上版本）")
        return 2
    import gw_docx as D

    src = Path(args.input)
    if not src.is_file():
        print(f"错误：找不到正文文件 {src}")
        return 2
    if src.suffix.lower() not in (".txt", ".md", ".markdown"):
        print("错误：正文文件须为 .txt 或 .md")
        return 2
    meta: dict = {}
    if args.meta:
        try:
            meta = json.loads(Path(args.meta).read_text(encoding="utf-8-sig"))
        except (OSError, ValueError) as exc:  # ValueError 含 JSON 格式错误和非 UTF-8 编码
            print(f"错误：版式信息 JSON 读取失败：{exc}")
            return 2
        if not isinstance(meta, dict):
            print("错误：版式信息 JSON 顶层必须是对象")
            return 2
    try:
        lines, read_notes = read_body_text(src)
    except (OSError, ValueError) as exc:  # 无法读取或无法识别编码
        print(f"错误：正文文件读取失败：{exc}")
        return 2
    try:
        doc, report = build(lines, meta, layout=args.layout, font_fallback=args.font_fallback)
    except BuildError as exc:
        print(f"错误：{exc}")
        return 2
    report["notes"] += read_notes
    unknown = [k for k in meta if k not in META_KEYS]
    if unknown:
        report["warnings"].append("版式信息中有未识别的键，已忽略：" + "、".join(unknown))
    # 默认输出名：正文.txt → 成稿.docx；正文-单位总结.txt → 成稿-单位总结.docx（output-format.md 的多版本命名）；
    # 其他文件用原文件名，如 送审稿.txt → 送审稿.docx，避免不同稿本都输出成“成稿”
    if src.stem == "正文":
        default_name = "成稿.docx"
    elif src.stem.startswith("正文-"):
        default_name = "成稿-" + src.stem[3:] + ".docx"
    else:
        default_name = src.stem + ".docx"
    out = Path(args.output) if args.output else src.with_name(default_name)
    if out.is_dir():  # -o 给的是已有文件夹：在文件夹内按默认名输出
        out = out / default_name
    final = D.unique_path(out)
    if final != out:
        report["notes"].append(f"{out.name} 已存在，未覆盖，另存为 {final.name}")
    final.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(final))
    report["input"] = str(src)
    report["output"] = str(final)
    if not args.no_check:
        from check_docx import check_document
        report["check"] = check_document(final, layout=report["layout"], allow=args.allow)
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(render_text(report))
    return 0


if __name__ == "__main__":
    sys.exit(main())
