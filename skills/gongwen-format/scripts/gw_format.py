# -*- coding: utf-8 -*-
"""按角色设置段落和文字格式。build_docx.py 生成时、fix_docx.py 修复时共用同一套规则，保证结果一致。"""
from __future__ import annotations

import re
from typing import List, Optional, Sequence, Set

import gw_docx as D
from gw_layout import ATTENDANCE_RE, text_width_chars

ATTACH_NUM_RE = re.compile(r"^\d{1,2}\s*[.．]\s*")
ATTACH_LABEL_RE = re.compile(r"^附\s*件\s*[：:]\s*")


def _font_runs(paragraph, ctx: D.Ctx, font_role: str, size_name, *, start: int = 0, end: Optional[int] = None,
               bold: Optional[bool] = None, color: Optional[str] = None, black: bool = True) -> Set[str]:
    """给 [start, end) 范围内的 run 设字体字号（按字符位置取 run，不拆分）。"""
    changes: Set[str] = set()
    name = ctx.font(font_role)
    size = ctx.size(size_name)
    pos = 0
    for run in paragraph.runs:
        length = len(run.text)
        in_range = length == 0 or (pos >= start and (end is None or pos < end))
        if in_range:
            rpr = run._r.get_or_add_rPr()
            if D.set_rpr(rpr, font=name):
                changes.add("字体")
            if D.set_rpr(rpr, size=size):
                changes.add("字号")
            if bold is not None and D.set_rpr(rpr, bold=bold):
                changes.add("加粗")
            if color:
                if D.set_rpr(rpr, color=color):
                    changes.add("颜色")
            elif black and D.set_rpr(rpr, black=True):
                changes.add("颜色")
        pos += length
    return changes


def _ensure_split(paragraph, offset: int) -> bool:
    """确保 offset 处是 run 边界；返回是否成功。"""
    if offset <= 0:
        return True
    pos = 0
    for run in paragraph.runs:
        length = len(run.text)
        if pos == offset:
            return True
        if pos < offset < pos + length:
            return D.split_run(run, offset - pos) is not None
        pos += length
    return pos == offset or offset >= pos


def strip_leading_spaces(paragraph) -> bool:
    """删除段首用于缩进的空格（半角、全角空格和制表符），缩进改由段落格式实现。"""
    changed = False
    for run in paragraph.runs:
        text = run.text
        if not text:
            continue
        new = text.lstrip(" 　\t")
        if new != text:
            if not D.run_splittable(run):
                return changed
            run.text = new
            changed = True
        if new:
            break
    return changed


def apply_role(paragraph, role: str, ctx: D.Ctx, *, split: int = -1, black: bool = True) -> Set[str]:
    """按角色设置格式，返回发生变化的类别集合（段落格式、字体、字号、加粗、颜色）。

    split：层次标题与正文同段时，标题部分的结束位置；-1 表示整段都是标题。
    署名和成文日期的缩进由 layout_signature_block 统一计算，这里只设字体。
    """
    changes: Set[str] = set()
    line = ctx.line_pt
    text = D.direct_text(paragraph)

    def para(**kw):
        if D.set_para(paragraph, **kw):
            changes.add("段落格式")

    if role == "body":
        el = ctx.el("body")
        para(align=el["align"], first_line=el["first_line_chars"], line_exact=line, before_pt=0, after_pt=0, snap=True)
        changes |= _font_runs(paragraph, ctx, el["font"], el["size"], black=black)
    elif role in ("h1", "h2", "h3", "h4"):
        el = ctx.el(role)
        body = ctx.el("body")
        end = len(text) if split is None or split < 0 else split
        # 独占一行、末尾不带句号的标题设“与下段同页”，标题不落在页末；与正文同段的、以句号结尾的条目不设
        standalone = end >= len(text) and text.rstrip()[-1:] not in "。；！？.;!?"
        para(align=body["align"], first_line=el["first_line_chars"], line_exact=line, before_pt=0, after_pt=0,
             outline=el["outline_level"], snap=True, keep_next=True if standalone else None)
        if end < len(text) and not _ensure_split(paragraph, end):
            end = len(text)  # 拆不开时整段按标题处理
        bold = bool(el.get("bold", False))  # 层次标题默认不加粗；三级标题可按单位规定加粗
        changes |= _font_runs(paragraph, ctx, el["font"], el["size"], start=0, end=end, bold=bold, black=black)
        if end < len(text):
            changes |= _font_runs(paragraph, ctx, body["font"], body["size"], start=end, black=black,
                                  bold=False if bold else None)
    elif role == "title":
        el = ctx.el("title")
        para(align=el["align"], line_exact=line, before_pt=0, after_pt=0, snap=False)
        changes |= _font_runs(paragraph, ctx, el["font"], el["size"], bold=False, black=black)
    elif role in ("subtitle", "title_note"):
        el = ctx.el("subtitle")
        para(align=el["align"], line_exact=line, before_pt=0, after_pt=0, snap=True)
        changes |= _font_runs(paragraph, ctx, el["font"], el["size"], bold=False, black=black)
    elif role == "table_caption":
        # 表题：居中，不缩进，与表格同页
        el = ctx.el("table")
        para(align="center", line_exact=line, before_pt=0, after_pt=0, snap=True, keep_next=True)
        changes |= _font_runs(paragraph, ctx, el["caption_font"], el["caption_size"], bold=False, black=black)
    elif role == "recipient":
        el = ctx.el("recipient")
        para(align=el["align"], line_exact=line, before_pt=0, after_pt=0, snap=True)
        changes |= _font_runs(paragraph, ctx, el["font"], el["size"], black=black)
    elif role == "attendance":
        el = ctx.el("attendance")
        m = ATTENDANCE_RE.match(text)
        label_end = m.end() if m else 0
        label_w = text_width_chars(text[:label_end]) if m else 3
        para(align="justify", left=el["left_chars"] + label_w, hanging=label_w, line_exact=line,
             before_pt=0, after_pt=0, snap=True)
        if label_end and _ensure_split(paragraph, label_end):
            changes |= _font_runs(paragraph, ctx, el["label_font"], el["size"], start=0, end=label_end, black=black)
            changes |= _font_runs(paragraph, ctx, el["font"], el["size"], start=label_end, black=black)
        else:
            changes |= _font_runs(paragraph, ctx, el["font"], el["size"], black=black)
    elif role == "attachment":
        el = ctx.el("attachment")
        base = el["left_chars"]
        m_label = ATTACH_LABEL_RE.match(text)
        if m_label:
            rest = text[m_label.end():]
            m_num = ATTACH_NUM_RE.match(rest)
            num_w = text_width_chars(m_num.group(0)) if m_num else 0
            label_w = 3
            para(align="left", left=base + label_w + num_w, hanging=label_w + num_w, line_exact=line,
                 before_pt=0, after_pt=0, snap=True)
        else:
            m_num = ATTACH_NUM_RE.match(text)
            num_w = text_width_chars(m_num.group(0)) if m_num else 0
            para(align="left", left=base + 3 + num_w, hanging=num_w if num_w else None, line_exact=line,
                 before_pt=0, after_pt=0, snap=True)
        changes |= _font_runs(paragraph, ctx, el["font"], el["size"], black=black)
    elif role == "note":
        el = ctx.el("note")
        para(align="left", first_line=el["first_line_chars"], line_exact=line, before_pt=0, after_pt=0, snap=True)
        changes |= _font_runs(paragraph, ctx, el["font"], el["size"], black=black)
    elif role in ("signature", "date"):
        el = ctx.el("signature_noseal")
        changes |= _font_runs(paragraph, ctx, el["font"], el["size"], black=black)
    elif role in ("copy_no", "secret", "urgency"):
        el = ctx.el(role)
        para(align=el["align"], line_exact=line, before_pt=0, after_pt=0, snap=True)
        changes |= _font_runs(paragraph, ctx, el["font"], el["size"], black=black)
    elif role == "issuer":
        # 版头标志只统一字体，保留原有字号、颜色和位置
        el = ctx.el("issuer")
        name = ctx.font(el["font"])
        for run in paragraph.runs:
            if D.set_rpr(run._r.get_or_add_rPr(), font=name):
                changes.add("字体")
    elif role == "docno":
        # 发文字号用仿宋；同一行有签发人时，“签发人：”用仿宋、姓名用楷体（7.2.6）
        el = ctx.el("docno")
        sg = ctx.el("signer")
        m = re.search(r"签发人\s*[：:]\s*", text)
        name_start = m.end() if m and _ensure_split(paragraph, m.end()) else len(text)
        changes |= _font_runs(paragraph, ctx, el["font"], el["size"], start=0, end=name_start, black=False)
        if name_start < len(text):
            changes |= _font_runs(paragraph, ctx, sg["name_font"], sg["size"], start=name_start, black=False)
    elif role == "lingno":
        # 命令（令）格式的令号：居中，三号仿宋（10.2）
        el = ctx.el("order")
        para(align="center", line_exact=line, before_pt=0, after_pt=0, snap=True)
        changes |= _font_runs(paragraph, ctx, el["number_font"], el["number_size"], black=black)
    elif role == "attach_label":
        # 附件另面编排的“附件”“附件1”：三号黑体，顶格（7.3.7）
        el = ctx.el("attach_page")
        para(align="left", line_exact=line, before_pt=0, after_pt=0, snap=True, keep_next=True)
        changes |= _font_runs(paragraph, ctx, el["label_font"], el["label_size"], bold=False, black=black)
    elif role in ("cc", "print", "banji_main", "brief_foot"):
        el = ctx.el("banji")
        changes |= _font_runs(paragraph, ctx, el["font"], el["size"], black=black)
    return changes


def ensure_upward_docno_indent(paragraph, doc, ctx: D.Ctx) -> bool:
    """上行文发文字号居左空一字（7.2.5）：行首空白和左缩进合计不足一字时，补足左缩进；不改文字。
    居中的改为左对齐、左缩进一字。返回是否有改动。"""
    need = float(ctx.el("docno")["left_chars_upward"])
    have = D.leading_offset_chars(paragraph, doc)
    centered = D.para_align(paragraph, doc) in ("center", "right") and not paragraph.text.startswith("\t")
    if have >= need - 0.1 and not centered:
        return False
    left = D.para_left_chars(paragraph, doc)
    new_left = need if centered else left + (need - have)
    return D.set_para(paragraph, align="left", left=max(0.0, new_left), reset_indent=True, line_exact=ctx.line_pt,
                      before_pt=0, after_pt=0)


def layout_signature_block(sig_paras: Sequence, date_para, ctx: D.Ctx, *, seal: bool,
                           before_lines: Optional[int] = None) -> List[Set[str]]:
    """排署名和成文日期，返回每段的改动类别（顺序为各署名段、成文日期段）。

    - 加盖印章（7.3.5.1）：成文日期右空四字，署名以成文日期为准居中。两段都设为居中对齐、
      左右缩进相同，中心点完全一致；中心点按“右空四字加日期宽度的一半”估算，日期右空字数是近似值。
    - 不加盖印章（7.3.5.2）：署名右空二字，左对齐、左缩进 = 28 − 2 − 署名字数；成文日期左对齐、
      左缩进 = 署名左缩进 + 2，首字比署名首字右移二字（用左缩进定位，不受日期中半角数字宽度影响）。
      日期加 2 字后会超出“右空二字”时，日期右对齐、右空二字，署名左缩进 = 28 − 2 − 日期字数 − 2。
    before_lines：第一个要素前的段前空行数（修复已有文档时用段前距代替插入空段）。
    """
    width = float(D.pv(ctx.params, "grid.chars_per_line"))
    sig_w = [text_width_chars(D.direct_text(p)) for p in sig_paras]
    date_w = text_width_chars(D.direct_text(date_para)) if date_para is not None else 0.0
    # 每段的 (对齐, 左缩进, 右缩进)
    places: List[tuple] = []
    if seal:
        el = ctx.el("signature_seal")
        r_date = float(el["date_right_chars"])
        if date_para is None:
            places = [("right", None, r_date) for _ in sig_w]
        else:
            # 居中区域取以中心点为轴、尽量宽的对称区间，给半角数字的实际宽度留余量，避免日期回行
            center = width - r_date - date_w / 2
            need = max(sig_w + [date_w]) / 2 + 0.5
            half = width - center
            if half < need:  # 署名太长，以日期为准居中会超出版心：整体左移
                center, half = width - need, need
            left = max(0.0, center - half)
            right = max(0.0, width - center - half)
            places = [("center", left, right) for _ in sig_w] + [("center", left, right)]
    else:
        el = ctx.el("signature_noseal")
        r_sig = float(el["sign_right_chars"])
        shift = float(el["date_shift_chars"])
        if sig_w and date_para is not None:
            w_max = max(sig_w)
            start = width - r_sig - w_max
            if date_w + shift > w_max:  # 日期较长：日期右空二字，署名相应左移
                start = width - r_sig - date_w - shift
                date_place = ("right", None, r_sig)
            else:
                date_place = ("left", start + shift, None)
            places = [("left", max(0.0, start), None) for _ in sig_w] + [date_place]
        else:
            places = [("right", None, r_sig) for _ in sig_w] + ([("right", None, r_sig)] if date_para is not None else [])
    targets = list(sig_paras) + ([date_para] if date_para is not None else [])
    result: List[Set[str]] = []
    line = ctx.line_pt
    for k, (p, (align, left, right)) in enumerate(zip(targets, places)):
        changes: Set[str] = set()
        before = (before_lines or 0) * line if k == 0 and before_lines else 0
        if D.set_para(p, align=align, left=round(left, 2) if left else None, right=round(right, 2) if right else None,
                      line_exact=line, before_pt=before, after_pt=0, snap=True, keep_lines=True):
            changes.add("段落格式")
        changes |= _font_runs(p, ctx, el["font"], el["size"])
        result.append(changes)
    return result


def keep_signoff_together(paragraphs: Sequence, first_idx: int, last_idx: int) -> int:
    """落款不单独成页：从正文最后一段（或附件说明）起，经空行、署名到成文日期（有附注时到附注），
    依次设“与下段同页”。正文最后一段不设“段中不分页”，换页时至少有一行正文随落款下移。

    paragraphs 为正文段落列表，first_idx 为第一个署名（或日期）段的下标，last_idx 为链条最后一段的下标。
    返回设置的段数。
    """
    count = 0
    for j in range(first_idx, last_idx):
        if D.set_para(paragraphs[j], reset_indent=False, keep_next=True):
            count += 1
    j = first_idx - 1
    while j >= 0:
        if D.set_para(paragraphs[j], reset_indent=False, keep_next=True):
            count += 1
        if paragraphs[j].text.strip():
            break
        j -= 1
    return count
