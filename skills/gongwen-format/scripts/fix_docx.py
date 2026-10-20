#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按 GB/T 9704-2012 修复已有 docx 的格式。

按段落文本识别角色（标题、主送、各级层次标题、正文、附件说明、署名、日期等），统一页面设置、
字体字号、行距、缩进和页码。只改格式，不改文字（段首用于缩进的空格除外）。
输出新文件，不修改原文件；同时输出一份修复报告（Markdown），并对新文件运行 check_docx.py。

退出码：0 成功；2 输入错误或缺少依赖。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence

sys.path.insert(0, str(Path(__file__).resolve().parent))

import gw_fonts as F  # noqa: E402
from gw_layout import (BODY_ROLES, ROLE_LABELS, doc_kind, draft_note_index, has_brief_header,  # noqa: E402
                       parse_brief, parse_lines, summarize)

# 两个非空段落之间应保留的空行数（键为前一段角色、后一段角色；None 表示任意）
GAP_RULES = [
    (("issuer", None), 2), (("docno", None), 2), (("title", None), 1), (("subtitle", None), 1),
    (("title_note", None), 1), (("brief_editor", None), 1), ((None, "attachment"), 1), ((None, "signature"), -1), ((None, "date"), -1),
]


def _gap_wanted(prev_role: Optional[str], next_role: Optional[str], seal: bool) -> int:
    for (a, b), n in GAP_RULES:
        if (a is None or a == prev_role) and (b is None or b == next_role):
            if n == -1:
                return 2 if seal else 1
            return n
    return 0


GAP_BEFORE_ATTENDANCE = ("body", "h1", "h2", "h3", "h4", "attachment")


def _first_attendance(parsed):
    """第一个出席（请假、列席）段；它前面是正文、层次标题或附件说明时才需要空一行，否则返回 None。"""
    items = [it for it in parsed.items if it.src >= 0]
    for k, it in enumerate(items):
        if it.role == "attendance":
            return it if k > 0 and items[k - 1].role in GAP_BEFORE_ATTENDANCE else None
    return None


def _is_empty(p) -> bool:
    from docx.oxml.ns import qn
    if p.text.strip():
        return False
    el = p._p
    for tag in ("w:drawing", "w:pict", "w:object", "w:fldChar", "w:sectPr"):
        if el.find(".//" + qn(tag)) is not None:
            return False
    return True


def _snapshot_page(section) -> Dict[str, str]:
    from docx.oxml.ns import qn
    out: Dict[str, str] = {}
    sect = section._sectPr
    for tag, attrs in (("pgSz", ("w", "h")), ("pgMar", ("top", "bottom", "left", "right", "footer")),
                       ("docGrid", ("type", "linePitch", "charSpace"))):
        el = sect.find(qn(f"w:{tag}"))
        for a in attrs:
            out[f"{tag}.{a}"] = el.get(qn(f"w:{a}")) if el is not None else None
    return out


def _describe_page_change(before: Dict[str, str], after: Dict[str, str]) -> List[str]:
    import gw_docx as D
    labels = {"pgSz.w": "纸张宽", "pgSz.h": "纸张高", "pgMar.top": "上边距", "pgMar.bottom": "下边距",
              "pgMar.left": "左边距", "pgMar.right": "右边距", "pgMar.footer": "页脚距边界"}
    out = []
    for key, label in labels.items():
        if before.get(key) != after.get(key):
            b = f"{D.twips_to_mm(int(before[key])):.1f}" if before.get(key) else "未设置"
            a = f"{D.twips_to_mm(int(after[key])):.1f}"
            out.append(f"{label} {b} → {a} mm")
    if any(before.get(k) != after.get(k) for k in ("docGrid.type", "docGrid.linePitch", "docGrid.charSpace")):
        out.append("文档网格改为每页 22 行、每行 28 字")
    return out


def fix_document(src: Path, dst: Path, *, layout: Optional[str] = None,
                 seal: Optional[bool] = None, installed: Optional[set] = None, params: Optional[dict] = None,
                 unit_fonts=None) -> dict:
    """修复 src，保存到 dst（调用方保证 dst 不覆盖已有文件）。返回报告字典。

    字体严格模式：各角色一律改为标准字体名；installed（已安装字体名集合，空集合表示无法检测）只用于核对，
    缺字体时在报告中给出字体核对结果（font_status、missing_fonts、font_notice 等），不替代、不降级。
    unit_fonts：按单位规定指定标准字体（“仿宋=方正仿宋_GBK”字符串或 {角色: 字体名}），指定后该角色改为单位字体名，严格模式不变；
    角色或字体名不合法时抛 ValueError。
    """
    import gw_docx as D
    from docx import Document
    from docx.oxml.ns import qn
    from gw_format import apply_role, keep_signoff_together, layout_signature_block, strip_leading_spaces

    params = params or D.load_params()
    params = F.apply_unit_fonts(params, F.parse_unit_fonts(unit_fonts, params))
    doc = Document(str(src))
    report: dict = {"input": str(src), "output": str(dst), "page": [], "roles": {}, "notes": [], "warnings": [],
                    "untouched": []}
    method = ""
    if installed is None:
        installed, method = F.installed_fonts()
    fonts, check = F.resolve_fonts(params, installed, method)
    report["fonts"] = fonts
    report.update(check)
    ctx = D.Ctx(params)

    # ---- 版式判断 ----
    paras = list(doc.paragraphs)
    texts = [D.direct_text(p) or p.text for p in paras]
    # 【起草说明】之后不是正文，不排版、保持原样（output-format.md 第二部分）
    cut = draft_note_index(texts)
    after_cut: set = set()  # 【起草说明】之后的表格元素
    if cut is not None:
        cut_el = paras[cut]._p
        seen = False
        for child in doc.element.body.iterchildren():
            seen = seen or child is cut_el
            if seen and child.tag == qn("w:tbl"):
                after_cut.add(child)
        report["untouched"].append(f"第 {cut + 1} 段为【起草说明】，从该段起共 {len(paras) - cut} 段未排版；"
                                   "起草说明不是正文，成稿中应删除")
        paras, texts = paras[:cut], texts[:cut]
    # 正文表格（不含版记）在识别时作占位，表格上一行的表题按表题排，不当成署名或正文
    anchors = [a for a, _ in D.body_tables(doc) if cut is None or a < cut]
    red_flags = [D.para_is_red(p, doc) for p in paras]
    sec0 = doc.sections[0]
    is_letter = bool(sec0.different_first_page_header_footer and
                     "".join(t.text or "" for t in sec0.first_page_header._element.iter(
                         "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")).strip())
    brief = layout == "简报" or (layout is None and not is_letter and has_brief_header(texts, red_flags))
    parsed = (parse_brief(texts, red=red_flags, tables_after=anchors, force=layout == "简报") if brief
              else parse_lines(texts, red=red_flags, tables_after=anchors))
    kind, doc_type = doc_kind(str(parsed.fields.get("标题") or ""))
    if layout is None:
        if brief:
            layout = "简报"
        elif is_letter:
            layout = "信函"
        elif parsed.find("issuer"):
            layout = "纪要" if parsed.find("issuer")[0].text.endswith("纪要") else "红头"
        else:
            layout = "白头" if kind == "法定" else "材料"
    is_letter = is_letter or layout == "信函"
    if seal is None:
        seal = bool(params["seal_default"].get(layout, False))
    report["layout"] = layout
    report["seal"] = seal
    report["summary"] = summarize(parsed)
    report["warnings"] += [w for w in parsed.warnings if "空格缩进" not in w]

    # ---- 页面 ----
    D.setup_document(doc, ctx)
    for section in doc.sections:
        before = _snapshot_page(section)
        D.setup_section(section, ctx, letter=is_letter)
        report["page"] += _describe_page_change(before, _snapshot_page(section))

    # ---- 段落 ----
    role_of: Dict[int, str] = {}
    stats: Dict[str, Dict[str, int]] = {}
    stripped = 0
    for it in parsed.items:
        if it.role == "table":
            continue  # 正文表格不调整（见下文“表格”）
        p = paras[it.src]
        role_of[it.src] = it.role
        # 发文字号行的行首制表符用于“居左空一字”，版头要素不删行首空白
        if it.role not in ("docno", "signer", "issuer", "copy_no", "secret", "urgency", "brief_name", "brief_issue",
                           "brief_editor") and strip_leading_spaces(p):
            stripped += 1
        changes = apply_role(p, it.role, ctx, split=it.split)
        bucket = stats.setdefault(it.role, {"段数": 0})
        bucket["段数"] += 1
        for c in changes:
            bucket[c] = bucket.get(c, 0) + 1
    if stripped:
        report["notes"].append(f"删除段首空格缩进 {stripped} 段，缩进改由“首行缩进 2 字符”实现")

    # 空行：结构性空行保留，其余压到 1 磅
    nonempty = [i for i in range(len(paras)) if not _is_empty(paras[i])]
    compressed = 0
    normalized = 0
    for k, i in enumerate(nonempty + [len(paras)]):
        prev_i = nonempty[k - 1] if k > 0 else -1
        gap = list(range(prev_i + 1, i))
        if not gap or (prev_i < 0 and layout in ("红头", "纪要")):
            continue  # 红头文件首行前的空段可能用于定位发文机关标志，保持原样
        want = _gap_wanted(role_of.get(prev_i), role_of.get(i), seal) if prev_i >= 0 and i < len(paras) else 0
        if prev_i >= 0 and i < len(paras) and role_of.get(prev_i) in BODY_ROLES and role_of.get(i) in BODY_ROLES:
            want = 0
        if role_of.get(i) == "attendance" and role_of.get(prev_i) in GAP_BEFORE_ATTENDANCE:
            want = 1  # 纪要出席名单在正文下空一行（10.3）
        if layout == "简报" and role_of.get(i) == "title" and role_of.get(prev_i) not in (None, "title") \
                and not str(role_of.get(prev_i)).startswith("brief_"):
            want = 1  # 简报两篇之间空一行
        for n, j in enumerate(gap):
            if n < want:
                if D.set_para(paras[j], line_exact=ctx.line_pt, before_pt=0, after_pt=0, reset_indent=False):
                    normalized += 1
            else:
                if D.set_para(paras[j], line_exact=1, before_pt=0, after_pt=0, reset_indent=False, snap=False):
                    compressed += 1
    if compressed:
        report["notes"].append(f"多余空行 {compressed} 个已压缩为 1 磅高（未删除，可在 Word 中显示编辑标记查看）")

    def empties_before(idx: int) -> int:
        n = 0
        j = idx - 1
        while j >= 0 and _is_empty(paras[j]):
            n += 1
            j -= 1
        return min(n, 2 if seal else 1)

    # 标题后空一行、附件说明前空一行：没有空段时用段前段后距补足
    titles = [it for it in parsed.items if it.role in ("title", "subtitle", "title_note")]
    if layout == "简报":  # 每篇标题后空一行
        for k, it in enumerate(titles):
            nxt = titles[k + 1].src if k + 1 < len(titles) else None
            if it.src + 1 < len(paras) and not _is_empty(paras[it.src + 1]) and nxt != it.src + 1:
                D.set_para(paras[it.src], after_pt=ctx.line_pt, before_pt=0, reset_indent=False)
    elif titles:
        last = titles[-1].src
        if last + 1 < len(paras) and not _is_empty(paras[last + 1]):
            D.set_para(paras[last], after_pt=ctx.line_pt, before_pt=0, reset_indent=False)
    atts = parsed.find("attachment")
    if atts and atts[0].src > 0 and not _is_empty(paras[atts[0].src - 1]):
        D.set_para(paras[atts[0].src], before_pt=ctx.line_pt, after_pt=0, reset_indent=False)
    for it in atts[:-1]:
        D.set_para(paras[it.src], reset_indent=False, keep_next=True)  # 附件说明各行不分页
    # 纪要出席名单前空一行：前面没有空段时用段前距补足
    first_att = _first_attendance(parsed)
    if first_att is not None and first_att.src > 0 and not _is_empty(paras[first_att.src - 1]):
        D.set_para(paras[first_att.src], before_pt=ctx.line_pt, after_pt=0, reset_indent=False)
        report["notes"].append("纪要出席名单前已空一行（段前距一行，GB/T 9704-2012 10.3）")

    sigs = [paras[it.src] for it in parsed.find("signature")]
    dates = parsed.find("date")
    date_p = paras[dates[0].src] if dates else None
    if sigs or date_p is not None:
        first_idx = (parsed.find("signature") or dates)[0].src
        need = (2 if seal else 1) - empties_before(first_idx)
        per_para = layout_signature_block(sigs, date_p, ctx, seal=seal, before_lines=max(0, need))
        roles_seq = ["signature"] * len(sigs) + (["date"] if date_p is not None else [])
        for role, changes in zip(roles_seq, per_para):
            bucket = stats.setdefault(role, {"段数": 0})
            if "段落格式" in changes:
                bucket["段落格式"] = bucket.get("段落格式", 0) + 1
        if any(per_para):
            report["notes"].append("署名和成文日期已按" + ("加盖印章（7.3.5.1）" if seal else "不加盖印章（7.3.5.2）") + "重新排列")
        notes = parsed.find("note")
        last_idx = notes[0].src if notes else (dates[0].src if dates else parsed.find("signature")[-1].src)
        if keep_signoff_together(paras, first_idx, last_idx):
            report["notes"].append("正文最后一段至落款已设“与下段同页”，落款不会单独成页")

    for role, bucket in stats.items():
        report["roles"][ROLE_LABELS.get(role, role)] = bucket

    # ---- 表格 ----
    other_tables = 0
    for table in doc.tables:
        if table._tbl in after_cut:
            continue  # 起草说明中的附表不处理
        cells_text = "".join(c.text for row in table.rows for c in row.cells)
        if "抄送" in cells_text or cells_text.strip().endswith("印发") or (
                layout == "简报" and D.is_banji_table(table._tbl)):
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if p.text.strip():
                            apply_role(p, "cc", ctx)
            report["notes"].append(("简报报尾表格" if layout == "简报" else "版记表格") + "已统一为 4 号仿宋")
        else:
            other_tables += 1
    if other_tables:
        report["untouched"].append(f"正文中的表格 {other_tables} 个未调整（国标未规定表格字号，按本单位要求处理）")

    # ---- 页码 ----
    for section in doc.sections:
        D.setup_page_numbers(section, ctx, first_page_blank=False)
    report["notes"].append("页码已重设为 4 号宋体“— 1 —”，单页居右空一字、双页居左空一字" +
                           ("；信函首页页脚保持原样" if is_letter else ""))
    if any(not h.is_linked_to_previous and "".join(t.text or "" for t in h._element.iter(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")).strip()
           for s in doc.sections for h in (s.header, s.even_page_header)):
        report["untouched"].append("页眉中有内容，公文一般不用页眉，请确认是否删除（未自动删除）")
    report["stripped"] = stripped
    report["compressed"] = compressed
    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(dst))
    return report


def render_report(report: dict) -> str:
    out = ["# 修复报告", "", f"- 原文件：{report['input']}（未修改）", f"- 新文件：{report['output']}",
           f"- 版式判断：{report['layout']}；署名按{'加盖印章' if report['seal'] else '不加盖印章'}排列"]
    status = F.status_line(report)
    if status:
        out.append(f"- 字体：文件中一律写入标准字体名；本机检测：{status}")
    out += [f"- 按单位规定使用：{line}" for line in F.unit_notes(report)]
    out.append("")
    notice = F.render_notice(report, markdown=True)  # 缺字体或无法检测字体时，醒目提示放在最前面（引用块，不占章节编号）
    if notice:
        out += [notice, ""]
    nums = iter("一二三四五六七八九十")  # 各部分按实际出现的顺序连续编号
    out.append(f"## {next(nums)}、识别结果（请核对）")
    out += [f"{i}. {s}" for i, s in enumerate(report.get("summary", []), start=1)] or ["（无）"]
    out.append("")
    out.append(f"## {next(nums)}、页面设置")
    out += [f"- {s}" for s in report["page"]] or ["- 原设置已符合，未改动"]
    out.append("")
    out.append(f"## {next(nums)}、段落与文字")
    out.append("| 角色 | 段数 | 段落格式 | 字体 | 字号 | 颜色 |")
    out.append("|---|---|---|---|---|---|")
    for role, b in report["roles"].items():
        out.append(f"| {role} | {b.get('段数', 0)} | {b.get('段落格式', 0)} | {b.get('字体', 0)} | {b.get('字号', 0)} | {b.get('颜色', 0)} |")
    out.append("")
    out.append("表中数字为实际发生改动的段数。")
    out.append("")
    if report.get("notes"):
        out.append(f"## {next(nums)}、其他处理")
        out += [f"- {n}" for n in report["notes"]]
        out.append("")
    if report.get("untouched") or report.get("warnings"):
        out.append(f"## {next(nums)}、未处理和提示")
        out += [f"- {n}" for n in report.get("untouched", []) + report.get("warnings", [])]
        out.append("")
    chk = report.get("check")
    if chk:
        s = chk["summary"]
        out.append(f"## {next(nums)}、修复后检查")
        out.append(f"P0 {s['P0']} 项，P1 {s['P1']} 项，P2 {s['P2']} 项。" + ("可以交付。" if s["P0"] == 0 else "仍有 P0，需人工处理。"))
        for it in chk["issues"]:
            loc = f"第 {it['line']} 段" if it.get("line") else "全文"
            out.append(f"- [{it['level']}] {it['message']}（{loc}：{it['excerpt']}）建议：{it['suggestion']}")
    return "\n".join(out) + "\n"


def main(argv: Optional[Sequence[str]] = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    parser = argparse.ArgumentParser(description="按 GB/T 9704-2012 修复已有 docx 的格式，输出新文件和修复报告。",
                                     epilog="示例：python scripts/fix_docx.py 旧稿.docx -o 旧稿-修复.docx")
    parser.add_argument("file", help="待修复的 .docx（不会被修改）")
    parser.add_argument("-o", "--output", help="输出路径；默认“原文件名-修复.docx”，已存在时加 -v2、-v3")
    parser.add_argument("--report", help="修复报告路径；默认与输出文件同名加“-修复报告.md”")
    parser.add_argument("--layout", help="版式：红头、信函、纪要、白头、材料、简报；不写时自动判断")
    parser.add_argument("--seal", choices=["是", "否"], help="署名是否按加盖印章排列；不写时法定公文按加盖处理")
    parser.add_argument("--json", action="store_true", help="以 JSON 输出报告")
    parser.add_argument("--allow", metavar="DOCX-050", help="修复后检查时按本单位制度保留的检查项（逗号分隔），命中降为 P2")
    parser.add_argument("--unit-fonts", metavar="角色=字体名,…", help=F.UNIT_FONTS_HELP)
    args = parser.parse_args(argv)
    try:
        import docx  # noqa: F401
    except ImportError:
        print("缺少 python-docx。请先安装：pip install -U python-docx（需要 1.2.0 及以上版本）")
        return 2
    import gw_docx as D
    from check_docx import LAYOUT_NAMES, check_document

    src = Path(args.file)
    if not src.is_file() or src.suffix.lower() != ".docx":
        print(f"错误：找不到 docx 文件 {src}")
        return 2
    if args.layout and args.layout not in LAYOUT_NAMES:
        print(f"错误：无法识别的版式：{args.layout}。可选：{'、'.join(LAYOUT_NAMES)}")
        return 2
    out = Path(args.output) if args.output else src.with_name(src.stem + "-修复.docx")
    if out.is_dir():  # -o 给的是已有文件夹：在文件夹内按默认名输出
        out = out / (src.stem + "-修复.docx")
    if out.resolve() == src.resolve():
        out = src.with_name(src.stem + "-修复.docx")
    final = D.unique_path(out)
    seal = None if args.seal is None else args.seal == "是"
    try:
        unit_fonts = F.parse_unit_fonts(args.unit_fonts, F.load_params())
    except ValueError as exc:
        print(f"错误：{exc}")
        return 2
    try:
        report = fix_document(src, final, layout=args.layout, seal=seal, unit_fonts=unit_fonts)
    except Exception as exc:
        print(f"错误：无法处理 {src}：{exc}")
        return 2
    report["check"] = check_document(final, layout=report["layout"], allow=args.allow, unit_fonts=unit_fonts)
    rpath = Path(args.report) if args.report else final.with_name(final.stem + "-修复报告.md")
    rpath = D.unique_path(rpath)
    rpath.parent.mkdir(parents=True, exist_ok=True)  # --report 指定的文件夹不存在时先建好
    rpath.write_text(render_report(report), encoding="utf-8")
    report["report"] = str(rpath)
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        s = report["check"]["summary"]
        notice = F.render_notice(report)  # 缺字体或无法检测字体时，醒目提示放在最前面
        if notice:
            print(notice + "\n")
        print(f"已输出：{final}\n修复报告：{rpath}\n修复后检查：P0 {s['P0']} 项，P1 {s['P1']} 项，P2 {s['P2']} 项")
    return 0


if __name__ == "__main__":
    sys.exit(main())
