#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从已有 docx 抽取全文和版式信息\uff0c供 build_docx.py 重排\uff08已有 docx 的版头\u3001版记\u3001附件等结构有缺失或错误时用\uff09\u3002

输出两个文件\uff1a
- 正文-<原名>.txt\uff1a按段落抽取的全文\u3002版头\u3001版记要素\uff08见下\uff09已移入版式信息\uff0c不在正文中\uff1b正文表格转成管道表格\uff1b
  附件页\uff08\u201c附件\u201d\u201c附件1\u201d标识加标题\uff09转成\u201c【附件1】标题\u201d分界行\uff1b版记表格和落在附件之后的抄送\u3001印发行移到主文末尾\u3002
- 版式信息-<原名>.json\uff1a版式\u3001份号\u3001密级\u3001紧急程度\u3001发文机关标志\u3001令号\u3001发文字号\u3001签发人\u3001抄送\uff08原句\uff09\u3001印发机关\u3001印发日期\u3002
  识别不出的不写\uff0c由 build_docx.py 写【待补】\u3002主送机关\u3001标题\u3001署名\u3001成文日期\u3001附注\u3001附件说明仍留在正文中\uff0c由 build 识别\uff0c
  避免版式信息覆盖原文写法\u3002

不改任何文字\uff1a文字只做搬运\uff08段内制表符换成全角空格\uff0c软换行拆成两行\uff09\u3002之后用 build_docx.py --keep-text 重排\uff0c
build 会逐行核对文字\u3002退出码\uff1a0 成功\uff1b2 输入错误或缺少依赖\u3002
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import List, Optional, Sequence

sys.path.insert(0, str(Path(__file__).resolve().parent))

from gw_verify import compare_chars  # noqa: E402
from gw_layout import (ATTACH_MARK_RE, CC_RE, PRINT_RE, SENT_PUNCT, SIGNER_RE, doc_kind, has_brief_header,  # noqa: E402
                       is_draft_note_line, parse_lines)

_DUN = "\u3001"  # 顿号；f-string 表达式内不能出现反斜杠（Python < 3.12）
_LAYOUT_UNKNOWN = "未判定\uff08build 按文种决定\uff09"

STRIP_HEAD = ("copy_no", "secret", "urgency", "issuer", "lingno", "docno", "signer")
STRIP_TAIL = ("cc", "print")
STRIP_LAYOUTS = ("红头", "信函", "纪要", "命令")
TAIL_LAYOUTS = ("红头", "白头", "纪要", "命令")


def _norm(text: str) -> List[str]:
    """一个段落的文字\uff1a软换行拆成多行\uff0c段内制表符换成全角空格\uff0c去行尾空白\uff08行首空格保留\uff09\u3002"""
    out = []
    for part in text.replace("\r", "").split("\n"):
        part = re.sub(r"(?<=\S)\t+(?=\S)", "　", part).replace("\t", "").rstrip()
        if part.strip():
            out.append(part)
    return out


def _cell_lines(cell_el, doc) -> List[str]:
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn
    lines: List[str] = []
    for p in cell_el.findall(qn("w:p")):
        para = Paragraph(p, doc._body)
        lines += _norm(para.text)
    return lines


def _pipe_table(tbl, doc) -> List[str]:
    from docx.oxml.ns import qn
    rows = []
    for tr in tbl._tbl.findall(qn("w:tr")):
        cells = [" <br> ".join(_cell_lines(tc, doc)).replace("|", "\\|") for tc in tr.findall(qn("w:tc"))]
        rows.append(cells)
    if not rows:
        return []
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    fmt = lambda r: "| " + " | ".join(r) + " |"  # noqa: E731
    return [fmt(rows[0]), "|" + "|".join(["---"] * width) + "|"] + [fmt(r) for r in rows[1:]]


def extract(path: Path, layout: Optional[str] = None) -> dict:
    """返回 {"lines": 正文行, "meta": 版式信息, "recognized": 识别结果, "notes": 提示, "warnings": 警告}\u3002"""
    import gw_docx as D
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(str(path))
    notes: List[str] = []
    warnings: List[str] = []
    original: List[str] = []  # 原文正文段落、表格单元格文字（【起草说明】之前），抽取后核对用
    # 1. 按文档顺序收集：(类型, 文字, 是否红色)
    seq: list = []  # ("p", text, red) / ("raw", text, False) / ("banji", text, False)
    joint_names: List[str] = []
    cut = False
    for block in doc.element.body.iterchildren():
        tag = D._local(block.tag)
        if tag == "p":
            para = Paragraph(block, doc._body)
            raw = D.direct_text(para) or para.text
            if is_draft_note_line(raw.strip()):
                cut = True
                notes.append("遇到【起草说明】\uff0c其后内容未抽取")
                break
            red = D.para_is_red(para, doc)
            original.append(raw)
            seq += [("p", t, red) for t in _norm(raw)]
        elif tag == "tbl":
            tbl = Table(block, doc._body)
            original += [Paragraph(p, doc._body).text for p in block.iter(D.qn("w:p"))]
            mark = D.table_mark(block)
            if mark == D.JOINT_HEAD_MARK:
                first = block.findall(".//" + D.qn("w:tc"))[0]
                joint_names = [t.strip() for t in _cell_lines(first, doc)]
            elif mark == D.JOINT_SEAL_MARK:
                for row in tbl._tbl.findall(D.qn("w:tr")):
                    for tc in row.findall(D.qn("w:tc")):
                        seq += [("p", t, False) for t in _cell_lines(tc, doc)]
            elif D.is_banji_table(block):
                for tr in block.findall(D.qn("w:tr")):
                    cells = [" ".join(_cell_lines(tc, doc)) for tc in tr.findall(D.qn("w:tc"))]
                    seq.append(("banji", " ".join(c for c in cells if c), False))
            else:
                seq += [("raw", t, False) for t in _pipe_table(tbl, doc)]
    # 2. 拆主文、附件页；版记行归到主文末尾
    main: list = []
    banji: list = []
    secs: list = []  # [label, title, [行]]
    i = 0
    while i < len(seq):
        kind, text, red = seq[i]
        m = ATTACH_MARK_RE.match(text.strip()) if kind == "p" else None
        if kind == "banji":
            banji.append((text, False))
        elif m:
            title = ""
            if i + 1 < len(seq) and seq[i + 1][0] == "p" and not ATTACH_MARK_RE.match(seq[i + 1][1].strip()) \
                    and len(seq[i + 1][1].strip()) <= 60 and seq[i + 1][1].strip()[-1] not in SENT_PUNCT + "\uff1a:":
                title = seq[i + 1][1].strip()
                i += 1
            secs.append([m.group(1) or "", title, []])
        elif secs:
            if kind == "p" and (CC_RE.match(text.strip()) or PRINT_RE.match(text.strip().replace("　", " "))):
                banji.append((text, red))
                notes.append(f"附件之后的版记行“{text.strip()[:20]}”已移到主文末尾")
            else:
                secs[-1][2].append(text)
        else:
            main.append((text, red))
        i += 1
    texts = [t for t, _ in main] + [t for t, _ in banji]
    reds = [r for _, r in main] + [False] * len(banji)

    # 3. 识别版式
    sec0 = doc.sections[0]
    letter_head = D.letter_head_text(sec0)
    meta: dict = {}
    recognized: List[str] = []
    brief = has_brief_header(texts, reds)
    parsed = parse_lines(texts, red=reds)
    roles = parsed.roles()
    kind, _ = doc_kind(str(parsed.fields.get("标题") or ""))
    if layout:
        lay = layout
    elif letter_head:
        lay = "信函"
    elif brief:
        lay = "简报"
    elif "issuer" in roles:
        lay = "纪要" if parsed.find("issuer")[0].text.endswith("纪要") else ("命令" if "lingno" in roles else "红头")
    elif joint_names:
        lay = "红头"
    else:
        lay = None
    eff = lay or ("白头" if kind == "法定" else "材料")
    strip_roles: tuple = ()
    if lay == "简报":
        warnings.append("识别为简报\uff0c版头版记未移入版式信息\uff1b简报用 fix_docx.py --keep-text 处理已排好的 docx")
    else:
        if eff in STRIP_LAYOUTS:
            strip_roles += STRIP_HEAD
        if eff in TAIL_LAYOUTS:
            strip_roles += STRIP_TAIL
    if lay:
        meta["版式"] = lay
    drop: set = set()
    seen_docno = False
    signers: List[str] = []
    for it in parsed.items:
        if it.role not in strip_roles or it.src < 0:
            continue
        if it.role == "docno":
            if seen_docno:
                continue  # 第二个文号行留在正文，由 build 提示
            seen_docno = True
        drop.add(it.src)
        if it.role in ("docno", "signer"):
            m = SIGNER_RE.search(it.text)
            if m:
                signers.append(m.group(1).strip())
    f = parsed.fields
    if lay != "简报":
        if "issuer" in strip_roles and joint_names:
            meta["发文机关标志"] = joint_names
        elif "issuer" in strip_roles and f.get("发文机关标志"):
            meta["发文机关标志"] = f["发文机关标志"]
        elif letter_head and eff == "信函":
            meta["发文机关标志"] = letter_head
        for key in ("份号", "密级", "紧急程度", "令号", "发文字号"):
            if key in f and any(r in strip_roles for r in STRIP_HEAD):
                meta[key] = f[key]
        if signers:
            meta["签发人"] = signers if len(signers) > 1 else signers[0]
        cc = next((it.text for it in parsed.items if it.role == "cc" and it.src in drop), None)
        if cc:
            meta["抄送"] = cc.strip()
        if f.get("印发机关") and "print" in strip_roles:
            meta["印发机关"] = f["印发机关"]
            meta["印发日期"] = f["印发日期"]
    if lay in ("红头", "纪要", "命令", "信函") and "发文机关标志" not in meta:
        notes.append("未识别到发文机关标志\uff0cbuild 将写【待补】\uff0c请用户补齐")
    for k, label in (("标题", "标题"), ("主送机关", "主送机关"), ("发文机关署名", "落款署名"), ("成文日期", "成文日期"),
                     ("附注", "附注"), ("附件", "附件说明")):
        if f.get(k):
            recognized.append(f"{label}：{f[k] if not isinstance(f[k], list) else _DUN.join(f[k])}（留在正文，由 build 识别）")
    recognized += [f"{k}：{_DUN.join(v) if isinstance(v, list) else v}（移入版式信息）" for k, v in meta.items() if k != "版式"]
    if secs:
        recognized.append("附件页\uff1a" + "\uff1b".join(f"附件{s[0]} {s[1] or '【无标题】'}" for s in secs))
    warnings += [w for w in parsed.warnings if "空格缩进" not in w]

    # 4. 组装正文行：按 parse 时的下标去掉已移入版式信息的行
    lines: List[str] = [t for k, (t, _) in enumerate(main) if k not in drop]
    tail = [t for k, (t, _) in enumerate(banji, start=len(main)) if k not in drop]
    lines += tail
    for label, title, body in secs:
        lines.append(f"【附件{label}】{title}")
        lines += body
    # 5. 页眉、页脚里的页码和杂项不带入重排稿（页码由 build 在页脚按标准重设），如实告知
    hf: List[str] = []
    for idx, sec in enumerate(doc.sections, start=1):
        for name, part in (("页眉", sec.header), ("双页页眉", sec.even_page_header), ("页脚", sec.footer),
                           ("双页页脚", sec.even_page_footer), ("首页页眉", sec.first_page_header),
                           ("首页页脚", sec.first_page_footer)):
            if part.is_linked_to_previous or (name == "首页页眉" and lay == "信函" and letter_head):
                continue
            summary = D.part_summary(part._element)
            if summary:
                hf.append(("" if len(doc.sections) == 1 else f"第 {idx} 节") + f"{name}：{summary}")
    if hf:
        notes.append("原页眉\u3001页脚中的内容不带入重排稿\uff08公文不用页眉\uff0c页码由 build 在页脚按标准重设\uff09\uff1a" + "\uff1b".join(hf))
    # 6. 核对：版式信息加正文（还原成文字）应与原文各字一致，没有多字、少字、重复；不一致的先查明再重排
    if letter_head and "发文机关标志" in meta and lay == "信函":
        original.append(letter_head)
    recon: List[str] = [v for key, val in meta.items() if key != "版式"
                        for v in (val if isinstance(val, list) else [val])]
    for ln in lines:
        if re.match(r"^[|]", ln.strip()):
            recon += [c.replace("<br>", "").strip() for c in ln.strip().strip("|").split("|")
                      if not re.fullmatch(r"[\s:-]*", c)]
        else:
            recon.append(ln)
    chk = compare_chars(original, recon, ignore=("签发人\uff1a", "签发人:", "抄送\uff1a", "抄送:", "印发", "文件", "【", "】", "|", "-",
                                                 "<br>", "&nbsp;"))
    if not chk.ok:
        warnings.append("抽取核对发现版式信息加正文与原文的文字不一致\uff0c请先查明再重排\uff08不得声称已逐字核对\uff09\uff1a" + chk.summary())
    return {"lines": lines, "meta": meta, "recognized": recognized, "notes": notes, "warnings": warnings,
            "attachments": len(secs), "layout": lay,
            "text_check": {"ok": chk.ok, "chars": chk.chars, "problems": chk.problems}}


def main(argv: Optional[Sequence[str]] = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    parser = argparse.ArgumentParser(description="从已有 docx 抽取全文和版式信息\uff0c供 build_docx.py 重排\u3002",
                                     epilog="示例\uff1apython scripts/extract_docx.py 旧稿.docx --layout 红头 && "
                                            "python scripts/build_docx.py 正文-旧稿.txt --meta 版式信息-旧稿.json --keep-text")
    parser.add_argument("file", help="待抽取的 .docx\uff08不会被修改\uff09")
    parser.add_argument("-o", "--output", help="正文输出路径\uff1b默认在 docx 同目录写\u201c正文-<原名>.txt\u201d")
    parser.add_argument("--meta-output", help="版式信息输出路径\uff1b默认与正文同目录\u201c版式信息-<原名>.json\u201d")
    parser.add_argument("--layout", help="版式\uff1a红头\u3001信函\u3001纪要\u3001命令\u3001白头\u3001材料\uff08不写时按版头判断\uff1b"
                                         "版头缺失但应排红头的\uff0c指定红头\uff0cbuild 会把缺的要素写成【待补】\uff09")
    parser.add_argument("--json", action="store_true", help="以 JSON 输出结果")
    args = parser.parse_args(argv)
    try:
        import docx  # noqa: F401
    except ImportError:
        print("缺少 python-docx\u3002请先安装\uff1apip install -U python-docx\uff08需要 1.2.0 及以上版本\uff09")
        return 2
    import gw_docx as D
    src = Path(args.file)
    if not src.is_file() or src.suffix.lower() != ".docx":
        print(f"错误：找不到 .docx 文件 {src}")
        return 2
    try:
        res = extract(src, args.layout)
    except Exception as exc:  # noqa: BLE001 损坏的 docx 等
        print(f"错误：抽取失败：{exc}")
        return 2
    out = D.unique_path(Path(args.output) if args.output else src.with_name(f"正文-{src.stem}.txt"))
    meta_out = D.unique_path(Path(args.meta_output) if args.meta_output else out.with_name(f"版式信息-{src.stem}.json"))
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(res["lines"]) + "\n", encoding="utf-8")
    meta_out.write_text(json.dumps(res["meta"], ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    res.update({"text_file": str(out), "meta_file": str(meta_out)})
    if args.json:
        print(json.dumps({k: v for k, v in res.items() if k != "lines"}, ensure_ascii=False, indent=2))
        return 0
    print(f"已抽取：{len(res['lines'])} 行 -> {out}\n版式信息 -> {meta_out}\n版式：{res['layout'] or _LAYOUT_UNKNOWN}")
    for title, items in (("识别结果", res["recognized"]), ("提示", res["notes"]), ("警告", res["warnings"])):
        if items:
            print(title + "\uff1a\n" + "\n".join("  - " + x for x in items))
    tc = res["text_check"]
    print("抽取核对\uff1a版式信息加正文与原文各字一致\uff0c没有多字\u3001少字\u3001重复\uff08" + str(tc["chars"]) + " 字\uff09" if tc["ok"]
          else "抽取核对\uff1a不一致\uff0c见上面的警告\uff1b查明前不要声称已逐字核对")
    print("下一步\uff1apython scripts/build_docx.py <正文> --meta <版式信息> --keep-text -o <成稿.docx>\uff08内容别动时加 --keep-text\uff09")
    return 0


if __name__ == "__main__":
    sys.exit(main())
