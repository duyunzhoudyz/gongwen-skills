#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""保密标识检查\uff1a扫描 txt\u3001md\u3001docx\u3001xlsx 中的密级标识和内部标注\uff0c命中即退出码 1\u3002

用法\uff1apython secret_check.py <文件...>

docx 读取范围\uff1a正文段落\u3001表格\u3001页眉\u3001页脚\uff08含首页\u3001奇偶页\uff09\u3001文本框\u3001脚注\u3001尾注\u3001批注\uff0c
输出位置带来源标签\uff0c如\u201c[页眉(首页)]\u201d\u201c[文本框]\u201d\u201c[表格]\u201d\u3002xlsx 读取单元格文字\u3001批注\u3001
工作表名\u3001页眉页脚\u3002只依赖标准库\u3002

检出\uff1a\u201c秘密★\u201d\u201c机密★\u201d\u201c绝密★\u201d\uff08含 5年 等保密期限写法\uff09\u3001\u201c工作秘密\u201d\u201c内部资料\u201d
\u201c不得外传\u201d\u201c内部使用\u201d等\u3002\u201c保密工作\u201d\u201c涉密人员管理\u201d之类业务用语不算\u3002
输出只给位置和命中词\uff0c不复述材料内容\u3002

退出码\uff1a0 未发现\uff1b1 发现标识\uff1b2 文件无法读取\u3002
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Tuple

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
MAX_XML = 50 * 1024 * 1024

LEVEL = "绝密|机密|秘密"
STAR = "[★☆＊*]"
PATTERNS = [
    ("密级标识", re.compile("(?:%s)\\s*%s" % (LEVEL, STAR))),
    ("密级与保密期限", re.compile("(?:%s)\\s*%s?\\s*[\\d一二三四五六七八九十]+\\s*(?:年|个月)" % (LEVEL, STAR))),
    ("密级与保密期限", re.compile("(?:%s)\\s*%s?\\s*(?:长期|永久)" % (LEVEL, STAR))),
    ("保密期限", re.compile("保密期限")),
    ("内部标注", re.compile("工作秘密|内部资料|不得外传|严禁外传|内部使用|仅供内部参考|内部参阅|注意保密")),
]

Item = Tuple[str, str]  # (位置标签, 文字)


def _root(zf: zipfile.ZipFile, name: str) -> ET.Element:
    if zf.getinfo(name).file_size > MAX_XML:
        raise ValueError("部件过大\uff1a" + name)
    return ET.fromstring(zf.read(name))


def _para_text(p: ET.Element) -> str:
    """段落文字\uff1a含修订删除的文字\uff08保守起见一并检查\uff09\u3002"""
    parts = []
    for e in p.iter():
        if e.tag in (W + "t", W + "delText"):
            parts.append(e.text or "")
        elif e.tag == W + "tab":
            parts.append("\t")
    return "".join(parts)


def _paragraphs(root: ET.Element, base: str) -> List[Item]:
    """遍历部件内全部段落\uff1b文本框\u3001表格内的段落单独标注来源\u3002"""
    out: List[Item] = []
    counter = [0]

    def walk(el: ET.Element, ctx: str) -> None:
        for ch in el:
            if ch.tag == W + "txbxContent":
                walk(ch, "文本框")
            elif ch.tag == W + "tbl":
                walk(ch, ctx if ctx == "文本框" else "表格")
            elif ch.tag == W + "p":
                counter[0] += 1
                text = _para_text(ch)
                if text.strip():
                    if ctx == "正文":
                        label = "%s第%d段" % (base, counter[0]) if base == "正文" else base
                    else:
                        label = ctx if base == "正文" else base + ctx
                    out.append((label, text))
                walk(ch, ctx)
            else:
                walk(ch, ctx)

    walk(root, "正文")
    seen = set()  # mc:AlternateContent 中文本框常有两份相同内容，去重
    uniq = []
    for item in out:
        if item not in seen:
            seen.add(item)
            uniq.append(item)
    return uniq


def _hf_labels(zf: zipfile.ZipFile) -> Dict[str, str]:
    """部件路径 -> 页眉/页脚标签\uff08区分首页\u3001偶数页\uff09\u3002"""
    labels: Dict[str, str] = {}
    names = set(zf.namelist())
    if "word/document.xml" not in names:
        return labels
    rels = {}
    rp = "word/_rels/document.xml.rels"
    if rp in names:
        for r in _root(zf, rp):
            rels[r.get("Id")] = r.get("Target", "")
    kinds = {"default": "", "first": "(首页)", "even": "(偶数页)"}
    for el in _root(zf, "word/document.xml").iter():
        if el.tag in (W + "headerReference", W + "footerReference"):
            tgt = rels.get(el.get(R + "id"), "").lstrip("/")
            path = tgt if tgt.startswith("word/") else "word/" + tgt
            word = "页眉" if el.tag == W + "headerReference" else "页脚"
            labels.setdefault(path, word + kinds.get(el.get(W + "type", "default"), ""))
    return labels


def extract_docx(path: Path) -> List[Item]:
    """抽取 docx 全部文字\uff0c返回 [(位置标签, 文字)]\u3002"""
    items: List[Item] = []
    with zipfile.ZipFile(path) as zf:
        labels = _hf_labels(zf)
        for name in zf.namelist():
            if not re.fullmatch(r"word/[^/]+\.xml", name):
                continue
            base = name[5:-4]
            if name == "word/document.xml":
                items.extend(_paragraphs(_root(zf, name), "正文"))
            elif re.fullmatch(r"(header|footer)\d*", base):
                default = "页眉" if base.startswith("header") else "页脚"
                items.extend(_paragraphs(_root(zf, name), labels.get(name, default)))
            elif base in ("footnotes", "endnotes", "comments"):
                tag = {"footnotes": "脚注", "endnotes": "尾注", "comments": "批注"}[base]
                items.extend(_paragraphs(_root(zf, name), tag))
    return items


def _stext(el: ET.Element) -> str:
    return "".join(t.text or "" for t in el.iter(S + "t"))


def extract_xlsx(path: Path) -> List[Item]:
    items: List[Item] = []
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            if not name.startswith("xl/") or not name.endswith(".xml"):
                continue
            if name == "xl/workbook.xml":
                for sh in _root(zf, name).iter(S + "sheet"):
                    items.append(("工作表名", sh.get("name", "")))
            elif name == "xl/sharedStrings.xml":
                for si in _root(zf, name).iter(S + "si"):
                    items.append(("单元格", _stext(si)))
            elif re.fullmatch(r"xl/comments\d*\.xml", name):
                for c in _root(zf, name).iter(S + "comment"):
                    items.append(("批注", _stext(c)))
            elif re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name):
                sheet = name.rsplit("/", 1)[1][:-4]
                for el in _root(zf, name).iter():
                    tag = el.tag.replace(S, "")
                    if tag in ("oddHeader", "evenHeader", "firstHeader"):
                        items.append(("页眉(%s)" % sheet, el.text or ""))
                    elif tag in ("oddFooter", "evenFooter", "firstFooter"):
                        items.append(("页脚(%s)" % sheet, el.text or ""))
                    elif tag == "is":  # 内联字符串
                        items.append(("单元格(%s)" % sheet, _stext(el)))
    return items


def extract(path: Path) -> List[Item]:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return extract_docx(path)
    if suffix == ".xlsx":
        return extract_xlsx(path)
    data = path.read_bytes()
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("gb18030", errors="replace")
    return [("第%d行" % i, line) for i, line in enumerate(text.splitlines(), 1) if line.strip()]


def scan(items: List[Item]) -> List[Tuple[str, str, str]]:
    """返回 [(位置, 类型, 命中词)]\u3002"""
    hits = []
    for label, text in items:
        flat = re.sub(r"\s+", " ", text)
        for kind, rx in PATTERNS:
            for m in rx.finditer(flat):
                hits.append((label, kind, m.group(0)))
    return hits


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    ap = argparse.ArgumentParser(description="检查文件中的保密标识\uff08含页眉页脚\u3001文本框\u3001表格\uff09\uff0c命中退出码 1\u3002")
    ap.add_argument("files", nargs="+", help="txt\u3001md\u3001docx\u3001xlsx 文件")
    args = ap.parse_args(argv)
    code = 0
    for f in args.files:
        path = Path(f)
        try:
            items = extract(path)
        except (OSError, zipfile.BadZipFile, ET.ParseError, ValueError, KeyError) as exc:
            print("[读取失败] %s\uff1a%s" % (path.name, exc))
            code = 2
            continue
        hits = scan(items)
        if not hits:
            print("[未发现保密标识] " + path.name)
            continue
        if code == 0:
            code = 1
        print("[发现保密标识] %s\uff1a共 %d 处\uff0c请立即停止处理并按保密红线告知用户" % (path.name, len(hits)))
        seen = set()
        for label, kind, word in hits:
            if (label, word) not in seen:
                seen.add((label, word))
                print("  [%s] %s\uff1a%s" % (label, kind, word))
    return code


if __name__ == "__main__":
    sys.exit(main())
