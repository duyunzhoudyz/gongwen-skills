#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把核稿意见写成 Word 批注\u3002

读取 gongwen-review 输出的\u201c核稿意见.json\u201d\uff08docs/开发规范.md 第 8.2 节\uff09\uff0c按 excerpt 在 docx 中定位\uff0c
用 python-docx 1.2 的 add_comment 添加批注\u3002check_docx.py 或 gongwen_lint.py 的 --json 输出同样可用\u3002
找不到位置的条目列入报告\uff0c不写入文档\u3002输出新文件\uff0c不修改原文件\u3002

退出码\uff1a0 成功\uff08含部分未定位\uff09\uff1b2 输入错误\u3001缺少依赖或 python-docx 版本过低\u3002
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parent))

MIN_VERSION = (1, 2, 0)
SPACE_RE = re.compile(r"[\s　]")


def docx_version() -> Tuple[int, ...]:
    import docx
    raw = getattr(docx, "__version__", "0")
    nums = [int(x) for x in re.findall(r"\d+", str(raw))[:3]]
    return tuple(nums + [0] * (3 - len(nums)))


def supports_comments() -> bool:
    from docx.document import Document as _Doc
    return hasattr(_Doc, "add_comment")


def load_issues(path: Path) -> List[dict]:
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    issues = data.get("issues") if isinstance(data, dict) else data
    if not isinstance(issues, list):
        raise ValueError("核稿意见 JSON 中缺少 issues 数组")
    return [it for it in issues if isinstance(it, dict)]


def comment_lines(issue: dict) -> List[str]:
    head = "｜".join(str(issue[k]) for k in ("level", "gate", "category") if issue.get(k))
    first = f"【{head}】{issue.get('message', '')}" if head else str(issue.get("message", ""))
    lines = [first.strip()]
    if issue.get("suggestion"):
        lines.append(f"建议：{issue['suggestion']}")
    if issue.get("basis"):
        lines.append(f"依据：{issue['basis']}")
    return lines


def _find_exact(text: str, excerpt: str, used: List[Tuple[int, int]]) -> Tuple[Optional[Tuple[int, int]], int]:
    hits = []
    start = text.find(excerpt)
    while start != -1:
        hits.append((start, start + len(excerpt)))
        start = text.find(excerpt, start + 1)
    for h in hits:
        if not any(h[0] < u[1] and u[0] < h[1] for u in used):
            return h, len(hits)
    return None, len(hits)


def _find_loose(text: str, excerpt: str, used: List[Tuple[int, int]]) -> Tuple[Optional[Tuple[int, int]], int]:
    """忽略空格\uff08含全角空格\uff09匹配\uff0c返回原文中的区间\u3002"""
    keep = [i for i, ch in enumerate(text) if not SPACE_RE.match(ch)]
    compact = "".join(text[i] for i in keep)
    target = SPACE_RE.sub("", excerpt)
    if not target:
        return None, 0
    found, count = _find_exact(compact, target, [])
    if found is None:
        return None, count
    s, e = found
    span = (keep[s], keep[e - 1] + 1)
    if any(span[0] < u[1] and u[0] < span[1] for u in used):
        return None, count
    return span, count


def annotate(src: Path, issues: List[dict], dst: Path, *, author: str = "gongwen-review",
             initials: str = "GR") -> dict:
    import gw_docx as D
    from docx import Document

    doc = Document(str(src))
    paras = list(D.iter_all_paragraphs(doc))
    texts = [D.direct_text(p) for p in paras]
    used: Dict[int, List[Tuple[int, int]]] = {}
    located: List[dict] = []
    unlocated: List[dict] = []
    for n, issue in enumerate(issues, start=1):
        excerpt = str(issue.get("excerpt") or "").strip()
        brief = {"序号": n, "level": issue.get("level", ""), "excerpt": excerpt, "message": issue.get("message", "")}
        if not excerpt:
            unlocated.append(dict(brief, reason="excerpt 为空"))
            continue
        hit = None
        total = 0
        for finder in (_find_exact, _find_loose):  # 先逐字匹配，找不到再忽略空格匹配
            total = 0
            for idx, text in enumerate(texts):
                span, count = finder(text, excerpt, used.get(idx, []))
                total += count
                if span and hit is None:
                    hit = (idx, span)
            if hit or total:
                break
        if hit is None:
            reason = "原文中找不到这段文字\uff08可能跨段\u3001已改动\uff0c或在页眉页脚\u3001文本框中\uff09"
            if total:
                reason = "这段文字的每一处都已有批注"
            unlocated.append(dict(brief, reason=reason))
            continue
        idx, (start, end) = hit
        runs = D.runs_for_range(paras[idx], start, end)
        if runs is None:
            unlocated.append(dict(brief, reason="所在文字含域代码\u3001图片或超链接\uff0c无法拆分定位"))
            continue
        lines = comment_lines(issue)
        comment = doc.add_comment(list(runs), text=lines[0], author=author, initials=initials)
        for extra in lines[1:]:
            comment.add_paragraph(extra)
        used.setdefault(idx, []).append((start, end))
        texts[idx] = D.direct_text(paras[idx])
        item = dict(brief, paragraph=idx + 1)
        if total > 1:
            item["note"] = f"原文有 {total} 处相同文字，已批注第一处未批注的位置"
        located.append(item)
    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(dst))
    return {"input": str(src), "output": str(dst), "author": author, "total": len(issues),
            "located": located, "unlocated": unlocated}


def render_report(report: dict) -> str:
    out = ["# 批注报告", "", f"- 原文件：{report['input']}（未修改）", f"- 新文件：{report['output']}",
           f"- 批注作者：{report['author']}",
           f"- 核稿意见 {report['total']} 条：已写入批注 {len(report['located'])} 条，未定位 {len(report['unlocated'])} 条", ""]
    if report["unlocated"]:
        out.append("## 未定位的意见\uff08未写入文档\uff0c请人工处理\uff09")
        out.append("")
        out.append("| 序号 | 级别 | 原文片段 | 问题 | 原因 |")
        out.append("|---|---|---|---|---|")
        for it in report["unlocated"]:
            out.append(f"| {it['序号']} | {it['level']} | {it['excerpt']} | {it['message']} | {it['reason']} |")
        out.append("")
    notes = [it for it in report["located"] if it.get("note")]
    if notes:
        out.append("## 多处匹配")
        out += [f"- 第 {it['序号']} 条“{it['excerpt']}”：{it['note']}" for it in notes]
    return "\n".join(out) + "\n"


def main(argv: Optional[Sequence[str]] = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    parser = argparse.ArgumentParser(description="把核稿意见 JSON 写成 Word 批注\uff0c输出新文件和批注报告\u3002",
                                     epilog="示例\uff1apython scripts/annotate_docx.py 成稿.docx 核稿意见.json")
    parser.add_argument("file", help="要加批注的 .docx\uff08不会被修改\uff09")
    parser.add_argument("issues", help="核稿意见.json\uff08开发规范 8.2\uff09或机检 JSON")
    parser.add_argument("-o", "--output", help="输出路径\uff1b默认\u201c原文件名-批注.docx\u201d\uff0c已存在时加 -v2\u3001-v3")
    parser.add_argument("--author", default="gongwen-review", help="批注作者名\uff0c默认 gongwen-review")
    parser.add_argument("--initials", default="GR", help="作者缩写\uff0c默认 GR")
    parser.add_argument("--report", help="批注报告路径\uff1b默认与输出文件同名加\u201c-批注报告.md\u201d")
    parser.add_argument("--json", action="store_true", help="以 JSON 输出报告")
    args = parser.parse_args(argv)
    try:
        import docx  # noqa: F401
    except ImportError:
        print("缺少 python-docx\u3002请先安装\uff1apip install -U python-docx\uff08需要 1.2.0 及以上版本\uff09")
        return 2
    if not supports_comments() or docx_version() < MIN_VERSION:
        ver = ".".join(str(x) for x in docx_version())
        print(f"python-docx 版本过低（当前 {ver}），写批注需要 1.2.0 及以上版本：pip install -U python-docx")
        return 2
    import gw_docx as D

    src = Path(args.file)
    if not src.is_file() or src.suffix.lower() != ".docx":
        print(f"错误：找不到 docx 文件 {src}")
        return 2
    try:
        issues = load_issues(Path(args.issues))
    except (OSError, ValueError) as exc:
        print(f"错误：核稿意见读取失败：{exc}")
        return 2
    out = Path(args.output) if args.output else src.with_name(src.stem + "-批注.docx")
    if out.is_dir():  # -o 给的是已有文件夹：在文件夹内按默认名输出
        out = out / (src.stem + "-批注.docx")
    final = D.unique_path(out)
    try:
        report = annotate(src, issues, final, author=args.author, initials=args.initials)
    except Exception as exc:  # 文件损坏、不是 docx 等
        print(f"错误：无法处理 {src}：{exc}")
        return 2
    rpath = D.unique_path(Path(args.report) if args.report else final.with_name(final.stem + "-批注报告.md"))
    rpath.parent.mkdir(parents=True, exist_ok=True)  # --report 指定的文件夹不存在时先建好
    rpath.write_text(render_report(report), encoding="utf-8")
    report["report"] = str(rpath)
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"已输出：{final}\n批注报告：{rpath}\n已写入批注 {len(report['located'])} 条，未定位 {len(report['unlocated'])} 条")
        for it in report["unlocated"]:
            print(f"  未定位：第 {it['序号']} 条“{it['excerpt']}”，{it['reason']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
