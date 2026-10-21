# -*- coding: utf-8 -*-
"""文字核对：输入全文与文档全文规范化后双向、按顺序逐字比对（只依赖标准库）。

多字（文档里有、输入里没有）、少字（输入里有、文档里没有）、乱序（位置不一致）、重复（同一段文字排了两次）都算不一致。
build_docx.py --keep-text 保存前用它核对，不一致就中止；抽取（extract_docx.py）后用它核对“版式信息 + 正文”是否等于原文。
规范化：去掉全部空白（含全角空格、制表符、换行）。核对通过前，不得对用户声称“已逐字核对”。
"""
from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field
from typing import List, Sequence

# 文档里由版式信息写出的固定标签字，核对“多出”的文字时不计（如“签发人：”“抄送：”“某日印发”“附件1”）
LABEL_WORDS = ("签发人：", "抄送：", "附件：", "总第", "★", "印发", "文件", "附件", "第", "号", "期", "编")


def norm(text: str) -> str:
    """规范化：去掉全部空白字符。"""
    return re.sub(r"\s+", "", text or "")


@dataclass
class TextCheck:
    ok: bool = True
    chars: int = 0
    problems: List[str] = field(default_factory=list)

    def summary(self, limit: int = 5) -> str:
        return "；".join(self.problems[:limit]) + (f"；……共 {len(self.problems)} 处" if len(self.problems) > limit else "")


def _pieces(a: str, b: str):
    """a（输入）与 b（文档）的字符级差异：返回 (少的文字列表, 多的文字列表)。"""
    missing: List[str] = []
    extra: List[str] = []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if op in ("delete", "replace") and a[i1:i2]:
            missing.append(a[i1:i2])
        if op in ("insert", "replace") and b[j1:j2]:
            extra.append(b[j1:j2])
    return missing, extra


def _covered(piece: str, values: Sequence[str]) -> bool:
    """piece 是否能完全由版式信息的值和固定标签字依次拼成。"""
    words = sorted({v for v in list(values) + list(LABEL_WORDS) if v}, key=len, reverse=True)
    reach = [False] * (len(piece) + 1)
    reach[0] = True
    for i in range(len(piece)):
        if not reach[i]:
            continue
        for w in words:
            if piece.startswith(w, i):
                reach[i + len(w)] = True
    return reach[len(piece)]


def compare_text(doc_texts: Sequence[str], input_texts: Sequence[str], *, allowed: Sequence[str] = (),
                 movable: Sequence[str] = ()) -> TextCheck:
    """比对文档全文与输入全文。

    doc_texts：文档各段落文字（按文档顺序，含表格单元格）；input_texts：输入各行、表格单元格文字（按输入顺序）。
    allowed：版式信息中提供的值（文档里由它们写出的文字不算“多出”，其他多出的文字报错）。
    movable：允许换位置的文字（版头要素、版记行，排版时按版式移到版头、版记位置）。
    """
    inp = [t for t in (norm(x) for x in input_texts) if t]
    doc = [t for t in (norm(x) for x in doc_texts) if t]
    allow = [a for a in (norm(x) for x in allowed) if a]
    mov = {norm(x) for x in movable if norm(x)}
    res = TextCheck(chars=sum(len(t) for t in inp))
    missing: List[str] = []
    extra: List[str] = []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, inp, doc, autojunk=False).get_opcodes():
        if op == "equal":
            continue
        ri, rd = "".join(inp[i1:i2]), "".join(doc[j1:j2])
        if ri == rd:
            continue  # 只是分段不同（如标题分行），文字一致
        m, e = _pieces(ri, rd)
        missing += m
        extra += e
    # 文字相同、位置不同：版头、版记要素按版式换位置，允许；其他是乱序
    for text in list(missing):
        if text in extra:
            missing.remove(text)
            extra.remove(text)
            if text not in mov:
                res.problems.append(f"乱序：“{text[:30]}”在文档中的位置与输入不一致")
    for text in missing:
        res.problems.append(f"少字：输入中的“{text[:30]}”文档里没有")
    inp_set = set(inp)
    for text in extra:
        if text in inp_set:
            res.problems.append(f"重复：“{text[:30]}”在文档中多排了一次")
        elif not _covered(text, allow):
            res.problems.append(f"多字：文档里多出输入中没有的“{text[:30]}”")
    res.ok = not res.problems
    return res


def compare_chars(original_texts: Sequence[str], rebuilt_texts: Sequence[str], *, ignore: Sequence[str] = ()) -> TextCheck:
    """按字统计核对（不看顺序）：原文全文与“版式信息加正文”重组后的全文，各字出现次数应一致，多了或少了都报出。

    抽取（extract_docx.py）用：版式信息里的要素已从正文移出，顺序与原文不同，所以只核对文字有没有多、少、重复；
    ignore：两边都去掉的固定标签字和记号（如“签发人：”“抄送：”“【”“】”）。"""
    from collections import Counter

    def flat(texts: Sequence[str]) -> str:
        out = norm("".join(texts))
        for word in ignore:
            out = out.replace(word, "")
        return out

    a, b = Counter(flat(original_texts)), Counter(flat(rebuilt_texts))
    res = TextCheck(chars=sum(a.values()))
    lost, gained = a - b, b - a
    if lost:
        res.problems.append("少字：原文有、版式信息和正文里没有的字：" + "".join(f"{c}×{n} " for c, n in lost.most_common(12)).strip())
    if gained:
        res.problems.append("多字或重复：版式信息和正文里多出的字：" + "".join(f"{c}×{n} " for c, n in gained.most_common(12)).strip())
    res.ok = not res.problems
    return res
