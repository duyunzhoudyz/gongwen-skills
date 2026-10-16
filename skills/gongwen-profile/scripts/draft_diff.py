# -*- coding: utf-8 -*-
"""花脸稿对比：按句对齐原稿和领导改稿，抽取删除、新增、替换（gongwen-profile）。

用法：
    python scripts/draft_diff.py <原稿> <改稿> [--json] [--min-sim 0.45] [--output 文件]
    python scripts/draft_diff.py <带修订痕迹的.docx> [--author N] [--json] [--output 文件]

两种输入：
    1. 两份文件：原稿和改稿，支持 .txt、.md、.docx。docx 中如有未接受的修订，按"接受全部修订"后的文本比对。
    2. 一份带修订痕迹的 docx：用标准库解析 word/document.xml，拒绝全部修订得到原稿，接受全部修订得到改稿；
       w:ins、w:moveTo 为新增，w:del、w:moveFrom 为删除。同时读取 word/comments.xml 中的批注及其锚定文字。
       --author N 只采用第 N 位修订人的修订（编号见输出的修订人清单），其他人的修订按拒绝处理。

对齐方法：
    1. 两稿各按句切分（句号、问号、感叹号、分号和段落），去掉空白后比较；
       比较时剥离句首的层次序号（"一、""（一）""1.""（1）"），序号变化不算改动，只在概况中计数；
    2. 先用完全相同的句子作锚点，再在锚点之间按相似度做动态规划对齐，允许一对一、一拆二、二合一；
    3. 对齐上的句子在字一级比对，得出删除、新增、替换片段；
    4. 未对齐的句子为"删除句""新增句"；删除句与新增句相同或高度相似的，记为"调整顺序"
       （整节挪动时，标题文字相同、序号不同的也记为调整顺序）。

标签是启发式提示，供归纳偏好时参考，需要人工判断：评价词删减、换词、数字改动（原有数字改成另一个数字，
可能是事实更正）、改为具体数字（"大幅增加""部分""近千人"等概括说法换成数字，不算事实更正）、删去数字、
删除含数据句、新增数据、删铺垫、删表态、语气改为要求、数据前移、句子压缩、拆句、合句、调整顺序、标点。

保密：比对前先查密级标识（"秘密★1年"等）、"内部资料""工作秘密""严禁外传""内部参阅"等内部标注和"未公开""未披露"说明，
      docx 同时查页眉、页脚和修订中删除的文字（规则见同目录 secret_marks.py）。任何一份命中即停止，不输出对比清单，
      只报"某份含某类标注，未比对"。
隐私：修订人姓名只显示首字（如"王**"），不输出全名；批注原文照录，归纳偏好时须去掉人名和对个人的评价。
修订人：有修订或批注的人都编入修订人清单（只写批注的，修订处数记 0）；--author N 时批注也只列该修订人的。
只比对【起草说明】之前的正文。

退出码：0 正常（含两稿无差异）；1 含密级或内部标注，未比对；2 输入错误。只依赖标准库。
"""
from __future__ import annotations

import argparse
import difflib
import json
import re
import sys
import zipfile
from collections import Counter, OrderedDict
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET

_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))
import secret_marks  # noqa: E402  （同目录的共用模块）

VERSION = "1.2.0"
BOUNDARY = "【起草说明】"
DRAFT_NOTE_RE = re.compile(r"^\s*(?:#{1,6}\s*)?(?:(?:\*\*|__)\s*)?" + re.escape(BOUNDARY))
SUPPORTED_SUFFIXES = (".txt", ".md", ".docx")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS

WS_RE = re.compile(r"\s+")
UNIT_RE = re.compile(r"[^。！？；]+[。！？；]?")
PUNCT_ONLY_RE = re.compile(r"^[\s，。、；：！？“”‘’（）《》〈〉【】—…·,.;:!?()\"'\-]+$")
DIGIT_RE = re.compile(r"\d|[一二三四五六七八九十百千万亿两]+(?:个|项|次|人|家|件|台|条|元|万|亿|倍|成)|百分点|%|％")
HAN_ONLY_RE = re.compile(r"^[\u4e00-\u9fff]+$")
CN_NUM_ONLY_RE = re.compile(r"^[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+$")
# \u6982\u62ec\u8bf4\u6cd5\uff1a\u6362\u6210\u5177\u4f53\u6570\u5b57\u7684\uff0c\u7b97"\u6539\u4e3a\u5177\u4f53\u6570\u5b57"\uff0c\u4e0d\u7b97\u539f\u6709\u6570\u5b57\u7684\u66f4\u6b63
APPROX_RE = re.compile(r"\u8fd1|\u7ea6|\u4f59|\u591a|\u5de6\u53f3|\u4e0a\u4e0b|\u4e00\u6279|\u90e8\u5206|\u5927\u91cf|\u82e5\u5e72|\u5927\u5e45|\u660e\u663e|\u8f83\u5927|\u4e0d\u5c11|\u4f17\u591a|\u4e00\u4e9b")
# \u53e5\u9996\u5c42\u6b21\u5e8f\u53f7\uff1a\u4e00\u3001\uff08\u4e00\uff091.\uff081\uff09\uff1b"1.2\u4e07"\u8fd9\u7c7b\u5c0f\u6570\u4e0d\u7b97
ORDINAL_RE = re.compile(r"^(?:[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]{1,3}[\u3001\uff0e.]|[\uff08(][\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]{1,3}[\uff09)]|"
                        r"\d{1,2}[.\uff0e\u3001](?!\d)|[\uff08(]\d{1,2}[\uff09)])")
CONTEXT = 6
MAX_DP_CELLS = 250000
MAX_MOVE_PAIRS = 40000   # 删除句数 × 新增句数超过此值时，调整顺序只认完全相同的句子
MOVE_SIM = 0.85          # 删除句与新增句相似度达到此值，记为调整顺序

# 评价性形容词、副词：领导删得最多的一类
EVAL_WORDS = (
    "显著", "明显", "大幅", "大力", "全面", "深入", "扎实", "积极", "有效", "切实", "高度", "充分",
    "不断", "进一步", "持续", "全力", "坚决", "认真", "着力", "努力", "较大", "很大", "极大", "巨大",
    "突出", "卓有成效", "圆满", "成功", "重大", "重要", "有力", "稳步", "广泛", "全方位", "多措并举",
    "取得了", "非常", "十分", "特别", "尤其", "首次", "率先", "领先", "一流", "优异", "丰硕",
)
PREAMBLE_RE = re.compile(r"随着|近年来|当前[，,]|新时代|新形势|新征程|在[^。]{0,20}领导下|为(?:深入)?贯彻落实|"
                         r"高度重视|重要意义|具有十分|必要性|紧迫性")
PLEDGE_RE = re.compile(r"我们将|坚决|一定要|全力以赴|确保[^。]{0,12}圆满|努力开创|再创佳绩|奋力|勇毅")
SOFT_RE = re.compile(r"建议|可以考虑|尽量|争取|希望|酌情|适时")
FIRM_RE = re.compile(r"要|必须|务必|确保|严禁|不得|于\d")


class InputError(Exception):
    """输入错误：文件不存在、类型不支持、无法解析。"""


# ---------------------------------------------------------------------------
# 读取
# ---------------------------------------------------------------------------

def _decode(data: bytes, name: str, notes: List[str]) -> str:
    for encoding in ("utf-8-sig", "gb18030"):
        try:
            text = data.decode(encoding)
        except UnicodeDecodeError:
            continue
        if encoding != "utf-8-sig":
            notes.append(f"{name} 不是 UTF-8 编码，已按 GB18030 读取")
        return text
    raise InputError(f"无法识别文件编码，请另存为 UTF-8：{name}")


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


class DocxRevisions:
    """解析带修订痕迹的 docx：原稿视图、改稿视图、修订人、批注。"""

    INS_TAGS = ("ins", "moveTo")
    DEL_TAGS = ("del", "moveFrom")

    def __init__(self, path: Path) -> None:
        self.path = path
        try:
            with zipfile.ZipFile(str(path)) as zf:
                names = set(zf.namelist())
                document = zf.read("word/document.xml")
                comments = zf.read("word/comments.xml") if "word/comments.xml" in names else None
        except (zipfile.BadZipFile, KeyError, OSError) as exc:
            raise InputError(f"无法打开 docx 文件：{path.name}（{exc}）")
        try:
            self.root = ET.fromstring(document)
            self.comments_root = ET.fromstring(comments) if comments else None
        except ET.ParseError as exc:
            raise InputError(f"docx 内容无法解析：{path.name}（{exc}）")
        self.authors: "OrderedDict[str, dict]" = OrderedDict()
        self._collect_authors()

    # -- 修订人 ----------------------------------------------------------
    def _collect_authors(self) -> None:
        """先按修订出现顺序编号，再补上只写批注、没做修订的人（修订处数为 0）。"""
        for node in self.root.iter():
            name = _local(node.tag)
            if name in self.INS_TAGS or name in self.DEL_TAGS:
                author = node.get(W + "author", "")
                date = node.get(W + "date", "")
                info = self.authors.setdefault(author, {"count": 0, "comments": 0, "dates": set()})
                info["count"] += 1
                if date:
                    info["dates"].add(date[:10])
        if self.comments_root is not None:
            for c in self.comments_root.iter(W + "comment"):
                author = c.get(W + "author", "")
                info = self.authors.setdefault(author, {"count": 0, "comments": 0, "dates": set()})
                info["comments"] += 1
                date = c.get(W + "date", "") or ""
                if date:
                    info["dates"].add(date[:10])

    @property
    def has_revisions(self) -> bool:
        return any(info["count"] for info in self.authors.values())

    def alias_of(self, author: str) -> str:
        names = list(self.authors)
        return f"修订人{names.index(author) + 1}" if author in names else "修订人?"

    def author_rows(self) -> List[dict]:
        rows = []
        for idx, (author, info) in enumerate(self.authors.items(), start=1):
            dates = sorted(info["dates"])
            rows.append({
                "index": idx,
                "alias": f"修订人{idx}",
                "masked_name": mask_name(author),
                "revisions": info["count"],
                "comments": info["comments"],
                "dates": [dates[0], dates[-1]] if dates else [],
            })
        return rows

    # -- 文本视图 --------------------------------------------------------
    def lines(self, view: str, only_author: Optional[str] = None) -> List[str]:
        """view = "original"（拒绝全部修订）或 "revised"（接受修订）。

        only_author 非空时，只接受该修订人的修订，其他人的修订按拒绝处理。
        """
        body = self.root.find(W + "body")
        result: List[str] = []
        if body is None:
            return result

        def accepted(node: ET.Element) -> bool:
            if view == "original":
                return False
            return only_author is None or node.get(W + "author", "") == only_author

        def para_text(p: ET.Element) -> str:
            parts: List[str] = []

            def walk(node: ET.Element, state: str) -> None:
                # state: normal / ins-accepted / ins-rejected / del-accepted / del-rejected
                name = _local(node.tag)
                # 祖先已隐藏的（如甲插入、乙又删掉的 <w:ins><w:del>…），子节点保持隐藏
                if name in self.INS_TAGS:
                    state = "hide" if state == "hide" or not accepted(node) else "show"
                elif name in self.DEL_TAGS:
                    state = "hide" if state == "hide" or accepted(node) else "show"
                if state != "hide":
                    if name in ("t", "delText") and node.text:
                        parts.append(node.text)
                    elif name == "tab":
                        parts.append("\t")
                    elif name in ("br", "cr"):
                        parts.append("\n")
                for child in node:
                    if _local(child.tag) in ("pPr", "rPr", "instrText", "delInstrText"):
                        continue
                    walk(child, state)

            walk(p, "show")
            return "".join(parts)

        def visit(node: ET.Element) -> None:
            for child in node:
                name = _local(child.tag)
                if name == "p":
                    result.extend(para_text(child).split("\n"))
                elif name == "tbl":
                    for tr in child.iter(W + "tr"):
                        cells = ["".join(para_text(p) for p in tc.iter(W + "p")) for tc in tr.findall(W + "tc")]
                        result.append("\t".join(cells))
                elif name in ("sdt", "sdtContent", "customXml", "smartTag") or name in self.INS_TAGS \
                        or name in self.DEL_TAGS:
                    visit(child)

        visit(body)
        return result

    # -- 批注 ------------------------------------------------------------
    def comments(self) -> List[dict]:
        if self.comments_root is None:
            return []
        anchors: Dict[str, List[str]] = {}
        open_ids: List[str] = []

        def walk(node: ET.Element) -> None:
            name = _local(node.tag)
            if name == "commentRangeStart":
                cid = node.get(W + "id", "")
                open_ids.append(cid)
                anchors.setdefault(cid, [])
            elif name == "commentRangeEnd":
                cid = node.get(W + "id", "")
                if cid in open_ids:
                    open_ids.remove(cid)
            elif name in ("t", "delText") and node.text:
                for cid in open_ids:
                    anchors[cid].append(node.text)
            for child in node:
                walk(child)

        walk(self.root)
        rows = []
        for c in self.comments_root.iter(W + "comment"):
            cid = c.get(W + "id", "")
            author = c.get(W + "author", "")
            text = "".join(t.text or "" for t in c.iter(W + "t")).strip()
            anchor = WS_RE.sub("", "".join(anchors.get(cid, [])))
            rows.append({
                "id": cid,
                "author": self.alias_of(author),
                "masked_name": mask_name(author),
                "date": (c.get(W + "date", "") or "")[:10],
                "anchor": anchor[:40] + ("……" if len(anchor) > 40 else ""),
                "text": text,
            })
        return rows


def mask_name(name: str) -> str:
    name = (name or "").strip()
    if not name:
        return "未署名"
    return name[0] + "*" * min(len(name) - 1, 3)


def read_lines(path: Path, notes: List[str]) -> List[str]:
    if not path.exists():
        raise InputError(f"文件不存在：{path}")
    if not path.is_file():
        raise InputError(f"不是文件：{path}")
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise InputError(f"不支持的文件类型 {suffix or '（无扩展名）'}，只支持 .txt、.md、.docx：{path.name}")
    if suffix == ".docx":
        rev = DocxRevisions(path)
        if rev.has_revisions:
            notes.append(f"{path.name} 含未接受的修订，已按“接受全部修订”后的文本比对")
        return rev.lines("revised")
    try:
        data = path.read_bytes()
    except OSError as exc:
        raise InputError(f"读取失败：{path.name}（{exc}）")
    text = _decode(data, path.name, notes)
    return text.replace("\r\n", "\n").replace("\r", "\n").split("\n")


def cut_boundary(lines: Sequence[str]) -> List[str]:
    out: List[str] = []
    for line in lines:
        if DRAFT_NOTE_RE.match(line):
            break
        out.append(line)
    return out


# ---------------------------------------------------------------------------
# 切分与对齐
# ---------------------------------------------------------------------------

def strip_ordinal(text: str) -> str:
    """去掉句首的层次序号，用于比较；只剩序号的，原样返回。"""
    stripped = ORDINAL_RE.sub("", text, count=1)
    return stripped or text


class Unit:
    __slots__ = ("text", "key", "para", "sent")

    def __init__(self, text: str, para: int, sent: int) -> None:
        self.text = text
        self.key = strip_ordinal(text)  # 比较用：去掉层次序号
        self.para = para
        self.sent = sent

    @property
    def where(self) -> str:
        return f"第{self.para}段第{self.sent}句"


def to_units(lines: Sequence[str]) -> List[Unit]:
    units: List[Unit] = []
    para = 0
    for line in lines:
        compact = WS_RE.sub("", line)
        if not compact:
            continue
        para += 1
        sent = 0
        for match in UNIT_RE.finditer(compact):
            piece = match.group(0)
            if not piece:
                continue
            sent += 1
            units.append(Unit(piece, para, sent))
    return units


def similarity(a: str, b: str) -> float:
    """相似度：取 SequenceMatcher 比值与"重合字数占短句比例"（打八折）中的较大者。

    后者照顾领导把短句扩写成长句、把长句删成短句的情形。
    """
    if a == b:
        return 1.0
    if not a or not b:
        return 0.0
    matcher = difflib.SequenceMatcher(None, a, b, autojunk=False)
    if matcher.real_quick_ratio() < 0.2:
        return matcher.real_quick_ratio()
    matched = sum(block.size for block in matcher.get_matching_blocks())
    ratio = 2.0 * matched / (len(a) + len(b))
    overlap = matched / min(len(a), len(b))
    return max(ratio, 0.8 * overlap)


def plain_ratio(a: str, b: str) -> float:
    if a == b:
        return 1.0
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()


def align_block(a: Sequence[Unit], b: Sequence[Unit], min_sim: float) -> List[Tuple[List[int], List[int], float]]:
    """锚点之间的块内对齐。返回 [(a 下标列表, b 下标列表, 相似度)]，空列表表示删除或新增。"""
    n, m = len(a), len(b)
    if n == 0 or m == 0:
        return [([i], [], 0.0) for i in range(n)] + [([], [j], 0.0) for j in range(m)]
    if n * m > MAX_DP_CELLS:
        return _greedy_align(a, b, min_sim)
    neg = float("-inf")
    score = [[neg] * (m + 1) for _ in range(n + 1)]
    back: List[List[Optional[Tuple[int, int, float]]]] = [[None] * (m + 1) for _ in range(n + 1)]
    score[0][0] = 0.0
    sim_cache: Dict[Tuple[int, int, int, int], float] = {}

    def sim(i0: int, i1: int, j0: int, j1: int) -> float:
        key = (i0, i1, j0, j1)
        if key not in sim_cache:
            sa = "".join(u.key for u in a[i0:i1])
            sb = "".join(u.key for u in b[j0:j1])
            # 拆句、合句只用比值，不用"重合占短句比例"，避免把无关的相邻句并进来
            single = (i1 - i0 == 1 and j1 - j0 == 1)
            sim_cache[key] = similarity(sa, sb) if single else plain_ratio(sa, sb)
        return sim_cache[key]

    for i in range(n + 1):
        for j in range(m + 1):
            cur = score[i][j]
            if cur == neg:
                continue
            # 删除、新增：不得分
            if i < n and cur > score[i + 1][j]:
                score[i + 1][j], back[i + 1][j] = cur, (i, j, 0.0)
            if j < m and cur > score[i][j + 1]:
                score[i][j + 1], back[i][j + 1] = cur, (i, j, 0.0)
            # 一对一、一拆二、二合一：相似度达到门槛才可配对
            for di, dj, penalty in ((1, 1, 0.0), (1, 2, 0.05), (2, 1, 0.05)):
                if i + di <= n and j + dj <= m:
                    s = sim(i, i + di, j, j + dj)
                    if s >= min_sim:
                        # 得分按覆盖的句数加权：拆句、合句一次配上三句，相似度相当时优先于一对一加一条新增
                        val = cur + s * (di + dj) / 2.0 - penalty
                        if val > score[i + di][j + dj]:
                            score[i + di][j + dj], back[i + di][j + dj] = val, (i, j, s)
    pairs: List[Tuple[List[int], List[int], float]] = []
    i, j = n, m
    while i > 0 or j > 0:
        pi, pj, s = back[i][j]  # type: ignore[misc]
        pairs.append((list(range(pi, i)), list(range(pj, j)), s))
        i, j = pi, pj
    pairs.reverse()
    return pairs


def _greedy_align(a: Sequence[Unit], b: Sequence[Unit], min_sim: float) -> List[Tuple[List[int], List[int], float]]:
    """块太大时的退化方案：按顺序为每个原句在后续窗口内找最相似的改句。"""
    pairs: List[Tuple[List[int], List[int], float]] = []
    j = 0
    for i in range(len(a)):
        best, best_j = 0.0, None
        for k in range(j, min(len(b), j + 8)):
            s = similarity(a[i].key, b[k].key)
            if s > best:
                best, best_j = s, k
        if best_j is not None and best >= min_sim:
            for k in range(j, best_j):
                pairs.append(([], [k], 0.0))
            pairs.append(([i], [best_j], best))
            j = best_j + 1
        else:
            pairs.append(([i], [], 0.0))
    for k in range(j, len(b)):
        pairs.append(([], [k], 0.0))
    return pairs


def align(a: Sequence[Unit], b: Sequence[Unit], min_sim: float) -> List[Tuple[List[int], List[int], float]]:
    matcher = difflib.SequenceMatcher(None, [u.key for u in a], [u.key for u in b], autojunk=False)
    result: List[Tuple[List[int], List[int], float]] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            result.extend(([i1 + k], [j1 + k], 1.0) for k in range(i2 - i1))
        else:
            for ia, jb, s in align_block(a[i1:i2], b[j1:j2], min_sim):
                result.append(([i1 + x for x in ia], [j1 + y for y in jb], s))
    return result


# ---------------------------------------------------------------------------
# 片段比对与标签
# ---------------------------------------------------------------------------

def char_edits(before: str, after: str) -> List[dict]:
    matcher = difflib.SequenceMatcher(None, before, after, autojunk=False)
    ops = [list(op) for op in matcher.get_opcodes()]
    # 合并被 1 个相同字隔开的相邻改动，得到词一级的片段
    merged: List[List] = []
    k = 0
    while k < len(ops):
        op = ops[k]
        if (merged and op[0] == "equal" and op[2] - op[1] <= 1 and k + 1 < len(ops)
                and merged[-1][0] != "equal" and ops[k + 1][0] != "equal"):
            nxt = ops[k + 1]
            prev = merged[-1]
            merged[-1] = ["replace", prev[1], nxt[2], prev[3], nxt[4]]
            k += 2
            continue
        if merged and op[0] != "equal" and merged[-1][0] != "equal":
            prev = merged[-1]
            merged[-1] = ["replace", prev[1], op[2], prev[3], op[4]]
        else:
            merged.append(op)
        k += 1
    edits = []
    for tag, i1, i2, j1, j2 in merged:
        if tag == "equal":
            continue
        a_part, b_part = before[i1:i2], after[j1:j2]
        kind = "删除" if not b_part else ("新增" if not a_part else "替换")
        edits.append({
            "kind": kind,
            "before": a_part,
            "after": b_part,
            "context": before[max(0, i1 - CONTEXT):i1] + "【" + a_part + "→" + b_part + "】" + before[i2:i2 + CONTEXT],
            "tags": fragment_tags(a_part, b_part),
        })
    return edits


def fragment_tags(before: str, after: str) -> List[str]:
    tags: List[str] = []
    if PUNCT_ONLY_RE.match(before + after or " "):
        return ["标点"]
    digit_before, digit_after = DIGIT_RE.search(before), DIGIT_RE.search(after)
    if "百分点" in after and ("%" in before or "％" in before) and not re.search(r"\d", before + after):
        digit_before = digit_after = None  # 只把"%"改成"个百分点"，数值没变，由下面的"改为百分点"标出
    if digit_after and (not digit_before or APPROX_RE.search(before)):
        tags.append("改为具体数字")   # 概括说法换成数字，不算事实更正
    elif digit_before and digit_after:
        tags.append("数字改动")       # 原有数字改成另一个数字，可能是事实更正
    elif digit_before:
        tags.append("删去数字")
    removed = [w for w in EVAL_WORDS if w in before and w not in after]
    if removed:
        tags.append("评价词删减")
    if before and after and len(before) <= 6 and len(after) <= 6 and HAN_ONLY_RE.match(before) \
            and HAN_ONLY_RE.match(after) and not (digit_before or digit_after) \
            and not (CN_NUM_ONLY_RE.match(before) and CN_NUM_ONLY_RE.match(after)):
        tags.append("换词")
    if SOFT_RE.search(before) and not SOFT_RE.search(after) and FIRM_RE.search(after):
        tags.append("语气改为要求")
    if "百分点" in after and ("%" in before or "％" in before):
        tags.append("改为百分点")
    return tags


def sentence_tags(text: str, deleted: bool) -> List[str]:
    tags = []
    if deleted:
        if PREAMBLE_RE.search(text):
            tags.append("删铺垫")
        if PLEDGE_RE.search(text):
            tags.append("删表态")
        if DIGIT_RE.search(text):
            tags.append("删除含数据句")   # 整句删去，不是把数字改成另一个数字
        if any(w in text for w in EVAL_WORDS):
            tags.append("含评价词")
    else:
        if DIGIT_RE.search(text):
            tags.append("新增数据")
    return tags


def _first_digit_ratio(text: str) -> Optional[float]:
    match = re.search(r"\d", text)
    if not match or not text:
        return None
    return match.start() / len(text)


def build_report(original_lines: Sequence[str], revised_lines: Sequence[str], min_sim: float,
                 source: dict, notes: List[str], authors: Optional[List[dict]] = None,
                 comments: Optional[List[dict]] = None) -> dict:
    a = to_units(cut_boundary(original_lines))
    b = to_units(cut_boundary(revised_lines))
    pairs = align(a, b, min_sim)
    items: List[dict] = []
    deleted_idx: List[int] = []
    inserted_idx: List[int] = []
    renumbered = 0  # 只有层次序号变了的句子，不算改动
    for ia, jb, s in pairs:
        if ia and jb:
            before = "".join(a[i].text for i in ia)
            after = "".join(b[j].text for j in jb)
            key_before = "".join(a[i].key for i in ia)
            key_after = "".join(b[j].key for j in jb)
            if key_before == key_after:
                if before != after:
                    renumbered += 1
                continue
            kind = "拆句" if len(jb) > len(ia) else ("合句" if len(ia) > len(jb) else "改写")
            tags = [kind] if kind != "改写" else []
            ra, rb = _first_digit_ratio(key_before), _first_digit_ratio(key_after)
            if ra is not None and rb is not None and rb + 0.2 < ra:
                tags.append("数据前移")
            if len(key_after) <= 0.8 * len(key_before):
                tags.append("句子压缩")
            edits = char_edits(key_before, key_after)
            for e in edits:
                for t in e["tags"]:
                    if t not in tags:
                        tags.append(t)
            items.append({
                "type": kind,
                "where_before": a[ia[0]].where,
                "where_after": b[jb[0]].where,
                "before": before,
                "after": after,
                "similarity": round(s, 2),
                "edits": edits,
                "tags": tags,
            })
        elif ia:
            deleted_idx.extend(ia)
        else:
            inserted_idx.extend(jb)

    # 删除句与新增句相同或高度相似：调整顺序
    moved_pairs: List[Tuple[int, int, float]] = []
    used_ins = set()
    by_text: Dict[str, List[int]] = {}
    for j in inserted_idx:
        by_text.setdefault(b[j].key, []).append(j)
    # 先配完全相同的句子（不计层次序号）；句子很多时只做这一步，避免两两比较过慢
    fuzzy = len(deleted_idx) * len(inserted_idx) <= MAX_MOVE_PAIRS
    for i in deleted_idx:
        same = [j for j in by_text.get(a[i].key, []) if j not in used_ins]
        if same:
            moved_pairs.append((i, same[0], 1.0))
            used_ins.add(same[0])
            continue
        if not fuzzy:
            continue
        best, best_j = 0.0, None
        for j in inserted_idx:
            if j in used_ins:
                continue
            matcher = difflib.SequenceMatcher(None, a[i].key, b[j].key, autojunk=False)
            if matcher.real_quick_ratio() < MOVE_SIM or matcher.quick_ratio() < MOVE_SIM:
                continue
            s = matcher.ratio()
            if s > best:
                best, best_j = s, j
        if best_j is not None and best >= MOVE_SIM:
            moved_pairs.append((i, best_j, best))
            used_ins.add(best_j)
    moved_del = {i for i, _, _ in moved_pairs}
    for i, j, s in moved_pairs:
        edits = char_edits(a[i].key, b[j].key) if s < 1.0 else []
        tags = ["调整顺序"]
        for e in edits:
            for t in e["tags"]:
                if t not in tags:
                    tags.append(t)
        items.append({"type": "调整顺序", "where_before": a[i].where, "where_after": b[j].where,
                      "before": a[i].text, "after": b[j].text, "similarity": round(s, 2),
                      "edits": edits, "tags": tags})
    for i in deleted_idx:
        if i in moved_del:
            continue
        items.append({"type": "删除句", "where_before": a[i].where, "where_after": "",
                      "before": a[i].text, "after": "", "similarity": 0.0, "edits": [],
                      "tags": sentence_tags(a[i].text, True)})
    for j in inserted_idx:
        if j in used_ins:
            continue
        items.append({"type": "新增句", "where_before": "", "where_after": b[j].where,
                      "before": "", "after": b[j].text, "similarity": 0.0, "edits": [],
                      "tags": sentence_tags(b[j].text, False)})
    items.sort(key=lambda x: _sort_key(x))
    for k, item in enumerate(items, start=1):
        item["no"] = k

    # 汇总
    tag_counts: Counter = Counter()
    swaps: Counter = Counter()
    removed_words: Counter = Counter()
    for item in items:
        tag_counts.update(item["tags"])
        for e in item["edits"]:
            if "换词" in e["tags"]:
                swaps[f"{e['before']}→{e['after']}"] += 1
            for w in EVAL_WORDS:
                if w in e["before"] and w not in e["after"]:
                    removed_words[w] += 1
        if item["type"] == "删除句":
            for w in EVAL_WORDS:
                if w in item["before"]:
                    removed_words[w] += item["before"].count(w)
    chars_a = sum(len(u.text) for u in a)
    chars_b = sum(len(u.text) for u in b)
    type_counts = Counter(item["type"] for item in items)
    summary = OrderedDict([
        ("chars_before", chars_a),
        ("chars_after", chars_b),
        ("change_ratio", round((chars_b - chars_a) / chars_a, 3) if chars_a else 0.0),
        ("sentences_before", len(a)),
        ("sentences_after", len(b)),
        ("avg_sentence_before", round(chars_a / len(a), 1) if a else 0.0),
        ("avg_sentence_after", round(chars_b / len(b), 1) if b else 0.0),
        ("by_type", dict(type_counts.most_common())),
        ("by_tag", dict(tag_counts.most_common())),
        ("word_swaps", [{"pair": k, "count": v} for k, v in swaps.most_common()]),
        ("removed_eval_words", [{"word": k, "count": v} for k, v in removed_words.most_common()]),
        ("renumbered", renumbered),
        ("identical", not items and not renumbered),
    ])
    report = OrderedDict()
    report["tool"] = "draft_diff.py"
    report["version"] = VERSION
    report["source"] = source
    report["scope"] = "【起草说明】之前的正文；比较时去掉全部空白"
    report["notes"] = notes
    report["authors"] = authors or []
    report["summary"] = summary
    report["comments"] = comments or []
    report["items"] = items
    return report


def blocked_report(source: dict, notes: List[str], blocked: List[dict]) -> dict:
    """含密级或内部标注时的报告：只列文件名和标注类别，不含任何正文、批注和修订人。"""
    report = OrderedDict()
    report["tool"] = "draft_diff.py"
    report["version"] = VERSION
    report["source"] = source
    report["scope"] = "未比对：材料含密级或内部标注"
    report["notes"] = notes
    report["blocked"] = blocked
    report["authors"] = []
    report["summary"] = {}
    report["comments"] = []
    report["items"] = []
    return report


def marks_of(path: Path, lines: Sequence[str]) -> List[str]:
    """一份材料命中的标注类别；docx 另查页眉、页脚和修订中删除的文字。"""
    extra = secret_marks.docx_mark_lines(path) if path.suffix.lower() == ".docx" else []
    return secret_marks.find_marks(list(lines) + extra)


def _sort_key(item: dict) -> Tuple[int, int, int]:
    where = item["where_before"] or item["where_after"]
    match = re.match(r"第(\d+)段第(\d+)句", where)
    para, sent = (int(match.group(1)), int(match.group(2))) if match else (0, 0)
    return (para, sent, 0 if item["where_before"] else 1)


# ---------------------------------------------------------------------------
# 文本输出
# ---------------------------------------------------------------------------

def _cell(text: str, limit: int = 80) -> str:
    text = text.replace("|", "｜")
    return text if len(text) <= limit else text[:limit] + "……"


def render_text(report: dict) -> str:
    out: List[str] = []
    src = report["source"]
    s = report["summary"]
    out.append(f"改稿对比（draft_diff.py {report['version']}）")
    if src.get("mode") == "docx-revisions":
        out.append(f"输入：{src['file']}（按修订痕迹还原原稿和改稿）")
        if src.get("author"):
            out.append(f"只采用：{src['author']}的修订和批注")
    else:
        out.append(f"原稿：{src['original']}　改稿：{src['revised']}")
    for note in report["notes"]:
        out.append("提示：" + note)
    if report.get("blocked"):
        for b in report["blocked"]:
            out.append(f"已停止：{b['file']} 含{secret_marks.describe(b['marks'])}，未比对，未输出任何内容。")
        out.append("涉密信息和内部敏感信息不得输入联网 AI 工具。请改用单位内网部署、允许处理此类材料的模型；"
                   "或删除相关内容后再比对；标识系误加或材料已公开的，去掉标识后再比对。")
        return "\n".join(out)
    if s["identical"] and not report["comments"]:
        out.append("两稿正文没有差异。")
        return "\n".join(out)
    if report["authors"]:
        out.append("修订人：" + "；".join(
            f"{r['alias']}（{r['masked_name']}）修订 {r['revisions']} 处"
            + (f"、批注 {r['comments']} 条" if r.get("comments") else "")
            + (f"，{r['dates'][0]}" + (f"至{r['dates'][1]}" if r['dates'][1] != r['dates'][0] else "") if r["dates"] else "")
            for r in report["authors"]))
    pct = round(s["change_ratio"] * 100, 1)
    out.append(f"概况：原稿 {s['chars_before']} 字 → 改稿 {s['chars_after']} 字（{'+' if pct > 0 else ''}{pct}%）；"
               f"平均句长 {s['avg_sentence_before']} → {s['avg_sentence_after']} 字；"
               + ("、".join(f"{k} {v} 处" for k, v in s["by_type"].items()) or "无正文改动"))
    if s.get("renumbered"):
        out.append(f"层次序号变化 {s['renumbered']} 处（只是序号随增删、调序而变，不计入改动）")
    if s["by_tag"]:
        out.append("标签统计（启发式，需人工判断）：" + "、".join(f"{k} {v}" for k, v in s["by_tag"].items()))
    if s["word_swaps"]:
        out.append("换词：" + "、".join(f"{x['pair']}（{x['count']}）" for x in s["word_swaps"]))
    if s["removed_eval_words"]:
        out.append("删去的评价词：" + "、".join(f"{x['word']}（{x['count']}）" for x in s["removed_eval_words"]))
    if report["comments"]:
        out.append("")
        out.append("批注：")
        for k, c in enumerate(report["comments"], start=1):
            out.append(f"{k}. [{c['author']}，{c['date'] or '无日期'}] 锚定“{c['anchor']}”：{c['text']}")
    if report["items"]:
        out.append("")
        out.append("| 序号 | 位置（原稿/改稿） | 类型 | 原稿 | 改稿 | 字一级改动 | 标签 |")
        out.append("|---|---|---|---|---|---|---|")
        for item in report["items"]:
            where = "/".join(x for x in (item["where_before"], item["where_after"]) if x)
            edits = "；".join(e["context"] for e in item["edits"][:4])
            if len(item["edits"]) > 4:
                edits += f"；等 {len(item['edits'])} 处"
            out.append(f"| {item['no']} | {where} | {item['type']} | {_cell(item['before'])} | "
                       f"{_cell(item['after'])} | {_cell(edits, 120)} | {'、'.join(item['tags'])} |")
    out.append("")
    out.append("说明：本清单含原文，只作归纳偏好的工作底稿，不存入 .gongwen/；偏好卡只记归纳后的条目和不超过 20 字的例证。")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    k = 2
    while True:
        candidate = path.with_name(f"{path.stem}-v{k}{path.suffix}")
        if not candidate.exists():
            return candidate
        k += 1


def _setup_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


class _ChineseHelpFormatter(argparse.RawDescriptionHelpFormatter):
    def add_usage(self, usage, actions, groups, prefix=None):
        return super().add_usage(usage, actions, groups, "用法：" if prefix is None else prefix)


def make_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="draft_diff.py",
        description="花脸稿对比：按句对齐原稿和领导改稿，抽取删除、新增、替换；"
                    "也可以只给一份带修订痕迹的 docx，按修订还原两稿，并读取批注。",
        epilog="退出码：0 正常（含两稿无差异）；1 含密级或内部标注，未比对；2 输入错误。",
        formatter_class=_ChineseHelpFormatter,
        add_help=False,
    )
    parser._optionals.title = "选项"
    parser._positionals.title = "参数"
    parser.add_argument("-h", "--help", action="help", help="显示本帮助信息并退出")
    parser.add_argument("files", nargs="+", help="原稿和改稿两份文件；或一份带修订痕迹的 .docx")
    parser.add_argument("--json", action="store_true", help="以 JSON 格式输出")
    parser.add_argument("--min-sim", type=float, default=0.45,
                        help="两句相似度达到多少才算同一句的改写（0～1，默认 0.45）")
    parser.add_argument("--author", type=int, help="只采用第 N 位修订人的修订（仅单个 docx 时可用）")
    parser.add_argument("--output", help="把结果另存为文件；文件已存在时加 -v2、-v3，不覆盖")
    return parser


def run(args: argparse.Namespace) -> dict:
    notes: List[str] = []
    if not 0 < args.min_sim <= 1:
        raise InputError("--min-sim 必须在 0 到 1 之间")
    if len(args.files) == 2:
        if args.author is not None:
            raise InputError("--author 只在输入一份带修订痕迹的 docx 时使用")
        a_path, b_path = Path(args.files[0]), Path(args.files[1])
        a_lines = read_lines(a_path, notes)
        b_lines = read_lines(b_path, notes)
        source = {"mode": "two-files", "original": a_path.name, "revised": b_path.name}
        blocked = []
        for p, lines in ((a_path, a_lines), (b_path, b_lines)):
            marks = marks_of(p, lines)
            if marks:
                blocked.append({"file": p.name, "marks": marks})
        if blocked:
            return blocked_report(source, notes, blocked)
        return build_report(a_lines, b_lines, args.min_sim, source, notes)
    if len(args.files) == 1:
        path = Path(args.files[0])
        if not path.exists():
            raise InputError(f"文件不存在：{path}")
        if path.suffix.lower() != ".docx":
            raise InputError("只给一份文件时，必须是带修订痕迹的 .docx；否则请同时提供原稿和改稿")
        rev = DocxRevisions(path)
        source = {"mode": "docx-revisions", "file": path.name}
        marks = marks_of(path, rev.lines("original") + rev.lines("revised"))
        if marks:
            return blocked_report(source, notes, [{"file": path.name, "marks": marks}])
        authors = rev.author_rows()
        comments = rev.comments()
        if not rev.has_revisions:
            notes.append("未发现修订痕迹（w:ins、w:del）。领导直接改了没留痕迹的，请同时提供原稿和改稿两份文件")
        only = None
        if args.author is not None:
            names = list(rev.authors)
            if not 1 <= args.author <= len(names):
                raise InputError(f"--author 超出范围：本文件共有 {len(names)} 位修订人")
            only = names[args.author - 1]
            alias = f"修订人{args.author}"
            source["author"] = alias
            others = [c for c in comments if c["author"] != alias]
            comments = [c for c in comments if c["author"] == alias]
            if others:
                notes.append(f"其他修订人的批注 {len(others)} 条未列出（只列{alias}的批注）")
        return build_report(rev.lines("original"), rev.lines("revised", only), args.min_sim, source, notes,
                            authors=authors, comments=comments)
    raise InputError("请给出原稿和改稿两份文件，或一份带修订痕迹的 .docx")


def main(argv: Optional[Sequence[str]] = None) -> int:
    _setup_console()
    parser = make_parser()
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:
        return 0 if exc.code == 0 else 2
    try:
        report = run(args)
    except InputError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    text = json.dumps(report, ensure_ascii=False, indent=2) if args.json else render_text(report)
    if args.output:
        target = unique_path(Path(args.output))
        try:
            target.write_text(text + "\n", encoding="utf-8")
        except OSError as exc:
            print(f"错误：无法写入 {target}（{exc}）", file=sys.stderr)
            return 2
        print(f"已写入：{target}")
    else:
        print(text)
    return 1 if report.get("blocked") else 0


if __name__ == "__main__":
    sys.exit(main())
