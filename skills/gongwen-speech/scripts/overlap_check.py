# -*- coding: utf-8 -*-
"""文本重合检测：找出稿件与参照文件之间的连续重合片段，计算重合率。

源文件位于仓库 shared/scripts/，各 skill 中的副本由 tools/sync_shared.py 生成，请在源文件处修改。
维护者：gongwen-implement。接口见 docs/开发规范.md 8.3 节。

用途：
    gongwen-implement：检测实施方案、贯彻落实方案是否照搬照抄上位文件
        （《整治形式主义为基层减负若干规定》第2条）；
    gongwen-party：检测材料与往年稿、他人稿是否雷同。

用法：
    python scripts/overlap_check.py <稿件> <参照文件> [<参照文件>...] [--min-len 15] [--json] [--strict]

比对规则：
    1. 稿件只取【起草说明】之前的正文；参照文件全文参与比对。支持 .txt、.md、.docx（docx 按段落计行）。
    2. 比对前去掉空白和标点符号；全角字母、数字转半角，英文字母不分大小写。
       汉字、字母、数字各算一个字。
    3. 稿件中书名号《》内的文件名称、【】内的【待补】【待核实】标记，不计入正文字数，也不参与比对；
       重合片段不跨越这些位置，也不跨越稿件的段落（行）。
    4. 找出长度不少于 --min-len 个字（默认 15）的连续重合片段。参照文件不分段，
       稿件把参照的相邻两段合成一段照抄的，同样能查出。
    5. 重合率 = 稿件中被重合片段覆盖的字数 ÷ 稿件正文字数。每个参照文件单独计算，
       另给出全部参照合计的重合率（同一处只算一次）。
    6. 稿件中某一行（段）被覆盖的字数达到该段字数一半以上的，列为"高重合段落"。
    7. 重合片段整体落在稿件同一对引号（“”「」，或同一行内成对的半角双引号）之内的，标为 quoted，
       表示带引号的原文引用，是否属于正当引用由人工判断；重合率照常计入。

--json 输出字段：
    file、char_count（稿件正文字数）、min_len、threshold_percent（30）、
    references：每个参照一项，含 file、char_count、overlap_chars、ratio（0～1，保留 4 位小数）、
        percent（百分数，保留 1 位小数）、over_threshold（是否超过 30%）、
        fragments（重合片段：length、draft_line、draft_end_line、ref_line、ref_end_line、
        text（稿件原文）、ref_text（参照原文）、quoted）；
    combined：全部参照合计的 overlap_chars、ratio、percent；
    paragraphs：高重合段落，含 line、char_count、overlap_chars、ratio、percent、references（参照序号，从 1 起）、excerpt；
    over_threshold：任一参照超过 30% 时为 true；notes：提示信息。

退出码：0 正常；--strict 时任一参照的重合率超过 30% 返回 1；输入错误（文件不存在、格式不支持、读取失败、参数错误）返回 2。
只依赖标准库；读取 .docx 需要 python-docx，未安装时给出安装提示并返回 2。不访问网络，不写任何文件。
"""
from __future__ import annotations

import argparse
import bisect
import json
import os
import re
import sys
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

BOUNDARY = "【起草说明】"
# 分界行：去掉行首空白后以【起草说明】开头；兼容 Markdown 标题、加粗包裹的写法。
# 与 gongwen_lint.py 的 DRAFT_NOTE_RE、fact_diff.py 保持一致。
DRAFT_NOTE_RE = re.compile(r"^\s*(?:#{1,6}\s*)?(?:(?:\*\*|__)\s*)?" + re.escape(BOUNDARY))
SUPPORTED_SUFFIXES = (".txt", ".md", ".docx")

DEFAULT_MIN_LEN = 15
MIN_LEN_RANGE = (4, 1000)
THRESHOLD = 0.30          # --strict 的判定线：重合率超过 30%
PARAGRAPH_RATIO = 0.5     # 高重合段落：本段一半以上的字与参照重合
MAX_POSITIONS = 100       # 同一个片段在参照中最多记录的出现位置（防止重复文字拖慢比对）
MAX_MATCHES = 20000       # 每个参照最多收集的候选片段；超出时片段清单截断，重合率不受影响
TEXT_FRAGMENT_LIMIT = 30  # 文本输出时每个参照最多列出的片段数
EXCERPT_CHARS = 60        # 文本输出时片段的显示长度
PARA_EXCERPT_CHARS = 24
CHUNK = 64                # 片段向后延伸时按块比较

# 稿件中不参与比对的位置：书名号内的文件名称、【】标记
EXCLUDE_RES = (
    re.compile(r"《[^《》\n]{1,200}》"),
    re.compile(r"【[^【】\n]{0,200}】"),
)
# 引号：片段整体落在同一对引号内时标为 quoted
QUOTE_RES = (
    re.compile(r"“[^“”\n]{1,2000}”"),
    re.compile(r"「[^「」\n]{1,2000}」"),
    re.compile(r'"[^"\n]{1,2000}"'),  # 半角双引号不分左右，同一行内从左到右两两配对
)


class InputError(Exception):
    """输入错误，退出码 2。"""


# ---------------------------------------------------------------------------
# 读取
# ---------------------------------------------------------------------------

def _decode(data: bytes, path: Path) -> str:
    for encoding in ("utf-8-sig", "gb18030"):
        try:
            text = data.decode(encoding)
        except UnicodeDecodeError:
            continue
        if encoding != "utf-8-sig":
            print(f"提示：{path.name} 不是 UTF-8 编码，已按 GB18030 读取。", file=sys.stderr)
        return text
    raise InputError(f"无法识别文件编码，请另存为 UTF-8：{path}")


def _read_docx(path: Path) -> str:
    try:
        import docx  # type: ignore
        from docx.table import Table  # type: ignore
        from docx.text.paragraph import Paragraph  # type: ignore
    except ImportError:
        raise InputError("读取 .docx 需要 python-docx，请先运行：pip install python-docx；"
                         "或把文件另存为 .txt 后再比对。")
    try:
        document = docx.Document(str(path))
    except Exception as exc:  # 文件损坏、不是 docx 等
        raise InputError(f"无法打开 docx 文件：{path}（{exc}）")
    lines: List[str] = []
    for child in document.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            lines.append(Paragraph(child, document).text)
        elif tag == "tbl":
            # 合并单元格会被逐格重复返回，按底层 tc 元素去重
            seen = set()
            for row in Table(child, document).rows:
                texts = []
                for cell in row.cells:
                    if cell._tc in seen:
                        continue
                    seen.add(cell._tc)
                    texts.append(cell.text.replace("\n", " "))
                lines.append("\t".join(texts))
    return "\n".join(lines)


def _long_path(path: Path) -> Path:
    """Windows 下完整路径接近 260 字符上限时加长路径前缀，避免"文件不存在"误报。"""
    if os.name != "nt":
        return path
    try:
        full = str(path.resolve())
    except OSError:
        return path
    prefix = "\\\\?\\"
    if len(full) < 240 or full.startswith(prefix):
        return path
    if full.startswith("\\\\"):
        return Path(prefix + "UNC\\" + full[2:])
    return Path(prefix + full)


def read_text(path: Path) -> str:
    """读取 .txt、.md、.docx 文件的全文，换行统一为 \\n。"""
    shown = path
    path = _long_path(path)
    if not path.exists():
        raise InputError(f"文件不存在：{shown}")
    if not path.is_file():
        raise InputError(f"不是文件：{shown}")
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise InputError(f"不支持的文件类型 {suffix or '（无扩展名）'}，只支持 .txt、.md、.docx：{shown}")
    if suffix == ".docx":
        text = _read_docx(path)
    else:
        try:
            data = path.read_bytes()
        except OSError as exc:
            raise InputError(f"读取失败：{shown}（{exc}）")
        text = _decode(data, path)
    return text.replace("\r\n", "\n").replace("\r", "\n")


def split_body(text: str) -> Tuple[str, Optional[int]]:
    """返回【起草说明】之前的正文，以及分界行的行号（没有分界行时为 None）。"""
    lines = text.split("\n")
    for index, line in enumerate(lines):
        if DRAFT_NOTE_RE.match(line):
            return "\n".join(lines[:index]), index + 1
    return text, None


# ---------------------------------------------------------------------------
# 规范化
# ---------------------------------------------------------------------------

def norm_char(ch: str) -> Optional[str]:
    """汉字、字母、数字返回规范化后的字符；空白、标点和其他符号返回 None。"""
    code = ord(ch)
    if 0xFF01 <= code <= 0xFF5E:  # 全角 ASCII 转半角
        ch = chr(code - 0xFEE0)
    category = unicodedata.category(ch)
    if category[0] not in ("L", "N"):
        return None
    return ch.lower() if ch.isascii() else ch


class Normalized:
    """规范化后的文字序列，以及每个字在原文中的位置和比对分段。"""

    def __init__(self, text: str, chars: str, offsets: List[int], segments: List[Tuple[int, int]]) -> None:
        self.text = text
        self.chars = chars
        self.offsets = offsets
        self.segments = segments
        self.line_starts = [0] + [m.end() for m in re.finditer("\n", text)]

    def line_of(self, offset: int) -> int:
        """原文偏移量所在的行号（从 1 起）。"""
        return bisect.bisect_right(self.line_starts, offset)

    def original(self, start: int, length: int) -> str:
        """规范化序列 [start, start+length) 对应的原文（含其间的标点）。"""
        begin = self.offsets[start]
        end = self.offsets[start + length - 1] + 1
        return self.text[begin:end]


def _spans(text: str, regexes: Sequence["re.Pattern[str]"]) -> List[Tuple[int, int]]:
    spans: List[Tuple[int, int]] = []
    for regex in regexes:
        spans.extend((m.start(), m.end()) for m in regex.finditer(text))
    spans.sort()
    return spans


def normalize(text: str, is_draft: bool = False) -> Normalized:
    """规范化文本。

    is_draft 为 True 时（稿件）：去掉书名号和【】内的文字，并在这些位置和换行处切断比对分段，
    使重合片段不跨段落，避免"……以上。二、重点任务"这类由相邻段落拼出的误报。
    参照文件不切分，稿件把参照的两段合成一段照抄时也能查出。
    """
    excluded = bytearray(len(text))
    if is_draft:
        for start, end in _spans(text, EXCLUDE_RES):
            excluded[start:end] = b"\x01" * (end - start)
        for match in re.finditer("\n", text):
            excluded[match.start()] = 1
    chars: List[str] = []
    offsets: List[int] = []
    segments: List[Tuple[int, int]] = []
    seg_start = 0
    for offset, ch in enumerate(text):
        if excluded[offset]:
            if len(chars) > seg_start:
                segments.append((seg_start, len(chars)))
            seg_start = len(chars)
            continue
        value = norm_char(ch)
        if value is None:
            continue
        chars.append(value)
        offsets.append(offset)
    if len(chars) > seg_start:
        segments.append((seg_start, len(chars)))
    return Normalized(text, "".join(chars), offsets, segments)


# ---------------------------------------------------------------------------
# 比对
# ---------------------------------------------------------------------------

def build_index(chars: str, min_len: int) -> Dict[str, List[int]]:
    """参照文本中每个长度为 min_len 的片段 → 出现位置。"""
    index: Dict[str, List[int]] = {}
    for pos in range(len(chars) - min_len + 1):
        gram = chars[pos:pos + min_len]
        found = index.get(gram)
        if found is None:
            index[gram] = [pos]
        elif len(found) < MAX_POSITIONS:
            found.append(pos)
    return index


def _extend(draft: str, a: int, seg_end: int, ref: str, j: int, k: int) -> int:
    """从已知相同的 k 个字向后延伸，返回最长相同长度。"""
    limit = min(seg_end - a, len(ref) - j)
    while k < limit:
        step = min(CHUNK, limit - k)
        if draft[a + k:a + k + step] == ref[j + k:j + k + step]:
            k += step
            continue
        while k < limit and draft[a + k] == ref[j + k]:
            k += 1
        break
    return k


def compare(draft: Normalized, ref: Normalized, min_len: int) -> Tuple[bytearray, List[Tuple[int, int, int]], bool]:
    """返回 (稿件逐字覆盖标记, 极大重合片段 [(稿件位置, 参照位置, 长度)], 片段清单是否截断)。

    覆盖标记由全部长度为 min_len 的相同片段求并集得到，与极大片段的并集相同，且不受截断影响。
    """
    n = len(draft.chars)
    diff = [0] * (n + 1)
    matches: List[Tuple[int, int, int]] = []
    truncated = False
    if len(ref.chars) < min_len:
        return bytearray(n), matches, truncated
    index = build_index(ref.chars, min_len)
    d, r = draft.chars, ref.chars
    for seg_start, seg_end in draft.segments:
        for a in range(seg_start, seg_end - min_len + 1):
            positions = index.get(d[a:a + min_len])
            if not positions:
                continue
            diff[a] += 1
            diff[a + min_len] -= 1
            if truncated:
                continue
            # 同一稿件位置在参照中多处出现时，只保留最长的一处
            best_j, best_len = -1, 0
            for j in positions:
                # 左边还能延伸的，不是片段起点，会在更靠前的位置统计
                if a > seg_start and j > 0 and d[a - 1] == r[j - 1]:
                    continue
                length = _extend(d, a, seg_end, r, j, min_len)
                if length > best_len:
                    best_j, best_len = j, length
            if best_j >= 0:
                matches.append((a, best_j, best_len))
                if len(matches) >= MAX_MATCHES:
                    truncated = True
    covered = bytearray(n)
    running = 0
    for pos in range(n):
        running += diff[pos]
        if running > 0:
            covered[pos] = 1
    return covered, matches, truncated


def select_fragments(matches: List[Tuple[int, int, int]], n: int) -> List[Tuple[int, int, int]]:
    """从候选片段中挑出用于展示的片段：长的优先，已被展示片段完全覆盖的不再列出。"""
    shown = bytearray(n)
    selected: List[Tuple[int, int, int]] = []
    for a, j, length in sorted(matches, key=lambda m: (-m[2], m[0], m[1])):
        if shown[a:a + length].count(0) == 0:
            continue
        shown[a:a + length] = b"\x01" * length
        selected.append((a, j, length))
    selected.sort(key=lambda m: (m[0], m[1]))
    return selected


def _in_quote(spans: List[Tuple[int, int]], begin: int, end: int) -> bool:
    for start, stop in spans:
        if start <= begin and end <= stop:
            return True
        if start > begin:
            break
    return False


def _ratio(part: int, whole: int) -> float:
    return round(part / whole, 4) if whole else 0.0


def _percent(part: int, whole: int) -> float:
    return round(part * 100.0 / whole, 1) if whole else 0.0


def build_report(draft_path: Path, ref_paths: Sequence[Path], min_len: int = DEFAULT_MIN_LEN) -> dict:
    notes: List[str] = []
    draft_text, boundary_line = split_body(read_text(draft_path))
    if boundary_line is not None:
        notes.append(f"稿件第 {boundary_line} 行为【起草说明】，其后内容未参与比对。")
    draft = normalize(draft_text, is_draft=True)
    n = len(draft.chars)
    if n == 0:
        notes.append("稿件正文为空，或只有书名号、【】内的文字，未作比对。")
    quote_spans = _spans(draft_text, QUOTE_RES)

    combined = bytearray(n)
    per_ref_cover: List[bytearray] = []
    references: List[dict] = []
    for number, ref_path in enumerate(ref_paths, start=1):
        ref = normalize(read_text(ref_path))
        if len(ref.chars) < min_len:
            notes.append(f"参照 {number}（{ref_path}）正文不足 {min_len} 字，未找到可比对的片段。")
        covered, matches, truncated = compare(draft, ref, min_len)
        if truncated:
            notes.append(f"参照 {number} 的重合片段过多，片段清单已截断；重合率不受影响。")
        per_ref_cover.append(covered)
        for pos in range(n):
            if covered[pos]:
                combined[pos] = 1
        fragments = []
        for a, j, length in select_fragments(matches, n):
            begin = draft.offsets[a]
            end = draft.offsets[a + length - 1] + 1
            ref_begin = ref.offsets[j]
            ref_end = ref.offsets[j + length - 1] + 1
            fragments.append({
                "length": length,
                "draft_line": draft.line_of(begin),
                "draft_end_line": draft.line_of(end - 1),
                "ref_line": ref.line_of(ref_begin),
                "ref_end_line": ref.line_of(ref_end - 1),
                "text": draft_text[begin:end],
                "ref_text": ref.text[ref_begin:ref_end],
                "quoted": _in_quote(quote_spans, begin, end),
            })
        overlap = covered.count(1)
        references.append({
            "file": str(ref_path),
            "char_count": len(ref.chars),
            "overlap_chars": overlap,
            "ratio": _ratio(overlap, n),
            "percent": _percent(overlap, n),
            "over_threshold": n > 0 and overlap / n > THRESHOLD,
            "fragments": fragments,
        })

    paragraphs = _paragraph_stats(draft, combined, per_ref_cover, min_len)
    total = combined.count(1)
    return {
        "file": str(draft_path),
        "char_count": n,
        "min_len": min_len,
        "threshold_percent": int(THRESHOLD * 100),
        "references": references,
        "combined": {"overlap_chars": total, "ratio": _ratio(total, n), "percent": _percent(total, n)},
        "paragraphs": paragraphs,
        "over_threshold": any(item["over_threshold"] for item in references),
        "notes": notes,
    }


def _paragraph_stats(draft: Normalized, combined: bytearray, per_ref: List[bytearray], min_len: int) -> List[dict]:
    """按稿件的行（段）统计重合字数，返回一半以上与参照重合的段落。"""
    totals: Dict[int, int] = {}
    overlaps: Dict[int, int] = {}
    refs: Dict[int, set] = {}
    for pos, offset in enumerate(draft.offsets):
        line = draft.line_of(offset)
        totals[line] = totals.get(line, 0) + 1
        if combined[pos]:
            overlaps[line] = overlaps.get(line, 0) + 1
            hit = refs.setdefault(line, set())
            for number, cover in enumerate(per_ref, start=1):
                if cover[pos]:
                    hit.add(number)
    lines = draft.text.split("\n")
    result = []
    for line in sorted(overlaps):
        overlap, total = overlaps[line], totals[line]
        if overlap < min_len or overlap / total < PARAGRAPH_RATIO:
            continue
        content = lines[line - 1].strip()
        excerpt = content if len(content) <= PARA_EXCERPT_CHARS else content[:PARA_EXCERPT_CHARS] + "……"
        result.append({
            "line": line,
            "char_count": total,
            "overlap_chars": overlap,
            "ratio": _ratio(overlap, total),
            "percent": _percent(overlap, total),
            "references": sorted(refs.get(line, ())),
            "excerpt": excerpt,
        })
    return result


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------

def _short(text: str, limit: int = EXCERPT_CHARS) -> str:
    text = text.replace("\n", " ")
    return text if len(text) <= limit else text[:limit] + "……"


def _lines(start: int, end: int) -> str:
    return f"第 {start} 行" if start == end else f"第 {start}～{end} 行"


def render_text(report: dict) -> str:
    out: List[str] = []
    out.append(f"重合检测：{report['file']}（正文 {report['char_count']} 字，最小片段 {report['min_len']} 字）")
    threshold = report["threshold_percent"]
    for number, ref in enumerate(report["references"], start=1):
        out.append("")
        out.append(f"参照 {number}：{ref['file']}")
        flag = f"，超过 {threshold}%" if ref["over_threshold"] else ""
        out.append(f"  重合率 {ref['percent']}%（{ref['overlap_chars']} 字）{flag}")
        fragments = ref["fragments"]
        if not fragments:
            out.append(f"  未发现不少于 {report['min_len']} 字的重合片段。")
            continue
        out.append(f"  重合片段 {len(fragments)} 处：")
        for index, item in enumerate(fragments[:TEXT_FRAGMENT_LIMIT], start=1):
            mark = "（引号内）" if item["quoted"] else ""
            out.append(
                f"    {index}. 稿件{_lines(item['draft_line'], item['draft_end_line'])}"
                f" ← 参照{_lines(item['ref_line'], item['ref_end_line'])}，{item['length']} 字{mark}："
                f"{_short(item['text'])}"
            )
        if len(fragments) > TEXT_FRAGMENT_LIMIT:
            out.append(f"    其余 {len(fragments) - TEXT_FRAGMENT_LIMIT} 处见 --json 输出。")
    if len(report["references"]) > 1:
        combined = report["combined"]
        out.append("")
        out.append(f"全部参照合计：重合率 {combined['percent']}%（{combined['overlap_chars']} 字）")
    out.append("")
    if report["paragraphs"]:
        out.append("高重合段落（本段一半以上的文字与参照重合）：")
        for item in report["paragraphs"]:
            refs = "、".join(str(x) for x in item["references"])
            out.append(f"  稿件第 {item['line']} 行：{item['percent']}%（参照 {refs}）：{item['excerpt']}")
    else:
        out.append("高重合段落：无。")
    for note in report["notes"]:
        out.append(f"提示：{note}")
    out.append("说明：重合率只反映文字相同的程度，是否属于照搬照抄或雷同由人工判断；"
               "带引号并注明出处的原文引用可以保留。")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

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
        prog="overlap_check.py",
        description="找出稿件与参照文件之间不少于指定字数的连续重合片段，计算重合率，列出高重合段落。"
                    "稿件只比对【起草说明】之前的正文；比对前去掉空白和标点；"
                    "稿件中书名号和【】内的文字不参与比对。",
        epilog="退出码：0 正常；1 使用 --strict 且任一参照的重合率超过 30%；2 输入错误。",
        formatter_class=_ChineseHelpFormatter,
        add_help=False,
    )
    parser._optionals.title = "选项"
    parser._positionals.title = "参数"
    parser.add_argument("-h", "--help", action="help", help="显示本帮助信息并退出")
    parser.add_argument("draft", help="稿件（.txt、.md、.docx）")
    parser.add_argument("references", nargs="+", help="参照文件，可以有多个（.txt、.md、.docx）")
    parser.add_argument("--min-len", type=int, default=DEFAULT_MIN_LEN, dest="min_len",
                        help=f"重合片段的最小字数，默认 {DEFAULT_MIN_LEN}（去掉空白和标点后计）")
    parser.add_argument("--json", action="store_true", help="以 JSON 格式输出")
    parser.add_argument("--strict", action="store_true", help="任一参照的重合率超过 30%% 时返回 1")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    _setup_console()
    parser = make_parser()
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:
        return 0 if exc.code == 0 else 2
    low, high = MIN_LEN_RANGE
    if not low <= args.min_len <= high:
        print(f"错误：--min-len 须为 {low}～{high} 之间的整数，当前为 {args.min_len}。", file=sys.stderr)
        return 2
    try:
        report = build_report(Path(args.draft), [Path(p) for p in args.references], args.min_len)
    except InputError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(render_text(report))
    if args.strict and report["over_threshold"]:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
