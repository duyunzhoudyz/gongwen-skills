#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""讲话稿时长与口语化检查（gongwen-speech）。

按语速估算讲话时长，并找出不便口头表达的长句、数字过密和过长的段落，列出引语和来源不明的引文，
供起草者校核时长、整理引语核对清单。

计数口径：从称谓行（如“同志们：”）算到结尾，不含空白，含标点；标题、署名、日期行，
以及【待补：……】【待核实：……】标记内的文字不计。遇到单独成行的【起草说明】即停止读取。

用法：
  python scripts/speech_check.py <讲话稿> [--minutes 分钟] [--rate 200] [--style 讲|念]
                                 [--max-sentence N] [--max-numbers N] [--max-para N]
                                 [--json] [--strict]

退出码：默认 0；--strict 时存在 P1 返回 1；输入错误返回 2。
只用标准库，不联网；读取 .docx 需要 python-docx。本脚本不写任何文件。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

VERSION = "1.0.0"

# 稿子用法对应的默认阈值（工具约定，见 references/oral-style.md）
STYLE_DEFAULTS: Dict[str, Dict[str, int]] = {
    "讲": {"max_sentence": 40, "max_numbers": 3, "max_para": 250},
    "念": {"max_sentence": 60, "max_numbers": 5, "max_para": 400},
}
DEFAULT_RATE = 200
HOUR_LIMIT_MINUTES = 60
# 称谓行最长字数：礼宾称谓可能点名多位领导和来宾，比标题区的行长
SALUTATION_MAX_CHARS = 150
TOLERANCE = 0.10
# 讲的稿超过这个字数仍只在开头称呼一次的，提示穿插称谓
SALUTATION_HINT_CHARS = 2000

# 【起草说明】分界行（兼容 Markdown 标题、加粗包裹的写法，与 gongwen_lint.py 一致）
DRAFT_NOTE_RE = re.compile(r"^\s*(?:#{1,6}\s*)?(?:(?:\*\*|__)\s*)?【起草说明】")
MARKER_RE = re.compile(r"【(待补|待核实)[：:][^】]*】")
SENTENCE_END = "。！？!?；;"
SENTENCE_RE = re.compile(r"[^。！？!?；;]+[。！？!?；;]*")
HEADER_END_PUNCT = set("。！？!?；;")
TOP_HEADING_RE = re.compile(r"^\s*[一二三四五六七八九十]{1,3}、")
LIST_NUMBER_RE = re.compile(r"^\s*(?:\d{1,2}[.．]|[（(]\d{1,2}[)）])")
# 数字：整数、小数、带千分位，带百分号或千分号；后接年、月、日的视为日期，不计。
# 日期范围（如“1—9月”“2026—2028年”）中前一个数也视为日期
NUMBER_RE = re.compile(r"\d+(?:[.,]\d+)*(?:\s*[%％‰])?")
DATE_TAIL_RE = re.compile(r"^\s*(?:[—–\-－～~至到]\s*\d+\s*)?(年|月|日|号|时)")
# 引语：引号前紧挨着“指出、强调、说”等动词（可带冒号或逗号），
# 兼容“所指出的：”“强调的那样，”“重要论述：”“重要讲话：”等引法，
# 以及“强调，要“……””“指出，必须“……””这类动词后带逗号和连接词的部分引用
QUOTE_RE = re.compile(r"“([^”]{4,})”")
QUOTE_VERB_RE = re.compile(
    r"(指出|强调|说|讲|讲道|讲到|写道|告诫|提出|提到|要求|号召|说过|曾说|曾经说过|论述|重要讲话|重要指示|批示|云|曰)"
    r"(?:的(?:那样)?)?(?:[：:]|[，,]\s*(?:要|必须|就是)?)?\s*$"
)
SOURCE_RE = re.compile(
    r"《[^》]+》|在[^。！？“”]{2,40}(讲话|会议|会上|大会|文件|通知|意见|决定|方案|纪要|报告|信|批示|指示|座谈|考察|论述|回信|贺信)"
)
# 来源不明的引子：P1 为假托古人、名人；P2 为俗语和“正如……所说”（可能指同场讲话，需人工核对）
VAGUE_SOURCE_P1_RE = re.compile(r"(古人云|古语云|古人说|古话说|名人说|有句名言|一位哲人|一位伟人|伟人说)")
VAGUE_SOURCE_P2_RE = re.compile(r"(俗话说|常言道|有句话说|有句话叫|有句老话|正如[^，。！？“”]{1,15}所说)")
SALUTATION_WORDS_RE = re.compile(r"同志们|同学们|老师们|朋友们|各位同事|各位代表|各位来宾|各位居民")


class _ChineseHelpFormatter(argparse.RawDescriptionHelpFormatter):
    """把 argparse 的 usage 前缀改为中文。"""

    def add_usage(self, usage, actions, groups, prefix=None):  # type: ignore[override]
        return super().add_usage(usage, actions, groups, "用法：" if prefix is None else prefix)


class InputError(Exception):
    """输入文件或参数有误。"""


def setup_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
        except Exception:  # 被重定向为 StringIO 等对象时没有 reconfigure
            pass


def nonspace_len(text: str) -> int:
    return sum(1 for ch in text if not ch.isspace())


def strip_markers(text: str) -> str:
    return MARKER_RE.sub("", text)


def clip(text: str, limit: int = 30) -> str:
    text = text.strip()
    return text if len(text) <= limit else text[:limit] + "……"


# ---------------------------------------------------------------------------
# 读取
# ---------------------------------------------------------------------------

def read_lines(path: Path) -> List[Tuple[int, str]]:
    """读取文件，返回 [(行号或段号, 文本)]。"""
    if not path.is_file():
        raise InputError(f"未找到文件：{path}")
    suffix = path.suffix.lower()
    if suffix in (".txt", ".md", ".markdown", ".text"):
        try:
            raw = path.read_bytes()
        except OSError as exc:
            raise InputError(f"无法读取文件：{exc}") from exc
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            try:
                text = raw.decode("gb18030")
            except UnicodeDecodeError as exc:
                raise InputError("无法识别文件编码，请另存为 UTF-8 后重试。") from exc
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        return list(enumerate(text.split("\n"), start=1))
    if suffix == ".docx":
        try:
            import docx  # type: ignore
        except ImportError as exc:
            raise InputError("读取 .docx 需要 python-docx，请先运行：pip install python-docx") from exc
        try:
            document = docx.Document(str(path))
        except Exception as exc:  # 损坏或加密的 docx
            raise InputError(f"无法打开 docx 文件：{exc}") from exc
        return [(i, p.text) for i, p in enumerate(document.paragraphs, start=1)]
    raise InputError(f"不支持的文件类型：{path.suffix or '无扩展名'}。只支持 .txt、.md、.docx。")


# ---------------------------------------------------------------------------
# 结构识别
# ---------------------------------------------------------------------------

class Speech:
    """讲话稿正文：识别标题区、称谓行，按段落保存。"""

    def __init__(self, lines: Sequence[Tuple[int, str]]) -> None:
        self.truncated_line: Optional[int] = None
        paras: List[Tuple[int, str]] = []
        for number, text in lines:
            if DRAFT_NOTE_RE.match(text):
                self.truncated_line = number
                break
            if text.strip():
                paras.append((number, text.strip()))
        self.paras = paras
        self.header: List[Tuple[int, str]] = []
        self.salutation: Optional[Tuple[int, str]] = None
        start = 0
        for index, (number, text) in enumerate(paras[:8]):
            plain = strip_markers(text)
            no_end_punct = not any(ch in HEADER_END_PUNCT for ch in plain)
            not_heading = not TOP_HEADING_RE.match(text)
            if (no_end_punct and not_heading and plain.rstrip().endswith(("：", ":"))
                    and nonspace_len(plain) <= SALUTATION_MAX_CHARS):
                self.salutation = (number, text)
                start = index
                break
            if not (nonspace_len(plain) <= 60 and no_end_punct and not_heading):
                start = index
                break
            self.header.append((number, text))
            start = index + 1
        self.body: List[Tuple[int, str]] = paras[start:]

    def char_count(self) -> int:
        return sum(nonspace_len(strip_markers(text)) for _n, text in self.body)

    def marker_counts(self) -> Dict[str, int]:
        counts = {"待补": 0, "待核实": 0}
        for _n, text in self.paras:
            for match in MARKER_RE.finditer(text):
                counts[match.group(1)] += 1
        return counts

    def sections(self, rate: float) -> List[Dict[str, object]]:
        """按一级标题（一、二、……）分段统计字数；第一个一级标题之前记为“开场”。"""
        result: List[Dict[str, object]] = []
        current = {"heading": "开场", "line": self.body[0][0] if self.body else None, "chars": 0}
        for number, text in self.body:
            if TOP_HEADING_RE.match(text):
                if current["chars"] or current["heading"] != "开场":
                    result.append(current)
                current = {"heading": clip(text, 20), "line": number, "chars": 0}
            current["chars"] = int(current["chars"]) + nonspace_len(strip_markers(text))
        if current["chars"] or current["heading"] != "开场":
            result.append(current)
        for item in result:
            item["minutes"] = round(int(item["chars"]) / rate, 1)
        return result


# ---------------------------------------------------------------------------
# 检查
# ---------------------------------------------------------------------------

def make_issue(issue_id: str, level: str, category: str, line: Optional[int], excerpt: str,
               message: str, suggestion: str, basis: str) -> Dict[str, object]:
    return {
        "id": issue_id,
        "level": level,
        "category": category,
        "line": line,
        "excerpt": excerpt,
        "message": message,
        "suggestion": suggestion,
        "basis": basis,
    }


def count_numbers(text: str) -> int:
    count = 0
    for match in NUMBER_RE.finditer(text):
        if DATE_TAIL_RE.match(text[match.end():]):
            continue
        count += 1
    return count


def find_quotes(number: int, text: str) -> Tuple[List[Dict[str, object]], List[Dict[str, object]]]:
    """返回（引语清单，问题）。"""
    quotes: List[Dict[str, object]] = []
    issues: List[Dict[str, object]] = []
    for match in QUOTE_RE.finditer(text):
        before = text[:match.start()]
        if not QUOTE_VERB_RE.search(before):
            continue
        # 本句的起点：上一个句末标点之后
        sentence_start = max(before.rfind(ch) for ch in SENTENCE_END) + 1
        lead = before[sentence_start:].lstrip("”’）)」』 ，,、")
        after = text[match.end():match.end() + 40]
        has_source = bool(SOURCE_RE.search(lead))
        marked = "【待核实" in lead or "【待核实" in after.split("。")[0]
        content = match.group(1)
        quotes.append({
            "line": number,
            "lead": clip(lead, 30),
            "text": clip(content, 40),
            "has_source": has_source,
            "marked": marked,
        })
        vague = bool(VAGUE_SOURCE_P1_RE.search(lead) or VAGUE_SOURCE_P2_RE.search(lead))
        if not has_source and not marked and not vague:  # 来源不明的引子另由 QUOTE-002、003 报告
            issues.append(make_issue(
                "QUOTE-001", "P2", "引语", number, clip(lead + "“" + content + "”", 40),
                "引语未注明出处，也没有标【待核实】",
                "核对是否为用户提供的原文：是的注明出处，不是的删去或改为转述并标【待核实：请对照原文】",
                "references/_shared/guardrails.md 第二条第1、5款",
            ))
    for match in VAGUE_SOURCE_P1_RE.finditer(text):
        issues.append(make_issue(
            "QUOTE-002", "P1", "引语", number, clip(text[match.start():], 30),
            f"“{match.group(1)}”引出的话来源不明，讲话稿中的引文只能来自用户提供的原文",
            "有用户提供的原文和出处的，改为注明出处的引用；没有的删去引子，用自己的话直接说",
            "references/_shared/guardrails.md 第二条第5款",
        ))
    for match in VAGUE_SOURCE_P2_RE.finditer(text):
        issues.append(make_issue(
            "QUOTE-003", "P2", "引语", number, clip(text[match.start():], 30),
            f"“{match.group(1)}”引出的话请核对来源：是否为用户提供的原文，或同场讲话中确实说过",
            "核对属实的保留；核对不了的删去引子，用自己的话直接说",
            "references/_shared/guardrails.md 第二条第5款",
        ))
    return quotes, issues


def check(speech: Speech, minutes: Optional[float], rate: float, style: str,
          max_sentence: int, max_numbers: int, max_para: int) -> Dict[str, object]:
    issues: List[Dict[str, object]] = []
    quotes: List[Dict[str, object]] = []
    chars = speech.char_count()
    est = round(chars / rate, 1)
    first_line = speech.body[0][0] if speech.body else None
    basis_time = "工具约定：按每分钟约 200 字换算，误差±10%（references/oral-style.md 第五部分）"

    target_range: Optional[List[int]] = None
    within: Optional[bool] = None
    if minutes is not None:
        target = minutes * rate
        low, high = int(round(target * (1 - TOLERANCE))), int(round(target * (1 + TOLERANCE)))
        target_range = [low, high]
        within = low <= chars <= high
        if not within:
            direction = "超出" if chars > high else "不足"
            if chars > high:
                tip = "按铺垫、重复表态、数字罗列、次要事例的顺序删减"
            else:
                tip = "只展开要点清单中的事实；素材不够的按实际篇幅交付，在起草说明写明差距"
            issues.append(make_issue(
                "TIME-001", "P1", "时长", first_line, f"正文约 {chars} 字",
                f"预计 {est} 分钟，目标 {minutes:g} 分钟（{low}～{high} 字），篇幅{direction}",
                tip, basis_time,
            ))
    if est > HOUR_LIMIT_MINUTES:
        issues.append(make_issue(
            "TIME-002", "P1", "时长", first_line, f"正文约 {chars} 字",
            f"预计 {est} 分钟，超过 1 小时",
            "主要负责同志讲话、会议主报告不超过 1 小时，请压缩；用户坚持的，在起草说明提示",
            "《整治形式主义为基层减负若干规定》第6条",
        ))

    if speech.salutation is None:
        issues.append(make_issue(
            "ORAL-004", "P2", "称谓", first_line, clip(speech.body[0][1]) if speech.body else "",
            "开头没有找到单独成段的称谓（如“同志们：”）",
            "在开场前加称谓并单独成段；致辞按名单的礼宾顺序称呼",
            "GB/T 15834-2011 4.7.3.4；references/types/ceremony.md 第五部分",
        ))

    later_salutations = 0
    for index, (number, text) in enumerate(speech.body):
        plain = strip_markers(text)
        if speech.salutation is not None and index > 0:
            later_salutations += len(SALUTATION_WORDS_RE.findall(plain))
        para_len = nonspace_len(plain)
        if para_len > max_para:
            issues.append(make_issue(
                "ORAL-003", "P2", "口语化", number, clip(text),
                f"本段 {para_len} 字，超过 {max_para} 字",
                "拆成两三段，每段讲一层意思", "工具约定（references/oral-style.md 第二、三部分）",
            ))
        numbers = count_numbers(LIST_NUMBER_RE.sub("", plain, count=1))
        if numbers > max_numbers:
            issues.append(make_issue(
                "ORAL-002", "P2", "口语化", number, clip(text),
                f"本段连用 {numbers} 个数字，超过 {max_numbers} 个，听众记不住",
                "留最能说明问题的一两个，其余删去或写进书面材料；讲的稿可取整",
                "工具约定（references/oral-style.md 第二部分第3条）",
            ))
        is_salutation = speech.salutation is not None and index == 0
        for match in ([] if is_salutation else SENTENCE_RE.finditer(plain)):  # 称谓行不按整句检查
            sentence = match.group(0).rstrip(SENTENCE_END)
            length = nonspace_len(sentence)
            if length > max_sentence:
                issues.append(make_issue(
                    "ORAL-001", "P2", "口语化", number, clip(sentence),
                    f"整句 {length} 字，超过 {max_sentence} 字，不便口头表达",
                    "拆成两三句，一句一个意思；长定语拆开，动词提前",
                    "工具约定（references/oral-style.md 第二部分第1、2条）",
                ))
        found_quotes, quote_issues = find_quotes(number, text)
        quotes.extend(found_quotes)
        issues.extend(quote_issues)

    if (style == "讲" and speech.salutation is not None and chars > SALUTATION_HINT_CHARS
            and later_salutations == 0):
        issues.append(make_issue(
            "ORAL-005", "P2", "称谓", speech.salutation[0], clip(speech.salutation[1]),
            "全篇只在开头称呼一次",
            "讲的稿在段落转折、提出重要要求时用称谓唤起注意，全篇三到五次",
            "工具约定（references/oral-style.md 第二部分第4条）",
        ))

    summary = {"P0": 0, "P1": 0, "P2": 0}
    for issue in issues:
        summary[str(issue["level"])] += 1
    return {
        "char_count": chars,
        "rate": int(rate) if float(rate).is_integer() else rate,
        "style": style,
        "est_minutes": est,
        "target_minutes": minutes,
        "target_range": target_range,
        "within_target": within,
        "markers": speech.marker_counts(),
        "salutation": speech.salutation[1] if speech.salutation else None,
        "header": [text for _n, text in speech.header],
        "sections": speech.sections(rate),
        "quotes": quotes,
        "summary": summary,
        "issues": issues,
        "thresholds": {"max_sentence": max_sentence, "max_numbers": max_numbers, "max_para": max_para},
    }


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------

def render_text(result: Dict[str, object]) -> str:
    lines: List[str] = []
    style_name = "讲的稿" if result["style"] == "讲" else "念的稿"
    lines.append(f"讲话稿检查：{result['file']}（{style_name}，语速每分钟约 {result['rate']:g} 字）")
    head = f"正文约 {result['char_count']} 字，预计 {result['est_minutes']} 分钟"
    if result["target_minutes"] is not None:
        low, high = result["target_range"]  # type: ignore[misc]
        if result["within_target"]:
            verdict = "在范围内"
        else:
            verdict = "篇幅超出" if result["char_count"] > high else "篇幅不足"
        head += f"；目标 {result['target_minutes']:g} 分钟（{low}～{high} 字），{verdict}"
    lines.append(head + "。")
    markers = result["markers"]  # type: ignore[assignment]
    if markers["待补"] or markers["待核实"]:  # type: ignore[index]
        lines.append(f"待补 {markers['待补']} 处、待核实 {markers['待核实']} 处，补齐后字数会变化。")  # type: ignore[index]
    if result.get("truncated_line"):
        lines.append(f"第 {result['truncated_line']} 行为【起草说明】，其后内容未检查。")
    sections = result["sections"]  # type: ignore[assignment]
    if sections:
        parts = [f"{s['heading']} {s['chars']} 字（{s['minutes']} 分钟）" for s in sections]  # type: ignore[index]
        lines.append("分段：" + "；".join(parts))
    summary = result["summary"]  # type: ignore[assignment]
    lines.append(f"问题：P1 {summary['P1']} 条，P2 {summary['P2']} 条")  # type: ignore[index]
    for issue in result["issues"]:  # type: ignore[union-attr]
        where = f"第{issue['line']}行" if issue["line"] else ""
        lines.append(f"  [{issue['level']}] {issue['id']} {where}：{issue['message']}。原文：{issue['excerpt']}")
        lines.append(f"      建议：{issue['suggestion']}（依据：{issue['basis']}）")
    quotes = result["quotes"]  # type: ignore[assignment]
    if quotes:
        lines.append(f"引语 {len(quotes)} 处，请逐条对照用户提供的原文和出处，整理进引语核对清单：")
        for item in quotes:  # type: ignore[union-attr]
            if item["marked"]:
                state = "已标待核实"
            elif item["has_source"]:
                state = "已写出处"
            else:
                state = "未写出处"
            lines.append(f"  第{item['line']}行：{item['lead']}“{item['text']}”（{state}）")
    return "\n".join(lines)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="speech_check.py",
        description="讲话稿时长与口语化检查：估算时长，找出长句、数字过密的段落和来源不明的引文。",
        epilog="退出码：0 正常；--strict 且存在 P1 时为 1；输入错误为 2。docx 输入时“行号”为段落序号。",
        add_help=False,
        formatter_class=_ChineseHelpFormatter,
    )
    parser._positionals.title = "位置参数"
    parser._optionals.title = "选项"
    parser.add_argument("-h", "--help", action="help", help="显示本帮助信息并退出")
    parser.add_argument("file", help="讲话稿（.txt、.md、.docx），遇到【起草说明】行即停止读取")
    parser.add_argument("--minutes", type=float, default=None, help="目标时长（分钟），给出时检查字数是否在目标的±10%%以内")
    parser.add_argument("--rate", type=float, default=DEFAULT_RATE, help="语速，每分钟字数，默认 200")
    parser.add_argument("--style", choices=("讲", "念"), default="讲",
                        help="稿子用法：讲（讲的稿，默认）或念（念的稿，阈值放宽）")
    parser.add_argument("--max-sentence", type=int, default=None, help="整句字数上限，默认讲 40、念 60")
    parser.add_argument("--max-numbers", type=int, default=None, help="每段数字个数上限，默认讲 3、念 5")
    parser.add_argument("--max-para", type=int, default=None, help="每段字数上限，默认讲 250、念 400")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    parser.add_argument("--strict", action="store_true", help="存在 P1 时返回 1")
    parser.add_argument("--version", action="version", version=f"%(prog)s {VERSION}", help="显示版本号")
    return parser


def run(args: argparse.Namespace) -> Dict[str, object]:
    if args.minutes is not None and args.minutes <= 0:
        raise InputError("--minutes 必须大于 0")
    if args.rate <= 0:
        raise InputError("--rate 必须大于 0")
    defaults = STYLE_DEFAULTS[args.style]
    max_sentence = args.max_sentence if args.max_sentence is not None else defaults["max_sentence"]
    max_numbers = args.max_numbers if args.max_numbers is not None else defaults["max_numbers"]
    max_para = args.max_para if args.max_para is not None else defaults["max_para"]
    for name, value in (("--max-sentence", max_sentence), ("--max-numbers", max_numbers),
                        ("--max-para", max_para)):
        if value <= 0:
            raise InputError(f"{name} 必须大于 0")
    path = Path(args.file)
    speech = Speech(read_lines(path))
    if not speech.body:
        raise InputError("没有读到正文内容。")
    result = check(speech, args.minutes, args.rate, args.style, max_sentence, max_numbers, max_para)
    ordered: Dict[str, object] = {"file": path.name, "truncated_line": speech.truncated_line}
    ordered.update(result)
    return ordered


def main(argv: Optional[Sequence[str]] = None) -> int:
    setup_console()
    parser = build_parser()
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:  # argparse 参数错误返回 2，--help 返回 0
        return int(exc.code or 0)
    try:
        result = run(args)
    except InputError as exc:
        print(f"输入错误：{exc}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(render_text(result))
    summary = result["summary"]  # type: ignore[assignment]
    if args.strict and summary["P1"]:  # type: ignore[index]
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
