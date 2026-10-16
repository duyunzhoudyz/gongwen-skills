# -*- coding: utf-8 -*-
"""项目申报书形式审查辅助检查。

本脚本属于 gongwen-university（行业包·高校），只在本机运行，不访问网络，不写任何文件。
它只查形式问题，不评价选题和论证，也不改写申报书。

用法：
    python scripts/proposal_check.py <申报书> [--blind] [--names 词1,词2] [--names-file 名单.txt]
        [--max-chars 字数] [--section 栏目=字数 ...] [--title 项目名称]
        [--count word|chars|hanzi] [--json] [--strict]

检查项：
    匿名（--blind，用于活页等匿名评审材料）：
        PRC-ID-001  P0  出现 --names 提供的身份词（申请人、成员姓名，所在单位）
        PRC-ID-002  P1  可推断身份的自称：本人、笔者、拙著、我校、本单位、课题组前期等
                        （样本、基本、日本、资本、成本、文本、版本等复合词中的"本"不报；
                        "据我所知""我所见""我所谓"一类学术表述中的"我所"不报）
        PRC-ID-003  P1  项目批准号
        PRC-ID-004  P1  电子邮箱
        PRC-ID-005  P2  网址（个人主页、课题组网站会暴露身份；文献链接可保留）
        PRC-ID-006  P2  机构名称汇总（请确认不是本人单位，也不能推断本人单位）。机构名从"的""了""和""与"
                        等虚词和"借鉴""参考"等引导词之后起算，只列机构名本身；只剩"大学"一类通称的不计
        PRC-ID-007  P2  前期成果自述：出版专著《……》、主持完成……课题（可据此查到申请人）
        PRC-PRV-001 P0  身份证号、手机号
    隐私（非匿名材料）：
        PRC-PRV-002 P1  身份证号、手机号出现在没有"电话""手机""身份证""证件""联系"字样的行
    字数：
        PRC-LEN-001 P0  全文超过 --max-chars
        PRC-LEN-002 P0  栏目超过 --section 规定的字数
        PRC-LEN-003 P2  --section 指定的栏目没有找到
    填写：
        PRC-FIL-001 P1  【待补】【待核实】标记未补齐
        PRC-FIL-002 P1  占位符和模板提示语：XXX、××、某某、（略）、请填写、此处填写、：待定
                        （--blind 时"××""某某"多为有意隐去，降为 P2，提示"如是有意匿名可保留"）
        PRC-FIL-003 P1  栏目为空
    一致性（--title）：
        PRC-TTL-001 P0  "项目名称""课题名称"栏填写的名称与 --title 不一致
        PRC-TTL-002 P2  全文没有找到项目名称
    残留：
        PRC-RES-001 P1  AI 对话残留（作为 AI、以下是为您、希望对您有帮助）
        PRC-RES-002 P1  Markdown 符号（输入为 .md 时不查）

字数口径（--count）：
    word   默认。按 Word"字数"的近似口径：每个汉字、每个全角标点各算 1，连续的西文字母或数字算 1。
    chars  不计空白的字符数，与 Word"字符数（不计空格）"接近，最严格。
    hanzi  只数汉字。
    各口径与申报系统的统计可能有少量差异，临近上限的以申报系统统计为准。

栏目识别：一行不超过 40 字，去掉序号（一、（一）1. 第一部分）和"（限300字）"一类括注后，
    以已知栏目名（选题说明、选题依据、研究内容、创新之处、预期成果、参考文献、立项依据、研究方案等）
    或 --section 指定的栏目名开头。
    "栏目名：内容"写在同一行的，只在带序号（如"一、选题说明：……"）或该栏目名由 --section 指定时
    认作新栏目，冒号后的文字计入该栏目；否则整行当正文，计入上一栏目。这样正文里的"研究方法：采用……"
    不会被切成新栏目，代价是不带序号的"研究内容：……"会算进上一栏目的字数。需要分开统计的，
    用 --section 指定该栏目（如 --section 研究内容=3000）后重跑，或人工统计。

只检查【起草说明】之前的内容。身份词、号码、邮箱、网址、批准号和前期成果的书名在输出中一律打码，
可以直接贴进审查意见。

退出码：0 正常；1 使用 --strict 且存在 P0；2 输入错误（含文件没有可检查的内容）。
只依赖标准库；读取 .docx 需要 python-docx，未安装时给出安装提示并返回 2。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

BOUNDARY = "【起草说明】"
# 与 gongwen_lint.py、fact_diff.py 的分界行写法一致：兼容 Markdown 标题、加粗包裹
DRAFT_NOTE_RE = re.compile(r"^\s*(?:#{1,6}\s*)?(?:(?:\*\*|__)\s*)?" + re.escape(BOUNDARY))
SUPPORTED_SUFFIXES = (".txt", ".md", ".docx")
LEVELS = ("P0", "P1", "P2")
EXCERPT_RADIUS = 12
MAX_ORG_NAMES = 5
HEADING_MAX_LEN = 40

BASIS_BLIND = "当年申报公告的匿名要求（如2026年国家社科基金年度项目申报公告：活页不得出现任何可能透露申请人身份的信息）"
BASIS_LIMIT = "当年申报通知的字数上限（本次检查按 --max-chars、--section 设定）"
BASIS_PRIVACY = "references/_shared/guardrails.md 第一条第3款"
BASIS_FILL = "当年申请书模板的填写要求；references/templates/shenbao-xingshi.md 第三部分"
BASIS_TITLE = "当年申报通知：申请书、活页、系统填报的项目名称一致"
BASIS_RESIDUE = "references/_shared/guardrails.md 第三条；国科金发计〔2026〕2号 一（二）8"

# 申报书常见栏目名（长的在前，避免"研究内容"先于"研究内容和方法"命中）
KNOWN_SECTIONS = sorted({
    "选题说明", "选题依据", "研究内容", "创新之处", "预期成果", "参考文献", "主要参考文献",
    "研究基础", "前期研究成果", "前期相关研究成果", "课题设计论证", "立项依据", "研究目标",
    "研究方案", "研究计划", "年度研究计划", "可行性分析", "特色与创新之处", "研究基础与工作条件",
    "工作基础", "工作条件", "经费预算", "预算说明", "学术价值和应用价值", "研究思路和方法",
    "研究思路", "研究方法", "拟解决的关键科学问题", "预期研究结果", "其他需要说明的情况",
}, key=len, reverse=True)

# 模板中只作上级标题、下面直接接分栏目的栏目，为空不提示
PARENT_SECTIONS = {"课题设计论证", "研究基础与工作条件"}

SERIAL_RE = re.compile(
    r"^(?:第[一二三四五六七八九十]+(?:部分|章|节)|[一二三四五六七八九十]+[、．.]|"
    r"[（(]\s*[一二三四五六七八九十]+\s*[）)]|[（(]\s*\d+\s*[）)]|\d+(?:\.\d+)*[、．.]?(?=\s|[^\d.]))\s*"
)
NOTE_RE = re.compile(r"[（(][^）)]{0,20}[）)]")

SELF_REF_RE = re.compile(
    r"本人|笔者|拙著|拙作|拙文|我校|本校|我院|本院|我单位|本单位|我所|本所|本团队|我们团队|"
    r"申请人前期|课题组前期|本课题组前期"
)
# "我所"后面是这些字时是"据我所知"一类的助词结构，不是"研究所"的简称
WOSUO_PARTICLE_NEXT = frozenset("知见说谓述言想理认")
# "本"字开头的自称前面是这些字时，多半是复合词的一部分（样本单位、样本校、基本人权、日本人口、
# 资本所得、成本所占、文本院落、版本校对），不报
BEN_COMPOUND_PREFIX = frozenset("样基日资成根文版脚剧书原账课范标副底蓝稿抄话股绘译刻母")
# 但前两个字是这些词时，"本"仍是自称（完成本校、论文本人、出版本人、同样本人、近日本院）
SELF_REF_OVERRIDE_WORDS = frozenset({
    "完成", "形成", "构成", "促成", "达成", "组成", "造成", "变成", "生成", "建成", "做成", "写成",
    "制成", "编成", "集成", "养成", "合成", "赞成", "论文", "出版", "再版", "同样", "这样", "那样",
    "一样", "今日", "当日", "次日", "近日", "昨日", "明日", "平日", "每日", "投资", "出资", "工资", "融资",
})
# 前期成果自述：书名、课题名可据此查到申请人
PRIOR_WORK_RE = re.compile(
    r"(?:出版|撰写|主编|编著)了?(?:个人)?(?:学术)?(?:专著|著作|译著|教材)|主持(?:完成|承担|过)|"
    r"主持(?!人)了?[^，。；;\n]{0,12}?(?:国家|教育部|省级|省部级|全国)[^，。；;\n]{0,20}?(?:项目|课题)"
)
WORK_TITLE_RE = re.compile(r"((?:专著|著作|译著|教材|拙著|拙作)\s*《)[^》\n]{1,60}(》)")
GRANT_RE = re.compile(
    r"(?:批准号|项目编号|项目号|资助号|基金编号)\s*[:：为]?\s*[A-Za-z0-9&\-]{5,}|"
    r"(?<![A-Za-z0-9])\d{2}(?:&ZD|[A-Z]{1,4})\d{3,4}(?![A-Za-z0-9])"
)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
URL_RE = re.compile(r"https?://[^\s，。；）)]+|www\.[^\s，。；）)]+")
CONTACT_RE = re.compile(f"(?P<email>{EMAIL_RE.pattern})|(?P<url>{URL_RE.pattern})")
NAME_TOKEN = "〔身份词〕"
EMAIL_TOKEN = "〔邮箱〕"
URL_TOKEN = "〔网址〕"
TITLE_TOKEN = "〔书名〕"
GRANT_TOKEN = "〔批准号〕"
ORG_RE = re.compile(r"[\u4e00-\u9fa5]{2,12}?(?:大学(?!生)|学院(?!派)|研究院|研究所|研究中心|实验室)")
ID_RE = re.compile(r"(?<![0-9])\d{17}[0-9Xx](?![0-9])")
PHONE_RE = re.compile(r"(?<![0-9])1[3-9]\d{9}(?![0-9])")
NUMBER_FIELD_RE = re.compile(r"电话|手机|身份证|证件|联系")
ORG_SUFFIX_RE = re.compile(r"(?:大学|学院|研究院|研究所|研究中心|实验室)$")
# 机构名前面常连着的虚词和引导词，摘录从最后一个之后起算（"本研究借鉴了清华大学"只取"清华大学"）
# "据""联合"不在其中：避免切坏"大数据研究院""联合国大学"
ORG_CUT_CHARS = "的了和与及在从向或"
ORG_LEAD_RE = re.compile(r"^(?:本研究|本文|借鉴|参考|参照|根据|依据|引用|来自|选取|包括|例如|比如)+")
ORG_PREFIX_MAX = 10
MARKER_RE = re.compile(r"【(?:待补|待核实)[：:][^】]*】")
# 匿名材料中的"××""某某"多为有意隐去，--blind 时降为 P2
BLIND_SOFT_PLACEHOLDER_RE = re.compile(r"^(?:××+|某某)$")
PLACEHOLDER_RE = re.compile(
    r"(?<![A-Za-z])(?:XX+|xx+|××+)(?![A-Za-z])|某某|（略）|\(略\)|请填写|此处填写|在此填写|"
    r"[:：]\s*待定\s*$|（待定）"
)
TITLE_FIELD_RE = re.compile(r"(?:项目|课题)(?:名称|题目)\s*(?:[:：]|\t)\s*(.+)$")
FIELD_CUT_RE = re.compile(r"\s{2,}|\t|\u3000|(?:申请人|负责人|学科分类|项目类别)\s*[:：]")
AI_RESIDUE_RE = re.compile(
    r"作为(?:一个)?(?:AI|人工智能)|以下是(?:为您|我为你|根据您)|希望(?:以上|这些)?(?:内容)?对您有(?:所)?帮助|"
    r"如需(?:进一步)?(?:修改|调整)[，,]?请(?:随时)?告诉我"
)
MARKDOWN_RE = re.compile(r"^\s{0,3}#{1,6}\s|\*\*[^*\n]+\*\*|^\s*[-*+]\s+\S|^\s*>\s")
CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")
ASCII_TOKEN_RE = re.compile(r"[A-Za-z0-9]+")
NAME_SPLIT_RE = re.compile(r"[,，、;；\s]+")


class InputError(Exception):
    """输入错误，对应退出码 2。"""


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
                         "或把申报书另存为 .txt 后再检查。")
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
            # 合并单元格会被 python-docx 重复返回，按底层 tc 元素去重
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


def read_text(path: Path) -> str:
    """读取 .txt、.md、.docx 文件全文。"""
    if not path.exists():
        raise InputError(f"文件不存在：{path}")
    if not path.is_file():
        raise InputError(f"不是文件：{path}")
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise InputError(f"不支持的文件类型 {suffix or '（无扩展名）'}，只支持 .txt、.md、.docx：{path}")
    if suffix == ".docx":
        return _read_docx(path)
    try:
        data = path.read_bytes()
    except OSError as exc:
        raise InputError(f"读取失败：{path}（{exc}）")
    return _decode(data, path)


# ---------------------------------------------------------------------------
# 计数与工具函数
# ---------------------------------------------------------------------------

def count_text(text: str) -> Dict[str, int]:
    """返回三种口径的字数。"""
    chars = sum(1 for ch in text if not ch.isspace())
    hanzi = len(CJK_RE.findall(text))
    ascii_tokens = ASCII_TOKEN_RE.findall(text)
    ascii_chars = sum(len(tok) for tok in ascii_tokens)
    # Word 字数近似：非空白字符中，去掉西文字母数字后逐个计数，再加上西文词数
    word = chars - ascii_chars + len(ascii_tokens)
    return {"word": word, "chars": chars, "hanzi": hanzi}


def mask_number(value: str) -> str:
    """号码打码：保留前 3 位和后 4 位。"""
    if len(value) <= 7:
        return "*" * len(value)
    return value[:3] + "*" * (len(value) - 7) + value[-4:]


def make_excerpt(line: str, start: int, end: int, replace: Optional[str] = None) -> str:
    """截取命中处前后文；replace 不为空时用它替换命中片段。"""
    left = max(0, start - EXCERPT_RADIUS)
    right = min(len(line), end + EXCERPT_RADIUS)
    hit = line[start:end] if replace is None else replace
    text = line[left:start] + hit + line[end:right]
    text = text.strip()
    if left > 0:
        text = "……" + text
    if right < len(line):
        text = text + "……"
    return text


def mask_names(line: str, names: Sequence[str]) -> str:
    """把身份词替换为"〔身份词〕"。"""
    for name in names:
        line = line.replace(name, NAME_TOKEN)
    return line


def mask_numbers(line: str) -> str:
    """号码打码，长度不变，命中位置仍可对应。"""
    line = ID_RE.sub(lambda m: mask_number(m.group(0)), line)
    return PHONE_RE.sub(lambda m: mask_number(m.group(0)), line)


def mask_text(line: str, names: Sequence[str]) -> str:
    """号码以外的打码：邮箱、网址、批准号整个替换（先于身份词，避免邮箱被拆开后漏出片段），
    再换身份词和前期成果的书名。"""
    line = CONTACT_RE.sub(lambda m: EMAIL_TOKEN if m.group("email") else URL_TOKEN, line)
    line = GRANT_RE.sub(GRANT_TOKEN, line)
    line = mask_names(line, names)
    return WORK_TITLE_RE.sub(lambda m: m.group(1) + TITLE_TOKEN + m.group(2), line)


def mask_line(line: str, names: Sequence[str]) -> str:
    """把身份词、邮箱、网址、书名和号码打码，避免在输出中照抄。"""
    return mask_numbers(mask_text(line, names))


def is_compound_self_ref(text: str, start: int, word: str) -> bool:
    """命中的自称其实是复合词的一部分（样本单位、基本人权、自我校正）时返回 True，不报。"""
    if word.startswith("本") and start >= 1 and text[start - 1] in BEN_COMPOUND_PREFIX:
        return text[max(0, start - 2):start] not in SELF_REF_OVERRIDE_WORDS
    if word == "我校" and start >= 1 and text[start - 1] == "自" and text[start + 2:start + 3] in ("正", "准", "验", "对"):
        return True
    if word == "我所" and text[start + 2:start + 3] in WOSUO_PARTICLE_NEXT:
        return True
    return False


def org_display_name(match_text: str) -> Optional[str]:
    """从 ORG_RE 的命中中取机构名本身；只剩"大学""学院"一类通称的返回 None。"""
    suffix = ORG_SUFFIX_RE.search(match_text)
    if not suffix:
        return match_text
    prefix, tail = match_text[:suffix.start()], suffix.group(0)
    cut = max(prefix.rfind(ch) for ch in ORG_CUT_CHARS)
    if cut >= 0:
        prefix = prefix[cut + 1:]
    prefix = ORG_LEAD_RE.sub("", prefix)[-ORG_PREFIX_MAX:]
    if not prefix:
        return None
    return prefix + tail


def parse_names(raw: Optional[str], names_file: Optional[str]) -> List[str]:
    names: List[str] = []
    if raw:
        names.extend(NAME_SPLIT_RE.split(raw))
    if names_file:
        content = read_text(Path(names_file))
        for line in content.splitlines():
            names.extend(NAME_SPLIT_RE.split(line))
    cleaned = []
    for name in names:
        name = name.strip()
        if len(name) >= 2 and name not in cleaned:
            cleaned.append(name)
    # 长的在前，避免"某大学"先替换掉"某大学某学院"的一部分
    return sorted(cleaned, key=len, reverse=True)


def parse_sections(items: Optional[Sequence[str]]) -> Dict[str, int]:
    limits: Dict[str, int] = {}
    for item in items or []:
        if "=" not in item and "＝" not in item:
            raise InputError(f"--section 的写法应为 栏目名=字数，如 选题说明=300：{item}")
        name, _, value = item.replace("＝", "=").partition("=")
        name = name.strip()
        try:
            limit = int(value.strip())
        except ValueError:
            raise InputError(f"--section 的字数不是整数：{item}")
        if not name or limit <= 0:
            raise InputError(f"--section 的栏目名不能为空，字数须为正整数：{item}")
        limits[name] = limit
    return limits


def normalize_title(value: str) -> str:
    value = re.sub(r"\s+", "", value)
    value = value.strip("《》“”\"'‘’「」：:。．.；;，,")
    return value


# ---------------------------------------------------------------------------
# 栏目识别
# ---------------------------------------------------------------------------

def match_heading(line: str, section_names: Sequence[str],
                  inline_names: Sequence[str] = ()) -> Optional[Tuple[str, str]]:
    """判断一行是不是栏目标题。是的话返回 (栏目名, 同一行冒号后的内容)。

    标题单独成行的直接认定；"栏目名：内容"写在同一行的，只在带序号（如"一、选题说明："）
    或栏目名由 --section 指定时认定，避免把正文里的"研究方法：采用……"当成新栏目。
    """
    stripped = line.strip().lstrip("#").strip().strip("*").strip()
    if not stripped:
        return None
    body = SERIAL_RE.sub("", stripped, count=1).strip()
    has_serial = body != stripped
    for name in section_names:
        if not body.startswith(name):
            continue
        rest = body[len(name):].strip()
        rest = NOTE_RE.sub("", rest, count=1).strip() if rest.startswith(("（", "(")) else rest
        if rest == "" or rest in ("：", ":", "。", "."):
            if len(stripped) <= HEADING_MAX_LEN:
                return name, ""
            return None
        if rest[0] in "：:" and (has_serial or name in inline_names):
            return name, rest[1:].strip()
        return None
    return None


def split_sections(lines: Sequence[Tuple[int, str]], section_names: Sequence[str],
                   inline_names: Sequence[str] = ()) -> List[Dict[str, Any]]:
    sections: List[Dict[str, Any]] = []
    current: Optional[Dict[str, Any]] = None
    for no, line in lines:
        hit = match_heading(line, section_names, inline_names)
        if hit:
            name, inline = hit
            current = {"name": name, "line": no, "body": [inline] if inline else []}
            sections.append(current)
            continue
        if current is not None:
            current["body"].append(line)
    for sec in sections:
        sec["text"] = "\n".join(sec.pop("body"))
    return sections


# ---------------------------------------------------------------------------
# 检查
# ---------------------------------------------------------------------------

class Checker:
    def __init__(self, blind: bool, names: Sequence[str], is_markdown: bool) -> None:
        self.blind = blind
        self.names = list(names)
        self.is_markdown = is_markdown
        self.issues: List[Dict[str, Any]] = []

    def add(self, issue_id: str, level: str, category: str, line: Optional[int], excerpt: str,
            message: str, suggestion: str, basis: str) -> None:
        self.issues.append({
            "id": issue_id, "level": level, "category": category, "line": line,
            "excerpt": excerpt, "message": message, "suggestion": suggestion, "basis": basis,
        })

    def safe(self, line: str) -> str:
        return mask_line(line, self.names)

    def check_line(self, no: int, line: str) -> None:
        masked = self.safe(line)
        if self.blind:
            self._check_identity(no, line, masked)
        self._check_numbers(no, line)
        for m in MARKER_RE.finditer(masked):
            self.add("PRC-FIL-001", "P1", "填写", no, make_excerpt(masked, m.start(), m.end()),
                     "待补、待核实标记未补齐", "补齐后删除标记", BASIS_FILL)
        for m in PLACEHOLDER_RE.finditer(masked):
            hit = m.group(0).strip()
            if self.blind and BLIND_SOFT_PLACEHOLDER_RE.match(hit):
                self.add("PRC-FIL-002", "P2", "填写", no, make_excerpt(masked, m.start(), m.end()),
                         f"占位符：{hit}", "如是有意匿名可保留；如是模板占位符未替换，填写实际内容", BASIS_FILL)
                continue
            self.add("PRC-FIL-002", "P1", "填写", no, make_excerpt(masked, m.start(), m.end()),
                     f"占位符或模板提示语未替换：{hit}", "填写实际内容或删除提示语", BASIS_FILL)
        for m in AI_RESIDUE_RE.finditer(masked):
            self.add("PRC-RES-001", "P1", "残留", no, make_excerpt(masked, m.start(), m.end()),
                     "AI 对话残留", "删除；借助 AI 的内容须人工核实并按当年通知如实声明", BASIS_RESIDUE)
        if not self.is_markdown:
            m = MARKDOWN_RE.search(masked)
            if m:
                self.add("PRC-RES-002", "P1", "残留", no, make_excerpt(masked, m.start(), m.end()),
                         "Markdown 符号，粘贴进申报系统会原样显示", "删除 #、**、- 等符号", BASIS_RESIDUE)

    def _token_hits(self, masked: str, token: str, limit: int) -> List[int]:
        """打码后的行中 token 出现的位置，最多取原文实际命中的个数。"""
        positions: List[int] = []
        start = masked.find(token)
        while start != -1 and len(positions) < limit:
            positions.append(start)
            start = masked.find(token, start + len(token))
        return positions

    def _check_identity(self, no: int, line: str, masked: str) -> None:
        # 一律在打码后的行中定位和摘录，保证摘录不含身份词、邮箱、网址、书名和号码
        count = masked.count(NAME_TOKEN)
        if count:
            pos = masked.find(NAME_TOKEN)
            self.add("PRC-ID-001", "P0", "匿名", no, make_excerpt(masked, pos, pos + len(NAME_TOKEN)),
                     "出现申请人或成员的姓名、所在单位" + (f"（本行 {count} 处）" if count > 1 else ""),
                     "删除，或改为不指向具体人和单位的写法", BASIS_BLIND)
        for m in SELF_REF_RE.finditer(masked):
            if is_compound_self_ref(masked, m.start(), m.group(0)):
                continue
            self.add("PRC-ID-002", "P1", "匿名", no, make_excerpt(masked, m.start(), m.end()),
                     f"自称可推断身份：{m.group(0)}", "改为客观表述，如\"已有研究表明\"\"前期研究显示\"", BASIS_BLIND)
        contacts = list(CONTACT_RE.finditer(line))
        n_grant = len(GRANT_RE.findall(CONTACT_RE.sub(" ", line)))
        for pos in self._token_hits(masked, GRANT_TOKEN, n_grant):
            self.add("PRC-ID-003", "P1", "匿名", no, make_excerpt(masked, pos, pos + len(GRANT_TOKEN)),
                     "出现项目批准号，可据此查到负责人", "删除批准号", BASIS_BLIND)
        n_email = sum(1 for m in contacts if m.group("email"))
        for pos in self._token_hits(masked, EMAIL_TOKEN, n_email):
            self.add("PRC-ID-004", "P1", "匿名", no, make_excerpt(masked, pos, pos + len(EMAIL_TOKEN)),
                     "出现电子邮箱", "删除", BASIS_BLIND)
        for pos in self._token_hits(masked, URL_TOKEN, len(contacts) - n_email):
            self.add("PRC-ID-005", "P2", "匿名", no, make_excerpt(masked, pos, pos + len(URL_TOKEN)),
                     "出现网址：个人主页、课题组网站会暴露身份，公开文献链接可保留", "按行号找到原文，确认后处理",
                     BASIS_BLIND)
        for m in PRIOR_WORK_RE.finditer(masked):
            self.add("PRC-ID-007", "P2", "匿名", no, make_excerpt(masked, m.start(), m.end()),
                     "前期成果自述（专著、主持的课题），评审专家可据书名、课题名查到申请人",
                     "人工确认；能查到申请人的删除，或改为\"前期研究显示\"一类写法", BASIS_BLIND)

    def _check_numbers(self, no: int, line: str) -> None:
        # 在邮箱、网址、身份词打码后的行中找号码（打码不改长度），摘录取自完全打码的行
        named = mask_text(line, self.names)
        masked_line = mask_numbers(named)
        for regex, label in ((ID_RE, "身份证号"), (PHONE_RE, "手机号")):
            for m in regex.finditer(named):
                excerpt = make_excerpt(masked_line, m.start(), m.end())
                if self.blind:
                    self.add("PRC-PRV-001", "P0", "匿名", no, excerpt, f"匿名材料中出现{label}",
                             "删除", BASIS_BLIND)
                elif not NUMBER_FIELD_RE.search(line):
                    self.add("PRC-PRV-002", "P1", "隐私", no, excerpt,
                             f"{label}出现在非规定栏目", "只在申请书规定的联系方式栏填写，其他位置删除", BASIS_PRIVACY)

    def check_orgs(self, lines: Sequence[Tuple[int, str]]) -> None:
        found: List[str] = []
        first_line: Optional[int] = None
        total = 0
        for no, line in lines:
            masked = self.safe(line)
            for m in ORG_RE.finditer(masked):
                name = org_display_name(m.group(0))
                if name is None:
                    continue
                total += 1
                if first_line is None:
                    first_line = no
                if name not in found:
                    found.append(name)
        if total:
            shown = "、".join(found[:MAX_ORG_NAMES]) + ("等" if len(found) > MAX_ORG_NAMES else "")
            self.add("PRC-ID-006", "P2", "匿名", first_line, shown,
                     f"出现机构名称 {total} 处，请确认不是本人单位，也不能据此推断本人单位",
                     "涉及本人单位的删除或改为泛称", BASIS_BLIND)


def build_report(path: Path, blind: bool = False, names: Sequence[str] = (), max_chars: Optional[int] = None,
                 section_limits: Optional[Dict[str, int]] = None, title: Optional[str] = None,
                 count_mode: str = "word") -> Dict[str, Any]:
    text = read_text(path)
    raw_lines = text.splitlines()
    lines: List[Tuple[int, str]] = []
    truncated_at: Optional[int] = None
    for index, line in enumerate(raw_lines, start=1):
        if DRAFT_NOTE_RE.match(line):
            truncated_at = index
            break
        lines.append((index, line))
    if not any(line.strip() for _, line in lines):
        where = "【起草说明】之前没有内容" if truncated_at is not None else "正文为空或只有空白"
        raise InputError(f"文件没有可检查的内容（{where}）。内容在文本框或图片中的，请另存为 .txt 后再查：{path}")
    section_limits = dict(section_limits or {})
    names = list(names)
    checker = Checker(blind, names, path.suffix.lower() == ".md")
    notes: List[str] = []
    if truncated_at is not None:
        notes.append(f"第 {truncated_at} 行为【起草说明】，其后内容未检查。")
    if blind and not names:
        notes.append("未提供 --names，只查通用的自称、批准号、邮箱和号码；申请人和成员姓名、单位请人工核对。")

    for no, line in lines:
        checker.check_line(no, line)
    if blind:
        checker.check_orgs(lines)

    # 字数
    body_text = "\n".join(line for _, line in lines)
    counts = count_text(body_text)
    if max_chars is not None and counts[count_mode] > max_chars:
        checker.add("PRC-LEN-001", "P0", "字数", None, f"全文约 {counts[count_mode]} 字",
                    f"超过字数上限 {max_chars}（{count_label(count_mode)}），多出约 {counts[count_mode] - max_chars} 字",
                    "压缩篇幅，优先删减重复的文献综述和背景", BASIS_LIMIT)

    # 栏目
    section_names = sorted(set(KNOWN_SECTIONS) | set(section_limits), key=len, reverse=True)
    sections = split_sections(lines, section_names, list(section_limits))
    section_rows: List[Dict[str, Any]] = []
    for sec in sections:
        n = count_text(sec["text"])[count_mode]
        limit = section_limits.get(sec["name"])
        status = "未设上限" if limit is None else ("超出" if n > limit else "未超出")
        section_rows.append({"name": sec["name"], "line": sec["line"], "count": n, "limit": limit, "status": status})
        if n == 0 and sec["name"] not in PARENT_SECTIONS:
            checker.add("PRC-FIL-003", "P1", "填写", sec["line"], sec["name"], f"栏目「{sec['name']}」为空",
                        "按模板填写；确实不需要的，按通知要求注明", BASIS_FILL)
        if limit is not None and n > limit:
            checker.add("PRC-LEN-002", "P0", "字数", sec["line"], f"{sec['name']}约 {n} 字",
                        f"栏目「{sec['name']}」超过 {limit} 字（{count_label(count_mode)}），多出约 {n - limit} 字",
                        "压缩该栏目", BASIS_LIMIT)
    found_names = {sec["name"] for sec in sections}
    for name in section_limits:
        if name not in found_names:
            checker.add("PRC-LEN-003", "P2", "字数", None, name, f"未找到栏目「{name}」，未检查该栏目字数",
                        "确认栏目名称与模板一致，或人工统计", BASIS_LIMIT)

    # 项目名称一致性
    if title:
        target = normalize_title(title)
        field_found = False
        for no, line in lines:
            m = TITLE_FIELD_RE.search(line)
            if not m:
                continue
            value = FIELD_CUT_RE.split(m.group(1), maxsplit=1)[0]
            value = normalize_title(value)
            if not value:
                continue
            field_found = True
            if value != target:
                checker.add("PRC-TTL-001", "P0", "一致性", no, checker.safe(line.strip())[:60],
                            "项目名称与申请书不一致", "逐字改为与申请书、系统填报一致的名称", BASIS_TITLE)
        if not field_found and target not in normalize_title(body_text):
            checker.add("PRC-TTL-002", "P2", "一致性", None, title,
                        "全文没有找到项目名称", "确认是否需要填写项目名称，并与申请书一致", BASIS_TITLE)

    issues = sorted(checker.issues, key=lambda i: (LEVELS.index(i["level"]), i["line"] or 0))
    summary = {lv: sum(1 for i in issues if i["level"] == lv) for lv in LEVELS}
    return {
        "file": path.name,
        "mode": {"blind": blind, "count": count_mode, "names_given": len(names)},
        "scope": "【起草说明】之前" if truncated_at else "全文",
        "counts": counts,
        "max_chars": max_chars,
        "sections": section_rows,
        "summary": summary,
        "issues": issues,
        "notes": notes,
    }


def count_label(mode: str) -> str:
    return {"word": "按 Word 字数口径估算", "chars": "按不计空白的字符数", "hanzi": "按汉字数"}[mode]


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------

def render_text(report: Dict[str, Any]) -> str:
    counts = report["counts"]
    mode = report["mode"]
    lines = [
        f"申报书形式检查：{report['file']}（范围：{report['scope']}）",
        f"匿名检查：{'开启' if mode['blind'] else '未开启'}；字数口径：{count_label(mode['count'])}",
        f"字数：Word 字数约 {counts['word']}，字符数（不计空白）{counts['chars']}，汉字 {counts['hanzi']}"
        + (f"；上限 {report['max_chars']}" if report["max_chars"] else ""),
    ]
    if report["sections"]:
        lines.append("栏目：" + "；".join(
            f"{row['name']} {row['count']}" + (f"/{row['limit']}" if row["limit"] else "")
            for row in report["sections"]))
    s = report["summary"]
    lines.append(f"问题：P0 {s['P0']} 条，P1 {s['P1']} 条，P2 {s['P2']} 条")
    lines.append("")
    for index, issue in enumerate(report["issues"], start=1):
        where = f"第 {issue['line']} 行" if issue["line"] else "全文"
        lines.append(f"{index}. [{issue['level']}] {issue['category']}｜{where}｜{issue['excerpt']}")
        lines.append(f"   {issue['message']}；建议：{issue['suggestion']}")
        lines.append(f"   依据：{issue['basis']}")
    for note in report["notes"]:
        lines.append(f"说明：{note}")
    if not report["issues"]:
        lines.append("未发现形式问题。资格、限项、签章和附件请按当年通知人工核对。")
    else:
        lines.append("资格、限项、签章和附件不在脚本检查范围内，请按当年通知人工核对。")
    return "\n".join(lines)


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
        prog="proposal_check.py",
        description="项目申报书形式审查辅助检查：匿名、字数、填写、项目名称一致性、残留。"
                    "只查形式，不评价内容，不改写文件。",
        epilog="退出码：0 正常；1 使用 --strict 且存在 P0；2 输入错误。",
        formatter_class=_ChineseHelpFormatter,
        add_help=False,
    )
    parser._optionals.title = "选项"
    parser._positionals.title = "参数"
    parser.add_argument("-h", "--help", action="help", help="显示本帮助信息并退出")
    parser.add_argument("file", help="申报书或活页（.txt、.md、.docx）")
    parser.add_argument("--blind", action="store_true", help="按匿名评审材料检查（活页等）")
    parser.add_argument("--names", help="申请人、成员姓名和所在单位，用逗号或顿号分隔；只在本机使用")
    parser.add_argument("--names-file", help="身份词清单文件，每行一个或用逗号分隔")
    parser.add_argument("--max-chars", type=int, help="全文字数上限，如 7000")
    parser.add_argument("--section", action="append", metavar="栏目=字数",
                        help="栏目字数上限，可重复，如 --section 选题说明=300")
    parser.add_argument("--title", help="申请书中的项目名称，用于核对一致性")
    parser.add_argument("--count", choices=("word", "chars", "hanzi"), default="word",
                        help="字数口径：word（默认，近似 Word 字数）、chars（不计空白的字符数）、hanzi（汉字数）")
    parser.add_argument("--json", action="store_true", help="以 JSON 格式输出")
    parser.add_argument("--strict", action="store_true", help="存在 P0 时返回 1")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    _setup_console()
    parser = make_parser()
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:
        return 0 if exc.code == 0 else 2
    try:
        if args.max_chars is not None and args.max_chars <= 0:
            raise InputError("--max-chars 须为正整数")
        names = parse_names(args.names, args.names_file)
        limits = parse_sections(args.section)
        report = build_report(Path(args.file), blind=args.blind, names=names, max_chars=args.max_chars,
                              section_limits=limits, title=args.title, count_mode=args.count)
    except InputError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(render_text(report))
    if args.strict and report["summary"]["P0"]:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
