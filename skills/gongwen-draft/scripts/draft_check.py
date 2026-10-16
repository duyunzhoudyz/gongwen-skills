#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""法定公文专项检查（gongwen-draft）。

补充 gongwen_lint.py 没有覆盖的文种专项：标题叠用与文种一致、会议通知要素、
转发和印发通知的转发语、请示的一文多事和本单位意见、批复表态、函的语气、
通报的决定、签报主送、通告和公告的使用条件，以及面向群众的停水停电、
取消低保等强制表述。与 gongwen_lint.py 已有的
请示结语、报告夹带请示、多头主送、签发人等检查不重复。

用法：
  python draft_check.py <文件> [--doc-type 文种] [--json] [--strict]
  python draft_check.py --fold-title "被转发文件标题" [--issuer "本机关名称"] [--verb 转发|批转] [--json]

转发标题两种写法（计划 D19）：写本机关名称的写“本机关转发……”，不加“关于”；
不写本机关名称的写“关于转发……的通知”，省去被转发标题中的“关于”。

输出 issue 的字段与机检接口（开发规范第 8 节）一致：id、level、category、line、
excerpt、message、suggestion、basis，另带 check="draft"。
退出码：默认 0；--strict 时存在 P0 返回 1；输入错误返回 2。只用标准库，读取 docx 时需要 python-docx。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, NoReturn, Optional, Sequence, Tuple

DRAFT_NOTE_LINE = "【起草说明】"
TIAOLI = "《党政机关公文处理工作条例》"

# 标题末尾可识别的文种，长的在前
TITLE_KINDS = (
    "复函", "请示", "报告", "批复", "通知", "通报", "意见", "决定", "通告", "公告",
    "纪要", "议案", "决议", "公报", "命令", "签报", "函", "令",
)
# 被转发文件标题中可能出现的文种或材料名称，长的在前
FORWARD_KINDS = (
    "实施意见", "指导意见", "若干意见", "实施方案", "工作方案", "行动方案", "管理办法",
    "实施办法", "实施细则", "工作要点", "意见", "通知", "决定", "方案", "办法", "细则",
    "规定", "规划", "计划", "要点", "通报", "报告", "批复", "公告", "通告", "纪要", "条例", "函",
)
KNOWN_TYPES = {
    "通知", "请示", "报告", "函", "通报", "批复", "意见", "签报", "决定", "通告", "公告",
    "决议", "命令", "公报", "议案", "纪要",
}

MARKER_RE = re.compile(r"【(待补|待核实)：[^】]*】")
PURE_DATE_RE = re.compile(r"^\d{4}年\d{1,2}月\d{1,2}日$")
DOCNO_RE = re.compile(r"〔\d{4}〕\d+号|\[\d{4}\]\d+号|（\d{4}）\d+号")
HEADER_RE = re.compile(r"^(签发人|份号|密级|紧急程度|特急|加急|秘密|机密|绝密)")
# 签报单、呈批件抬头印的表头字样，不是标题
FORM_LABELS = {"签报", "呈批件", "请示单", "签报单", "校内请示报告单"}
OPENING_RE = re.compile(r"如下|^(为|根据|按照|依据|依照|经|现|关于|按|对|因|鉴于)")
# 法规、规章类名称，单独作为被转发对象时加书名号；按“结尾是”判断，“管理办法”“实施细则”同样适用
STATUTE_KINDS = ("条例", "办法", "规定", "细则")
# 政府及其办公厅（室）。它转发部门文件、批转下级文件时，转发这一层赋予文件效力，不能省略
GOV_ISSUER_RE = re.compile(r"(人民政府|政府|办公厅|办公室)$")
# 转发标题的两种写法（计划 D19），DRF-001 的建议和 --fold-title 的提示共用
FORWARD_TITLE_FORMS = ("写本机关名称的写作“本机关转发原发文机关关于……的通知”，不加“关于”；"
                       "不写本机关名称的写作“关于转发原发文机关……的通知”，省去被转发标题中的“关于”")
# 本机关名称之后接“关于转发”：名称部分不含转发、批转动作，如“某区教育局关于转发……”。
# “某某关于批转……的通知”是批转的惯用写法，不在此列
ISSUER_GUANYU_FORWARD_RE = re.compile(r"^(?:(?!转发|批转).)+?关于转发")
# 会议时间单独成行的小标题，如“一、会议时间”“（一）时间：”
TIME_HEADING_RE = re.compile(
    r"^(?:[一二三四五六七八九十]+、|（[一二三四五六七八九十]+）|\d+[.．、])?\s*(?:会议|开会|报到)?时间[：:]?$"
)
# 带钟点的日期，如“11月5日（星期四）14时30分”“11月5日下午2时30分”
CLOCK_DATE_RE = re.compile(
    r"\d{1,2}月\d{1,2}日(?:（星期.）|\(星期.\))?\s*(?:上午|下午|晚上)?\s*"
    r"\d{1,2}(?:时(?:\d{1,2}分)?|点(?:\d{1,2}分|半)?|[:：]\d{2})"
)
# 钟点，如“14时30分”“下午2时30分”“9点半”“14:30”
CLOCK_RE = re.compile(r"(?<!\d)\d{1,2}(?:时(?:\d{1,2}分)?|点(?:\d{1,2}分|半)?(?![要意建内事措])|[:：]\d{2})")
# 截止时间、回执时间所在的行，不算会议时间
DEADLINE_RE = re.compile(r"回执|报名|截止|报送|反馈|前将|前报|前发|之前")


@dataclass
class Line:
    no: int
    text: str


@dataclass
class Doc:
    lines: List[Line]
    title: Optional[Line] = None
    zhusong: Optional[Line] = None
    body: List[Line] = field(default_factory=list)

    @property
    def body_text(self) -> str:
        return "\n".join(line.text for line in self.body)

    @property
    def full_text(self) -> str:
        return "\n".join(line.text for line in self.lines)


@dataclass
class Issue:
    id: str
    level: str
    category: str
    line: int
    excerpt: str
    message: str
    suggestion: str
    basis: str

    def to_dict(self) -> Dict[str, object]:
        return {
            "id": self.id, "level": self.level, "category": self.category, "line": self.line,
            "excerpt": self.excerpt, "message": self.message, "suggestion": self.suggestion,
            "basis": self.basis, "check": "draft",
        }


class InputError(Exception):
    """输入文件无法读取。"""


# ---------------------------------------------------------------------------
# 读取与切分
# ---------------------------------------------------------------------------

def setup_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
        except (AttributeError, ValueError):
            pass


def read_lines(path: Path) -> List[Line]:
    """读取 .txt/.md/.docx，遇到【起草说明】行停止，返回非空行。"""
    if not path.is_file():
        raise InputError(f"未找到文件：{path}")
    suffix = path.suffix.lower()
    raw: List[str]
    if suffix == ".docx":
        try:
            import docx  # type: ignore
        except ImportError:
            raise InputError("读取 docx 需要 python-docx，请先运行：pip install python-docx")
        try:
            raw = [p.text for p in docx.Document(str(path)).paragraphs]
        except Exception as exc:  # python-docx 抛出的异常类型不固定
            raise InputError(f"无法读取 docx：{exc}")
    elif suffix in (".txt", ".md", ""):
        data = path.read_bytes()
        text = None
        for enc in ("utf-8-sig", "gb18030"):
            try:
                text = data.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise InputError("文件编码无法识别，请另存为 UTF-8")
        raw = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    else:
        raise InputError(f"不支持的文件类型：{suffix}（只支持 .txt、.md、.docx）")
    lines: List[Line] = []
    for index, text in enumerate(raw, start=1):
        stripped = text.strip()
        if stripped == DRAFT_NOTE_LINE:
            break
        if stripped:
            lines.append(Line(index, stripped))
    return lines


def split_doc(lines: List[Line]) -> Doc:
    """识别标题、主送和正文。标题前的发文字号、签发人等版头行跳过。"""
    doc = Doc(lines=lines)
    rest = list(lines)
    while rest and (HEADER_RE.match(rest[0].text) or rest[0].text in FORM_LABELS
                    or (DOCNO_RE.search(rest[0].text) and len(rest[0].text) <= 30)):
        rest.pop(0)
    if rest:
        first = rest[0].text
        if not first.endswith(("：", ":", "。")) and len(first) <= 120:
            doc.title = rest.pop(0)
            # 标题折成两三行：下一行以“关于”开头，或本行没有文种、下一行以文种结尾，合并为一行
            while rest and len(doc.title.text) <= 120:
                nxt = rest[0].text
                if nxt.endswith(("：", ":", "。", "；", ";")) or len(nxt) > 60:
                    break
                if nxt.startswith("关于") or (not title_kind(doc.title.text) and title_kind(nxt)):
                    doc.title = Line(doc.title.no, doc.title.text + nxt)
                    rest.pop(0)
                else:
                    break
    if rest:
        cand = rest[0].text
        # 主送是机关名称加冒号；“为……，现就有关事项通告如下：”这类开头语不是主送。
        # 签报单常写成“主送：院长、分管副院长”，同样算主送
        if cand.startswith(("主送：", "主送:")) and "。" not in cand and len(cand) <= 120:
            doc.zhusong = rest.pop(0)
        elif (cand.endswith(("：", ":")) and "。" not in cand and len(cand) <= 120
                and not OPENING_RE.search(cand)):
            doc.zhusong = rest.pop(0)
    doc.body = rest
    return doc


def infer_type(title: Optional[str]) -> str:
    if not title:
        return ""
    clean = title.rstrip("。.；;")
    for kind in TITLE_KINDS:
        if clean.endswith(kind):
            if kind == "复函":
                return "函"
            if kind == "令":
                return "命令"
            return kind
    return ""


def title_kind(title: str) -> str:
    clean = title.rstrip("。.；;")
    for kind in TITLE_KINDS:
        if clean.endswith(kind):
            return kind
    return ""


def title_verb(title: str) -> str:
    """判断标题是不是转发、批转、印发通知，返回动词；不是的返回空字符串。

    动词紧跟“关于”（“关于印发……”），或出现在第一个“关于”之前（“某某转发某某关于……”）
    才算。“关于做好文件印发管理工作的通知”中的“印发”在事由里，不算。
    """
    m = re.search(r"关于(转发|批转|印发)", title)
    if m:
        return m.group(1)
    head = title.split("关于", 1)[0]
    m = re.search(r"转发|批转|印发", head)
    return m.group(0) if m else ""


# ---------------------------------------------------------------------------
# 检查
# ---------------------------------------------------------------------------

class Checker:
    def __init__(self, doc: Doc, doc_type: str) -> None:
        self.doc = doc
        self.doc_type = doc_type
        self.issues: List[Issue] = []

    def add(self, iid: str, level: str, category: str, line: Optional[Line], excerpt: str,
            message: str, suggestion: str, basis: str) -> None:
        self.issues.append(Issue(iid, level, category, line.no if line else 0,
                                 excerpt[:60], message, suggestion, basis))

    def find_line(self, pattern: "re.Pattern[str]") -> Optional[Tuple[Line, str]]:
        for line in self.doc.lines:
            m = pattern.search(line.text)
            if m:
                return line, m.group(0)
        return None

    def has(self, pattern: str, text: Optional[str] = None) -> bool:
        return re.search(pattern, self.doc.full_text if text is None else text) is not None

    # ---- 标题 ----
    def check_title(self) -> None:
        t = self.doc.title
        if t is None:
            self.add("DRF-004", "P1", "标题", None, "", "未识别出标题行",
                     "第一行写公文标题：发文机关＋事由＋文种", f"{TIAOLI}第九条第（七）项")
            return
        text = t.text
        if "关于关于" in text or text.count("关于") >= 3:
            self.add("DRF-001", "P1", "标题", t, text, "标题中“关于”重复",
                     "层层转发时省略中间层次，全标题只留一个“关于”：" + FORWARD_TITLE_FORMS
                     + "（可用 --fold-title 生成）",
                     "references/doc-types/tongzhi.md 第四部分（实务）")
        elif text.count("关于") == 2 and re.search(r"转发|批转", text):
            self.add("DRF-001", "P2", "标题", t, text, "转发标题出现两个“关于”",
                     FORWARD_TITLE_FORMS + "；本单位制度另有明文规定的从其规定",
                     "references/doc-types/tongzhi.md 第四部分（实务）")
        elif text.count("关于") == 1 and ISSUER_GUANYU_FORWARD_RE.match(text):
            self.add("DRF-001", "P2", "标题", t, text, "本机关名称之后又写“关于转发”",
                     FORWARD_TITLE_FORMS + "；本单位制度另有明文规定的从其规定",
                     "references/doc-types/tongzhi.md 第四部分（实务）")
        if "的通知的通知" in text:
            self.add("DRF-002", "P1", "标题", t, text, "标题出现“的通知的通知”",
                     "被转发的是通知时，两个“的通知”合为一个",
                     "references/doc-types/tongzhi.md 第四部分（实务）")
        elif "通知的通知" in text:
            self.add("DRF-002", "P2", "标题", t, text, "标题出现“通知的通知”",
                     "本 skill 默认合为一个“的通知”；本单位惯例保留的可不改",
                     "references/doc-types/tongzhi.md 第四部分（实务）")
        if len(re.findall(r"转发|批转", text)) >= 2 and has_redundant_levels(text):
            self.add("DRF-003", "P1", "标题", t, text, "标题中层层转发",
                     "直接转发最初的文件，省略只加了转发语的中间层次；批转、政府转发部门文件的那一层保留",
                     "references/_shared/rules/doc-types.md 2.8（实务）；references/doc-types/tongzhi.md 第四部分")
        kind = title_kind(text)
        if not kind and self.doc_type in KNOWN_TYPES:
            self.add("DRF-004", "P1", "标题", t, text, "标题末尾缺少文种",
                     f"标题写作“……的{self.doc_type if self.doc_type != '签报' else '请示'}”",
                     f"{TIAOLI}第九条第（七）项")
        if kind and self.doc_type in KNOWN_TYPES:
            inferred = infer_type(text)
            ok = inferred == self.doc_type or (self.doc_type == "签报" and kind in ("请示", "报告"))
            if not ok:
                self.add("DRF-005", "P1", "标题", t, text,
                         f"标题文种“{kind}”与指定的文书类型“{self.doc_type}”不一致",
                         "核对文种：先按写给谁、要办什么事定文种，再统一标题和结语",
                         f"{TIAOLI}第八条")
        if self.doc_type in ("请示", "签报") and re.search(r"关于申请", text):
            self.add("DRF-006", "P2", "标题", t, text, "请示标题中“申请”与“请示”意思重复",
                     "写作“关于拨付（增设、购置）……的请示”",
                     "references/doc-types/qingshi.md 第三部分（实务）")

    # ---- 请示、签报 ----
    REQUEST_RE = re.compile(
        r"(恳请|请求|拟请|请予|申请)[^。；\n]{0,40}?"
        r"(批准|审批|拨付|同意|支持|核准|解决|安排|增加|追加|调剂|调整|立项|购置|增设|设立|配备|核定)"
    )
    CLOSING_RE = re.compile(r"^(妥否|当否|以上请示|特此)[^。]{0,20}[。]?$")
    SELF_DECISION_RE = re.compile(r"(我|本)(局|院|校|委|办|镇|街道|单位|公司|中心|科|处|室|部|社区)[^。，；]{0,6}决定")
    # “研究决定，拟……”“决定向县政府请示”是内部决策程序的正当表述，不是先斩后奏
    DECISION_OK_RE = re.compile(r"^[，,]?\s*(拟|向|报请|申请|请示|恳请)|^[^。]{0,6}(申请|请示|报请)")
    # “根据我局党组决定精神”是引用已有决定作依据，不是本单位先作决定
    DECISION_CITED_RE = re.compile(r"(根据|按照|依据|依照|遵照)[^。，；]{0,4}$")

    # 两个以上方案的迹象：“方案一”“两种方案”“几种办法”“备选方案”
    MULTI_OPTION_RE = re.compile(
        r"方案[一二三四五六1-9ＡＢA-D]|[两二三四五几多]种(方案|办法|做法|方式|选择|路径)"
        r"|[两二三四五几多]个(方案|选项)|备选方案|各有利弊"
    )
    # 倾向性意见的写法
    TENDENCY_RE = re.compile(
        r"倾向|推荐|建议(采用|选择|按)|拟(采用|选择|按|选)|(本|我)(单位|部门|科室|处|局|院|校|镇|公司|中心)意见"
        r"|【待补：[^】]*(意见|倾向)"
    )

    def find_self_decision(self) -> Optional[Tuple[Line, str]]:
        for line in self.doc.body:
            for m in self.SELF_DECISION_RE.finditer(line.text):
                if self.DECISION_CITED_RE.search(line.text[:m.start()]):
                    continue
                if not self.DECISION_OK_RE.search(line.text[m.end():]):
                    return line, m.group(0)
        return None

    def check_request(self, need_contact: bool = True) -> None:
        hits: List[Line] = []
        for line in self.doc.body:
            if self.CLOSING_RE.match(line.text):
                continue
            if self.REQUEST_RE.search(line.text):
                hits.append(line)
        if len(hits) >= 2:
            where = "、".join(f"第{h.no}行" for h in hits[:4])
            self.add("DRF-010", "P2", "结构", hits[1], hits[1].text,
                     f"{where}各有一处请求，疑似一文多事",
                     "同一事项的配套内容算一事；互不相干的事项拆成几份分别请示",
                     f"{TIAOLI}第十五条第（四）项")
        m = self.find_self_decision()
        if m:
            self.add("DRF-011", "P1", "措辞", m[0], m[1], "请示中写了本单位“决定”，未经批准的事项不能先决定",
                     "本单位的打算写作“拟……”；已经办了的，如实写明并请求追认，或改写报告",
                     "references/doc-types/qingshi.md 第三部分、第六部分")
        # 只有一个方案时，请求事项本身就是意见（qingshi.md 第三部分第5条）；两个以上方案才要写倾向
        body = self.doc.body_text
        if self.MULTI_OPTION_RE.search(body) and not self.TENDENCY_RE.search(body):
            m = self.find_line(self.MULTI_OPTION_RE)
            self.add("DRF-012", "P2", "结构", m[0] if m else None, m[1] if m else "",
                     "列了两个以上方案，未见本单位倾向性意见",
                     "比较后写明倾向哪个方案和理由；用户没给的标【待补：本单位倾向性意见】",
                     "references/_shared/style/defensive.md 第七部分；references/doc-types/qingshi.md 第三部分")
        if need_contact:
            self.check_contact("P2")

    def check_contact(self, level: str) -> None:
        if not self.has(r"联系人|联系电话|【待补：[^】]*联系"):
            self.add("DRF-013", level, "格式要素", None, "", "未见联系人和联系电话",
                     "在成文日期后附注“（联系人：……，电话：……）”，缺的标【待补：联系人及电话】",
                     "references/_shared/rules/doc-types.md 2.11（实务）")

    def check_qianbao(self) -> None:
        z = self.doc.zhusong
        if z is not None:
            leaders = re.findall(r"院长|校长|书记|总经理|董事长|局长|主任|镇长|处长|部长|厅长|主席|领导", z.text)
            if re.search(r"[、，,和及]", z.text) or len(leaders) >= 2:
                self.add("DRF-080", "P1", "行文规则", z, z.text, "签报主送了多位领导",
                         "只报一位领导，由办公室按程序流转；需会签的部门列入会签",
                         "《交通运输部公文处理办法》第二十六条；references/doc-types/qianbao.md 第三部分")
        if not self.has(r"会签"):
            self.add("DRF-081", "P2", "程序", None, "", "未见会签情况",
                     "涉及其他部门职责的先会签；未会签的写明需会签的部门",
                     f"{TIAOLI}第十九条第（六）项")

    # ---- 报告 ----
    def check_report(self) -> None:
        m = self.find_line(re.compile(r"(可能|初步判断|估计|疑似|大概)[^。；]{0,20}(原因|导致|造成|引起)"))
        if m and m[0] in self.doc.body:
            self.add("DRF-020", "P2", "事实", m[0], m[1], "原因表述含推测",
                     "只写已查明的原因；未查明的写“原因正在调查”；用户材料中明确写了初步判断的照写",
                     "references/doc-types/baogao.md 第三部分")

    # ---- 函 ----
    def check_letter(self) -> None:
        m = self.find_line(re.compile(r"务必|必须|责令|限你|要求你(单位|局|委|院|校|公司|中心|办)"))
        if m and m[0] in self.doc.body:
            self.add("DRF-030", "P2", "措辞", m[0], m[1], "函为平行文，出现指令性措辞",
                     "改为商洽口吻，如“商请……”“请予支持”",
                     f"{TIAOLI}第八条第（十四）项")
        title = self.doc.title.text if self.doc.title else ""
        if title.endswith("复函") and not self.has(r"收悉"):
            self.add("DRF-031", "P1", "结构", self.doc.title, title, "复函没有引叙来函",
                     "开头写“你单位《关于……的函》（文号）收悉。”，文号缺的标【待补：来函文号】",
                     "references/_shared/rules/doc-types.md 2.14")
        m = self.find_line(re.compile(r"万望|不胜感激|恳请务必|万分感谢"))
        if m:
            self.add("DRF-032", "P2", "措辞", m[0], m[1], "函的语气过分谦卑",
                     "平行文语气平和，用“请予支持”“请予函复”即可",
                     "references/doc-types/han.md 第四部分")

    # ---- 批复 ----
    VAGUE_RE = re.compile(r"待条件成熟|再行研究|另行研究|视情况|原则上可行|酌情|再议|待研究后")
    # “可酌情调整”“可视就诊人数酌情安排”是授权下级自主把握，不算表态含糊
    GRANT_RE = re.compile(r"可(以)?[^，。；]{0,8}$")

    def find_vague_stance(self) -> Optional[Tuple[Line, str]]:
        for line in self.doc.body:
            for m in self.VAGUE_RE.finditer(line.text):
                if m.group(0) in ("酌情", "视情况") and self.GRANT_RE.search(line.text[:m.start()]):
                    continue
                return line, m.group(0)
        return None

    def check_reply(self) -> None:
        if not self.has(r"收悉"):
            self.add("DRF-040", "P1", "结构", self.doc.body[0] if self.doc.body else None, "",
                     "批复没有引叙来文",
                     "开头写“你单位《关于……的请示》（文号）收悉。经研究，现批复如下：”",
                     "references/_shared/rules/doc-types.md 2.12")
        stance = re.compile(r"不同意|原则同意|同意|不予|准予|批准|【待补：[^】]*批复意见")
        stance_lines = [ln for ln in self.doc.body if stance.search(ln.text) and "收悉" not in ln.text]
        if not stance_lines:
            self.add("DRF-041", "P1", "态度", None, "", "批复未见明确表态",
                     "写明同意、原则同意、部分同意或不同意；批复意见由用户决定，未定的标【待补：批复意见】",
                     f"references/_shared/style/defensive.md 第七部分；{TIAOLI}第八条第（十二）项")
        else:
            first = stance_lines[0]
            body_index = self.doc.body.index(first)
            if body_index >= 3:
                self.add("DRF-043", "P2", "结构", first, first.text, "表态位置靠后",
                         "表态放在引叙来文之后的第一句或第一条", "references/doc-types/pifu.md 第三部分")
        m = self.find_vague_stance()
        if m:
            self.add("DRF-042", "P1", "态度", m[0], m[1], "批复表态含糊",
                     "明确同意或不同意；原则同意的写明条件，不同意的写明理由",
                     "references/_shared/style/defensive.md 第七部分")

    # ---- 通知 ----
    MEETING_TITLE_RE = re.compile(r"(召开|举办|举行|参加)[^，。]{0,40}(会议|会|培训|座谈|论坛|研讨|交流|推进会|调度会|动员会)")

    def is_meeting(self) -> bool:
        title = self.doc.title.text if self.doc.title else ""
        if self.MEETING_TITLE_RE.search(title):
            return True
        return bool(re.search(r"定于[^。]{0,40}(召开|举办|举行)", self.doc.body_text))

    def check_notice(self) -> None:
        title = self.doc.title.text if self.doc.title else ""
        body = self.doc.body_text
        verb = title_verb(title)
        if verb in ("转发", "批转"):
            if not re.search(r"(转发|批转)给你们|现予(转发|批转)|现将[^。]{0,120}(转发|批转)|(转发|批转)如下", body):
                self.add("DRF-060", "P1", "结构", self.doc.body[0] if self.doc.body else None, "",
                         "转发通知没有转发语",
                         "写“现将《……》（文号）转发给你们，请结合实际认真贯彻执行。”",
                         "references/doc-types/tongzhi.md 第四部分")
            if len(re.sub(r"\s", "", body)) > 2000:
                self.add("DRF-061", "P2", "篇幅", None, "", "转发通知正文较长",
                         "检查是否复述了被转发文件；本级要求只写本地措施",
                         "references/_shared/rules/jianfu.md JF-04")
        elif verb == "印发":
            if not re.search(r"印发给你们|现予印发|现将[^。]{0,120}印发", body):
                self.add("DRF-062", "P1", "结构", self.doc.body[0] if self.doc.body else None, "",
                         "印发通知没有印发语",
                         "写“现将《……》印发给你们，请结合实际认真贯彻执行。”",
                         "references/doc-types/tongzhi.md 第五部分")
        if self.is_meeting():
            self.check_meeting()

    MEETING_TIME_RE = re.compile(
        r"(会议时间|时间[：:]|定于)[^。\n]{0,30}(\d{1,2}月\d{1,2}日|【待补：[^】]*(时间|日期)[^】]*】)"
        r"|\d{1,2}月\d{1,2}日[^。\n]{0,40}(召开|举行|举办|报到)"
        r"|【待补：[^】]*会议(时间|日期)"
    )

    def has_meeting_time(self) -> bool:
        """只认与会议相连的日期；成文日期、回执截止日期不算会议时间。"""
        lines = self.doc.lines
        merged: List[str] = []
        for index, line in enumerate(lines):
            text = line.text
            # “一、会议时间”单独成行时，与下一行拼起来再判断
            if TIME_HEADING_RE.match(text) and index + 1 < len(lines):
                text = text.rstrip("：:") + "：" + lines[index + 1].text
            merged.append(text)
        if self.MEETING_TIME_RE.search("\n".join(merged)):
            return True
        # 小标题不含“时间”二字（如“一、会议安排”）时，带钟点的日期也算会议时间
        for line in self.doc.body:
            if PURE_DATE_RE.match(line.text) or DEADLINE_RE.search(line.text):
                continue
            for m in CLOCK_DATE_RE.finditer(line.text):
                if not line.text[m.end():m.end() + 4].lstrip().startswith(("前", "以前", "之前")):
                    return True
        return False

    def has_meeting_clock(self) -> bool:
        """会议时间写到钟点，或钟点已标【待补】。回执截止时间里的钟点不算。"""
        if re.search(r"【待补：[^】]*(会议时间|会议日期|时刻|钟点|几时)", self.doc.full_text):
            return True
        for line in self.doc.body:
            if PURE_DATE_RE.match(line.text) or DEADLINE_RE.search(line.text):
                continue
            for m in CLOCK_RE.finditer(line.text):
                if not line.text[m.end():m.end() + 4].lstrip().startswith(("前", "以前", "之前")):
                    return True
        return False

    def check_meeting(self) -> None:
        text = self.doc.full_text
        anchor = self.doc.title
        if not self.has_meeting_time():
            m = self.find_line(re.compile(r"近期|近日|下周|择日|另行通知|月底|月初"))
            self.add("DRF-050", "P0", "会议要素", m[0] if m else anchor, m[1] if m else "",
                     "会议通知缺少具体会议时间",
                     "写明年月日和几时几分；用户没给的标【待补：会议时间】",
                     "references/doc-types/tongzhi.md 第三部分")
        elif not self.has_meeting_clock():
            m = self.find_line(re.compile(r"定于[^。]{0,30}|\d{1,2}月\d{1,2}日[^。]{0,12}"))
            self.add("DRF-054", "P2", "会议要素", m[0] if m else anchor, m[1] if m else "",
                     "会议时间只有日期，没有写到几时几分",
                     "写作“×月×日（星期×）×时×分”；用户没给钟点的标【待补：会议开始时刻】",
                     "references/doc-types/tongzhi.md 第三部分")
        if not re.search(r"地点|会议室|报告厅|礼堂|会场|大厅|楼|线上|视频|网络|【待补：[^】]*地点", text):
            self.add("DRF-051", "P0", "会议要素", anchor, "", "会议通知缺少会议地点",
                     "写到楼栋和房间；线上会议写平台；用户没给的标【待补：会议地点】",
                     "references/doc-types/tongzhi.md 第三部分")
        if not re.search(r"参会人员|参加人员|参会范围|参会对象|出席|列席|与会|请[^。]{0,30}参加|参加[^。]{0,4}(人员|对象)|【待补：[^】]*参会", text):
            self.add("DRF-052", "P1", "会议要素", anchor, "", "会议通知缺少参会人员",
                     "写明参会范围和职务；用户没给的标【待补：参会范围】",
                     "references/doc-types/tongzhi.md 第三部分")
        if not re.search(r"议程|主要内容|会议内容|内容为|议题|听取|部署|传达|学习|通报|总结|培训内容|【待补：[^】]*内容", text):
            self.add("DRF-055", "P2", "会议要素", anchor, "", "未见会议主要内容",
                     "写明会议主要议程；用户没给的标【待补：会议主要内容】，不编议程",
                     "references/doc-types/tongzhi.md 第三部分")
        if not re.search(r"回执|报名|名单|无需回执|不需[^。]{0,4}回执|【待补：[^】]*回执", text):
            self.add("DRF-053", "P2", "会议要素", anchor, "", "未见报名或回执方式",
                     "写明回执方式和截止时间；用户说不需要回执的可不写",
                     "references/doc-types/tongzhi.md 第三部分")
        self.check_contact("P2")

    # ---- 通报 ----
    def check_bulletin(self) -> None:
        title = self.doc.title.text if self.doc.title else ""
        text = self.doc.full_text
        if "表彰" in title or "表扬" in title:
            if not re.search(r"决定|予以(通报)?表彰|通报表扬|予以表扬|【待补：[^】]*决定", text):
                self.add("DRF-070", "P2", "结构", self.doc.title, title, "表彰通报未见表彰决定",
                         "写明“经……研究，决定对……予以通报表彰”，研究主体来自用户",
                         "references/doc-types/tongbao.md 第三部分")
        elif re.search(r"问题|批评|违规|违纪|事故|事件", title):
            if not re.search(r"处理|处分|决定|问责|【待补：[^】]*(处理|决定)", text):
                self.add("DRF-070", "P2", "结构", self.doc.title, title, "批评通报未见处理决定",
                         "处理决定只写用户提供的；没有的标【待补：处理决定】",
                         "references/doc-types/tongbao.md 第三部分")

    # ---- 通告、公告 ----
    NO_POWER_RE = re.compile(r"居民委员会|村民委员会|居委会|村委会|社区|学校|小学|中学|大学|学院|医院|卫生院|公司|集团")
    # 标题没有“关于”时，取到第一个机关名称后缀为止，避免把事由里的“学校”“医院”当成发文机关
    ORG_SUFFIX = (r"(?:居民委员会|村民委员会|居委会|村委会|人民政府|办事处|委员会|政府|办公厅|办公室|"
                  r"局|厅|委|社区|小学|中学|大学|学院|学校|卫生院|医院|公司|集团|中心)")
    ISSUER_HEAD_RE = re.compile(r"^.*?" + ORG_SUFFIX + ORG_SUFFIX + "*")

    @classmethod
    def issuer_of(cls, title: str) -> str:
        """取标题中“关于”或文种之前的发文机关。"""
        if "关于" in title:
            return title.split("关于", 1)[0]
        head = title.rstrip("。.；;")
        kind = title_kind(head)
        if kind:
            head = head[: -len(kind)]
        m = cls.ISSUER_HEAD_RE.match(head)
        return m.group(0) if m else head

    def check_public_notice(self) -> None:
        title = self.doc.title.text if self.doc.title else ""
        issuer = self.issuer_of(title)
        m = self.NO_POWER_RE.search(issuer)
        if m:
            self.add("DRF-094", "P1", "权限", self.doc.title, issuer,
                     f"发文主体“{issuer}”一般无权发布约束公众的通告",
                     "改用通知或温馨提示，转达有权机关的规定并写明规定名称",
                     "references/doc-types/tonggao.md 第一部分")
        if self.doc.zhusong is not None:
            self.add("DRF-090", "P2", "格式要素", self.doc.zhusong, self.doc.zhusong.text,
                     "通告一般不设主送机关", "删去主送；有明确受文机关的改用通知",
                     "references/_shared/rules/doc-types.md 7.3")
        text = self.doc.full_text
        # 只看正文，标题中的“春节期间”和成文日期不算适用时间
        body = "\n".join(ln.text for ln in self.doc.body if not PURE_DATE_RE.match(ln.text))
        if not re.search(r"施行|生效|自[^。]{0,20}起|至[^。]{0,15}止|\d{1,2}月\d{1,2}日[^。]{0,10}[至到－～—]|【待补：[^】]*(时间|日期)", body):
            self.add("DRF-091", "P1", "结构", self.doc.title, "", "通告缺少适用时间或施行日期",
                     "写明起止日期或“本通告自×年×月×日起施行”", "references/doc-types/tonggao.md 第三部分")
        if not re.search(r"根据|依据|依照|按照", text):
            self.add("DRF-092", "P2", "依据", self.doc.title, "", "通告未写依据",
                     "写明依据的法律、法规、规章名称和条款；用户没给的标【待核实】",
                     "references/doc-types/tonggao.md 第一部分")
        m = self.find_line(re.compile(r"解释权归[^。]{0,20}"))
        if m:
            self.add("DRF-093", "P2", "措辞", m[0], m[1], "通告中写了“解释权归……所有”",
                     "删去", "references/doc-types/tonggao.md 第五部分")

    # 与通告、通知事项无关的权益限制，出现即报
    COERCE_ALWAYS_RE = re.compile(
        r"取消[^。，；]{0,6}低保|停发[^。，；]{0,6}(补贴|补助|救助|低保)|影响[^。，；]{0,6}入学"
        r"|不予办理[^。，；]{0,4}(户口|入学)|后果自负"
    )
    # 停水停电停气：只在带有迫使履行的语境时报，供水供电企业的检修停水告知不报
    UTILITY_CUT_RE = re.compile(r"停止?[^。，；]{0,4}(供水|供电|供气|供热|供暖)|停(水|电|气)|断(水|电|气)")
    COERCE_CONTEXT_RE = re.compile(r"违反|违者|违规|拒不|不配合|拒绝|屡教不改|对其|处罚|责令|一律|强制")

    def check_coercion(self) -> None:
        for line in self.doc.body:
            for sentence in re.split(r"[。；！？]", line.text):
                m = self.COERCE_ALWAYS_RE.search(sentence)
                if not m:
                    cut = self.UTILITY_CUT_RE.search(sentence)
                    if cut and self.COERCE_CONTEXT_RE.search(sentence):
                        m = cut
                if m:
                    self.add("DRF-095", "P1", "法律表述", line, sentence,
                             f"面向群众的文书出现无依据的强制或权益限制：“{m.group(0)}”",
                             "删除；需要处理的写“由……（有权机关）依法处理”，依据来自用户",
                             "《行政强制法》第四十三条第二款；国办发〔2018〕37号；"
                             "references/doc-types/tonggao.md 第四部分")
                    break  # 同一行只报一次

    def check_announcement(self) -> None:
        m = self.find_line(re.compile(r"停水|停电|停气|施工|维修|考试|招聘|放假|闭馆|搬迁|封闭|检修"))
        if m:
            self.add("DRF-100", "P2", "文种", m[0], m[1], "公告用于向国内外宣布重要事项或法定事项",
                     "局部事项改用通知或通告；招聘、招标等习惯用法按专门规定",
                     f"{TIAOLI}第八条第（五）项；references/doc-types/gonggao.md")

    def run(self) -> List[Issue]:
        self.check_title()
        dt = self.doc_type
        if dt == "请示":
            self.check_request()
        elif dt == "签报":
            self.check_qianbao()
            title = self.doc.title.text if self.doc.title else ""
            if not title.endswith("报告"):
                self.check_request(need_contact=False)
        elif dt == "报告":
            self.check_report()
        elif dt == "函":
            self.check_letter()
        elif dt == "批复":
            self.check_reply()
        elif dt == "通知":
            self.check_notice()
            self.check_coercion()
        elif dt == "通报":
            self.check_bulletin()
        elif dt == "通告":
            self.check_public_notice()
            self.check_coercion()
        elif dt == "公告":
            self.check_announcement()
            self.check_coercion()
        order = {"P0": 0, "P1": 1, "P2": 2}
        self.issues.sort(key=lambda i: (order.get(i.level, 3), i.line, i.id))
        return self.issues


def check_text_lines(lines: List[Line], doc_type: str = "") -> Tuple[str, List[Issue], List[str]]:
    """对已切好的行做检查，返回（实际文种、问题清单、说明）。"""
    doc = split_doc(lines)
    notes: List[str] = []
    dtype = doc_type.strip()
    if not dtype:
        if lines and lines[0].text in FORM_LABELS:
            dtype = "签报"  # 稿首有签报表头：内部签报，标题“关于……的请示”不据此推断为请示
        else:
            dtype = infer_type(doc.title.text if doc.title else None)
        if dtype:
            basis = "稿首表头" if dtype == "签报" and lines and lines[0].text in FORM_LABELS else "标题"
            notes.append(f"未指定 --doc-type，按{basis}推断为“{dtype}”")
        else:
            notes.append("未指定 --doc-type，标题也推断不出文种，只做标题检查")
    elif dtype not in KNOWN_TYPES:
        notes.append(f"“{dtype}”不在本脚本的法定文种范围内，只做标题检查；其余检查请用 gongwen_lint.py")
    issues = Checker(doc, dtype).run()
    return dtype, issues, notes


def build_report(path: Path, lines: List[Line], dtype: str, issues: List[Issue], notes: List[str]) -> Dict[str, object]:
    text = "\n".join(line.text for line in lines)
    markers = {"待补": 0, "待核实": 0}
    for m in MARKER_RE.finditer(text):
        markers[m.group(1)] += 1
    summary = {"P0": 0, "P1": 0, "P2": 0, "markers": markers}
    for issue in issues:
        summary[issue.level] = int(summary.get(issue.level, 0)) + 1  # type: ignore[arg-type]
    return {
        "file": path.name,
        "doc_type": dtype or "其他",
        "char_count": len(re.sub(r"\s", "", text)),
        "summary": summary,
        "notes": notes,
        "issues": [i.to_dict() for i in issues],
    }


def render_text(report: Dict[str, object]) -> str:
    out = [f"文件：{report['file']}　文书类型：{report['doc_type']}　字数：{report['char_count']}"]
    for note in report["notes"]:  # type: ignore[union-attr]
        out.append(f"说明：{note}")
    summary = report["summary"]  # type: ignore[assignment]
    out.append(f"P0 {summary['P0']} 条，P1 {summary['P1']} 条，P2 {summary['P2']} 条；"
               f"待补 {summary['markers']['待补']} 处，待核实 {summary['markers']['待核实']} 处")
    for item in report["issues"]:  # type: ignore[union-attr]
        where = f"第{item['line']}行" if item["line"] else "全文"
        excerpt = f"「{item['excerpt']}」" if item["excerpt"] else ""
        out.append(f"[{item['level']}] {item['id']} {where}{excerpt}：{item['message']}。"
                   f"建议：{item['suggestion']}（{item['basis']}）")
    if not report["issues"]:
        out.append("未发现本脚本检查范围内的问题。")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# 转发标题折叠
# ---------------------------------------------------------------------------

FWD_A = re.compile(r"^(?P<issuer>.*?)关于(?P<verb>转发|批转)(?P<inner>.+)的通知$")
FWD_B = re.compile(r"^(?P<issuer>.*?)(?P<verb>转发|批转)(?P<inner>.+)的通知$")


def _strip_title(text: str) -> str:
    text = re.sub(r"\s", "", text).strip("。.；;")
    if text.startswith("《") and text.endswith("》"):
        text = text[1:-1]
    return text


def _has_book_marks(text: str) -> bool:
    text = re.sub(r"\s", "", text).strip("。.；;")
    return text.startswith("《") and text.endswith("》")


def _kind_of(core: str) -> str:
    for kind in FORWARD_KINDS:
        if core.endswith(kind):
            return kind
    return ""


def _issuer_of(title: str) -> str:
    return title.split("关于", 1)[0] if "关于" in title else ""


@dataclass
class Level:
    """转发链中的一层：这一层的发文机关、动词和这一层文件的完整标题。"""
    issuer: str
    verb: str
    title: str


def parse_levels(title: str) -> Tuple[List[Level], str]:
    """拆出转发链，返回（由外向内的各层，最初文件标题）。"""
    levels: List[Level] = []
    current = _strip_title(title)
    for _ in range(10):
        candidates = [m for m in (FWD_A.match(current), FWD_B.match(current)) if m]
        # “关于做好文件转发工作的通知”这类标题中的“转发”不是转发动作：发文机关部分不能含“关于”
        candidates = [m for m in candidates if "关于" not in m.group("issuer")]
        if not candidates:
            break
        # 两种写法都能匹配时取发文机关最短的，即最外层的动词
        m = min(candidates, key=lambda c: len(c.group("issuer")))
        inner = _strip_title(m.group("inner"))
        levels.append(Level(m.group("issuer"), m.group("verb"), current))
        # 两个“的通知”合写时，被转发的通知少了“的通知”，补回。两种情形：看不出文种；
        # 被转发的本身是转发、印发通知（“某某关于印发……方案”“某某转发某某关于……意见”）
        if not inner.endswith("通知") and (not _kind_of(inner) or title_verb(inner)):
            inner += "的通知"
        current = inner
    return levels, current


def parse_forward_chain(title: str) -> Tuple[List[str], str]:
    """拆出转发链，返回（中间各层发文机关，最初文件标题）。"""
    levels, original = parse_levels(title)
    return [level.issuer for level in levels], original


def keep_index(levels: List[Level], original: str, start: int = 0) -> Optional[int]:
    """返回第一个不能省略的层次。

    批转，或政府、办公厅（室）转发部门文件的，文件经这一层才对下级产生效力，这一层保留。
    政府办公室转发上级政府办公厅文件这类只加转发语的层次，可以省略。
    """
    orig_issuer = _issuer_of(original)
    for i in range(start, len(levels)):
        level = levels[i]
        inner_issuer = levels[i + 1].issuer if i + 1 < len(levels) else orig_issuer
        if level.verb == "批转":
            return i
        if GOV_ISSUER_RE.search(level.issuer) and inner_issuer and not GOV_ISSUER_RE.search(inner_issuer):
            return i
    return None


def has_redundant_levels(title: str) -> bool:
    """本文标题中，除本文这一层外有没有可以省略的中间层次；拆不出转发链的按有处理。"""
    levels, original = parse_levels(title)
    if len(levels) < 2:
        return True
    k = keep_index(levels, original, start=1)
    return (k if k is not None else len(levels)) > 1


def _compose(target: str, issuer: str, verb: str, keep_marks: bool) -> Tuple[str, str, List[str]]:
    """按被转发文件的标题组标题，返回（推荐标题，备选标题，提示）。"""
    notes: List[str] = []
    has_guanyu = "关于" in target
    orig_issuer, core = target.split("关于", 1) if has_guanyu else ("", target)
    kind = _kind_of(core)
    if kind == "通知":
        body = core[:-2]
        body = body[:-1] if body.endswith("的") else body
    else:
        body = core
        if kind and body[: -len(kind)].endswith("的"):
            body = body[: -len(kind) - 1] + kind
        # 法规、规章和规范性文件名称加书名号；输入时原本带书名号的照留
        if not has_guanyu and (body.endswith(STATUTE_KINDS) or keep_marks) and not body.startswith("《"):
            body = f"《{body}》"
        if not kind:
            notes.append("未识别出被转发文件的文种，请人工核对标题")
    # 不写本机关名称的写法：关于转发……的通知，省去被转发标题中的“关于”（计划 D19）
    without_issuer = f"关于{verb}{orig_issuer}{body}的通知"
    if has_guanyu:
        if issuer:
            recommended = f"{issuer}{verb}{orig_issuer}关于{body}的通知"
            alternative = without_issuer if orig_issuer else ""
        else:
            recommended, alternative = without_issuer, ""
        if not orig_issuer:
            notes.append("被转发文件的标题没有写原发文机关，"
                         + ("推荐标题保留了“关于”；" if issuer else "")
                         + "请补上原发文机关后重新生成")
    else:
        recommended = f"{issuer}{verb}{body}的通知" if issuer else without_issuer
        alternative = without_issuer if issuer else ""
        notes.append("未识别出原发文机关（标题中没有“关于”），请人工核对")
    if not issuer:
        notes.append(f"未提供本机关名称，只给出不写本机关名称的写法；标题写本机关名称的，写作“本机关{verb}……”，"
                     "不加“关于”，可加 --issuer 生成")
    return recommended, alternative, notes


def fold_forward_title(title: str, issuer: str, verb: str = "转发") -> Dict[str, object]:
    levels, original = parse_levels(title)
    k = keep_index(levels, original)
    if k is None:
        target, skipped = original, [level.issuer for level in levels]
    else:
        target, skipped = levels[k].title, [level.issuer for level in levels[:k]]
    recommended, alternative, notes = _compose(target, issuer, verb, keep_marks=_has_book_marks(title) and not levels)
    if k is not None:
        kept = levels[k]
        reason = "批转" if kept.verb == "批转" else "政府、办公厅（室）转发部门文件"
        folded, _alt, _notes = _compose(original, issuer, verb, keep_marks=False)
        notes.append(f"保留了“{kept.issuer}{kept.verb}”这一层：属于{reason}，文件经这一层才对下级产生效力，"
                     f"正文引用这一层的标题和文号。本单位确认可以直接转发最初文件的，可改为：{folded}")
    if skipped:
        notes.append("已省略中间层次：" + "、".join(skipped) + "。中间机关提出了具体要求、需要一并执行的，"
                     "改为转发中间机关的文件，正文写明“连同……一并转发”")
    return {
        "input": _strip_title(title),
        "issuer": issuer,
        "verb": verb,
        "original_issuer": _issuer_of(original),
        "original_title": original,
        "forwarded_title": target,
        "skipped_levels": skipped,
        "recommended": recommended,
        "alternative": alternative,
        "notes": notes,
        "basis": "references/doc-types/tongzhi.md 第四部分（实务，以本单位规定为准）",
    }


def render_fold(result: Dict[str, object]) -> str:
    out = [f"推荐标题：{result['recommended']}"]
    if result["alternative"]:
        out.append(f"备选标题（标题不写本机关名称时）：{result['alternative']}")
    out.append(f"最初文件：{result['original_title']}")
    if result["forwarded_title"] != result["original_title"]:
        out.append(f"实际转发：{result['forwarded_title']}")
    for note in result["notes"]:  # type: ignore[union-attr]
        out.append(f"提示：{note}")
    out.append(f"依据：{result['basis']}")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

class _ChineseHelpFormatter(argparse.HelpFormatter):
    def add_usage(self, usage, actions, groups, prefix=None):  # type: ignore[override]
        return super().add_usage(usage, actions, groups, prefix="用法：")


class _ChineseParser(argparse.ArgumentParser):
    """参数错误时用中文提示。argparse 自带的错误说明是英文，常见的几种换成中文。"""

    _MESSAGES = (
        ("unrecognized arguments", "无法识别的参数"),
        ("invalid choice", "取值无效"),
        ("expected one argument", "缺少取值"),
        ("the following arguments are required", "缺少必需参数"),
    )

    def error(self, message: str) -> "NoReturn":  # type: ignore[override]
        for en, zh in self._MESSAGES:
            message = message.replace(en, zh)
        self.print_usage(sys.stderr)
        self.exit(2, f"错误：{message}\n")


def make_parser() -> argparse.ArgumentParser:
    parser = _ChineseParser(
        prog="draft_check.py",
        description="法定公文专项检查：标题叠用、会议通知要素、请示、批复、函、签报、通告、公告等。"
                    "与 gongwen_lint.py 配合使用。",
        epilog="退出码：0 正常；1 --strict 时存在 P0；2 输入错误。",
        add_help=False,
        formatter_class=_ChineseHelpFormatter,
    )
    parser._positionals.title = "位置参数"  # argparse 未提供公开接口修改分组标题
    parser._optionals.title = "选项"
    parser.add_argument("-h", "--help", action="help", help="显示帮助并退出")
    parser.add_argument("file", nargs="?", metavar="文件", help="要检查的文件（.txt、.md、.docx），读到【起草说明】为止")
    parser.add_argument("--doc-type", default="", metavar="文种", help="文书类型，如 通知、请示、函、批复、签报；不填时按标题推断")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    parser.add_argument("--strict", action="store_true", help="存在 P0 时返回 1")
    parser.add_argument("--fold-title", metavar="标题", help="生成转发通知标题：传入被转发文件的完整标题")
    parser.add_argument("--issuer", default="", metavar="机关名称",
                        help="与 --fold-title 配合：本机关名称；不传时只给“关于转发……的通知”写法")
    parser.add_argument("--verb", default="转发", choices=("转发", "批转"), help="与 --fold-title 配合：转发或批转")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    setup_console()
    parser = make_parser()
    args = parser.parse_args(argv)

    if args.fold_title is not None:
        if not args.fold_title.strip():
            print("错误：--fold-title 不能为空", file=sys.stderr)
            return 2
        result = fold_forward_title(args.fold_title, args.issuer.strip(), args.verb)
        print(json.dumps(result, ensure_ascii=False, indent=2) if args.json else render_fold(result))
        return 0

    if not args.file:
        parser.print_usage(sys.stderr)
        print("错误：请提供要检查的文件，或使用 --fold-title", file=sys.stderr)
        return 2
    path = Path(args.file)
    try:
        lines = read_lines(path)
    except InputError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    if not lines:
        print("错误：文件没有正文内容", file=sys.stderr)
        return 2
    dtype, issues, notes = check_text_lines(lines, args.doc_type)
    report = build_report(path, lines, dtype, issues, notes)
    print(json.dumps(report, ensure_ascii=False, indent=2) if args.json else render_text(report))
    if args.strict and report["summary"]["P0"]:  # type: ignore[index]
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
