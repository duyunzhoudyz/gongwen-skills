#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""面向群众文书体检（gongwen-grassroots）。

适用于居民（村民）通知、告知书、倡议书、防汛防火等应急通知、村（居）务公开和公示、群消息。
检查八件事：
1. 要素是否齐全：称呼、时间、联系方式、落款、日期（按文书类型取舍）；
2. 日期与星期是否一致，日期是否存在（阿拉伯数字、中文数字日期，以及“星期二（1月12日）”的写法）；
3. 对群众的文书是否通俗：长句、公文腔词语；
4. 倡议书有没有强制性措辞；
5. 村（居）务公开的公布期限（低保初审公示按 7 天另行核对；起止日期恰好卡在临界值时另行提示）、
   公示中有没有与事项无关的个人信息；
6. 转发预警时有没有写明发布单位和发布时间；
7. 有没有引用已废止的法规（如《森林防火条例》）；
8. 面向群众的文书中，村（社区）常见、gongwen_lint.py legal 类容易漏报的强制措施（NTC-018）：
   取消资格或分红、要求签责任书、“不得进山”、停发扣除福利。

罚款、拘留、断水断电等一般法律表述和个人信息（身份证号、手机号）由 gongwen_lint.py 的
legal、privacy 两类检查负责，本脚本不重复。

用法：python scripts/notice_check.py <文件> [--kind 类型] [--year 年份] [--json] [--strict]
不填 --kind 时从标题推断；文件名含“群消息”的按群消息处理。
退出码：0 正常；1 使用 --strict 且存在 P0；2 输入错误。
只用标准库；读取 .docx 时需要 python-docx。不联网，不写任何文件。
"""
from __future__ import annotations

import argparse
import datetime as _dt
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

VERSION = "1.2.0"
KINDS = ("通知", "告知书", "倡议书", "应急通知", "公示", "群消息")
LEVELS = ("P0", "P1", "P2")

# 工具约定：群众文书单句超过 60 字提示；群消息超过 300 字提示（约一屏）
LONG_SENTENCE = 60
MAX_LONG_REPORTS = 5
GROUP_MSG_LIMIT = 300
# 村（居）务公开的公布期限下限（《村民委员会组织法》第三十五条、《城市居民委员会组织法》第三十四条）
PUBLIC_DAYS_MIN = 15
# 低保初审公示期（《最低生活保障审核确认办法》第十九条）
DIBAO_DAYS = 7

DRAFT_NOTE_RE = re.compile(r"^\s*【起草说明】\s*$")
TODO_RE = re.compile(r"【待(?:补|核实)：[^】]*】")

EMERGENCY_TITLE_RE = re.compile(
    r"防汛|防台|台风|暴雨|降雨|强降水|洪水|汛期|汛情|山洪|大风|冰雹|雷电|森林防火|草原防火|防灭火|防火|火险|"
    r"寒潮|暴雪|冰冻|高温|地质灾害|应急|预警|避险|转移|紧急通知"
)
PUBLIC_TITLE_RE = re.compile(r"公示|公开|公布")
PUBLIC_AUDIENCE_RE = re.compile(
    r"居民|村民|群众|市民|业主|住户|商户|家长|乡亲|朋友|游客|租户|街坊|邻居|老年人|各位"
)

TIME_RE = re.compile(
    r"\d{1,2}月\d{1,2}日|[一二三四五六七八九十]{1,2}月[一二三四五六七八九十]{1,3}日|"
    r"\d{1,2}[:：]\d{2}|\d{1,2}时|\d{1,2}点|星期[一二三四五六日天]|周[一二三四五六日天]|"
    r"今天|明天|后天|今晚|今夜|即日起|本周|下周|【待(?:补|核实)：[^】]*(?:时间|日期|期限)[^】]*】"
)
DEADLINE_RE = re.compile(r"前|之前|以前|以内|日内|天内|期限|截止|届满")
PHONE_RE = re.compile(
    r"(?<![\d.])(?:0\d{2,3}-?)?\d{7,8}(?![\d.]|\s*(?:元|万|亿|人|户|个|名|次|件|套|台|辆|吨|斤|公斤|千克|亩|平方|立方|米|千瓦|度|%|％))|"
    r"(?<!\d)1[3-9]\d{9}(?!\d)|"
    r"(?<!\d)1[3-9]\d\*{4}\d{4}(?!\d)|(?<!\d)\d{3}-\d{3,4}-\d{4}(?!\d)|"
    r"(?:拨打|电话|热线|报警|求助)[^。\n]{0,4}(?<!\d)(?:110|119|120|122|12345|12315|12319|12350)(?!\d)|"
    r"【待(?:补|核实)：[^】]*(?:电话|联系)[^】]*】"
)
ORG_TAIL_RE = re.compile(
    r"(?:委员会|居委会|村委会|办事处|人民政府|党委|党工委|党支部|党总支|社区|街道|镇|乡|村|"
    r"指挥部|办公室|服务中心|服务站|中心|物业|物业服务中心|工作站|小组|站|所)$"
)
SIGN_TODO_RE = re.compile(r"^【待(?:补|核实)：[^】]*(?:落款|单位|署名)[^】]*】$")
CN_NUM = "〇零一二三四五六七八九"
# 成文日期行；紧急通知可写到“时”“分”，缺时的可写“2026年7月12日【待补：时】”
DATE_LINE_RE = re.compile(
    r"^(?:\d{4}年\d{1,2}月\d{1,2}日(?:\d{1,2}时(?:\d{1,2}分)?|【待(?:补|核实)：[^】]*】)?|"
    r"[〇零一二三四五六七八九]{4}年[一二三四五六七八九十]{1,2}月[一二三四五六七八九十]{1,3}日|"
    r"【待(?:补|核实)：[^】]*(?:日期|时间)[^】]*】)$"
)
YEAR_RE = re.compile(r"(\d{4})年")
_MD = r"\d{1,2}|[一二三四五六七八九十]{1,3}"
_YEAR = r"\d{4}|[〇零一二三四五六七八九]{4}"
# 日期加星期：“10月20日（星期二）”“十月二十日（星期二）”“星期二（10月20日）”
WEEKDAY_PATTERNS: Tuple["re.Pattern[str]", ...] = (
    re.compile(
        rf"(?:(?P<year>{_YEAR})年)?(?P<month>{_MD})月(?P<day>{_MD})日\s*[（(]?\s*"
        r"(?:星期|周|礼拜)(?P<wd>[一二三四五六日天])\s*[）)]?"
    ),
    re.compile(
        r"(?:星期|周|礼拜)(?P<wd>[一二三四五六日天])\s*[（(]\s*"
        rf"(?:(?P<year>{_YEAR})年)?(?P<month>{_MD})月(?P<day>{_MD})日\s*[）)]"
    ),
)
# 没有和日期写在一起的“星期X”（“周X”常见于“每周一”，不计）
LONE_WEEKDAY_RE = re.compile(r"(?:星期|礼拜)[一二三四五六日天]")
CN_YEAR_DIGITS = {"〇": 0, "零": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
WEEKDAY_NAMES = "一二三四五六日"
SENTENCE_SPLIT_RE = re.compile(r"(?<=[。！？；!?;])")

# 面向群众时的公文腔词语：(正则, 说明, 建议)
JARGON: Tuple[Tuple[str, str, str], ...] = (
    (r"兹定于", "“兹定于”是旧式公文用语", "改为“定于”或“将在”"),
    (r"届时", "“届时”对居民不够直白", "改为“到时”或写明具体时间"),
    (r"予以", "“予以”是书面公文用语", "删去，或改成具体动作，如“帮您办理”"),
    (r"务必", "“务必”语气生硬", "改为“请”"),
    (r"(?:相关|有关)事宜", "“相关事宜”没有说清是什么事", "写明具体事项"),
    (r"高度重视", "“高度重视”是机关套话，居民不需要", "删去，直接写要做什么"),
    (r"切实", "“切实”是机关套话", "删去"),
    (r"落实落细|走深走实|压实责任|抓手|闭环|赋能|见行见效", "机关内部用语，居民看不懂", "改成具体做法或删去"),
    (r"属地管理|履职|台账|网格化管理", "内部工作用语", "改成居民能懂的说法，如“由社区工作人员上门登记”"),
    (r"广大居民群众|广大村民群众|广大人民群众", "称呼生硬", "改为“各位居民”“各位村民”"),
    (r"配合做好", "“配合做好”把居民当成下级", "改为“请您……”，写清需要居民做的具体事"),
    (r"即日起", "“即日起”对居民不够直白", "写明从几月几日起"),
)
COERCIVE_RE = re.compile(
    r"必须|严禁|禁止|不得|一律|务必|责令|处罚|罚款|追究|问责|考核|通报批评|曝光|黑名单|后果自负"
)
PUBLICITY_PERIOD_RE = re.compile(r"公示期|公示时间|公示期限|公布期限|公开期限|公示至|公示日期|公开时间|公布时间|公示从")
# 村（居）务公开：标题含这些词，或正文出现“村务公开”等（“村务公开栏”“村务监督委员会”不算）
WUGONGKAI_TITLE_RE = re.compile(r"(?:村务|居务|财务)公开|(?:村|居)务(?!监督)")
WUGONGKAI_TEXT_RE = re.compile(r"(?:村务|居务|财务)公开(?!栏)")
# 村（居）务公开本身的公布期限行；公示子项（如低保初审公示）用“公示期”“公示时间”，不在此列
DOC_PERIOD_LABEL_RE = re.compile(r"(?:公开|公布)(?:时间|期限|日期)")
SUB_PERIOD_LABEL_RE = re.compile(r"^(?:公示期|公示时间|公示期限|公示日期|公示从|公示至)")
# 有自己公示期规定的事项：不按村（居）务公开的 15 日判断
EXEMPT_PERIOD_RE = re.compile(r"低保|最低生活保障|特困|救助|初审")
DIBAO_RE = re.compile(r"低保|最低生活保障")
HEADING_RE = re.compile(r"^[一二三四五六七八九十]+、")
PERIOD_DAYS_RE = re.compile(
    r"(?:公示期|公布期限|公开期限|公示时间|公示期限|公布时间|公开时间)[^。\n]{0,10}?"
    r"(?<![月\d一二三四五六七八九十])(\d+|[一二三四五六七八九十]+)\s*(?:个工作日|日|天)"
)
DATE_RANGE_RE = re.compile(r"(?:(\d{4})年)?(\d{1,2})月(\d{1,2})日[^。\n]{0,6}?(?:至|到|—|～|~|-)\s*(?:(\d{4})年)?(?:(\d{1,2})月)?(\d{1,2})日")
SENSITIVE_RE = re.compile(
    r"病情|病史|诊断|疾病名称|精神疾病|精神病|艾滋|癌症|透析|服刑|吸毒|戒毒|残疾等级|离婚原因|"
    r"家庭住址|门牌号|身份证号|银行卡号"
)
PREWARN_RE = re.compile(
    r"(?:蓝色|黄色|橙色|红色)(?:预警信号|预警)|[ⅠⅡⅢⅣ]级(?:应急响应|响应|预警)|"
    r"[一二三四1234]级(?:应急响应|响应|预警)"
)
PREWARN_SOURCE_RE = re.compile(
    r"气象台|气象局|气象部门|防汛抗旱指挥部|防汛指挥部|防指|应急管理局|应急管理部门|水文|自然资源|"
    r"指挥部|防灭火指挥部|人民政府|【待(?:补|核实)：[^】]*(?:发布|来源|单位)[^】]*】"
)
RELEASE_TIME_RE = re.compile(r"\d{1,2}时|\d{1,2}[:：]\d{2}|\d{1,2}点|【待(?:补|核实)：[^】]*时间[^】]*】")

# 已废止的法规：(书名号内名称, 替代说明)
REPEALED_LAWS: Tuple[Tuple[str, str], ...] = (
    ("森林防火条例", "《森林草原防灭火条例》（2025年11月公布，2026年1月1日起施行）"),
    ("草原防火条例", "《森林草原防灭火条例》（2025年11月公布，2026年1月1日起施行）"),
)
BASIS_REPEALED = "《森林草原防灭火条例》第六十三条（《森林防火条例》《草原防火条例》同时废止）"

BASIS_ELEMENTS = "references/templates/jumin-wenshu.md；references/_shared/profiles/units/shequ.md 第八、九部分"
BASIS_PLAIN = "references/_shared/profiles/units/shequ.md 第九部分（对居民说大白话）；60 字为工具约定"
BASIS_INITIATIVE = "references/templates/jumin-wenshu.md（倡议书邀请自愿参与，不带强制性）"
BASIS_GROUP = "references/_shared/profiles/units/shequ.md 第八部分第 3 条；300 字为工具约定"
BASIS_PERIOD = "《村民委员会组织法》第三十五条、《城市居民委员会组织法》第三十四条（2026年1月1日起施行）"
BASIS_DIBAO = "《最低生活保障审核确认办法》第十九条（初审公示期为 7 天）"
BASIS_SENSITIVE = "《最低生活保障审核确认办法》第二十四条；《个人信息保护法》第二十八条"
BASIS_PREWARN = "《气象灾害预警信号发布与传播办法》第七条、第十条；《突发事件应对法》第七条"
BASIS_COERCION = ("references/redlines.md 第一条；references/templates/jumin-wenshu.md 第七部分"
                  "（村（居）委会是基层群众性自治组织，没有行政执法权）")
BASIS_BOUNDARY = "《民法典》第二百零一条（按日计算期间的，开始的当日不计入）；各地公示期算法不一，以本地惯例为准"

# 面向群众的文书中，村（社区）常见、gongwen_lint.py legal 类容易漏报的强制措施：(正则, 说明)
COERCION_EXTRA: Tuple[Tuple["re.Pattern[str]", str], ...] = (
    (re.compile(r"取消[^，。；！？\n]{0,12}(?:资格|待遇|分红|福利|补贴|补助)"), "以取消资格、待遇、分红或补贴作为强制手段"),
    (re.compile(r"(?:签订|签署)[^，。；！？\n]{0,8}(?:责任书|承诺书|保证书)"), "要求群众签责任书、承诺书或保证书"),
    (re.compile(r"不得(?:进|入)(?:山|村|小区|楼)|禁止(?:进|入)(?:山|村|小区)"), "限制群众进山、进村、进小区"),
    (re.compile(r"(?:停发|扣除|扣发)[^，。；！？\n]{0,6}(?:分红|福利|补贴|补助)"), "以停发、扣除分红或福利作为强制手段"),
)
# 句中写明了依据（转述上级命令、法规）：只提示核对依据，不直接定为 P0
BASIS_CUE_RE = re.compile(r"根据|按照|依据|依照|《[^》]+》|人民政府[^。\n]{0,10}(?:令|命令|通告)")
# 落款是村（居）委会、社区或村党组织
GRASSROOTS_SIGN_RE = re.compile(
    r"(?:村|居)民委员会|村委会|居委会|村(?:党支部|党总支|党委)|社区(?:党支部|党总支|党委)|(?:社区|村)$"
)


class InputError(Exception):
    """输入文件错误（退出码 2）。"""


# ---------------------------------------------------------------------------
# 读取
# ---------------------------------------------------------------------------

def read_lines(path: Path) -> Tuple[List[Tuple[int, str]], List[str]]:
    """读取文件，返回 [(行号, 文本)] 和提示信息。"""
    notes: List[str] = []
    if not path.is_file():
        raise InputError(f"未找到文件：{path}")
    suffix = path.suffix.lower()
    if suffix in (".txt", ".md", ".markdown", ".text"):
        raw = path.read_bytes()
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            try:
                text = raw.decode("gb18030")
                notes.append("文件不是 UTF-8 编码，已按 GB18030 读取。")
            except UnicodeDecodeError as exc:
                raise InputError("无法识别文件编码，请另存为 UTF-8 后重试。") from exc
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        return [(i, line) for i, line in enumerate(text.split("\n"), start=1)], notes
    if suffix == ".docx":
        try:
            import docx  # type: ignore
        except ImportError as exc:
            raise InputError("读取 .docx 需要 python-docx，请先运行：pip install python-docx") from exc
        try:
            document = docx.Document(str(path))
        except Exception as exc:  # 损坏或加密的 docx
            raise InputError(f"无法打开 docx 文件：{exc}") from exc
        return [(i, p.text) for i, p in enumerate(document.paragraphs, start=1)], notes
    raise InputError(f"不支持的文件类型：{path.suffix or '无扩展名'}。只支持 .txt、.md、.docx。")


def body_lines(lines: Sequence[Tuple[int, str]]) -> Tuple[List[Tuple[int, str]], Optional[int]]:
    """取【起草说明】之前的非空行，返回 (行列表, 起草说明所在行号)。"""
    kept: List[Tuple[int, str]] = []
    for number, text in lines:
        if DRAFT_NOTE_RE.match(text):
            return kept, number
        stripped = text.strip()
        if stripped:
            kept.append((number, stripped))
    return kept, None


def nonspace_len(text: str) -> int:
    return len(re.sub(r"\s", "", text))


# ---------------------------------------------------------------------------
# 判断类型与对象
# ---------------------------------------------------------------------------

def detect_kind(title: str, explicit: Optional[str], filename: str = "") -> Tuple[str, str]:
    if explicit:
        return explicit, "参数"
    if "群消息" in filename:
        return "群消息", "文件名推断"
    if "倡议" in title:
        return "倡议书", "标题推断"
    if re.search(r"告知书|告知函|一封信|告居民书|告村民书", title):
        return "告知书", "标题推断"
    if PUBLIC_TITLE_RE.search(title):
        return "公示", "标题推断"
    if EMERGENCY_TITLE_RE.search(title):
        return "应急通知", "标题推断"
    return "通知", "默认"


def find_salutation(body: Sequence[Tuple[int, str]]) -> Optional[Tuple[int, str]]:
    """以冒号结尾的短行视为称呼或主送对象：没有标题时看第一行，有标题时看标题后 3 行。"""
    for number, text in body[0:4]:
        if text.endswith(("：", ":")) and nonspace_len(text) <= 30:
            return number, text
    return None


def is_public(kind: str, title: str, salutation: Optional[Tuple[int, str]]) -> bool:
    if kind in ("倡议书", "群消息", "公示", "告知书"):
        return True
    if salutation and PUBLIC_AUDIENCE_RE.search(salutation[1]):
        return True
    if salutation is None and PUBLIC_AUDIENCE_RE.search(title):
        return True
    return False


# ---------------------------------------------------------------------------
# 日期工具
# ---------------------------------------------------------------------------

def cn_to_int(text: str) -> Optional[int]:
    """把“十五”“二十”“三”之类的中文数字转为整数（1–99）。"""
    if text.isdigit():
        return int(text)
    digits = {c: i for i, c in enumerate("零一二三四五六七八九")}
    if text == "十":
        return 10
    if text.startswith("十"):
        return 10 + digits.get(text[1:], 0) if len(text) == 2 else None
    if "十" in text:
        tens, _, ones = text.partition("十")
        if tens not in digits:
            return None
        return digits[tens] * 10 + (digits.get(ones, 0) if ones else 0)
    return digits.get(text)


def signoff_year(body: Sequence[Tuple[int, str]]) -> Tuple[Optional[int], Optional[int]]:
    """从落款日期取年份和月份。"""
    for _number, text in reversed(body[-4:]):
        match = re.match(r"^(\d{4})年(\d{1,2})月\d{1,2}日(?:\d{1,2}时(?:\d{1,2}分)?|【待(?:补|核实)：[^】]*】)?$", text)
        if match:
            return int(match.group(1)), int(match.group(2))
    return None, None


def parse_year(text: Optional[str]) -> Optional[int]:
    """“2026”或“二〇二六”转为整数。"""
    if not text:
        return None
    if text.isdigit():
        return int(text)
    if all(c in CN_YEAR_DIGITS for c in text):
        return int("".join(str(CN_YEAR_DIGITS[c]) for c in text))
    return None


def resolve_year(explicit: Optional[int], month: int, arg_year: Optional[int],
                 sign_year: Optional[int], sign_month: Optional[int], first_year: Optional[int]) -> Optional[int]:
    if explicit:
        return explicit
    if arg_year:
        return arg_year
    if sign_year:
        # 12 月发文、写次年 1 月的事：月份比发文月份小 6 个月以上时按次年算
        if sign_month and sign_month - month > 6:
            return sign_year + 1
        return sign_year
    return first_year


# ---------------------------------------------------------------------------
# 检查
# ---------------------------------------------------------------------------

class Checker:
    def __init__(self, body: List[Tuple[int, str]], kind: str, year: Optional[int]) -> None:
        self.body = body
        self.kind = kind
        self.arg_year = year
        self.issues: List[Dict[str, Any]] = []
        self.notes: List[str] = []
        self.title = body[0][1] if body else ""
        self.salutation = find_salutation(body)
        self.public = is_public(kind, self.title, self.salutation)
        self.full_text = "\n".join(text for _n, text in body)
        # 落款日期行不算正文中的时间要素
        self.text_wo_date = "\n".join(text for _n, text in body if not DATE_LINE_RE.match(text))

    def add(self, issue_id: str, level: str, category: str, line: Optional[int], excerpt: str,
            message: str, suggestion: str, basis: str) -> None:
        self.issues.append({
            "id": issue_id,
            "level": level,
            "category": category,
            "line": line,
            "excerpt": excerpt[:60],
            "message": message,
            "suggestion": suggestion,
            "basis": basis,
        })

    # 一、要素 -------------------------------------------------------------
    def check_elements(self) -> None:
        kind = self.kind
        # 公示、村（居）务公开和群消息可以不写称呼
        if kind not in ("群消息", "公示") and self.salutation is None:
            self.add("NTC-001", "P1", "要素", self.body[0][0], self.title,
                     "缺少称呼或主送对象行（如“各位居民：”“各村（社区）：”）",
                     "标题下一行写明对象，以全角冒号结尾", BASIS_ELEMENTS)
        has_time = bool(TIME_RE.search(self.text_wo_date))
        if kind in ("通知", "应急通知", "群消息") and not has_time:
            self.add("NTC-002", "P1", "要素", None, self.title,
                     "正文没有具体时间（日期、几点、星期）",
                     "写明日期和时间；用户没提供的写【待补：时间】", BASIS_ELEMENTS)
        if kind == "告知书" and not (has_time or DEADLINE_RE.search(self.text_wo_date)):
            self.add("NTC-002", "P2", "要素", None, self.title,
                     "告知书没有写办理期限或时间",
                     "需要居民在期限内办理的，写明截止日期", BASIS_ELEMENTS)
        if kind != "倡议书" and not PHONE_RE.search(self.full_text):
            self.add("NTC-003", "P1", "要素", None, self.title,
                     "没有联系电话或咨询、监督电话",
                     "写明联系人和电话；用户没提供的写【待补：联系电话】", BASIS_ELEMENTS)
        if kind == "群消息":
            return
        tail = self.body[-4:] if len(self.body) > 1 else []
        has_date = any(DATE_LINE_RE.match(text) for _n, text in tail)
        has_sign = any(
            (ORG_TAIL_RE.search(text) and "。" not in text and nonspace_len(text) <= 40
             and not DATE_LINE_RE.match(text) and not text.endswith(("：", ":")))
            or SIGN_TODO_RE.match(text)
            for _n, text in tail[-3:]
        )
        last_line = self.body[-1][0] if self.body else None
        if not has_sign:
            self.add("NTC-004", "P1", "要素", last_line, self.body[-1][1] if self.body else "",
                     "末尾没有落款单位", "正文后另起一行写落款单位全称（张贴版须加盖公章）", BASIS_ELEMENTS)
        if not has_date:
            self.add("NTC-005", "P1", "要素", last_line, self.body[-1][1] if self.body else "",
                     "末尾没有成文日期", "落款下一行写成文日期，如“2026年10月16日”", BASIS_ELEMENTS)

    # 二、日期与星期 ---------------------------------------------------------
    def check_weekday(self) -> None:
        sign_year, sign_month = signoff_year(self.body)
        first = YEAR_RE.search(self.full_text)
        first_year = int(first.group(1)) if first else None
        skipped = 0
        unparsed = 0
        for number, text in self.body:
            matches = self._weekday_matches(text)
            for match in matches:
                month, day = cn_to_int(match.group("month")), cn_to_int(match.group("day"))
                if month is None or day is None or not 1 <= month <= 12:
                    unparsed += 1
                    continue
                year = resolve_year(parse_year(match.group("year")), month, self.arg_year,
                                    sign_year, sign_month, first_year)
                if year is None:
                    skipped += 1
                    continue
                try:
                    date = _dt.date(year, month, day)
                except ValueError:
                    self.add("NTC-006", "P0", "日期", number, match.group(0),
                             f"{year}年没有{month}月{day}日", "核实日期", "核对事实：日期必须真实存在")
                    continue
                actual = WEEKDAY_NAMES[date.weekday()]
                written = "日" if match.group("wd") == "天" else match.group("wd")
                if written != actual:
                    self.add("NTC-006", "P0", "日期", number, match.group(0),
                             f"日期与星期不一致：{year}年{month}月{day}日是星期{actual}，稿中写的是星期{match.group('wd')}",
                             "向用户核实日期和星期哪个正确，不要自行改动",
                             "核对事实：居民会按错误的时间行动")
            spans = [m.span() for m in matches]
            unparsed += sum(
                1 for m in LONE_WEEKDAY_RE.finditer(text)
                if not any(start <= m.start() < end for start, end in spans)
            )
        if skipped:
            self.notes.append(f"有 {skipped} 处“日期加星期”无法确定年份，未核对星期；可用 --year 指定年份。")
        if unparsed:
            self.notes.append(f"有 {unparsed} 处星期未核对（没有和日期写在一起，或日期写法无法识别）；"
                              "用户只给了星期几的，保留原话或标【待补：具体日期】，不要推算日期；"
                              "用户给了日期的，请人工核对星期。")

    @staticmethod
    def _weekday_matches(text: str) -> List["re.Match[str]"]:
        """找出“日期加星期”和“星期加（日期）”，重叠的只取先找到的。"""
        found: List["re.Match[str]"] = []
        for pattern in WEEKDAY_PATTERNS:
            for match in pattern.finditer(text):
                start, end = match.span()
                if any(start < m.end() and m.start() < end for m in found):
                    continue
                found.append(match)
        return found

    # 三、通俗程度 -----------------------------------------------------------
    def check_plain(self) -> None:
        if not self.public:
            return
        skip = {self.body[0][0]}
        if self.salutation:
            skip.add(self.salutation[0])
        long_hits: List[Tuple[int, str]] = []
        for number, text in self.body:
            if number in skip or DATE_LINE_RE.match(text):
                continue
            for sentence in SENTENCE_SPLIT_RE.split(text):
                clean = TODO_RE.sub("", sentence)
                if nonspace_len(clean) > LONG_SENTENCE:
                    long_hits.append((number, sentence))
        for number, sentence in long_hits[:MAX_LONG_REPORTS]:
            self.add("NTC-007", "P2", "通俗", number, sentence,
                     f"句子超过 {LONG_SENTENCE} 字，居民不容易一次看懂",
                     "拆成几句，一句话说一件事", BASIS_PLAIN)
        if len(long_hits) > MAX_LONG_REPORTS:
            self.notes.append(f"长句共 {len(long_hits)} 句，只列出前 {MAX_LONG_REPORTS} 句。")
        for pattern, message, suggestion in JARGON:
            regex = re.compile(pattern)
            hits = [(n, m.group(0)) for n, t in self.body for m in regex.finditer(t)]
            if not hits:
                continue
            number, word = hits[0]
            extra = f"（全文共 {len(hits)} 处）" if len(hits) > 1 else ""
            self.add("NTC-008", "P2", "通俗", number, word, message + extra, suggestion, BASIS_PLAIN)

    # 四、倡议书 -------------------------------------------------------------
    def check_initiative(self) -> None:
        if self.kind != "倡议书":
            return
        seen = set()
        for number, text in self.body:
            for match in COERCIVE_RE.finditer(text):
                word = match.group(0)
                if word in seen:
                    continue
                seen.add(word)
                self.add("NTC-009", "P1", "语气", number, word,
                         f"倡议书出现强制性措辞“{word}”",
                         "倡议只邀请自愿参与：改为“请大家”“欢迎”；确需强制执行的，须由有权限的主体改用通知并写明依据和执法主体（村（社区）无此权限，转请乡镇（街道）或执法部门依法办理）",
                         BASIS_INITIATIVE)

    # 五、群消息篇幅 ---------------------------------------------------------
    def check_group_message(self) -> None:
        if self.kind != "群消息":
            return
        count = nonspace_len(self.full_text)
        if count > GROUP_MSG_LIMIT:
            self.add("NTC-010", "P1", "篇幅", None, self.title,
                     f"群消息共 {count} 字，超过 {GROUP_MSG_LIMIT} 字，手机上一屏看不完",
                     "只留时间、地点、对象、要做的事和联系电话，其余放进张贴版", BASIS_GROUP)

    # 六、公示与村（居）务公开 -------------------------------------------------
    def check_publicity(self) -> None:
        if self.kind != "公示":
            return
        periods = self._periods()
        text_wo_title = "\n".join(text for _n, text in self.body[1:])
        if not PUBLICITY_PERIOD_RE.search(text_wo_title) and not periods:
            self.add("NTC-011", "P1", "要素", None, self.title,
                     "公示没有写明公示期（起止日期或天数）",
                     "写明公示起止日期，以及对公示内容有异议时的反映渠道", BASIS_ELEMENTS)
        if self._is_wugongkai():
            # 公布期限优先取“公开时间”“公布期限”行；没有的，逐个看不属于低保等专项公示的期限
            doc_level = [p for p in periods if p["doc_label"] and not p["exempt"]]
            targets = doc_level or [p for p in periods if not p["exempt"]]
            if periods and not targets:
                self.add("NTC-011", "P2", "要素", None, self.title,
                         "只写了低保等专项公示的期限，没有写这次村（居）务公开本身的公开时间",
                         f"另起一行写“公开时间：某月某日至某月某日”，不少于 {PUBLIC_DAYS_MIN} 日",
                         BASIS_PERIOD)
            for period in targets:
                if period["days"] < PUBLIC_DAYS_MIN:
                    self.add("NTC-012", "P1", "期限", period["line"], period["excerpt"],
                             f"村（居）务公开的公布期限为 {period['days']} 天，少于 {PUBLIC_DAYS_MIN} 日",
                             f"公布期限不少于 {PUBLIC_DAYS_MIN} 日；本地规定更长的按本地规定",
                             BASIS_PERIOD)
                elif period["days"] == PUBLIC_DAYS_MIN and period["from_range"]:
                    self._boundary(period, "村（居）务公开的公布期限", "结束日期往后延一天，或按本地惯例核对起算日")
        for period in periods:
            if not period["dibao"]:
                continue
            days = period["days"]
            if period["from_range"] and days == DIBAO_DAYS + 1:
                continue  # 开始当日不计入时正好 7 天
            if days == DIBAO_DAYS:
                if period["from_range"]:
                    self._boundary(period, "低保初审公示期", "按乡镇（街道）民政部门的算法核对起止日期")
                continue
            level = "P1" if days < DIBAO_DAYS else "P2"
            self.add("NTC-017", level, "期限", period["line"], period["excerpt"],
                     f"低保初审公示期写的是 {period['days']} 天，规定为 {DIBAO_DAYS} 天",
                     f"低保初审公示期为 {DIBAO_DAYS} 天；本地另有规定的，请用户提供依据",
                     BASIS_DIBAO)
        seen = set()
        for number, text in self.body:
            for match in SENSITIVE_RE.finditer(text):
                word = match.group(0)
                if word in seen:
                    continue
                seen.add(word)
                self.add("NTC-013", "P1", "个人信息", number, word,
                         f"公示中出现“{word}”，可能公开了与公示事项无关的个人信息",
                         "公示只列必要信息（如姓名、家庭人口、金额），删去病情、住址门牌、证件号等",
                         BASIS_SENSITIVE)

    def _boundary(self, period: Dict[str, Any], what: str, suggestion: str) -> None:
        days = period["days"]
        self.add("NTC-019", "P2", "期限", period["line"], period["excerpt"],
                 f"{what}按首尾两天都算正好 {days} 天，按开始当日不计入的算法只有 {days - 1} 天",
                 suggestion, BASIS_BOUNDARY)

    def _is_wugongkai(self) -> bool:
        """标题是村（居）务公开，或正文写明“村务公开”等；“村务监督委员会”“村务公开栏”不算。"""
        return bool(WUGONGKAI_TITLE_RE.search(self.title) or WUGONGKAI_TEXT_RE.search(self.full_text))

    def _adjacent(self, upper: int, lower: int) -> bool:
        """body 中相邻两行之间没有空行。"""
        return self.body[lower][0] - self.body[upper][0] == 1

    def _period_context(self, idx: int) -> str:
        """取判断期限性质用的上下文：本行；公示子项标签行加上一行；所在一级标题；独立公示的标题。

        只在期限行与上文之间没有空行时才继承上一行和小节标题：空一行单独写在末尾的期限行，
        是整份公开的期限（如低保小节之后、落款之前的“公示时间”）。
        """
        text = self.body[idx][1]
        parts = [text]
        if DOC_PERIOD_LABEL_RE.search(text) and not EXEMPT_PERIOD_RE.search(text):
            return text  # “公开时间”“公布期限”行是整份公开的期限
        if SUB_PERIOD_LABEL_RE.match(text) and idx > 1 and self._adjacent(idx - 1, idx):
            parts.append(self.body[idx - 1][1])
        for j in range(idx - 1, 0, -1):
            if not self._adjacent(j, j + 1):
                break
            if HEADING_RE.match(self.body[j][1]):
                parts.append(self.body[j][1])
                break
        if not self._is_wugongkai():
            parts.append(self.title)  # 单独一份公示：标题说明公示的是什么
        return "\n".join(parts)

    def _periods(self) -> List[Dict[str, Any]]:
        """逐行找公示、公开期限（起止日期或天数），跳过标题、成文日期行和一级标题行。"""
        sign_year, _m = signoff_year(self.body)
        default_year = self.arg_year or sign_year or _dt.date.today().year
        found: List[Dict[str, Any]] = []
        for idx in range(1, len(self.body)):
            number, text = self.body[idx]
            if DATE_LINE_RE.match(text) or HEADING_RE.match(text):
                continue
            if not re.search(r"公示|公布|公开", text):
                continue
            spans: List[Tuple[int, str, bool]] = []
            for match in DATE_RANGE_RE.finditer(text):
                days = self._range_days(match, default_year)
                if days is not None:
                    spans.append((days, match.group(0), True))
            for match in PERIOD_DAYS_RE.finditer(text):
                days = cn_to_int(match.group(1))
                if days is not None:
                    spans.append((days, match.group(0), False))
            if not spans:
                continue
            context = self._period_context(idx)
            exempt = bool(EXEMPT_PERIOD_RE.search(context))
            for days, excerpt, from_range in spans:
                found.append({
                    "line": number,
                    "days": days,
                    "excerpt": excerpt,
                    "from_range": from_range,
                    "doc_label": bool(DOC_PERIOD_LABEL_RE.search(text)),
                    "exempt": exempt,
                    "dibao": exempt and bool(DIBAO_RE.search(context)),
                })
        return found

    @staticmethod
    def _range_days(match: "re.Match[str]", default_year: int) -> Optional[int]:
        year1 = int(match.group(1)) if match.group(1) else default_year
        month1, day1 = int(match.group(2)), int(match.group(3))
        year2 = int(match.group(4)) if match.group(4) else year1
        month2 = int(match.group(5)) if match.group(5) else month1
        day2 = int(match.group(6))
        try:
            start = _dt.date(year1, month1, day1)
            end = _dt.date(year2, month2, day2)
        except ValueError:
            return None
        if end < start:
            try:
                end = _dt.date(year2 + 1, month2, day2)
            except ValueError:
                return None
        return (end - start).days + 1

    # 七、转发预警 -----------------------------------------------------------
    def check_prewarning(self) -> None:
        for number, text in self.body:
            match = PREWARN_RE.search(text)
            if not match:
                continue
            if not PREWARN_SOURCE_RE.search(text):
                self.add("NTC-014", "P1", "预警", number, match.group(0),
                         "提到预警或应急响应，但没有写明由哪个单位发布",
                         "照原文写明发布单位、发布时间和级别，如“县气象台10月16日8时发布暴雨橙色预警信号”；不改动预警内容",
                         BASIS_PREWARN)
            elif not RELEASE_TIME_RE.search(text):
                self.add("NTC-015", "P2", "预警", number, match.group(0),
                         "转发预警没有写发布时间", "补上发布时间，用户没提供的写【待补：预警发布时间】",
                         BASIS_PREWARN)

    # 八、已废止的法规 ---------------------------------------------------------
    def check_repealed(self) -> None:
        for name, replacement in REPEALED_LAWS:
            hits = [n for n, t in self.body if f"《{name}》" in t]
            if hits:
                extra = f"（全文共 {len(hits)} 处）" if len(hits) > 1 else ""
                self.add("NTC-016", "P1", "依据", hits[0], f"《{name}》",
                         f"引用的《{name}》已废止{extra}",
                         f"改为引用{replacement}，条号按新条例对应调整", BASIS_REPEALED)

    # 九、村（社区）常见的强制措施 ----------------------------------------------
    def _signed_by_grassroots(self) -> bool:
        """落款是村（居）委会、社区或村党组织。群消息看最后一句之后的署名。"""
        if self.kind == "群消息" and self.body:
            tail = re.split(r"[。！？!?]", self.body[-1][1].rstrip("。！？!? "))[-1].strip()
            if tail and nonspace_len(tail) <= 40 and GRASSROOTS_SIGN_RE.search(tail):
                return True
        for _n, text in self.body[-3:]:
            if DATE_LINE_RE.match(text) or "。" in text or text.endswith(("：", ":")):
                continue
            if nonspace_len(text) <= 40 and any(
                    GRASSROOTS_SIGN_RE.search(part) for part in re.split(r"\s+", text) if part):
                return True
        return False

    def check_coercion(self) -> None:
        if not self.public:
            return
        grassroots = self._signed_by_grassroots()
        for number, text in self.body:
            for sentence in SENTENCE_SPLIT_RE.split(text):
                for pattern, what in COERCION_EXTRA:
                    match = pattern.search(sentence)
                    if not match:
                        continue
                    has_basis = bool(BASIS_CUE_RE.search(sentence))
                    if grassroots and not has_basis:
                        level = "P0"
                        message = f"村（社区）发布的面向群众的文书{what}（“{match.group(0)}”），村（居）委会没有这项权力"
                    else:
                        level = "P1"
                        message = (f"面向群众的文书{what}（“{match.group(0)}”）；发文主体是村（社区）的按 P0"
                                   + ("；句中写了依据，请核对依据原文是否有这项措施、由谁执行" if has_basis else ""))
                    self.add("NTC-018", level, "强制", number, match.group(0), message,
                             "村（社区）发布的删去，改为提醒和服务（如“请不要进山”“有困难打村委会电话”）；"
                             "乡镇（街道）发布的，有法律依据和本级职权才写，并写明依据和执法主体",
                             BASIS_COERCION)

    def run(self) -> None:
        self.check_elements()
        self.check_weekday()
        self.check_plain()
        self.check_initiative()
        self.check_group_message()
        self.check_publicity()
        self.check_prewarning()
        self.check_repealed()
        self.check_coercion()
        order = {lv: i for i, lv in enumerate(LEVELS)}
        self.issues.sort(key=lambda i: (order[i["level"]], i["line"] or 0, i["id"]))


def check_file(path: Path, kind: Optional[str] = None, year: Optional[int] = None) -> Dict[str, Any]:
    lines, notes = read_lines(path)
    body, note_line = body_lines(lines)
    if not body:
        raise InputError("正文为空（【起草说明】之前没有内容）")
    detected, source = detect_kind(body[0][1], kind, path.stem)
    checker = Checker(body, detected, year)
    checker.run()
    all_notes = list(notes)
    if note_line is not None:
        all_notes.append(f"第 {note_line} 行为【起草说明】，其后内容未检查。")
    all_notes.extend(checker.notes)
    all_notes.append("罚款、拘留等法律表述和个人信息请另行运行：python scripts/gongwen_lint.py <文件> --checks legal,privacy；"
                     "两个脚本都可能漏报，按 references/review-addon.md 第一部分第 1 项的检索词再人工搜一遍。")
    summary = {lv: sum(1 for i in checker.issues if i["level"] == lv) for lv in LEVELS}
    return {
        "file": path.name,
        "kind": detected,
        "kind_source": source,
        "audience": "群众" if checker.public else "工作对象",
        "char_count": sum(nonspace_len(t) for _n, t in body),
        "summary": summary,
        "issues": checker.issues,
        "notes": all_notes,
        "version": VERSION,
    }


def render_text(result: Dict[str, Any]) -> str:
    out = [
        f"文件：{result['file']}　类型：{result['kind']}（{result['kind_source']}）　对象：{result['audience']}　字数：{result['char_count']}",
        "结果：P0 {P0} 条，P1 {P1} 条，P2 {P2} 条".format(**result["summary"]),
    ]
    for issue in result["issues"]:
        where = f"第 {issue['line']} 行" if issue["line"] else "全文"
        out.append(f"[{issue['level']}] {issue['id']} {where}｜{issue['excerpt']}")
        out.append(f"    问题：{issue['message']}")
        out.append(f"    建议：{issue['suggestion']}")
        out.append(f"    依据：{issue['basis']}")
    for note in result["notes"]:
        out.append(f"提示：{note}")
    return "\n".join(out)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="notice_check.py",
        description="面向群众文书体检：要素、日期与星期、通俗程度、倡议书语气、公示期限、预警来源、已废止法规、"
                    "村（社区）常见的强制措施。",
        epilog="退出码：0 正常；1 使用 --strict 且存在 P0；2 输入错误。罚款、拘留等法律表述和个人信息请用 gongwen_lint.py 检查。",
    )
    parser.add_argument("file", help="要检查的文件（.txt、.md、.docx），读到单独一行的【起草说明】为止")
    parser.add_argument("--kind", choices=KINDS, help="文书类型；不填时从标题推断，文件名含“群消息”的按群消息处理")
    parser.add_argument("--year", type=int, help="日期没写年份时，用这个年份核对星期")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    parser.add_argument("--strict", action="store_true", help="存在 P0 时返回 1")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
        except Exception:
            pass
    args = build_parser().parse_args(argv)
    if args.year is not None and not 1900 <= args.year <= 2100:
        print("错误：--year 应为 1900 至 2100 之间的四位年份", file=sys.stderr)
        return 2
    try:
        result = check_file(Path(args.file), args.kind, args.year)
    except InputError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(render_text(result))
    if args.strict and result["summary"]["P0"] > 0:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
