# -*- coding: utf-8 -*-
"""改稿前后事实一致性检查。

源文件位于仓库 shared/scripts/，各 skill 中的副本由 tools/sync_shared.py 生成，请在源文件处修改。
维护者：gongwen-humanize。接口见 docs/开发规范.md 8.1 节。

用法：
    python scripts/fact_diff.py <原稿> <改稿> [--ref <另一份原稿>]... [--json] [--strict]
    python scripts/fact_diff.py <素材台账.md> <正文.txt> --ledger [--ref <另一份材料>]... [--json] [--strict]

--ledger（台账对账）：只输出"改稿新增"，即正文中有、台账中没有的事实；按集合比较，同一个数在正文中
用两次不报；台账中 Markdown 表格表头为"序号""来源"的列不参与比对。

--ref（多份原稿，可重复）：多稿合并、原稿加补充材料时使用。各稿分别截掉【起草说明】及其后内容，
合起来作比对基准；同一事实按出现次数最多的那一份计数；"原稿有而改稿没有"的每项注明出自哪份原稿。

比对内容（只比对【起草说明】之前的正文）：
    数字：阿拉伯数字和带量词的汉字数字，连同单位一起比对，如"3.5天""12.5%""2个百分点"；
          汉字数字换算成阿拉伯数字后比对，"三家"与"3家"视为相同；"三分之一"按"1/3"比对；
          含万、亿的汉字数字不带量词也比对，"三百二十万"按"320万"比对，"一点五亿"按"1.5亿"比对；
          "万"与"万元"、"亿"与"亿元"在一方缺"元"时按同一数值比对（只限这两种）；
          数值范围中前一个数省略的单位按后一个数补全，"3～5万元"与"3万～5万元"视为相同；
    日期：如"2026年10月16日""2026-10-16"；时刻如"9:30"；月、日的虚位 0 不计，"11月05日"与"11月5日"视为相同；
    标记：【待补：……】【待核实：……】；
    书名号：《……》整体比对；
    引号：“……”‘……’「……」『……』及成对的英文双引号，整体比对；英文双引号、「」按“”比对，『』按‘’比对；
    发文字号：如"某政发〔2026〕12号"。

比对前去掉数字与单位、量词之间的空格（含全角空格，不含制表符），"3216 人次"与"3216人次"视为相同；
"第 3 次""2026 年 10 月 16 日"中的空格同样去掉；紧挨英文字母的编号（"E1 家校协同"）不并入后面的文字。

不算事实、不参与比对的内容：
    层次序号 一、（一）1.（1），段内的 1. 1、1）（1）列举序号，作列举用的"第一，""第二、"；
    单独的汉字"一"加量词（如"一个""一批"），以及不带量词、不含万亿的汉字数字（如"统一""一是"）。

可能是序数或概数的内容单列第三张清单"可能是序数或概数（请人工判断）"，不计入前两张清单和 --strict：
    结构序数：十以内的"X点意见""X个方面""X项要求""X条建议""X件事"，"讲三点""以下三点"，"第三点"；
    概数："X成"，如"九成以上""9成患者满意"；"数""几"开头的汉字数字，如"数十万人次""几十家"；
    概括数：做法、措施、经验等的总条数，如"具体做法有三条""采取了三条措施""措施共5项"，提示"概括数，请核对与分项一致"
            （P2）；带"新增""发现""完成""培训"等业务动词的，或后面不是概括词的（"新增5项制度"），仍按业务数字比对。

前两张清单为"原稿有而改稿没有""改稿新增"。按出现次数比对：原稿出现 2 次、改稿 1 次，也会列出。
--json 输出中，第三张清单在 uncertain 字段（每项带 hint：序数、概数或概括数；list：missing 或 added），
docx 修订等说明在 notes 字段，各原稿文件名在 original_files 字段。

读取 .docx 时按"接受全部修订"后的文字比对：含修订插入（w:ins）的文字，不含修订删除（w:del）的文字，
输出中注明"已按接受修订后的文字比对"。

退出码：0 正常；--strict 时前两张清单任一非空返回 1；输入错误（文件不存在、格式不支持、读取失败）返回 2。
只依赖标准库（.docx 用 zipfile 和 ElementTree 解析）。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, OrderedDict
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

BOUNDARY = "【起草说明】"
# 分界行：去掉行首空白后以【起草说明】开头；兼容 Markdown 标题、加粗包裹的写法。
# 与 gongwen_lint.py 的 DRAFT_NOTE_RE、gongwen-format 的 gw_layout.is_draft_note_line 保持一致。
DRAFT_NOTE_RE = re.compile(r"^\s*(?:#{1,6}\s*)?(?:(?:\*\*|__)\s*)?" + re.escape(BOUNDARY))
SUPPORTED_SUFFIXES = (".txt", ".md", ".docx")
MASK = "\ue000"  # 私用区字符，用于屏蔽已识别的片段，长度不变，便于回溯原文位置
MAX_LOCATIONS = 3
EXCERPT_RADIUS = 14
VALUE_DISPLAY_LIMIT = 60

# 类型名称（输出用）
T_NUMBER = "数字"
T_DATE = "日期"
T_MARKER = "标记"
T_TITLE = "书名号"
T_QUOTE = "引号"
T_DOCNO = "发文字号"
TYPE_ORDER = (T_NUMBER, T_DATE, T_MARKER, T_TITLE, T_QUOTE, T_DOCNO)

# 全角数字和符号转半角（不用 NFKC，避免把 ① 之类的序号转成数字）
FULLWIDTH_TABLE = str.maketrans("０１２３４５６７８９％．", "0123456789%.")

# ---------------------------------------------------------------------------
# 正则
# ---------------------------------------------------------------------------

MARKER_RE = re.compile(r"【(?:待补|待核实)(?:[：:][^】\n]*)?】")
DOCNO_RE = re.compile(
    r"([\u4e00-\u9fffA-Za-z]{0,12})"
    r"([〔\[［(（【])(\d{4})([〕\]］)）】])第?(\d{1,5})号"
)
# 发文字号前面常见的非代字成分，截掉最后一次出现及其之前的部分，使改写上下文不影响比对。
# 单字只收不会出现在机关代字中的介词、连词（按、依、据、以、在、为、是、由、从、见、即、并、把）
DOCNO_PREFIX_CUT_RE = re.compile(
    r".*(?:根据|按照|依据|依照|参照|对照|遵照|收到|接到|贯彻|落实|学习|传达|印发|转发|执行|你局|你委|你厅|你办|你院|你校|"
    r"你公司|你单位|贵单位|贵局|贵委|贵厅|我局|我委|我厅|我院|我校|本局|本院|关于|文件|的|了|和|及|与|将|"
    r"按|依|据|以|在|为|是|由|从|见|即|并|把)"
)
TITLE_RE = re.compile(r"《[^《》\n]{1,200}》")
QUOTE_RES = (
    re.compile(r"“[^“”\n]{1,300}”"),
    re.compile(r"「[^「」\n]{1,300}」"),
    re.compile(r"『[^『』\n]{1,300}』"),
    re.compile(r"‘[^‘’\n]{1,300}’"),
    re.compile(r"\"[^\"\n]{1,300}\""),
)

# 层次序号与列举序号
MD_PREFIX = r"(?:#{1,6}\s*|[-*+]\s+|>\s*)?(?:\*\*)?"
SEQ_LINE_START_RE = re.compile(
    r"^\s*" + MD_PREFIX + r"(?:[一二三四五六七八九十]{1,3}[、.．]|"
    r"[（(][一二三四五六七八九十]{1,3}[）)]|"
    r"\d{1,3}[.．、](?!\d)|"
    r"[（(]\d{1,3}[）)]|"
    r"\d{1,3}[)）])"
)
SEQ_INLINE_RES = (
    re.compile(r"[（(][一二三四五六七八九十]{1,3}[）)]"),
    re.compile(r"[（(]\d{1,2}[）)]"),
    re.compile(r"(?:(?<=^)|(?<=[：:；;。！？!?，,\s\ue000]))\d{1,2}[.．、)）](?!\d)"),
    # 作列举用的序数词："第一，""第二、""第三：""其一，"
    re.compile(r"第[一二三四五六七八九十\d]{1,3}(?=[，,、：:；;])"),
)

DATE_RES = (
    re.compile(r"(?<!\d)\d{4}年\d{1,2}月(?:\d{1,2}[日号])?"),
    re.compile(r"(?<!\d)\d{1,2}月\d{1,2}[日号]"),
    re.compile(r"(?<![\d.])\d{4}[-/.]\d{1,2}[-/.]\d{1,2}(?![\d.])"),
)
TIME_RE = re.compile(r"(?<![\d.])(\d{1,2})[:：](\d{2})(?![\d])")
# 日期中月、日的虚位 0（年份中的 0 前面是数字，不受影响）
LEADING_ZERO_RE = re.compile(r"(?<!\d)0+(?=\d)")
# 数值范围的连接号：前一个数没写单位或"万""亿"时，沿用后一个数的（"3～5万元"与"3万～5万元"视为相同）
RANGE_SEP_RE = re.compile(r"\s*(?:[～~〜－\-]|—{1,2}|至|到)\s*")
# 汉字书写的日期，如"二〇二六年十一月十六日""十月十六日"，换算成阿拉伯数字后与阿拉伯数字日期统一比对
CN_DATE_RES = (
    re.compile(r"([〇零一二三四五六七八九]{4})年([一二三四五六七八九十]{1,3})月(?:([一二三四五六七八九十]{1,3})[日号])?"),
    re.compile(r"(?<![〇零一二三四五六七八九十])()([一二三四五六七八九十]{1,3})月([一二三四五六七八九十]{1,3})[日号]"),
)

CN_DIGITS = "零〇一二两三四五六七八九"
CN_SMALL_UNITS = "十百千"
CN_BIG_UNITS = "万亿"
CN_NUM_CHARS = CN_DIGITS + CN_SMALL_UNITS
PERCENT_CN_RE = re.compile(r"百分之([" + CN_NUM_CHARS + r"]+)")
# 汉字分数“三分之一”“千分之五”（“百分之”由 PERCENT_CN_RE 先行处理）
FRACTION_CN_RE = re.compile(r"([" + CN_NUM_CHARS + r"万]+)分之([" + CN_NUM_CHARS + r"]+)")
_FRACTION_UNIT_DENOM = {"十": 10, "百": 100, "千": 1000, "万": 10000}

ARABIC_RE = re.compile(
    r"(?<![\d.])(第)?(\d+(?:,\d{3})+|\d+)(\.\d+)?(余|多)?([十百千万亿]{1,2})?(余|多)?"
)
CN_RE = re.compile(
    r"(第)?([" + CN_NUM_CHARS + r"]+(?:[万亿][" + CN_NUM_CHARS + r"]+)*)(余|多)?([万亿]{1,2})?(余|多)?"
)
# 汉字小数"一点五亿""三点五个百分点"；小数部分逐位书写，后面紧跟十百千的是时刻（"十二点三十分"），不算
CN_DECIMAL_RE = re.compile(
    r"([" + CN_NUM_CHARS + r"]+)点([〇零一二三四五六七八九]+)(?![十百千])([万亿])?"
)

# 单位：先匹配多字单位，再匹配单字单位
MULTI_UNITS = (
    "个百分点", "百分点", "个工作日", "工作日", "个月", "万人次", "人次", "例次", "台次", "户次", "架次",
    "平方公里", "平方千米", "平方米", "立方米", "公里", "千米", "公顷", "千克", "公斤", "小时", "分钟",
    "课时", "学时", "学分", "季度", "周岁", "周年",
)
SINGLE_UNITS = frozenset(
    "%‰个项次家名人位所件条批期轮场起篇份台辆处座户村镇张本床例套栋间册门节"
    "年月日号时秒天周岁倍元角吨亩米层级类种支队组届点股斤箱瓶块段部章成"
)
# 含数字但不表示数量的成语和惯用语，抽取汉字数字前先屏蔽
CN_IDIOMS = ("第一时间", "千家万户", "百家争鸣", "三番五次", "三天两头", "五花八门", "一五一十", "三令五申",
             "十万火急")
LATIN_UNIT_RE = re.compile(r"[A-Za-z]{1,3}[²³]?(?![A-Za-z])")

# 数字与单位之间可去掉的空白：半角、不换行、全角空格等；不含制表符（表格各列以制表符分隔）
SPACE_CHARS = frozenset("      　")
DIGIT_CHARS = frozenset("0123456789０１２３４５６７８９")
UNIT_LEAD_CHARS = frozenset("%‰％万亿余多")
# "万"与"万元"、"亿"与"亿元"按同一数值比对：比对键去掉万、亿（可带余、多）之后的"元"
YUAN_AFTER_MULT_RE = re.compile(r"(?<=[万亿])(余|多)?元$")
# 前面是"数""几"的汉字数字（"数十万人次"）属概数，列入第三张清单
CN_VAGUE_PREFIXES = frozenset("数几")
# 引号统一：英文双引号、直角引号按中文双引号比对，『』按中文单引号比对
QUOTE_NORM = {"\"": ("“", "”"), "“": ("“", "”"), "「": ("“", "”"), "‘": ("‘", "’"), "『": ("‘", "’")}

# 第三张清单：可能是序数或概数
HINT_ORDINAL = "序数"
HINT_APPROX = "概数"
STRUCT_MAX = 10  # 结构序数只认十以内
# 量词后紧跟这些词时，多为讲话、材料的结构序数（"三点意见""五项要求""两件事""三个方面"）
STRUCT_FOLLOW = {
    "点": ("意见", "要求", "建议", "看法", "体会", "想法", "打算", "感受", "认识", "希望", "考虑", "思考",
          "启示", "收获"),
    "条": ("意见", "建议", "要求", "体会", "启示", "感受"),
    "项": ("要求", "意见", "建议", "原则"),
    "件": ("事",),
    "个": ("方面",),
}
STRUCT_NOT_FOLLOW = ("事项", "事业", "事务", "事件", "事宜", "事迹")
# "讲三点""以下三点"：量词"点"前面是这些词时按结构序数（"下午3点"不算）
POINT_LEAD_RE = re.compile(r"(?:讲|谈|提|说|强调|以下|如下|下面|这)$")

# 概括数：做法、措施等的“总条数”（“具体做法有三条”“采取了三条措施”“措施共5项”），数字本身不是新事实，
# 只要与后面的分项个数一致即可；单列为第三张清单的“概括数”，提示“概括数，请核对与分项一致”（P2）。
# 判断标准（宁可少放行，拿不准的一律留在“正文有、台账无”）：
#   1. 数字是十以内的整数，量词只限条、项、个、点、类、种，或阿拉伯数字后直接接“方面”；
#   2. 数字前面（到最近一个句读为止，最多 8 个字）不能出现业务动词：新增、发现、完成、培训、出台、整改等
#      带这些动词的是真实的业务数量（“新增3项措施”“发现3个问题”），不放行；
#   3. 满足下列任一形式才放行：
#      甲、数字在前：数字+量词后接概括词（可夹“主要”“重点”“具体”等修饰语，后面可带“的”），
#          概括词分强弱——强概括词（做法、措施、举措、经验、方法、办法、亮点、特点、步骤、要点、对策、方面）直接放行；
#          弱概括词（任务、问题、原因、成效、目标）容易是真实数量（“3项任务已完成”），
#          只在数字前紧挨“主要”“重点”“以下”“如下”“包括”“分为”等引导词时放行；
#      乙、数字在后：强概括词（不含任务、问题等弱概括词）+“有/共有/共/为/是/包括/分为”+数字+量词，
#          且量词后是句读或行尾（“具体做法有三条。”“措施共5项：”）；
#   4. “新增5项制度”（制度不是概括词）、“培训3场次”（量词不在限定范围）、“措施有12项”（超过十）都不放行。
HINT_SUMMARY = "概括数"
SUMMARY_NOTE = "概括数，请核对与分项一致"
SUMMARY_LEVEL = "P2"
SUMMARY_UNITS = frozenset("条项个点类种")
_SUMMARY_STRONG = "做法|措施|举措|经验|方法|办法|亮点|特点|步骤|要点|对策|方面"
_SUMMARY_WEAK = "任务|问题|原因|成效|目标"
_SUMMARY_MODIFIER = r"(?:(?:主要|重点|具体|基本|重要|突出|关键|核心|有效|典型)的?)*"
SUMMARY_AFTER_STRONG_RE = re.compile(_SUMMARY_MODIFIER + "(?:" + _SUMMARY_STRONG + ")")
SUMMARY_AFTER_WEAK_RE = re.compile(_SUMMARY_MODIFIER + "(?:" + _SUMMARY_WEAK + ")")
# 弱概括词的引导词：数字前紧挨这些词（可再带一个“有”）才算概括
SUMMARY_WEAK_LEAD_RE = re.compile(r"(?:主要|重点|突出|以下|如下|下面|其中|着重|集中|聚焦|归纳为|概括为|总结为|梳理出|分为|包括)有?$")
# 数字在后的形式：强概括词（可带修饰语）+连接词，紧挨数字
SUMMARY_BEFORE_RE = re.compile(
    "(?:" + _SUMMARY_STRONG + ")(?:主要|重点|大致|大体|基本)?(?:共有|有以下|有如下|可概括为|可归纳为|包括|分为|有|共|为|是)$"
)
SUMMARY_END_CHARS = frozenset("。，、；：:,;.！!")
# 业务动词：数字前面出现这些词，是真实的业务数量，不按概括数放行
SUMMARY_BUSINESS_RE = re.compile(
    "新增|新建|新设|新开|新上|新出台|新推出|新制定|增设|增加|减少|压减|取消|精简|推出|出台|制定|印发|发布|"
    "建立|建成|设立|设置|完成|办结|办理|解决|整改|落实|兑现|发放|开展|举办|组织|培训|实施|投入|收到|受理|"
    "处理|排查|查处|查出|发现|清理|拆除|落地|上线|入库|签约|立案|约谈|通报|曝光"
)
_CLAUSE_BREAKS = "，。；：、,;:.（）()！!？?\n"


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------

class Fact:
    """一项事实。value 用于显示；norm 用于比对（去掉单位前空格、统一引号、万元按万等）；
    hint 非空时表示可能是序数或概数；source 为所在文件名（多份原稿时用于注明出处）。"""

    __slots__ = ("kind", "value", "line", "pos", "excerpt", "norm", "hint", "source")

    def __init__(self, kind: str, value: str, line: int, pos: int, excerpt: str,
                 norm: Optional[str] = None, hint: str = "", source: str = "") -> None:
        self.kind = kind
        self.value = value
        self.line = line
        self.pos = pos
        self.excerpt = excerpt
        self.norm = value if norm is None else norm
        self.hint = hint
        self.source = source

    @property
    def key(self) -> Tuple[str, str, str]:
        # 序数、概数与同写法的数量分开计数："讲三点"与"下午3点"不相抵
        return (self.kind, self.norm, self.hint)


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


# .docx 解析（标准库 zipfile + ElementTree），按"接受全部修订"后的文字读取
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W_BODY, W_P, W_T, W_TBL, W_TR, W_TC = (W_NS + t for t in ("body", "p", "t", "tbl", "tr", "tc"))
W_TAB, W_BR, W_CR, W_SECTPR = (W_NS + t for t in ("tab", "br", "cr", "sectPr"))
W_INS, W_DEL, W_MOVE_FROM, W_MOVE_TO = (W_NS + t for t in ("ins", "del", "moveFrom", "moveTo"))
MAX_XML_BYTES = 50 * 1024 * 1024


def _docx_para_text(paragraph: ET.Element) -> str:
    """段落文字：含 w:ins、w:moveTo 内的文字，不含 w:del、w:moveFrom 内的文字；文本框等嵌套段落不并入。"""
    parts: List[str] = []

    def walk(element: ET.Element) -> None:
        for child in element:
            tag = child.tag
            if tag in (W_DEL, W_MOVE_FROM, W_P):
                continue
            if tag == W_T:
                parts.append(child.text or "")
            elif tag == W_TAB:
                parts.append("\t")
            elif tag in (W_BR, W_CR):
                parts.append("\n")
            else:
                walk(child)

    walk(paragraph)
    return "".join(parts)


def _docx_lines(body: ET.Element) -> List[str]:
    """正文逐段取文字；表格一行为一行，单元格之间用制表符分隔。

    合并单元格在 XML 中只有一个 w:tc（纵向合并的后续行是空单元格），一处数字不会计成多处；
    嵌套表格的文字并入所在单元格。
    """
    lines: List[str] = []

    def cell_text(cell: ET.Element) -> str:
        texts: List[str] = []

        def walk(element: ET.Element) -> None:
            for child in element:
                if child.tag == W_P:
                    texts.append(_docx_para_text(child).replace("\n", " ").strip())
                else:
                    walk(child)

        walk(cell)
        return " ".join(t for t in texts if t)

    def walk(element: ET.Element) -> None:
        for child in element:
            tag = child.tag
            if tag == W_P:
                lines.extend(_docx_para_text(child).split("\n"))
            elif tag == W_TBL:
                for row in (r for r in child if r.tag == W_TR):
                    lines.append("\t".join(cell_text(tc) for tc in row if tc.tag == W_TC))
            elif tag != W_SECTPR:
                walk(child)  # 内容控件（w:sdt）、w:customXml 等容器

    walk(body)
    return lines


def _read_docx(path: Path) -> Tuple[str, int]:
    """返回（接受全部修订后的正文，修订痕迹处数）。"""
    try:
        with zipfile.ZipFile(str(path)) as archive:
            try:
                info = archive.getinfo("word/document.xml")
            except KeyError:
                raise InputError(f"不是有效的 Word 文档（缺少 word/document.xml）：{path}")
            if info.file_size > MAX_XML_BYTES:
                raise InputError(f"word/document.xml 过大，已停止读取：{path}")
            data = archive.read(info)
    except (zipfile.BadZipFile, OSError, RuntimeError) as exc:  # 文件损坏、不是 docx、加密等
        raise InputError(f"无法打开 docx 文件：{path}（{exc}）")
    try:
        root = ET.fromstring(data)
    except ET.ParseError as exc:
        raise InputError(f"docx 文件内容损坏：{path}（{exc}）")
    body = root.find(W_BODY)
    revisions = sum(1 for el in root.iter() if el.tag in (W_INS, W_DEL, W_MOVE_FROM, W_MOVE_TO))
    return "\n".join(_docx_lines(root if body is None else body)), revisions


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


def read_document(path: Path) -> Tuple[str, int]:
    """读取 .txt、.md、.docx 文件的全文，返回（全文，docx 修订痕迹处数）。"""
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
        return _read_docx(path)
    try:
        data = path.read_bytes()
    except OSError as exc:
        raise InputError(f"读取失败：{path}（{exc}）")
    return _decode(data, path), 0


def read_text(path: Path) -> str:
    """读取 .txt、.md、.docx 文件的全文（docx 按接受全部修订后的文字）。"""
    return read_document(path)[0]


TABLE_ROW_RE = re.compile(r"^\s*\|.*\|\s*$")
TABLE_SEP_RE = re.compile(r"^\s*\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)*\|?\s*$")
LEDGER_SKIP_COLUMNS = ("序号", "来源")


def _table_cells(row: str) -> List[str]:
    return [cell.strip() for cell in row.strip().strip("|").split("|")]


def ledger_lines(lines: Sequence[str]) -> List[str]:
    """台账对账用：去掉 Markdown 表格中表头为"序号""来源"的列，其余单元格用制表符连接。

    分隔行（|---|）变为空行，行数不变，便于按行号定位。表格以外的行原样保留。
    """
    result: List[str] = []
    skip: Optional[set] = None  # 当前表格中要去掉的列号；None 表示不在表格中
    for i, line in enumerate(lines):
        if not TABLE_ROW_RE.match(line):
            skip = None
            result.append(line)
            continue
        if TABLE_SEP_RE.match(line):
            result.append("")
            continue
        cells = _table_cells(line)
        if skip is None:  # 表头：下一行是分隔行
            nxt = lines[i + 1] if i + 1 < len(lines) else ""
            names = [re.sub(r"[\s*_]", "", c) for c in cells]
            skip = {k for k, name in enumerate(names) if name in LEDGER_SKIP_COLUMNS} \
                if TABLE_SEP_RE.match(nxt) else set()
        result.append("\t".join(c for k, c in enumerate(cells) if k not in skip))
    return result


def body_lines(text: str) -> List[str]:
    """返回【起草说明】之前的正文行。"""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines: List[str] = []
    for line in text.split("\n"):
        if DRAFT_NOTE_RE.match(line):
            break
        lines.append(line)
    return lines


# ---------------------------------------------------------------------------
# 汉字数字
# ---------------------------------------------------------------------------

_CN_DIGIT_VALUE = {"零": 0, "〇": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
                   "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
_CN_SMALL_VALUE = {"十": 10, "百": 100, "千": 1000}
_CN_BIG_VALUE = {"万": 10 ** 4, "亿": 10 ** 8}


def cn_to_int(text: str) -> Optional[int]:
    """把汉字数字转成整数，不合法时返回 None。

    支持"十五""三百零五""一万二千"，以及逐位书写的年份"二〇二六"。
    "三四"（概数）、"十四五"之类不是确定数值，返回 None。
    """
    if not text:
        return None
    if all(ch in _CN_DIGIT_VALUE for ch in text):
        if len(text) == 1:
            return _CN_DIGIT_VALUE[text]
        if "〇" in text or "零" in text or len(text) == 4:
            if "两" in text:
                return None
            return int("".join(str(_CN_DIGIT_VALUE[ch]) for ch in text))
        return None
    no_unit = 10 ** 9
    total = 0
    section = 0
    number: Optional[int] = None
    prev = ""
    last_small_unit = no_unit
    for ch in text:
        if ch in _CN_DIGIT_VALUE:
            if number is not None and prev not in ("零", "〇"):
                return None  # 连续数字，如"三四"（概数）
            number = _CN_DIGIT_VALUE[ch]
        elif ch in _CN_SMALL_VALUE:
            unit = _CN_SMALL_VALUE[ch]
            if unit >= last_small_unit:
                return None  # 单位不递减，如"十百""千千"
            if number is None:
                if ch != "十":
                    return None
                number = 1  # "十五"中的"十"
            section += number * unit
            number = None
            last_small_unit = unit
        elif ch in _CN_BIG_VALUE:
            if number is not None:
                section += number
            if section == 0:
                return None
            total += section * _CN_BIG_VALUE[ch]
            section = 0
            number = None
            last_small_unit = no_unit
        else:
            return None
        prev = ch
    if number is not None:
        before = text[-2] if len(text) >= 2 else ""
        if before in ("百", "千"):
            number = number * _CN_SMALL_VALUE[before] // 10  # "三百五"按 350
        elif before in _CN_BIG_VALUE:
            number = number * _CN_BIG_VALUE[before] // 10  # "一万五"按 15000
        section += number
    return total + section


# ---------------------------------------------------------------------------
# 抽取
# ---------------------------------------------------------------------------

def _excerpt(line: str, start: int, end: int) -> str:
    left = max(0, start - EXCERPT_RADIUS)
    right = min(len(line), end + EXCERPT_RADIUS)
    text = line[left:right].strip()
    return ("……" if left > 0 else "") + text + ("……" if right < len(line) else "")


def _mask(buf: List[str], start: int, end: int) -> None:
    for i in range(start, end):
        buf[i] = MASK


def _norm_number(int_part: str, frac: Optional[str]) -> str:
    int_part = int_part.replace(",", "")
    int_part = int_part.lstrip("0") or "0"
    if frac:
        frac_digits = frac[1:].rstrip("0")
        if frac_digits:
            return f"{int_part}.{frac_digits}"
    return int_part


def _match_unit(text: str, pos: int) -> str:
    """从 pos 开始匹配单位，返回单位文字（可能为空）。允许数字与 % 之间有一个空格。"""
    probe = pos
    if probe < len(text) and text[probe] == " " and probe + 1 < len(text) and text[probe + 1] in "%‰":
        probe += 1
    for unit in MULTI_UNITS:
        if text.startswith(unit, probe):
            return text[pos:probe] + unit
    if probe < len(text) and text[probe] in SINGLE_UNITS:
        return text[pos:probe] + text[probe]
    match = LATIN_UNIT_RE.match(text, pos)
    if match:
        return match.group(0)
    return ""


def _scan_spans(regex: "re.Pattern[str]", masked: str,
                allow_inner_mask: bool = False) -> List[Tuple[int, int, "re.Match[str]"]]:
    """在已屏蔽的文本中查找匹配。默认跳过含屏蔽字符的匹配；引号、书名号允许内部含已识别片段。"""
    spans = []
    for m in regex.finditer(masked):
        text = m.group(0)
        if MASK in text:
            if not allow_inner_mask or text[0] == MASK or text[-1] == MASK:
                continue
        spans.append((m.start(), m.end(), m))
    return spans


def _docno_value(match: "re.Match[str]") -> Tuple[str, int]:
    """返回 (发文字号文本, 起始位置偏移)，截掉代字前面的动词、介词等。"""
    prefix = match.group(1)
    cut = DOCNO_PREFIX_CUT_RE.match(prefix)
    offset = cut.end() if cut else 0
    return match.group(0)[offset:], offset


def _unit_follows(text: str, pos: int) -> bool:
    """pos 处是否以单位、量词或万、亿、余、多开头。"""
    ch = text[pos]
    if ch in UNIT_LEAD_CHARS or ch in SINGLE_UNITS:
        return True
    return any(text.startswith(unit, pos) for unit in MULTI_UNITS)


def _space_removable(kept: List[str], text: str, nxt: int) -> bool:
    """kept 为已保留的字符，nxt 为空白之后第一个字符的位置。"""
    prev, following = kept[-1], text[nxt]
    if prev in DIGIT_CHARS:
        # 数字前紧挨英文字母的是编号（"指标E1 家校协同"），不并入后面的文字
        k = len(kept) - 1
        while k >= 0 and (kept[k] in DIGIT_CHARS or kept[k] in ".,．"):
            k -= 1
        if k >= 0 and (kept[k].isascii() and kept[k].isalpha() or kept[k] in "-_"):
            return False
        return _unit_follows(text, nxt)  # "3216 人次""12 %""320 万元"
    if prev == "第":
        return following in DIGIT_CHARS  # "第 3 次"
    if prev in "年月" and len(kept) >= 2 and kept[-2] in DIGIT_CHARS:
        return following in DIGIT_CHARS  # "2026 年 10 月 16 日"
    return False


def squeeze_spaces(line: str) -> Tuple[str, List[int]]:
    """去掉数字与单位、量词之间的空白，返回（新文本，新文本每个字符在原行中的位置）。"""
    kept: List[str] = []
    positions: List[int] = []
    i, n = 0, len(line)
    while i < n:
        if line[i] not in SPACE_CHARS:
            kept.append(line[i])
            positions.append(i)
            i += 1
            continue
        j = i
        while j < n and line[j] in SPACE_CHARS:
            j += 1
        if not (kept and j < n and _space_removable(kept, line, j)):
            kept.extend(line[i:j])
            positions.extend(range(i, j))
        i = j
    return "".join(kept), positions


def _number_norm(value: str) -> str:
    """数字的比对键："320万元"与"320万"、"1.5亿元"与"1.5亿"相同。"""
    return YUAN_AFTER_MULT_RE.sub(lambda m: m.group(1) or "", value)


def _quote_norm(text: str) -> str:
    """引号的比对键：英文双引号、「」统一为“”，『』统一为‘’。"""
    pair = QUOTE_NORM.get(text[:1])
    return pair[0] + text[1:-1] + pair[1] if pair and len(text) >= 2 else text


def _number_hint(text: str, start: int, end: int, unit: str, value: Optional[int], ordinal: bool) -> str:
    """可能是序数或概数时返回 HINT_ORDINAL 或 HINT_APPROX，否则返回空字符串。"""
    if unit == "成":
        return HINT_APPROX  # "九成以上""9成患者满意"
    if value is None or not 1 <= value <= STRUCT_MAX:
        return ""
    if ordinal:
        return HINT_ORDINAL if unit == "点" else ""  # "第三点"
    follows = STRUCT_FOLLOW.get(unit)
    if follows and any(text.startswith(word, end) for word in follows) \
            and not any(text.startswith(word, end) for word in STRUCT_NOT_FOLLOW):
        return HINT_ORDINAL  # "三点意见""五项要求""两件事""三个方面"
    if unit == "点" and POINT_LEAD_RE.search(text[max(0, start - 2):start]):
        return HINT_ORDINAL  # "讲三点""以下三点"
    if _is_summary_count(text, start, end, unit):
        return HINT_SUMMARY  # "具体做法有三条""采取了三条措施"
    return ""


def _is_summary_count(text: str, start: int, end: int, unit: str) -> bool:
    """是否为做法、措施等的纯概括数，判断标准见 HINT_SUMMARY 处的注释。value 已在调用处限定为十以内整数。"""
    window = text[max(0, start - 8):start]
    for brk in _CLAUSE_BREAKS:
        cut = window.rfind(brk)
        if cut >= 0:
            window = window[cut + 1:]
    if SUMMARY_BUSINESS_RE.search(window):
        return False
    if unit == "" and text.startswith("方面", end):
        return True  # 阿拉伯数字直接接“方面”："3方面"
    if unit not in SUMMARY_UNITS:
        return False
    if SUMMARY_AFTER_STRONG_RE.match(text, end):
        return True
    if SUMMARY_AFTER_WEAK_RE.match(text, end) and SUMMARY_WEAK_LEAD_RE.search(window):
        return True
    if SUMMARY_BEFORE_RE.search(text[max(0, start - 12):start]):
        return end >= len(text) or text[end] in SUMMARY_END_CHARS
    return False


def extract_line(line: str, lineno: int) -> List[Fact]:
    """抽取一行正文中的事实。"""
    facts: List[Fact] = []
    original = line
    stripped, positions = squeeze_spaces(line)  # 去掉数字与单位之间的空白，positions 用于回溯原文位置
    norm = stripped.translate(FULLWIDTH_TABLE)  # 逐字替换，长度不变
    buf = list(norm)

    def current() -> str:
        return "".join(buf)

    def add(kind: str, value: str, start: int, end: int, key: Optional[str] = None, hint: str = "") -> None:
        o_start, o_end = positions[start], positions[end - 1] + 1
        if key is None:
            key = value
        if kind in (T_MARKER, T_TITLE, T_QUOTE):
            value = original[o_start:o_end]  # 显示原文写法
        facts.append(Fact(kind, value, lineno, o_start, _excerpt(original, o_start, o_end), key, hint))

    # 1. 标记
    for start, end, m in _scan_spans(MARKER_RE, current()):
        add(T_MARKER, stripped[start:end], start, end)
        _mask(buf, start, end)
    # 2. 发文字号
    for start, end, m in _scan_spans(DOCNO_RE, current()):
        value, offset = _docno_value(m)
        add(T_DOCNO, value, start + offset, end)
        _mask(buf, start + offset, end)
    # 3. 书名号（内部可含已识别的发文字号、标记）
    for start, end, m in _scan_spans(TITLE_RE, current(), allow_inner_mask=True):
        add(T_TITLE, stripped[start:end], start, end)
        _mask(buf, start, end)
    # 4. 引号（内部可含已识别的书名号等）；中英文引号统一后比对
    for regex in QUOTE_RES:
        for start, end, m in _scan_spans(regex, current(), allow_inner_mask=True):
            add(T_QUOTE, stripped[start:end], start, end, key=_quote_norm(stripped[start:end]))
            _mask(buf, start, end)
    # 5. 层次序号和列举序号（不算事实，只屏蔽）
    m = SEQ_LINE_START_RE.match(current())
    if m:
        _mask(buf, m.start(), m.end())
    for regex in SEQ_INLINE_RES:
        for start, end, _m in _scan_spans(regex, current()):
            _mask(buf, start, end)
    # 6. 日期、时刻（"16号"与"16日"视为相同）
    for regex in DATE_RES:
        for start, end, m in _scan_spans(regex, current()):
            value = m.group(0)
            if value.endswith("号"):
                value = value[:-1] + "日"
            value = LEADING_ZERO_RE.sub("", value)  # "11月05日"与"11月5日"视为相同
            add(T_DATE, value, start, end)
            _mask(buf, start, end)
    for regex in CN_DATE_RES:
        for start, end, m in _scan_spans(regex, current()):
            year, month, day = m.group(1), m.group(2), m.group(3)
            y = cn_to_int(year) if year else None
            mo = cn_to_int(month)
            d = cn_to_int(day) if day else None
            if mo is None or not 1 <= mo <= 12 or (year and y is None) or (day and (d is None or not 1 <= d <= 31)):
                continue
            value = (f"{y}年" if year else "") + f"{mo}月" + (f"{d}日" if day else "")
            add(T_DATE, value, start, end)
            _mask(buf, start, end)
    for start, end, m in _scan_spans(TIME_RE, current()):
        add(T_DATE, f"{int(m.group(1))}:{m.group(2)}", start, end)
        _mask(buf, start, end)
    # 7. "百分之三十""百分之百"
    for start, end, m in _scan_spans(PERCENT_CN_RE, current()):
        value = 100 if m.group(1) == "百" else cn_to_int(m.group(1))
        if value is not None:
            add(T_NUMBER, f"{value}%", start, end)
            _mask(buf, start, end)
    # 7.1 "三分之一""千分之五"：按"分子/分母"比对，千分之按 ‰
    for start, end, m in _scan_spans(FRACTION_CN_RE, current()):
        denom = _FRACTION_UNIT_DENOM.get(m.group(1)) or cn_to_int(m.group(1))
        numer = cn_to_int(m.group(2))
        if not denom or numer is None:
            continue
        add(T_NUMBER, f"{numer}‰" if denom == 1000 else f"{numer}/{denom}", start, end)
        _mask(buf, start, end)
    # 8. 阿拉伯数字
    text = current()
    numbers = []  # [match, 倍数（万、亿）, 单位, 结束位置]
    for m in ARABIC_RE.finditer(text):
        if MASK in m.group(0):
            continue
        unit = _match_unit(text, m.end())
        numbers.append([m, m.group(5) or "", unit, m.end() + len(unit)])
    # 数值范围：前一个数缺倍数或单位的，沿用后一个数的（只改写法、数值不变的不报差异）
    for cur, nxt in zip(numbers, numbers[1:]):
        if not cur[2] and cur[3] < nxt[0].start() and RANGE_SEP_RE.fullmatch(text, cur[3], nxt[0].start()):
            cur[1] = cur[1] or nxt[1]
            cur[2] = nxt[2]
    for m, mult, unit, end in numbers:
        ordinal, int_part, frac, approx1, _mult, approx2 = m.groups()
        unit = unit.replace(" ", "")
        value = (ordinal or "") + _norm_number(int_part, frac) + (approx1 or "") + mult + (approx2 or "") + unit
        plain = None if (frac or approx1 or approx2 or mult) else int(int_part.replace(",", ""))
        hint = _number_hint(norm, m.start(), end, unit, plain, bool(ordinal))
        add(T_NUMBER, value, m.start(), end, key=_number_norm(value), hint=hint)
        _mask(buf, m.start(), end)
    # 9. 带量词的汉字数字（先屏蔽成语）
    for idiom in CN_IDIOMS:
        for m in re.finditer(re.escape(idiom), current()):
            _mask(buf, m.start(), m.end())
    # 9.1 汉字小数：须带万、亿或单位，"一点五亿"按"1.5亿"比对
    text = current()
    for start, end, m in _scan_spans(CN_DECIMAL_RE, text):
        whole = cn_to_int(m.group(1))
        if whole is None:
            continue
        unit = _match_unit(text, end)
        if unit and LATIN_UNIT_RE.fullmatch(unit):
            unit = ""
        if not (m.group(3) or unit):
            continue
        frac = "".join(str(_CN_DIGIT_VALUE[ch]) for ch in m.group(2)).rstrip("0")
        unit = unit.replace(" ", "")
        value = str(whole) + (f".{frac}" if frac else "") + (m.group(3) or "") + unit
        add(T_NUMBER, value, start, end + len(unit), key=_number_norm(value))
        _mask(buf, start, end + len(unit))
    text = current()
    for m in CN_RE.finditer(text):
        ordinal, digits, approx1, mult, approx2 = m.groups()
        number = cn_to_int(digits)
        if number is None:
            continue
        unit = _match_unit(text, m.end())
        if unit and LATIN_UNIT_RE.fullmatch(unit) and not ordinal:
            unit = ""
        if not ordinal and m.start() > 0 and text[m.start() - 1] in CN_VAGUE_PREFIXES:
            # "数十万人次""几十家"是概数：原样列入第三张清单；"数十""几百"不带量词、不含万亿的不算
            if not unit and not mult:
                continue
            start = m.start() - 1
            unit = unit.replace(" ", "")
            end = m.end() + len(unit)
            value = text[start] + digits + (approx1 or "") + (mult or "") + (approx2 or "") + unit
            add(T_NUMBER, value, start, end, key=_number_norm(value), hint=HINT_APPROX)
            _mask(buf, start, end)
            continue
        if not ordinal:
            if not unit and not mult:
                continue  # 不带量词的汉字数字（统一、一是、十四五）不算；含万、亿的（三百二十万）照常比对
            if digits == "一" and not (approx1 or mult or approx2):
                continue  # "一个""一批""一次"多为泛指
        unit = unit.replace(" ", "")
        value = (ordinal or "") + str(number) + (approx1 or "") + (mult or "") + (approx2 or "") + unit
        end = m.end() + len(unit)
        plain = None if (approx1 or mult or approx2) else number
        hint = _number_hint(norm, m.start(), end, unit, plain, bool(ordinal))
        add(T_NUMBER, value, m.start(), end, key=_number_norm(value), hint=hint)
        _mask(buf, m.start(), end)
    facts.sort(key=lambda f: f.pos)
    return facts


def extract_facts(lines: Sequence[str], source: str = "") -> List[Fact]:
    facts: List[Fact] = []
    for index, line in enumerate(lines, start=1):
        if line.strip():
            facts.extend(extract_line(line, index))
    if source:
        for fact in facts:
            fact.source = source
    return facts


# ---------------------------------------------------------------------------
# 比对
# ---------------------------------------------------------------------------

def _entry(items: List[Fact], original_count: int, revised_count: int, with_source: bool) -> dict:
    """生成一项差异。items 为该项在一侧的全部出现位置；可能是序数或概数的带 hint 字段。"""
    first = items[0]
    locations = []
    for f in items[:MAX_LOCATIONS]:
        loc = {"line": f.line, "excerpt": f.excerpt}
        if with_source and f.source:
            loc["file"] = f.source
        locations.append(loc)
    entry = {
        "type": first.kind,
        "value": first.value,
        "original_count": original_count,
        "revised_count": revised_count,
        "locations": locations,
    }
    if with_source:
        entry["sources"] = list(OrderedDict.fromkeys(f.source for f in items if f.source))
    if first.hint:  # 比对键含 hint，同一项的各处 hint 相同
        entry["hint"] = first.hint
    if first.hint == HINT_SUMMARY:
        entry["level"] = SUMMARY_LEVEL
        entry["note"] = SUMMARY_NOTE
    return entry


def _diff_entries(base: List[Fact], base_counts: Counter, other_counts: Counter,
                  base_is_original: bool, with_source: bool = False) -> List[dict]:
    """列出 base 中出现次数多于另一稿的项（按首次出现顺序），位置取自 base。"""
    grouped: "OrderedDict[Tuple[str, str, str], List[Fact]]" = OrderedDict()
    for fact in base:
        grouped.setdefault(fact.key, []).append(fact)
    entries: List[dict] = []
    for key, items in grouped.items():
        mine, theirs = base_counts[key], other_counts.get(key, 0)
        if mine <= theirs:
            continue
        if base_is_original:
            entries.append(_entry(items, mine, theirs, with_source))
        else:
            entries.append(_entry(items, theirs, mine, False))
    return entries


def compare(original: List[Fact], revised: List[Fact], original_counts: Optional[Counter] = None,
            with_source: Optional[bool] = None) -> Tuple[List[dict], List[dict]]:
    """返回 (原稿有而改稿没有, 改稿新增)。按出现次数比对。

    original_counts：多份原稿时传入各稿计数取最大值后的结果；不传时按 original 计数。
    with_source：是否在"原稿有而改稿没有"中注明出自哪份原稿；不传时按 original 中的文件数判断。
    可能是序数或概数的项带 hint 字段，由 split_uncertain 分出。
    """
    orig_counts = original_counts if original_counts is not None else Counter(f.key for f in original)
    rev_counts = Counter(f.key for f in revised)
    if with_source is None:
        with_source = len({f.source for f in original}) > 1
    missing = _diff_entries(original, orig_counts, rev_counts, base_is_original=True, with_source=with_source)
    added = _diff_entries(revised, rev_counts, orig_counts, base_is_original=False)
    return missing, added


def compare_ledger(ledger: List[Fact], body: List[Fact]) -> List[dict]:
    """台账对账：返回正文中有、台账中没有的事实。按集合比较，同一事实在正文中用几次都只看有没有。"""
    known = {f.key for f in ledger}
    body_counts = Counter(f.key for f in body)
    grouped: "OrderedDict[Tuple[str, str, str], List[Fact]]" = OrderedDict()
    for fact in body:
        if fact.key not in known:
            grouped.setdefault(fact.key, []).append(fact)
    return [_entry(items, 0, body_counts[key], False) for key, items in grouped.items()]


def split_uncertain(entries: List[dict], side: str) -> Tuple[List[dict], List[dict]]:
    """把可能是序数或概数的项分出来，返回 (确定的项, 可能是序数或概数的项)。side 记入后者的 list 字段。"""
    certain = [e for e in entries if "hint" not in e]
    uncertain = [dict(e, list=side) for e in entries if "hint" in e]
    return certain, uncertain


def _type_counts(facts: List[Fact]) -> Dict[str, int]:
    counts = Counter(f.kind for f in facts)
    return {kind: counts.get(kind, 0) for kind in TYPE_ORDER}


def _display_names(paths: Sequence[Path]) -> List[str]:
    names = [p.name for p in paths]
    if len(set(names)) != len(names):  # 不同目录下的同名文件，改用完整路径区分
        names = [str(p) for p in paths]
    return names


def _revision_note(name: str, revisions: int) -> str:
    return f"{name} 含修订痕迹 {revisions} 处，已按接受修订后的文字比对"


def build_report(original_path: Path, revised_path: Path, ledger: bool = False,
                 refs: Optional[Sequence[Path]] = None) -> dict:
    """refs：追加的原稿（台账模式下为追加的台账或材料），与 original_path 合起来作比对基准。"""
    paths = [original_path] + list(refs or [])
    names = _display_names(paths)
    notes: List[str] = []
    orig_facts: List[Fact] = []
    orig_counts: Counter = Counter()
    orig_line_total = 0
    for path, name in zip(paths, names):
        text, revisions = read_document(path)
        if revisions:
            notes.append(_revision_note(name, revisions))
        lines = body_lines(text)
        if ledger:
            lines = ledger_lines(lines)
        facts = extract_facts(lines, source=name)
        orig_facts += facts
        orig_counts |= Counter(f.key for f in facts)  # 多份原稿：同一事实按出现次数最多的一份计
        orig_line_total += len(lines)
    text, revisions = read_document(revised_path)
    if revisions:
        notes.append(_revision_note(revised_path.name, revisions))
    rev_lines = body_lines(text)
    rev_facts = extract_facts(rev_lines)
    report = {
        "original_file": names[0],
        "original_files": names,
        "revised_file": revised_path.name,
        "scope": "【起草说明】之前的正文",
        "notes": notes,
        "original_lines": orig_line_total,
        "revised_lines": len(rev_lines),
        "fact_counts": {"original": _type_counts(orig_facts), "revised": _type_counts(rev_facts)},
    }
    if ledger:
        added, uncertain = split_uncertain(compare_ledger(orig_facts, rev_facts), "added")
        report = dict({"mode": "ledger"}, **report)
        report["scope"] = "【起草说明】之前的正文；台账表格的“序号”“来源”列不参与比对"
        report.update({"added": added, "uncertain": uncertain, "consistent": not added})
        return report
    missing, added = compare(orig_facts, rev_facts, orig_counts, with_source=len(paths) > 1)
    missing, uncertain_missing = split_uncertain(missing, "missing")
    added, uncertain_added = split_uncertain(added, "added")
    report.update({
        "missing": missing,
        "added": added,
        "uncertain": uncertain_missing + uncertain_added,
        "consistent": not missing and not added,
    })
    return report


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------

UNCERTAIN_TITLE = "可能是序数或概数（请人工判断，不计入 --strict）"
UNCERTAIN_SIDES = {"missing": "原稿有而改稿没有", "added": "改稿新增"}


def _short(value: str) -> str:
    if len(value) <= VALUE_DISPLAY_LIMIT:
        return value
    return value[:VALUE_DISPLAY_LIMIT] + "……"


def _render_item(index: int, item: dict, where: str, names: Tuple[str, str], extra: str = "") -> List[str]:
    sources = item.get("sources") or []
    tail = f"；出自 {'、'.join(sources)}" if sources else ""
    lines = [f"{index}. [{item['type']}] {_short(item['value'])}"
             f"（{names[0]} {item['original_count']} 处，{names[1]} {item['revised_count']} 处{tail}{extra}）"]
    for loc in item["locations"]:
        place = f"{loc['file']} " if loc.get("file") else where
        lines.append(f"   {place}第 {loc['line']} 行：{loc['excerpt']}")
    return lines


def _render_list(title: str, entries: List[dict], where: str,
                 names: Tuple[str, str] = ("原稿", "改稿")) -> List[str]:
    lines = [f"{title}（{len(entries)} 项）"]
    if not entries:
        lines.append("无")
        return lines
    for index, item in enumerate(entries, start=1):
        lines += _render_item(index, item, where, names)
    return lines


def _render_uncertain(title: str, entries: List[dict], names: Tuple[str, str],
                      sides: Dict[str, str], wheres: Dict[str, str]) -> List[str]:
    lines = [f"{title}（{len(entries)} 项）"]
    for index, item in enumerate(entries, start=1):
        side = item.get("list", "added")
        if item.get("hint") == HINT_SUMMARY:
            extra = f"；{sides.get(side, side)}；{item.get('level', SUMMARY_LEVEL)} 提示：{item.get('note', SUMMARY_NOTE)}"
        else:
            extra = f"；{sides.get(side, side)}；可能是{item.get('hint', '序数或概数')}"
        lines += _render_item(index, item, wheres.get(side, ""), names, extra)
    return lines


def _uncertain_conclusion(report: dict) -> List[str]:
    count = len(report.get("uncertain") or [])
    if not count:
        return []
    lines = [f"另有 {count} 项可能是序数或概数（如“三点意见”“九成”），未计入差异和 --strict，"
             "请人工判断：属于事实数据的（如“9成患者满意”），按前面的清单同样处理。"]
    summary = sum(1 for item in report["uncertain"] if item.get("hint") == HINT_SUMMARY)
    if summary:
        lines.append(f"其中 {summary} 项为{SUMMARY_NOTE}（{SUMMARY_LEVEL} 提示，如“具体做法有三条”）。")
    return lines


def _file_label(report: dict) -> str:
    files = report.get("original_files") or [report["original_file"]]
    return "、".join(files) + (f"（共 {len(files)} 份）" if len(files) > 1 else "")


def render_ledger_text(report: dict) -> str:
    counts_o = report["fact_counts"]["original"]
    counts_r = report["fact_counts"]["revised"]
    summary = "，".join(f"{k} {counts_o[k]}/{counts_r[k]}" for k in TYPE_ORDER)
    lines = [
        f"台账对账：台账 {_file_label(report)}，正文 {report['revised_file']}",
        f"比对范围：{report['scope']}（台账 {report['original_lines']} 行，正文 {report['revised_lines']} 行）",
        f"识别到的事实项（台账/正文）：{summary}",
    ]
    lines += [f"说明：{note}。" for note in report.get("notes") or []]
    lines.append("")
    lines += _render_list("正文中有、台账中没有", report["added"], "正文", ("台账", "正文"))
    lines.append("")
    if report.get("uncertain"):
        lines += _render_uncertain(UNCERTAIN_TITLE, report["uncertain"], ("台账", "正文"),
                                   {"added": "正文中有、台账中没有"}, {"added": "正文"})
        lines.append("")
    if report["added"]:
        lines.append(f"结论：{len(report['added'])} 项不在台账中。逐项处理：补进台账并写明出处，"
                     "或在正文中改为【待补：说明】；新加的【待补】【待核实】标记核对无误即可。")
    else:
        lines.append("结论：正文中的数字、日期、标记等均能在台账中找到。")
    lines += _uncertain_conclusion(report)
    return "\n".join(lines)


def render_text(report: dict) -> str:
    if report.get("mode") == "ledger":
        return render_ledger_text(report)
    counts_o = report["fact_counts"]["original"]
    counts_r = report["fact_counts"]["revised"]
    summary = "，".join(f"{k} {counts_o[k]}/{counts_r[k]}" for k in TYPE_ORDER)
    lines = [
        f"事实一致性检查：原稿 {_file_label(report)}，改稿 {report['revised_file']}",
        f"比对范围：{report['scope']}（原稿 {report['original_lines']} 行，改稿 {report['revised_lines']} 行）",
        f"识别到的事实项（原稿/改稿）：{summary}",
    ]
    lines += [f"说明：{note}。" for note in report.get("notes") or []]
    if len(report.get("original_files") or []) > 1:
        lines.append("说明：多份原稿合并作比对基准，同一事实按出现次数最多的一份计数。")
    lines.append("")
    lines += _render_list("一、原稿有而改稿没有", report["missing"], "原稿")
    lines.append("")
    lines += _render_list("二、改稿新增", report["added"], "改稿")
    lines.append("")
    if report.get("uncertain"):
        lines += _render_uncertain("三、" + UNCERTAIN_TITLE, report["uncertain"], ("原稿", "改稿"),
                                   UNCERTAIN_SIDES, {"missing": "原稿", "added": "改稿"})
        lines.append("")
    total = len(report["missing"]) + len(report["added"])
    if total:
        lines.append(f"结论：发现差异 {total} 项。逐项核对：属于事实改动的，恢复原稿写法；"
                     "确属表达调整的（如去掉强调用的引号、删去复述句后重复的数字少了一处、按保密要求脱敏），"
                     "在修改说明中写明。")
    else:
        lines.append("结论：未发现事实差异。人名、单位名、承诺事项不在本脚本比对范围内，请人工核对。")
    lines += _uncertain_conclusion(report)
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
        prog="fact_diff.py",
        description="比对改稿前后的数字、日期、【待补】【待核实】标记、书名号和引号内容、发文字号，"
                    "确认改稿没有改动事实。只比对【起草说明】之前的正文。",
        epilog="退出码：0 正常；1 使用 --strict 且前两张清单有差异（“可能是序数或概数”一张不计）；2 输入错误。",
        formatter_class=_ChineseHelpFormatter,
        add_help=False,
    )
    parser._optionals.title = "选项"
    parser._positionals.title = "参数"
    parser.add_argument("-h", "--help", action="help", help="显示本帮助信息并退出")
    parser.add_argument("original", help="原稿文件（.txt、.md、.docx）")
    parser.add_argument("revised", help="改稿文件（.txt、.md、.docx）")
    parser.add_argument("--json", action="store_true", help="以 JSON 格式输出")
    parser.add_argument("--strict", action="store_true",
                        help="“原稿有而改稿没有”“改稿新增”任一非空时返回 1（“可能是序数或概数”不计）")
    parser.add_argument("--ledger", action="store_true",
                        help="台账对账：第一个文件为素材台账、第二个为正文，只列正文中有、台账中没有的事实；"
                             "按是否出现比较，不计次数；台账表格的“序号”“来源”列不参与比对")
    parser.add_argument("--ref", action="append", default=[], metavar="文件",
                        help="追加一份原稿（台账模式下为追加的台账或材料），可重复使用；各稿分别截掉【起草说明】"
                             "及其后内容，与第一个文件合起来作比对基准，同一事实按出现次数最多的一份计数，"
                             "“原稿有而改稿没有”的每项注明出自哪份原稿")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    _setup_console()
    parser = make_parser()
    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:
        return 0 if exc.code == 0 else 2
    try:
        report = build_report(Path(args.original), Path(args.revised), ledger=args.ledger,
                              refs=[Path(p) for p in args.ref])
    except InputError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(render_text(report))
    if args.strict and not report["consistent"]:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
