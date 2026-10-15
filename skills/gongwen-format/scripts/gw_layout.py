# -*- coding: utf-8 -*-
"""公文正文结构识别（只依赖标准库）。

从按行排列的纯文本中识别版头要素、标题、主送机关、层次标题、正文、附件说明、
署名、成文日期、附注、版记等角色。build_docx.py、check_docx.py、fix_docx.py 共用。

识别规则：
1. 首个非空行以"关于"开头或以文种结尾时作标题；标题回行的，合并后续行。
2. 紧随标题、以全角冒号结尾的短行作主送机关。
3. 以"附件："开头的行及其后的"2. ××"行作附件说明。
4. 从末尾倒推：印发行、抄送行、附注、成文日期、署名。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

# ---------------------------------------------------------------------------
# 文种与模式
# ---------------------------------------------------------------------------

# 法定文种（《党政机关公文处理工作条例》第八条）
LEGAL_TYPES = ("决议", "决定", "命令", "公报", "公告", "通告", "意见", "通知",
               "通报", "报告", "请示", "批复", "议案", "函", "纪要")
INTERNAL_TYPES = ("签报",)
# 常用材料；"调研报告"等虽以"报告"结尾，按材料处理，须先于法定文种匹配
MATERIAL_TYPES = ("调研报告", "述职报告", "述责述廉报告", "述职述廉报告", "工作报告",
                  "汇报材料", "经验材料", "工作要点", "实施方案", "工作方案", "讲话稿",
                  "总结", "计划", "要点", "方案", "讲话", "发言", "致辞", "信息", "简报",
                  "规划", "办法", "规定", "细则", "预案", "汇报")
UPWARD_TYPES = ("请示", "报告", "议案")  # 上行文，红头版式需要签发人

CN_NUM = "一二三四五六七八九十百零〇"
H1_RE = re.compile(rf"^[{CN_NUM}]+、")
H2_RE = re.compile(rf"^（[{CN_NUM}]+）")
H3_RE = re.compile(r"^\d{1,2}[.．](?!\d)")
H4_RE = re.compile(r"^（\d{1,2}）")
# 不规范的层次序数写法（仍按对应层级识别，检查时报出）
H2_BAD_RE = re.compile(rf"^\([{CN_NUM}]+\)")
H4_BAD_RE = re.compile(r"^\(\d{1,2}\)")
H3_BAD_RE = re.compile(r"^\d{1,2}、")

DATE_AR_RE = re.compile(r"^(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日$")
DATE_CN_RE = re.compile(r"^[〇○零一二三四五六七八九]{4}年[一二三四五六七八九十]{1,2}月[一二三四五六七八九十]{1,3}日$")
DATE_MARK_RE = re.compile(r"^【待补[：:][^】]*日期[^】]*】$")
DOCNO_PAT = r"[\u4e00-\u9fffA-Za-z]{1,15}\s*[〔\[［(（【﹝{｛]\s*(?:19|20)?\d{2,4}\s*[〕\]］)）】﹞}｝]\s*第?\s*\d+\s*号"
DOCNO_RE = re.compile(DOCNO_PAT)
# 发文字号行：可带份号（信函格式同一行）和签发人（上行文）
DOCNO_LINE_RE = re.compile(rf"^(?:(?P<copy>\d{{1,6}})\s+)?(?P<docno>{DOCNO_PAT})\s*(?:签发人\s*[：:]\s*(?P<signer>.+))?$")
# 排版时写入的占位标记（修复、检查时要能认出来）
DOCNO_MARK_RE = re.compile(r"^(?P<docno>【待补[：:][^】]*(?:发文字号|文号|序号)[^】]*】)\s*(?:签发人\s*[：:]\s*(?P<signer>.+))?$")
ISSUER_MARK_RE = re.compile(r"^【待补[：:][^】]*(?:发文机关标志|纪要标志)[^】]*】$")
QIHAO_RE = re.compile(r"^(?:[\u4e00-\u9fff]{0,15}\s*[〔\[［(（]\s*\d{4}\s*[〕\]］)）]\s*)?第\s*\d+\s*期$")
SIGNER_RE = re.compile(r"签发人\s*[：:]\s*(.+)$")
COPY_NO_RE = re.compile(r"^\d{1,6}$")
SECRET_RE = re.compile(r"^(绝密|机密|秘密)(\s*★.*)?$")
URGENCY_RE = re.compile(r"^(特急|加急|特提|平急)$")
# 印发机关以非空白字符结尾（贪婪匹配），避免“懒惰匹配 + 空白串”在超长空白行上二次方回溯
PRINT_RE = re.compile(r"^(?P<org>\S(?:.*\S)?)\s+(?P<date>\d{4}年\d{1,2}月\d{1,2}日)\s*印发$")
CC_RE = re.compile(r"^抄\s*送\s*[：:]")
BANJI_MAIN_RE = re.compile(r"^主\s*送\s*[：:]")
ATTACH_RE = re.compile(r"^附\s*件\s*[：:]")
ATTACH_CONT_RE = re.compile(r"^\d{1,2}\s*[.．]\s*\S")
NOTE_RE = re.compile(r"^[（(].+[）)]$")
ATTENDANCE_RE = re.compile(r"^(出\s*席|请\s*假|列\s*席|缺\s*席)\s*[：:]")
SUBTITLE_RE = re.compile(r"^(——|—|－－|--)")
TITLE_NOTE_RE = re.compile(r"^[（(]\s*\d{4}年\d{1,2}月\d{1,2}日\s*[）)]$")
CLOSING_RE = re.compile(r"^(特此|此复|此致|敬礼|妥否|以上|当否|专此)")
SENT_PUNCT = "。，；！？、"
MARKERS_RE = re.compile(r"【(待补|待核实)[：:][^】]*】")
SECRET_WORDS_RE = re.compile(r"(绝密|机密|秘密)\s*★|★\s*\d+\s*年|内部资料|工作秘密|不得外传|涉密")
DRAFT_NOTE_LINE = "【起草说明】"

ROLE_LABELS = {
    "copy_no": "份号", "secret": "密级和保密期限", "urgency": "紧急程度",
    "issuer": "发文机关标志", "docno": "发文字号", "signer": "签发人",
    "title": "标题", "subtitle": "副标题", "title_note": "题注",
    "recipient": "主送机关", "h1": "一级标题", "h2": "二级标题", "h3": "三级标题",
    "h4": "四级标题", "body": "正文", "attendance": "出席列席", "attachment": "附件说明",
    "signature": "发文机关署名", "date": "成文日期", "note": "附注",
    "banji_main": "版记主送", "cc": "抄送", "print": "印发机关和印发日期", "blank": "空行",
}
HEADING_ROLES = ("h1", "h2", "h3", "h4")
BODY_ROLES = ("body", "h1", "h2", "h3", "h4", "attendance")


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------

def is_draft_note_line(line: str) -> bool:
    """是否【起草说明】分界行。

    与 gongwen_lint.py 一致：去掉行首空白（含全角空格）后以【起草说明】开头即算；
    另外兼容 Markdown 标题、加粗包裹的写法（“## 【起草说明】”“**【起草说明】**”），避免起草说明被排进正文。
    """
    text = line.strip()
    if text.startswith(DRAFT_NOTE_LINE):
        return True
    text = re.sub(r"^#{1,6}\s*", "", text)
    text = re.sub(r"^(?:\*\*|__)\s*", "", text)
    return text.startswith(DRAFT_NOTE_LINE)


def draft_note_index(texts: Sequence[str]) -> Optional[int]:
    """返回第一个【起草说明】分界行的下标；没有时返回 None。读取 docx 段落时用。"""
    for i, text in enumerate(texts):
        if is_draft_note_line(text or ""):
            return i
    return None


def read_body_text(path: Path) -> Tuple[List[str], List[str]]:
    """读取 .txt/.md 正文，遇到【起草说明】停止。返回 (行列表, 提示)。

    先按 UTF-8（可带 BOM）解码，失败时按 GB18030 解码（与 gongwen_lint.py 一致）；
    都失败时抛出 ValueError，由调用方提示用户另存为 UTF-8。
    """
    notes: List[str] = []
    data = path.read_bytes()
    try:
        raw = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        try:
            raw = data.decode("gb18030")
            notes.append("文件不是 UTF-8 编码，已按 GB18030 读取。")
        except UnicodeDecodeError as exc:
            raise ValueError("无法识别文件编码，请另存为 UTF-8 后重试。") from exc
    lines: List[str] = []
    md_hit = False
    for line in raw.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if is_draft_note_line(line):
            break
        if path.suffix.lower() in (".md", ".markdown"):
            new = re.sub(r"^\s{0,3}#{1,6}\s+", "", line)
            new = re.sub(r"(\*\*|__)(.+?)\1", r"\2", new)
            if new != line:
                md_hit = True
            line = new
        lines.append(line)
    if md_hit:
        notes.append("输入含 Markdown 标记（# 或 **），已去掉标记符号，文字未改动")
    return lines, notes


def clean_line(line: str) -> str:
    """去掉行首行尾空白（含全角空格），行首空格缩进由排版统一处理。"""
    return line.strip().strip("　").strip()


FULL_WIDTH_PUNCT = "—…“”‘’"  # 码位低于 0x2E80，但按 eastAsia 字体排时占一个字位置（GB/T 15834）


def text_width_chars(text: str) -> float:
    """按字估算宽度：汉字和全角符号（含中文引号、破折号、省略号）记 1 字，ASCII 字符记 0.5 字。"""
    width = 0.0
    for ch in text:
        width += 0.5 if ord(ch) < 0x2E80 and ch not in FULL_WIDTH_PUNCT else 1.0
    return width


def is_heading(text: str) -> Optional[str]:
    """返回层次标题级别（h1～h4），不是标题时返回 None。

    一级标题独占一行、末尾不加标点；"一、二季度共完成……。"这类带句号的长句按正文处理。
    """
    if H1_RE.match(text):
        return "h1" if "。" not in text and len(text) <= 50 else None
    if H2_RE.match(text) or H2_BAD_RE.match(text):
        return "h2"
    if H4_RE.match(text) or H4_BAD_RE.match(text):
        return "h4"
    if H3_RE.match(text) or H3_BAD_RE.match(text):
        return "h3"
    return None


def heading_split(text: str) -> int:
    """标题与正文同段时，返回标题部分的结束位置（含句号）；整段是标题时返回全长。"""
    idx = text.find("。")
    if idx == -1 or idx >= len(text) - 1:
        return len(text)
    return idx + 1


# 标题回行时优先在这些词之后断开（词意完整）
TITLE_BREAK_AFTER = ("关于印发", "关于转发", "关于", "印发", "转发", "开展", "做好", "加强", "进一步", "委员会",
                     "办公室", "办公厅", "公司", "集团", "政府", "中心", "大学", "学院", "医院", "学校", "年度",
                     "局", "部", "院", "委", "厅", "办", "处", "年", "的", "》", "）")
TITLE_NO_LINE_START = "的了和与及、，。；：）》"
# 公文标题常用词：回行时不从词中间断开（贪心最长匹配分词，只用于找断点）
TITLE_WORDS = set("""
基本 医疗 保险 参保 缴费 工作 安全 生产 管理 建设 发展 服务 改革 教育 卫生 健康 社会 经济 质量 实施 方案 意见
通知 规定 办法 计划 总结 报告 请示 检查 治理 整治 推进 落实 加强 做好 开展 进一步 全面 深入 城乡 居民 人力 资源
集团 有限 公司 有限公司 委员会 党风 廉政 反腐败 防控 应急 防汛 消防 环境 保护 生态 文明 创建 考核 评估 审核 监督
执法 专项 行动 活动 会议 培训 招生 预算 资金 项目 申报 认定 试点 标准 规范 制度 机制 体系 能力 水平 队伍 人才
科技 创新 数字 信息 平台 系统 基层 减负 党建 组织 宣传 纪律 作风 年度 冬季 夏季 春季 秋季 全市 全区 全县 全省
全国 学生 宿舍 医院 学校 大学 学院 部门 单位 企业 职工 干部 群众 研究生 硕士 本科 教学 门诊 住院 病历 患者 护理
政务 营商 环境 乡村 振兴 农村 社区 街道 交通 运输 道路 住房 城市 更新 改造 老旧 小区 养老 托育 就业 创业 财政
税务 审计 统计 市场 监管 食品 药品 文化 旅游 体育 民政 残疾人 退役 军人 水利 自然 林业 粮食 能源 电力 燃气 供水
工程 设备 采购 招标 合同 绩效 薪酬 任期制 契约化 经理层 安全生产 大检查 实施方案 工作方案 工作要点 年终 情况
冬春季 冬春 传染病 卫生健康委 住宿费 收费 医保 慢特病 联合 审核
""".split())


def _segment_bounds(text: str) -> set:
    """贪心最长匹配分词，返回落在词中间的位置集合。"""
    inside = set()
    i = 0
    while i < len(text):
        for size in (5, 4, 3, 2):
            if text[i:i + size] in TITLE_WORDS:
                inside.update(range(i + 1, i + size))
                i += size
                break
        else:
            i += 1
    return inside


def _protected(text: str):
    """返回 (不能断开的位置, 括号内部的位置)：数字连写和“2026年”中间不能断；书名号、括号内部尽量不断。"""
    hard = [False] * (len(text) + 1)
    soft = [False] * (len(text) + 1)
    depth = 0
    for i, ch in enumerate(text):
        if depth > 0:
            soft[i] = True
        if ch in "《（(":
            depth += 1
        elif ch in "》）)" and depth > 0:
            depth -= 1
        prev = text[i - 1] if i > 0 else ""
        if prev and prev.isascii() and prev.isalnum() and ch.isascii() and ch.isalnum():
            hard[i] = True
        if prev.isdigit() and ch in "年月日号期次届":
            hard[i] = True
    return hard, soft


def break_title(text: str, max_chars: float = 20.0, min_chars: float = 3.0) -> List[str]:
    """把超长标题分成几行：每行不超过 max_chars，上行不短于下行，尽量在词边界断开（7.3.1）。

    宽度按汉字 1 字、半角字符 0.5 字估算。放得下一行时原样返回。
    """
    total = text_width_chars(text)
    if total <= max_chars or len(text) < 4:
        return [text]
    hard, soft = _protected(text)
    inside_word = _segment_bounds(text)
    prefix = [0.0]
    for ch in text:
        prefix.append(prefix[-1] + text_width_chars(ch))

    def boundary_cost(i: int) -> Optional[float]:
        if i <= 0 or i >= len(text) or hard[i]:
            return None
        if any(text[:i].endswith(w) for w in TITLE_BREAK_AFTER):
            cost = 0.0
        elif i in inside_word:
            cost = 30.0
        elif text[i - 1].isascii() != text[i].isascii():
            cost = 4.0
        else:
            cost = 6.0
        if soft[i]:
            cost += 15.0
        if text[i] in TITLE_NO_LINE_START:
            cost += 20.0
        return cost

    def score(cuts: Sequence[int]) -> Optional[float]:
        edges = [0] + list(cuts) + [len(text)]
        widths = [prefix[b] - prefix[a] for a, b in zip(edges, edges[1:])]
        if any(w > max_chars or w < min_chars for w in widths):
            return None
        cost = sum(boundary_cost(c) for c in cuts)
        cost += sum(max(0.0, b - a) * 3 for a, b in zip(widths, widths[1:]))  # 上行不短于下行
        cost += (max(widths) - min(widths)) * 0.5
        return cost

    candidates = [i for i in range(1, len(text)) if boundary_cost(i) is not None]
    lines = int(-(-total // max_chars))
    best = None
    for n in (lines, lines + 1):
        if n == 2:
            combos = ([c] for c in candidates)
        elif n == 3:
            combos = ([a, b] for a in candidates for b in candidates if b > a)
        else:
            break
        for cuts in combos:
            sc = score(cuts)
            if sc is None:
                continue
            sc += (n - lines) * 8  # 能少分一行就少分
            if best is None or sc < best[0]:
                best = (sc, cuts)
    if best:
        edges = [0] + best[1] + [len(text)]
        return [text[a:b] for a, b in zip(edges, edges[1:])]
    return [text]


def doc_kind(title: str) -> Tuple[str, Optional[str]]:
    """根据标题末尾判断文书类别。返回 (类别, 文种)，类别为"法定""内部""材料""未知"。"""
    t = re.sub(r"[\s　]", "", title or "")
    if not t:
        return "未知", None
    for name in MATERIAL_TYPES:
        if t.endswith(name):
            return "材料", name
    for name in LEGAL_TYPES:
        if t.endswith(name):
            return "法定", name
    for name in INTERNAL_TYPES:
        if t.endswith(name):
            return "内部", name
    return "未知", None


def ends_with_doc_type(text: str) -> bool:
    return doc_kind(text)[0] != "未知"


def has_sentence_punct(text: str) -> bool:
    return any(ch in text for ch in SENT_PUNCT)


def looks_like_date(text: str) -> bool:
    t = text.replace(" ", "").replace("　", "")
    return bool(DATE_AR_RE.match(t) or DATE_CN_RE.match(t) or DATE_MARK_RE.match(t))


def normalize_date(value: str) -> Tuple[str, bool]:
    """把 2026-10-16、2026/10/16、2026.10.16 写法转成"2026年10月16日"。返回 (结果, 是否转换)。

    已是"××××年×月×日"写法的原样保留，由 check_docx.py 检查虚位等问题。
    """
    v = (value or "").strip()
    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", v)
    if m:
        y, mo, d = m.groups()
        return f"{y}年{int(mo)}月{int(d)}日", True
    return v, False


# ---------------------------------------------------------------------------
# 识别结果
# ---------------------------------------------------------------------------

@dataclass
class Item:
    role: str
    text: str
    src: int = -1  # 在输入行列表中的下标；来自版式信息时为 -1
    split: int = -1  # 标题与正文同段时标题部分的结束位置
    first: bool = True  # 附件说明的第一行


@dataclass
class Parsed:
    items: List[Item] = field(default_factory=list)
    fields: Dict[str, object] = field(default_factory=dict)
    detected: List[str] = field(default_factory=list)  # 识别报告
    warnings: List[str] = field(default_factory=list)
    kind: str = "未知"
    doc_type: Optional[str] = None

    def roles(self) -> List[str]:
        return [it.role for it in self.items]

    def find(self, role: str) -> List[Item]:
        return [it for it in self.items if it.role == role]


# ---------------------------------------------------------------------------
# 识别主流程
# ---------------------------------------------------------------------------

def _is_signature_line(text: str) -> bool:
    text = MARKERS_RE.sub("某", text)
    if len(text) > 40 or has_sentence_punct(text) or "：" in text or ":" in text:
        return False
    if is_heading(text) or ATTACH_RE.match(text) or CLOSING_RE.match(text):
        return False
    return not looks_like_date(text)


def parse_lines(raw_lines: Sequence[str], red: Optional[Sequence[bool]] = None) -> Parsed:
    """识别各行角色。raw_lines 可含空行，空行不进入结果。

    red：与 raw_lines 对应的“整段为红色”标记（读已有 docx 时传入），开头的红色短行按发文机关标志处理。
    """
    res = Parsed()
    lines: List[Tuple[int, str]] = []
    indent_hit = False
    for idx, raw in enumerate(raw_lines):
        text = clean_line(raw)
        if text:
            if raw[:1] in (" ", "　", "\t"):
                indent_hit = True
            lines.append((idx, text))
    if indent_hit:
        res.warnings.append("部分段落用空格缩进，排版时已去掉行首空格，缩进统一按首行缩进 2 字处理")
    if not lines:
        res.warnings.append("正文为空")
        return res

    n = len(lines)
    head: List[Item] = []
    pos = 0

    # 1. 版头要素（正文文件里一般没有；修复已有 docx 时会遇到）
    while pos < n:
        idx, text = lines[pos]
        if COPY_NO_RE.match(text) and pos < 3:
            head.append(Item("copy_no", text, idx))
            res.fields["份号"] = text
        elif SECRET_RE.match(text):
            head.append(Item("secret", text, idx))
            res.fields["密级"] = text
        elif URGENCY_RE.match(text):
            head.append(Item("urgency", text, idx))
            res.fields["紧急程度"] = text
        elif _is_issuer(text, lines[pos + 1][1] if pos + 1 < n else "") or                 (red is not None and idx < len(red) and red[idx] and pos < 4 and len(text) <= 40):
            head.append(Item("issuer", text, idx))
            res.fields["发文机关标志"] = text
        elif DOCNO_LINE_RE.match(text) or QIHAO_RE.match(text) or DOCNO_MARK_RE.match(text):
            head.append(Item("docno", text, idx))
            _record_docno(res, text)
        elif SIGNER_RE.match(text):
            head.append(Item("signer", text, idx))
            res.fields["签发人"] = SIGNER_RE.match(text).group(1).strip()
        else:
            break
        pos += 1

    # 2. 标题（可能回行）
    title_items: List[Item] = []
    if pos < n:
        idx, text = lines[pos]
        if (text.startswith("关于") or ends_with_doc_type(text)) and not is_heading(text) \
                and not text.endswith(("：", ":")) and len(text) <= 80:
            title_items.append(Item("title", text, idx))
            pos += 1
            # 回行标题：首行未以文种结尾，后续短行以文种结尾时合并
            if not ends_with_doc_type(text):
                look = pos
                while look < n and look < pos + 2:
                    t2 = lines[look][1]
                    if has_sentence_punct(t2) or t2.endswith(("：", ":")) or is_heading(t2) or len(t2) > 40:
                        break
                    if ends_with_doc_type(t2):
                        for k in range(pos, look + 1):
                            title_items.append(Item("title", lines[k][1], lines[k][0]))
                        pos = look + 1
                        break
                    look += 1
            res.detected.append("标题：按首行" + ("以“关于”开头" if text.startswith("关于") else "以文种结尾") + "识别")
        elif _looks_title(text) and pos + 1 < n:
            title_items.append(Item("title", text, idx))
            pos += 1
            res.detected.append("标题：首行较短且无句末标点，推断为标题，请核对")
            res.warnings.append(f"标题“{text}”是按首行推断的，不是按“关于”或文种识别，请核对")
        else:
            res.warnings.append("未识别到标题：首行既不以“关于”开头，也不以文种结尾。如首行是标题，请在版式信息中填写“标题”")
    if title_items:
        title = "".join(it.text for it in title_items)
        res.fields["标题"] = title
        res.kind, res.doc_type = doc_kind(title)
    # 副标题、题注（讲话稿的日期行、讲话人行）
    while pos < n and title_items:
        idx, text = lines[pos]
        nxt = lines[pos + 1][1] if pos + 1 < n else ""
        if DOCNO_LINE_RE.match(text) or DOCNO_MARK_RE.match(text):
            # AI 稿常把发文字号写在标题下、主送机关前：认作发文字号，排版时移到版头，正文中不再保留
            title_items.append(Item("docno", text, idx))
            if res.fields.get("发文字号"):
                res.warnings.append(f"版头已有发文字号“{res.fields['发文字号']}”，标题下另有文号行“{text}”，已以版头为准")
            else:
                _record_docno(res, text)
            res.detected.append("发文字号：标题下的文号行，排版时移到版头")
            pos += 1
        elif SUBTITLE_RE.match(text) and len(text) <= 60:
            title_items.append(Item("subtitle", text, idx))
            res.fields["副标题"] = text
            pos += 1
        elif TITLE_NOTE_RE.match(text):
            title_items.append(Item("title_note", text, idx))
            pos += 1
        elif len(text) <= 25 and not has_sentence_punct(text) and not is_heading(text)                 and not text.endswith(("：", ":")) and (nxt.endswith(("：", ":")) or TITLE_NOTE_RE.match(nxt)):
            title_items.append(Item("title_note", text, idx))
            pos += 1
        else:
            break

    # 3. 主送机关
    recipient: Optional[Item] = None
    if pos < n:
        idx, text = lines[pos]
        # “现将有关事项通告如下：”这类引出下文的句子是正文首段（通告、公告没有主送机关），不是主送机关
        if text.endswith(("：", ":")) and len(text) <= 80 and "。" not in text and "如下" not in text \
                and not is_heading(text) and not ATTENDANCE_RE.match(text) and not ATTACH_RE.match(text):
            recipient = Item("recipient", text, idx)
            res.fields["主送机关"] = text.rstrip("：:")
            res.detected.append("主送机关：标题后以冒号结尾的行")
            if text.endswith(":"):
                res.warnings.append("主送机关后用了半角冒号，应为全角冒号（GB/T 9704-2012 7.3.2）")
            pos += 1

    # 4. 从末尾倒推版记、附注、成文日期、署名
    end = n
    tail: List[Item] = []
    if end > pos:
        idx, text = lines[end - 1]
        m = PRINT_RE.match(text.replace("　", " "))
        if m:
            tail.insert(0, Item("print", text, idx))
            res.fields["印发机关"] = m.group("org").strip()
            res.fields["印发日期"] = m.group("date")
            end -= 1
    cc_items: List[Item] = []
    while end > pos and (CC_RE.match(lines[end - 1][1]) or BANJI_MAIN_RE.match(lines[end - 1][1])):
        idx, text = lines[end - 1]
        role = "cc" if CC_RE.match(text) else "banji_main"
        cc_items.insert(0, Item(role, text, idx))
        end -= 1
    if cc_items:
        for it in cc_items:
            if it.role == "cc":
                res.fields["抄送"] = re.sub(r"^抄\s*送\s*[：:]", "", it.text).strip()
        tail = cc_items + tail
    if any(it.role in ("print", "cc") for it in tail):
        res.detected.append("版记：末尾的抄送行、印发行")

    # 附注（在成文日期下一行）
    note_item: Optional[Item] = None
    if end - 1 > pos and NOTE_RE.match(lines[end - 1][1]) and looks_like_date(lines[end - 2][1]):
        note_item = Item("note", lines[end - 1][1], lines[end - 1][0])
        res.fields["附注"] = lines[end - 1][1]
        end -= 1
    # 成文日期
    date_item: Optional[Item] = None
    sig_items: List[Item] = []
    if end > pos and looks_like_date(lines[end - 1][1]):
        date_item = Item("date", lines[end - 1][1], lines[end - 1][0])
        res.fields["成文日期"] = lines[end - 1][1]
        end -= 1
        res.detected.append("成文日期：末尾的“××××年×月×日”行")
        if DATE_CN_RE.match(date_item.text.replace(" ", "")):
            res.warnings.append("成文日期用了汉字数字，应使用阿拉伯数字（GB/T 9704-2012 7.3.5.4）")
        # 署名：日期前不超过 3 行短行
        while end > pos and len(sig_items) < 3 and _is_signature_line(lines[end - 1][1]):
            sig_items.insert(0, Item("signature", lines[end - 1][1], lines[end - 1][0]))
            end -= 1
        if sig_items:
            res.fields["发文机关署名"] = [it.text for it in sig_items]
            res.detected.append("发文机关署名：成文日期前的单位名称行")

    # 附件说明：主体末尾以"附件："开头的行及其续行
    attach_items: List[Item] = []
    for k in range(end - 1, pos - 1, -1):
        if ATTACH_RE.match(lines[k][1]):
            rest = lines[k + 1:end]
            if all(ATTACH_CONT_RE.match(t) and len(t) <= 80 for _, t in rest):
                attach_items.append(Item("attachment", lines[k][1], lines[k][0], first=True))
                for i2, t2 in rest:
                    attach_items.append(Item("attachment", t2, i2, first=False))
                end = k
            break
        if not ATTACH_CONT_RE.match(lines[k][1]):
            break
    if attach_items:
        names = []
        first = re.sub(r"^附\s*件\s*[：:]\s*", "", attach_items[0].text)
        names.append(first)
        names.extend(it.text for it in attach_items[1:])
        res.fields["附件"] = names
        res.detected.append("附件说明：以“附件：”开头的行")

    # 5. 主体：层次标题、出席列席、正文
    body: List[Item] = []
    for k in range(pos, end):
        idx, text = lines[k]
        level = is_heading(text)
        if ATTENDANCE_RE.match(text):
            body.append(Item("attendance", text, idx))
        elif level:
            body.append(Item(level, text, idx, split=heading_split(text)))
        else:
            body.append(Item("body", text, idx))

    res.items = head + title_items + ([recipient] if recipient else []) + body + attach_items \
        + sig_items + ([date_item] if date_item else []) + ([note_item] if note_item else []) + tail
    if not body:
        res.warnings.append("没有识别到正文段落")
    for _, text in lines:
        if MARKERS_RE.search(text):
            res.warnings.append("正文中仍有【待补】或【待核实】标记，请补齐后再定稿")
            break
    for _, text in lines[:6]:
        if SECRET_WORDS_RE.search(text) or SECRET_RE.match(text):
            res.warnings.append("保密提示：材料含密级或“内部资料”等标识。涉密信息不得输入联网 AI 工具，请在单位内网环境处理")
            break
    return res


def _record_docno(res: Parsed, text: str) -> None:
    """从发文字号行取发文字号，以及同一行的份号、签发人。"""
    m = DOCNO_LINE_RE.match(text)
    if m:
        res.fields["发文字号"] = m.group("docno").strip()
        if m.group("copy"):
            res.fields["份号"] = m.group("copy")
        if m.group("signer"):
            res.fields["签发人"] = m.group("signer").strip()
    else:
        m2 = DOCNO_MARK_RE.match(text)
        res.fields["发文字号"] = m2.group("docno") if m2 else text
        if m2 and m2.group("signer"):
            res.fields["签发人"] = m2.group("signer").strip()


def _is_issuer(text: str, nxt: str) -> bool:
    """发文机关标志：以"文件"结尾的短行，或以"纪要"结尾且下一行是发文字号、期号。"""
    if ISSUER_MARK_RE.match(text):
        return True
    if len(text) > 30 or text.startswith("关于") or has_sentence_punct(text):
        return False
    if text.endswith("文件"):
        return True
    return text.endswith("纪要") and bool(DOCNO_LINE_RE.match(nxt) or QIHAO_RE.match(nxt))


def _looks_title(text: str) -> bool:
    """标题兜底判断：短、无句末标点、不是层次标题、不以冒号结尾。"""
    return (len(text) <= 40 and not has_sentence_punct(text) and not is_heading(text)
            and not text.endswith(("：", ":")) and not looks_like_date(text)
            and not DOCNO_RE.search(text) and not ATTACH_RE.match(text))


def summarize(parsed: Parsed) -> List[str]:
    """生成识别结果报告，便于用户核对。"""
    out: List[str] = []
    f = parsed.fields
    for key in ("发文机关标志", "发文字号", "签发人", "标题", "副标题", "主送机关", "附件", "发文机关署名", "成文日期",
                "附注", "抄送", "印发机关", "印发日期"):
        if key in f and f[key]:
            val = f[key]
            if isinstance(val, list):
                val = "；".join(str(v) for v in val)
            label = "主送机关（讲话稿为称谓）" if key == "主送机关" and parsed.kind != "法定" else key
            out.append(f"{label}：{val}")
    counts: Dict[str, int] = {}
    for it in parsed.items:
        counts[it.role] = counts.get(it.role, 0) + 1
    parts = [f"{ROLE_LABELS[r]} {counts[r]} 段" for r in ("h1", "h2", "h3", "h4", "body", "attendance") if counts.get(r)]
    if parts:
        out.append("主体：" + "，".join(parts))
    if parsed.doc_type:
        out.append(f"文种：{parsed.doc_type}（{parsed.kind}）")
    return out
