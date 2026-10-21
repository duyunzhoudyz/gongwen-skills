#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""基于公开公文语料，按文种重算量化基线（shared/style/baseline.json）。

输入（语料目录，位于仓库之外的临时目录，不入库）：
  <语料目录>/sources.json         来源清单：[{doc_type, file, title, org, date, url, segment, subtype}, ...]
                                  也可写成 {"meta": {...}, "docs": [...]}
  <语料目录>/<文种>/<文件>.txt    纯文本，一行一个自然段（UTF-8）

输出：
  baseline.json，含各文种的样本数、机关数、全文字数、段落、句长、层次标题、“一是……二是……”、
  顿号密度、四字格比例、开头与结语方式分布、套话与评价词的千字频次，
  以及 shared/scripts/style_patterns.json 中每条词表规则在真实公文里的命中情况。
  仓库只保存统计结果和出处（标题、发文机关、日期、URL、字数、文本哈希），不保存全文。

输出是确定性的：同样的语料，生成的 JSON 逐字节相同（不含时间戳，键按固定顺序）。
退出码：0 成功；1 --check 发现不一致；2 输入有误（语料目录、sources.json、文件缺失等）。

用法示例：
  python tools/build_baseline.py --corpus D:/tmp/corpus
  python tools/build_baseline.py --corpus D:/tmp/corpus --check
  python tools/build_baseline.py --corpus D:/tmp/corpus --print-tables
"""
from __future__ import annotations

import hashlib
import json
import re
import sys
from collections import Counter, OrderedDict
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

_TOOLS_DIR = str(Path(__file__).resolve().parent)
if _TOOLS_DIR not in sys.path:
    sys.path.insert(0, _TOOLS_DIR)

from sync_shared import REPO_ROOT, make_parser, setup_console  # noqa: E402

BASELINE_VERSION = "1.0.0"
DEFAULT_OUT_REL = "shared/style/baseline.json"
DEFAULT_PATTERNS_REL = "shared/scripts/style_patterns.json"

# 本基线的文种与 gongwen_lint.py 文书类型的对应（用于词表规则的 doc_types / exempt_doc_types 判断）
LINT_DOC_TYPE = {
    "通知": "通知", "通知（印发转发）": "通知", "意见": "意见", "方案": "实施方案", "批复": "批复",
    "函": "函", "通报": "通报", "报告": "报告", "政府工作报告": "报告", "讲话": "讲话稿", "工作总结": "总结",
}

# 文种展示顺序；不在表中的文种排在后面，按名称排序
DOC_TYPE_ORDER = ["通知", "通知（印发转发）", "意见", "方案", "批复", "函", "通报", "报告", "政府工作报告", "讲话", "工作总结"]

MIN_SAMPLES = 10   # 低于此样本数标记 low_sample
MIN_ORGS = 3       # 低于此机关数标记 low_orgs
MIN_SUBGROUP = 5   # 子类至少要有这么多样本才单独成组

# ---------------------------------------------------------------------------
# 基础统计
# ---------------------------------------------------------------------------


def percentile(values: Sequence[float], q: float) -> Optional[float]:
    """线性插值分位数（与 numpy 默认一致）；q 取 0～100；空序列返回 None。"""
    if not values:
        return None
    data = sorted(values)
    if len(data) == 1:
        return float(data[0])
    pos = (len(data) - 1) * q / 100.0
    low = int(pos)
    high = min(low + 1, len(data) - 1)
    frac = pos - low
    return data[low] + (data[high] - data[low]) * frac


def summarize(values: Sequence[float], digits: int = 1) -> Dict[str, object]:
    """样本数、最小、P10、四分位、中位数、均值、P90、最大。样本为空时只给 n=0。"""
    vals = [float(v) for v in values]
    if not vals:
        return {"n": 0}

    def r(x: Optional[float]) -> float:
        return round(x, digits) if digits else int(round(x))  # type: ignore[arg-type]

    return OrderedDict([
        ("n", len(vals)),
        ("min", r(min(vals))),
        ("p10", r(percentile(vals, 10))),
        ("q1", r(percentile(vals, 25))),
        ("median", r(percentile(vals, 50))),
        ("q3", r(percentile(vals, 75))),
        ("p90", r(percentile(vals, 90))),
        ("max", r(max(vals))),
        ("mean", r(sum(vals) / len(vals))),
    ])


# ---------------------------------------------------------------------------
# 文本预处理与单篇指标
# ---------------------------------------------------------------------------

WENHAO_RE = re.compile(r"[\u4e00-\u9fa5A-Za-z]{0,12}[〔\[]\s*\d{4}\s*[〕\]]\s*\d+\s*号")
DATE_LINE_RE = re.compile(
    r"(?:\d{4}年\s*\d{1,2}月\s*\d{1,2}日|[一二三四五六七八九〇零○]{4}年[一二三四五六七八九十]{1,3}月[一二三四五六七八九十]{1,3}日)\s*$")
SENT_END = "。？！；"
BODY_END_RE = re.compile(r"[。？！；]")

HEADING_RES = (
    (1, re.compile(r"^\s*[一二三四五六七八九十]{1,3}\s*[、，,.．]")),
    (2, re.compile(r"^\s*[（(]\s*[一二三四五六七八九十]{1,3}\s*[）)]")),
    (3, re.compile(r"^\s*\d{1,2}\s*(?:[.．](?!\d)|、(?!\d))")),
    (4, re.compile(r"^\s*[（(]\s*\d{1,2}\s*[）)]")),
)

# “一是……二是……”：序数词后紧跟“是”，且前面是句首或标点（排除“统一是”“唯一是”“之一是”）
YISHI_RE = re.compile(r"(?:^|(?<=[，。；：！？、）」”\s]))[一二三四五六七八九十]是")
YISHI_SET_RE = re.compile(r"(?:^|(?<=[，。；：！？、）」”\s]))一是")

CLAUSE_SPLIT_RE = re.compile(r"[^\u4e00-\u9fa5]+")


def clean_paragraphs(lines: Iterable[str]) -> List[str]:
    """去掉空行、统一空白、去掉发文字号（不同来源的页面把文号和标题粘在同一行）。"""
    paras: List[str] = []
    for line in lines:
        text = WENHAO_RE.sub("", line.replace("\u3000", " ").replace("\xa0", " "))
        text = re.sub(r"[（(]\s*[）)]", "", text)  # 去掉文号后留下的空括号
        text = re.sub(r"\s+", " ", text).strip()
        if text:
            paras.append(text)
    return paras


def count_chars(paras: Sequence[str]) -> int:
    """字数：全文非空白字符数（含标点）。"""
    return sum(len(re.sub(r"\s+", "", p)) for p in paras)


def split_sentences(paras: Sequence[str]) -> List[int]:
    """按句号、问号、叹号、分号切分，返回每句的字数（不含句末标点）。
    不含句末标点的自然段（标题、署名行）不计；含句末标点的自然段末尾的残句不足 8 字不计。"""
    lengths: List[int] = []
    for p in paras:
        if not BODY_END_RE.search(p):
            continue
        parts = re.split(r"[。？！；]", p)
        for i, seg in enumerate(parts):
            n = len(re.sub(r"\s+", "", seg))
            if n == 0:
                continue
            if i == len(parts) - 1 and n < 8:  # 末尾没有标点收尾的残句
                continue
            lengths.append(n)
    return lengths


def heading_level(para: str) -> int:
    for level, regex in HEADING_RES:
        if regex.match(para):
            return level
    return 0


def count_headings(paras: Sequence[str]) -> Dict[int, int]:
    counts = {1: 0, 2: 0, 3: 0, 4: 0}
    for p in paras:
        level = heading_level(p)
        if level:
            counts[level] += 1
    return counts


def count_yishi(paras: Sequence[str]) -> Tuple[int, int]:
    """返回 (“一是”出现次数即列举组数, 全部“X是”分项数)。"""
    sets = sum(len(YISHI_SET_RE.findall(p)) for p in paras)
    items = sum(len(YISHI_RE.findall(p)) for p in paras)
    return sets, items


def four_char_ratio(paras: Sequence[str]) -> Optional[float]:
    """四字格比例（近似算法）：以非汉字字符（标点、数字、字母、空白）为界切出连续汉字片段，
    再以“的、了、和、与、及、等”之外不再细分；片段恰为 4 个汉字的占全部片段的比例。
    这是“四字短语密度”的代理指标，不做成语词典匹配。"""
    total = four = 0
    for p in paras:
        for seg in CLAUSE_SPLIT_RE.split(p):
            if not seg:
                continue
            total += 1
            if len(seg) == 4:
                four += 1
    return four / total if total else None


# ---------------------------------------------------------------------------
# 开头与结语方式
# ---------------------------------------------------------------------------

SIGNOFF_ORG_RE = re.compile(r"(?:局|厅|委|部|办|院|府|会|处|室|中心|政府|办公室|办公厅|单位|党组|党委|公司|集团|总署|总局)$")

OPENING_RULES = (
    ("代表政府报告式", re.compile(r"^(?:现在，?)?我(?:代表|受)")),
    ("来文答复式", re.compile(r"^[^。]{0,80}(?:收悉|已收到)|^你[^。，]{0,30}(?:来函|报来的|上报的)")),
    ("目的起笔式", re.compile(r"^为(?!何)")),
    ("依据起笔式", re.compile(r"^(?:根据|按照|依据|遵照|依照|经|鉴于|按|据)")),
    ("直接印发式", re.compile(r"^(?:现将|《|现印发|现转发|现就|现通知|现批复|现函复)")),
    ("分条直陈式", re.compile(r"^(?:[一二三四五六七八九十]{1,3}\s*[、，,.．]|[（(]\s*[一二三四五六七八九十\d]{1,3}\s*[）)]|\d{1,2}\s*[.．、])")),
    ("会议起笔式", re.compile(r"^(?:同志们|各位|尊敬的|女士们|先生们|大家好|同学们|老师们|很高兴|刚才|今天|这次|本次|此次|在.{0,30}(?:会议|大会|座谈会|论坛))")),
    ("背景起笔式", re.compile(r"^(?:近年来|近期|当前|目前|随着|今年以来|去年以来|自\S{0,12}以来|党的|中央|习近平|\d{4}年|二〇|\d{1,2}月|过去|刚刚过去|新)")),
)

CLOSING_RULES = (
    ("特此X式", re.compile(r"特此(?:通知|报告|函告|函复|函|通报|公告|通告|批复|说明|复函|答复|函达)")),
    ("请批示审议式", re.compile(r"(?:妥否|当否|可否|以上|上述).{0,16}请(?:予)?(?:批示|审批|批复|批准|审定|审议|审查|指示|审阅)|请(?:予)?(?:批示|审批|批准|审定|审议|审查)")),
    ("此复式", re.compile(r"此复")),
    ("重大事项请示报告式", re.compile(r"重大(?:事项|情况|问题).{0,10}(?:及时)?按程序.{0,12}请示报告")),
    ("报送反馈联系式", re.compile(r"请(?:于|在).{0,30}(?:前|内).{0,30}(?:报送|反馈|报告|报我)|(?:情况和问题|问题).{0,6}请.{0,6}(?:联系|沟通|及时)|联系人|联系电话")),
    ("请贯彻落实式", re.compile(r"请.{0,24}(?:贯彻|执行|落实|遵照|组织实施|抓好|做好|认真研究|遵循)|要.{0,6}(?:认真|切实|严格)?(?:贯彻|执行|落实)")),
    ("请函复支持式", re.compile(r"请(?:予)?(?:研究|函复|复函|回复|支持|协助|配合|协调)|函复|盼复")),
    ("号召致谢式", re.compile(r"谢谢|让我们|共同努力|团结奋斗|奋斗|努力奋斗|再创|更大贡献|[！!]$")),
    ("展望表态式", re.compile(r"我们将|将继续|下一步|今后|坚决|一定|有信心")),
)


def body_paragraphs(paras: Sequence[str]) -> Tuple[List[str], bool]:
    """去掉标题、称呼行、署名日期等，返回 (正文候选自然段, 是否出现过称呼行)。"""
    out: List[str] = []
    saw_address = False
    for i, p in enumerate(paras):
        if len(p) <= 600 and re.search(r"[：:]$", p) and not BODY_END_RE.search(p):  # 主送机关、称呼行
            saw_address = True
            continue
        if not BODY_END_RE.search(p) and len(p) < 70:  # 标题、副标题、署名、日期、版记
            continue
        out.append(p)
    return out, saw_address


def classify_opening(paras: Sequence[str]) -> str:
    body, _saw_address = body_paragraphs(paras)
    if not body:
        return "无法判断"
    first = body[0].strip()
    for name, regex in OPENING_RULES:
        if regex.search(first):
            return name
    return "直陈式（其他）"


def has_address_line(paras: Sequence[str]) -> bool:
    """正文前是否有“XX：”形式的主送机关或称呼行。"""
    return body_paragraphs(paras)[1]


def strip_signoff(paras: Sequence[str]) -> List[str]:
    """去掉末尾的日期行、署名行、附件说明、公开属性说明。"""
    rest = list(paras)
    while rest:
        last = rest[-1].strip()
        compact = last.replace(" ", "")
        if DATE_LINE_RE.search(compact) and not re.search(r"[。；，,]", compact[-80:]) and len(compact) <= 200:
            rest.pop(); continue
        if re.match(r"^(?:附件|抄送|主送|（?此件|联系人|联系方式|联系电话|电话|邮箱|电子邮箱|地址|传真|图\s*[\d一二三四五六七八九十]|表\s*[\d一二三四五六七八九十])", last):
            rest.pop(); continue
        if len(last) <= 40 and not BODY_END_RE.search(last) and not re.search(r"[，,：:]", last):
            rest.pop(); continue
        break
    return rest


def classify_closing(paras: Sequence[str]) -> str:
    rest = strip_signoff(paras)
    if not rest:
        return "无法判断"
    last = rest[-1]
    sentences = [s for s in re.split(r"(?<=[。？！；!])", last) if s.strip()]
    tail = " ".join(sentences[-2:]) if sentences else last
    for name, regex in CLOSING_RULES:
        if regex.search(tail):
            return name
    return "无固定结语（以实质内容收束）"


# ---------------------------------------------------------------------------
# 套话与评价词（千字频次基准）
# ---------------------------------------------------------------------------

PHRASE_GROUPS: "OrderedDict[str, List[str]]" = OrderedDict([
    ("程度副词与动词套话", [
        "进一步", "不断", "持续", "切实", "全面", "深入", "扎实", "稳步", "有序", "着力", "加快", "严格",
        "积极", "充分", "大力", "统筹", "强化", "加强", "推进", "推动", "提升", "落实", "贯彻", "坚持",
        "抓好", "抓实", "压实", "确保", "做好", "完善", "健全", "优化", "加大",
    ]),
    ("政策口号与抽象名词", [
        "高质量", "高水平", "现代化", "长效机制", "常态化", "体系", "机制", "格局", "合力", "协同",
        "联动", "抓手", "载体", "平台", "闭环", "赋能", "助力", "打造", "责任制", "清单",
    ]),
    ("评价词", [
        "重要", "显著", "明显", "良好", "有效", "突出", "巨大", "重大", "深刻", "关键", "基本", "较好",
        "取得成效", "成效", "不足", "问题", "短板", "亮点",
    ]),
    ("惯用句式", [
        "现将", "特此", "如下", "以上", "请予", "同意", "认真", "切实做好", "按照", "根据", "为了",
        "要求", "必须", "应当", "不得", "严禁", "可以",
    ]),
])


def phrase_counts(text: str) -> Dict[str, int]:
    counts: Dict[str, int] = {}
    for words in PHRASE_GROUPS.values():
        for w in words:
            counts[w] = text.count(w)
    return counts


# ---------------------------------------------------------------------------
# 词表规则（shared/scripts/style_patterns.json）在真实公文上的命中
# ---------------------------------------------------------------------------


class StylePat:
    """与 gongwen_lint.py 的 style_hits 口径一致：逐自然段匹配，剔除白名单与 exclude 重叠的命中。"""

    def __init__(self, item: dict) -> None:
        self.id = item["id"]
        self.group = item.get("group", "style")
        self.category = item.get("category", "")
        self.level = item.get("level", "P2")
        self.regex = re.compile(item["pattern"])
        self.density = float(item["density_per_1000"]) if item.get("density_per_1000") is not None else None
        self.doc_types = tuple(item.get("doc_types") or ())
        self.exempt = tuple(item.get("exempt_doc_types") or ())
        self.exclude = tuple(re.compile(p) for p in (item.get("exclude") or ()))

    def applies(self, lint_type: str) -> bool:
        if self.doc_types and lint_type not in self.doc_types:
            return False
        return lint_type not in self.exempt


def load_style_patterns(path: Path) -> Tuple[List[StylePat], List["re.Pattern[str]"], str]:
    data = json.loads(path.read_text(encoding="utf-8"))
    pats = [StylePat(item) for item in data.get("patterns", [])]
    whitelist = [re.compile(w["pattern"]) for w in data.get("whitelist", [])]
    return pats, whitelist, str(data.get("version", ""))


def count_pattern_hits(paras: Sequence[str], pat: StylePat, whitelist: Sequence["re.Pattern[str]"]) -> int:
    hits = 0
    for para in paras:
        wl = [m.span() for regex in whitelist for m in regex.finditer(para) if m.end() > m.start()]
        own: Optional[List[Tuple[int, int]]] = None
        for m in pat.regex.finditer(para):
            start, end = m.span()
            if start == end:
                continue
            if any(ws < end and start < we for ws, we in wl):
                continue
            if pat.exclude:
                if own is None:
                    own = [mm.span() for regex in pat.exclude for mm in regex.finditer(para) if mm.end() > mm.start()]
                if any(ws < end and start < we for ws, we in own):
                    continue
            hits += 1
    return hits


# ---------------------------------------------------------------------------
# 单篇文档与语料加载
# ---------------------------------------------------------------------------


class Doc:
    def __init__(self, meta: dict, paras: List[str]) -> None:
        self.meta = meta
        self.doc_type: str = meta["doc_type"]
        self.subtype: str = meta.get("subtype") or ""
        self.paras = paras
        self.chars = count_chars(paras)
        self.text = "\n".join(paras)
        self.sha = hashlib.sha256(self.text.encode("utf-8")).hexdigest()[:16]


class CorpusError(Exception):
    pass


def load_corpus(corpus_dir: Path) -> Tuple[List[Doc], dict]:
    src_path = corpus_dir / "sources.json"
    if not src_path.is_file():
        raise CorpusError(f"语料目录缺少 sources.json：{src_path}")
    try:
        raw = json.loads(src_path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise CorpusError(f"sources.json 读取失败：{exc}")
    if isinstance(raw, dict):
        meta = raw.get("meta") or {}
        entries = raw.get("docs")
    else:
        meta, entries = {}, raw
    if not isinstance(entries, list) or not entries:
        raise CorpusError("sources.json 中没有文档清单（应为数组，或 {\"docs\": [...]}）")
    docs: List[Doc] = []
    for i, entry in enumerate(entries, start=1):
        for key in ("doc_type", "file", "title", "org", "url"):
            if not entry.get(key):
                raise CorpusError(f"sources.json 第 {i} 项缺少字段 {key}")
        path = corpus_dir / entry["file"]
        if not path.is_file():
            raise CorpusError(f"sources.json 第 {i} 项的文件不存在：{entry['file']}")
        paras = clean_paragraphs(path.read_text(encoding="utf-8").split("\n"))
        if not paras:
            raise CorpusError(f"文件为空：{entry['file']}")
        docs.append(Doc(entry, paras))
    return docs, meta


# ---------------------------------------------------------------------------
# 分组统计
# ---------------------------------------------------------------------------

LENGTH_BINS = [(0, 1000), (1000, 2000), (2000, 4000), (4000, 8000), (8000, 16000), (16000, 10 ** 9)]


def bin_label(low: int, high: int) -> str:
    if low == 0:
        return f"<{high}"
    if high >= 10 ** 9:
        return f">={low}"
    return f"{low}-{high}"


def org_key(org: str) -> str:
    """机关去重键：取第一个机关，去掉“国家”前缀与空白，使“发展改革委”“国家发展改革委”算同一机关。"""
    first = re.split(r"[\s、，,]+", org.strip())[0] if org.strip() else org
    return re.sub(r"^国家", "", first)


def per_doc_metrics(doc: Doc) -> dict:
    paras = doc.paras
    sent = split_sentences(paras)
    heads = count_headings(paras)
    yi_sets, yi_items = count_yishi(paras)
    body = [p for p in paras if BODY_END_RE.search(p)]
    chars = max(doc.chars, 1)
    dun = doc.text.count("、")
    return {
        "chars": doc.chars,
        "paras": len(paras),
        "body_paras": len(body),
        "body_para_len_mean": (sum(len(re.sub(r"\s+", "", p)) for p in body) / len(body)) if body else None,
        "body_para_len_median": percentile([len(re.sub(r"\s+", "", p)) for p in body], 50) if body else None,
        "sentences": len(sent),
        "sent_len_mean": (sum(sent) / len(sent)) if sent else None,
        "sent_len_median": percentile(sent, 50) if sent else None,
        "h1": heads[1], "h2": heads[2], "h3": heads[3], "h4": heads[4],
        "h_total": sum(heads.values()),
        "yishi_sets": yi_sets, "yishi_items": yi_items,
        "dunhao_per_1000": dun * 1000.0 / chars,
        "four_char_ratio": four_char_ratio(paras),
    }


def distribution(labels: Sequence[str]) -> "OrderedDict[str, object]":
    counts = Counter(labels)
    total = sum(counts.values())
    out: "OrderedDict[str, object]" = OrderedDict()
    for name, n in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0])):
        out[name] = OrderedDict([("n", n), ("share", round(n / total, 3))])
    return out


def aggregate_group(docs: Sequence[Doc], metrics: Sequence[dict], phrase_rows: Sequence[Dict[str, int]],
                    pattern_rows: Sequence[Dict[str, int]], pats: Sequence[StylePat], lint_type: str) -> dict:
    """汇总一个分组（文种或文种下的子类）。"""
    n = len(docs)
    orgs = sorted({org_key(d.meta["org"]) for d in docs})
    years = Counter((d.meta.get("date") or "")[:4] or "未知" for d in docs)

    def col(name: str) -> List[float]:
        return [m[name] for m in metrics if m[name] is not None]

    result: "OrderedDict[str, object]" = OrderedDict()
    result["n"] = n
    result["org_count"] = len(orgs)
    result["low_sample"] = n < MIN_SAMPLES
    result["low_orgs"] = len(orgs) < MIN_ORGS
    result["years"] = OrderedDict(sorted(years.items()))
    result["chars"] = summarize(col("chars"), 0)
    result["paras"] = summarize(col("paras"), 0)
    result["body_paras"] = summarize(col("body_paras"), 0)
    result["body_para_len_median"] = summarize(col("body_para_len_median"), 0)
    result["body_para_len_mean"] = summarize(col("body_para_len_mean"), 0)
    result["sent_len_mean"] = summarize(col("sent_len_mean"), 1)
    result["sent_len_median"] = summarize(col("sent_len_median"), 1)

    heading: "OrderedDict[str, object]" = OrderedDict()
    totals = {lv: sum(m["h%d" % lv] for m in metrics) for lv in (1, 2, 3, 4)}
    grand = sum(totals.values())
    heading["share_of_docs_with_heading"] = round(sum(1 for m in metrics if m["h_total"] > 0) / n, 3)
    heading["level_share"] = OrderedDict(
        (f"L{lv}", round(totals[lv] / grand, 3) if grand else 0.0) for lv in (1, 2, 3, 4))
    for lv in (1, 2, 3, 4):
        heading[f"h{lv}_per_doc"] = summarize([m["h%d" % lv] for m in metrics], 1)
    heading["h_total_per_doc"] = summarize([m["h_total"] for m in metrics], 1)
    chars_per_h1 = [m["chars"] / m["h1"] for m in metrics if m["h1"] > 0]
    heading["chars_per_h1"] = summarize(chars_per_h1, 0)
    bins = []
    for low, high in LENGTH_BINS:
        idx = [i for i, m in enumerate(metrics) if low <= m["chars"] < high]
        if len(idx) < 3:
            continue
        bins.append(OrderedDict([
            ("chars", bin_label(low, high)), ("n", len(idx)),
            ("h1_median", round(percentile([metrics[i]["h1"] for i in idx], 50), 1)),
            ("h2_median", round(percentile([metrics[i]["h2"] for i in idx], 50), 1)),
            ("h3_median", round(percentile([metrics[i]["h3"] for i in idx], 50), 1)),
            ("paras_median", round(percentile([metrics[i]["paras"] for i in idx], 50), 0)),
        ]))
    heading["by_length"] = bins
    result["headings"] = heading

    yishi: "OrderedDict[str, object]" = OrderedDict()
    yishi["share_of_docs_with_yishi"] = round(sum(1 for m in metrics if m["yishi_sets"] > 0) / n, 3)
    yishi["sets_per_doc"] = summarize([m["yishi_sets"] for m in metrics], 1)
    yishi["items_per_doc"] = summarize([m["yishi_items"] for m in metrics], 1)
    yishi["sets_per_1000_chars"] = summarize([m["yishi_sets"] * 1000.0 / max(m["chars"], 1) for m in metrics], 2)
    result["yishi"] = yishi

    result["dunhao_per_1000"] = summarize(col("dunhao_per_1000"), 1)
    result["four_char_ratio"] = summarize(col("four_char_ratio"), 3)

    result["address_line_share"] = round(sum(1 for d in docs if has_address_line(d.paras)) / n, 3)
    result["openings"] = distribution([classify_opening(d.paras) for d in docs])
    result["closings"] = distribution([classify_closing(d.paras) for d in docs])

    # 套话与评价词：每千字频次（逐篇计算后取分位）
    phrases: "OrderedDict[str, object]" = OrderedDict()
    for gname, words in PHRASE_GROUPS.items():
        rows: "OrderedDict[str, object]" = OrderedDict()
        for w in words:
            per_k = [row[w] * 1000.0 / max(d.chars, 1) for d, row in zip(docs, phrase_rows)]
            rows[w] = [
                round(percentile(per_k, 50), 2),
                round(sum(per_k) / n, 2),
                round(percentile(per_k, 90), 2),
                round(sum(1 for v in per_k if v > 0) / n, 3),
            ]
        phrases[gname] = rows
    result["phrases_per_1000"] = phrases

    # 词表规则：在该组真实公文里的命中情况
    sp: "OrderedDict[str, object]" = OrderedDict()
    for pat in pats:
        if not pat.applies(lint_type):
            continue
        counts = [row[pat.id] for row in pattern_rows]
        if not any(counts):
            continue
        per_k = [c * 1000.0 / max(d.chars, 1) for c, d in zip(counts, docs)]
        entry: "OrderedDict[str, object]" = OrderedDict([
            ("docs_with_hit", sum(1 for c in counts if c > 0)),
            ("doc_share", round(sum(1 for c in counts if c > 0) / n, 3)),
            ("per_1000_pooled", round(sum(counts) * 1000.0 / max(sum(d.chars for d in docs), 1), 3)),
            ("per_1000_p90", round(percentile(per_k, 90), 3)),
            ("per_1000_max", round(max(per_k), 3)),
        ])
        if pat.density is not None:
            over = sum(1 for c, v in zip(counts, per_k) if c >= 2 and v > pat.density)
            entry["threshold"] = pat.density
            entry["docs_over_threshold"] = over
            entry["over_threshold_share"] = round(over / n, 3)
        sp[pat.id] = entry
    result["style_patterns"] = sp
    return result


def build_baseline(docs: List[Doc], meta: dict, patterns_path: Optional[Path]) -> dict:
    pats: List[StylePat] = []
    whitelist: List["re.Pattern[str]"] = []
    pat_version = ""
    if patterns_path is not None and patterns_path.is_file():
        pats, whitelist, pat_version = load_style_patterns(patterns_path)

    metrics_all = [per_doc_metrics(d) for d in docs]
    phrase_all = [phrase_counts(d.text) for d in docs]
    pattern_all = [{p.id: count_pattern_hits(d.paras, p, whitelist) for p in pats} for d in docs]

    order = {name: i for i, name in enumerate(DOC_TYPE_ORDER)}
    types = sorted({d.doc_type for d in docs}, key=lambda t: (order.get(t, 99), t))

    out: "OrderedDict[str, object]" = OrderedDict()
    out["version"] = BASELINE_VERSION
    out["corpus"] = OrderedDict([
        ("fetched", meta.get("fetched", "")),
        ("note", meta.get("note", "")),
        ("docs", len(docs)),
        ("style_patterns_version", pat_version),
    ])
    out["method"] = OrderedDict([
        ("chars", "全文非空白字符数（含标点）；已去掉页面导航、版权、发文字号"),
        ("sentence", "按句号、问号、叹号、分号切分；句长为句内非空白字符数；不含句末标点的自然段（标题、署名）不计"),
        ("paragraph", "一个非空行为一个自然段；body_paras 指含句末标点的自然段（排除标题、称呼、署名、日期）"),
        ("heading", "自然段开头匹配 一、/（一）/1./（1），分别记 L1～L4；一个自然段最多计一级"),
        ("yishi", "“一是”“二是”……前面为句首或标点；sets 为“一是”出现次数（列举组数），items 为全部“X是”分项数"),
        ("dunhao", "每千字顿号个数"),
        ("four_char_ratio", "以非汉字字符为界切出连续汉字片段，恰为 4 个汉字的片段占全部片段的比例（四字短语密度的代理指标，不做成语词典匹配）"),
        ("opening", "取首个正文自然段，按固定规则归类，规则见 tools/build_baseline.py 的 OPENING_RULES"),
        ("closing", "去掉署名、日期、附件说明后取末段末两句，按固定规则归类，规则见 CLOSING_RULES"),
        ("phrases_per_1000", "每个词一行 [中位数, 均值, P90, 出现该词的文档占比]；每篇先算千字频次，再在该组内取统计量"),
        ("style_patterns", "与 gongwen_lint.py 同口径（逐自然段匹配，剔除白名单与 exclude 重叠）；"
                           "per_1000_pooled 为合计命中数除以合计字数；带阈值的规则另给超阈值文档占比"
                           "（命中不少于 2 处且千字密度大于阈值，与机检判定一致）"),
    ])

    groups: "OrderedDict[str, object]" = OrderedDict()
    for t in types:
        idx = [i for i, d in enumerate(docs) if d.doc_type == t]
        lint_type = LINT_DOC_TYPE.get(t, t)
        groups[t] = aggregate_group([docs[i] for i in idx], [metrics_all[i] for i in idx],
                                    [phrase_all[i] for i in idx], [pattern_all[i] for i in idx], pats, lint_type)
        # 子类
        subs = sorted({docs[i].subtype for i in idx if docs[i].subtype})
        if len(subs) >= 2:
            for s in subs:
                sidx = [i for i in idx if docs[i].subtype == s]
                if len(sidx) < MIN_SUBGROUP:
                    continue
                groups[f"{t}/{s}"] = aggregate_group(
                    [docs[i] for i in sidx], [metrics_all[i] for i in sidx],
                    [phrase_all[i] for i in sidx], [pattern_all[i] for i in sidx], pats, lint_type)
    out["groups"] = groups

    # 词表规则总览：把各文种合在一起看（每个规则只统计其适用文种的文档）
    overview: "OrderedDict[str, object]" = OrderedDict()
    for pat in pats:
        applicable = [i for i, d in enumerate(docs) if pat.applies(LINT_DOC_TYPE.get(d.doc_type, d.doc_type))]
        if not applicable:
            continue
        counts = [pattern_all[i][pat.id] for i in applicable]
        if not any(counts):
            continue
        n = len(applicable)
        row: "OrderedDict[str, object]" = OrderedDict([
            ("group", pat.group), ("category", pat.category), ("level", pat.level),
            ("docs_applicable", n),
            ("docs_with_hit", sum(1 for c in counts if c > 0)),
            ("doc_share", round(sum(1 for c in counts if c > 0) / n, 3)),
            ("per_1000_pooled", round(sum(counts) * 1000.0 / max(sum(docs[i].chars for i in applicable), 1), 3)),
        ])
        if pat.density is not None:
            over = sum(1 for i in applicable
                       if pattern_all[i][pat.id] >= 2 and pattern_all[i][pat.id] * 1000.0 / max(docs[i].chars, 1) > pat.density)
            row["threshold"] = pat.density
            row["docs_over_threshold"] = over
            row["over_threshold_share"] = round(over / n, 3)
        overview[pat.id] = row
    out["style_patterns_overview"] = overview
    # 在全部适用文档中一次都没有命中的规则：真实公文里不出现，是区分 AI 腔的好信号
    out["style_patterns_zero_hit"] = [p.id for p in pats if p.id not in overview]

    sources: List[object] = []
    for d in sorted(docs, key=lambda x: (order.get(x.doc_type, 99), x.doc_type, x.meta.get("date") or "", x.meta["url"])):
        sources.append(OrderedDict([
            ("doc_type", d.doc_type),
            ("subtype", d.subtype),
            ("title", d.meta["title"]),
            ("org", d.meta["org"]),
            ("date", d.meta.get("date") or ""),
            ("url", d.meta["url"]),
            ("segment", d.meta.get("segment") or "main"),
            ("chars", d.chars),
            ("sha256_16", d.sha),
        ]))
    out["sources"] = sources
    return out


def _is_scalar(v: object) -> bool:
    return v is None or isinstance(v, (str, int, float, bool))


def _inline(v: object) -> str:
    return json.dumps(v, ensure_ascii=False, separators=(", ", ": "))


def _dump(v: object, level: int, out: List[str]) -> None:
    pad = "  " * level
    if isinstance(v, dict):
        if not v:
            out.append("{}")
            return
        if all(_is_scalar(x) for x in v.values()) or (
                all(isinstance(x, list) and all(_is_scalar(y) for y in x) for x in v.values()) and level >= 4):
            out.append(_inline(v))
            return
        out.append("{")
        items = list(v.items())
        for i, (k, x) in enumerate(items):
            buf: List[str] = []
            _dump(x, level + 1, buf)
            comma = "," if i < len(items) - 1 else ""
            out.append(f"{pad}  {json.dumps(k, ensure_ascii=False)}: {buf[0]}")
            out.extend(buf[1:])
            out[-1] += comma
        out.append(pad + "}")
    elif isinstance(v, list):
        if all(_is_scalar(x) for x in v):
            out.append(_inline(v))
            return
        out.append("[")
        for i, x in enumerate(v):
            buf = []
            _dump(x, level + 1, buf)
            comma = "," if i < len(v) - 1 else ""
            out.append(f"{pad}  {buf[0]}")
            out.extend(buf[1:])
            out[-1] += comma
        out.append(pad + "]")
    else:
        out.append(json.dumps(v, ensure_ascii=False))


def dump_json(data: dict) -> str:
    """确定性输出：叶子字典与纯标量数组写成一行，其余按 2 空格缩进。"""
    out: List[str] = []
    _dump(data, 0, out)
    return "\n".join(out) + "\n"


# ---------------------------------------------------------------------------
# Markdown 表格（供 baseline.md 引用，保证数字与 JSON 一致）
# ---------------------------------------------------------------------------


TABLE_PHRASE_GROUPS = ["通知", "意见", "方案", "通报", "政府工作报告", "讲话"]
TABLE_KEY_WORDS = ["进一步", "切实", "加强", "推进", "推动", "落实", "强化", "高质量", "打造", "助力", "赋能", "有效", "显著", "成效", "问题"]


def _pct(x: float) -> str:
    """百分比取整；不足 1% 的非零值写成“<1%”，避免四舍五入成 0%。"""
    if x <= 0:
        return "0%"
    if x < 0.005:
        return "<1%"
    return f"{x * 100:.0f}%"


def _iqr(s: dict, digits: int = 0) -> str:
    if not s or s.get("n", 0) == 0:
        return "—"
    fmt = "{:.%df}" % digits
    return f"{fmt.format(s['median'])}（{fmt.format(s['q1'])}～{fmt.format(s['q3'])}）"


def render_tables(baseline: dict) -> str:
    groups: dict = baseline["groups"]
    lines: List[str] = []
    lines.append("### 样本与篇幅")
    lines.append("")
    lines.append("| 文种 | 样本数 | 机关数 | 全文字数 中位数（四分位） | 自然段数 | 平均句长（字） |")
    lines.append("|---|---:|---:|---|---|---|")
    for name, g in groups.items():
        flag = "（样本偏少）" if g["low_sample"] else ""
        lines.append(f"| {name}{flag} | {g['n']} | {g['org_count']} | {_iqr(g['chars'])} | {_iqr(g['paras'])} | {_iqr(g['sent_len_mean'], 1)} |")
    lines.append("")
    lines.append("### 层次标题与列举")
    lines.append("")
    lines.append("| 文种 | 有标题的文档占比 | 一级标题数/篇 | 二级标题数/篇 | 三级标题数/篇 | 一、二、三级占比 | “一是”组数/篇 | 顿号/千字 | 四字格比例 |")
    lines.append("|---|---:|---|---|---|---|---|---|---|")
    for name, g in groups.items():
        h = g["headings"]
        ls = h["level_share"]
        share = f"{ls['L1']:.0%}/{ls['L2']:.0%}/{ls['L3']:.0%}" if sum(ls.values()) else "—"
        lines.append(
            f"| {name} | {h['share_of_docs_with_heading']:.0%} | {_iqr(h['h1_per_doc'], 1)} | {_iqr(h['h2_per_doc'], 1)} | "
            f"{_iqr(h['h3_per_doc'], 1)} | {share} | {_iqr(g['yishi']['sets_per_doc'], 1)} | "
            f"{_iqr(g['dunhao_per_1000'], 1)} | {_iqr(g['four_char_ratio'], 2)} |")
    lines.append("")
    lines.append("### 开头与结语方式（占该文种文档的比例）")
    lines.append("")
    lines.append("| 文种 | 开头方式 | 结语方式 |")
    lines.append("|---|---|---|")
    for name, g in groups.items():
        op = "；".join(f"{k} {v['share']:.0%}" for k, v in list(g["openings"].items())[:4])
        cl = "；".join(f"{k} {v['share']:.0%}" for k, v in list(g["closings"].items())[:4])
        lines.append(f"| {name} | {op} | {cl} |")
    lines.append("")

    # 套话与评价词：每千字均值（出现该词的文档占比）
    cols = [g for g in TABLE_PHRASE_GROUPS if g in groups]
    lines.append("### 常见套话与评价词的千字频次（均值，括号内为出现该词的文档占比）")
    lines.append("")
    lines.append("| 词 | " + " | ".join(cols) + " |")
    lines.append("|---|" + "---:|" * len(cols))
    for word in TABLE_KEY_WORDS:
        cells = []
        for name in cols:
            row = None
            for rows in groups[name]["phrases_per_1000"].values():
                if word in rows:
                    row = rows[word]
            cells.append("—" if row is None else f"{row[1]:.2f}（{row[3]:.0%}）")
        lines.append(f"| {word} | " + " | ".join(cells) + " |")
    lines.append("")

    # 词表规则在真实公文里的命中
    overview: dict = baseline.get("style_patterns_overview", {})
    if overview:
        lines.append("### style_patterns.json 规则在真实公文里的命中（只列有命中的规则，按命中文档占比排序）")
        lines.append("")
        lines.append("| 规则 | 类别 | 级别 | 阈值（每千字） | 命中文档占比 | 超阈值文档占比 | 合计千字密度 | 命中占比最高的文种 |")
        lines.append("|---|---|---|---:|---:|---:|---:|---|")
        main_types = [g for g in groups if "/" not in g]
        for pid, r in sorted(overview.items(), key=lambda kv: (-kv[1]["doc_share"], kv[0])):
            tops = []
            for name in main_types:
                sp = groups[name]["style_patterns"].get(pid)
                if sp:
                    tops.append((sp["doc_share"], name))
            tops.sort(key=lambda x: (-x[0], x[1]))
            top_text = "、".join(f"{n} {v:.0%}" for v, n in tops[:2])
            thr = f"{r['threshold']:g}" if "threshold" in r else "—"
            over = _pct(r["over_threshold_share"]) if "over_threshold_share" in r else "—"
            lines.append(f"| {pid} | {r['category']} | {r['level']} | {thr} | {_pct(r['doc_share'])} | {over} | {r['per_1000_pooled']:.2f} | {top_text} |")
        lines.append("")
    return "\n".join(lines) + "\n"


# ---------------------------------------------------------------------------
# 命令行
# ---------------------------------------------------------------------------


def main(argv: Optional[Sequence[str]] = None) -> int:
    setup_console()
    parser = make_parser(
        "build_baseline.py",
        "从仓库外的公开公文语料重算文种量化基线（shared/style/baseline.json）。",
        "退出码：0 成功；1 --check 发现不一致；2 输入有误。",
    )
    parser.add_argument("--corpus", required=True, help="语料目录（含 sources.json 与按文种分子目录的 .txt）")
    parser.add_argument("--out", default=None, help="输出路径，默认 <仓库>/shared/style/baseline.json")
    parser.add_argument("--patterns", default=None, help="style_patterns.json 路径，默认 <仓库>/shared/scripts/style_patterns.json（只读）")
    parser.add_argument("--check", action="store_true", help="只比较不写入；结果与现有 baseline.json 不一致时返回 1")
    parser.add_argument("--print-tables", action="store_true", help="输出 Markdown 摘要表格后退出（不写文件）")
    args = parser.parse_args(argv)

    out_path = Path(args.out) if args.out else REPO_ROOT / DEFAULT_OUT_REL
    patterns_path = Path(args.patterns) if args.patterns else REPO_ROOT / DEFAULT_PATTERNS_REL
    try:
        docs, meta = load_corpus(Path(args.corpus))
    except CorpusError as exc:
        print(f"错误：{exc}")
        return 2
    if not patterns_path.is_file():
        print(f"提示：未找到词表 {patterns_path}，将跳过词表对照。")
    baseline = build_baseline(docs, meta, patterns_path if patterns_path.is_file() else None)

    if args.print_tables:
        print(render_tables(baseline))
        return 0

    text = dump_json(baseline)
    if args.check:
        current = out_path.read_text(encoding="utf-8") if out_path.is_file() else None
        if current is None:
            print(f"不一致：{out_path} 不存在。")
            return 1
        if current.replace("\r\n", "\n") != text:
            print(f"不一致：{out_path} 与按当前语料重算的结果不同。")
            return 1
        print("baseline.json 与语料一致。")
        return 0

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(text)
    print(f"已生成 {out_path}：{len(docs)} 篇，{len(baseline['groups'])} 组。")  # type: ignore[arg-type]
    return 0


if __name__ == "__main__":
    sys.exit(main())
