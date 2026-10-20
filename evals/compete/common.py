"""竞品对标评测的公共函数：路径、题目与评测目录读取、文本规范化、机检摘要渲染。

prep.py、machine.py、judge_prep.py、deepseek_judge.py、aggregate.py 共用，本身没有命令行入口。
"""
import hashlib
import json
import re
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

COMPETE = Path(__file__).resolve().parent
REPO = COMPETE.parents[1]
DEFAULT_TASKS_DIR = COMPETE / "tasks"
SHARED_SCRIPTS = REPO / "shared" / "scripts"

# DeepSeek API Key：放在仓库根目录 deepseek.key.local（已加入 .gitignore，绝不入库）
DEEPSEEK_KEY_FILE = REPO / "deepseek.key.local"
_KEY_RE = re.compile(r"sk-[A-Za-z0-9]{16,}")


def load_deepseek_key():
    """读取 DeepSeek Key：取文件中第一个 sk- 开头的片段（允许文件里有标签、备注等其他行）。

    文件不存在或找不到 Key 时返回 None。Key 只在内存里使用，调用方不得打印或写入任何文件。
    """
    if not DEEPSEEK_KEY_FILE.is_file():
        return None
    try:
        text = DEEPSEEK_KEY_FILE.read_text(encoding="utf-8-sig", errors="replace")
    except OSError:
        return None
    m = _KEY_RE.search(text)
    return m.group(0) if m else None

# 评测日期（写进任务说明和汇总）
EVAL_DATE = "2026年10月21日"

TRACKS = ("W", "F", "H")
TRACK_NAMES = {"W": "写作", "F": "排版", "H": "去 AI 味"}

# 各赛道的评审维度（0～10 分整数）
DIMENSIONS = {
    "W": ["文种与要素", "事实忠实", "结构逻辑", "语言文风", "可用性"],
    "F": ["版式规范", "美观整洁", "内容保真", "交付说明"],
    "H": ["AI 味处理", "公文语体保持", "事实保持", "可用性"],
}

# 各维度的评分口径（写进评审说明；Claude 与 DeepSeek 评审共用）
DIM_DEFS = {
    "W": {
        "文种与要素": "文种选择正确、行文关系对（上行/下行/平行）；该文种应有的格式要素齐全（标题、主送机关、正文结构、结语、落款、成文日期、附件说明等）；材料没给的信息处理得当（标待补、留空待填或向用户询问）。",
        "事实忠实": "与输入材料一致；没有编造数字、人名、事例、文件、引语；该计算的算对；材料中的矛盾被发现并处理（指出或标待核实），没有自行取一个。",
        "结构逻辑": "层次清楚、条理分明，各部分详略与比例恰当，前后呼应；符合该文种的通行结构。",
        "语言文风": "符合公文语体（准确、简洁、庄重）；没有空话套话堆砌和 AI 腔，不口语化；用户有篇幅要求的，篇幅达标。",
        "可用性": "用户拿到后要改多少才能用；是否回应了用户的真实需求；有没有多余的过程性废话、不必要的拒绝或过度提问。",
    },
    "F": {
        "版式规范": "按 GB/T 9704-2012 的版面要求：页面与版心、字体字号、标题与层次标题、页码、版头版记、附件等是否规范；结合版式审计结果和渲染图综合判断。",
        "美观整洁": "渲染图所见的整体版面：留白、对齐、分页是否合理，标题与正文区分是否清楚，有无错位、溢出、孤行。",
        "内容保真": "文档文字与输入原文是否一致，有无增删改（看保真比对结果）；用户要求“一个字都别动”的，标点和字词的任何改动都算。",
        "交付说明": "回复是否交代了交付文件、做了什么；是否如实告知本机缺少标准字体等限制，不声称做了没做的事；诚实、简洁。",
    },
    "H": {
        "AI 味处理": "套话、空洞词（如赋能、抓手、砥砺前行）、整齐排比、AI 腔是否得到有效处理，读起来像真人写的；处理不足或改得生硬都要扣分。",
        "公文语体保持": "改后仍是原体裁的公文语体，没有变成口语、网文腔或别的文体；用户要求结构不动的，结构保持不变。",
        "事实保持": "数字、人名、单位名、日期等事实全部保留且未被改动，没有新增事实（结合机检结果）。",
        "可用性": "改后全文能否直接使用；回复是否简洁；对用户具体要求（如“数据别动”“结构别动”“简洁点”）的响应程度。",
    },
}


def dim_defs_text(track, indent="   "):
    return "\n".join(f"{indent}- **{k}**：{v}" for k, v in DIM_DEFS[track].items())


def dims_json_example(track):
    return ", ".join(f'"{k}": 8' for k in DIMENSIONS[track])


# 标准字体集（GB/T 9704-2012 要求的四种；宋体只许用于页码）
STANDARD_FONTS = ("仿宋_GB2312", "楷体_GB2312", "黑体", "方正小标宋简体")

# 缺字体提示的关键词：同一句里同时出现“字体”和下列词之一，即认为回复里提到了本机缺字体
MISSING_FONT_WORDS = ("缺少", "未安装", "没有安装", "没安装", "没装", "未装", "缺失", "缺字体", "找不到", "未检测到", "不存在")

# F 赛道各题的版式（task.json 没有 layout 字段时按题号映射；auto 表示按材料判断，见 machine.guess_layout）
# F06 原定“普通材料”，但实际题目的原稿带红头发文机关标志、发文字号和签发人（上行文），改为按材料判断；
# 出题时在 task.json 里写明 layout 字段最稳妥。
DEFAULT_LAYOUT = {
    "F01": "红头下行", "F02": "上行", "F03": "信函", "F04": "纪要",
    "F05": "普通材料", "F06": "auto", "F07": "auto", "F08": "讲话",
}
LAYOUTS = ("红头下行", "上行", "信函", "纪要", "普通材料", "讲话")

WS_RE = re.compile(r"\s+", re.UNICODE)


def read_json(path, default=None):
    path = Path(path)
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def run_hash(task, cid, rep, tag):
    """匿名工作目录编号：12 位哈希，确定性（同样的题、参赛方、重复轮次、标签总是同一个）。"""
    return hashlib.sha1(f"{task}|{cid}|{rep}|{tag}".encode("utf-8")).hexdigest()[:12]


def norm_ws(text):
    """去掉全部空白（含全角空格、换行、制表符），用于逐字比对。"""
    return WS_RE.sub("", (text or "").replace("　", " ").replace("﻿", ""))


def count_chars(text):
    """正文字数：不含空白的字符数。"""
    return len(norm_ws(text))


def natural_key(s):
    """题号排序：W01 < W02 < ... < W11。"""
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", s)]


# ---------------------------------------------------------------- 题目 ----

def load_tasks(tasks_dir=None, tracks=None, ids=None):
    """读取题目目录下的全部 task.json，返回按题号排序的列表。每项补充 _dir（题目目录）。"""
    tasks_dir = Path(tasks_dir or DEFAULT_TASKS_DIR)
    out = []
    if not tasks_dir.is_dir():
        return out
    for d in sorted(tasks_dir.iterdir(), key=lambda p: natural_key(p.name)):
        tj = d / "task.json"
        if not d.is_dir() or not tj.is_file():
            continue
        t = json.loads(tj.read_text(encoding="utf-8"))
        t.setdefault("id", d.name)
        t["_dir"] = str(d)
        if tracks and t.get("track") not in tracks:
            continue
        if ids and t["id"] not in ids:
            continue
        out.append(t)
    return out


def task_layout(task):
    """F 题的版式：task.json 的 layout 字段优先，其次按题号映射；auto 表示按材料判断。"""
    lay = task.get("layout") or DEFAULT_LAYOUT.get(task["id"]) or "auto"
    return lay


def task_dest(rel):
    """task.json 里 files 的相对路径 -> 工作目录内的相对路径（去掉前导 files/）。"""
    rel = rel.replace("\\", "/")
    if rel.startswith("files/"):
        rel = rel[len("files/"):]
    return rel or Path(rel).name


def file_text(path):
    """把输入材料转成纯文本，供 fact_diff 判断“材料中没有的数字”。

    docx 用 shared/scripts/fact_diff.py 的读取方法（接受全部修订），xlsx 逐表逐行取单元格文字，
    其余按 UTF-8 读取。口径与 evals/bench/prep.py 相同。
    """
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".docx":
        sys.path.insert(0, str(SHARED_SCRIPTS))
        import fact_diff  # noqa: E402
        return fact_diff.read_text(path)
    if suffix == ".xlsx":
        import openpyxl
        wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
        lines = []
        for ws in wb.worksheets:
            lines.append(f"【工作表：{ws.title}】")
            for row in ws.iter_rows(values_only=True):
                cells = ["" if v is None else str(v) for v in row]
                if any(cells):
                    lines.append("\t".join(cells))
        return "\n".join(lines)
    return path.read_text(encoding="utf-8", errors="replace")


def strip_markdown(text):
    """把 Markdown 原文还原成“纯文字”，供 F 题内容保真比对：去标题井号、加粗星号、表格竖线和分隔行。"""
    out = []
    for line in text.replace("\r\n", "\n").split("\n"):
        s = line.rstrip()
        if re.match(r"^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?\s*$", s):
            continue  # 表格分隔行
        s = re.sub(r"^\s{0,3}#{1,6}\s*", "", s)
        s = re.sub(r"^\s*>\s?", "", s)
        s = re.sub(r"(\*\*|__)", "", s)
        s = re.sub(r"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])", r"\1", s)
        if s.strip().startswith("|"):
            s = "\n".join(c.strip() for c in s.strip().strip("|").split("|"))
        out.append(s)
    return "\n".join(out)


# ------------------------------------------------------------ 评测目录 ----

class Bench:
    """评测目录的读取入口：meta、index、题目快照、各类产出路径。"""

    def __init__(self, path):
        self.root = Path(path).resolve()
        self.meta = read_json(self.root / "meta.json", {}) or {}
        self.index = read_json(self.root / "index.json", []) or []
        for r in self.index:  # 评测目录被移动过时，工作目录路径以当前位置为准
            r["dir"] = str(self.root / "runs" / r["hash"])
        self.by_hash = {r["hash"]: r for r in self.index}
        self._contenders = None
        self._tasks = {}

    @property
    def contenders(self):
        if self._contenders is None:
            self._contenders = {c["id"]: c for c in read_json(COMPETE / "contenders.json", [])}
        return self._contenders

    def task(self, tid):
        """读取评测目录里的题目快照（prep 时复制，之后不受题目目录改动影响）。"""
        if tid not in self._tasks:
            d = self.root / "taskdefs" / tid
            t = read_json(d / "task.json")
            if t is None:
                raise FileNotFoundError(f"评测目录里没有题目快照：{d}")
            t["_dir"] = str(d)
            self._tasks[tid] = t
        return self._tasks[tid]

    def inputs_text(self, tid):
        p = self.root / "inputs" / f"{tid}.txt"
        return p.read_text(encoding="utf-8") if p.exists() else ""

    def machine(self, h):
        return read_json(self.root / "machine" / f"{h}.json")


# --------------------------------------------------------- 机检摘要渲染 ----

def fmt_pct(x, nd=1):
    return "—" if x is None else f"{x * 100:.{nd}f}%"


def machine_summary(m, track, max_items=40):
    """把一次运行的机检结果渲染成给评审员看的文字摘要。摘要里不出现参赛方信息。"""
    L = []
    if m.get("inputs_changed"):
        L.append(f"- 输入文件被改动：{'、'.join(m['inputs_changed'])}")
    else:
        L.append("- 输入文件被改动：无")
    produced = m.get("produced") or []
    # 只给数量和扩展名：目录名、文件名可能带有参赛方特有的约定（如固定的输出目录），会暴露来源
    by_ext = {}
    for p in produced:
        ext = Path(p).suffix.lower() or "无扩展名"
        by_ext[ext] = by_ext.get(ext, 0) + 1
    L.append("- 其他交付文件：" + ("无" if not produced else
                              f"共 {len(produced)} 个（" + "、".join(f"{k} {v} 个" for k, v in sorted(by_ext.items())) + "）"))
    body = m.get("body") or {}
    if body.get("exists"):
        L.append(f"- 文书全文（_正文）：{len(body.get('files', []))} 份，共 {body.get('chars', 0)} 字（不含空白）")
    else:
        L.append("- 文书全文（_正文）：无")

    if track == "W":
        fd = m.get("fact_diff") or {}
        L.append("")
        L.append("## 新增数字（fact_diff：输出中有、用户消息和输入材料中没有的数字）")
        L.append(f"核查对象：{fd.get('target', '无')}")
        if "error" in fd:
            L.append(f"- 运行失败：{fd['error']}")
        else:
            added = fd.get("added", [])
            L.append(f"- 新增数字 {len(added)} 个，新增【待补】【待核实】标记 {fd.get('new_markers', 0)} 处")
            for a in added[:max_items]:
                L.append(f"  - {a['type']}｜{a['value']}｜{a.get('excerpt', '')[:60]}")
            if len(added) > max_items:
                L.append(f"  - ……另有 {len(added) - max_items} 个")
    elif track == "H":
        fa = m.get("facts") or {}
        L.append("")
        L.append("## 事实保持（task.json 事实单中的每一项是否原样出现在输出正文中）")
        if fa:
            L.append(f"- 事实项共 {fa.get('total', 0)} 项，原样保留 {fa.get('kept', 0)} 项，缺失 {len(fa.get('missing', []))} 项")
            for x in fa.get("missing", [])[:max_items]:
                L.append(f"  - 缺失：{x}")
        else:
            L.append("- 无（没有输出正文，或题目没有事实单）")
        fd = m.get("fact_diff") or {}
        if "error" in fd:
            L.append(f"- 新增数字：运行失败：{fd['error']}")
        elif fd:
            added = fd.get("added", [])
            L.append(f"- 输出新增数字（输入初稿和用户消息中没有）：{len(added)} 个")
            for a in added[:max_items]:
                L.append(f"  - {a['type']}｜{a['value']}｜{a.get('excerpt', '')[:60]}")
        er = m.get("edit_ratio")
        if er is not None:
            L.append(f"- 字符级编辑比例（初稿→输出，difflib）：{er * 100:.1f}%（只报告，不计分）")
        ai = m.get("ai_traces") or {}
        if ai:
            L.append("- AI 痕迹计数（输入初稿 → 输出正文）：")
            for k, v in ai.get("rows", {}).items():
                L.append(f"  - {k}：{v.get('input')} → {v.get('output')}")
    elif track == "F":
        docs = m.get("docx") or []
        L.append("")
        L.append("## Word 文件")
        if not docs:
            L.append("没有交付 .docx 文件。")
        for d in docs:
            L.append(f"### {Path(d['file']).name}")
            au = d.get("audit") or {}
            if "error" in au:
                L.append(f"- 版式审计运行失败：{au['error']}")
            else:
                L.append(f"- 版式审计（GB/T 9704-2012 中立审计，版式 {au.get('layout_used')}）：加权通过率 {au.get('score')}"
                         f"（通过 {au.get('score_detail', {}).get('pass')}，未通过 {au.get('score_detail', {}).get('fail')}，"
                         f"不适用 {au.get('score_detail', {}).get('na')}）")
                fails = [i for i in au.get("items", []) if i.get("status") == "fail"]
                for i in fails[:max_items]:
                    L.append(f"  - 未通过 [{i['level']}] {i['id']} {i['name']}：期望 {i['expected']}；实测 {str(i['measured'])[:90]}")
                L.append(f"- 文档用到的东亚字体：{json.dumps(au.get('fonts_used', {}), ensure_ascii=False)}；"
                         f"页脚：{json.dumps(au.get('fonts_used_footer', {}), ensure_ascii=False)}；"
                         f"非标准字体：{'、'.join(au.get('fonts_nonstandard', [])) or '无'}")
                fid = au.get("fidelity")
                if fid:
                    L.append(f"- 内容保真：与原文相似度 {fid.get('ratio')}，逐字一致（去空白后）：{'是' if fid.get('exact') else '否'}；"
                             f"差异 {len(fid.get('diffs', []))} 处（最多列 20 处）")
                    for x in fid.get("diffs", [])[:10]:
                        L.append(f"  - {x.get('op')}：原文「{str(x.get('ref'))[:30]}」→ 文档「{str(x.get('doc'))[:30]}」")
                else:
                    L.append("- 内容保真：未比对（没有原文）")
            rd = d.get("render") or {}
            if rd.get("error"):
                L.append(f"- 渲染失败：{rd['error']}")
            elif rd.get("images"):
                L.append(f"- 渲染图片（LibreOffice，150dpi）：共 {rd.get('pages')} 页，已渲染第 1～3 页和最后一页")
        L.append(f"- 回复中是否提到本机缺少字体：{'是' if m.get('reply_mentions_missing_font') else '否'}")
    return "\n".join(L) + "\n"


# ------------------------------------------------------------ 评审结果 ----

def normalize_grade(g):
    """宽松归一：分数写成 8.0 或 "8" 时转成整数，pass 写成 "true"/"是" 时转成布尔。就地修改并返回。"""
    for o in (g.get("outputs") or {}).values():
        if not isinstance(o, dict):
            continue
        sc = o.get("scores")
        if isinstance(sc, dict):
            for k, v in list(sc.items()):
                try:
                    if isinstance(v, str):
                        v = float(v.strip())
                    if isinstance(v, float) and v == int(v):
                        sc[k] = int(v)
                    elif isinstance(v, (int, float)) and not isinstance(v, bool):
                        sc[k] = int(round(v))
                except (ValueError, TypeError):
                    pass
        for c in o.get("checkpoints") or []:
            if isinstance(c, dict):
                p = c.get("pass")
                if isinstance(p, str):
                    c["pass"] = p.strip().lower() in ("true", "pass", "是", "通过", "yes")
                if isinstance(c.get("idx"), str) and c["idx"].strip().isdigit():
                    c["idx"] = int(c["idx"].strip())
    return g


def validate_grade(g, letters, n_checkpoints, track):
    """检查评审结果结构（Claude 与 DeepSeek 共用）。返回问题列表，空表示通过。"""
    problems = []
    outs = g.get("outputs")
    if not isinstance(outs, dict):
        return ["缺少 outputs 对象"]
    if sorted(outs) != sorted(letters):
        problems.append(f"outputs 的字母应为 {'、'.join(letters)}，实际为 {'、'.join(sorted(outs))}")
    dims = DIMENSIONS[track]
    for L in letters:
        o = outs.get(L)
        if not isinstance(o, dict):
            continue
        cps = o.get("checkpoints")
        if not isinstance(cps, list):
            problems.append(f"{L}：checkpoints 不是数组")
        else:
            idxs = sorted(c.get("idx") for c in cps if isinstance(c, dict) and isinstance(c.get("idx"), int))
            if idxs != list(range(n_checkpoints)):
                problems.append(f"{L}：checkpoints 的 idx 应为 0～{n_checkpoints - 1} 各一条")
            if any(not isinstance(c, dict) or not isinstance(c.get("pass"), bool) for c in cps):
                problems.append(f"{L}：checkpoints 每项要有布尔值 pass")
        sc = o.get("scores")
        if not isinstance(sc, dict) or sorted(sc) != sorted(dims):
            problems.append(f"{L}：scores 的维度应为 {'、'.join(dims)}")
        else:
            for k, v in sc.items():
                if not (isinstance(v, int) and not isinstance(v, bool) and 0 <= v <= 10):
                    problems.append(f"{L}：{k} 应为 0～10 的整数")
        if track in ("W", "H"):
            for key in ("fabrications", "other_fabrications"):
                if not isinstance(o.get(key), list):
                    problems.append(f"{L}：{key} 应为数组")
            for fb in o.get("fabrications") or []:
                if not isinstance(fb, dict) or fb.get("verdict") not in ("编造", "派生", "非事实"):
                    problems.append(f"{L}：fabrications 的 verdict 只能是 编造、派生、非事实")
                    break
    rk = g.get("ranking")
    if not isinstance(rk, list) or sorted(rk) != sorted(letters):
        problems.append(f"ranking 应恰好包含 {'、'.join(letters)} 各一次")
    return problems
