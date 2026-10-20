"""竞品对标评测机检：对每次运行的产出做机器核查，并生成给评审员看的脱敏副本。

用法：
    python evals/compete/machine.py --bench <评测目录> [--only 哈希|题号|参赛方,...] [--no-render] [--jobs 3]

对每个工作目录，写 <评测目录>/machine/<哈希>.json，内容按赛道：
    通用   输入文件是否被改动、交付文件清单、_正文 是否存在及字数、_回复/_过程 是否存在
    W      fact_diff：输出中有、用户消息和输入材料（台账）中没有的数字
    H      事实保持（task.json 的 facts 逐项是否原样出现）、输出新增数字、字符级编辑比例、
           AI 痕迹计数（lieflat-less-ai-tone 的两个统计脚本 + 本仓库 gongwen_lint，对输入初稿同样统计一遍作基线）
    F      docx_audit 版式审计与内容保真、LibreOffice 渲染第 1～3 页和最后一页为 PNG（150dpi）、
           docx 中出现的东亚字体、回复里是否提到本机缺字体
并生成：
    sanitized/<哈希>/    _回复.md、_正文/*、机检摘要.md，其中能识别参赛方的词已替换为“[工具]”（评审员只看这里）
    render/<哈希>/       F 赛道的 PDF 与页面 PNG

PDF 转图片优先用已安装的 PyMuPDF；没有时自动用 `uv run --with pymupdf` 调用本脚本的 --render-pdf 模式。
"""
import argparse
import difflib
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from common import (Bench, COMPETE, MISSING_FONT_WORDS, SHARED_SCRIPTS, count_chars, machine_summary,  # noqa: E402
                    norm_ws, read_json, strip_markdown, task_layout, write_json)

LINT = SHARED_SCRIPTS / "gongwen_lint.py"
FACT = SHARED_SCRIPTS / "fact_diff.py"
AUDIT = COMPETE / "docx_audit.py"
ENV = {**os.environ, "PYTHONIOENCODING": "utf-8"}
HARNESS_FILES = {"_回复.md", "_过程.md"}
BODY_SUFFIXES = (".txt", ".md")
SOFFICE_CANDIDATES = [r"C:\Program Files\LibreOffice\program\soffice.exe",
                      r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]
RENDER_DPI = 150
UNKNOWN = "（无法判断）"


# ------------------------------------------------------------ 子进程工具 ----

def run_capture(cmd, timeout=180, cwd=None):
    """运行命令，返回 (退出码, 标准输出, 标准错误)；超时返回 (-1, '', '超时')。"""
    try:
        p = subprocess.run([str(x) for x in cmd], capture_output=True, text=True, encoding="utf-8",
                           errors="replace", env=ENV, timeout=timeout, cwd=cwd)
        return p.returncode, p.stdout, p.stderr
    except subprocess.TimeoutExpired:
        return -1, "", "超时"
    except OSError as e:
        return -2, "", str(e)


def run_json(cmd, timeout=180):
    rc, out, err = run_capture([sys.executable, *cmd], timeout=timeout)
    try:
        return json.loads(out)
    except json.JSONDecodeError:
        return {"error": ((err or out) or f"退出码 {rc}")[-400:]}


# --------------------------------------------------------- fact_diff / lint ----

def fact_ledger(inputs_txt, path):
    """台账模式：输出中有、台账（用户消息 + 输入材料）中没有的数字。"""
    r = run_json([FACT, "--ledger", inputs_txt, path, "--json"])
    if "error" in r:
        return r
    added = [{"type": a["type"], "value": a["value"],
              "excerpt": (a.get("locations") or [{}])[0].get("excerpt", "")}
             for a in r.get("added", []) if a["type"] != "标记"]
    markers = [a["value"] for a in r.get("added", []) if a["type"] == "标记"]
    return {"added": added, "new_markers": len(markers)}


def fact_compare(draft, path):
    """普通模式：初稿有而输出没有的事实（缺失），仅 H 题参考。"""
    r = run_json([FACT, draft, path, "--json"])
    if "error" in r:
        return r
    return {"missing": [{"type": a["type"], "value": a["value"]} for a in r.get("missing", [])]}


def lint_stats(path, doc_type=None):
    cmd = [LINT, path, "--json"]
    if doc_type:
        cmd += ["--doc-type", doc_type]
    r = run_json(cmd)
    if "error" in r:
        return r
    cats = {}
    for i in r.get("issues", []):
        cats[i.get("category", "?")] = cats.get(i.get("category", "?"), 0) + 1
    s = r.get("summary", {})
    return {"doc_type": r.get("doc_type"), "P0": s.get("P0", 0), "P1": s.get("P1", 0), "P2": s.get("P2", 0),
            "total": s.get("P0", 0) + s.get("P1", 0) + s.get("P2", 0), "by_category": cats}


def lint_doc_type(genre):
    """task.json 的 genre 里含有 gongwen_lint 认识的文书类型时，传给它；否则让它自己推断。"""
    if not genre:
        return None
    cands = ["对照检查材料", "党课讲稿", "心得体会", "研讨发言", "会议记录", "议题材料", "实施方案", "述职报告", "调研报告",
             "经验材料", "汇报材料", "工作要点", "讲话稿", "新闻稿", "简报", "整改报告", "总结", "计划", "信息", "通知",
             "报告", "请示", "函", "纪要"]
    for c in cands:
        if c in genre:
            return c
    return None


# ----------------------------------------------------- lieflat 去 AI 味脚本 ----

TRANS_ROW = re.compile(r"^(\S+)\s+(\d+\.\d+)\s+(\d+)\s+(\d+/\d+)\s")
STRUCT_ROW = re.compile(r"^(\S+)\s+([\d.]+)\s+([\d.]+)\s+(inf|[\d.]+)×")


def lieflat_stats(bench_root, text, tmp_dir):
    """用参赛方副本里的 check-translationese.py、check-structure.py 统计一段文字。

    两个脚本都按“目录下的 *.md”读取，所以把文字写成临时目录里的一个 md 文件再调用。
    """
    sdir = Path(bench_root) / "skills" / "lieflat-less-ai-tone" / "scripts"
    f1, f2 = sdir / "check-translationese.py", sdir / "check-structure.py"
    if not (f1.is_file() and f2.is_file()):
        return {"available": False, "reason": "评测目录里没有 lieflat-less-ai-tone 的 skill 副本"}
    tmp_dir = Path(tmp_dir)
    if tmp_dir.exists():
        shutil.rmtree(tmp_dir)
    tmp_dir.mkdir(parents=True)
    (tmp_dir / "text.md").write_text(text, encoding="utf-8")
    out = {"available": True, "translationese": {}, "baseline": {}, "structure": {}}
    rc, so, se = run_capture([sys.executable, f1, tmp_dir], timeout=120)
    if rc != 0:
        out["error_translationese"] = (se or so)[-300:]
    for line in so.splitlines():
        m = TRANS_ROW.match(line)
        if m:
            tgt = out["baseline"] if m.group(1).startswith("[基准]") else out["translationese"]
            tgt[m.group(1)] = int(m.group(3))
    rc, so, se = run_capture([sys.executable, f2, "--human", tmp_dir, "--ai", tmp_dir], timeout=120)
    if rc != 0:
        out["error_structure"] = (se or so)[-300:]
    m = re.search(r"AI\s+(\d+)\s+段", so)
    out["paragraphs"] = int(m.group(1)) if m else None
    for line in so.splitlines():
        m = STRUCT_ROW.match(line)
        if m:
            out["structure"][m.group(1)] = float(m.group(2))
    out["translationese_total"] = sum(out["translationese"].values())
    out["baseline_total"] = sum(out["baseline"].values())
    shutil.rmtree(tmp_dir, ignore_errors=True)
    return out


# AI 痕迹计数的行标签。给评审员看的摘要里必须中立：不出现任何参赛方或项目的名字，
# 否则带有参赛方自称的行在脱敏时会被替换成“[工具]”，反而暴露来源。
# 统计脚本一 = lieflat-less-ai-tone 自带的 check-translationese.py、check-structure.py；统计脚本二 = 本仓库 gongwen_lint。
TRACE_LABELS = {
    "trans": "统计脚本一：翻译腔标记合计（20 类，次）",
    "base": "统计脚本一：对照标记合计（段首序数词、不是…而是、破折号、提示性冒号，次）",
    "iso2": "统计脚本一：连续两句同构（每百段）",
    "iso3": "统计脚本一：连续三句同构（每百段）",
    "zero": "统计脚本一：段首零主语评论（占非首段 %）",
    "lint_total": "统计脚本二：问题总数（P0+P1+P2）",
    "lint_ai": "统计脚本二：AI味类问题数",
    "lint_def": "统计脚本二：防御性类问题数",
}


def ai_trace_rows(inp, out):
    """把输入初稿和输出正文的各项统计整理成“输入→输出”的对照行。"""
    rows = {}

    def add(key, a, b):
        if a is not None or b is not None:
            rows[TRACE_LABELS[key]] = {"input": a, "output": b}

    li, lo = inp.get("lieflat") or {}, out.get("lieflat") or {}
    if li.get("available") and lo.get("available"):
        add("trans", li.get("translationese_total"), lo.get("translationese_total"))
        add("base", li.get("baseline_total"), lo.get("baseline_total"))
        for key, k in (("iso2", "连续两句同构/百段"), ("iso3", "连续三句同构/百段"), ("zero", "段首零主语评论/非首段%")):
            add(key, li.get("structure", {}).get(k), lo.get("structure", {}).get(k))
    gi, go = inp.get("lint") or {}, out.get("lint") or {}
    if "error" not in gi and "error" not in go and gi and go:
        add("lint_total", gi.get("total"), go.get("total"))
        add("lint_ai", gi.get("by_category", {}).get("AI味", 0), go.get("by_category", {}).get("AI味", 0))
        add("lint_def", gi.get("by_category", {}).get("防御性", 0), go.get("by_category", {}).get("防御性", 0))
    return rows


# ------------------------------------------------------------- 事实保持 ----

def norm_fact(s):
    return norm_ws(unicodedata.normalize("NFKC", s or "")).lower()


def flatten_facts(facts):
    """task.json 的 facts 可能是 {"numbers": [...], "names": [...]} 之类，逐组摊平。"""
    groups = {}
    if isinstance(facts, dict):
        for k, v in facts.items():
            if isinstance(v, list):
                groups[k] = [str(x) for x in v if str(x).strip()]
    elif isinstance(facts, list):
        groups["facts"] = [str(x) for x in facts if str(x).strip()]
    return groups


def check_facts(facts, body_text):
    groups = flatten_facts(facts)
    if not groups:
        return None
    nb = norm_fact(body_text)
    total = kept = 0
    missing, by_group = [], {}
    for g, items in groups.items():
        gk = 0
        for it in items:
            total += 1
            if norm_fact(it) in nb:
                kept += 1
                gk += 1
            else:
                missing.append(it)
        by_group[g] = {"total": len(items), "kept": gk}
    return {"total": total, "kept": kept, "missing": missing, "by_group": by_group}


def edit_stats(a, b):
    a, b = norm_ws(strip_markdown(a)), norm_ws(strip_markdown(b))
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    ratio = sm.ratio()
    ins = dele = 0
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag in ("replace", "delete"):
            dele += i2 - i1
        if tag in ("replace", "insert"):
            ins += j2 - j1
    return {"edit_ratio": round(1 - ratio, 4), "deleted": dele, "inserted": ins, "orig_chars": len(a), "new_chars": len(b)}


# ---------------------------------------------------------------- F 赛道 ----

def find_soffice():
    for c in [os.environ.get("SOFFICE")] + SOFFICE_CANDIDATES + [shutil.which("soffice")]:
        if c and Path(c).is_file():
            return str(c)
    return None


def docx_to_pdf(docx, outdir, profile_dir, timeout=240):
    soffice = find_soffice()
    if not soffice:
        return None, "找不到 LibreOffice（soffice.exe）"
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    cmd = [soffice, f"-env:UserInstallation={Path(profile_dir).resolve().as_uri()}", "--headless", "--norestore",
           "--convert-to", "pdf", "--outdir", str(outdir), str(docx)]
    rc, so, se = run_capture(cmd, timeout=timeout)
    pdf = outdir / (Path(docx).stem + ".pdf")
    if not pdf.is_file():
        return None, f"转换失败（退出码 {rc}）：{(se or so)[-200:]}"
    return pdf, None


def import_mupdf():
    """导入 PyMuPDF：新版包名是 pymupdf，旧版是 fitz（新版再 import fitz 会打印弃用警告）。"""
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    return fitz


def render_pdf_inproc(pdf, outdir, dpi):
    """用 PyMuPDF 把第 1～3 页和最后一页渲染为 PNG。"""
    fitz = import_mupdf()
    doc = fitz.open(str(pdf))
    n = doc.page_count
    pages = list(range(1, min(3, n) + 1))
    names = {p: f"page-{p:02d}.png" for p in pages}
    if n > 3:
        pages.append(n)
        names[n] = "page-last.png"
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    zoom = dpi / 72.0
    imgs = []
    for p in pages:
        pix = doc[p - 1].get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        pix.save(str(outdir / names[p]))
        imgs.append({"page": p, "file": names[p]})
    doc.close()
    return {"pages": n, "images": imgs}


def render_pdf(pdf, outdir, dpi=RENDER_DPI):
    try:
        import_mupdf()
        return render_pdf_inproc(pdf, outdir, dpi)
    except ImportError:
        pass
    uv = shutil.which("uv")
    if not uv:
        return {"error": "没有安装 PyMuPDF，也找不到 uv"}
    rc, so, se = run_capture([uv, "run", "--quiet", "--with", "pymupdf", "python", __file__, "--render-pdf",
                              str(pdf), str(outdir), str(dpi)], timeout=300)
    try:
        return json.loads(so.strip().splitlines()[-1])
    except Exception:
        return {"error": f"渲染失败（退出码 {rc}）：{(se or so)[-200:]}"}


def docx_declared_fonts(path):
    """扫描 docx 里全部部件，收集写过的东亚字体名（包括样式、页脚、编号定义里声明的）。"""
    import zipfile
    fonts = set()
    try:
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if n.startswith("word/") and n.endswith(".xml"):
                    xml = z.read(n).decode("utf-8", errors="replace")
                    fonts.update(re.findall(r'w:eastAsia="([^"]+)"', xml))
    except Exception:
        pass
    return sorted(fonts)


def reply_font_mentions(reply):
    """回复里是否提到本机缺少字体：同一句里有“字体”且有“缺少/未安装/没有安装……”之类的词。"""
    hits = []
    for s in re.split(r"[。！？\n]", reply or ""):
        low = s.lower()
        if ("字体" in s and any(w in s for w in MISSING_FONT_WORDS)) or \
                ("font" in low and any(w in low for w in ("not installed", "missing", "unavailable", "not found", "isn't installed"))):
            hits.append(s.strip()[:120])
    return hits


def guess_layout(text):
    """task.json 没有明确版式（auto）时，按材料判断。"""
    t = text or ""
    if re.search(r"〔\s*\d{4}\s*〕\s*\d+\s*号", t):
        return "上行" if "签发人" in t else "红头下行"
    first = next((ln.strip() for ln in t.splitlines() if ln.strip()), "")
    if first.endswith("纪要"):
        return "纪要"
    if "同志们" in t and "讲话" in t:
        return "讲话"
    return "普通材料"


def reference_text(bench, task):
    """F 题的内容保真基准：reference/_原文.txt；没有时取输入材料本身（md 去掉标记，docx 取文字）。"""
    d = Path(bench.root) / "taskdefs" / task["id"]
    ref = d / "reference" / "_原文.txt"
    if ref.is_file():
        return ref.read_text(encoding="utf-8", errors="replace")
    for rel in task.get("files", []):
        p = d / rel
        if not p.is_file():
            continue
        if p.suffix.lower() in (".txt", ".md"):
            t = p.read_text(encoding="utf-8", errors="replace")
            return strip_markdown(t) if p.suffix.lower() == ".md" else t
        if p.suffix.lower() == ".docx":
            from common import file_text
            return file_text(p)
    return None


def pick_primary(docxs, reply_text):
    """多个 docx 时选主交付物：回复里提到文件名的优先，其次最近修改的。"""
    if len(docxs) == 1:
        return docxs[0]
    named = [p for p in docxs if p.name in (reply_text or "")]
    pool = named or docxs
    return max(pool, key=lambda p: p.stat().st_mtime)


def process_f(bench, row, task, wdir, produced, reply_text, opts, res):
    docxs = [wdir / p for p in produced if p.lower().endswith(".docx") and not Path(p).name.startswith("~$")]
    # 用户的原 docx 被原地修改时（说明里要求另存新文件，但个别 skill 习惯原地改），被改过的原件也算交付物
    changed_in = {(wdir / f).resolve() for f in res.get("inputs_changed", []) if f.lower().endswith(".docx")}
    for f in res.get("inputs_changed", []):
        if f.lower().endswith(".docx") and (wdir / f).is_file() and (wdir / f) not in docxs:
            docxs.append(wdir / f)
    res["docx"] = []
    ref_text = reference_text(bench, task)
    layout = task_layout(task)
    if layout == "auto":
        layout = guess_layout(ref_text)
    res["layout"] = layout
    tmp = Path(bench.root) / "machine" / "_tmp"
    tmp.mkdir(parents=True, exist_ok=True)
    ref_file = None
    if ref_text is not None:
        ref_file = tmp / f"{row['hash']}-ref.txt"
        ref_file.write_text(ref_text, encoding="utf-8")
    primary = pick_primary(docxs, reply_text) if docxs else None
    for d in docxs:
        rec = {"file": d.relative_to(wdir).as_posix(), "primary": d == primary}
        if d.resolve() in changed_in:
            rec["modified_input"] = True
        cmd = [AUDIT, d, "--layout", layout, "--json"]
        if ref_file:
            cmd += ["--ref", ref_file]
        rec["audit"] = run_json(cmd, timeout=180)
        rec["fonts_declared"] = docx_declared_fonts(d)
        if d == primary and not opts.no_render:
            rdir = Path(bench.root) / "render" / row["hash"]
            if rdir.exists():
                shutil.rmtree(rdir, ignore_errors=True)
            pdf, err = docx_to_pdf(d, rdir, tmp / f"lo-{row['hash']}")
            shutil.rmtree(tmp / f"lo-{row['hash']}", ignore_errors=True)
            if err:
                rec["render"] = {"error": err}
            else:
                r = render_pdf(pdf, rdir)
                if "error" in r:
                    rec["render"] = {"error": r["error"]}
                else:
                    rec["render"] = {"pdf": f"render/{row['hash']}/{pdf.name}", "pages": r["pages"],
                                     "images": [{"page": i["page"], "path": f"render/{row['hash']}/{i['file']}"}
                                                for i in r["images"]]}
        res["docx"].append(rec)
    if ref_file and ref_file.exists():
        ref_file.unlink()
    hits = reply_font_mentions(reply_text)
    res["reply_mentions_missing_font"] = bool(hits)
    res["reply_font_mentions"] = hits[:3]
    return res


# ------------------------------------------------------------------ 脱敏 ----

def path_variants(p):
    s = str(p)
    v = {s, s.replace("\\", "/"), s.replace("\\", "\\\\")}
    m = re.match(r"^([A-Za-z]):[\\/](.*)$", s)
    if m:
        v.add("/" + m.group(1).lower() + "/" + m.group(2).replace("\\", "/"))
    return sorted(v, key=len, reverse=True)


def build_scrubber(bench, cid, wdir):
    """返回 scrub(text)：把能识别参赛方的词和路径替换掉。"""
    c = bench.contenders.get(cid, {})
    terms = set(c.get("identify_terms", []))
    sk = Path(bench.root) / "skills" / cid
    if c.get("kind") == "suite" and sk.is_dir():
        terms.update(d.name for d in sk.iterdir() if d.is_dir())
    if c.get("skill_name"):
        terms.add(c["skill_name"])
    terms = sorted((t for t in terms if t and len(t) >= 3), key=len, reverse=True)
    term_res = [re.compile(re.escape(t), re.I) for t in terms]
    path_subs = [(str(wdir), "[工作目录]"), (str(sk), "[技能目录]"), (str(bench.root), "[评测目录]")]
    path_res = [(re.compile(re.escape(v), re.I), repl) for p, repl in path_subs for v in path_variants(p)]

    def scrub(text):
        for rx, repl in path_res:
            text = rx.sub(repl, text)
        for rx in term_res:
            text = rx.sub("[工具]", text)
        return text
    return scrub


def write_sanitized(bench, row, wdir, bodies, scrub, summary_text):
    sdir = Path(bench.root) / "sanitized" / row["hash"]
    if sdir.exists():
        shutil.rmtree(sdir)
    sdir.mkdir(parents=True)
    files = []
    reply = wdir / "_回复.md"
    if reply.is_file():
        (sdir / "_回复.md").write_text(scrub(reply.read_text(encoding="utf-8", errors="replace")), encoding="utf-8")
        files.append("_回复.md")
    if bodies:
        (sdir / "_正文").mkdir()
        for b in bodies:
            (sdir / "_正文" / b.name).write_text(scrub(b.read_text(encoding="utf-8", errors="replace")), encoding="utf-8")
            files.append(f"_正文/{b.name}")
    (sdir / "机检摘要.md").write_text(scrub(summary_text), encoding="utf-8")
    files.append("机检摘要.md")
    return files


# ---------------------------------------------------------------- 主流程 ----

def draft_of(bench, task):
    """H 题的输入初稿：优先文件名为“初稿”的材料，否则取第一份；读题目快照里的原件。"""
    rel = next((f for f in task.get("files", []) if Path(f).stem == "初稿"), (task.get("files") or [None])[0])
    p = Path(bench.root) / "taskdefs" / task["id"] / rel if rel else None
    return p, (p.read_text(encoding="utf-8", errors="replace") if p and p.is_file() else "")


def baseline_for(bench, task):
    """输入初稿的 AI 痕迹基线，缓存在 machine/_baseline-<题号>.json。"""
    base_path = Path(bench.root) / "machine" / f"_baseline-{task['id']}.json"
    base = read_json(base_path)
    if base is None:
        _, draft_text = draft_of(bench, task)
        # 初稿是 Markdown，输出正文是纯文本：先去掉标记再统计，否则“残留”类问题会让基线虚高
        draft_text = strip_markdown(draft_text)
        tmp = Path(bench.root) / "machine" / "_tmp"
        tmp.mkdir(parents=True, exist_ok=True)
        dtmp = tmp / f"{task['id']}-draft.txt"
        dtmp.write_text(draft_text, encoding="utf-8")
        base = {"lint": lint_stats(dtmp, lint_doc_type(task.get("genre"))),
                "lieflat": lieflat_stats(bench.root, draft_text, tmp / f"{task['id']}-lf-draft")}
        write_json(base_path, base)
        dtmp.unlink(missing_ok=True)
    return base


def process(bench, row, hashes, opts):
    wdir = Path(row["dir"])
    task = bench.task(row["task"])
    track = row["track"]
    res = {"hash": row["hash"], "task": row["task"], "track": track, "contender": row["contender"],
           "rep": row["rep"], "dir": str(wdir)}
    res["has_reply"] = (wdir / "_回复.md").is_file()
    res["has_log"] = (wdir / "_过程.md").is_file()
    reply_text = (wdir / "_回复.md").read_text(encoding="utf-8", errors="replace") if res["has_reply"] else ""
    log = (wdir / "_过程.md").read_text(encoding="utf-8", errors="replace") if res["has_log"] else ""
    res["model_line"] = log.splitlines()[0][:80] if log else ""

    # 输入材料是否被改动
    changed, input_paths = [], set()
    for f in row["files"]:
        p = wdir / f
        input_paths.add(os.path.normcase(str(p)))
        key = f"runs/{row['hash']}/{f}"
        if not p.is_file():
            changed.append(f"{f}（已不存在）")
        elif key in hashes and hashlib.sha256(p.read_bytes()).hexdigest() != hashes[key]:
            changed.append(f)
    res["inputs_changed"] = changed

    # 交付文件清单（不含输入材料、运行记录和 _正文）
    produced = []
    for p in sorted(wdir.rglob("*")):
        if not p.is_file() or os.path.normcase(str(p)) in input_paths or p.name in HARNESS_FILES:
            continue
        rel = p.relative_to(wdir).as_posix()
        if rel.startswith("_正文/") or "__pycache__" in rel or p.suffix == ".pyc":
            continue
        produced.append(rel)
    res["produced"] = produced

    bodies = sorted(p for p in (wdir / "_正文").glob("*") if p.is_file() and p.suffix.lower() in BODY_SUFFIXES) \
        if (wdir / "_正文").is_dir() else []
    body_text = "\n\n".join(b.read_text(encoding="utf-8", errors="replace") for b in bodies)
    res["body"] = {"exists": bool(bodies), "files": [b.name for b in bodies], "chars": count_chars(body_text)}

    inputs_txt = Path(bench.root) / "inputs" / f"{row['task']}.txt"
    tmp = Path(bench.root) / "machine" / "_tmp"
    tmp.mkdir(parents=True, exist_ok=True)

    if track == "W":
        target = bodies if bodies else ([wdir / "_回复.md"] if res["has_reply"] else [])
        res["fact_diff"] = {"target": "、".join(t.relative_to(wdir).as_posix() for t in target) or "无", "added": [],
                            "new_markers": 0}
        seen = set()
        for t in target:
            r = fact_ledger(inputs_txt, t)
            if "error" in r:
                res["fact_diff"]["error"] = r["error"]
                continue
            for a in r["added"]:
                if (a["type"], a["value"]) not in seen:
                    seen.add((a["type"], a["value"]))
                    res["fact_diff"]["added"].append(a)
            res["fact_diff"]["new_markers"] += r["new_markers"]
    elif track == "H":
        draft_path, draft_text = draft_of(bench, task)
        res["facts"] = check_facts(task.get("facts"), body_text) if bodies else None
        if bodies:
            merged = tmp / f"{row['hash']}-body.txt"
            merged.write_text(body_text, encoding="utf-8")
            fd = fact_ledger(inputs_txt, merged)
            fd["target"] = "、".join(b.name for b in bodies)
            if draft_path and draft_path.is_file():
                cmp_ = fact_compare(draft_path, merged)
                fd["missing_vs_draft"] = cmp_.get("missing", []) if "error" not in cmp_ else {"error": cmp_["error"]}
            res["fact_diff"] = fd
            res["edit_stats"] = edit_stats(draft_text, body_text)
            res["edit_ratio"] = res["edit_stats"]["edit_ratio"]
            dt = lint_doc_type(task.get("genre"))
            out_stats = {"lint": lint_stats(merged, dt),
                         "lieflat": lieflat_stats(bench.root, body_text, tmp / f"{row['hash']}-lf")}
            base = baseline_for(bench, task)  # 输入初稿的基线（每题只算一次）
            res["ai_traces"] = {"input": base, "output": out_stats, "rows": ai_trace_rows(base, out_stats)}
            merged.unlink(missing_ok=True)
        else:
            res["fact_diff"] = {}
    elif track == "F":
        process_f(bench, row, task, wdir, produced, reply_text, opts, res)

    summary = machine_summary(res, track)
    scrub = build_scrubber(bench, row["contender"], wdir)
    res["sanitized"] = {"dir": f"sanitized/{row['hash']}", "files": write_sanitized(bench, row, wdir, bodies, scrub, summary)}
    return res


def select_rows(bench, only):
    if not only:
        return list(bench.index)
    keys = set(only.split(","))
    return [r for r in bench.index if r["hash"] in keys or r["task"] in keys or r["contender"] in keys]


def main():
    if len(sys.argv) >= 5 and sys.argv[1] == "--render-pdf":  # 内部模式：由 uv 带 pymupdf 调用
        print(json.dumps(render_pdf_inproc(sys.argv[2], sys.argv[3], float(sys.argv[4])), ensure_ascii=False))
        return 0
    ap = argparse.ArgumentParser(description="竞品对标评测机检")
    ap.add_argument("--bench", required=True, help="评测目录")
    ap.add_argument("--only", default="", help="只处理这些哈希、题号或参赛方（逗号分隔）")
    ap.add_argument("--no-render", action="store_true", help="F 赛道不转 PDF、不渲染图片")
    ap.add_argument("--jobs", type=int, default=3, help="并发处理的运行数（LibreOffice 转换较慢），默认 3")
    opts = ap.parse_args()
    bench = Bench(opts.bench)
    if not bench.index:
        print("评测目录里没有 index.json，请先运行 prep.py", file=sys.stderr)
        return 2
    hashes = read_json(bench.root / "hashes.json", {})
    (bench.root / "machine").mkdir(exist_ok=True)
    rows = select_rows(bench, opts.only)

    def work(row):
        try:
            res = process(bench, row, hashes, opts)
        except Exception as e:  # 单次运行出错不影响其他运行
            res = {"hash": row["hash"], "task": row["task"], "track": row["track"], "contender": row["contender"],
                   "rep": row["rep"], "dir": row["dir"], "error": f"{type(e).__name__}: {e}"}
        write_json(bench.root / "machine" / f"{row['hash']}.json", res)
        return row, res

    # H 题的输入初稿基线先串行算好，避免并发时重复计算
    for t in sorted({r["task"] for r in rows if r["track"] == "H"}):
        baseline_for(bench, bench.task(t))
    n = 0
    with ThreadPoolExecutor(max_workers=max(1, opts.jobs)) as ex:
        for row, res in ex.map(work, rows):
            n += 1
            tag = f"{row['task']}-{row['contender']}-r{row['rep']}"
            if "error" in res:
                print(f"{tag}: 机检出错：{res['error']}")
                continue
            extra = ""
            if row["track"] == "F":
                extra = f"，docx {len(res.get('docx', []))} 个"
            if row["track"] == "H" and res.get("facts"):
                extra = f"，事实保留 {res['facts']['kept']}/{res['facts']['total']}"
            flag = "" if res["has_reply"] else "  缺少 _回复.md"
            print(f"{tag}: 正文 {len(res['body']['files'])} 份{extra}{flag}")
    shutil.rmtree(bench.root / "machine" / "_tmp", ignore_errors=True)
    print(f"共处理 {n} 次运行")
    return 0


if __name__ == "__main__":
    sys.exit(main())
