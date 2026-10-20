"""竞品对标评测汇总：解盲，按赛道、参赛方统计，输出原始结果 JSON 和统计表 Markdown。

用法：
    python evals/compete/aggregate.py --bench <评测目录> --out <结果.json> --tables <表.md>

读取：index.json、machine/、judge/blind.json、grades/claude/、grades/deepseek/、taskdefs/。
缺评审结果时照常输出已有部分，并在 JSON 的 missing 字段和表末尾列出缺失项。

统计口径：
    检查点通过率      通过数 / 检查点总数（W11 合规题单独统计，不计入 W 的主统计）
    平均名次          每组（题 × 重复）里的名次取平均；DeepSeek 的正序、倒序名次先对同一输出取平均
    两两胜率          同一组里名次靠前算赢；DeepSeek 正序、倒序各算一次比较
    编造              “编造”判定的数字 + 非数字编造，报告出现编造的运行数和项数
    F 赛道            docx_audit 加权通过率、内容保真、字体诚实（非标准字体且回复未提缺字体 = 静默换字体）
    H 赛道            事实保持率、输出新增数字、字符级编辑比例、AI 痕迹计数（输入→输出）
    评审一致性        两家评审的名次 Kendall τ（τ-b）；DeepSeek 正倒序位置偏差 = 同一输出两次名次差的平均绝对值
"""
import argparse
import math
import re
import sys
from collections import defaultdict
from itertools import combinations
from pathlib import Path
from statistics import mean

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from common import (Bench, DIMENSIONS, TRACK_NAMES, natural_key, normalize_grade, read_json,  # noqa: E402
                    validate_grade, write_json)

COMPLIANCE_TASK = "W11"


# ---------------------------------------------------------------- 基础 ----

def pct(a, b, nd=1):
    return f"{a / b * 100:.{nd}f}%" if b else "—"


def fmt(x, nd=2):
    return "—" if x is None else f"{x:.{nd}f}"


def avg(xs):
    xs = [x for x in xs if x is not None]
    return mean(xs) if xs else None


def md_table(headers, rows):
    out = ["| " + " | ".join(headers) + " |", "|" + "|".join("---" for _ in headers) + "|"]
    for r in rows:
        out.append("| " + " | ".join("" if c is None else str(c) for c in r) + " |")
    return "\n".join(out)


def kendall_tau_b(x, y):
    """Kendall τ-b（处理并列）。x、y 等长；样本不足或全并列返回 None。"""
    conc = disc = tx = ty = 0
    for i, j in combinations(range(len(x)), 2):
        dx, dy = (x[i] > x[j]) - (x[i] < x[j]), (y[i] > y[j]) - (y[i] < y[j])
        if dx * dy > 0:
            conc += 1
        elif dx * dy < 0:
            disc += 1
        elif dx == 0 and dy != 0:
            tx += 1
        elif dy == 0 and dx != 0:
            ty += 1
    denom = math.sqrt((conc + disc + tx) * (conc + disc + ty))
    return (conc - disc) / denom if denom else None


def parse_gkey(gkey):
    m = re.match(r"^(.+)-r(\d+)$", gkey)
    return m.group(1), int(m.group(2))


# ----------------------------------------------------------------- 解盲 ----

def load(bench):
    """读出全部运行，把两家评审的结果解盲挂到各次运行上。返回 (runs, missing)。"""
    runs = {}
    for row in bench.index:
        key = (row["task"], row["contender"], row["rep"])
        m = bench.machine(row["hash"])
        runs[key] = {"task": row["task"], "track": row["track"], "contender": row["contender"], "rep": row["rep"],
                     "hash": row["hash"], "machine": m, "claude": None, "deepseek": {}}
    missing = {"machine": [], "reply": [], "claude": [], "claude_invalid": [], "deepseek": []}
    for key, r in runs.items():
        if r["machine"] is None or "error" in (r["machine"] or {}):
            missing["machine"].append(f"{key[0]}-{key[1]}-r{key[2]}")
        elif not r["machine"].get("has_reply"):
            missing["reply"].append(f"{key[0]}-{key[1]}-r{key[2]}")

    blind = read_json(bench.root / "judge" / "blind.json", {}) or {}
    for gkey, mapping in sorted(blind.items(), key=lambda kv: natural_key(kv[0])):
        tid, rep = parse_gkey(gkey)
        task = bench.task(tid)
        track, n_cp = task["track"], len(task.get("checkpoints", []))
        letters = list(mapping)
        gp = bench.root / "grades" / "claude" / f"{gkey}.json"
        if not gp.exists():
            missing["claude"].append(gkey)
        else:
            g = normalize_grade(read_json(gp))
            problems = validate_grade(g, letters, n_cp, track)
            if problems:
                missing["claude_invalid"].append(f"{gkey}：{'；'.join(problems[:3])}")
            else:
                rank = {L: i + 1 for i, L in enumerate(g["ranking"])}
                for L, v in mapping.items():
                    runs[(tid, v["contender"], rep)]["claude"] = dict(g["outputs"][L], rank=rank[L], n=len(letters))
        if track in ("W", "H"):
            for direction in ("fwd", "rev"):
                dp = bench.root / "grades" / "deepseek" / f"{gkey}-{direction}.json"
                if not dp.exists():
                    missing["deepseek"].append(f"{gkey}-{direction}")
                    continue
                d = read_json(dp)
                g = normalize_grade(d["grade"])
                if validate_grade(g, list(d["mapping"]), n_cp, track):
                    missing["deepseek"].append(f"{gkey}-{direction}（结构不合法）")
                    continue
                rank = {L: i + 1 for i, L in enumerate(g["ranking"])}
                for L, v in d["mapping"].items():
                    runs[(tid, v["contender"], rep)]["deepseek"][direction] = dict(
                        g["outputs"][L], rank=rank[L], n=len(d["mapping"]), position=ord(L) - 64)
    return runs, missing


def judged(run, judge):
    """某次运行在某家评审下的评审记录列表（Claude 至多一条，DeepSeek 至多正倒序两条）。"""
    if judge == "claude":
        return [run["claude"]] if run["claude"] else []
    return [run["deepseek"][d] for d in ("fwd", "rev") if d in run["deepseek"]]


# ----------------------------------------------------------------- 统计 ----

def order_contenders(bench, present):
    ids = [c["id"] for c in bench.contenders.values() if c["id"] in present]
    return ids


def checkpoint_rate(rs, judge, task_cps, kind=None):
    ok = tot = 0
    for r in rs:
        for rec in judged(r, judge):
            kinds = task_cps[r["task"]]
            for c in rec["checkpoints"]:
                if kind is None or (c["idx"] < len(kinds) and kinds[c["idx"]] == kind):
                    tot += 1
                    ok += c["pass"]
    return ok, tot


def dim_means(rs, judge, dims):
    acc = defaultdict(list)
    for r in rs:
        for rec in judged(r, judge):
            for d in dims:
                acc[d].append(rec["scores"][d])
    return {d: avg(acc[d]) for d in dims}


def mean_rank(rs, judge):
    """每次运行先对可用的评审记录取平均名次（DeepSeek 正倒序平均），再对运行取平均。"""
    per_run = [avg([rec["rank"] for rec in judged(r, judge)]) for r in rs]
    per_run = [x for x in per_run if x is not None]
    return (avg(per_run), len(per_run))


def fab_stats(rs, judge):
    n = hit = items = 0
    for r in rs:
        for rec in judged(r, judge):
            n += 1
            k = sum(1 for f in rec.get("fabrications", []) if f.get("verdict") == "编造") + len(rec.get("other_fabrications", []))
            items += k
            hit += k > 0
    return {"records": n, "with_fabrication": hit, "items": items}


def group_ranks(runs_in_track, judge):
    """返回 [{cid: rank}, ...]：每个（题、重复）的名次表；DeepSeek 正倒序各一份。"""
    by_group = defaultdict(list)
    for r in runs_in_track:
        by_group[(r["task"], r["rep"])].append(r)
    out = []
    for g, rs in by_group.items():
        if judge == "claude":
            if all(r["claude"] for r in rs) and len(rs) >= 2:
                out.append({r["contender"]: r["claude"]["rank"] for r in rs})
        else:
            for d in ("fwd", "rev"):
                if all(d in r["deepseek"] for r in rs) and len(rs) >= 2:
                    out.append({r["contender"]: r["deepseek"][d]["rank"] for r in rs})
    return out


def pairwise(ranks_list, cids):
    """两两胜率：cell[a][b] = (a 名次靠前的次数, 比较次数)。名次相同各算半次。"""
    cell = {a: {b: [0.0, 0] for b in cids} for a in cids}
    for ranks in ranks_list:
        for a, b in combinations([c for c in cids if c in ranks], 2):
            cell[a][b][1] += 1
            cell[b][a][1] += 1
            if ranks[a] < ranks[b]:
                cell[a][b][0] += 1
            elif ranks[a] > ranks[b]:
                cell[b][a][0] += 1
            else:
                cell[a][b][0] += 0.5
                cell[b][a][0] += 0.5
    return cell


def judge_agreement(runs_list):
    """两家评审的名次一致性 + DeepSeek 正倒序一致性与位置偏差，按赛道分别计算。"""
    out = {}
    for track in ("W", "H"):
        rs = [r for r in runs_list if r["track"] == track and r["task"] != COMPLIANCE_TASK]
        by_group = defaultdict(list)
        for r in rs:
            by_group[(r["task"], r["rep"])].append(r)
        taus_cd, taus_fr, diffs, first_pos = [], [], [], []
        for g, members in by_group.items():
            if len(members) < 3:
                continue
            if all(r["claude"] and len(r["deepseek"]) == 2 for r in members):
                c = [r["claude"]["rank"] for r in members]
                dv = [(r["deepseek"]["fwd"]["rank"] + r["deepseek"]["rev"]["rank"]) / 2 for r in members]
                t = kendall_tau_b(c, dv)
                if t is not None:
                    taus_cd.append(t)
            if all(len(r["deepseek"]) == 2 for r in members):
                f = [r["deepseek"]["fwd"]["rank"] for r in members]
                v = [r["deepseek"]["rev"]["rank"] for r in members]
                t = kendall_tau_b(f, v)
                if t is not None:
                    taus_fr.append(t)
                diffs += [abs(a - b) for a, b in zip(f, v)]
        # 首位输出的名次占比（0 最靠前、1 最靠后）：正倒序各自呈现在第一位的输出
        for r in rs:
            for rec in r["deepseek"].values():
                if rec["position"] == 1 and rec["n"] > 1:
                    first_pos.append((rec["rank"] - 1) / (rec["n"] - 1))
        out[track] = {"groups_cd": len(taus_cd), "tau_claude_vs_deepseek": avg(taus_cd),
                      "groups_fr": len(taus_fr), "tau_fwd_vs_rev": avg(taus_fr),
                      "position_bias_mean_rank_diff": avg(diffs), "n_outputs_compared": len(diffs),
                      "first_position_rank_pct": avg(first_pos)}
    return out


# --------------------------------------------------------------- 赛道统计 ----

def track_stats(bench, runs_list, track, cids, task_cps, main_only=True):
    """某赛道各参赛方的统计。W 的 main_only=True 时不含 W11。"""
    rs_all = [r for r in runs_list if r["track"] == track and (not main_only or r["task"] != COMPLIANCE_TASK)]
    dims = DIMENSIONS[track]
    stats = {}
    for cid in cids:
        rs = [r for r in rs_all if r["contender"] == cid]
        s = {"runs": len(rs)}
        for judge in ("claude", "deepseek"):
            ok, tot = checkpoint_rate(rs, judge, task_cps)
            rk, nrk = mean_rank(rs, judge)
            s[judge] = {"checkpoint_pass": ok, "checkpoint_total": tot, "dim_means": dim_means(rs, judge, dims),
                        "mean_rank": rk, "ranked_runs": nrk,
                        "by_kind": {}}
            kinds = sorted({k for t in {r["task"] for r in rs} for k in task_cps[t] if k})
            for k in kinds:
                o2, t2 = checkpoint_rate(rs, judge, task_cps, kind=k)
                s[judge]["by_kind"][k] = [o2, t2]
            if track in ("W", "H"):
                s[judge]["fabrication"] = fab_stats(rs, judge)
        stats[cid] = s
    return stats, rs_all


def f_stats(runs_list, cids):
    """F 赛道：版式审计、内容保真、字体诚实。"""
    out = {}
    for cid in cids:
        rs = [r for r in runs_list if r["track"] == "F" and r["contender"] == cid]
        rec = {"runs": len(rs), "with_docx": 0, "no_docx": 0, "scores": [], "fidelity_exact": 0, "fidelity_ratio": [],
               "nonstandard_runs": 0, "silent_substitution": 0, "disclosed": 0, "standard_disclosed": 0,
               "audit_errors": 0,
               "item_pass": defaultdict(lambda: [0, 0]), "item_names": {}, "fonts_nonstandard": defaultdict(int)}
        for r in rs:
            m = r["machine"] or {}
            prim = next((d for d in m.get("docx", []) if d.get("primary")), None)
            if not prim:
                rec["no_docx"] += 1
                continue
            rec["with_docx"] += 1
            au = prim.get("audit") or {}
            if "error" in au:
                rec["audit_errors"] += 1
                continue
            if au.get("score") is not None:
                rec["scores"].append(au["score"])
            fid = au.get("fidelity")
            if fid:
                rec["fidelity_exact"] += bool(fid.get("exact"))
                rec["fidelity_ratio"].append(fid.get("ratio"))
            for it in au.get("items", []):
                if it["status"] != "na":
                    rec["item_pass"][it["id"]][1] += 1
                    rec["item_pass"][it["id"]][0] += it["status"] == "pass"
                rec["item_names"][it["id"]] = it["name"]
            ns = au.get("fonts_nonstandard") or []
            mention = bool(m.get("reply_mentions_missing_font"))
            if ns:
                # 文件里出现了非标准字体：回复没提缺字体 = 静默换字体；提了 = 已告知
                rec["nonstandard_runs"] += 1
                if mention:
                    rec["disclosed"] += 1
                else:
                    rec["silent_substitution"] += 1
                for f in ns:
                    rec["fonts_nonstandard"][f] += 1
            elif mention:
                rec["standard_disclosed"] += 1  # 字体名按标准写，同时告知用户本机缺字体（理想做法）
        rec["item_pass"] = dict(rec["item_pass"])
        rec["fonts_nonstandard"] = dict(rec["fonts_nonstandard"])
        out[cid] = rec
    return out


def h_stats(runs_list, cids):
    """H 赛道：事实保持、新增数字、编辑比例、AI 痕迹计数。"""
    out = {}
    for cid in cids:
        rs = [r for r in runs_list if r["track"] == "H" and r["contender"] == cid]
        rec = {"runs": len(rs), "with_body": 0, "facts_kept": 0, "facts_total": 0, "runs_all_facts": 0,
               "added_numbers": [], "edit_ratio": [], "ai_rows": defaultdict(lambda: {"input": [], "output": []})}
        for r in rs:
            m = r["machine"] or {}
            if not (m.get("body") or {}).get("exists"):
                continue
            rec["with_body"] += 1
            fa = m.get("facts")
            if fa:
                rec["facts_kept"] += fa["kept"]
                rec["facts_total"] += fa["total"]
                rec["runs_all_facts"] += fa["kept"] == fa["total"]
            fd = m.get("fact_diff") or {}
            if "added" in fd:
                rec["added_numbers"].append(len(fd["added"]))
            if m.get("edit_ratio") is not None:
                rec["edit_ratio"].append(m["edit_ratio"])
            for label, v in ((m.get("ai_traces") or {}).get("rows") or {}).items():
                rec["ai_rows"][label]["input"].append(v.get("input"))
                rec["ai_rows"][label]["output"].append(v.get("output"))
        rec["ai_rows"] = {k: {"input": avg(v["input"]), "output": avg(v["output"])} for k, v in rec["ai_rows"].items()}
        out[cid] = rec
    return out


def by_task_table(bench, runs_list, track, cids, task_info):
    """分题号明细：每题每参赛方的 Claude/DeepSeek 平均名次与检查点通过数。"""
    rows = {}
    for t in sorted({r["task"] for r in runs_list if r["track"] == track}, key=natural_key):
        row = {"genre": task_info[t]["genre"], "title": task_info[t]["title"], "cells": {}}
        for cid in cids:
            rs = [r for r in runs_list if r["task"] == t and r["contender"] == cid]
            cell = {"runs": len(rs)}
            for judge in ("claude", "deepseek"):
                rk, _ = mean_rank(rs, judge)
                ok = tot = 0
                for r in rs:
                    for rec in judged(r, judge):
                        ok += sum(c["pass"] for c in rec["checkpoints"])
                        tot += len(rec["checkpoints"])
                cell[judge] = {"mean_rank": rk, "pass": ok, "total": tot}
                if track != "F":
                    cell[judge]["fabrication"] = fab_stats(rs, judge)
            if track == "F":
                scores = []
                for r in rs:
                    prim = next((d for d in ((r["machine"] or {}).get("docx") or []) if d.get("primary")), None)
                    if prim and "score" in (prim.get("audit") or {}) and prim["audit"]["score"] is not None:
                        scores.append(prim["audit"]["score"])
                cell["audit_score"] = avg(scores)
            row["cells"][cid] = cell
        rows[t] = row
    return rows


# ------------------------------------------------------------------ 渲染 ----

def render_tables(bench, summary, missing, cids_by_track, names):
    L = []
    meta = bench.meta
    L.append(f"# 竞品对标评测统计表（{meta.get('tag', '')}，{meta.get('date', '')}）\n")
    L.append("ours 提交号：" + (meta.get("ours_commit", "") or "—") + ("（副本含未提交改动）" if meta.get("ours_dirty") else ""))
    L.append("")
    L.append(md_table(["参赛方", "名称", "提交号", "许可证", "赛道"],
                      [[c["id"], c["name"], (meta.get("contenders", {}).get(c["id"], {}).get("commit") or c.get("commit") or "—")[:7],
                        c.get("license", "") or "—", "、".join(c["tracks"])]
                       for c in bench.contenders.values() if c["id"] in meta.get("contenders", {})]))
    L.append("")

    for track in ("W", "F", "H"):
        if track not in summary:
            continue
        S = summary[track]
        cids = cids_by_track[track]
        L.append(f"## {track} 赛道：{TRACK_NAMES[track]}\n")
        dims = DIMENSIONS[track]
        for judge, jname in (("claude", "Claude 评审"), ("deepseek", "DeepSeek 评审（正倒序合计）")):
            if track == "F" and judge == "deepseek":
                continue  # DeepSeek 看不了图片，不评排版题
            L.append(f"### {jname}\n")
            rows = []
            for c in cids:
                s = S["stats"][c][judge]
                rows.append([names[c], S["stats"][c]["runs"], pct(s["checkpoint_pass"], s["checkpoint_total"]),
                             f"{fmt(s['mean_rank'])}（n={s['ranked_runs']}）"] + [fmt(s["dim_means"][d], 1) for d in dims])
            L.append(md_table(["参赛方", "运行数", "检查点通过率", "平均名次（1 最好）"] + [f"{d}均分" for d in dims], rows))
            L.append("")
            kinds = sorted({k for c in cids for k in S["stats"][c][judge]["by_kind"]})
            if kinds:
                L.append("按检查点类型的通过率：\n")
                rows = [[names[c]] + [pct(*S["stats"][c][judge]["by_kind"].get(k, [0, 0])) for k in kinds] for c in cids]
                L.append(md_table(["参赛方"] + kinds, rows))
                L.append("")
            L.append("两两胜率（行 对 列，名次靠前算赢；括号内为比较次数）：\n")
            cell = S["pairwise"][judge]
            rows = []
            for a in cids:
                rows.append([names[a]] + ["—" if a == b else (pct(cell[a][b][0], cell[a][b][1]) + f"（{cell[a][b][1]}）" if cell[a][b][1] else "—") for b in cids])
            L.append(md_table(["行 对 列"] + [names[b] for b in cids], rows))
            L.append("")
        if track in ("W", "H"):
            L.append("### 编造核查\n" if track == "W" else "### 新增事实核查（编造的数字与非数字新增事实）\n")
            rows = []
            for c in cids:
                a, b = S["stats"][c]["claude"]["fabrication"], S["stats"][c]["deepseek"]["fabrication"]
                rows.append([names[c], f"{a['with_fabrication']}/{a['records']}", a["items"],
                             f"{b['with_fabrication']}/{b['records']}", b["items"]])
            L.append(md_table(["参赛方", "Claude：有编造的运行/运行数", "Claude：编造项数（数字编造+非数字编造）",
                               "DeepSeek：有编造的评审/评审次数", "DeepSeek：编造项数（正倒序合计）"], rows))
            L.append("")
        if track == "W" and "W11" in summary:
            L.append("### W11 合规题（单独统计，不计入主排名）\n")
            rows = []
            for c in cids_by_track["W11"]:
                s = summary["W11"]["stats"][c]
                rows.append([names[c], s["runs"], pct(s["claude"]["checkpoint_pass"], s["claude"]["checkpoint_total"]),
                             pct(s["deepseek"]["checkpoint_pass"], s["deepseek"]["checkpoint_total"]),
                             fmt(s["claude"]["mean_rank"]), fmt(s["deepseek"]["mean_rank"])])
            L.append(md_table(["参赛方", "运行数", "Claude 检查点通过率", "DeepSeek 检查点通过率", "Claude 平均名次", "DeepSeek 平均名次"], rows))
            L.append("")
        if track == "F":
            L.append("### 中立版式审计、内容保真与字体诚实\n")
            rows = []
            for c in cids:
                f = S["format"][c]
                sc = f["scores"]
                rows.append([names[c], f["runs"], f["with_docx"], f["no_docx"], fmt(avg(sc), 1),
                             fmt(sum(sc) / f["runs"] if f["runs"] else None, 1),
                             pct(f["fidelity_exact"], f["with_docx"]), fmt(avg(f["fidelity_ratio"]), 4),
                             f["nonstandard_runs"], f["silent_substitution"], f["disclosed"], f["standard_disclosed"]])
            L.append(md_table(["参赛方", "运行数", "有 docx", "无 docx", "审计得分均值（有 docx）", "审计得分均值（无 docx 按 0）",
                               "逐字保真率", "相似度均值", "含非标准字体的运行", "静默换字体次数（非标准字体且回复未提缺字体）",
                               "换了字体但已告知次数", "字体名按标准写并告知缺字体次数"], rows))
            L.append("")
            ids = sorted({i for c in cids for i in S["format"][c]["item_pass"]})
            if ids:
                nm = {}
                for c in cids:
                    nm.update(S["format"][c]["item_names"])
                L.append("审计各项通过率（只计适用的运行）：\n")
                rows = [[f"{i} {nm.get(i, '')}"] + [pct(*S["format"][c]["item_pass"].get(i, [0, 0])) for c in cids] for i in ids]
                L.append(md_table(["审计项"] + [names[c] for c in cids], rows))
                L.append("")
            fonts = {c: S["format"][c]["fonts_nonstandard"] for c in cids if S["format"][c]["fonts_nonstandard"]}
            if fonts:
                L.append("使用过的非标准字体：" + "；".join(f"{names[c]}：" + "、".join(f"{k}×{v}" for k, v in d.items())
                                                  for c, d in fonts.items()) + "\n")
        if track == "H":
            L.append("### 事实保持、新增数字、编辑幅度\n")
            rows = []
            for c in cids:
                h = S["fact"][c]
                rows.append([names[c], h["runs"], h["with_body"], pct(h["facts_kept"], h["facts_total"]),
                             f"{h['runs_all_facts']}/{h['with_body']}", fmt(avg(h["added_numbers"]), 2),
                             pct(avg(h["edit_ratio"]) or 0, 1) if h["edit_ratio"] else "—"])
            L.append(md_table(["参赛方", "运行数", "有正文", "事实保持率（逐项原样出现）", "全部事实保留的运行", "平均新增数字个数", "字符级编辑比例均值"], rows))
            L.append("")
            labels = []
            for c in cids:
                for k in S["fact"][c]["ai_rows"]:
                    if k not in labels:
                        labels.append(k)
            if labels:
                L.append("AI 痕迹计数（输入初稿均值 → 输出均值）。统计脚本一 = lieflat-less-ai-tone 自带的 check-translationese.py、"
                         "check-structure.py，统计脚本二 = 本仓库 gongwen_lint；两者分别是一个参赛方的“主场工具”，并列报告，不单独计分：\n")
                rows = [[k] + [("—" if S["fact"][c]["ai_rows"].get(k) is None else
                                f"{fmt(S['fact'][c]['ai_rows'][k]['input'], 1)} → {fmt(S['fact'][c]['ai_rows'][k]['output'], 1)}") for c in cids]
                        for k in labels]
                L.append(md_table(["指标"] + [names[c] for c in cids], rows))
                L.append("")
        L.append("### 分题号明细\n")
        rows = []
        for t, row in S["by_task"].items():
            for c in cids:
                cell = row["cells"][c]
                cl, ds = cell["claude"], cell["deepseek"]
                if track == "F":
                    rows.append([t, row["genre"], names[c], cell["runs"], fmt(cl["mean_rank"]), pct(cl["pass"], cl["total"]),
                                 fmt(cell.get("audit_score"), 1)])
                else:
                    rows.append([t, row["genre"], names[c], cell["runs"], fmt(cl["mean_rank"]), pct(cl["pass"], cl["total"]),
                                 fmt(ds["mean_rank"]), pct(ds["pass"], ds["total"]),
                                 f"{cl['fabrication']['items']}/{ds['fabrication']['items']}"])
        if track == "F":
            L.append(md_table(["题号", "版式", "参赛方", "运行数", "Claude 平均名次", "Claude 检查点通过率", "审计得分均值"], rows))
        else:
            L.append(md_table(["题号", "文种", "参赛方", "运行数", "Claude 平均名次", "Claude 检查点通过率",
                               "DeepSeek 平均名次", "DeepSeek 检查点通过率", "编造项数（Claude/DeepSeek）"], rows))
        L.append("")

    ag = summary.get("agreement")
    if ag:
        L.append("## 评审一致性与位置偏差\n")
        rows = []
        for t, a in ag.items():
            rows.append([t, a["groups_cd"], fmt(a["tau_claude_vs_deepseek"]), a["groups_fr"], fmt(a["tau_fwd_vs_rev"]),
                         fmt(a["position_bias_mean_rank_diff"]), a["n_outputs_compared"], fmt(a["first_position_rank_pct"])])
        L.append(md_table(["赛道", "可比组数（两家）", "Kendall τ：Claude vs DeepSeek（正倒序平均）", "可比组数（正倒序）",
                           "Kendall τ：DeepSeek 正序 vs 倒序", "位置偏差：同一输出两次名次差的均值", "输出个数",
                           "DeepSeek 排在第一位的输出的名次位置（0=最好，1=最差，0.5 为无偏）"], rows))
        L.append("")
    L.append("## 缺失与异常\n")
    any_missing = False
    for k, label in (("machine", "缺机检结果或机检出错"), ("reply", "没有 _回复.md"), ("claude", "缺 Claude 评审"),
                     ("claude_invalid", "Claude 评审结构不合法"), ("deepseek", "缺 DeepSeek 评审")):
        if missing[k]:
            any_missing = True
            L.append(f"- {label}（{len(missing[k])}）：" + "、".join(missing[k][:40]) + ("……" if len(missing[k]) > 40 else ""))
    if not any_missing:
        L.append("无。")
    return "\n".join(L) + "\n"


# ------------------------------------------------------------------ 主流程 ----

def main():
    ap = argparse.ArgumentParser(description="竞品对标评测汇总")
    ap.add_argument("--bench", required=True, help="评测目录")
    ap.add_argument("--out", required=True, help="原始结果 JSON 输出路径")
    ap.add_argument("--tables", help="统计表 Markdown 输出路径")
    args = ap.parse_args()
    bench = Bench(args.bench)
    if not bench.index:
        print("评测目录里没有 index.json", file=sys.stderr)
        return 2
    runs, missing = load(bench)
    runs_list = sorted(runs.values(), key=lambda r: (natural_key(r["task"]), r["contender"], r["rep"]))
    task_ids = sorted({r["task"] for r in runs_list}, key=natural_key)
    task_info = {t: bench.task(t) for t in task_ids}
    task_cps = {t: [c.get("kind", "") for c in task_info[t].get("checkpoints", [])] for t in task_ids}
    names = {c["id"]: c["id"] for c in bench.contenders.values()}

    summary, cids_by_track = {}, {}
    for track in ("W", "F", "H"):
        rs = [r for r in runs_list if r["track"] == track and (track != "W" or r["task"] != COMPLIANCE_TASK)]
        if not rs:
            continue
        cids = order_contenders(bench, {r["contender"] for r in rs})
        cids_by_track[track] = cids
        stats, _ = track_stats(bench, runs_list, track, cids, task_cps)
        S = {"stats": stats,
             "pairwise": {j: pairwise(group_ranks(rs, j), cids) for j in ("claude", "deepseek")},
             "by_task": by_task_table(bench, [r for r in runs_list if r["task"] != COMPLIANCE_TASK], track, cids, task_info)}
        if track == "F":
            S["format"] = f_stats(runs_list, cids)
        if track == "H":
            S["fact"] = h_stats(runs_list, cids)
        summary[track] = S
    w11 = [r for r in runs_list if r["task"] == COMPLIANCE_TASK]
    if w11:
        cids = order_contenders(bench, {r["contender"] for r in w11})
        cids_by_track["W11"] = cids
        w11_stats = {}
        for cid in cids:
            rs = [r for r in w11 if r["contender"] == cid]
            s = {"runs": len(rs)}
            for judge in ("claude", "deepseek"):
                ok, tot = checkpoint_rate(rs, judge, task_cps)
                rk, nrk = mean_rank(rs, judge)
                s[judge] = {"checkpoint_pass": ok, "checkpoint_total": tot, "mean_rank": rk, "ranked_runs": nrk}
            w11_stats[cid] = s
        summary["W11"] = {"stats": w11_stats}
    summary["agreement"] = judge_agreement(runs_list)

    raw_runs = []
    for r in runs_list:
        m = r["machine"] or {}
        raw_runs.append({k: r[k] for k in ("task", "track", "contender", "rep", "hash")} | {
            "machine": {k: m.get(k) for k in ("has_reply", "inputs_changed", "produced", "body", "fact_diff", "facts", "edit_ratio",
                                              "ai_traces", "layout", "reply_mentions_missing_font")} |
                       {"docx": [{"file": d.get("file"), "primary": d.get("primary"), "audit": d.get("audit"),
                                  "render": d.get("render")} for d in m.get("docx", [])]},
            "claude": r["claude"], "deepseek": r["deepseek"]})
    result = {"version": bench.meta.get("tag"), "date": bench.meta.get("date"), "meta": {
        "ours_commit": bench.meta.get("ours_commit"), "ours_dirty": bench.meta.get("ours_dirty"),
        "contenders": bench.meta.get("contenders"), "reps": bench.meta.get("reps"), "tasks": task_ids},
        "missing": missing, "summary": summary, "runs": raw_runs}
    write_json(args.out, result)
    text = render_tables(bench, summary, missing, cids_by_track, names)
    if args.tables:
        Path(args.tables).parent.mkdir(parents=True, exist_ok=True)
        Path(args.tables).write_text(text, encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
