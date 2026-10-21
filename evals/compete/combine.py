# -*- coding: utf-8 -*-
"""合并多轮竞品对标的汇总结果（aggregate.py 输出的 JSON），按参赛方给出合计口径的统计表。

用法：
    python evals/compete/combine.py <agg1.json> <agg2.json> ... --labels c4,c5 --out <合并表.md>

合计口径：名次、维度分按“参与排名的运行数”加权平均；检查点、编造、事实保持等计数直接相加。
"""
import argparse
import json
import sys
from collections import defaultdict
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

TRACK_NAMES = {"W": "写作", "F": "排版", "H": "去 AI 味"}
JUDGES = (("claude", "Claude"), ("deepseek", "DeepSeek"))


def pct(a, b):
    return f"{a / b * 100:.1f}%（{a}/{b}）" if b else "—"


def combine_track(aggs, track):
    """返回 {参赛方: {judge: 合计}}，以及各轮名次明细。"""
    acc = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    per_round = defaultdict(dict)
    for label, agg in aggs:
        stats = agg["summary"][track]["stats"]
        for cid, st in stats.items():
            for jk, _ in JUDGES:
                j = st.get(jk)
                if not j:
                    continue
                a = acc[cid][jk]
                n = j.get("ranked_runs") or 0
                a["cp_pass"] += j.get("checkpoint_pass", 0)
                a["cp_total"] += j.get("checkpoint_total", 0)
                if n and j.get("mean_rank") is not None:
                    a["rank_sum"] += j["mean_rank"] * n
                    a["rank_n"] += n
                    per_round[(cid, jk)][label] = j["mean_rank"]
                for dim, v in (j.get("dim_means") or {}).items():
                    if v is not None and n:
                        a["dim:" + dim] += v * n
                        a["dimn:" + dim] += n
                fab = j.get("fabrication") or {}
                a["fab_records"] += fab.get("records", 0)
                a["fab_with"] += fab.get("with_fabrication", 0)
                a["fab_items"] += fab.get("items", 0)
    return acc, per_round


def render(aggs, labels):
    out = [f"# 竞品对标合并统计（{'、'.join(labels)}）", "",
           "合计口径：平均名次与维度分按参与排名的运行数加权；检查点、编造等计数相加。名次 1 最好。", ""]
    for track in ("W", "F", "H"):
        acc, per_round = combine_track(aggs, track)
        out += [f"## {track} 赛道：{TRACK_NAMES[track]}", ""]
        for jk, jname in JUDGES:
            rows = [(cid, a[jk]) for cid, a in acc.items() if a[jk]["rank_n"]]
            if not rows:
                continue
            rows.sort(key=lambda x: x[1]["rank_sum"] / x[1]["rank_n"])
            dims = sorted({k[4:] for _, a in rows for k in a if k.startswith("dim:")})
            out += [f"### {jname} 评审", "",
                    "| 名次 | 参赛方 | 平均名次（合计） | " + " | ".join(f"{l} 名次" for l in labels)
                    + " | 检查点通过率 | " + " | ".join(dims) + " | 有编造的评审/评审数 | 编造项数 |",
                    "|" + "---|" * (6 + len(labels) + len(dims))]
            for i, (cid, a) in enumerate(rows, 1):
                rounds = " | ".join(f"{per_round[(cid, jk)].get(l, float('nan')):.2f}" if l in per_round[(cid, jk)] else "—"
                                    for l in labels)
                dimv = " | ".join(f"{a['dim:' + d] / a['dimn:' + d]:.2f}" if a.get("dimn:" + d) else "—" for d in dims)
                out.append(f"| {i} | {cid} | {a['rank_sum'] / a['rank_n']:.2f} | {rounds} | "
                           f"{pct(int(a['cp_pass']), int(a['cp_total']))} | {dimv} | "
                           f"{int(a['fab_with'])}/{int(a['fab_records'])} | {int(a['fab_items'])} |")
            out.append("")
        if track == "F":
            fmt = defaultdict(lambda: defaultdict(float))
            for _, agg in aggs:
                for cid, f in agg["summary"]["F"]["format"].items():
                    x = fmt[cid]
                    x["runs"] += f["runs"]
                    x["score_sum"] += sum(f["scores"])
                    x["score_n"] += len(f["scores"])
                    for k in ("fidelity_exact", "nonstandard_runs", "silent_substitution", "disclosed", "standard_disclosed"):
                        x[k] += f.get(k, 0)
            out += ["### 中立版式审计与字体诚实（合计）", "",
                    "| 参赛方 | 运行数 | 审计得分均值 | 逐字保真 | 写入非标准字体的运行 | 其中未告知用户 | 写标准字体名并告知缺字体 |",
                    "|---|---|---|---|---|---|---|"]
            for cid, x in sorted(fmt.items(), key=lambda kv: -kv[1]["score_sum"] / max(kv[1]["score_n"], 1)):
                out.append(f"| {cid} | {int(x['runs'])} | {x['score_sum'] / max(x['score_n'], 1):.1f} | "
                           f"{int(x['fidelity_exact'])}/{int(x['runs'])} | {int(x['nonstandard_runs'])} | "
                           f"{int(x['silent_substitution'])} | {int(x['standard_disclosed'])} |")
            out.append("")
        if track == "H":
            fact = defaultdict(lambda: defaultdict(float))
            for _, agg in aggs:
                for cid, f in agg["summary"]["H"]["fact"].items():
                    x = fact[cid]
                    x["runs"] += f["runs"]
                    x["with_body"] += f["with_body"]
                    x["kept"] += f["facts_kept"]
                    x["total"] += f["facts_total"]
                    x["all"] += f["runs_all_facts"]
                    x["added"] += sum(f["added_numbers"])
                    x["edit_sum"] += sum(f["edit_ratio"])
                    x["edit_n"] += len(f["edit_ratio"])
            out += ["### 事实保持（合计）", "",
                    "| 参赛方 | 运行数 | 有正文 | 事实保持率 | 全部事实保留的运行 | 新增数字合计 | 编辑比例均值 |",
                    "|---|---|---|---|---|---|---|"]
            for cid, x in sorted(fact.items(), key=lambda kv: -kv[1]["kept"] / max(kv[1]["total"], 1)):
                out.append(f"| {cid} | {int(x['runs'])} | {int(x['with_body'])} | {pct(int(x['kept']), int(x['total']))} | "
                           f"{int(x['all'])}/{int(x['with_body'])} | {int(x['added'])} | "
                           f"{x['edit_sum'] / max(x['edit_n'], 1) * 100:.1f}% |")
            out.append("")
    # W11 涉密题
    out += ["## W11 涉密合规题（各轮单列）", "",
            "| 参赛方 | " + " | ".join(f"{l} Claude 检查点" for l in labels) + " | "
            + " | ".join(f"{l} DeepSeek 检查点" for l in labels) + " |",
            "|---|" + "---|" * (2 * len(labels))]
    cids = sorted({cid for _, agg in aggs for cid in agg["summary"]["W11"]["stats"]})
    for cid in cids:
        cells = []
        for jk, _ in JUDGES:
            for _, agg in aggs:
                j = (agg["summary"]["W11"]["stats"].get(cid) or {}).get(jk) or {}
                cells.append(pct(j.get("checkpoint_pass", 0), j.get("checkpoint_total", 0)))
        out.append(f"| {cid} | " + " | ".join(cells) + " |")
    out.append("")
    # 评审一致性
    out += ["## 评审一致性（各轮）", "", "| 轮次 | 赛道 | Claude 与 DeepSeek 的 Kendall τ | DeepSeek 正倒序 τ |", "|---|---|---|---|"]
    for label, agg in aggs:
        for t, v in agg["summary"].get("agreement", {}).items():
            out.append(f"| {label} | {t} | {v.get('tau_claude_vs_deepseek', float('nan')):.2f} | {v.get('tau_fwd_vs_rev', float('nan')):.2f} |")
    out.append("")
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("aggs", nargs="+")
    ap.add_argument("--labels", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    labels = args.labels.split(",")
    if len(labels) != len(args.aggs):
        ap.error("--labels 个数须与输入文件个数一致")
    aggs = [(l, json.loads(Path(p).read_text(encoding="utf-8"))) for l, p in zip(labels, args.aggs)]
    Path(args.out).write_text(render(aggs, labels), encoding="utf-8")
    print(f"已写入 {args.out}")


if __name__ == "__main__":
    main()
