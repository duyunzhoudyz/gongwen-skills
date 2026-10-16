"""跨 skill 触发评测：计分。

判定：
    should_trigger 为 true 的说法（正例）：选中来源 skill 为对；
    should_trigger 为 false 的说法（反例）：没有选中来源 skill 为对（选别的 skill 或"无"都算对）。

用法：
    python evals/bench/trigger_score.py --bench <评测目录> [--out evals/results/v1.0-trigger.json] [--tables 表.md]
"""
import argparse
import json
import sys
from collections import Counter, defaultdict
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


def pct(a, b):
    return f"{a / b * 100:.1f}%" if b else "—"


def main():
    ap = argparse.ArgumentParser(description="跨 skill 触发评测计分")
    ap.add_argument("--bench", required=True)
    ap.add_argument("--out", help="结果 JSON 输出路径")
    ap.add_argument("--tables", help="统计表 Markdown 输出路径")
    args = ap.parse_args()

    base = Path(args.bench).resolve() / "trigger"
    queries = json.loads((base / "queries.json").read_text(encoding="utf-8"))
    answers = {}
    for f in sorted((base / "answers").glob("batch-*.json")):
        data = json.loads(f.read_text(encoding="utf-8"))
        for a in data.get("answers", []):
            answers[a["id"]] = a.get("skill", "")

    rows, missing = [], []
    per_skill = defaultdict(lambda: {"pos": [0, 0], "neg": [0, 0]})
    confusion = Counter()
    for q in queries:
        got = answers.get(q["qid"])
        if got is None:
            missing.append(q["qid"])
            continue
        if q["should_trigger"]:
            ok = got == q["owner"]
            per_skill[q["owner"]]["pos"][1] += 1
            per_skill[q["owner"]]["pos"][0] += ok
            if not ok:
                confusion[(q["owner"], got)] += 1
        else:
            ok = got != q["owner"]
            per_skill[q["owner"]]["neg"][1] += 1
            per_skill[q["owner"]]["neg"][0] += ok
            if not ok:
                confusion[(q["owner"] + "（反例）", got)] += 1
        rows.append({**q, "chosen": got, "correct": ok})

    total = len(rows)
    correct = sum(r["correct"] for r in rows)
    pos = [r for r in rows if r["should_trigger"]]
    neg = [r for r in rows if not r["should_trigger"]]
    summary = {
        "total": total, "correct": correct,
        "positive": [sum(r["correct"] for r in pos), len(pos)],
        "negative": [sum(r["correct"] for r in neg), len(neg)],
        "missing": missing,
        "by_skill": {s: v for s, v in sorted(per_skill.items())},
        "confusions": [{"from": a, "chosen": b, "count": n} for (a, b), n in confusion.most_common()],
    }

    lines = ["| 口径 | 正确 | 比例 |", "|---|---|---|",
             f"| 全部说法 | {correct}/{total} | {pct(correct, total)} |",
             f"| 正例（应触发本 skill） | {summary['positive'][0]}/{summary['positive'][1]} | {pct(*summary['positive'])} |",
             f"| 反例（不应触发本 skill） | {summary['negative'][0]}/{summary['negative'][1]} | {pct(*summary['negative'])} |",
             "", "| skill | 正例 | 反例 |", "|---|---|---|"]
    for s, v in summary["by_skill"].items():
        lines.append(f"| {s} | {v['pos'][0]}/{v['pos'][1]} | {v['neg'][0]}/{v['neg'][1]} |")
    lines += ["", "主要混淆（来源 → 被选中）：", ""]
    for c in summary["confusions"][:30]:
        lines.append(f"- {c['from']} → {c['chosen']}：{c['count']} 次")
    if missing:
        lines += ["", f"缺少答案：{len(missing)} 条"]
    text = "\n".join(lines) + "\n"
    if args.out:
        Path(args.out).write_text(json.dumps({"summary": summary, "rows": rows}, ensure_ascii=False, indent=1),
                                  encoding="utf-8")
    if args.tables:
        Path(args.tables).write_text(text, encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
