"""盲评准备：把每个用例的 A、B 两组输出随机标为"甲""乙"，生成评分员任务说明。

用法：
    python evals/bench/grade_prep.py --bench <评测目录> [--chunk 批次名 ...] [--seed 20261016]

产出：
    grade_tasks/<批次>.md         评分员任务说明（含 expectations，不含分组信息）
    grade_tasks/<批次>/<skill>-<id>-甲.md  机检摘要的盲化副本
    blind.json                     "甲""乙"与 A、B 的对应关系（评分员不可见）
"""
import argparse
import json
import random
import shutil
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

REPO = Path(__file__).resolve().parents[2]
HERE = Path(__file__).resolve().parent


def main():
    ap = argparse.ArgumentParser(description="生成盲评任务")
    ap.add_argument("--bench", required=True)
    ap.add_argument("--chunk", nargs="*")
    ap.add_argument("--seed", type=int, default=20261016)
    args = ap.parse_args()
    bench = Path(args.bench).resolve()
    index = json.loads((bench / "index.json").read_text(encoding="utf-8"))
    tpl = (HERE / "templates" / "grade.md").read_text(encoding="utf-8")
    blind_path = bench / "blind.json"
    blind = json.loads(blind_path.read_text(encoding="utf-8")) if blind_path.exists() else {}
    (bench / "grade_tasks").mkdir(exist_ok=True)
    (bench / "grades").mkdir(exist_ok=True)
    scratch = bench / "grade_tmp"
    scratch.mkdir(exist_ok=True)

    evals = {}
    for row in index:
        if row["skill"] not in evals:
            data = json.loads((REPO / "skills" / row["skill"] / "evals" / "evals.json").read_text(encoding="utf-8"))
            evals[row["skill"]] = {e["id"]: e for e in data["evals"]}

    chunks = {}
    for row in index:
        chunks.setdefault(row["chunk"], {}).setdefault((row["skill"], row["id"]), {})[row["group"]] = row

    for chunk, cases in chunks.items():
        if args.chunk and chunk not in args.chunk:
            continue
        rng = random.Random(f"{args.seed}|{chunk}")
        cdir = bench / "grade_tasks" / chunk
        cdir.mkdir(exist_ok=True)
        blocks = []
        for (skill, eid), groups in cases.items():
            key = f"{skill}-{eid}"
            order = ["A", "B"]
            rng.shuffle(order)
            mapping = {"甲": order[0], "乙": order[1]}
            blind[key] = mapping
            e = evals[skill][eid]
            exps = "\n".join(f"{i}. {x}" for i, x in enumerate(e["expectations"]))
            outs = []
            for label in ("甲", "乙"):
                row = groups[mapping[label]]
                src = bench / "machine" / f"{key}-{mapping[label]}.md"
                dst = cdir / f"{key}-{label}.md"
                shutil.copyfile(src, dst)
                outs.append(f"输出{label}：工作目录 {row['dir']}；机检摘要 {dst}")
            files = "、".join(groups["A"]["files"]) or "无"
            blocks.append(
                f"---\n\n## 用例：{skill} #{eid}\n\n用户消息：\n\n<<<\n{e['prompt']}\n>>>\n\n"
                f"用户提供的文件：{files}\n\n输入材料全文（含用户消息）：{bench / 'inputs' / (key + '.txt')}\n\n"
                f"场景标注：单位类型 {e['unit_type']}，角色模式 {e['role_mode']}，标签 {'、'.join(e['tags'])}\n\n"
                f"期望结果简述：{e['expected_output']}\n\n期望（expectations，共 {len(e['expectations'])} 条）：\n\n{exps}\n\n"
                + "\n\n".join(outs) + "\n"
            )
        text = (tpl.replace("{cases}", "\n".join(blocks)).replace("{n}", str(len(cases)))
                .replace("{grades_dir}", str(bench / "grades")).replace("{scratch}", str(scratch)))
        (bench / "grade_tasks" / f"{chunk}.md").write_text(text, encoding="utf-8")
        print(f"{chunk}: {len(cases)} 个用例")
    blind_path.write_text(json.dumps(blind, ensure_ascii=False, indent=1), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
