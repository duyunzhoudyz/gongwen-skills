"""盲评准备：把每个用例的 A、B 两组输出随机标为"甲""乙"，生成评分员任务说明。

用法：
    python evals/bench/grade_prep.py --bench <评测目录> [--chunk 批次名 ...] [--seed 20261016]

产出：
    grade_tasks/<批次>.md         评分员任务说明（含 expectations，不含分组信息）
    grade_tasks/<批次>/<skill>-<id>-甲.md  机检摘要的盲化副本
    grade_refs/<skill>.txt         公开文件清单：skill 内置资料中含文件名称、文号的行，评分员据此判断引用是否有出处
    blind.json                     "甲""乙"与 A、B 的对应关系（评分员不可见）
"""
import argparse
import json
import random
import re
import shutil
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

REPO = Path(__file__).resolve().parents[2]
HERE = Path(__file__).resolve().parent

# 公开文件清单：skill 内置资料中出现文件名称（书名号）或文号的行。评分员据此判断引用是否有出处，
# 两组用同一份清单，不泄露分组。
REF_LINE_RE = re.compile(r"《[^》]{2,60}》|〔\d{4}〕\s*\d+\s*号|GB/T\s*\d+")


def write_ref_list(skill, out_dir):
    """把 skills/<skill> 下 SKILL.md 和 references 中含文件名称、文号的行写成清单，返回清单路径。"""
    root = REPO / "skills" / skill
    files = [root / "SKILL.md"] + sorted((root / "references").rglob("*.md"))
    seen, lines = set(), []
    for f in files:
        if not f.exists():
            continue
        for line in f.read_text(encoding="utf-8").splitlines():
            text = line.strip()
            if text and REF_LINE_RE.search(text) and text not in seen:
                seen.add(text)
                lines.append(text)
    path = out_dir / f"{skill}.txt"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


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
    ref_dir = bench / "grade_refs"
    ref_dir.mkdir(exist_ok=True)

    evals, refs = {}, {}
    for row in index:
        if row["skill"] not in evals:
            data = json.loads((REPO / "skills" / row["skill"] / "evals" / "evals.json").read_text(encoding="utf-8"))
            evals[row["skill"]] = {e["id"]: e for e in data["evals"]}
            refs[row["skill"]] = write_ref_list(row["skill"], ref_dir)

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
                f"公开文件清单：{refs[skill]}\n\n"
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
