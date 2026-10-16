"""跨 skill 触发评测：准备任务说明。

模拟 Agent 工具按 description 选择 skill 的过程：给模型看全部 skill 的 name 和 description，
再给一批用户说法，要求为每条说法选出最该处理它的一个 skill（或"无"）。

用法：
    python evals/bench/trigger_prep.py --out <评测目录> [--batch 55] [--seed 20261016]

产出（都在 <评测目录>/trigger/ 下）：
    queries.json        全部用例：编号、说法、来源 skill、should_trigger（不给模型看）
    tasks/batch-<n>.md  子智能体任务说明；答案写到 answers/batch-<n>.json
"""
import argparse
import json
import random
import re
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))
from skill_list import select_skills  # noqa: E402

FRONT_RE = re.compile(r"^---\s*\n(.*?)\n---\s*\n", re.S)


def description(skill):
    text = (REPO / "skills" / skill / "SKILL.md").read_text(encoding="utf-8")
    m = FRONT_RE.match(text)
    if not m:
        raise SystemExit(f"{skill} 的 SKILL.md 没有 frontmatter")
    for line in m.group(1).splitlines():
        if line.startswith("description:"):
            return line.split(":", 1)[1].strip()
    raise SystemExit(f"{skill} 的 SKILL.md 没有 description")


TEMPLATE = """# 任务说明：为用户说法选择 skill

你在一个装有下列 {k} 个 skill 的 Agent 工具中工作。工具会根据用户说的话，按各 skill 的描述（description）决定调用哪一个。请你扮演这个选择过程。

## 已安装的 skill

{skills}

## 要求

1. 对下面每条用户说法，选出最应该处理它的**一个** skill，填 skill 名称；如果没有任何 skill 适合，填"无"。
2. 只依据上面的描述判断，不要读取任何文件。
3. 把答案写成 JSON 文件：{answer_path}
   格式：{{"answers": [{{"id": "编号", "skill": "skill 名称或无"}}, ...]}}，覆盖全部 {n} 条，UTF-8 编码。
4. 最终回复只写一行："完成 {n} 条"。

## 用户说法

{queries}
"""


def main():
    ap = argparse.ArgumentParser(description="准备跨 skill 触发评测")
    ap.add_argument("--out", required=True, help="评测目录（建议放在仓库之外）")
    ap.add_argument("--batch", type=int, default=55, help="每个子智能体处理的说法条数，默认 55")
    ap.add_argument("--seed", type=int, default=20261016)
    ap.add_argument("--skills", nargs="*", help="只取这些 skill（名称或预设 v0.1），默认全部")
    args = ap.parse_args()

    out = Path(args.out).resolve() / "trigger"
    if REPO in out.parents:
        print("评测目录不要放在仓库内", file=sys.stderr)
        return 2
    (out / "tasks").mkdir(parents=True, exist_ok=True)
    (out / "answers").mkdir(parents=True, exist_ok=True)

    skills = select_skills(REPO, args.skills)
    descs = {s: description(s) for s in skills}
    queries = []
    for s in skills:
        data = json.loads((REPO / "skills" / s / "evals" / "trigger.json").read_text(encoding="utf-8"))
        for i, q in enumerate(data, start=1):
            queries.append({"id": f"{s}#{i}", "query": q["query"], "owner": s,
                            "should_trigger": bool(q["should_trigger"])})
    rng = random.Random(args.seed)
    order = queries[:]
    rng.shuffle(order)
    for k, q in enumerate(order, start=1):
        q["qid"] = f"Q{k:03d}"
    (out / "queries.json").write_text(json.dumps(order, ensure_ascii=False, indent=1), encoding="utf-8")

    skill_block = "\n".join(f"- **{s}**：{descs[s]}" for s in skills)
    batches = [order[i:i + args.batch] for i in range(0, len(order), args.batch)]
    for b, items in enumerate(batches, start=1):
        answer_path = out / "answers" / f"batch-{b}.json"
        qtext = "\n".join(f"- {q['qid']}：{q['query']}" for q in items)
        text = TEMPLATE.format(k=len(skills), skills=skill_block, answer_path=answer_path,
                               n=len(items), queries=qtext)
        (out / "tasks" / f"batch-{b}.md").write_text(text, encoding="utf-8")
    print(f"skill {len(skills)} 个，说法 {len(order)} 条，批次 {len(batches)} 个：{out / 'tasks'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
