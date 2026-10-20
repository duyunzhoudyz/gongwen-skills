"""评审准备：把同一题同一轮的全部参赛方输出（脱敏后）随机分配字母，生成 Claude 评审任务说明。

用法：
    python evals/compete/judge_prep.py --bench <评测目录> [--seed 20261021] [--only W01,F03,...]

前置：已运行 prep.py、完成各批次执行、运行过 machine.py（需要 sanitized/ 和 machine/）。

产出（都在 <评测目录>/judge/ 下，评审员只读说明中列出的文件）：
    blind.json                        每组的“字母 → 参赛方、哈希”映射（评审员不可读；aggregate.py 解盲用）
    tasks/<题号>-r<重复>.md            评审任务说明（题目、用户消息、检查点、各匿名输出的路径与机检摘要）
    validate/<题号>-r<重复>.py         评审结果校验脚本（评审员写完结果后运行）
    judge_batches.json                批次清单：batch_id、instruction_path、grade_path、task、rep、track……
评审结果写到 <评测目录>/grades/claude/<题号>-r<重复>.json（格式见 templates/judge-*.md）。
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

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from common import (Bench, COMPETE, DIMENSIONS, dim_defs_text, dims_json_example, natural_key,  # noqa: E402
                    read_json, task_layout, write_json)

LETTERS = "ABCDEFGHIJKLMNOP"

VALIDATOR = '''"""评审结果校验脚本（由 judge_prep.py 生成）：检查 JSON 合法、字母齐全、检查点条数、分数范围和排名。"""
import json
import sys

GRADE = r"@GRADE@"
LETTERS = @LETTERS@
N_CHECKPOINTS = @NCP@
DIMS = @DIMS@
TRACK = "@TRACK@"


def main():
    problems = []
    try:
        with open(GRADE, encoding="utf-8") as f:
            g = json.load(f)
    except FileNotFoundError:
        print("还没有写评审结果文件：" + GRADE)
        return 1
    except json.JSONDecodeError as e:
        print("JSON 不合法：" + str(e))
        return 1
    outs = g.get("outputs")
    if not isinstance(outs, dict):
        print("缺少 outputs 对象")
        return 1
    if sorted(outs) != sorted(LETTERS):
        problems.append("outputs 的字母应为 " + "、".join(LETTERS) + "，实际为 " + "、".join(sorted(outs)))
    for L in LETTERS:
        o = outs.get(L)
        if not isinstance(o, dict):
            continue
        cps = o.get("checkpoints")
        if not isinstance(cps, list):
            problems.append(L + "：checkpoints 不是数组")
        else:
            idxs = sorted(c.get("idx") for c in cps if isinstance(c, dict))
            if idxs != list(range(N_CHECKPOINTS)):
                problems.append(L + "：checkpoints 的 idx 应为 0～%d 各一条，实际 %s" % (N_CHECKPOINTS - 1, idxs))
            for c in cps:
                if not isinstance(c, dict) or not isinstance(c.get("pass"), bool):
                    problems.append(L + "：checkpoints 每项要有布尔值 pass")
                    break
                if not str(c.get("evidence", "")).strip():
                    problems.append(L + "：checkpoints idx=%s 缺证据 evidence" % c.get("idx"))
                    break
        sc = o.get("scores")
        if not isinstance(sc, dict) or sorted(sc) != sorted(DIMS):
            problems.append(L + "：scores 的维度应为 " + "、".join(DIMS))
        else:
            for k, v in sc.items():
                if not (isinstance(v, int) and not isinstance(v, bool) and 0 <= v <= 10):
                    problems.append(L + "：" + k + " 应为 0～10 的整数，实际 " + repr(v))
        if TRACK in ("W", "H"):
            for key in ("fabrications", "other_fabrications"):
                if not isinstance(o.get(key), list):
                    problems.append(L + "：" + key + " 应为数组（没有则为空数组）")
            for fb in o.get("fabrications", []) if isinstance(o.get("fabrications"), list) else []:
                if not isinstance(fb, dict) or fb.get("verdict") not in ("编造", "派生", "非事实"):
                    problems.append(L + "：fabrications 的 verdict 只能是 编造、派生、非事实")
                    break
    rk = g.get("ranking")
    if not isinstance(rk, list) or sorted(rk) != sorted(LETTERS):
        problems.append("ranking 应恰好包含 " + "、".join(LETTERS) + " 各一次，实际 " + repr(rk))
    if problems:
        print("校验未通过：")
        for p in problems:
            print(" - " + p)
        return 1
    print("校验通过：%d 份输出，排名 %s" % (len(LETTERS), ">".join(rk)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
'''


def checkpoints_block(task):
    lines = []
    for i, c in enumerate(task.get("checkpoints", [])):
        kind = c.get("kind")
        lines.append(f"{i}. {c['text']}" + (f"（{kind}）" if kind else ""))
    return "\n".join(lines)


def output_block(bench, letter, row, m, track):
    """一份匿名输出的说明块：回复、文书全文、渲染图片的路径，以及内联的机检摘要。"""
    s = Path(bench.root) / "sanitized" / row["hash"]
    L = [f"### 输出 {letter}", ""]
    reply = s / "_回复.md"
    L.append(f"- 回复：{reply}" if reply.is_file() else "- 回复：（没有 _回复.md）")
    if track in ("W", "H"):
        bodies = sorted((s / "_正文").glob("*")) if (s / "_正文").is_dir() else []
        if bodies:
            L.append("- 文书全文：" + "、".join(str(b) for b in bodies))
        else:
            L.append("- 文书全文：无（没有交付文书全文，以回复为准）")
    else:
        prim = next((d for d in (m.get("docx") or []) if d.get("primary")), None)
        imgs = (prim or {}).get("render", {}).get("images") if prim else None
        if imgs:
            L.append("- 渲染图片（LibreOffice 渲染，150dpi；第 1～3 页和最后一页，用 Read 工具查看）：")
            for im in imgs:
                L.append(f"  - 第 {im['page']} 页：{Path(bench.root) / im['path']}")
        else:
            why = ((prim or {}).get("render") or {}).get("error", "没有交付 .docx 文件")
            L.append(f"- 渲染图片：无（{why}）")
    L.append("- 机检摘要：")
    L.append("")
    summ = (s / "机检摘要.md").read_text(encoding="utf-8") if (s / "机检摘要.md").is_file() else "（缺少机检摘要）"
    L.append(re.sub(r"^", "    ", summ.rstrip(), flags=re.M))
    L.append("")
    return "\n".join(L)


def fill(tpl, mapping):
    for k, v in mapping.items():
        tpl = tpl.replace("{" + k + "}", str(v))
    return tpl


def main():
    ap = argparse.ArgumentParser(description="生成 Claude 评审任务")
    ap.add_argument("--bench", required=True, help="评测目录")
    ap.add_argument("--seed", type=int, default=20261021)
    ap.add_argument("--only", default="", help="只处理这些题号（逗号分隔）")
    ap.add_argument("--reshuffle", action="store_true", help="忽略已有 blind.json，重新随机分配字母")
    args = ap.parse_args()
    bench = Bench(args.bench)
    if not bench.index:
        print("评测目录里没有 index.json，请先运行 prep.py", file=sys.stderr)
        return 2
    only = {x for x in args.only.split(",") if x}
    jdir = bench.root / "judge"
    for sub in ("tasks", "validate"):
        (jdir / sub).mkdir(parents=True, exist_ok=True)
    (bench.root / "grades" / "claude").mkdir(parents=True, exist_ok=True)
    (bench.root / "grades" / "deepseek").mkdir(parents=True, exist_ok=True)
    blind = {} if args.reshuffle else (read_json(jdir / "blind.json", {}) or {})
    tpls = {t: (COMPETE / "templates" / f"judge-{t}.md").read_text(encoding="utf-8") for t in "WFH"}

    groups = {}
    for r in bench.index:
        if only and r["task"] not in only:
            continue
        groups.setdefault((r["task"], r["rep"]), []).append(r)
    batches, skipped = [], []
    for (tid, rep), rows in sorted(groups.items(), key=lambda kv: (natural_key(kv[0][0]), kv[0][1])):
        if len(rows) < 2:
            skipped.append(f"{tid}-r{rep}（只有 {len(rows)} 份输出）")
            continue
        gkey = f"{tid}-r{rep}"
        members = sorted((r["contender"], r["hash"]) for r in rows)
        old = blind.get(gkey)
        if old and sorted((v["contender"], v["hash"]) for v in old.values()) == members:
            mapping = old
        else:
            rng = random.Random(f"{args.seed}|{tid}|{rep}")
            order = [r for r in rows]
            order.sort(key=lambda r: r["contender"])  # 先定序再洗牌，保证同一种子可复现
            rng.shuffle(order)
            mapping = {LETTERS[i]: {"contender": r["contender"], "hash": r["hash"]} for i, r in enumerate(order)}
            blind[gkey] = mapping
        task = bench.task(tid)
        track = task["track"]
        by_hash = {r["hash"]: r for r in rows}
        blocks, missing, layouts = [], [], []
        for letter, v in mapping.items():
            m = bench.machine(v["hash"])
            if m is None or "error" in m:
                missing.append(v["hash"])
                m = m or {}
            layouts.append(m.get("layout"))
            blocks.append(output_block(bench, letter, by_hash[v["hash"]], m, track))
        if missing:
            print(f"{gkey}：缺少 machine 结果（{len(missing)} 份），请先运行 machine.py", file=sys.stderr)
        letters = list(mapping)
        grade_path = bench.root / "grades" / "claude" / f"{gkey}.json"
        validator = jdir / "validate" / f"{gkey}.py"
        validator.write_text(VALIDATOR.replace("@GRADE@", str(grade_path)).replace("@LETTERS@", json.dumps(letters))
                             .replace("@NCP@", str(len(task.get("checkpoints", [])))).replace("@DIMS@", json.dumps(DIMENSIONS[track], ensure_ascii=False))
                             .replace("@TRACK@", track), encoding="utf-8")
        lay = next((x for x in layouts if x), None) or task_layout(task)  # 优先用机检实际采用的版式（auto 时是按材料判断的结果）
        text = fill(tpls[track], {
            "task_id": tid, "rep": rep, "title": task.get("title", ""), "genre": task.get("genre", ""),
            "layout": "按材料判断" if lay == "auto" else lay, "user_message": task["user_message"],
            "inputs_path": bench.root / "inputs" / f"{tid}.txt", "n_checkpoints": len(task.get("checkpoints", [])),
            "n_outputs": len(letters), "letters": "、".join(letters), "checkpoints": checkpoints_block(task),
            "outputs": "\n".join(blocks), "dim_defs": dim_defs_text(track), "dims_json": dims_json_example(track),
            "grade_path": grade_path, "validator": validator})
        text = text.replace("<!--RULES-->\n", "").replace("<!--/RULES-->\n", "")  # 共享规则块的标记只给 DeepSeek 用
        ipath = jdir / "tasks" / f"{gkey}.md"
        ipath.write_text(text, encoding="utf-8")
        batches.append({"batch_id": gkey, "instruction_path": str(ipath), "grade_path": str(grade_path),
                        "validator": str(validator), "task": tid, "rep": rep, "track": track, "n_outputs": len(letters)})
    write_json(jdir / "blind.json", blind)
    write_json(jdir / "judge_batches.json", batches)
    print(f"评审任务 {len(batches)} 个（每个一个子智能体；说明在 {jdir / 'tasks'}，清单 judge_batches.json）")
    if skipped:
        print("跳过（输出不足两份）：" + "、".join(skipped))
    return 0


if __name__ == "__main__":
    sys.exit(main())
