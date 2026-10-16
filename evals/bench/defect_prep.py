"""缺陷归因准备：解盲后，按 skill 列出 A 组（装 skill）未通过的期望，生成归因任务说明。

用法：
    python evals/bench/defect_prep.py --bench <评测目录> [--skill gongwen-report ...]

产出：
    defects/<skill>.md   归因任务说明（含 A 组失败项、对应 B 组结果、A 组工作目录）
    defects/<skill>.json 由归因代理写入
"""
import argparse
import json
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parent))
from skill_list import ordered_unique  # noqa: E402

HEADER = """# 缺陷归因任务：{skill}

背景：我们用真实模型对 skill「{skill}」的评测用例做了"装 skill（A 组）与不装 skill（B 组）"对比。下面列出 A 组没有通过的期望（expectations）。A 组运行时先读了 SKILL.md 及其引用的 references，再按用户消息作答。你的任务是逐条判断失败原因，找出 skill 说明本身的缺陷。

只读，不要修改 D:\\公文写作skill 仓库中的任何文件。

## 你可以读取
- skill 目录：{skill_dir}（SKILL.md、references、scripts；也可以看 evals\\evals.json 了解用例全貌）
- 每条失败项给出的 A 组工作目录：`_回复.md`（给用户的回复）、`_过程.md`（读了哪些文件、跑了哪些命令）、`_正文/` 和其他交付文件
- 机检摘要：{bench}\\machine\\<skill>-<id>-A.md

## 归因类别（每条失败项选一个）
- "skill缺陷"：SKILL.md 或 references 没写、写得含糊、前后矛盾，或者规则与期望冲突，导致模型照着做也会失败；
- "执行偏差"：skill 写得清楚，模型没照做（如没读到该读的 reference、漏了步骤）；如果是因为 SKILL.md 没有提示去读那份 reference，算"skill缺陷"；
- "期望缺陷"：期望本身不合理、过严、与 skill 设计或用户消息冲突，或与同一用例其他期望矛盾；
- "评分误判"：A 组输出其实满足了期望。

## 输出
写入 {bench}\\defects\\{skill}.json，UTF-8，格式：

```json
{{
  "skill": "{skill}",
  "items": [
    {{"case": 1, "idx": 8, "category": "skill缺陷", "location": "SKILL.md 第三步 / references/types/zongjie.md 第二节", "analysis": "不超过 120 字的原因分析，引用 skill 原文", "fix": "不超过 80 字的修改方向"}}
  ],
  "defects": [
    {{"title": "不超过 25 字的缺陷名", "location": "文件和章节", "cases": [1, 6], "description": "不超过 150 字", "fix": "不超过 100 字的修改方向", "severity": "高/中/低"}}
  ]
}}
```

items 覆盖下面全部失败项；defects 是归并后的 skill 说明缺陷清单（只放"skill缺陷"类，以及多个用例反复出现、值得在 skill 中加强提示的"执行偏差"），按严重程度排序。另外，如果你在阅读 A 组输出时发现了期望没有覆盖、但明显属于 skill 说明问题的情况（例如篇幅普遍偏离、脚本报错、规则互相冲突），也写进 defects，cases 写相关用例号。写完后用 Python 校验 JSON。

最终回复只写一行："{skill}：失败项 N 条，skill缺陷 X 条，执行偏差 Y 条，期望缺陷 Z 条，评分误判 W 条"。

## A 组失败项
"""


def main():
    ap = argparse.ArgumentParser(description="生成缺陷归因任务")
    ap.add_argument("--bench", required=True)
    ap.add_argument("--skill", nargs="*")
    args = ap.parse_args()
    bench = Path(args.bench).resolve()
    index = json.loads((bench / "index.json").read_text(encoding="utf-8"))
    blind = json.loads((bench / "blind.json").read_text(encoding="utf-8"))
    dirs = {(r["skill"], r["id"], r["group"]): r["dir"] for r in index}
    (bench / "defects").mkdir(exist_ok=True)

    for skill in ordered_unique([r["skill"] for r in index], REPO):
        if args.skill and skill not in args.skill:
            continue
        data = json.loads((REPO / "skills" / skill / "evals" / "evals.json").read_text(encoding="utf-8"))
        blocks, n = [], 0
        for e in data["evals"]:
            key = f"{skill}-{e['id']}"
            g = json.loads((bench / "grades" / f"{key}.json").read_text(encoding="utf-8"))
            inv = {v: k for k, v in blind[key].items()}
            a = {x["idx"]: x for x in g["outputs"][inv["A"]]["expectations"]}
            b = {x["idx"]: x for x in g["outputs"][inv["B"]]["expectations"]}
            fails = [i for i in range(len(e["expectations"])) if not a.get(i, {}).get("pass")]
            if not fails:
                continue
            lines = [f"### 用例 {e['id']}（{e['unit_type']}，{e['role_mode']}）", "",
                     f"用户消息：{e['prompt']}", "",
                     f"A 组工作目录：{dirs[(skill, e['id'], 'A')]}", ""]
            for i in fails:
                n += 1
                lines.append(f"- idx {i}：{e['expectations'][i]}")
                lines.append(f"  - A 组判定证据：{a.get(i, {}).get('evidence', '（缺）')}")
                bv = b.get(i, {})
                lines.append(f"  - B 组（不装 skill）：{'通过' if bv.get('pass') else '未通过'}；{bv.get('evidence', '')}")
            blocks.append("\n".join(lines) + "\n")
        text = HEADER.format(skill=skill, skill_dir=REPO / "skills" / skill, bench=bench) + "\n".join(blocks)
        (bench / "defects" / f"{skill}.md").write_text(text, encoding="utf-8")
        print(f"{skill}: A 组失败项 {n} 条")
    return 0


if __name__ == "__main__":
    sys.exit(main())
