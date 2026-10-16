"""评测汇总：解盲评分结果，计算通过率、编造次数、机检问题数，输出原始结果 JSON 和统计表。

用法：
    python evals/bench/aggregate.py --bench <评测目录> --out evals/results/v0.1-raw.json [--tables 统计表.md]
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

REPO = Path(__file__).resolve().parents[2]
SKILLS = ["gongwen", "gongwen-report", "gongwen-title", "gongwen-humanize",
          "gongwen-direct", "gongwen-review", "gongwen-format"]
GROUPS = ("A", "B")


def pct(a, b):
    return f"{a / b * 100:.1f}%" if b else "—"


def build_cases(bench):
    """从评测目录解盲，组装每条用例两组的结果。"""
    index = json.loads((bench / "index.json").read_text(encoding="utf-8"))
    blind = json.loads((bench / "blind.json").read_text(encoding="utf-8"))
    classes = {(c["skill"], c["id"], c["idx"]): c["class"]
               for c in json.loads((bench / "exp_class.json").read_text(encoding="utf-8"))}
    evals = {}
    for s in SKILLS:
        data = json.loads((REPO / "skills" / s / "evals" / "evals.json").read_text(encoding="utf-8"))
        evals[s] = {e["id"]: e for e in data["evals"]}

    model_lines = defaultdict(int)
    for row in index:
        mpath = bench / "machine" / f"{row['skill']}-{row['id']}-{row['group']}.json"
        if mpath.exists():
            model_lines[json.loads(mpath.read_text(encoding="utf-8")).get("model_line", "")] += 1

    cases, missing = [], []
    for s in SKILLS:
        for eid, e in evals[s].items():
            key = f"{s}-{eid}"
            gpath = bench / "grades" / f"{key}.json"
            if not gpath.exists() or key not in blind:
                missing.append(key)
                continue
            grade = json.loads(gpath.read_text(encoding="utf-8"))
            inv = {v: k for k, v in blind[key].items()}  # A/B -> 甲/乙
            case = {"skill": s, "id": eid, "prompt": e["prompt"], "tags": e["tags"],
                    "unit_type": e["unit_type"], "role_mode": e["role_mode"], "runs": {}}
            for g in GROUPS:
                label = inv[g]
                out = grade["outputs"][label]
                m = json.loads((bench / "machine" / f"{key}-{g}.json").read_text(encoding="utf-8"))
                verdicts = {v["idx"]: v for v in out["expectations"]}
                exps = []
                for i, text in enumerate(e["expectations"]):
                    v = verdicts.get(i)
                    if v is None:
                        missing.append(f"{key}-{g}-exp{i}")
                    exps.append({"idx": i, "text": text, "class": classes.get((s, eid, i), "未分类"),
                                 "pass": bool(v and v.get("pass")), "evidence": (v or {}).get("evidence", "")})
                fabs = out.get("fabrications", [])
                lint_sum = [l.get("summary") for l in m["lint"] if "summary" in l]
                case["runs"][g] = {
                    "blind_label": label,
                    "dir_token": Path(m["dir"]).name,
                    "expectations": exps,
                    "numeric_fabrications": [f for f in fabs if f.get("verdict") == "编造"],
                    "numeric_added_total": sum(len(f.get("added", [])) for f in m["fact_diff"] if "added" in f),
                    "other_fabrications": out.get("other_fabrications", []),
                    "grader_notes": out.get("notes", ""),
                    "bodies": m["bodies"],
                    "lint": [{"file": l.get("file"), "doc_type": l.get("doc_type"), "char_count": l.get("char_count"),
                              "body_char_count": l.get("body_char_count"), "summary": l.get("summary"),
                              "p0_p1": [f"{i['level']} {i['id']} {i['category']}" for i in l.get("issues", [])
                                        if i["level"] in ("P0", "P1")]} for l in m["lint"]],
                    "check_docx": [{"file": Path(c["file"]).name, "summary": c.get("summary")}
                                   for c in m["check_docx"] if "summary" in c],
                    "produced": m["produced"],
                    "inputs_changed": m["inputs_changed"],
                    "repo_mentions": m.get("repo_mentions", []),
                }
            cases.append(case)
    return cases, missing, dict(model_lines)


def summarize(cases, model_lines):
    """由逐条判定计算全部统计数字。只依赖 cases，给定原始结果 JSON 即可复算。"""
    def tally(filter_case=lambda c: True, filter_exp=lambda x: True):
        r = {g: [0, 0] for g in GROUPS}
        for c in cases:
            if not filter_case(c):
                continue
            for g in GROUPS:
                for x in c["runs"][g]["expectations"]:
                    if filter_exp(x):
                        r[g][1] += 1
                        r[g][0] += x["pass"]
        return r

    summary = {}
    summary["overall"] = tally()
    summary["general_only"] = tally(filter_exp=lambda x: x["class"] == "通用")
    summary["convention_only"] = tally(filter_exp=lambda x: x["class"] == "约定")
    summary["by_skill"] = {s: tally(lambda c, s=s: c["skill"] == s) for s in SKILLS}
    summary["by_skill_general"] = {s: tally(lambda c, s=s: c["skill"] == s, lambda x: x["class"] == "通用") for s in SKILLS}
    for tag in ("防编造", "保密"):
        summary[f"tag_{tag}"] = tally(lambda c, t=tag: t in c["tags"])

    # 用例全部通过
    full = {g: sum(all(x["pass"] for x in c["runs"][g]["expectations"]) for c in cases) for g in GROUPS}
    summary["cases_all_pass"] = {g: [full[g], len(cases)] for g in GROUPS}

    # 配对差异：同一条期望 A 过 B 不过 / B 过 A 不过
    a_only = b_only = 0
    for c in cases:
        for xa, xb in zip(c["runs"]["A"]["expectations"], c["runs"]["B"]["expectations"]):
            a_only += xa["pass"] and not xb["pass"]
            b_only += xb["pass"] and not xa["pass"]
    summary["paired"] = {"A_only": a_only, "B_only": b_only}

    # 编造
    def fab(filter_case):
        r = {}
        for g in GROUPS:
            num = oth = runs_with = 0
            for c in cases:
                if not filter_case(c):
                    continue
                run = c["runs"][g]
                n1, n2 = len(run["numeric_fabrications"]), len(run["other_fabrications"])
                num += n1
                oth += n2
                runs_with += (n1 + n2) > 0
            r[g] = {"numeric": num, "other": oth, "runs_with_fabrication": runs_with,
                    "runs": sum(1 for c in cases if filter_case(c))}
        return r

    summary["fabrication_anti_cases"] = fab(lambda c: "防编造" in c["tags"])
    summary["fabrication_all_cases"] = fab(lambda c: True)

    # 机检（文字类 skill 的正文）
    def lint_tally(filter_case):
        r = {}
        for g in GROUPS:
            t = {"runs_with_body": 0, "bodies": 0, "P0": 0, "P1": 0, "P2": 0, "待补": 0, "待核实": 0,
                 "runs_with_P0": 0}
            for c in cases:
                if not filter_case(c):
                    continue
                run = c["runs"][g]
                sums = [l["summary"] for l in run["lint"] if l.get("summary")]
                if not sums:
                    continue
                t["runs_with_body"] += 1
                t["bodies"] += len(sums)
                p0 = 0
                for s in sums:
                    for k in ("P0", "P1", "P2"):
                        t[k] += s.get(k, 0)
                    p0 += s.get("P0", 0)
                    t["待补"] += s.get("markers", {}).get("待补", 0)
                    t["待核实"] += s.get("markers", {}).get("待核实", 0)
                t["runs_with_P0"] += p0 > 0
            # 【待补】【待核实】标记每处报 P2，单列不含标记的 P2
            t["P2_非标记"] = t["P2"] - t["待补"] - t["待核实"]
            r[g] = t
        return r

    both_body = lambda c: c["skill"] != "gongwen-format" and all(
        any(l.get("summary") for l in c["runs"][g]["lint"]) for g in GROUPS)
    summary["lint_all"] = lint_tally(lambda c: c["skill"] != "gongwen-format")
    summary["lint_paired"] = lint_tally(both_body)
    summary["lint_paired_cases"] = sum(1 for c in cases if both_body(c))

    # Word 格式检查
    docx = {}
    for g in GROUPS:
        t = {"runs_with_docx": 0, "docx": 0, "P0": 0, "P1": 0, "P2": 0, "docx_P0_zero": 0}
        for c in cases:
            if c["skill"] != "gongwen-format":
                continue
            ds = c["runs"][g]["check_docx"]
            if ds:
                t["runs_with_docx"] += 1
            for d in ds:
                s = d["summary"] or {}
                t["docx"] += 1
                for k in ("P0", "P1", "P2"):
                    t[k] += s.get(k, 0)
                t["docx_P0_zero"] += s.get("P0", 0) == 0
        docx[g] = t
    summary["check_docx"] = docx
    summary["inputs_changed_runs"] = {g: sum(bool(c["runs"][g]["inputs_changed"]) for c in cases) for g in GROUPS}
    summary["B_repo_mentions"] = [f"{c['skill']}-{c['id']}" for c in cases if c["runs"]["B"]["repo_mentions"]]
    summary["model_lines"] = dict(model_lines)
    return summary


def render(cases, missing, summary):
    """输出统计表文本。"""
    a_only, b_only = summary["paired"]["A_only"], summary["paired"]["B_only"]
    docx = summary["check_docx"]
    lines = []
    lines.append(f"用例 {len(cases)} 条；缺失 {len(missing)} 项")
    lines.append("")
    lines.append("| 口径 | A 组（装 skill） | B 组（不装） |")
    lines.append("|---|---|---|")
    for name, key in (("全部期望", "overall"), ("通用期望", "general_only"), ("约定期望", "convention_only"),
                      ("防编造用例", "tag_防编造"), ("保密用例", "tag_保密")):
        t = summary[key]
        lines.append(f"| {name} | {t['A'][0]}/{t['A'][1]}（{pct(*t['A'])}） | {t['B'][0]}/{t['B'][1]}（{pct(*t['B'])}） |")
    ca = summary["cases_all_pass"]
    lines.append(f"| 用例全部通过 | {ca['A'][0]}/{ca['A'][1]}（{pct(*ca['A'])}） | {ca['B'][0]}/{ca['B'][1]}（{pct(*ca['B'])}） |")
    lines.append("")
    lines.append(f"配对：A 过 B 不过 {a_only} 条；B 过 A 不过 {b_only} 条")
    lines.append("")
    lines.append("| skill | 用例 | A 全部 | B 全部 | A 通用 | B 通用 |")
    lines.append("|---|---|---|---|---|---|")
    for s in SKILLS:
        t, u = summary["by_skill"][s], summary["by_skill_general"][s]
        n = sum(1 for c in cases if c["skill"] == s)
        lines.append(f"| {s} | {n} | {pct(*t['A'])}（{t['A'][0]}/{t['A'][1]}） | {pct(*t['B'])}（{t['B'][0]}/{t['B'][1]}） | "
                     f"{pct(*u['A'])}（{u['A'][0]}/{u['A'][1]}） | {pct(*u['B'])}（{u['B'][0]}/{u['B'][1]}） |")
    lines.append("")
    for name, key in (("防编造用例", "fabrication_anti_cases"), ("全部用例", "fabrication_all_cases")):
        f = summary[key]
        lines.append(f"编造（{name}）：A 数字编造 {f['A']['numeric']}、非数字编造 {f['A']['other']}、有编造的运行 "
                     f"{f['A']['runs_with_fabrication']}/{f['A']['runs']}；B 数字编造 {f['B']['numeric']}、非数字编造 "
                     f"{f['B']['other']}、有编造的运行 {f['B']['runs_with_fabrication']}/{f['B']['runs']}")
    lines.append("")
    for name, key in (("全部正文", "lint_all"), (f"双方都有正文的 {summary['lint_paired_cases']} 个用例", "lint_paired")):
        t = summary[key]
        lines.append(f"机检（{name}）：A {t['A']}；B {t['B']}")
    lines.append(f"check_docx：A {docx['A']}；B {docx['B']}")
    lines.append(f"输入被改动的运行：{summary['inputs_changed_runs']}；B 组过程记录提到仓库路径：{summary['B_repo_mentions']}")
    lines.append(f"模型：{summary['model_lines']}")
    return "\n".join(lines) + "\n"


def main():
    ap = argparse.ArgumentParser(description="评测汇总：从评测目录解盲汇总，或从原始结果 JSON 复算统计")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--bench", help="评测目录（含 index.json、blind.json、grades、machine、exp_class.json）")
    src.add_argument("--from-raw", help="已有的原始结果 JSON，只复算统计，不需要评测目录")
    ap.add_argument("--out", help="原始结果 JSON 输出路径（--bench 时必填）")
    ap.add_argument("--tables", help="统计表 Markdown 输出路径")
    args = ap.parse_args()

    if args.from_raw:
        raw = json.loads(Path(args.from_raw).read_text(encoding="utf-8"))
        cases, missing = raw["cases"], raw.get("missing", [])
        summary = summarize(cases, raw["summary"].get("model_lines", {}))
        if summary != raw["summary"]:
            print("注意：复算结果与文件中保存的 summary 不一致", file=sys.stderr)
    else:
        if not args.out:
            ap.error("使用 --bench 时必须给出 --out")
        cases, missing, model_lines = build_cases(Path(args.bench).resolve())
        summary = summarize(cases, model_lines)
        raw = {"version": "v0.1", "date": "2026-10-16", "cases_total": len(cases),
               "missing": missing, "summary": summary, "cases": cases}
        Path(args.out).parent.mkdir(parents=True, exist_ok=True)
        Path(args.out).write_text(json.dumps(raw, ensure_ascii=False, indent=1), encoding="utf-8")

    text = render(cases, missing, summary)
    if args.tables:
        Path(args.tables).write_text(text, encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
