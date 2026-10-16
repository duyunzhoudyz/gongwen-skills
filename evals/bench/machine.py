"""评测机检：对每次运行的产出跑 gongwen_lint、fact_diff（--ledger）和 check_docx，并检查输入材料是否被改动。

用法：
    python evals/bench/machine.py --bench <评测目录> [--chunk 批次名 ...]

产出：
    machine/<skill>-<id>-<A|B>.json   完整机检结果
    machine/<skill>-<id>-<A|B>.md     给评分员看的机检摘要（不含分组信息）
"""
import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

REPO = Path(__file__).resolve().parents[2]
LINT = REPO / "shared" / "scripts" / "gongwen_lint.py"
FACT = REPO / "shared" / "scripts" / "fact_diff.py"
CHECK = REPO / "skills" / "gongwen-format" / "scripts" / "check_docx.py"
ENV = {**os.environ, "PYTHONIOENCODING": "utf-8"}
BODY_RE = re.compile(r"约 (\d+) 字")
# 运行产生的记录文件，不算交付物
HARNESS_FILES = {"_回复.md", "_过程.md"}


def run_json(cmd):
    """运行脚本并解析 JSON 输出；失败时返回错误说明。"""
    try:
        p = subprocess.run([sys.executable, *map(str, cmd)], capture_output=True, text=True,
                           encoding="utf-8", errors="replace", env=ENV, timeout=180)
    except subprocess.TimeoutExpired:
        return {"error": "超时"}
    try:
        return json.loads(p.stdout)
    except json.JSONDecodeError:
        return {"error": (p.stderr or p.stdout)[-400:]}


def lint(path):
    r = run_json([LINT, path, "--json"])
    if "error" in r:
        return r
    body = None
    for note in r.get("notes", []):
        m = BODY_RE.search(note)
        if m and "篇幅口径" in note:
            body = int(m.group(1))
    return {
        "file": Path(path).name,
        "doc_type": r.get("doc_type"),
        "doc_type_source": r.get("doc_type_source"),
        "char_count": r.get("char_count"),
        "body_char_count": body,
        "summary": r.get("summary"),
        "issues": [{k: i.get(k) for k in ("id", "level", "category", "line", "excerpt", "message")}
                   for i in r.get("issues", [])],
    }


def fact(inputs, path):
    r = run_json([FACT, "--ledger", inputs, path, "--json"])
    if "error" in r:
        return r
    added = [{"type": a["type"], "value": a["value"],
              "excerpt": (a.get("locations") or [{}])[0].get("excerpt", "")}
             for a in r.get("added", []) if a["type"] != "标记"]
    markers = [a["value"] for a in r.get("added", []) if a["type"] == "标记"]
    return {"file": Path(path).name, "added": added, "new_markers": len(markers)}


def check_docx(path):
    r = run_json([CHECK, path, "--json"])
    if "error" in r:
        return r
    return {"file": str(path), "summary": r.get("summary"), "layout": r.get("layout") or r.get("doc_type"),
            "issues": [{k: i.get(k) for k in ("id", "level", "category", "excerpt", "message")}
                       for i in r.get("issues", []) if i.get("level") in ("P0", "P1")]}


def process(bench, row, hashes):
    wdir = Path(row["dir"])
    inputs = bench / "inputs" / f"{row['skill']}-{row['id']}.txt"
    input_paths = {os.path.normcase(str(wdir / f)) for f in row["files"]}
    res = {"skill": row["skill"], "id": row["id"], "group": row["group"], "dir": str(wdir)}
    res["has_reply"] = (wdir / "_回复.md").is_file()
    res["has_log"] = (wdir / "_过程.md").is_file()

    # 输入材料是否被改动
    changed = []
    for f in row["files"]:
        p = wdir / f
        key = str(p)
        if not p.is_file():
            changed.append(f"{f}（已不存在）")
        elif key in hashes and hashlib.sha256(p.read_bytes()).hexdigest() != hashes[key]:
            changed.append(f)
    res["inputs_changed"] = changed

    # 交付文件清单（不含输入材料和运行记录）
    produced = []
    for p in sorted(wdir.rglob("*")):
        if p.is_file() and os.path.normcase(str(p)) not in input_paths and p.name not in HARNESS_FILES:
            produced.append(p.relative_to(wdir).as_posix())
    res["produced"] = produced

    bodies = sorted((wdir / "_正文").glob("*.txt")) if (wdir / "_正文").is_dir() else []
    res["bodies"] = [b.name for b in bodies]
    res["lint"] = [lint(b) for b in bodies]

    docx_out = [wdir / p for p in produced if p.lower().endswith(".docx")]
    res["check_docx"] = [check_docx(d) for d in docx_out]

    # 编造核查对象：有正文时查正文；格式类再加生成的 docx；都没有时查回复全文
    targets = list(bodies) + docx_out
    if not targets and res["has_reply"]:
        targets = [wdir / "_回复.md"]
    res["fact_targets"] = [str(t.relative_to(wdir)) for t in targets]
    res["fact_diff"] = [fact(inputs, t) for t in targets]

    # B 组是否读过仓库（自述过程中出现仓库路径即标记，供人工复核）
    log = (wdir / "_过程.md").read_text(encoding="utf-8", errors="replace") if res["has_log"] else ""
    res["model_line"] = log.splitlines()[0][:80] if log else ""
    res["repo_mentions"] = sorted(set(re.findall(r"公文写作skill[\\/][^\s`'\"）)]*", log)))[:20]
    return res


def to_md(res):
    """给评分员的机检摘要，不出现分组信息。"""
    out = ["# 机检摘要", ""]
    out.append(f"- 输入文件被改动：{'、'.join(res['inputs_changed']) or '无'}")
    out.append(f"- 交付文件：{'、'.join(res['produced']) or '无'}")
    out.append("")
    out.append("## gongwen_lint（对 _正文 下的文书全文）")
    if not res["lint"]:
        out.append("无 _正文，未运行。")
    for l in res["lint"]:
        if "error" in l:
            out.append(f"- 运行失败：{l['error']}")
            continue
        s = l["summary"] or {}
        out.append(f"- {l['file']}：文书类型 {l['doc_type']}（{l['doc_type_source']}），全文 {l['char_count']} 字，"
                   f"正文（标题后、起草说明前）约 {l['body_char_count']} 字；P0={s.get('P0')} P1={s.get('P1')} "
                   f"P2={s.get('P2')} 待补={s.get('markers', {}).get('待补')} 待核实={s.get('markers', {}).get('待核实')}")
        for i in l["issues"]:
            if i["level"] in ("P0", "P1"):
                out.append(f"  - {i['level']} {i['id']} {i['category']}：{(i['excerpt'] or '')[:40]}｜{i['message']}")
    out.append("")
    out.append("## Word 格式检查（check_docx，对生成的 .docx）")
    if not res["check_docx"]:
        out.append("没有生成 docx。")
    for c in res["check_docx"]:
        if "error" in c:
            out.append(f"- 运行失败：{c['error']}")
            continue
        s = c["summary"] or {}
        out.append(f"- {Path(c['file']).name}：P0={s.get('P0')} P1={s.get('P1')} P2={s.get('P2')}")
        for i in c["issues"]:
            out.append(f"  - {i['level']} {i['id']} {i['category']}：{(i['excerpt'] or '')[:40]}｜{i['message']}")
    out.append("")
    out.append("## fact_diff --ledger：输出中出现、输入材料（含用户消息）中没有的数字")
    out.append(f"核查对象：{'、'.join(res['fact_targets']) or '无'}")
    for f in res["fact_diff"]:
        if "error" in f:
            out.append(f"- {f.get('file', '')} 运行失败：{f['error']}")
            continue
        out.append(f"- {f['file']}：新增数字 {len(f['added'])} 个，新增【待补】【待核实】标记 {f['new_markers']} 处")
        for a in f["added"]:
            out.append(f"  - {a['type']}｜{a['value']}｜{a['excerpt'][:60]}")
    return "\n".join(out) + "\n"


def main():
    ap = argparse.ArgumentParser(description="评测机检")
    ap.add_argument("--bench", required=True, help="评测目录")
    ap.add_argument("--chunk", nargs="*", help="只处理这些批次")
    args = ap.parse_args()
    bench = Path(args.bench).resolve()
    index = json.loads((bench / "index.json").read_text(encoding="utf-8"))
    hashes = json.loads((bench / "hashes.json").read_text(encoding="utf-8"))
    (bench / "machine").mkdir(exist_ok=True)
    n = 0
    for row in index:
        if args.chunk and row["chunk"] not in args.chunk:
            continue
        res = process(bench, row, hashes)
        stem = f"{row['skill']}-{row['id']}-{row['group']}"
        (bench / "machine" / f"{stem}.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
        (bench / "machine" / f"{stem}.md").write_text(to_md(res), encoding="utf-8")
        n += 1
        flag = "" if res["has_reply"] else "  缺少 _回复.md"
        print(f"{stem}: 正文 {len(res['bodies'])} 份，docx {len(res['check_docx'])} 个{flag}")
    print(f"共处理 {n} 次运行")
    return 0


if __name__ == "__main__":
    sys.exit(main())
