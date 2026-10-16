"""评测准备：为每条用例的 A 组（装 skill）和 B 组（不装 skill）建立独立工作目录，生成子智能体任务说明。

用法：
    python evals/bench/prep.py --out <评测目录> [--chunk 4] [--skills 名称或预设 ...] [--tag v1.0]

    --skills 默认取插件市场清单中的全部 skill；写 v0.1 复现 v0.1 的 7 个 skill。
    用例可以用 file_dest 字段指定材料放在工作目录中的位置，如 {"evals/files/已有画像.md": ".gongwen/单位画像.md"}。

产出（都在 <评测目录> 下）：
    w/<编号>/            每次运行的工作目录，已复制输入材料；编号为不透明的哈希，避免从路径看出分组
    inputs/<skill>-<id>.txt   用户消息 + 全部输入材料的纯文本，供 fact_diff 判断"材料中没有的数字"
    tasks/<批次>-A.md、tasks/<批次>-B.md   子智能体任务说明（不含 expectations）
    index.json           用例、分组、工作目录、批次的对应关系
    hashes.json          输入材料的 SHA-256，用于检查"原文件未被修改"
"""
import argparse
import hashlib
import json
import shutil
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

REPO = Path(__file__).resolve().parents[2]
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from skill_list import select_skills  # noqa: E402


def token(skill, eid, group, tag="v0.1"):
    """生成不透明的工作目录编号（确定性，便于复现）。"""
    return hashlib.sha1(f"{skill}|{eid}|{group}|{tag}".encode("utf-8")).hexdigest()[:10]


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def file_text(path):
    """把输入材料转成纯文本，供 fact_diff 判断"材料中没有的数字"。

    docx 用 shared/scripts/fact_diff.py 的读取方法（接受全部修订：计入插入、不计删除），与机检口径一致；
    xlsx 逐表逐行取单元格文字（需要 openpyxl）；其余按 UTF-8 读取。
    """
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".docx":
        sys.path.insert(0, str(REPO / "shared" / "scripts"))
        import fact_diff  # noqa: E402
        return fact_diff.read_text(path)
    if suffix == ".xlsx":
        try:
            import openpyxl
        except ImportError:
            print("缺少 openpyxl，请先 pip install openpyxl", file=sys.stderr)
            sys.exit(2)
        wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
        lines = []
        for ws in wb.worksheets:
            lines.append(f"【工作表：{ws.title}】")
            for row in ws.iter_rows(values_only=True):
                cells = ["" if v is None else str(v) for v in row]
                if any(cells):
                    lines.append("\t".join(cells))
        return "\n".join(lines)
    return path.read_text(encoding="utf-8", errors="replace")


def load_cases(skills):
    cases = []
    for skill in skills:
        data = json.loads((REPO / "skills" / skill / "evals" / "evals.json").read_text(encoding="utf-8"))
        for e in data["evals"]:
            cases.append({"skill": skill, **e})
    return cases


def main():
    ap = argparse.ArgumentParser(description="准备装与不装 skill 的对比评测")
    ap.add_argument("--out", required=True, help="评测目录（建议放在仓库之外）")
    ap.add_argument("--chunk", type=int, default=4, help="每个子智能体处理的用例数，默认 4")
    ap.add_argument("--skills", nargs="*", help="只评这些 skill（名称或预设 v0.1），默认全部")
    ap.add_argument("--tag", default="v1.0", help="评测版本标签，用于生成工作目录编号和汇总，默认 v1.0")
    args = ap.parse_args()
    skills = select_skills(REPO, args.skills)

    out = Path(args.out).resolve()
    if REPO in out.parents or out == REPO:
        print("评测目录不要放在仓库内，以免 B 组读到 skill 文件", file=sys.stderr)
        return 2
    for sub in ("w", "inputs", "tasks"):
        (out / sub).mkdir(parents=True, exist_ok=True)

    common = (HERE / "templates" / "common.md").read_text(encoding="utf-8")
    tpl = {g: (HERE / "templates" / f"run-{g}.md").read_text(encoding="utf-8") for g in "AB"}

    cases = load_cases(skills)
    index, hashes = [], {}
    by_skill = {}
    for c in cases:
        by_skill.setdefault(c["skill"], []).append(c)

    for skill, items in by_skill.items():
        chunks = [items[i:i + args.chunk] for i in range(0, len(items), args.chunk)]
        for ci, chunk in enumerate(chunks, start=1):
            chunk_id = f"{skill}-{ci}"
            blocks = {"A": [], "B": []}
            for k, c in enumerate(chunk, start=1):
                eid = c["id"]
                # 全部输入材料的纯文本（含用户消息本身，用户口头给出的数字不算编造）
                parts = [f"【用户消息】\n{c['prompt']}"]
                for rel in c["files"]:
                    src = REPO / "skills" / skill / rel
                    parts.append(f"【材料：{Path(rel).name}】\n{file_text(src)}")
                (out / "inputs" / f"{skill}-{eid}.txt").write_text("\n\n".join(parts) + "\n", encoding="utf-8")

                for g in "AB":
                    wdir = out / "w" / token(skill, eid, g, args.tag)
                    if wdir.exists():
                        shutil.rmtree(wdir)
                    wdir.mkdir(parents=True)
                    shown = []
                    for rel in c["files"]:
                        src = REPO / "skills" / skill / rel
                        dest_rel = (c.get("file_dest") or {}).get(rel) or Path(rel).name
                        dest = wdir / dest_rel
                        dest.parent.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(src, dest)
                        hashes[str(dest)] = sha256(dest)
                        shown.append(dest_rel)
                    files_line = "、".join(shown) if shown else "无"
                    blocks[g].append(
                        f"---\n\n## 请求 {k}\n\n工作目录：{wdir}\n\n用户消息：\n\n<<<\n{c['prompt']}\n>>>\n\n"
                        f"用户提供的文件（都在工作目录中）：{files_line}\n"
                    )
                    index.append({"skill": skill, "id": eid, "group": g, "chunk": chunk_id,
                                  "order": k, "dir": str(wdir), "files": shown})

            for g in "AB":
                body = common.replace("{cases}", "\n".join(blocks[g])).replace("{n}", str(len(chunk)))
                text = tpl[g].replace("{common}", body).replace("{n}", str(len(chunk)))
                text = text.replace("{skill}", skill).replace("{skill_dir}", str(REPO / "skills" / skill))
                (out / "tasks" / f"{chunk_id}-{g}.md").write_text(text, encoding="utf-8")

    (out / "index.json").write_text(json.dumps(index, ensure_ascii=False, indent=1), encoding="utf-8")
    (out / "hashes.json").write_text(json.dumps(hashes, ensure_ascii=False, indent=1), encoding="utf-8")
    (out / "meta.json").write_text(json.dumps({"tag": args.tag, "skills": skills}, ensure_ascii=False, indent=1),
                                   encoding="utf-8")
    chunks = sorted({r["chunk"] for r in index})
    print(f"用例 {len(cases)} 条，运行 {len(index)} 次，批次 {len(chunks)} 个（每批 A、B 各一个子智能体）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
