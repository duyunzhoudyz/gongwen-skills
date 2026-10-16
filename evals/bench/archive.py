"""归档评测产出：把每次运行的回复、过程记录和交付文件复制到仓库，便于复核。

用法：
    python evals/bench/archive.py --bench <评测目录> --out evals/results/v0.1-runs

处理规则：
- 目录名用 <skill>-<id>-<A|B>，A 组装 skill，B 组不装；
- 不复制输入材料副本，不复制预览图片和 PDF（可用 docx 重新生成），不复制运行中生成的辅助脚本；
- .gitignore 忽略 公文/ 和 .gongwen/，归档时改名为 _公文/ 和 _gongwen/；
- 文本文件中的本机绝对路径替换为 <评测目录>、<仓库>、<用户目录>，不公开本机用户名和目录结构；
- JSON 文件先校验，不合法的改存为 .json.txt，避免 tools/validate.py 报错。
"""
import argparse
import json
import os
import re
import shutil
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

SKIP_EXT = {".png", ".pdf", ".py", ".pyc"}
RENAME = {"公文": "_公文", ".gongwen": "_gongwen"}
TEXT_EXT = {".md", ".txt", ".json", ".csv"}
REPO = Path(__file__).resolve().parents[2]
SEP = r"[\\/]{1,2}"


def _path_regex(path):
    """把绝对路径转成兼容 \\、/、JSON 转义 \\\\ 和 /c/ 写法的正则。"""
    parts = Path(path).parts
    drive = parts[0].rstrip("\\/").rstrip(":")
    head = rf"(?:{re.escape(drive)}:|/{re.escape(drive.lower())})"
    return re.compile(head + "".join(SEP + re.escape(p) for p in parts[1:]), re.IGNORECASE)


def scrub(text, rules):
    for pat, repl in rules:
        text = pat.sub(repl, text)
    return text


def main():
    ap = argparse.ArgumentParser(description="归档评测产出")
    ap.add_argument("--bench", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    bench = Path(args.bench).resolve()
    out = Path(args.out).resolve()
    index = json.loads((bench / "index.json").read_text(encoding="utf-8"))
    # 先替换最长的路径：评测目录、仓库，再兜底替换用户目录
    rules = [(_path_regex(bench), "<评测目录>"), (_path_regex(REPO), "<仓库>"),
             (_path_regex(Path.home()), "<用户目录>")]
    n_files = n_bad_json = 0
    for row in index:
        wdir = Path(row["dir"])
        inputs = {os.path.normcase(str(wdir / f)) for f in row["files"]}
        dest_root = out / f"{row['skill']}-{row['id']}-{row['group']}"
        if dest_root.exists():
            shutil.rmtree(dest_root)
        for p in sorted(wdir.rglob("*")):
            if not p.is_file() or os.path.normcase(str(p)) in inputs:
                continue
            if p.suffix.lower() in SKIP_EXT or "__pycache__" in p.parts:
                continue
            rel = [RENAME.get(part, part) for part in p.relative_to(wdir).parts]
            dest = dest_root.joinpath(*rel)
            dest.parent.mkdir(parents=True, exist_ok=True)
            n_files += 1
            if p.suffix.lower() not in TEXT_EXT:
                shutil.copy2(p, dest)
                continue
            try:
                text = scrub(p.read_text(encoding="utf-8"), rules)
            except UnicodeDecodeError:
                shutil.copy2(p, dest)
                continue
            if p.suffix.lower() == ".json":
                try:
                    json.loads(text)
                except json.JSONDecodeError:
                    dest = dest.with_name(dest.name + ".txt")
                    n_bad_json += 1
            dest.write_text(text, encoding="utf-8")
    print(f"已归档 {len(index)} 次运行、{n_files} 个文件；不合法 JSON 改存 .txt {n_bad_json} 个")
    return 0


if __name__ == "__main__":
    sys.exit(main())
