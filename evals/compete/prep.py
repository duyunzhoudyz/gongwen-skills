"""竞品对标评测准备：复制各参赛方的 skill、建立匿名工作目录、分批生成执行任务说明。

用法：
    python evals/compete/prep.py --out <评测目录> [--reps 2] [--tracks W,F,H] [--tasks W01,F01,...]
                                 [--ours-src <仓库路径>] [--tasks-dir <题目目录>]
                                 [--competitors-dir <竞品克隆目录>] [--contenders ours,none,...]
                                 [--tag compete-v1] [--chunk 4] [--skip-setup] [--no-smoke] [--force]

做什么（都在 <评测目录> 下）：
    skills/<参赛方>/        各参赛方的 skill 副本。ours 是 22 个 skill（去掉各自的 evals/ 和 __pycache__），
                            复制前先运行 tools/sync_shared.py --check；竞品从克隆目录复制（去掉 .git）并执行 setup 命令
    runs/<12位哈希>/        每（题 × 参赛方 × 重复）一个匿名工作目录，已复制输入材料（不含 reference/）
    inputs/<题号>.txt       用户消息 + 全部输入材料的纯文本，供 fact_diff 判断“材料中没有的数字”
    taskdefs/<题号>/        题目快照（task.json、files/、reference/），之后的机检、评审、汇总都读它
    tasks/<批号>.md         子智能体执行任务说明（同一参赛方、同一赛道、同一重复轮次每 3～4 题一批）
    run_batches.json        批次清单：batch_id、instruction_path、contender、track、rep、cases
    index.json / hashes.json / meta.json   哈希与（题、参赛方、重复）的对应、输入材料 SHA-256、版本信息

评测目录必须放在仓库之外，并建议用短路径（竞品的 node_modules 层级较深，Windows 路径长度有限制）。
"""
import argparse
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import time
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "bench"))
from common import (COMPETE, REPO, TRACKS, EVAL_DATE, DEFAULT_TASKS_DIR, file_text, load_tasks,  # noqa: E402
                    natural_key, read_json, run_hash, sha256, task_dest, write_json)
from skill_list import marketplace_skills  # noqa: E402

FRONT_RE = re.compile(r"^\ufeff?---[ \t]*\r?\n(.*?)\r?\n---[ \t]*(?:\r?\n|$)", re.S)


# ---------------------------------------------------------------- 工具 ----

def rm_tree(path):
    """删除目录；Windows 上只读文件会让 rmtree 失败，先改权限重试。"""
    def onerror(func, p, exc_info):
        try:
            os.chmod(p, stat.S_IWRITE)
            func(p)
        except Exception:
            raise exc_info[1]
    if Path(path).exists():
        shutil.rmtree(path, onerror=onerror)


def to_msys(path):
    """D:\\a\\b -> /d/a/b，给 Git Bash 用。"""
    p = str(path).replace("\\", "/")
    m = re.match(r"^([A-Za-z]):/(.*)$", p)
    return f"/{m.group(1).lower()}/{m.group(2)}" if m else p


def git(repo, *args, check=True):
    p = subprocess.run(["git", "-C", str(repo), *args], capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if check and p.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} 失败：{(p.stderr or p.stdout).strip()[:300]}")
    return p.stdout.strip()


def parse_frontmatter(text):
    """解析 SKILL.md 开头的 YAML frontmatter（只支持本评测用到的写法）。

    支持：key: value、带引号的值、| > >- |- 块标量、缩进续行。返回 {key: str}；没有 frontmatter 返回 {}。
    """
    m = FRONT_RE.match(text)
    if not m:
        return {}
    lines = m.group(1).splitlines()
    out, key, mode, buf = {}, None, None, []

    def flush():
        if key is None:
            return
        if mode in ("|", "|-", "|+"):
            val = "\n".join(buf)
        else:  # 折叠或普通多行：用空格连接
            val = " ".join(s.strip() for s in buf if s.strip())
        out[key] = val.strip()

    for line in lines:
        km = re.match(r"^([A-Za-z_][\w-]*)\s*:\s*(.*)$", line)
        if km and not line.startswith((" ", "\t")):
            flush()
            key, rest = km.group(1), km.group(2).strip()
            if rest in ("|", ">", "|-", ">-", "|+", ">+"):
                mode, buf = rest, []
            else:
                mode = "plain"
                if len(rest) >= 2 and rest[0] == rest[-1] and rest[0] in "\"'":
                    rest = rest[1:-1]
                buf = [rest]
        elif key is not None:
            buf.append(line)
    flush()
    return out


def skill_meta(skill_dir):
    """读取一个 skill 目录的 (name, description)。没有 frontmatter 时 name 取目录名，description 为空。"""
    p = Path(skill_dir) / "SKILL.md"
    fm = parse_frontmatter(p.read_text(encoding="utf-8", errors="replace")) if p.is_file() else {}
    return fm.get("name") or Path(skill_dir).name, fm.get("description", "")


def split_batches(items, max_size):
    """把 n 个题尽量均匀地分成每批不超过 max_size 的若干批（11 题→4,4,3；8 题→4,4）。"""
    n = len(items)
    if n == 0:
        return []
    k = -(-n // max_size)
    base, extra = divmod(n, k)
    out, i = [], 0
    for j in range(k):
        size = base + (1 if j < extra else 0)
        out.append(items[i:i + size])
        i += size
    return out


def subst(cmd, ctx):
    for k, v in ctx.items():
        cmd = cmd.replace("{" + k + "}", str(v))
    return cmd


def run_cmds(cmds, ctx, base_cwd, label, timeout=600, extra_path=None):
    """依次执行 setup 或 smoke 命令（字符串或 {cmd, cwd, desc}），返回每条的结果记录。

    extra_path：放到 PATH 最前面的目录（PATH 替身所在目录），让 smoke 命令和运行时的环境一致。
    """
    env = os.environ.copy()
    if extra_path:
        env["PATH"] = str(extra_path) + os.pathsep + env.get("PATH", "")
    results = []
    for item in cmds:
        if isinstance(item, str):
            item = {"cmd": item, "cwd": "", "desc": ""}
        cmd = subst(item["cmd"], ctx)
        cwd = Path(base_cwd) / item.get("cwd", "") if item.get("cwd") else Path(base_cwd)
        t0 = time.time()
        try:
            p = subprocess.run(cmd, shell=True, cwd=str(cwd), capture_output=True, text=True,
                               encoding="utf-8", errors="replace", timeout=timeout, env=env)
            rc, tail = p.returncode, ((p.stdout or "") + (p.stderr or "")).strip()[-300:]
        except subprocess.TimeoutExpired:
            rc, tail = -1, f"超时（{timeout} 秒）"
        results.append({"cmd": cmd, "cwd": str(cwd), "desc": item.get("desc", ""), "rc": rc,
                        "ok": rc == 0, "seconds": round(time.time() - t0, 1), "tail": tail})
        print(f"    [{label}] {'通过' if rc == 0 else '失败(%s)' % rc}：{item.get('desc') or cmd[:60]}")
        if rc != 0 and tail:
            print("      " + tail.replace("\n", "\n      ")[:400])
    return results


def ensure_clone(c, comp_dir):
    """确认竞品克隆目录在指定提交上；不存在时克隆，提交不符时检出。返回克隆目录。"""
    d = Path(comp_dir) / c["repo"].replace("/", "_")
    if not d.is_dir():
        print(f"  克隆 {c['repo_url']} -> {d}")
        subprocess.run(["git", "clone", "--quiet", c["repo_url"], str(d)], check=True)
    head = git(d, "rev-parse", "HEAD")
    if head != c["commit"]:
        print(f"  {c['id']}：当前 {head[:7]}，检出 {c['commit'][:7]}")
        git(d, "fetch", "--quiet", "origin", check=False)
        git(d, "checkout", "--quiet", c["commit"])
        head = git(d, "rev-parse", "HEAD")
    if head != c["commit"]:
        raise RuntimeError(f"{c['id']} 无法检出提交 {c['commit']}")
    return d


# ------------------------------------------------------ 参赛方 skill 副本 ----

def copy_tree(src, dst, ignore_names):
    shutil.copytree(src, dst, ignore=shutil.ignore_patterns(*ignore_names), dirs_exist_ok=False)


def build_ours(c, src_repo, dst, args):
    """复制 22 个 skill 到 dst/<skill名>，返回 [(name, 描述)]。"""
    sync = src_repo / "tools" / "sync_shared.py"
    if sync.is_file():
        p = subprocess.run([sys.executable, str(sync), "--check"], capture_output=True, text=True,
                           encoding="utf-8", errors="replace", cwd=str(src_repo))
        if p.returncode != 0:
            msg = (p.stdout + p.stderr).strip()[-600:]
            if args.allow_unsynced:
                print(f"  警告：共享资料未同步（已按 --allow-unsynced 继续）：{msg}")
            else:
                raise RuntimeError(f"tools/sync_shared.py --check 未通过，请先运行 sync_shared.py 同步共享资料：\n{msg}")
        else:
            print("  共享资料已同步（sync_shared.py --check 通过）")
    names = marketplace_skills(src_repo)
    dst.mkdir(parents=True, exist_ok=True)
    listing = []
    for n in names:
        s = src_repo / "skills" / n
        if not (s / "SKILL.md").is_file():
            raise RuntimeError(f"插件市场清单里的 skill 不存在：{s}")
        copy_tree(s, dst / n, c["exclude"])
        listing.append(n)
    return listing


def build_contender(c, ctx, args, comp_dir):
    """建立 skills/<cid>/，执行 setup 和 smoke，返回写进 meta 的记录。"""
    cid = c["id"]
    cdir = ctx["skills_root"] / cid
    if cdir.exists():
        rm_tree(cdir)
    rec = {"id": cid, "dir": str(cdir), "setup": [], "smoke": []}
    if c["kind"] == "suite":
        src = Path(args.ours_src).resolve()
        rec["commit"] = git(src, "rev-parse", "HEAD")
        dirty = git(src, "status", "--porcelain", "--", "skills", "shared", "tools", ".claude-plugin", check=False)
        rec["dirty"] = [x for x in dirty.splitlines() if x.strip()]
        if rec["dirty"]:
            print(f"  警告：{src} 的 skills/shared/tools 有 {len(rec['dirty'])} 处未提交改动，"
                  f"副本内容与提交号 {rec['commit'][:7]} 不完全一致")
        rec["skills"] = build_ours(c, src, cdir, args)
        rec["skill_dir"] = str(cdir)
    else:
        src = ensure_clone(c, comp_dir)
        rec["commit"] = git(src, "rev-parse", "HEAD")
        copy_tree(src, cdir, c["exclude"])
        rec["skill_dir"] = str(cdir / c["skill_subdir"]) if c["skill_subdir"] else str(cdir)
        if not (Path(rec["skill_dir"]) / "SKILL.md").is_file():
            raise RuntimeError(f"{cid}：{rec['skill_dir']} 下没有 SKILL.md")
        name, _ = skill_meta(rec["skill_dir"])
        rec["skill_name_in_file"] = name
    shim_dir = cdir / "_shim"
    if c["env_path"]:
        shim_dir.mkdir(parents=True, exist_ok=True)
    cctx = dict(ctx["subst"], dir=cdir, shim_dir=shim_dir, smoke_out=ctx["smoke_dir"] / "out" / cid)
    (ctx["smoke_dir"] / "out" / cid).mkdir(parents=True, exist_ok=True)
    if not args.skip_setup:
        extra = shim_dir if c["env_path"] else None
        rec["setup"] = run_cmds(c["setup"], cctx, cdir, "setup", extra_path=extra)
        if not args.no_smoke:
            rec["smoke"] = run_cmds(c["smoke"], cctx, cdir, "smoke", extra_path=extra)
    rec["ready"] = all(x["ok"] for x in rec["setup"] + rec["smoke"])
    return rec


def make_smoke_inputs(smoke_dir):
    """烟雾测试用的样例输入：一份 md、一份 docx、一个只含 md 的语料目录。"""
    (smoke_dir / "corpus").mkdir(parents=True, exist_ok=True)
    md = ("# 关于做好冬季供暖保障工作的通知\n\n各区县人民政府：\n\n"
          "一、总体要求\n\n为确保群众温暖过冬，现就有关事项通知如下.\n\n"
          "（一）强化责任落实,\"压实\"各方责任。\n\n特此通知。\n")
    (smoke_dir / "sample.md").write_text(md, encoding="utf-8")
    (smoke_dir / "corpus" / "sample.md").write_text(md, encoding="utf-8")
    import docx
    d = docx.Document()
    d.add_paragraph("关于做好冬季供暖保障工作的通知")
    d.add_paragraph("各区县人民政府:")
    d.add_paragraph("一、总体要求")
    d.add_paragraph("为确保群众温暖过冬,现就有关事项通知如下.")
    d.save(str(smoke_dir / "sample.docx"))


# ---------------------------------------------------------- 任务说明 ----

def env_note_for(c, cdir):
    """需要 PATH 替身（zip、python3）的参赛方，在“已安装的 skill”一段末尾补一条运行环境说明。"""
    if not c["env_path"]:
        return ""
    shim = Path(cdir) / c["env_path"][0]
    cmds = "、".join(f"`{x}`" for x in c["env_commands"])
    return (f"4. 运行环境：该 skill 的文档或脚本里用到 {cmds} 命令，Windows 默认没有，已在 `{shim}` 目录提供同名命令。"
            f"运行 skill 的脚本之前，先把该目录加到 PATH 最前面——PowerShell：`$env:PATH = \"{shim};\" + $env:PATH`；"
            f"Git Bash：`export PATH=\"{to_msys(shim)}:$PATH\"`。shell 的环境变量不会在多次命令调用之间保留，"
            f"请在运行脚本的同一条命令里先设置 PATH。")


def build_instruction(c, rec, tpl, cases, n_total_skills=0):
    """拼出一份执行任务说明。cases：[{task, dir, files, user_message}]。"""
    n = len(cases)
    blocks = []
    for i, case in enumerate(cases, start=1):
        files_line = "、".join(case["files"]) or "无"
        blocks.append(f"---\n\n## 请求 {i}\n\n工作目录：{case['dir']}\n\n用户消息：\n\n<<<\n{case['user_message']}\n>>>\n\n"
                      f"用户提供的文件（都在工作目录中）：{files_line}\n")
    common = tpl["common"].replace("{n}", str(n)).replace("{cases}", "\n".join(blocks))
    kind = c["kind"]
    text = tpl[kind]
    text = text.replace("{n}", str(n))
    if kind == "suite":
        lines = []
        for name in rec["skills"]:
            sdir = Path(rec["skill_dir"]) / name
            _, desc = skill_meta(sdir)
            lines.append(f"- **{name}**（目录：{sdir}）：{desc}")
        text = text.replace("{k}", str(len(rec["skills"]))).replace("{skill_list}", "\n".join(lines))
    elif kind == "single":
        text = (text.replace("{skill_name}", c["skill_name"] or rec.get("skill_name_in_file", ""))
                .replace("{skill_dir}", rec["skill_dir"])
                .replace("{env_note}", env_note_for(c, rec["dir"])))
    text = text.replace("{common}", common)
    return re.sub(r"\n{3,}", "\n\n", text).rstrip() + "\n"


# --------------------------------------------------------------- 主流程 ----

def main():
    ap = argparse.ArgumentParser(description="准备竞品对标评测")
    ap.add_argument("--out", required=True, help="评测目录（必须在仓库之外，建议短路径）")
    ap.add_argument("--reps", type=int, default=2, help="每个（题 × 参赛方）的重复次数，默认 2")
    ap.add_argument("--tracks", default="W,F,H", help="参加评测的赛道，逗号分隔，默认 W,F,H")
    ap.add_argument("--tasks", default="", help="只评这些题号，逗号分隔，默认题目目录下的全部题")
    ap.add_argument("--tasks-dir", default=str(DEFAULT_TASKS_DIR), help="题目目录，默认 evals/compete/tasks")
    ap.add_argument("--ours-src", default=str(REPO), help="本项目 skill 的来源仓库，默认当前仓库")
    ap.add_argument("--competitors-dir", default=str(Path.home() / ".gongwen-bench" / "competitors"),
                    help="竞品克隆目录（<owner>_<repo>）；不存在的会按 contenders.json 克隆")
    ap.add_argument("--contenders", default="", help="只含这些参赛方，逗号分隔，默认全部")
    ap.add_argument("--tag", default="compete-v1", help="评测标签，参与匿名目录哈希，默认 compete-v1")
    ap.add_argument("--chunk", type=int, default=4, help="每个子智能体最多处理的题数，默认 4（每批 3～4 题）")
    ap.add_argument("--skip-setup", action="store_true", help="跳过 setup 与 smoke 命令（只复制文件）")
    ap.add_argument("--no-smoke", action="store_true", help="执行 setup，但不跑 smoke 烟雾测试")
    ap.add_argument("--allow-unsynced", action="store_true", help="共享资料未同步时只警告、不中止")
    ap.add_argument("--force", action="store_true", help="评测目录里已有运行产出时仍然重建（会删除这些产出）")
    args = ap.parse_args()

    out = Path(args.out).resolve()
    ours_src = Path(args.ours_src).resolve()
    for root in {REPO, ours_src}:
        if out == root or root in out.parents:
            print(f"评测目录不要放在仓库内（{root}），以免参赛方读到仓库文件", file=sys.stderr)
            return 2
    tracks = [t for t in args.tracks.split(",") if t]
    if not set(tracks) <= set(TRACKS):
        print(f"--tracks 只能取 {','.join(TRACKS)}", file=sys.stderr)
        return 2
    ids = [t for t in args.tasks.split(",") if t]
    tasks = load_tasks(args.tasks_dir, tracks=tracks, ids=ids or None)
    if not tasks:
        print(f"题目目录 {args.tasks_dir} 下没有符合条件的题（需要 <题号>/task.json）", file=sys.stderr)
        return 2
    if ids:
        missing = sorted(set(ids) - {t["id"] for t in tasks})
        if missing:
            print(f"题目目录里找不到这些题：{'、'.join(missing)}", file=sys.stderr)
            return 2
    # 题目的输入材料必须齐全：先全部检查，再做耗时的复制与 setup，免得做到一半才发现缺文件
    missing_files = [f"{t['id']}：{rel}" for t in tasks for rel in t.get("files", []) if not (Path(t["_dir"]) / rel).is_file()]
    if missing_files:
        print("这些题的输入材料不存在（题目目录里缺文件）：\n  " + "\n  ".join(missing_files), file=sys.stderr)
        return 2
    bad_tasks = [t["id"] for t in tasks if not t.get("user_message") or not t.get("checkpoints")]
    if bad_tasks:
        print(f"这些题的 task.json 缺少 user_message 或 checkpoints：{'、'.join(bad_tasks)}", file=sys.stderr)
        return 2
    contenders = read_json(COMPETE / "contenders.json")
    if args.contenders:
        want = set(args.contenders.split(","))
        contenders = [c for c in contenders if c["id"] in want]
    contenders = [c for c in contenders if set(c["tracks"]) & set(tracks)]
    if not contenders:
        print("没有参加所选赛道的参赛方", file=sys.stderr)
        return 2

    # 已有运行产出时不允许悄悄重建
    if (out / "runs").is_dir() and any((out / "runs").glob("*/_回复.md")) and not args.force:
        print("评测目录里已有运行产出（runs/*/_回复.md）；重建会删除它们。确需重建请加 --force", file=sys.stderr)
        return 2
    out.mkdir(parents=True, exist_ok=True)
    for sub in ("skills", "runs", "inputs", "tasks", "taskdefs", "smoke"):
        (out / sub).mkdir(exist_ok=True)
    if args.force:
        rm_tree(out / "runs")
        (out / "runs").mkdir()

    # 评测日期：题目目录下 _meta.json 的 date（每套题按设计时的日期），没有时用默认值
    eval_date = EVAL_DATE
    meta_file = Path(args.tasks_dir) / "_meta.json"
    if meta_file.is_file():
        eval_date = json.loads(meta_file.read_text(encoding="utf-8")).get("date", EVAL_DATE)
    tpl = {"common": (COMPETE / "templates" / "run-common.md").read_text(encoding="utf-8").replace("{eval_date}", eval_date),
           "suite": (COMPETE / "templates" / "run-ours.md").read_text(encoding="utf-8"),
           "single": (COMPETE / "templates" / "run-single.md").read_text(encoding="utf-8"),
           "none": (COMPETE / "templates" / "run-none.md").read_text(encoding="utf-8")}

    # 一、参赛方的 skill 副本
    print(f"[1/3] 建立 skill 副本（{len(contenders)} 个参赛方）")
    smoke_dir = out / "smoke"
    make_smoke_inputs(smoke_dir)
    ctx = {"skills_root": out / "skills", "smoke_dir": smoke_dir,
           "subst": {"compete": COMPETE, "python": sys.executable, "smoke_md": smoke_dir / "sample.md",
                     "smoke_docx": smoke_dir / "sample.docx", "smoke_dir": smoke_dir / "corpus"}}
    recs = {}
    for c in contenders:
        if c["kind"] == "none":
            recs[c["id"]] = {"id": "none", "commit": "", "ready": True}
            continue
        print(f"  {c['id']}")
        recs[c["id"]] = build_contender(c, ctx, args, args.competitors_dir)
    not_ready = [cid for cid, r in recs.items() if not r.get("ready", True)]

    # 二、题目快照、输入文本、匿名工作目录
    print(f"[2/3] 建立匿名工作目录（{len(tasks)} 题 × 参赛方 × {args.reps} 次）")
    index, hashes = [], {}
    cases_by_group = {}
    for t in tasks:
        tdir = Path(t["_dir"])
        snap = out / "taskdefs" / t["id"]
        if snap.exists():
            rm_tree(snap)
        shutil.copytree(tdir, snap)
        parts = [f"【用户消息】\n{t['user_message']}"]
        for rel in t.get("files", []):
            src = tdir / rel
            if not src.is_file():
                print(f"题 {t['id']} 的输入材料不存在：{src}", file=sys.stderr)
                return 2
            parts.append(f"【材料：{Path(rel).name}】\n{file_text(src)}")
        (out / "inputs" / f"{t['id']}.txt").write_text("\n\n".join(parts) + "\n", encoding="utf-8")
        for c in contenders:
            if t["track"] not in c["tracks"]:
                continue
            for rep in range(1, args.reps + 1):
                h = run_hash(t["id"], c["id"], rep, args.tag)
                wdir = out / "runs" / h
                if wdir.exists():
                    rm_tree(wdir)
                wdir.mkdir(parents=True)
                dests = []
                for rel in t.get("files", []):
                    dest_rel = task_dest(rel)
                    dest = wdir / dest_rel
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(tdir / rel, dest)
                    hashes[f"runs/{h}/{dest_rel}"] = sha256(dest)  # 键用相对路径，评测目录移动后仍然有效
                    dests.append(dest_rel)
                row = {"hash": h, "task": t["id"], "track": t["track"], "contender": c["id"], "rep": rep,
                       "dir": str(wdir), "files": dests}
                index.append(row)
                cases_by_group.setdefault((c["id"], t["track"], rep), []).append(
                    {"task": t["id"], "hash": h, "dir": str(wdir), "files": dests, "user_message": t["user_message"]})

    # 三、分批与执行任务说明
    print("[3/3] 生成执行任务说明")
    batches = []
    by_cid = {c["id"]: c for c in contenders}
    for (cid, track, rep), cases in sorted(cases_by_group.items(),
                                           key=lambda kv: (kv[0][0], kv[0][1], kv[0][2])):
        cases.sort(key=lambda x: natural_key(x["task"]))
        for bi, chunk in enumerate(split_batches(cases, args.chunk), start=1):
            bid = f"{cid}-{track}-r{rep}-{bi}"
            text = build_instruction(by_cid[cid], recs[cid], tpl, chunk)
            p = out / "tasks" / f"{bid}.md"
            p.write_text(text, encoding="utf-8")
            batches.append({"batch_id": bid, "instruction_path": str(p), "contender": cid, "track": track,
                            "rep": rep, "cases": [{"task": x["task"], "hash": x["hash"], "dir": x["dir"]} for x in chunk]})
            for x in chunk:
                for r in index:
                    if r["hash"] == x["hash"]:
                        r["batch"] = bid
    index.sort(key=lambda r: (natural_key(r["task"]), r["contender"], r["rep"]))
    write_json(out / "index.json", index)
    write_json(out / "hashes.json", hashes)
    write_json(out / "run_batches.json", batches)
    ours = recs.get("ours", {})
    write_json(out / "meta.json", {
        "tag": args.tag, "date": eval_date, "reps": args.reps, "tracks": tracks,
        "tasks": [t["id"] for t in tasks], "tasks_dir": str(Path(args.tasks_dir).resolve()),
        "ours_src": str(ours_src), "ours_commit": ours.get("commit", ""), "ours_dirty": ours.get("dirty", []),
        "contenders": {cid: {"commit": r.get("commit", ""), "dir": r.get("dir", ""), "ready": r.get("ready", True)}
                       for cid, r in recs.items()},
        "setup_log": {cid: {"setup": r.get("setup", []), "smoke": r.get("smoke", [])} for cid, r in recs.items()},
        "created": time.strftime("%Y-%m-%d %H:%M:%S"),
    })
    print(f"题目 {len(tasks)} 道，运行 {len(index)} 次（匿名目录 runs/），批次 {len(batches)} 个"
          f"（run_batches.json；每批一个子智能体，并发不超过 10 个）")
    if ours.get("commit"):
        print(f"ours 提交号：{ours['commit']}" + ("（有未提交改动）" if ours.get("dirty") else ""))
    if not_ready:
        print(f"注意：这些参赛方的 setup 或烟雾测试有失败，运行时可能缺依赖：{'、'.join(not_ready)}（详情见 meta.json 的 setup_log）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
