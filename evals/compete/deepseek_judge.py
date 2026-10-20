"""DeepSeek 评审：对 W、H 赛道每组输出，把与 Claude 评审相同的材料拼成一个提示词，正序、倒序各评一次。

用法：
    python evals/compete/deepseek_judge.py --bench <评测目录> [--model deepseek-chat] [--only W01,H03,...]
                                           [--workers 4] [--rerun] [--dry-run] [--dump-prompts <目录>]

- F 赛道不评（DeepSeek 看不了图片）。
- 接口：OpenAI 兼容的 https://api.deepseek.com/chat/completions，只用标准库 urllib；温度 0，要求 JSON 输出；
  解析或校验失败重试 3 次；并发 4；已有结果默认跳过（--rerun 重评）。
- Key 只从仓库根目录 deepseek.key.local 读取（取第一个 sk- 开头的片段），不打印、不写入任何文件和日志；文件不存在时提示后退出，退出码 2。
- 倒序时把输出倒过来重新分配字母（原来最后一份变成 A），映射记录在结果里。
- 提示词估算超过 100k 字时，先截断最长的输入材料，再截断最长的输出，并在提示中注明。
结果：<评测目录>/grades/deepseek/<题号>-r<重复>-<fwd|rev>.json
    {task, rep, direction, model, mapping: {字母: {orig_letter, contender, hash}}, grade: {……与 Claude 评审同结构……}, usage, attempts, truncated}
"""
import argparse
import json
import re
import sys
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from common import (Bench, COMPETE, DEEPSEEK_KEY_FILE, load_deepseek_key, dim_defs_text, dims_json_example, natural_key,  # noqa: E402
                    normalize_grade, read_json, validate_grade, write_json)

API_URL = "https://api.deepseek.com/chat/completions"
KEY_PATH = DEEPSEEK_KEY_FILE
MAX_CHARS = 100_000
LETTERS = "ABCDEFGHIJKLMNOP"
RULES_RE = re.compile(r"<!--RULES-->(.*?)<!--/RULES-->", re.S)
MATERIAL_RE = re.compile(r"(?m)^(【材料：[^】]*】)\n")

SYSTEM = ("你是中文公文写作评测的评审员。用户会给出一个写作请求、检查点，以及多份匿名输出；"
          "请严格按下面的评审要求评审，并且只输出一个合法的 JSON 对象（JSON 之外不要有任何文字）。")


def read_key():
    """读取 API Key；不存在时返回 None。Key 只在内存里使用，绝不打印。"""
    return load_deepseek_key()


def rules_text(track, task_id, rep, letters, n_checkpoints):
    tpl = (COMPETE / "templates" / f"judge-{track}.md").read_text(encoding="utf-8")
    m = RULES_RE.search(tpl)
    text = m.group(1)
    for k, v in {"task_id": task_id, "rep": rep, "letters": "、".join(letters), "n_checkpoints": n_checkpoints,
                 "dim_defs": dim_defs_text(track), "dims_json": dims_json_example(track)}.items():
        text = text.replace("{" + k + "}", str(v))
    return text.strip()


def split_materials(inputs_text):
    """把 inputs/<题号>.txt 拆成 [(名称, 文本)]：第一段是用户消息，其后每份材料一段。"""
    parts = MATERIAL_RE.split(inputs_text)
    blocks = [("用户消息", parts[0].replace("【用户消息】", "", 1).strip())]
    for i in range(1, len(parts) - 1, 2):
        blocks.append((parts[i], parts[i + 1].strip()))
    return blocks


def truncate(text, keep):
    n = len(text)
    return text[:keep] + f"\n【……此处因超长截断，原文共 {n} 字，保留前 {keep} 字……】"


def build_prompt(bench, task, rep, order, track):
    """order：[(orig_letter, contender, hash)]，按呈现顺序。返回 (system, user, 呈现字母列表, 截断说明)。"""
    letters = list(LETTERS[:len(order)])
    n_cp = len(task.get("checkpoints", []))
    materials = split_materials(bench.inputs_text(task["id"]))
    outs = []
    for L, (orig, cid, h) in zip(letters, order):
        s = Path(bench.root) / "sanitized" / h
        reply = (s / "_回复.md").read_text(encoding="utf-8") if (s / "_回复.md").is_file() else "（没有 _回复.md）"
        bodies = sorted((s / "_正文").glob("*")) if (s / "_正文").is_dir() else []
        body = "\n\n".join(f"〈{b.name}〉\n" + b.read_text(encoding="utf-8") for b in bodies) if bodies \
            else "（没有交付文书全文，以回复为准）"
        summ = (s / "机检摘要.md").read_text(encoding="utf-8") if (s / "机检摘要.md").is_file() else "（缺少机检摘要）"
        outs.append({"letter": L, "reply": reply, "body": body, "summary": summ})

    # 超长时先截断最长的输入材料，再截断最长的输出部分
    truncated = []

    def total():
        return (sum(len(t) for _, t in materials)
                + sum(len(o["reply"]) + len(o["body"]) + len(o["summary"]) for o in outs) + 6000)

    guard = 0
    while total() > MAX_CHARS and guard < 200:
        guard += 1
        excess = total() - MAX_CHARS
        cands = [("m", i, len(t)) for i, (_, t) in enumerate(materials) if len(t) > 3000]
        if cands:
            _, i, ln = max(cands, key=lambda x: x[2])
            name, t = materials[i]
            keep = max(2000, ln - excess - 200)
            materials[i] = (name, truncate(t, keep))
            truncated.append(f"输入材料 {name}：{ln} 字 → 保留前 {keep} 字")
            continue
        cands = [(o, k, len(o[k])) for o in outs for k in ("body", "reply") if len(o[k]) > 3000]
        if not cands:
            break
        o, k, ln = max(cands, key=lambda x: x[2])
        keep = max(2000, ln - excess - 200)
        o[k] = truncate(o[k], keep)
        truncated.append(f"输出 {o['letter']} 的{'文书全文' if k == 'body' else '回复'}：{ln} 字 → 保留前 {keep} 字")

    cps = "\n".join(f"{i}. {c['text']}" + (f"（{c['kind']}）" if c.get("kind") else "")
                    for i, c in enumerate(task.get("checkpoints", [])))
    mats = "\n\n".join(f"{name}\n{text}" if name.startswith("【") else f"【{name}】\n{text}" for name, text in materials)
    blocks = []
    for o in outs:
        blocks.append(f"### 输出 {o['letter']}\n\n【回复】\n{o['reply'].strip()}\n\n【文书全文】\n{o['body'].strip()}\n\n"
                      f"【机检摘要】\n{o['summary'].strip()}")
    note = ("\n\n注意：下列材料因超长已截断：" + "；".join(truncated) + "。") if truncated else ""
    user = (f"# 评审任务（{task['id']}，第 {rep} 轮）\n\n题目：{task.get('title', '')}（"
            f"{'文种' if track == 'W' else '体裁'}：{task.get('genre', '')}）\n\n"
            f"## 输入材料（含用户消息；docx、xlsx 已转为文本）\n\n{mats}{note}\n\n"
            f"## 检查点（共 {len(task.get('checkpoints', []))} 条，idx 从 0 起）\n\n{cps}\n\n"
            f"## 待评输出（{len(outs)} 份：{'、'.join(letters)}）\n\n" + "\n\n".join(blocks) + "\n\n"
            f"请按评审要求输出 JSON 对象，outputs 里要有全部字母（{'、'.join(letters)}）。")
    system = SYSTEM + "\n\n" + rules_text(track, task["id"], rep, letters, n_cp)
    return system, user, letters, truncated


def call_api(key, model, system, user, timeout=420, url=API_URL):
    body = json.dumps({"model": model, "temperature": 0, "max_tokens": 8192,
                       "response_format": {"type": "json_object"},
                       "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}]},
                      ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="POST",
                                 headers={"Content-Type": "application/json", "Authorization": "Bearer " + key})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def judge_once(key, model, system, user, letters, n_cp, track, attempts=3, url=API_URL):
    """调用并校验，失败重试；返回 (grade, usage, 次数, 错误说明)。"""
    last = ""
    for i in range(1, attempts + 1):
        try:
            data = call_api(key, model, system, user, url=url)
            content = data["choices"][0]["message"]["content"]
            usage = data.get("usage", {})
            g = normalize_grade(json.loads(content))
            problems = validate_grade(g, letters, n_cp, track)
            if not problems:
                return g, usage, i, ""
            last = "校验未通过：" + "；".join(problems[:4])
        except urllib.error.HTTPError as e:
            last = f"HTTP {e.code}"
            if e.code in (401, 403):  # 认证问题重试没有意义
                return None, {}, i, last
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            last = f"网络错误：{type(e).__name__}"
        except (json.JSONDecodeError, KeyError, IndexError) as e:
            last = f"解析失败：{type(e).__name__}"
        time.sleep(min(2 ** i, 20))
    return None, {}, attempts, last


def main():
    ap = argparse.ArgumentParser(description="DeepSeek 评审（W、H 赛道，正序倒序各一次）")
    ap.add_argument("--bench", required=True, help="评测目录")
    ap.add_argument("--model", default="deepseek-chat")
    ap.add_argument("--only", default="", help="只评这些题号（逗号分隔）")
    ap.add_argument("--workers", type=int, default=4, help="并发数，默认 4")
    ap.add_argument("--rerun", action="store_true", help="已有结果也重评")
    ap.add_argument("--dry-run", action="store_true", help="只拼提示词、估算长度，不调用接口")
    ap.add_argument("--api-url", default=API_URL, help="接口地址，默认 DeepSeek 官方；仅自测时改成本地模拟服务")
    ap.add_argument("--dump-prompts", default="", help="把提示词写到这个目录，便于检查（--dry-run 时默认写到 grades/deepseek/_prompts）")
    args = ap.parse_args()
    bench = Bench(args.bench)
    blind = read_json(bench.root / "judge" / "blind.json")
    if not blind:
        print("缺少 judge/blind.json，请先运行 judge_prep.py", file=sys.stderr)
        return 2
    key = None
    if not args.dry_run:
        key = read_key()
        if key is None:
            print(f"找不到 DeepSeek Key：请把 Key 存放在 {KEY_PATH}（文件中含 sk- 开头的 Key 即可），再重新运行。", file=sys.stderr)
            return 2
    only = {x for x in args.only.split(",") if x}
    out_dir = bench.root / "grades" / "deepseek"
    out_dir.mkdir(parents=True, exist_ok=True)
    dump = Path(args.dump_prompts) if args.dump_prompts else (out_dir / "_prompts" if args.dry_run else None)
    if dump:
        dump.mkdir(parents=True, exist_ok=True)

    jobs = []
    for gkey, mapping in sorted(blind.items(), key=lambda kv: natural_key(kv[0])):
        tid, rep = re.match(r"^(.+)-r(\d+)$", gkey).group(1), int(re.match(r"^(.+)-r(\d+)$", gkey).group(2))
        task = bench.task(tid)
        if task["track"] not in ("W", "H"):
            continue
        if only and tid not in only:
            continue
        fwd = [(L, v["contender"], v["hash"]) for L, v in sorted(mapping.items())]
        for direction, order in (("fwd", fwd), ("rev", list(reversed(fwd)))):
            path = out_dir / f"{gkey}-{direction}.json"
            if path.exists() and not args.rerun:
                continue
            jobs.append((gkey, tid, rep, task, direction, order, path))
    if not jobs:
        print("没有需要评审的组（W、H 赛道已全部有结果，或没有符合条件的组）")
        return 0
    print(f"待评 {len(jobs)} 次调用（模型 {args.model}{'，仅预演' if args.dry_run else ''}）")

    def work(job):
        gkey, tid, rep, task, direction, order, path = job
        system, user, letters, truncated = build_prompt(bench, task, rep, order, task["track"])
        size = len(system) + len(user)
        if dump:
            (dump / f"{gkey}-{direction}.md").write_text("【system】\n" + system + "\n\n【user】\n" + user, encoding="utf-8")
        if args.dry_run:
            return gkey, direction, f"提示词 {size} 字（system {len(system)}，user {len(user)}）" + \
                   (f"，截断：{'；'.join(truncated)}" if truncated else ""), True
        g, usage, n, err = judge_once(key, args.model, system, user, letters, len(task.get("checkpoints", [])), task["track"],
                                      url=args.api_url)
        if g is None:
            return gkey, direction, f"失败（{n} 次）：{err}", False
        mapping = {L: {"orig_letter": orig, "contender": cid, "hash": h} for L, (orig, cid, h) in zip(letters, order)}
        write_json(path, {"task": tid, "rep": rep, "direction": direction, "model": args.model, "mapping": mapping,
                          "grade": g, "usage": usage, "attempts": n, "truncated": truncated})
        return gkey, direction, f"完成（{n} 次调用，输入 {size} 字）", True

    bad = 0
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as ex:
        for gkey, direction, msg, ok in ex.map(work, jobs):
            print(f"{gkey}-{direction}: {msg}")
            bad += not ok
    if bad:
        print(f"有 {bad} 次调用失败；重新运行同一命令会只补评失败的部分")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
