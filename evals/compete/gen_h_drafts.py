#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""为 H 赛道（去 AI 味）生成初稿。

做什么
    对每个 H 题读取 tasks/Hxx/reference/生成提示词.md，调用 DeepSeek（OpenAI 兼容接口）生成正文，
    写入 tasks/Hxx/files/初稿.md，并把模型名、温度、生成时间等写入 tasks/Hxx/reference/生成记录.json。
    提示词是普通的“请根据以下事实写一篇……”，不要求模型刻意写差，保留模型自然产生的 AI 味。

用法
    python gen_h_drafts.py --dry-run                 # 只打印将要调用的题号和提示词长度（不读 key、不联网、不写文件）
    python gen_h_drafts.py                           # 为尚无初稿的 H 题生成初稿
    python gen_h_drafts.py --only H01,H03            # 只处理指定题号
    python gen_h_drafts.py --force                   # 已有初稿也重新生成并覆盖
    python gen_h_drafts.py --model deepseek-chat --temperature 0.7

Key
    只从仓库根目录 deepseek.key.local 读取（取第一个 sk- 开头的片段）。绝不打印，也不写入任何文件或日志。
    文件不存在时给出中文提示并以退出码 2 结束。

退出码
    0  成功（含全部跳过、--dry-run）
    1  有题目生成失败
    2  key 缺失或不可用、参数有误
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import socket
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
TASKS_DIR = ROOT / "tasks"
KEY_FILE = Path(__file__).resolve().parents[2] / "deepseek.key.local"

API_BASE = "https://api.deepseek.com"
API_URL = API_BASE + "/chat/completions"

DEFAULT_MODEL = "deepseek-chat"
DEFAULT_TEMPERATURE = 0.7
DEFAULT_MAX_TOKENS = 4096  # 1200～1500 字的正文约 2000 token，4096 足够；deepseek-chat 上限为 8192
DEFAULT_TIMEOUT = 180  # 单次请求超时（秒）
DEFAULT_RETRIES = 3  # 失败后最多重试 3 次（每题最多请求 4 次）
BACKOFF_SECONDS = 2.0  # 重试间隔：2、4、8 秒

PROMPT_REL = Path("reference") / "生成提示词.md"
DRAFT_REL = Path("files") / "初稿.md"
RECORD_REL = Path("reference") / "生成记录.json"

# 这些状态码重试可能恢复（限流、服务端繁忙等）
RETRYABLE_HTTP = {408, 409, 425, 429, 500, 502, 503, 504}
# 认证失败、余额不足：重试无意义，其余题目也会同样失败，直接终止
FATAL_HTTP = {401, 402}

_SECRETS: list[str] = []  # 已读到的 key，输出前统一遮蔽，防止任何路径把 key 带进日志


def _setup_stdio() -> None:
    """让中文在管道/重定向时也以 UTF-8 输出，且遇到无法编码的字符时不抛异常。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def scrub(text: str) -> str:
    for secret in _SECRETS:
        if secret:
            text = text.replace(secret, "***")
    return text


def log(msg: str = "", *, err: bool = False) -> None:
    print(scrub(str(msg)), file=sys.stderr if err else sys.stdout, flush=True)


class KeyFileError(Exception):
    """key 文件缺失或内容不可用。"""


class ApiError(Exception):
    """调用 API 失败。retryable：重试可能恢复；fatal：应终止后续所有题目。"""

    def __init__(self, message: str, retryable: bool = False, fatal: bool = False):
        super().__init__(message)
        self.retryable = retryable
        self.fatal = fatal


# ---------------------------------------------------------------------------
# key
# ---------------------------------------------------------------------------
def read_key() -> str:
    if not KEY_FILE.is_file():
        raise KeyFileError(
            f"未找到 DeepSeek key 文件：{KEY_FILE}\n"
            "请先创建该文件，写入 sk- 开头的 API Key，再重新运行。\n"
            "也可以先加 --dry-run 检查将要调用的题目。"
        )
    try:
        text = KEY_FILE.read_text(encoding="utf-8-sig", errors="replace")
    except OSError as exc:
        raise KeyFileError(f"无法读取 key 文件：{KEY_FILE}（{exc.__class__.__name__}）") from None
    # 文件中可以有标签、备注等其他行，只取第一个 sk- 开头的片段
    m = re.search(r"sk-[A-Za-z0-9]{16,}", text)
    if not m:
        raise KeyFileError(f"key 文件中没有找到 sk- 开头的 Key：{KEY_FILE}")
    key = m.group(0)
    _SECRETS.append(key)
    return key


# ---------------------------------------------------------------------------
# 任务发现与计划
# ---------------------------------------------------------------------------
def discover_h_tasks() -> list[str]:
    if not TASKS_DIR.is_dir():
        return []
    return sorted(p.name for p in TASKS_DIR.iterdir() if p.is_dir() and re.fullmatch(r"H\d{2}", p.name))


class ArgError(Exception):
    """参数类错误，main 里转成退出码 2。"""


def select_tasks(only: str | None, all_ids: list[str]) -> list[str]:
    if not only:
        return all_ids
    wanted: list[str] = []
    for part in only.split(","):
        tid = part.strip().upper()
        if tid and tid not in wanted:
            wanted.append(tid)
    unknown = [t for t in wanted if t not in all_ids]
    if unknown:
        raise ArgError(f"未知题号：{', '.join(unknown)}。可用题号：{', '.join(all_ids)}")
    return wanted


def make_plan(ids: list[str], force: bool) -> list[dict]:
    plan = []
    for tid in ids:
        tdir = TASKS_DIR / tid
        item = {
            "id": tid,
            "dir": tdir,
            "prompt": "",
            "draft_path": tdir / DRAFT_REL,
            "has_draft": False,
            "action": "error",
            "reason": "",
        }
        prompt_path = tdir / PROMPT_REL
        draft_path = item["draft_path"]
        item["has_draft"] = draft_path.is_file() and draft_path.stat().st_size > 0
        if not prompt_path.is_file():
            item["reason"] = f"缺少 {PROMPT_REL.as_posix()}"
        else:
            prompt = prompt_path.read_text(encoding="utf-8-sig").strip()
            if not prompt:
                item["reason"] = f"{PROMPT_REL.as_posix()} 是空的"
            else:
                item["prompt"] = prompt
                if item["has_draft"] and not force:
                    item["action"] = "skip"
                    item["reason"] = "已有初稿，跳过（加 --force 覆盖）"
                else:
                    item["action"] = "generate"
                    item["reason"] = "将调用（--force，覆盖已有初稿）" if item["has_draft"] else "将调用"
        plan.append(item)
    return plan


# ---------------------------------------------------------------------------
# 调用 API
# ---------------------------------------------------------------------------
def _error_detail(exc: urllib.error.HTTPError) -> str:
    """取出服务端返回的错误说明（截断），用于提示；不含任何请求头。"""
    try:
        raw = exc.read(2000).decode("utf-8", errors="replace")
    except Exception:  # noqa: BLE001 - 读错误体失败不影响主流程
        return ""
    try:
        msg = json.loads(raw).get("error", {}).get("message", "")
    except (ValueError, AttributeError):
        msg = ""
    return (msg or raw).strip().replace("\n", " ")[:200]


def call_api(key: str, prompt: str, args: argparse.Namespace) -> dict:
    """发一次请求，返回 {"content", "finish_reason", "usage", "model"}；失败抛 ApiError。"""
    payload = {
        "model": args.model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": args.temperature,
        "max_tokens": args.max_tokens,
        "stream": False,
    }
    req = urllib.request.Request(
        API_URL,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        method="POST",
        headers={
            "Content-Type": "application/json",
            "Authorization": "Bearer " + key,
            "User-Agent": "gongwen-compete-gen/1.0",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=args.timeout) as resp:
            raw = resp.read()
    except urllib.error.HTTPError as exc:  # 必须先于 URLError（它是 URLError 的子类）
        code = exc.code
        hint = {401: "认证失败，请检查 key 是否有效", 402: "账户余额不足",
                400: "请求格式错误", 422: "请求参数无效", 429: "请求过于频繁（限流）"}.get(code, "")
        detail = _error_detail(exc)
        msg = f"HTTP {code}" + (f"（{hint}）" if hint else "") + (f"：{detail}" if detail else "")
        raise ApiError(msg, retryable=code in RETRYABLE_HTTP, fatal=code in FATAL_HTTP) from None
    except (urllib.error.URLError, socket.timeout, TimeoutError, ConnectionError, OSError) as exc:
        reason = getattr(exc, "reason", exc)
        raise ApiError(f"网络错误：{reason}", retryable=True) from None

    try:
        obj = json.loads(raw.decode("utf-8"))
        choice = obj["choices"][0]
        content = choice["message"]["content"]
        finish = choice.get("finish_reason")
    except (ValueError, KeyError, IndexError, TypeError):
        raise ApiError("返回内容无法解析", retryable=True) from None
    if not isinstance(content, str) or not content.strip():
        raise ApiError("返回的正文为空", retryable=True)
    if finish not in (None, "stop"):
        raise ApiError(f"生成未正常结束（finish_reason={finish}）", retryable=True)
    return {"content": content, "finish_reason": finish, "usage": obj.get("usage"), "model": obj.get("model")}


def generate_with_retry(key: str, prompt: str, args: argparse.Namespace) -> tuple[dict, int]:
    """带重试的生成。返回 (结果, 实际请求次数)。"""
    attempt = 0
    while True:
        attempt += 1
        try:
            return call_api(key, prompt, args), attempt
        except ApiError as exc:
            if exc.fatal or not exc.retryable or attempt > args.retries:
                raise
            wait = BACKOFF_SECONDS * (2 ** (attempt - 1))
            log(f"  第 {attempt} 次请求失败：{exc}；{wait:g} 秒后重试（{attempt}/{args.retries}）")
            time.sleep(wait)


# ---------------------------------------------------------------------------
# 写出结果
# ---------------------------------------------------------------------------
def _write_utf8(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    with open(tmp, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(text)
    tmp.replace(path)  # 先写临时文件再替换，失败时不会留下半截初稿


def write_outputs(item: dict, result: dict, attempts: int, args: argparse.Namespace) -> int:
    draft = result["content"].strip() + "\n"
    _write_utf8(item["draft_path"], draft)
    record = {
        "task": item["id"],
        "model": args.model,
        "model_returned": result.get("model"),
        "temperature": args.temperature,
        "max_tokens": args.max_tokens,
        "api_base": API_BASE,
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "prompt_file": PROMPT_REL.as_posix(),
        "prompt_chars": len(item["prompt"]),
        "prompt_sha256": hashlib.sha256(item["prompt"].encode("utf-8")).hexdigest(),
        "output_file": DRAFT_REL.as_posix(),
        "output_chars": len(re.sub(r"\s+", "", draft)),
        "finish_reason": result.get("finish_reason"),
        "usage": result.get("usage"),
        "attempts": attempts,
    }
    _write_utf8(item["dir"] / RECORD_REL, json.dumps(record, ensure_ascii=False, indent=2) + "\n")
    return record["output_chars"]


# ---------------------------------------------------------------------------
# 命令行
# ---------------------------------------------------------------------------
def parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="gen_h_drafts.py",
        description="为 H 赛道（去 AI 味）调用 DeepSeek 生成初稿，写入 tasks/Hxx/files/初稿.md。",
        epilog="Key 只从仓库根目录 deepseek.key.local 读取；不会被打印或写入任何文件。",
    )
    parser.add_argument("--only", metavar="H01,H03", help="只处理指定题号，逗号分隔")
    parser.add_argument("--dry-run", action="store_true",
                        help="只打印将要调用的题号和提示词长度；不读 key、不联网、不写文件")
    parser.add_argument("--force", action="store_true", help="已有初稿也重新生成并覆盖（默认跳过）")
    parser.add_argument("--model", default=DEFAULT_MODEL, help=f"模型名（默认 {DEFAULT_MODEL}）")
    parser.add_argument("--temperature", type=float, default=DEFAULT_TEMPERATURE,
                        help=f"采样温度，0～2（默认 {DEFAULT_TEMPERATURE}）")
    parser.add_argument("--max-tokens", type=int, default=DEFAULT_MAX_TOKENS,
                        help=f"单次生成的 token 上限（默认 {DEFAULT_MAX_TOKENS}）")
    parser.add_argument("--retries", type=int, default=DEFAULT_RETRIES,
                        help=f"失败后最多重试次数（默认 {DEFAULT_RETRIES}，即每题最多请求 {DEFAULT_RETRIES + 1} 次）")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help=f"单次请求超时秒数（默认 {DEFAULT_TIMEOUT}）")
    args = parser.parse_args(argv)
    if not 0 <= args.temperature <= 2:
        parser.error("--temperature 应在 0～2 之间")
    if args.retries < 0:
        parser.error("--retries 不能为负数")
    if args.max_tokens < 1 or args.timeout < 1:
        parser.error("--max-tokens 和 --timeout 必须为正整数")
    return args


def print_plan(plan: list[dict], args: argparse.Namespace) -> None:
    log(f"[dry-run] 模型 {args.model}，温度 {args.temperature}，接口 {API_URL}")
    log(f"[dry-run] key 文件：{KEY_FILE}（{'存在' if KEY_FILE.is_file() else '未找到'}；dry-run 不读取内容）")
    for it in plan:
        size = f"提示词 {len(it['prompt'])} 字符" if it["prompt"] else "提示词 -"
        log(f"[dry-run] {it['id']}  {size:<16}  {it['reason']}")
    n_gen = sum(1 for it in plan if it["action"] == "generate")
    n_skip = sum(1 for it in plan if it["action"] == "skip")
    n_err = sum(1 for it in plan if it["action"] == "error")
    log(f"[dry-run] 共 {len(plan)} 题：将调用 {n_gen} 题，跳过 {n_skip} 题" + (f"，有问题 {n_err} 题" if n_err else "")
        + "。未读取 key，未发起任何请求，未写任何文件。")


def main(argv: list[str] | None = None) -> int:
    _setup_stdio()
    args = parse_args(argv)

    all_ids = discover_h_tasks()
    if not all_ids:
        log(f"未在 {TASKS_DIR} 下找到 H 题目录（H01～H08）。", err=True)
        return 2
    try:
        ids = select_tasks(args.only, all_ids)
    except ArgError as exc:
        log(str(exc), err=True)
        return 2

    plan = make_plan(ids, args.force)

    if args.dry_run:
        print_plan(plan, args)
        return 1 if any(it["action"] == "error" for it in plan) else 0

    failed: list[str] = []
    for it in plan:
        if it["action"] == "error":
            log(f"{it['id']}：{it['reason']}", err=True)
            failed.append(it["id"])
    todo = [it for it in plan if it["action"] == "generate"]
    for it in plan:
        if it["action"] == "skip":
            log(f"{it['id']}：{it['reason']}")
    if not todo:
        log("没有需要生成的题目。")
        return 1 if failed else 0

    try:
        key = read_key()
    except KeyFileError as exc:
        log(str(exc), err=True)
        return 2

    ok = 0
    for idx, it in enumerate(todo, 1):
        log(f"[{idx}/{len(todo)}] {it['id']}：提示词 {len(it['prompt'])} 字符，调用 {args.model} ……")
        try:
            result, attempts = generate_with_retry(key, it["prompt"], args)
        except ApiError as exc:
            log(f"  失败：{exc}", err=True)
            failed.append(it["id"])
            if exc.fatal:
                rest = [t["id"] for t in todo[idx:]]
                if rest:
                    log(f"  认证或余额问题，终止后续题目：{', '.join(rest)}", err=True)
                    failed.extend(rest)
                break
            continue
        n_chars = write_outputs(it, result, attempts, args)
        ok += 1
        log(f"  完成：{n_chars} 字（不含空白），请求 {attempts} 次，已写入 {DRAFT_REL.as_posix()} 和 {RECORD_REL.as_posix()}")

    log(f"结束：成功 {ok} 题，失败 {len(failed)} 题" + (f"（{', '.join(failed)}）" if failed else "") + "。")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
