# -*- coding: utf-8 -*-
"""工具链单测的公共夹具：在临时目录中构造假仓库和假 skill，不触碰真实的 skills/。"""
from __future__ import annotations

import contextlib
import io
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

REPO_ROOT = Path(__file__).resolve().parents[2]
TOOLS_DIR = REPO_ROOT / "tools"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

REQUIRED_SHARED = (
    "guardrails.md",
    "style/output-format.md",
    "style/ai-tone.md",
    "style/defensive.md",
)

SHARED_FILES: Dict[str, str] = {
    "guardrails.md": "# 六条写作红线\n\n## 一、保密先于生成\n\n1. 请勿输入涉密信息。\n",
    "style/output-format.md": "# 输出格式规范\n\n正文用纯文本公文体。\n",
    "style/ai-tone.md": "# AI 味特征\n\n少用评论语。\n",
    "style/defensive.md": "# 防御性表述\n\n少用模糊措辞。\n",
    "rules/gbt9704.md": "# GB/T 9704 要点\n\n版心尺寸见 5.2.1。\n",
}

UNITS = ["党政机关", "基层", "社区", "公立医院", "高校", "中小学", "国有企业"]
ROLES = ["教学模式", "高效模式", "转译模式", "审稿模式", "严谨模式"]


def write_text(path: Path, text: str) -> Path:
    """统一以 UTF-8、LF 换行写入（不用 write_text(newline=...)，保持 3.9 兼容）。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(text.encode("utf-8"))
    return path


def write_json(path: Path, data: object) -> Path:
    return write_text(path, json.dumps(data, ensure_ascii=False, indent=2) + "\n")


def skill_md(name: str, description: Optional[str] = None, extra_body: str = "", extra_fields: str = "") -> str:
    desc = description or "演示用的公文 skill，用于测试工具链。用户说\"测试一下\"时使用。其他情况不使用。"
    return (
        "---\n"
        f"name: {name}\n"
        f"description: {desc}\n"
        "license: MIT\n"
        f"{extra_fields}"
        "---\n"
        "\n"
        f"# 演示技能（{name}）\n"
        "\n"
        "一段话说明解决什么问题。\n"
        "\n"
        "## 开工检查\n"
        "1. 保密：按 `references/_shared/guardrails.md` 第一条执行。\n"
        "2. 输出前读 references/_shared/style/output-format.md。\n"
        "\n"
        "## 自检清单\n"
        "有 Python 环境时运行：`python scripts/<脚本>.py <文件>`，示例资料见 `references/*.md`。\n"
        "\n"
        "## 参考资料\n"
        "| 文件 | 何时读取 |\n"
        "|---|---|\n"
        "| `references/checklist.md` | 自检时 |\n"
        "| `evals/evals.json` | 评测时 |\n"
        f"{extra_body}"
    )


def make_evals(name: str, count: int = 10) -> dict:
    items = []
    for i in range(1, count + 1):
        tags = ["测试"]
        if i <= 2:
            tags.append("防编造")
        if i == 3:
            tags.append("保密")
        items.append({
            "id": i,
            "prompt": f"帮我写第 {i} 份材料",
            "expected_output": "一份合规的材料",
            "files": ["evals/files/素材.txt"] if i == 1 else [],
            "expectations": ["正文不含 Markdown 符号"],
            "unit_type": UNITS[i % len(UNITS)],
            "role_mode": ROLES[i % len(ROLES)],
            "tags": tags,
        })
    return {"skill_name": name, "evals": items}


def make_trigger(positive: int = 10, negative: int = 10) -> List[dict]:
    data = [{"query": f"应触发的说法 {i}", "should_trigger": True} for i in range(positive)]
    data += [{"query": f"不应触发的说法 {i}", "should_trigger": False} for i in range(negative)]
    return data


class FakeRepo:
    """临时假仓库。"""

    def __init__(self, root: Path) -> None:
        self.root = root
        for rel, text in SHARED_FILES.items():
            write_text(root / "shared" / rel, text)
        (root / "tools").mkdir(parents=True, exist_ok=True)

    def skill_dir(self, name: str) -> Path:
        return self.root / "skills" / name

    def add_skill(
        self,
        name: str = "gongwen-demo",
        *,
        status: str = "beta",
        shared: Optional[List[dict]] = None,
        bundle: Optional[List[str]] = None,
        skill_text: Optional[str] = None,
        evals_count: int = 10,
        sync: bool = True,
    ) -> Path:
        d = self.skill_dir(name)
        write_text(d / "SKILL.md", skill_text if skill_text is not None else skill_md(name))
        write_text(d / "README.md", f"# {name}\n\n给人看的说明。\n")
        write_text(d / "references" / "checklist.md", "# 自检清单\n\n- [ ] 标题规范\n- [ ] 无 Markdown 符号\n")
        write_text(d / "evals" / "files" / "素材.txt", "某市某局 2026 年工作记录（虚构）。\n")
        write_json(d / "evals" / "evals.json", make_evals(name, evals_count))
        write_json(d / "evals" / "trigger.json", make_trigger())
        if shared is None:
            shared = [{"src": s} for s in REQUIRED_SHARED]
        if bundle is None:
            bundle = [
                "references/_shared/guardrails.md",
                "references/_shared/style/output-format.md",
                "references/checklist.md",
            ]
        write_json(d / "manifest.json", {
            "name": name,
            "version": "0.1.0",
            "status": status,
            "shared": shared,
            "prompt_bundle": bundle,
        })
        if sync:
            self.sync()
        return d

    def add_marketplace(self, skills: Optional[List[str]] = None, plugin_name: str = "gongwen",
                        version: Optional[str] = "0.1.0") -> Path:
        if skills is None:
            skills = sorted(p.name for p in (self.root / "skills").iterdir() if p.is_dir())
        data = {
            "name": "gongwen-skills",
            "owner": {"name": "test"},
            "plugins": [{
                "name": plugin_name,
                "source": "./",
                "strict": False,
                "skills": [f"./skills/{s}" for s in skills],
            }],
        }
        if version is not None:
            data["metadata"] = {"description": "测试", "version": version}
        return write_json(self.root / ".claude-plugin" / "marketplace.json", data)

    def sync(self) -> None:
        import sync_shared
        code, out = run_quiet(sync_shared.main, ["--root", str(self.root)])
        if code != 0:
            raise AssertionError(f"同步失败：{out}")


def run_quiet(func: Callable[..., int], argv: List[str], **kwargs) -> Tuple[int, str]:
    """调用脚本的 main，捕获输出，返回 (退出码, 输出文本)。"""
    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer), contextlib.redirect_stderr(buffer):
        try:
            code = func(argv, **kwargs)
        except SystemExit as exc:  # argparse 参数错误
            code = exc.code if isinstance(exc.code, int) else 2
    return code, buffer.getvalue()


class TempRepoTestCase(unittest.TestCase):
    """每个用例一个干净的临时仓库。"""

    def setUp(self) -> None:
        self._tmp = tempfile.mkdtemp(prefix="gongwen-tools-")
        self.tmp = Path(self._tmp)
        self.repo = FakeRepo(self.tmp / "repo")
        self.root = self.repo.root

    def tearDown(self) -> None:
        shutil.rmtree(self._tmp, ignore_errors=True)
