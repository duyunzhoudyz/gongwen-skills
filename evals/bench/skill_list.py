"""评测涉及的 skill 清单：默认取插件市场清单中的全部 skill，也可以用名称或预设筛选。

预设：
    v0.1   v0.1 的 7 个 skill（复现 v0.1 评测时使用）
"""
import json
from pathlib import Path

V01 = ["gongwen", "gongwen-report", "gongwen-title", "gongwen-humanize",
       "gongwen-direct", "gongwen-review", "gongwen-format"]
PRESETS = {"v0.1": V01}


def marketplace_skills(repo):
    """按插件市场清单的顺序返回全部 skill 名称（去重）。"""
    data = json.loads((Path(repo) / ".claude-plugin" / "marketplace.json").read_text(encoding="utf-8"))
    names = []
    for plugin in data.get("plugins", []):
        for path in plugin.get("skills", []):
            name = path.rstrip("/").split("/")[-1]
            if name not in names:
                names.append(name)
    return names


def select_skills(repo, wanted=None):
    """wanted 为空时返回全部；可以写 skill 名称或预设名（如 v0.1）。名称不存在时报错。"""
    all_names = marketplace_skills(repo)
    if not wanted:
        return all_names
    chosen = []
    for item in wanted:
        for name in PRESETS.get(item, [item]):
            if name not in all_names:
                raise SystemExit(f"插件市场清单中没有 skill：{name}")
            if name not in chosen:
                chosen.append(name)
    return [n for n in all_names if n in chosen]


def ordered_unique(names, repo=None):
    """保持插件市场清单中的顺序去重；清单中没有的放在最后。"""
    seen = []
    for n in names:
        if n not in seen:
            seen.append(n)
    if repo is None:
        return seen
    order = marketplace_skills(repo)
    return sorted(seen, key=lambda n: order.index(n) if n in order else len(order))
