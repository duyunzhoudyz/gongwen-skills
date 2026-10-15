# -*- coding: utf-8 -*-
"""测试公用：加载机检引擎和规则数据。"""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

REPO_ROOT = Path(__file__).resolve().parents[2]
SCRIPTS_DIR = REPO_ROOT / "shared" / "scripts"
LINT_PATH = SCRIPTS_DIR / "gongwen_lint.py"
RULES_PATH = SCRIPTS_DIR / "lint_rules.json"
PATTERNS_PATH = SCRIPTS_DIR / "style_patterns.json"
EVAL_FILES = REPO_ROOT / "skills" / "gongwen-review" / "evals" / "files"


def load_lint():
    """按文件路径加载引擎（不依赖 sys.path），同一进程内只加载一次。"""
    name = "gongwen_lint_under_test"
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, LINT_PATH)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)  # type: ignore[union-attr]
    return module


gl = load_lint()
LINTER = gl.Linter()


def load_json(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def run(text: str, doc_type: Optional[str] = None, checks: str = "all") -> Dict[str, Any]:
    return gl.lint_text(text, doc_type, checks, linter=LINTER)


def ids(result: Dict[str, Any]) -> List[str]:
    return [i["id"] for i in result["issues"]]


def issues_of(result: Dict[str, Any], rule_id: str) -> List[Dict[str, Any]]:
    return [i for i in result["issues"] if i["id"] == rule_id]
