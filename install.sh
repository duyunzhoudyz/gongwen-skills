#!/usr/bin/env bash
# gongwen-skills 一键安装脚本（macOS、Linux）
#
# 用法：在仓库根目录打开终端，运行
#   bash install.sh                    # 安装到本机探测到的全部 Agent 工具
#   bash install.sh --dry-run          # 只演练，不写入任何文件
#   bash install.sh --install-deps     # 同时安装排版依赖 python-docx、openpyxl
#   bash install.sh --pack hospital    # 加装行业包（可选 hospital、university、k12、soe、grassroots、all）
#   bash install.sh --target qwen      # 只装到指定工具
#   bash install.sh --uninstall        # 列出可卸载的 skill，加 --yes 才删除
#
# 本脚本做三件事：
#   1. 检查 Python 3.9 及以上，没有就给出下载指引；
#   2. 检查排版依赖 python-docx、openpyxl，只有你同意或传了 --install-deps 时才执行 pip install；
#   3. 调用 python3 tools/install.py --target auto（你自己传了 --target 时以你的为准）。
# 除 --install-deps 外，其余参数原样传给 tools/install.py，可用 --help 查看全部参数。
# 本脚本自身不联网；只有执行 pip install 时，pip 会联网下载依赖。
# 兼容 macOS 自带的 bash 3.2。

set -u

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
INSTALL_PY="$SCRIPT_DIR/tools/install.py"
PIP_MIRROR="https://pypi.tuna.tsinghua.edu.cn/simple"
export PYTHONIOENCODING=utf-8

INSTALL_DEPS=0
DRY_RUN=0
HAS_TARGET=0
SKIP_DEPS=0
SHOW_HELP=0
UNINSTALL=0
PASS_ARGS=()

for arg in "$@"; do
  case "$arg" in
    --install-deps) INSTALL_DEPS=1 ;;
    --dry-run) DRY_RUN=1; PASS_ARGS+=("$arg") ;;
    --target|--target=*) HAS_TARGET=1; PASS_ARGS+=("$arg") ;;
    --uninstall) UNINSTALL=1; SKIP_DEPS=1; PASS_ARGS+=("$arg") ;;
    --detect|--list-targets) SKIP_DEPS=1; PASS_ARGS+=("$arg") ;;
    -h|--help) SHOW_HELP=1; SKIP_DEPS=1; PASS_ARGS+=("--help") ;;
    *) PASS_ARGS+=("$arg") ;;
  esac
done

if [ -t 1 ]; then
  C_STEP=$'\033[36m'; C_WARN=$'\033[33m'; C_END=$'\033[0m'
else
  C_STEP=""; C_WARN=""; C_END=""
fi
step() { printf '\n%s== %s ==%s\n' "$C_STEP" "$1" "$C_END"; }
warn() { printf '%s%s%s\n' "$C_WARN" "$1" "$C_END"; }

# 依次尝试 python3、python，找到 3.9 及以上的解释器后写入 PY 和 PY_VERSION
PY=""
PY_VERSION=""
TOO_OLD=""
find_python() {
  local cand out status
  for cand in python3 python; do
    command -v "$cand" >/dev/null 2>&1 || continue
    out="$("$cand" -c 'import sys; v=sys.version_info; print("%d.%d.%d" % (v[0], v[1], v[2])); sys.exit(0 if v >= (3, 9) else 1)' 2>/dev/null)"
    status=$?
    [ -n "$out" ] || continue
    if [ "$status" -eq 0 ]; then
      PY="$cand"
      PY_VERSION="$out"
      return 0
    fi
    TOO_OLD="$TOO_OLD $cand $out"
  done
  return 1
}

show_python_guide() {
  warn "未找到 Python 3.9 及以上版本。"
  if [ -n "$TOO_OLD" ]; then
    warn "找到的版本过旧：$TOO_OLD"
  fi
  echo "安装方法："
  if [ "$(uname -s 2>/dev/null)" = "Darwin" ]; then
    echo "  macOS：打开 https://www.python.org/downloads/macos/ 下载 3.9 及以上版本的安装包（推荐 3.12），"
    echo "         或者用 Homebrew：brew install python"
  else
    echo "  Ubuntu、Debian、统信 UOS、银河麒麟等：sudo apt install python3 python3-pip"
    echo "  CentOS、RHEL、Fedora、openEuler 等：sudo dnf install python3 python3-pip"
    echo "  其他发行版：使用系统自带的软件包管理器安装 python3"
  fi
  echo "  装好后重新打开终端，再运行：bash install.sh"
  echo "暂时不装 Python 也能用："
  echo "  Claude Code 用户：在 Claude Code 中运行 /plugin marketplace add duyunzhoudyz/gongwen-skills"
  echo "  网页大模型用户：打开 dist/prompts/，把提示词文件内容复制到对话开头。"
}

if [ "$SHOW_HELP" -eq 1 ]; then
  echo "install.sh 自身参数：--install-deps 自动安装排版依赖；--dry-run 只演练不写入。以下为 tools/install.py 的参数："
fi
if [ "$DRY_RUN" -eq 1 ]; then
  echo "[演练模式] 只打印将要执行的操作，不安装依赖、不写入任何文件。"
fi

if [ ! -f "$INSTALL_PY" ]; then
  warn "找不到 $INSTALL_PY。请在 gongwen-skills 仓库根目录运行本脚本，并确认下载的文件完整。"
  exit 2
fi

step "检查 Python"
if ! find_python; then
  show_python_guide
  exit 1
fi
echo "已找到 Python $PY_VERSION（$PY）"

if [ "$SKIP_DEPS" -eq 0 ]; then
  step "检查排版依赖（python-docx、openpyxl）"
  MISSING="$("$PY" -c 'import importlib.util as u; print(" ".join(n for n, m in (("python-docx", "docx"), ("openpyxl", "openpyxl")) if u.find_spec(m) is None))' 2>/dev/null)" \
    || MISSING="python-docx openpyxl"
  if [ -z "$MISSING" ]; then
    echo "依赖齐全，排版功能可以直接使用。"
  else
    PIP_SHOWN="$PY -m pip install python-docx openpyxl"
    warn "缺少：$MISSING。不装也能写公文，只是排版会改为输出 WPS/Word 手动设置清单。"
    DO_INSTALL=0
    if [ "$DRY_RUN" -eq 1 ]; then
      echo "[演练] 不安装依赖。正式安装时可加 --install-deps，或手动运行：$PIP_SHOWN"
    elif [ "$INSTALL_DEPS" -eq 1 ]; then
      DO_INSTALL=1
    elif [ -t 0 ] && [ -t 1 ]; then
      printf '是否现在用 pip 安装？输入 y 确认，直接回车跳过：'
      ANSWER=""
      read -r ANSWER || ANSWER=""
      case "$ANSWER" in
        y|Y|yes|YES|是) DO_INSTALL=1 ;;
        *) echo "已跳过。需要时运行：$PIP_SHOWN" ;;
      esac
    else
      echo "已跳过。需要时运行：$PIP_SHOWN（或给本脚本加 --install-deps）"
    fi
    if [ "$DO_INSTALL" -eq 1 ]; then
      echo "正在运行：$PIP_SHOWN"
      if "$PY" -m pip install python-docx openpyxl; then
        echo "排版依赖安装完成。"
      else
        warn "依赖安装失败，不影响安装 skill。可以稍后手动处理："
        echo "  网络较慢时改用国内镜像：$PIP_SHOWN -i $PIP_MIRROR"
        echo "  提示没有 pip 时：Debian、Ubuntu 运行 sudo apt install python3-pip；其他系统运行 $PY -m ensurepip --upgrade"
        echo "  提示 externally-managed-environment 时：Debian、Ubuntu 可运行 sudo apt install python3-docx python3-openpyxl，"
        echo "    或者运行 $PIP_SHOWN --user --break-system-packages（只为当前用户安装）"
      fi
    fi
  fi
fi

if [ "$UNINSTALL" -eq 1 ]; then
  step "卸载 skill"
else
  step "安装 skill"
fi
FINAL_ARGS=()
if [ "$HAS_TARGET" -eq 0 ]; then
  FINAL_ARGS+=(--target auto)
fi
# bash 3.2 在 set -u 下展开空数组会报错，用 ${arr[@]+"${arr[@]}"} 写法规避
"$PY" "$INSTALL_PY" ${FINAL_ARGS[@]+"${FINAL_ARGS[@]}"} ${PASS_ARGS[@]+"${PASS_ARGS[@]}"}
CODE=$?
if [ "$CODE" -eq 3 ]; then
  warn "可以先安装 Claude Code、Qwen Code 等 Agent 工具后重新运行，或用 --target 指定目标。"
elif [ "$CODE" -ne 0 ]; then
  warn "安装器返回错误（退出码 $CODE），请查看上方提示。"
fi
exit "$CODE"
