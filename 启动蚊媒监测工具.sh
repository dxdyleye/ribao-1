#!/usr/bin/env bash
# ============================================================================
# 蚊媒监测数据处理程序 —— 银河麒麟（Kylin V10 / Linux）启动器
#
# 用法：
#   双击运行          直接启动图形界面
#   bash 启动蚊媒监测工具.sh           同上（终端里也能看到日志）
#   bash 启动蚊媒监测工具.sh --check   只做环境自检（Python/tkinter/依赖），不启动界面
#   bash 启动蚊媒监测工具.sh --install 检查依赖，缺失时尝试自动安装到用户目录
#
# 说明：脚本会自动切到自身所在目录，因此从桌面快捷方式、终端、任意路径运行都一样。
#       启动失败时会把原因写入 启动日志.log，并尽量弹窗提示。
# ============================================================================
set -u

APP_DIR="$(cd "$(dirname "$(readlink -f "${BASH_SOURCE[0]}")")" && pwd)"
cd "$APP_DIR" || exit 1
LOG="$APP_DIR/启动日志.log"
MODE="${1:-run}"

# 运行必需的模块（pyinstaller/lxml 是打包用，运行 GUI 不需要单独装）
DEPS="tkinter pandas openpyxl docx"
PIP_DEPS=("pandas==2.0.3" "openpyxl==3.1.5" "python-docx==1.1.2")

C_OK=$'\033[32m'; C_ERR=$'\033[31m'; C_WARN=$'\033[33m'; C_OFF=$'\033[0m'
say()  { printf '%s\n' "$*"; }
ok()   { printf '%s✓%s %s\n' "$C_OK" "$C_OFF" "$*"; }
bad()  { printf '%s✗%s %s\n' "$C_ERR" "$C_OFF" "$*"; }
warn() { printf '%s!%s %s\n' "$C_WARN" "$C_OFF" "$*"; }

# ---------------------------------------------------------------- 弹窗提示 --
# 图形界面起不来时，双击启动看不到终端输出，用弹窗把原因告诉用户
popup() {
    local title="$1" text="$2"
    if command -v zenity >/dev/null 2>&1; then
        zenity --error --title="$title" --width=560 --text="$text" 2>/dev/null && return 0
    fi
    if command -v xmessage >/dev/null 2>&1 && [ -n "${DISPLAY:-}" ]; then
        xmessage -title "$title" "$text" 2>/dev/null && return 0
    fi
    printf '%s\n%s\n' "$title" "$text" >&2
}

# ------------------------------------------------------------ 找 Python -----
find_python() {
    local cand
    for cand in "${PYTHON3:-}" python3 python3.8 python3.9 python3.10 python3.11 python3.12; do
        [ -n "$cand" ] || continue
        if command -v "$cand" >/dev/null 2>&1; then
            printf '%s' "$cand"; return 0
        fi
    done
    return 1
}

# ------------------------------------------------------------ 依赖自检 ------
# 返回 0=齐全  1=缺模块（缺失模块名写入 MISSING）
MISSING=""
check_deps() {
    local py="$1" mod out
    MISSING=""
    for mod in $DEPS; do
        out="$("$py" -c "import $mod" 2>&1)" || { MISSING="$MISSING $mod"; }
    done
    [ -z "$MISSING" ]
}

print_env() {
    say "程序目录：$APP_DIR"
    say "Python  ：$PY_BIN（$("$PY_BIN" -V 2>&1)）"
    if check_deps "$PY_BIN"; then
        ok "运行依赖齐全：$DEPS"
    else
        bad "缺少模块：$MISSING"
        return 1
    fi
}

# ------------------------------------------------------------ 自动安装 ------
try_install() {
    local py="$1" pip_cmd
    if "$py" -m pip --version >/dev/null 2>&1; then
        pip_cmd=("$py" -m pip)
    elif command -v pip3 >/dev/null 2>&1; then
        pip_cmd=(pip3)
    else
        bad "未找到 pip，无法自动安装依赖"; return 1
    fi

    say "尝试安装缺失模块到用户目录（${PIP_DEPS[*]}）…"
    if "${pip_cmd[@]}" install --user --no-input --disable-pip-version-check "${PIP_DEPS[@]}" 2>&1 | tail -n 20; then
        check_deps "$py" && { ok "依赖安装完成"; return 0; }
    fi
    warn "按锁定版本安装未成功，改用不锁版本再试一次…"
    "${pip_cmd[@]}" install --user --no-input --disable-pip-version-check pandas openpyxl python-docx 2>&1 | tail -n 20
    check_deps "$py" && { ok "依赖安装完成"; return 0; }
    return 1
}

# ================================================================ 主流程 ====
PY_BIN="$(find_python)" || {
    bad "未找到 python3"
    popup "蚊媒监测数据处理 —— 启动失败" \
"未找到 python3。

请先安装：打开终端执行
    sudo apt install python3 python3-tk"

    exit 1
}
say "使用解释器：$PY_BIN"

if ! "$PY_BIN" -c "import tkinter" >/dev/null 2>&1; then
    bad "该 Python 缺少 tkinter（图形界面库）"
    popup "蚊媒监测数据处理 —— 启动失败" \
"$PY_BIN 缺少 tkinter 图形界面库。

请在终端执行（需要输入登录密码）：
    sudo apt install python3-tk

装好后重新双击本程序即可。"
    exit 1
fi

case "$MODE" in
    --check|check)
        print_env && { ok "自检通过，可正常启动"; exit 0; } || exit 1
        ;;
    --install|install)
        print_env || {
            try_install "$PY_BIN" || {
                bad "自动安装失败，请手动执行：pip3 install --user pandas openpyxl python-docx"
                exit 1
            }
        }
        ok "环境就绪"
        exit 0
        ;;
esac

# 正常启动：依赖缺失时先尝试自动补齐
if ! check_deps "$PY_BIN"; then
    warn "缺少模块：$MISSING —— 尝试自动安装…"
    if ! try_install "$PY_BIN"; then
        popup "蚊媒监测数据处理 —— 缺少依赖" \
"缺少 Python 模块：$MISSING

请联网后在终端执行：
    pip3 install --user pandas openpyxl python-docx

或执行程序目录下的：
    bash 启动蚊媒监测工具.sh --install"
        exit 1
    fi
fi

say "正在启动图形界面…（日志：$LOG）"
: > "$LOG"
"$PY_BIN" main.py >>"$LOG" 2>&1
code=$?
if [ $code -ne 0 ]; then
    bad "程序退出码 $code，最后 20 行日志："
    tail -n 20 "$LOG" >&2
    popup "蚊媒监测数据处理 —— 运行出错" \
"程序异常退出（退出码 $code）。

日志文件：$LOG

最后几行信息：
$(tail -n 12 "$LOG")"
    exit $code
fi
exit 0
