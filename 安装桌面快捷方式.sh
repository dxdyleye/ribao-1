#!/usr/bin/env bash
# ============================================================================
# 把「蚊媒监测数据处理」安装成银河麒麟桌面图标 + 开始菜单项
#
# 用法：bash 安装桌面快捷方式.sh
#   安装后：① 桌面出现图标，双击即用；② 开始菜单（左下角）可搜索到
#   卸载：bash 安装桌面快捷方式.sh --uninstall
# ============================================================================
set -u

APP_DIR="$(cd "$(dirname "$(readlink -f "${BASH_SOURCE[0]}")")" && pwd)"
LAUNCHER="$APP_DIR/启动蚊媒监测工具.sh"
ICON="$APP_DIR/assets/蚊媒监测工具.svg"
DESK_NAME="蚊媒监测工具.desktop"
APP_LABEL="蚊媒监测数据处理"

BIN_DIR="$HOME/.local/bin"
DESKTOP_BIN="$BIN_DIR/蚊媒监测工具"
APP_DIR_DESKTOP="$HOME/.local/share/applications"
ICON_DIR="$HOME/.local/share/icons/hicolor/scalable/apps"
ICON_INSTALLED="$ICON_DIR/mosquito-monitor.svg"
MENU_FILE="$APP_DIR_DESKTOP/$DESK_NAME"

C_OK=$'\033[32m'; C_ERR=$'\033[31m'; C_OFF=$'\033[0m'
ok()  { printf '%s✓%s %s\n' "$C_OK" "$C_OFF" "$*"; }
bad() { printf '%s✗%s %s\n' "$C_ERR" "$C_OFF" "$*"; }

# 桌面目录：优先中文「桌面」，其次英文「Desktop」
desktop_dir() {
    if [ -d "$HOME/桌面" ]; then printf '%s' "$HOME/桌面"
    elif [ -d "$HOME/Desktop" ]; then printf '%s' "$HOME/Desktop"
    else printf '%s' "$HOME/桌面"; fi
}

# ------------------------------------------------------------- 卸载 ---------
if [ "${1:-}" = "--uninstall" ]; then
    rm -f "$DESKTOP_BIN" "$MENU_FILE" "$(desktop_dir)/$DESK_NAME" "$ICON_INSTALLED"
    command -v update-desktop-database >/dev/null 2>&1 && \
        update-desktop-database "$APP_DIR_DESKTOP" >/dev/null 2>&1
    ok "已卸载桌面图标与菜单项（程序文件未删除）"
    exit 0
fi

# ------------------------------------------------------------- 前置检查 -----
[ -f "$LAUNCHER" ] || { bad "找不到启动脚本：$LAUNCHER"; exit 1; }
[ -f "$ICON" ]     || { bad "找不到图标文件：$ICON"; exit 1; }
chmod +x "$LAUNCHER" 2>/dev/null

bash "$LAUNCHER" --check || {
    bad "环境自检未通过，先解决上面的问题再安装"
    exit 1
}
echo

# ------------------------------------------------------------- 写入文件 -----
DESK_DIR="$(desktop_dir)"
mkdir -p "$BIN_DIR" "$APP_DIR_DESKTOP" "$ICON_DIR" "$DESK_DIR" 2>/dev/null || true

# 1) 中转启动器（放在无空格路径下，供 .desktop 调用）
cat > "$DESKTOP_BIN" <<EOF
#!/usr/bin/env bash
# 由「安装桌面快捷方式.sh」生成：转发到程序目录下的启动器
exec bash "$LAUNCHER" "\$@"
EOF
chmod +x "$DESKTOP_BIN"
ok "中转启动器：$DESKTOP_BIN"

# 2) 图标
cp -f "$ICON" "$ICON_INSTALLED"
chmod 644 "$ICON_INSTALLED" 2>/dev/null || true
ok "图标：$ICON_INSTALLED"

# 3) .desktop 内容（写入失败返回非 0，便于上层如实提示）
write_desktop() {
    if ! cat > "$1" <<EOF
[Desktop Entry]
Type=Application
Version=1.0
Name=$APP_LABEL
Name[zh_CN]=$APP_LABEL
GenericName=蚊媒监测数据处理程序
Comment=读取总库表与广州市表，生成日报 Word、村居一览表 Excel、计算过程与基础数据集
Exec=$DESKTOP_BIN %U
Path=$APP_DIR
Icon=$ICON_INSTALLED
Terminal=false
StartupNotify=true
Categories=Office;
Keywords=蚊媒;BI;ADI;监测;日报;村居一览表;
EOF
    then
        bad "写入失败：$1（目录只读或无权限）"
        return 1
    fi
    chmod +x "$1" 2>/dev/null || true
    # 麒麟/Peony 与 GNOME 双击运行需要“受信任”标记
    if command -v gio >/dev/null 2>&1; then
        gio set "$1" metadata::trusted true 2>/dev/null || true
    fi
    return 0
}

if write_desktop "$MENU_FILE"; then
    ok "开始菜单项：$MENU_FILE"
else
    bad "开始菜单项写入失败：$MENU_FILE"
    exit 1
fi

DESK_PATH="$DESK_DIR/$DESK_NAME"
if write_desktop "$DESK_PATH"; then
    ok "桌面图标：$DESK_PATH"
else
    bad "桌面图标写入失败：$DESK_PATH（可手动把 $MENU_FILE 复制到桌面）"
fi

# 4) 刷新桌面数据库 + 校验
command -v update-desktop-database >/dev/null 2>&1 && \
    update-desktop-database "$APP_DIR_DESKTOP" >/dev/null 2>&1
if command -v desktop-file-validate >/dev/null 2>&1; then
    if desktop-file-validate "$MENU_FILE"; then
        ok "desktop 文件校验通过"
    else
        bad "desktop 文件校验有告警（一般不影响使用）"
    fi
fi

# 若 ~/.local/bin 不在 PATH 中，提示（不影响桌面图标使用）
case ":$PATH:" in
    *":$BIN_DIR:"*) ;;
    *) printf '提示：%s 不在 PATH 中，终端里可直接用完整路径调用\n' "$BIN_DIR" ;;
esac

echo
ok "安装完成"
echo "  · 双击桌面上的「$APP_LABEL」即可启动"
echo "  · 也可在开始菜单搜索「蚊媒」"
echo "  · 终端启动：bash \"$LAUNCHER\"   或   bash \"$LAUNCHER\" --check"
