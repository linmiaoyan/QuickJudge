#!/usr/bin/env bash
# QuickJudge 主机（处理中心）在 Linux 上装成 systemd 服务。
#
#   sudo bash deploy/install_linux.sh
#
# 装完之后：
#   systemctl status quickjudge          看状态
#   systemctl restart quickjudge         重启
#   journalctl -u quickjudge -f          跟日志
#
# 卸载：
#   sudo systemctl disable --now quickjudge
#   sudo rm /etc/systemd/system/quickjudge.service
#
# 这台机器只做批阅：扫描仪在别的电脑上，通过局域网把图传进来。
# 所以这里不需要 NAPS2，也不需要 X11/图形环境。

set -euo pipefail

APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
UNIT=/etc/systemd/system/quickjudge.service
PORT="${PORTAL_PORT:-5001}"

say()  { printf '\n\033[1m== %s\033[0m\n' "$*"; }
warn() { printf '\033[33m[注意] %s\033[0m\n' "$*"; }
die()  { printf '\033[31m[失败] %s\033[0m\n' "$*" >&2; exit 1; }

# ---------- 前置检查 ----------
[ "$(id -u)" -eq 0 ] || die "需要用 root 装系统服务：sudo bash deploy/install_linux.sh"
[ -f "$APP_DIR/app.py" ] || die "在 $APP_DIR 下没找到 app.py，这个脚本要放在仓库的 deploy/ 里"

command -v systemctl >/dev/null 2>&1 || die "这台机器没有 systemd，无法装成服务"

# 挑一个够新的 python3
PY=""
for c in python3.12 python3.11 python3.10 python3.9 python3.8 python3; do
  if command -v "$c" >/dev/null 2>&1; then
    if "$c" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 8) else 1)' 2>/dev/null; then
      PY="$(command -v "$c")"
      break
    fi
  fi
done
[ -n "$PY" ] || die "需要 Python 3.8 或更新的版本。Debian/Ubuntu: apt install python3 python3-venv"

say "Python: $PY ($("$PY" -V 2>&1))"
say "目录:   $APP_DIR"

# ---------- 运行身份 ----------
# 不用 root 跑 Flask：这个进程对外开端口，出事时权限越小损失越小。
RUN_USER="${QJ_USER:-}"
if [ -z "$RUN_USER" ]; then
  if [ -n "${SUDO_USER:-}" ] && [ "$SUDO_USER" != "root" ]; then
    RUN_USER="$SUDO_USER"
  else
    RUN_USER=quickjudge
  fi
fi

if ! id "$RUN_USER" >/dev/null 2>&1; then
  say "创建系统用户 $RUN_USER"
  # --user-group 显式建一个同名组：systemd 的 Group= 找不到组会直接启动失败，
  # 而 useradd --system 在不同发行版上建不建组并不一致。
  useradd --system --user-group --create-home --shell /usr/sbin/nologin "$RUN_USER"
fi
say "运行身份: $RUN_USER"

# ---------- 虚拟环境 ----------
if [ ! -x "$APP_DIR/.venv/bin/python" ]; then
  say "创建虚拟环境 .venv"
  "$PY" -m venv "$APP_DIR/.venv" || die "建虚拟环境失败。Debian/Ubuntu 需要先 apt install python3-venv"
fi

say "安装依赖"
"$APP_DIR/.venv/bin/python" -m pip install --upgrade pip --quiet
"$APP_DIR/.venv/bin/python" -m pip install -r "$APP_DIR/requirements.txt" --quiet
"$APP_DIR/.venv/bin/python" -c 'import flask, openai, requests, PIL; print("依赖自检通过")' \
  || die "依赖装上了但 import 不进来，看上面的报错"

# ---------- 目录与权限 ----------
mkdir -p "$APP_DIR/logs" "$APP_DIR/data"
# 服务以 $RUN_USER 身份写 data/ 和 logs/，整个仓库交给它最省事；
# 只 chown 这两个目录会在"导入练习""保存配置"等路径上继续踩权限。
chown -R "$RUN_USER:$RUN_USER" "$APP_DIR"

# ---------- 密钥 ----------
if [ ! -f "$APP_DIR/.env" ]; then
  warn "没找到 $APP_DIR/.env —— 模型接口的密钥在这里配，服务能起但批阅会失败。"
  warn "照着 .env.example 建一份，然后 systemctl restart quickjudge"
fi

# ---------- 安装服务 ----------
say "写入 $UNIT"
sed -e "s#__APP_DIR__#$APP_DIR#g" \
    -e "s#__RUN_USER__#$RUN_USER#g" \
    "$APP_DIR/deploy/quickjudge.service" > "$UNIT"
chmod 644 "$UNIT"

systemctl daemon-reload
systemctl enable quickjudge >/dev/null
systemctl restart quickjudge

# ---------- 等它就绪 ----------
# 用 .venv 里的 python 探活，不用 curl：最小化安装的服务器上 curl 不一定有，
# 而 python 刚刚才确认过能用。
say "等待服务就绪（最多 60 秒）"
probe() {
  "$APP_DIR/.venv/bin/python" - "$PORT" <<'PY' >/dev/null 2>&1
import sys, urllib.request
urllib.request.urlopen('http://127.0.0.1:%s/favicon.ico' % sys.argv[1], timeout=2)
PY
}

ok=0
for i in $(seq 1 60); do
  if probe; then
    ok=1
    printf '就绪，用时约 %s 秒\n' "$i"
    break
  fi
  sleep 1
done

if [ "$ok" -ne 1 ]; then
  warn "60 秒还没响应。看日志："
  journalctl -u quickjudge -n 40 --no-pager || true
  die "服务没起来"
fi

# ---------- 打印访问地址 ----------
say "装好了"
printf '本机：      http://127.0.0.1:%s/\n' "$PORT"
for ip in $(hostname -I 2>/dev/null || true); do
  case "$ip" in
    *:*) ;;
    *) printf '局域网：    http://%s:%s/        （扫描点填这个）\n' "$ip" "$PORT" ;;
  esac
done
cat <<EOF

下一步
  1. 确认密钥：  $APP_DIR/.env 里有 LLM_API_KEY / LLM_BASE_URL / LLM_MODEL
     改完执行：   systemctl restart quickjudge
  2. 放行端口：  防火墙要开 $PORT，否则扫描点和手机连不进来
                 ufw allow $PORT/tcp     （或 firewall-cmd --add-port=$PORT/tcp --permanent）
  3. 建扫描点：  浏览器打开上面的地址 → 侧边栏「扫描点」→ 新建 → 复制链接
                 拿去扫描仪那台 Windows 电脑打开，按页面提示装。

常用命令
  systemctl status quickjudge
  systemctl restart quickjudge
  journalctl -u quickjudge -f
EOF
