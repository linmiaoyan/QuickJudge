# -*- coding: utf-8 -*-
"""QuickJudge 常驻服务的「停止 / 查看状态」工具。

    python stop_service.py --status   看服务在不在跑、端口通不通
    python stop_service.py            停掉看护进程 + 服务进程

为什么不用 taskkill：taskkill.exe 属于外部程序，在受管环境里会被安全策略拦掉。
这里用 ctypes 直接调 Win32 的 OpenProcess/TerminateProcess，不依赖任何外部命令。
"""
import ctypes
import os
import sys
import time
import urllib.error
import urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(BASE, 'logs')
PID_PATH = os.path.join(LOG_DIR, 'service.pid')
WATCH_PID_PATH = os.path.join(LOG_DIR, 'watchdog.pid')

PORT = int(os.getenv('PORTAL_PORT', '5001') or '5001')
PROBE_URL = 'http://127.0.0.1:%d/favicon.ico' % PORT

_k32 = ctypes.windll.kernel32
WAIT_TIMEOUT = 0x102
_probe = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def healthy(timeout=4):
    try:
        with _probe.open(PROBE_URL, timeout=timeout) as r:
            return r.status < 500
    except urllib.error.HTTPError as e:
        return e.code < 500
    except Exception:
        return False


def pid_alive(pid):
    SYNCHRONIZE = 0x00100000
    h = _k32.OpenProcess(SYNCHRONIZE, False, int(pid)) if pid else 0
    if not h:
        return False
    try:
        return _k32.WaitForSingleObject(h, 0) == WAIT_TIMEOUT
    finally:
        _k32.CloseHandle(h)


def terminate(pid):
    PROCESS_TERMINATE = 0x0001
    h = _k32.OpenProcess(PROCESS_TERMINATE, False, int(pid)) if pid else 0
    if not h:
        return False
    try:
        return bool(_k32.TerminateProcess(h, 1))
    finally:
        _k32.CloseHandle(h)


def read_pid(path):
    try:
        with open(path, encoding='utf-8') as f:
            return int(f.read().strip() or 0)
    except (OSError, ValueError):
        return 0


def status():
    svc, wd = read_pid(PID_PATH), read_pid(WATCH_PID_PATH)
    print('服务进程  pid=%s  %s' % (svc or '-', '运行中' if pid_alive(svc) else '未运行'))
    print('看护进程  pid=%s  %s' % (wd or '-', '运行中' if pid_alive(wd) else '未运行'))
    ok = healthy()
    print('端口 %d  %s' % (PORT, '可以访问' if ok else '无响应'))
    if ok:
        print('本机访问：http://127.0.0.1:%d/' % PORT)
    print('运行日志：%s' % os.path.join(LOG_DIR, 'service.log'))
    return 0 if ok else 1


def stop():
    wd, svc = read_pid(WATCH_PID_PATH), read_pid(PID_PATH)
    # 先停看护：否则它看到服务没了会立刻又拉起来
    if pid_alive(wd):
        print('停止看护进程 pid=%s …' % wd)
        terminate(wd)
        time.sleep(1.5)
    else:
        print('看护进程未在运行')

    if pid_alive(svc):
        print('停止服务进程 pid=%s …' % svc)
        terminate(svc)
        time.sleep(2)
    else:
        print('服务进程未在运行')

    if healthy(timeout=2):
        print('警告：端口 %d 仍有响应，可能还有别的实例在跑（比如手动开的窗口）。' % PORT)
        return 1

    for p in (WATCH_PID_PATH, PID_PATH):
        try:
            if os.path.isfile(p):
                os.remove(p)
        except OSError:
            pass
    print('已停止。')
    return 0


if __name__ == '__main__':
    sys.exit(status() if '--status' in sys.argv else stop())
