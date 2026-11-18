# -*- coding: utf-8 -*-
"""QuickJudge 常驻服务入口 —— 不依赖 WorkBuddy，也不依赖任何终端窗口。

用 pythonw.exe 启动本文件：
  * 没有控制台窗口（关掉 WorkBuddy、关掉所有终端都不影响它）；
  * 所有输出写进 logs/service.log（pythonw 下 sys.stdout 是 None，
    而 app.py 里有大量 print()，不接管就会一启动就抛异常）；
  * 两层自愈：服务进程内部异常自己重启；进程被杀/卡死由看护进程拉起；
  * 开机自启靠「启动」文件夹里的 start_quickjudge.vbs 调用 `--watch`。

用法：
    pythonw.exe run_service.py            常驻运行（= --watch，带看护，开机自启用这个）
    pythonw.exe run_service.py --serve    只跑服务本体（由看护进程拉起，一般不用手点）
    python   run_service.py --self-test   只验证日志与模块加载，不起服务、不占端口
"""
import ctypes
import datetime
import os
import runpy
import subprocess
import sys
import time
import traceback
import urllib.error
import urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

LOG_DIR = os.path.join(BASE, 'logs')
os.makedirs(LOG_DIR, exist_ok=True)
LOG_PATH = os.path.join(LOG_DIR, 'service.log')
PID_PATH = os.path.join(LOG_DIR, 'service.pid')       # 服务本体的 pid
WATCH_PID_PATH = os.path.join(LOG_DIR, 'watchdog.pid')  # 看护进程的 pid

# app.py 自己从这里读端口（第 7029 行），看护进程必须用同一个值做探活，
# 否则会出现"端口探不通 → 反复起新实例 → 全都绑不上"的死循环。
PORT = int(os.getenv('PORTAL_PORT', '5001') or '5001')
PROBE_URL = 'http://127.0.0.1:%d/favicon.ico' % PORT

# 日志超过 5MB 就滚一份，避免长期跑把磁盘塞满
try:
    if os.path.isfile(LOG_PATH) and os.path.getsize(LOG_PATH) > 5 * 1024 * 1024:
        os.replace(LOG_PATH, LOG_PATH + '.1')
except OSError:
    pass

# 必须在 import app 之前接管：pythonw 下 sys.stdout 为 None
_log = open(LOG_PATH, 'a', encoding='utf-8', buffering=1)
sys.stdout = _log
sys.stderr = _log

_k32 = ctypes.windll.kernel32
CREATE_NO_WINDOW = 0x08000000
WAIT_TIMEOUT = 0x102

# 本机走代理；不绕开的话探活请求会被送去代理服务器，永远失败
_probe = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def log(msg=''):
    print('[%s] %s' % (datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), msg), flush=True)


def healthy(timeout=5):
    """探活用 favicon：静态小文件、不需要登录、不碰任何业务逻辑。

    返回 404/403 也算活着——只要端口在用 HTTP 回话，服务就是好的。
    """
    try:
        with _probe.open(PROBE_URL, timeout=timeout) as r:
            return r.status < 500
    except urllib.error.HTTPError as e:
        return e.code < 500
    except Exception:
        return False


def tcp_alive(timeout=2):
    """端口能不能握手。

    这是判断「该死还是该等」的关键：批阅高峰时 8 个模型请求把服务压满，
    HTTP 探活可能超时，但监听队列照样能被系统 accept——连接是通的。
    真死了的表现是**连接被拒绝**，不是「慢」。
    只按 HTTP 超时就去杀进程，会把正在批阅的服务打断（2026-11-07 凌晨就因为
    这个误杀了 6 次）。所以要分两种失败分别计数。
    """
    import socket
    try:
        with socket.create_connection(('127.0.0.1', PORT), timeout=timeout):
            return True
    except OSError:
        return False


def _pid_alive(pid):
    """Win32 下**不能**用 os.kill(pid, 0) 探活：Windows 上它走 TerminateProcess，
    信号 0 也照样把进程杀掉。必须用 OpenProcess + WaitForSingleObject。
    """
    SYNCHRONIZE = 0x00100000
    try:
        h = _k32.OpenProcess(SYNCHRONIZE, False, int(pid))
    except Exception:
        return False
    if not h:
        return False
    try:
        return _k32.WaitForSingleObject(h, 0) == WAIT_TIMEOUT
    finally:
        _k32.CloseHandle(h)


def _terminate(pid):
    """结束指定进程。用 Win32 而不是 taskkill：taskkill 属于外部程序，
    在受管环境里可能被安全策略拦掉，ctypes 直接调 API 没有这个风险。
    """
    PROCESS_TERMINATE = 0x0001
    try:
        h = _k32.OpenProcess(PROCESS_TERMINATE, False, int(pid))
    except Exception:
        return False
    if not h:
        return False
    try:
        return bool(_k32.TerminateProcess(h, 1))
    finally:
        _k32.CloseHandle(h)


def _read_pid(path):
    try:
        with open(path, encoding='utf-8') as f:
            return int(f.read().strip() or 0)
    except (OSError, ValueError):
        return 0


def _write_pid(path):
    try:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(str(os.getpid()))
    except OSError:
        pass


def _drop_pid(path):
    try:
        if os.path.isfile(path) and _read_pid(path) == os.getpid():
            os.remove(path)
    except OSError:
        pass


def self_test():
    """只验证「日志接管 + 模块能加载」，不起服务、不绑端口、不碰业务数据。"""
    log('=== self-test 开始 ===')
    import importlib.util
    spec = importlib.util.spec_from_file_location('qj_app', os.path.join(BASE, 'app.py'))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    log('app 模块加载成功：PORTAL_PORT=%s' % getattr(mod, 'PORTAL_PORT', '?'))
    log('DATA_DIR=%s' % getattr(mod, 'DATA_DIR', '?'))
    log('=== self-test 通过 ===')
    return 0


def serve():
    """服务本体：反复执行 app.py，内部异常就地重启。

    这一层管的是「Python 里抛异常/进程被 flask 退出」；进程被外部杀掉、
    或者整个卡死不动，由下面的 watch() 兜底。
    """
    # 别人已经在跑同一端口就别抢：重复自启（开机自启 + 手动点）不会起两个实例。
    # ⚠️ 必须在写 pid 文件之前判断：pid 文件是全机唯一的一份记录，
    # 一个马上要退出的重复实例把它改成自己的 pid，真正的服务就再也停不掉了。
    if healthy():
        log('端口 %d 已有实例在服务，本次启动退出（这是正常的，不算失败）' % PORT)
        return 0

    _write_pid(PID_PATH)
    log('=' * 60)
    log('QuickJudge 服务启动（pythonw 常驻模式）')
    log('工作目录：%s' % BASE)
    log('监听端口：%d' % PORT)
    log('日志：%s' % LOG_PATH)
    log('=' * 60)

    app_path = os.path.join(BASE, 'app.py')
    restarts = 0
    try:
        while True:
            try:
                # run_name='__main__'：完全等同「python app.py」，包括启动横幅、
                # 残留批次收尾、NAPS2 初始化这些只在主程序里跑的逻辑。
                runpy.run_path(app_path, run_name='__main__')
                log('app.py 正常返回（服务被主动停止），5 秒后重启')
            except SystemExit as e:
                log('app.py 退出（SystemExit=%s），5 秒后重启' % e)
            except KeyboardInterrupt:
                log('收到中断，退出')
                return 0
            except BaseException:
                log('app.py 异常退出，5 秒后重启：')
                traceback.print_exc()
            restarts += 1
            log('已重启 %d 次' % restarts)
            time.sleep(5)
    finally:
        _drop_pid(PID_PATH)


def _spawn_service():
    """拉起服务子进程。父进程只需拿着句柄，不做别的。"""
    exe = sys.executable or ''
    if not exe.lower().endswith('pythonw.exe'):
        # 从 python.exe 里跑 --watch 时，也要让子进程没窗口
        alt = os.path.join(os.path.dirname(exe), 'pythonw.exe')
        if os.path.isfile(alt):
            exe = alt
    args = [exe, os.path.join(BASE, 'run_service.py'), '--serve']
    log('拉起服务进程：%s' % ' '.join('"%s"' % a for a in args))
    return subprocess.Popen(
        args, cwd=BASE, close_fds=True,
        stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        creationflags=CREATE_NO_WINDOW,
    )


def _wait_healthy(limit=180):
    """等端口能回话。app.py 启动要跑 NAPS2 初始化 + 密钥保险箱，留足 3 分钟。"""
    deadline = time.time() + limit
    while time.time() < deadline:
        if healthy():
            return True
        time.sleep(2)
    return False


def watch():
    """看护进程：保证「端口随时有人应答」，缺了就拉起来，卡死就换掉。

    这是外层的保命机制——服务进程被任务管理器结束、关机残留、扫描时崩掉，
    都会在一分钟内恢复到可用状态，不需要老师去点任何东西。

    但「重启」是破坏性动作：正在批阅的批次会被打断。所以两类失败分开计数——
      * 连端口都被拒绝   → 进程真的死了，等 1 分钟就重启（DEAD_LIMIT）
      * 端口通但 HTTP 慢 → 多半是批阅高峰把服务压满，等 10 分钟再动手（SLOW_LIMIT）
    """
    DEAD_LIMIT = 2      # 2 次 × 30 秒 = 1 分钟
    SLOW_LIMIT = 20     # 20 次 × 30 秒 = 10 分钟
    TICK = 30

    _write_pid(WATCH_PID_PATH)
    log('=' * 60)
    log('QuickJudge 看护进程启动（每 %d 秒体检一次，端口 %d）' % (TICK, PORT))
    log('=' * 60)

    child = None
    slow = 0
    dead = 0
    try:
        while True:
            if child is not None and child.poll() is not None:
                log('服务进程已退出（返回码 %s），3 秒后重新拉起' % child.returncode)
                child = None
                time.sleep(3)
                continue

            if child is None:
                # 可能是上一次看护留下的实例在跑（重复自启），先认它。
                # 端口在监听就说明有人顶着，不去抢。
                if healthy() or tcp_alive():
                    slow = dead = 0
                    time.sleep(TICK)
                    continue
                child = _spawn_service()
                log('等待服务就绪…')
                if _wait_healthy():
                    log('服务就绪：http://127.0.0.1:%d/' % PORT)
                else:
                    log('警告：180 秒内端口仍无响应，下一轮体检继续处理')
                slow = dead = 0
                time.sleep(TICK)
                continue

            if healthy():
                if slow or dead:
                    log('端口已恢复响应（此前 HTTP 超时 %d 次 / 连不上 %d 次）' % (slow, dead))
                slow = dead = 0
            elif tcp_alive():
                slow += 1
                dead = 0
                log('HTTP 超时（第 %d 次），但端口仍在监听，按「忙」处理，继续观察' % slow)
                if slow >= SLOW_LIMIT:
                    log('已连续 %d 分钟 HTTP 无响应，判定卡死，重启服务' % (slow * TICK // 60))
                    _terminate(child.pid)
                    time.sleep(3)
                    child = None
                    slow = dead = 0
                    continue
            else:
                dead += 1
                slow = 0
                log('连不上端口（第 %d 次）' % dead)
                if dead >= DEAD_LIMIT:
                    log('连续 %d 次连不上端口，判定进程已死，重启服务' % dead)
                    _terminate(child.pid)
                    time.sleep(3)
                    child = None
                    slow = dead = 0
                    continue
            time.sleep(TICK)
    finally:
        _drop_pid(WATCH_PID_PATH)


def main():
    if '--serve' in sys.argv:
        return serve()
    return watch()


if __name__ == '__main__':
    if '--self-test' in sys.argv:
        sys.exit(self_test())
    sys.exit(main())
