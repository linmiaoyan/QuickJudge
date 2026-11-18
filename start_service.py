# -*- coding: utf-8 -*-
"""在后台启动 QuickJudge 服务，并等到端口真的能访问了再报结果。

    python start_service.py

为什么要这个脚本而不是在 .bat 里直接 start：
  1. .bat 里写中文注释会被 cmd 按本地代码页错误解码（实测会把注释当成命令执行），
     所以 .bat 只保留纯 ASCII 的调用，所有中文提示由 Python 打印
     （Python 走宽字符控制台接口，不管 chcp 是多少都不会乱码）；
  2. 启动后需要等几秒（NAPS2 初始化、密钥保险箱加载），
     脚本在这里把「正在启动 → 起来了 → 地址」一次讲清楚，
     而不是让老师对着一个一闪而过的黑窗口猜。
"""
import os
import subprocess
import sys
import time

BASE = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(BASE, 'logs')
LOG_PATH = os.path.join(LOG_DIR, 'service.log')

sys.path.insert(0, BASE)
import stop_service as ss  # 复用同一套探活/pid 判断，避免两处逻辑走偏


def pick_pythonw():
    """挑一个不会弹控制台的解释器。

    当前是 python.exe 时，找同目录的 pythonw.exe；找不到就退回 python.exe
    （服务照样能跑，只是会带一个窗口）。
    """
    exe = sys.executable or 'python'
    if exe.lower().endswith('pythonw.exe'):
        return exe
    alt = os.path.join(os.path.dirname(exe), 'pythonw.exe')
    if os.path.isfile(alt):
        return alt
    return exe


def _spawn(args):
    """拉起服务进程，尽量让它脱离当前的进程树/作业对象。

    DETACHED_PROCESS + CREATE_BREAKAWAY_FROM_JOB 是这里的关键：只要服务仍是
    「谁启动它、谁退出就带走它」的子进程，那"关掉 WorkBuddy / 关掉终端"就会
    把它一起带走——那正是要解决的问题。BREAKAWAY 让子进程跳出作业对象，
    父进程退出后它照常活着。某些环境不允许 breakaway（会报 ACCESS_DENIED），
    所以逐级降级，最后至少保证能起来。
    """
    DETACHED_PROCESS = 0x00000008
    CREATE_BREAKAWAY_FROM_JOB = 0x01000000
    CREATE_NO_WINDOW = 0x08000000
    last = None
    for flags in (DETACHED_PROCESS | CREATE_BREAKAWAY_FROM_JOB,
                  DETACHED_PROCESS,
                  CREATE_NO_WINDOW):
        try:
            return subprocess.Popen(
                args, cwd=BASE, close_fds=True,
                stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                creationflags=flags,
            )
        except OSError as e:
            last = e
    raise last


def main():
    if ss.healthy():
        print('服务已经在运行了，不用重复启动。')
        print('本机访问：http://127.0.0.1:%d/' % ss.PORT)
        return 0

    os.makedirs(LOG_DIR, exist_ok=True)
    exe = pick_pythonw()
    args = [exe, os.path.join(BASE, 'run_service.py'), '--watch']
    print('正在后台启动…（用的是 %s）' % os.path.basename(exe))
    try:
        _spawn(args)
    except OSError as e:
        print('启动失败：%s' % e)
        return 1

    # 最多等 90 秒。正常情况下 5~15 秒就绪，慢的是磁盘和 NAPS2 初始化。
    for i in range(90):
        time.sleep(1)
        if ss.healthy(timeout=2):
            print('已启动，用时约 %d 秒。' % (i + 1))
            print('本机访问：http://127.0.0.1:%d/' % ss.PORT)
            print('运行日志：%s' % LOG_PATH)
            return 0
        if (i + 1) % 10 == 0:
            print('  还在启动中…已等 %d 秒' % (i + 1))

    print('等了 90 秒端口还没响应，启动可能失败了。')
    print('请打开下面的日志看最后几行：')
    print('  %s' % LOG_PATH)
    return 1


if __name__ == '__main__':
    sys.exit(main())
