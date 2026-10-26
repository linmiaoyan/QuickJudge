# -*- coding: utf-8 -*-
"""验证卡纸自救相关的三块逻辑：令牌锁、卡纸识别、僵尸锁回收。

运行： python tools/test_scan_lock.py
"""
import os
import sys
import time
import threading

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app  # noqa: E402

fail = []


def check(name, cond, extra=''):
    print(('  [OK] ' if cond else '  [FAIL] ') + name + (('  ' + extra) if extra else ''))
    if not cond:
        fail.append(name)


print('1) 卡纸关键字识别')
for text, want in [
    ('Paper jam detected in ADF', True),
    ('请整理纸，清除卡纸后继续', True),
    ('Scanner cover open', True),
    ('Beginning scan...', False),
    ('1 page(s) scanned.', False),
    ('', False),
]:
    check('识别 %r -> %s' % (text[:28], want), app._looks_like_jam(text) is want)

print('2) 令牌锁：被强制回收后，老持有者的 release 不能破坏新持有者')
lock = app._ScanLock()
t1 = lock.acquire(who='A')
check('A 拿到锁', bool(t1))
lock.steal()
t2 = lock.acquire(who='B')
check('B 接管成功', bool(t2) and t2 != t1)
lock.release(t1)                       # A 迟到的 finally
check('A 的迟到 release 不影响 B', lock.locked() and lock.holder()[0] == 'B')
lock.release(t2)
check('B 正常释放后为空闲', not lock.locked())

print('3) _acquire_scan_lock：未超时 -> 拒绝；超时 -> 强制回收')
old_max = app.SCAN_LOCK_MAX_HOLD
try:
    app.SCAN_LOCK_MAX_HOLD = 1
    tok = app._SCAN_LOCK.acquire(who='zombie')
    app._SCAN_LOCK._since = time.time() - 0.2     # 刚持有，未超时
    got, held, reclaimed = app._acquire_scan_lock('t')
    check('未超时直接拒绝', got == 0 and reclaimed is False, 'held=%.1f' % held)

    app._SCAN_LOCK._since = time.time() - 5       # 模拟卡死 5 秒
    t0 = time.time()
    got, held, reclaimed = app._acquire_scan_lock('t')
    cost = time.time() - t0
    check('超时后被强制回收并可继续', bool(got) and reclaimed is True,
          'held=%.1f 耗时=%.1fs' % (held, cost))
    check('回收耗时可控（<20s）', cost < 20, '%.1fs' % cost)
    app._SCAN_LOCK.release(got)
finally:
    app.SCAN_LOCK_MAX_HOLD = old_max
    # 收尾，避免污染后续
    try:
        app._SCAN_LOCK.steal()
    except Exception:
        pass

print('4) 并发抢锁：同一时刻只有一个成功')
app._SCAN_LOCK.steal()
results = []
barrier = threading.Barrier(5)


def grab(i):
    barrier.wait()
    t = app._SCAN_LOCK.acquire(who='th%d' % i)
    results.append((i, t))
    if t:
        time.sleep(0.2)
        app._SCAN_LOCK.release(t)


ths = [threading.Thread(target=grab, args=(i,)) for i in range(5)]
[t.start() for t in ths]
[t.join() for t in ths]
odd = [i for i, t in results if t]
check('5 个线程只有 1 个拿到锁', len(odd) == 1, str(results))

print()
if fail:
    print('失败 %d 项：%s' % (len(fail), '; '.join(fail)))
    sys.exit(1)
print('全部通过')
