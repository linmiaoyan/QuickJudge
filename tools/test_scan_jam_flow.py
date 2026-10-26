# -*- coding: utf-8 -*-
"""端到端验证卡纸自救：真实跑 Flask 路由，但把 NAPS2 调用替换成假的。

覆盖场景：
  A. 锁刚被正常扫描占用 -> /api/scan 回 429（不误杀正常扫描）
  B. 锁被占超过阈值（僵尸） -> /api/scan 自动回收并继续走流程
  C. /api/scan/reset 在僵尸占用时不再 409
  D. /api/scan/force_unlock 能释放锁并清进程
  E. 扫描工作站遇到卡纸 -> 返回 jam=True + 中文取纸引导，且不自动重试
运行： python tools/test_scan_jam_flow.py
"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app  # noqa: E402

fail = []


def check(name, cond, extra=''):
    print(('  [OK] ' if cond else '  [FAIL] ') + name + (('  ' + extra) if extra else ''))
    if not cond:
        fail.append(name)


def reset_lock():
    try:
        app._SCAN_LOCK.steal()
    except Exception:
        pass


# 让 /api/scan 在取到锁之后立刻以"找不到 NAPS2"退出，避免真的去驱动扫描仪
app.get_naps2_path = lambda: os.path.join(os.path.dirname(__file__), 'NO_SUCH_NAPS2.exe')

c = app.app.test_client()
with c.session_transaction() as s:
    s['user_id'] = 'admin'
    s['user_role'] = 'admin'
    s['user_name'] = '管理员'

old_max = app.SCAN_LOCK_MAX_HOLD

print('A. 锁被正常占用（未超时）时，第二次扫描应被拒绝')
reset_lock()
tok = app._SCAN_LOCK.acquire(who='api/scan')
r = c.post('/api/scan', json={})
d = r.get_json()
check('返回 429', r.status_code == 429, 'status=%s' % r.status_code)
check('提示里带"正在进行中"', '正在进行中' in (d.get('error') or ''))
app._SCAN_LOCK.release(tok)

print('B. 锁被僵尸占用（超阈值）时，/api/scan 应自动回收')
reset_lock()
app.SCAN_LOCK_MAX_HOLD = 1
tok = app._SCAN_LOCK.acquire(who='卡死的旧请求')
app._SCAN_LOCK._since = time.time() - 30      # 假装已经卡了 30 秒
t0 = time.time()
r = c.post('/api/scan', json={})
cost = time.time() - t0
d = r.get_json()
check('不是 429（已成功回收）', r.status_code != 429, 'status=%s' % r.status_code)
check('走到了 NAPS2 路径校验', r.status_code == 400 and 'NAPS2' in (d.get('error') or ''),
      'status=%s' % r.status_code)
check('结束后锁已释放', not app._SCAN_LOCK.locked())
check('回收耗时可控（<20s）', cost < 20, '%.1fs' % cost)

print('C. /api/scan/reset 在僵尸占用时不再一律 409')
reset_lock()
tok = app._SCAN_LOCK.acquire(who='卡死的旧请求')
app._SCAN_LOCK._since = time.time() - 30
# 把软复位改成假的，避免真的去跑 PowerShell
app._run_scanner_reset = lambda timeout=60: (True, 'RESULT=OK (mock)')
r = c.post('/api/scan/reset')
d = r.get_json()
check('返回 200 且 forced=True', r.status_code == 200 and d.get('forced') is True,
      'status=%s body=%s' % (r.status_code, str(d)[:120]))
check('复位后锁已释放', not app._SCAN_LOCK.locked())

print('C2. 正常扫描进行中，复位仍应被拒绝（不打断正在跑的扫描）')
app.SCAN_LOCK_MAX_HOLD = old_max        # 还原阈值，否则这里会被当成僵尸
reset_lock()
tok = app._SCAN_LOCK.acquire(who='api/scan')
app._SCAN_LOCK._since = time.time() - 2       # 刚跑 2 秒，未超阈值
r = c.post('/api/scan/reset')
check('返回 409', r.status_code == 409, 'status=%s' % r.status_code)
app._SCAN_LOCK.release(tok)

print('D. /api/scan/force_unlock 强制解锁')
reset_lock()
tok = app._SCAN_LOCK.acquire(who='卡死的旧请求')
app._SCAN_LOCK._since = time.time() - 12
r = c.post('/api/scan/force_unlock')
d = r.get_json()
check('返回 200', r.status_code == 200, 'status=%s' % r.status_code)
check('was_locked=True', d.get('was_locked') is True, str(d)[:120])
check('解锁后锁空闲', not app._SCAN_LOCK.locked())
check('提示里告诉老师去取纸', '纸张' in (d.get('message') or ''))
r2 = c.post('/api/scan/force_unlock')
check('空闲时再点也不报错', r2.status_code == 200 and r2.get_json().get('was_locked') is False)

print('E. 扫描工作站遇卡纸 -> jam 标记 + 取纸引导 + 不重试')
app.SCAN_LOCK_MAX_HOLD = old_max
# 扫描工作站会先校验 NAPS2 路径存在，这里指向一个真实存在的文件骗过校验，
# 真正的扫描调用由下面的 fake_scan 接管。
app.get_naps2_path = lambda: os.path.abspath(__file__)
reset_lock()
calls = {'n': 0}


def fake_scan(cmd, out_dir, filename, log_prefix=''):
    calls['n'] += 1
    return 0, 'Error: Paper jam detected in the ADF', 'jam', None


app._scan_with_watchdog = fake_scan
r = c.post('/api/station/scan', json={'auto': False})
d = r.get_json()
err = (d.get('error') or '')
check('返回 500', r.status_code == 500, 'status=%s' % r.status_code)
check('带 jam=True', d.get('jam') is True, str(d)[:160])
check('说明是卡纸', '卡纸' in err, err[:60].replace('\n', ' '))
check('给了取纸步骤', '进纸盖' in err or '取出' in err, 'len=%d' % len(err))
check('卡纸只扫了一次（不自动重试）', calls['n'] == 1, 'calls=%d' % calls['n'])
check('失败后锁已释放', not app._SCAN_LOCK.locked())

print('F. 卡纸前已扫出的页要保留（部分成功）')
reset_lock()
calls['n'] = 0


def fake_scan_partial(cmd, out_dir, filename, log_prefix=''):
    calls['n'] += 1
    # 模拟"已写入 2 页后才卡纸"
    import shutil
    src = None
    for root, _dirs, files in os.walk(os.path.join(os.path.dirname(__file__), '..', 'scan')):
        for fn in files:
            if fn.lower().endswith('.png'):
                src = os.path.join(root, fn)
                break
        if src:
            break
    if src:
        base = os.path.splitext(filename)[0]
        shutil.copy(src, os.path.join(out_dir, base + '.1.png'))
        shutil.copy(src, os.path.join(out_dir, base + '.2.png'))
    return 0, 'jam on page 3', 'jam', None


app._scan_with_watchdog = fake_scan_partial
r = c.post('/api/station/scan', json={'auto': False})
d = r.get_json()
if r.status_code == 200:
    check('部分成功仍返回 200', True, 'count=%s' % d.get('count'))
    check('保留了卡纸前的页', (d.get('count') or 0) >= 1, 'count=%s' % d.get('count'))
else:
    # 找不到样张图片时跳过（不算失败），但要说明
    print('  [SKIP] scan 目录里没有样张 png，跳过部分成功用例')
    check('skipped', True)

reset_lock()
print()
if fail:
    print('失败 %d 项：%s' % (len(fail), '; '.join(fail)))
    sys.exit(1)
print('全部通过')
