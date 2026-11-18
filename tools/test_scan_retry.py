# -*- coding: utf-8 -*-
"""验证 _run_scanner 的重试策略（不碰真实扫描仪）。

模拟场景：
  A. 第 1 次返回码 1 + 0 页（设备没就绪），第 2 次正常出 1 页 -> 应重试并成功
  B. 第 1 次返回码 0 + 0 页（确实没放纸）-> 不重试，报「未检测到原稿」
  C. 第 1 次卡纸（jam）-> 不重试，报卡纸
"""
import json
import os
import shutil
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import app  # noqa: E402

# 避免真的去 taskkill 用户正在用的扫描进程
app._cleanup_stale_scanner_processes = lambda: None
# 让 NAPS2 路径校验通过（用任意存在的文件冒充）
app.get_naps2_path = lambda: os.path.join(ROOT, 'app.py')
# 别去枚举真实扫描设备（否则会尝试执行 app.py --listdevices）
app._list_scan_devices = lambda *a, **k: (None, [])


def _run_case(name, results, expect_ok, expect_pages, expect_calls, expect_title_kw=''):
    out_dir = tempfile.mkdtemp(prefix='qj_scan_')
    calls = {'n': 0}

    def fake_watchdog(cmd, d, f, log_prefix=''):
        i = calls['n']
        calls['n'] += 1
        rc, text, diag, npages = results[min(i, len(results) - 1)]
        if npages:
            with open(os.path.join(d, f), 'wb') as fh:
                fh.write(b'\x89PNG fake')
        return rc, text, diag, npages

    app._scan_with_watchdog = fake_watchdog
    try:
        ok, msg, pages, meta = app._run_scanner(out_dir, 'scan1.png')
    finally:
        shutil.rmtree(out_dir, ignore_errors=True)

    okk = (ok == expect_ok and len(pages) == expect_pages and calls['n'] == expect_calls)
    if expect_title_kw:
        okk = okk and (expect_title_kw in (msg or ''))
    print(('PASS' if okk else 'FAIL'), '|', name,
          '| ok=%s pages=%d calls=%d' % (ok, len(pages), calls['n']),
          '| msg[:40]=%r' % (msg or '')[:40])
    return okk


fails = 0
fails += not _run_case(
    'A 返回码1+0页 -> 重试后成功',
    [(1, 'Beginning scan...\nStarting scan 1 of 1...\n0 page(s) scanned.', 'ok', 0),
     (0, '1 page(s) scanned.\nFinished saving', 'ok', 1)],
    expect_ok=True, expect_pages=1, expect_calls=2)

fails += not _run_case(
    'B 返回码0+0页 -> 不重试，报未检测到原稿',
    [(0, '0 page(s) scanned.', 'ok', 0)],
    expect_ok=False, expect_pages=0, expect_calls=1, expect_title_kw='未检测到原稿')

fails += not _run_case(
    'C 卡纸 -> 不重试，报卡纸',
    [(1, 'Paper jam detected', 'jam', 0)],
    expect_ok=False, expect_pages=0, expect_calls=1, expect_title_kw='卡纸')

print('\n结果：%s' % ('全部通过' if not fails else '%d 项失败' % fails))
raise SystemExit(1 if fails else 0)
