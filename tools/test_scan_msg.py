# -*- coding: utf-8 -*-
"""只抽取 _humanize_scan_error 做单元验证，不 import 整个 app（避免副作用）。"""
import ast
import io
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
src = io.open(os.path.join(ROOT, 'app.py'), encoding='utf-8').read()
tree = ast.parse(src)
fn = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == '_humanize_scan_error')
mod = ast.Module(body=[fn], type_ignores=[])
ns = {'re': re}
exec(compile(ast.fix_missing_locations(mod), '<extracted>', 'exec'), ns)
humanize = ns['_humanize_scan_error']

cases = [
    # 设备没就绪：返回码 1 + 0 页 —— 必须报「没出图/设备没启动」，不能报「无纸」
    ('Beginning scan...\nStarting scan 1 of 1...\n0 page(s) scanned.', 1, '没有出图'),
    # 真的没放纸：返回码 0 + 0 页 —— 仍报「未检测到原稿」
    ('0 page(s) scanned.', 0, '未检测到原稿'),
    # 正常扫到多页但下游误用（防 10 页被 '0 page' 子串误判）
    ('10 page(s) scanned.', 0, '扫描失败'),
    # 卡纸
    ('Paper jam detected', 1, '卡纸'),
]

fails = 0
for text, rc, expect in cases:
    info = humanize(text, rc)
    title = info['title']
    ok = expect in title
    print(('PASS' if ok else 'FAIL'), '| rc=%s' % rc, '| expect~%s' % expect, '| got:', title)
    if not ok:
        fails += 1

print('\n结果：%d 通过，%d 失败' % (len(cases) - fails, fails))
raise SystemExit(1 if fails else 0)
