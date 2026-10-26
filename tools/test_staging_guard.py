# -*- coding: utf-8 -*-
"""验证「临时区幽灵条目 / 僵尸批次 / 重复归档」这三处修复。

运行： python tools/test_staging_guard.py
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app  # noqa: E402

fail = []


def check(name, cond, extra=''):
    print(('  [OK] ' if cond else '  [FAIL] ') + name + (('  ' + extra) if extra else ''))
    if not cond:
        fail.append(name)


c = app.app.test_client()
with c.session_transaction() as s:
    s['user_id'] = 'admin'
    s['user_role'] = 'admin'
    s['user_name'] = '管理员'


def mk_batch(bid, boot, status='running', item_ids=None, records=None):
    b = {
        'id': bid, 'created_at': app._now_str(), 'subject': 'english_essay',
        'max_score': 25, 'task_id': '', 'note': '',
        'item_ids': item_ids or [], 'topic': {}, 'corpus': {},
        'awards': {'awards': []}, 'record_ids': records or [],
        'job': {'status': status, 'stage': 'committing', 'total': 1, 'done': 1,
                'message': '正在归档到用户端…', 'updated_at': app._now_str()},
        'runner': {'pid': 1, 'boot': boot},
    }
    app._save_batch(b)
    return b


print('1. 临时区列表不再返回"图片已不在"的幽灵条目')
r = c.get('/api/station/staging')
d = r.get_json()
idx = app.load_staging()
noimg = [x for x in idx if not app._staging_image_path(x)]
check('接口 200', r.status_code == 200, 'status=%s' % r.status_code)
check('返回的条目全部有图', all(x.get('has_image') for x in (d.get('items') or [])))
check('索引里已无缺图条目', len(noimg) == 0, '缺图 %d 条' % len(noimg))
print('   当前临时区 %d 条（%s）' % (d.get('count'), ', '.join(
    str(x.get('student_id')) for x in (d.get('items') or []))))

print('2. 别的进程留下的 running 批次，读一次就自动标成中断')
bid = 'testguard0000001'
mk_batch(bid, boot='some-old-boot')
b2 = app._load_batch(bid)
job = (b2 or {}).get('job') or {}
check('状态已变 error', job.get('status') == 'error', 'status=%s' % job.get('status'))
check('消息说明被中断', '中断' in (job.get('message') or ''), job.get('message', '')[:50])
check('打了 interrupted 标记', b2.get('interrupted') is True)
os.remove(os.path.join(app.BATCHES_DIR, bid + '.json'))
app._BATCH_MEM.pop(os.path.join(app.BATCHES_DIR, bid + '.json'), None)

print('3. 本进程正在跑的批次不能被误判为僵尸')
bid2 = 'testguard0000002'
mk_batch(bid2, boot=app._BOOT_ID)
b3 = app._load_batch(bid2)
check('仍保持 running', ((b3 or {}).get('job') or {}).get('status') == 'running',
      str(((b3 or {}).get('job') or {}).get('status')))

print('4. 同一批次重复点「重新批阅」要被挡住（409），不能起第二个线程')
r = c.post('/api/station/jobs/%s/retry' % bid2)
check('返回 409', r.status_code == 409, 'status=%s' % r.status_code)
check('提示说明会重复归档', '重复' in (r.get_json().get('error') or ''),
      (r.get_json().get('error') or '')[:40])
os.remove(os.path.join(app.BATCHES_DIR, bid2 + '.json'))
app._BATCH_MEM.pop(os.path.join(app.BATCHES_DIR, bid2 + '.json'), None)

print('5. 同一批图片不会起两个批次')
bid3 = 'testguard0000003'
mk_batch(bid3, boot=app._BOOT_ID, item_ids=['aaa111'])
found = app._find_running_batch(['aaa111', 'bbb222'])
check('能查出重叠的进行中批次', bool(found) and found.get('id') == bid3,
      str((found or {}).get('id')))
check('不重叠时返回空', app._find_running_batch(['zzz999']) is None)
os.remove(os.path.join(app.BATCHES_DIR, bid3 + '.json'))
app._BATCH_MEM.pop(os.path.join(app.BATCHES_DIR, bid3 + '.json'), None)

print('6. 归档幂等：已有记录对应的临时区条目会被识别出来')
# 造一条临时假记录指向不存在的条目，只看映射函数本身
old = app.load_records()
try:
    fake = dict(old[0]) if old else {}
    if fake:
        fake['id'] = 'testrec00000001'
        fake['staging_item'] = 'item-abc'
        app.save_records(old + [fake])
        m = app._records_by_staging_item(['item-abc', 'item-xyz'])
        check('能按临时区条目找到已入库记录', m.get('item-abc') == 'testrec00000001', str(m))
        check('没入库的条目查不到', 'item-xyz' not in m)
finally:
    app.save_records(old)

print('7. 幽灵条目清理函数幂等（没有幽灵时不动索引）')
before = app.load_staging()
n = app._purge_ghost_staging_items()
after = app.load_staging()
check('无幽灵时返回 0', n == 0, 'n=%d' % n)
check('索引未被改动', len(before) == len(after))

print()
if fail:
    print('失败 %d 项：%s' % (len(fail), '; '.join(fail)))
    sys.exit(1)
print('全部通过')
