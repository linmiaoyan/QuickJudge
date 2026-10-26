# -*- coding: utf-8 -*-
"""修复「已批阅的卷子还挂在临时区」的残留数据。

背景（2026-10-26 那批 39 张）：
  流水线跑到「归档」阶段时进程被重启，31 条记录已经写进档案库，
  但临时区的清理是整批跑完才做的最后一步 → 索引里 39 条全留着，
  图片却早被 _commit_staging_item 搬进 records/images 了。
  界面上看起来就是"明明批阅了，临时区还有一堆点开空白"。

本脚本做三件事（都不删任何档案记录、不动任何图片）：
  1. 把「进程已死但状态还 running」的批次标成中断，扫描端不再转圈；
  2. 从临时区索引里清掉图片已不存在的条目（它们已经入库）；
  3. 打印重复档案的清单，供人工确认后再处理（不自动删）。

用法：
  python tools/repair_staging.py            # 只看，不改
  python tools/repair_staging.py --apply    # 真正执行（先自动备份索引）
"""
import collections
import json
import os
import shutil
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

APPLY = '--apply' in sys.argv

DATA = os.path.join(ROOT, 'data')
STAGING_IDX = os.path.join(DATA, 'staging', 'index.json')
STAGING_IMG = os.path.join(DATA, 'staging', 'images')
RECORDS_IDX = os.path.join(DATA, 'records', 'index.json')
BATCHES = os.path.join(DATA, 'batches')


def load(p, default):
    try:
        with open(p, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return default


def backup(p):
    if not os.path.exists(p):
        return None
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    dst = p + '.bak-repair-' + ts
    shutil.copy2(p, dst)
    return dst


def main():
    print('=' * 62)
    print('临时区残留修复' + ('' if APPLY else '（演练模式，不会改动任何文件）'))
    print('=' * 62)

    staging = load(STAGING_IDX, [])
    files = set(os.listdir(STAGING_IMG)) if os.path.isdir(STAGING_IMG) else set()
    print('临时区索引 %d 条 / images 目录 %d 个文件' % (len(staging), len(files)))

    def has_img(it):
        iid = str(it.get('id') or '')
        return any(f.startswith(iid) for f in files)

    ghosts = [it for it in staging if not has_img(it) or str(it.get('status')) == 'committed']
    keep = [it for it in staging if it not in ghosts]
    print('图片已不存在（=已入库，应当移除）: %d 条' % len(ghosts))
    print('仍有图片（=确实还没入库）:        %d 条' % len(keep))
    if keep:
        print('  保留的学号:', sorted({str(x.get('student_id')) for x in keep}))

    # ---- 批次状态 ----
    stale = []
    for name in sorted(os.listdir(BATCHES)) if os.path.isdir(BATCHES) else []:
        if not name.endswith('.json'):
            continue
        p = os.path.join(BATCHES, name)
        b = load(p, None)
        if not isinstance(b, dict):
            continue
        job = b.get('job') or {}
        if job.get('status') not in ('running', 'queued'):
            continue
        runner = b.get('runner') or {}
        # 脚本是独立进程，任何 running 批次都不可能是「本进程」在跑
        stale.append((name[:-5], b, runner, job))

    print()
    print('卡在 running/queued 的批次: %d 个' % len(stale))
    for bid, b, runner, job in stale:
        left = sum(1 for i in (b.get('item_ids') or []) if any(
            str(x.get('id')) == str(i) for x in keep))
        print('  %s  创建于 %s  stage=%s 已入库 %d 份  临时区剩余 %d 张'
              % (bid, b.get('created_at'), job.get('stage'),
                 len(b.get('record_ids') or []), left))
        print('     runner=%s  最后更新=%s' % (runner or '（无 runner 字段，老批次）',
                                             job.get('updated_at')))

    # ---- 重复档案（只报告）----
    recs = load(RECORDS_IDX, [])
    dup = {sid: n for sid, n in collections.Counter(
        str(r.get('student_id')) for r in recs).items() if n > 1}
    print()
    print('档案总数 %d 条；同一学号多条记录: %d 个学生' % (len(recs), len(dup)))
    for sid, n in sorted(dup.items(), key=lambda kv: -kv[1]):
        rows = [r for r in recs if str(r.get('student_id')) == sid]
        print('  学号 %s 有 %d 条：' % (sid, n))
        for r in rows:
            print('     %s  %s  %s  %s' % (r.get('id', '')[:8], r.get('created_at'),
                                           r.get('score_text') or '-', r.get('batch_id') or '-'))
    print('  （重复记录本脚本不会自动删除，需要你确认后单独处理）')

    if not APPLY:
        print()
        print('演练结束。确认无误后加 --apply 执行。')
        return

    # ---- 执行 ----
    print()
    for p in (STAGING_IDX, RECORDS_IDX):
        dst = backup(p)
        if dst:
            print('已备份:', os.path.basename(dst))

    if ghosts:
        with open(STAGING_IDX, 'w', encoding='utf-8') as f:
            json.dump(keep, f, ensure_ascii=False, indent=2)
        print('临时区：移除 %d 条幽灵条目，剩 %d 条' % (len(ghosts), len(keep)))

    for bid, b, runner, job in stale:
        alive_ids = {str(x.get('id')) for x in keep}
        left = sum(1 for i in (b.get('item_ids') or []) if str(i) in alive_ids)
        done_n = len(b.get('record_ids') or [])
        msg = '上一次归档被中断（服务重启或进程退出）'
        if done_n:
            msg += '，已入库的 %d 份没丢' % done_n
        if left:
            msg += '；还有 %d 张留在临时区，点「重新批阅」可续跑（已入库的不会重复建档）' % left
        job.update(status='error', stage='error', message=msg,
                   updated_at=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        b['job'] = job
        b['interrupted'] = True
        b['runner'] = {'pid': 0, 'boot': 'repaired'}
        p = os.path.join(BATCHES, bid + '.json')
        with open(p + '.tmp', 'w', encoding='utf-8') as f:
            json.dump(b, f, ensure_ascii=False, indent=2)
        os.replace(p + '.tmp', p)
        print('批次 %s 已标记为中断' % bid)

    print()
    print('完成。')


if __name__ == '__main__':
    main()
