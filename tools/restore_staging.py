# -*- coding: utf-8 -*-
"""临时区「图片还在、索引却空了」时，把条目从 images/ 重建回来。

适用场景（2026-10-26 真实事故）：扫描端扫了几十张，图片落在 data/staging/images/，
但 staging/index.json 被另一个进程用过期快照覆盖成空列表或少数几条，
于是扫描端一张都看不到——图片没丢，只是没被索引。

原则：**只读图片，不删、不移动、不改名任何文件**；写索引前先备份当前索引。

用法：
    python tools/restore_staging.py             # 预览要恢复什么
    python tools/restore_staging.py --apply     # 真正写回索引
    python tools/restore_staging.py --min-kb 100   # 调整「多大才算真卷面」的门槛

注意：恢复出来的条目只有图片和体检结果，没有学号/班级（那些信息原本也没算过，要重新识别）。
如果它们本该被批阅，恢复后在扫描端对这些条目重新走一遍批阅流程即可。
"""
import argparse
import glob
import os
import shutil
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import app as A  # noqa: E402


def ts(path):
    return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(os.path.getmtime(path)))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--apply', action='store_true', help='真正写回索引（默认只预览）')
    ap.add_argument('--min-kb', type=int, default=100, help='小于这个体积的图视为测试图，跳过')
    args = ap.parse_args()
    min_bytes = args.min_kb * 1024

    existing = A.load_staging()
    indexed = {os.path.basename((it.get('image') or '')) for it in existing}
    files = sorted(glob.glob(os.path.join(A.STAGING_IMAGES_DIR, '*')), key=os.path.getmtime)
    orphan = [f for f in files
              if os.path.getsize(f) >= min_bytes and os.path.basename(f) not in indexed]

    print('索引现有 %d 条；图片目录 %d 个文件；待恢复 %d 张'
          % (len(existing), len(files), len(orphan)))
    if not orphan:
        print('没有需要恢复的图片。')
        return 0
    print('时间范围 %s ~ %s' % (ts(orphan[0]), ts(orphan[-1])))

    items, skipped = [], []
    for i, f in enumerate(orphan, 1):
        stem = os.path.splitext(os.path.basename(f))[0]
        if len(stem) != 32:                     # 正常条目 id 是 32 位 uuid4 hex
            skipped.append(os.path.basename(f))
            continue
        items.append({
            'id': stem,
            'image': 'images/' + os.path.basename(f),
            'original_name': 'restored-%02d%s' % (i, os.path.splitext(f)[1].lower()),
            'source': 'scan',
            'created_at': ts(f),
            'student_id': '', 'student_id_guess': '', 'class_name': '', 'task_id': '',
            'subject': '', 'crop': None, 'regions': [], 'rotation': 0, 'ocr_text': '',
            'note': '从图片恢复（索引曾丢失）',
            'status': 'staged',
        })
        q = A.image_quality_metrics(f)
        if q:
            items[-1]['quality'] = q

    print('可恢复 %d 条，跳过文件名异常 %d 条' % (len(items), len(skipped)))
    if not args.apply:
        for it in items[:3] + items[-1:]:
            print('  %s  %s  %s' % (it['id'][:8], it['original_name'], it['created_at']))
        print('\n（预览模式，未写入；加 --apply 生效）')
        return 0

    if os.path.exists(A.STAGING_INDEX_FILE):
        bak = os.path.join(A.STAGING_DIR, 'index.json.bak-%s' % time.strftime('%Y%m%d_%H%M%S'))
        shutil.copy2(A.STAGING_INDEX_FILE, bak)
        print('已备份原索引 ->', bak)

    A.save_staging(existing + items)
    print('已写入 %d 条（原有 %d + 恢复 %d）' % (len(existing) + len(items), len(existing), len(items)))
    print('复核：', len(A.load_staging()), '条')
    return 0


if __name__ == '__main__':
    sys.exit(main())
