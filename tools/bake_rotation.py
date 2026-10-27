#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把记录上的 rotation 烘进图片本身，让「预览 / 看原图 / 导出 / 重新批阅」看到同一张正立图。

背景
----
`_commit_staging_item()` 归档时其实**已经把 rotation 和 crop 烘进图片**了
（`_transform_image(img, rotation, crop).save(dest_path)`），所以正常入库的记录
`rotation` 应该永远是 0。2026-10-26 那批是用恢复工具从孤儿图片重建的，跳过了定向那一步，
39 条全是倒图；后来只把 `rotation` 标成 180——这只修好了走 `?processed=1` 的预览，
「看原图」（`/image` 原样返回）、导出、重新批阅拿到的还是倒的字节。

这个脚本把角度真正烧进像素，然后把 `rotation` 归零：
  · 预览（processed=1）→ 变成无操作，仍是正图
  · 看原图（/image）   → 正图
  · 导出 / 报告 / 重新批阅 → 正图

安全措施
--------
  · 只处理 rotation ∈ {90,180,270} 且**没有 regions/crop** 的记录（有区域框的话，
    坐标是在旋转前的画面上量的，直接烧角度会错位，这种记录跳过并提示）。
  · 先整批备份图片到 temp/backup-<日期>/records-images/，再改。
  · 图片用 .tmp + os.replace 原子替换，中途失败不会留下半张图。
  · 备份索引为 index.json.bak-bakerot-<时间>。
  · **先停 5001 服务再 --apply**（索引是读-改-写）。

用法
----
    python tools/bake_rotation.py            # 预览
    python tools/bake_rotation.py --apply    # 真的改
"""
import argparse
import json
import os
import shutil
import sys
import time

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, HERE)

INDEX = os.path.join(HERE, 'data', 'records', 'index.json')
IMAGES = os.path.join(HERE, 'data', 'records', 'images')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--apply', action='store_true', help='真的改（默认只预览）')
    args = ap.parse_args()

    import app as A          # 复用 _open_image / _transform_image，口径和线上一致
    from PIL import Image

    with open(INDEX, encoding='utf-8') as f:
        records = json.load(f)

    todo, skipped = [], []
    for r in records:
        try:
            rot = int(r.get('rotation') or 0) % 360
        except (TypeError, ValueError):
            rot = 0
        if rot not in (90, 180, 270):
            continue
        if r.get('crop') or (r.get('regions') or []):
            skipped.append((r, rot, '有 crop/regions，角度烘进去会让区域框错位'))
            continue
        path = A._records_image_path(r)
        if not path or not os.path.exists(path):
            skipped.append((r, rot, '图片文件缺失'))
            continue
        todo.append((r, rot, path))

    print('需要烘角度的记录 %d 条，跳过 %d 条' % (len(todo), len(skipped)))
    for r, rot, why in skipped:
        print('  ⚠ %s %s rot=%d：%s' % (r.get('student_id') or '?', r.get('student_name') or '', rot, why))
    if not todo:
        print('没有需要处理的记录')
        return 0

    for r, rot, path in todo[:10]:
        print('  %s %s rot=%d°  %s' % (r.get('student_id') or '?', r.get('student_name') or '',
                                       rot, os.path.basename(path)))
    if len(todo) > 10:
        print('  … 另有 %d 条' % (len(todo) - 10))

    if not args.apply:
        print('\n（预览模式，未改动。确认后加 --apply，并先停掉 5001 服务）')
        return 0

    stamp = time.strftime('%Y%m%d-%H%M%S')
    bak_dir = os.path.join(HERE, 'temp', 'backup-' + time.strftime('%Y%m%d'), 'records-images')
    os.makedirs(bak_dir, exist_ok=True)
    ok_n, fail = 0, []
    for r, rot, path in todo:
        name = os.path.basename(path)
        try:
            if not os.path.exists(os.path.join(bak_dir, name)):
                shutil.copy2(path, os.path.join(bak_dir, name))
            img = A._open_image(path)
            out = A._transform_image(img, rotation=rot, crop=None)
            tmp = path + '.tmp'
            out.save(tmp, format=os.path.splitext(path)[1].lstrip('.').upper() or 'PNG')
            os.replace(tmp, path)
            r['rotation'] = 0
            # 留痕：以后排查「这张为什么看不出转过」时有个依据
            r['rotation_baked'] = rot
            r['rotation_baked_at'] = time.strftime('%Y-%m-%d %H:%M:%S')
            ok_n += 1
        except Exception as e:
            fail.append((r.get('student_id'), name, str(e)))

    shutil.copy2(INDEX, INDEX + '.bak-bakerot-' + stamp)
    with open(INDEX, 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

    print('\n已烘进图片 %d 条，失败 %d 条' % (ok_n, len(fail)))
    for sid, name, e in fail:
        print('  ✗ %s %s：%s' % (sid, name, e))
    print('图片备份：%s' % bak_dir)
    print('索引备份：%s' % os.path.basename(INDEX + '.bak-bakerot-' + stamp))
    return 0


if __name__ == '__main__':
    sys.exit(main())
