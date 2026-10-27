#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""按比例折算已有记录的满分口径（例如 25 分制 → 15 分制）。

为什么需要它
------------
2026-10-26 那批 39 张，老师扫的时候在扫描端填了满分 15，但那个值在临时区被覆盖、
用恢复工具捞回来之后就丢了，整批按方案默认的 25 分制出分。重批要再烧一轮视觉调用，
而分数本身是线性的，按比例折算就够用。

它改哪些字段（一条都不能漏，漏一个界面就会自相矛盾）
------------------------------------------------------
  record.max_score           25 → 15
  record.score_text          '18/25' → '10.8/15'
  grading.max_total          25 → 15
  grading.total              18 → 10.8
  grading.items[*].max_score 25 → 15
  grading.items[*].score     18 → 10.8
  essay.max_score            25 → 15
  essay.score                18 → 10.8
`level`（档次名，如「良好」）不动：它是定性档位，不随满分刻度变。

用法
----
    python tools/rescale_scores.py                     # 只预览
    python tools/rescale_scores.py --class 2614        # 只预览某个班
    python tools/rescale_scores.py --apply             # 真的写
    python tools/rescale_scores.py --from 25 --to 15 --apply

⚠️ 这个脚本会写 data/records/index.json。**先停服务**再 --apply，否则服务进程
内存里的旧索引会在下一次保存时把结果覆盖回去（索引是读-改-写）。
"""
import argparse
import json
import os
import shutil
import sys
import time

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INDEX = os.path.join(HERE, 'data', 'records', 'index.json')


def _round(v):
    """保留一位小数；整数就返回 int，避免出现 15.0/15 这种别扭写法。"""
    v = round(float(v) + 1e-9, 1)
    return int(v) if float(v).is_integer() else v


def _fmt(v):
    return ('%g' % v)


def rescale_record(rec, old_max, new_max):
    """就地折算一条记录，返回改动清单（空列表 = 没动）。"""
    ratio = float(new_max) / float(old_max)
    changes = []

    def hit(path, before, after):
        changes.append((path, before, after))

    if _as_num(rec.get('max_score')) == old_max:
        before = rec['max_score']
        rec['max_score'] = new_max
        hit('max_score', before, new_max)

    st = rec.get('score_text')
    if isinstance(st, str) and '/' in st:
        num, _, den = st.partition('/')
        if _as_num(den) == old_max:
            try:
                rec['score_text'] = '%s/%s' % (_fmt(_round(float(num) * ratio)), _fmt(new_max))
                hit('score_text', st, rec['score_text'])
            except ValueError:
                pass

    g = rec.get('grading')
    if isinstance(g, dict):
        if _as_num(g.get('max_total')) == old_max:
            before = g['max_total']
            g['max_total'] = new_max
            hit('grading.max_total', before, new_max)
            if _as_num(g.get('total')) is not None:
                before = g['total']
                g['total'] = _round(float(g['total']) * ratio)
                hit('grading.total', before, g['total'])
        for i, it in enumerate(g.get('items') or []):
            if not isinstance(it, dict):
                continue
            if _as_num(it.get('max_score')) == old_max:
                before = it['max_score']
                it['max_score'] = new_max
                hit('grading.items[%d].max_score' % i, before, new_max)
                if _as_num(it.get('score')) is not None:
                    before = it['score']
                    it['score'] = _round(float(it['score']) * ratio)
                    hit('grading.items[%d].score' % i, before, it['score'])

    e = rec.get('essay')
    if isinstance(e, dict):
        if _as_num(e.get('max_score')) == old_max:
            before = e['max_score']
            e['max_score'] = new_max
            hit('essay.max_score', before, new_max)
            if _as_num(e.get('score')) is not None:
                before = e['score']
                e['score'] = _round(float(e['score']) * ratio)
                hit('essay.score', before, e['score'])
    return changes


def _as_num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--from', dest='old', type=float, default=25.0, help='原满分（默认 25）')
    ap.add_argument('--to', dest='new', type=float, default=15.0, help='目标满分（默认 15）')
    ap.add_argument('--class', dest='cls', default='', help='只处理某个班（class_name）')
    ap.add_argument('--apply', action='store_true', help='真的写盘（默认只预览）')
    args = ap.parse_args()

    if args.old <= 0 or args.new <= 0:
        print('满分必须是正数')
        return 2
    if not os.path.exists(INDEX):
        print('找不到 %s' % INDEX)
        return 2

    with open(INDEX, encoding='utf-8') as f:
        records = json.load(f)

    touched, skipped = [], 0
    for r in records:
        if args.cls and (r.get('class_name') or '') != args.cls:
            continue
        ch = rescale_record(r, args.old, args.new)
        if ch:
            touched.append((r, ch))
        else:
            skipped += 1

    print('记录总数 %d ｜ 命中 %d 条 ｜ 未动 %d 条'
          % (len(records), len(touched), skipped))
    print('折算比例：%g → %g（×%g），保留一位小数\n' % (args.old, args.new, args.new / args.old))
    for r, ch in touched[:15]:
        print('  %s %s %s' % (r.get('student_id') or '?', r.get('student_name') or '',
                              ' ｜ '.join('%s %s→%s' % (p, b, a) for p, b, a in ch)))
    if len(touched) > 15:
        print('  … 另有 %d 条' % (len(touched) - 15))

    if not args.apply:
        print('\n（预览模式，未写盘。确认无误后加 --apply，并先停掉 5001 服务）')
        return 0

    if sys.stdin.isatty() is False:
        pass
    bak = INDEX + '.bak-rescale-' + time.strftime('%Y%m%d-%H%M%S')
    shutil.copy2(INDEX, bak)
    with open(INDEX, 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    print('\n已写盘 %d 条，原索引备份为 %s' % (len(touched), os.path.basename(bak)))
    return 0


if __name__ == '__main__':
    sys.exit(main())
