# -*- coding: utf-8 -*-
"""把「学生姓名与学号对照表」导入 QuickJudge，并取代现有学生名单。

为什么要有这个脚本：
  项目自带的名单导入只认 CSV（姓名,账号），而学校下发的通常是 xlsx。
  本脚本支持 xlsx / csv 两种来源，且**以学号为准重建整套名单**：
  账号表、班级名单、班级目录三者一起对齐，避免出现「账号在、名单不在」的脏状态。

数据口径（与学校现行惯例一致）：
  · 学号为 6 位：前 2 位年级 + 第 3-4 位班号 + 第 5-6 位座号（260101 -> 26 级 1 班 01 号）
  · 班级编码 = 学号前 4 位（260101 -> 2601）
  · 学生密码 = 学号（沿用原设计，学生端可直接登录）
  · **高一学生卷面上只写学号后四位**（0101），需靠班级号回查完整学号

用法：
    python tools/import_roster.py "C:/path/学生姓名与学号对照表.xlsx"
    python tools/import_roster.py 名单.csv --keep-students

默认行为（不加 --keep-students）：
    · 学生账号整体替换为名单内容（非学生账号如 admin/teacher 原样保留）
    · data/students/<班级>.json 按名单重建
    · data/classes/<班级>/ 目录按名单重建
    · 名单里已不存在的旧班级，其名单文件与空目录会被清掉（不递归删业务文件）
执行前会自动备份 users.json。
"""
import argparse
import csv
import io
import json
import os
import re
import shutil
import sys
import time
import zipfile
from xml.etree import ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, 'data')
USERS_FILE = os.path.join(DATA_DIR, 'config', 'users.json')
STUDENTS_DIR = os.path.join(DATA_DIR, 'students')
CLASSES_DIR = os.path.join(DATA_DIR, 'classes')

_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
_NAME_KEYS = ('学生姓名', '姓名', '名字', 'name')
_ID_KEYS = ('学号', '考号', '账号', 'student_id', 'id')


def _cell_col(ref):
    """把单元格坐标（如 C12）换算成 0 基列号。"""
    m = re.match(r'([A-Z]+)', str(ref or ''))
    if not m:
        return None
    n = 0
    for ch in m.group(1):
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def read_xlsx(path):
    """用标准库读 xlsx（避免为一个导入动作给运行环境加 openpyxl 依赖）。"""
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall(_NS + 'si'):
                shared.append(''.join(t.text or '' for t in si.iter(_NS + 't')))
        sheets = sorted(n for n in z.namelist()
                        if re.match(r'xl/worksheets/sheet\d+\.xml$', n))
        if not sheets:
            raise RuntimeError('xlsx 里找不到工作表')
        rows = []
        root = ET.fromstring(z.read(sheets[0]))
        for row in root.iter(_NS + 'row'):
            vals, pos = {}, 0
            for c in row.findall(_NS + 'c'):
                t, v, isel = c.get('t'), c.find(_NS + 'v'), c.find(_NS + 'is')
                if t == 's' and v is not None:
                    val = shared[int(v.text)]
                elif isel is not None:
                    val = ''.join(x.text or '' for x in isel.iter(_NS + 't'))
                elif v is not None:
                    val = v.text
                else:
                    val = ''
                idx = _cell_col(c.get('ref'))
                idx = pos if idx is None else idx
                vals[idx] = val
                pos = idx + 1
            if vals:
                rows.append([vals.get(i, '') for i in range(max(vals) + 1)])
        return rows


def read_csv(path):
    raw = open(path, 'rb').read()
    for enc in ('utf-8-sig', 'utf-8', 'gbk'):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise RuntimeError('CSV 编码无法识别（试过 utf-8 / gbk）')
    return list(csv.reader(io.StringIO(text)))


def norm_id(v):
    """学号统一成纯数字字符串。Excel 常把学号存成 260101.0 这类浮点文本。"""
    s = str(v if v is not None else '').strip()
    if not s:
        return ''
    s = s.replace(' ', '').replace('\u3000', '')
    if re.fullmatch(r'\d+\.0+', s):
        s = s.split('.')[0]
    return re.sub(r'[^\d]', '', s)


def norm_name(v):
    s = re.sub(r'\s+', '', str(v if v is not None else ''))
    return s


def parse_roster(rows):
    """从二维表里抽出 (学号, 姓名)，自动识别表头列序。"""
    if not rows:
        raise RuntimeError('表格是空的')
    header = [str(x).strip().lower() for x in rows[0]]

    id_col = name_col = None
    for i, h in enumerate(header):
        if id_col is None and any(k.lower() in h for k in _ID_KEYS):
            id_col = i
        if name_col is None and any(k.lower() in h for k in _NAME_KEYS):
            name_col = i
    if id_col is None or name_col is None:
        raise RuntimeError('表头里没找到「学号」「姓名」列，实际表头：%s' % rows[0])

    out, seen, dups, bad = [], set(), [], []
    for r in rows[1:]:
        if max(id_col, name_col) >= len(r):
            continue
        sid, name = norm_id(r[id_col]), norm_name(r[name_col])
        if not sid:
            continue
        if len(sid) not in (4, 6):
            bad.append((sid, name))
            continue
        if sid in seen:
            dups.append(sid)
            continue
        seen.add(sid)
        out.append((sid, name))
    return out, dups, bad


def class_of(sid):
    """班级编码：6 位取前 4，4 位取前 2（与 app.student_id_to_class_code 一致）。"""
    return sid[:4] if len(sid) == 6 else sid[:2]


def load_json(path, default):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return default


def write_json(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def main():
    ap = argparse.ArgumentParser(description='导入学生姓名与学号对照表（取代现有学生名单）')
    ap.add_argument('src', help='xlsx 或 csv 路径')
    ap.add_argument('--keep-students', action='store_true',
                    help='只增补/更新名单中的学生，不删除名单外的旧学生账号')
    ap.add_argument('--dry-run', action='store_true', help='只打印将要发生的变化，不写盘')
    args = ap.parse_args()

    if not os.path.exists(args.src):
        print('文件不存在：%s' % args.src)
        return 2

    rows = read_xlsx(args.src) if args.src.lower().endswith(('.xlsx', '.xlsm')) else read_csv(args.src)
    roster, dups, bad = parse_roster(rows)

    print('解析到学生 %d 人' % len(roster))
    if dups:
        print('  ! 表里学号重复 %d 个（已只保留首次出现）：%s' % (len(dups), dups[:5]))
    if bad:
        print('  ! 学号长度非 4/6 位被跳过 %d 条：%s' % (len(bad), bad[:5]))

    by_class = {}
    for sid, name in roster:
        by_class.setdefault(class_of(sid), []).append(sid)
    print('覆盖班级 %d 个：%s%s' % (
        len(by_class), ', '.join(sorted(by_class)[:6]),
        ' …' if len(by_class) > 6 else ''))
    nameless = [s for s, n in roster if not n]
    if nameless:
        print('  ! 无姓名学生 %d 人：%s' % (len(nameless), nameless[:5]))

    users = load_json(USERS_FILE, [])
    kept = [u for u in users if u.get('role') != 'student']
    old_students = {str(u.get('id')): u for u in users if u.get('role') == 'student'}

    if args.keep_students:
        merged = dict(old_students)
        for sid, name in roster:
            prev = merged.get(sid) or {}
            # 姓名一律以名单为准；密码沿用已有密码，避免已登录学生被踢下线
            merged[sid] = {'id': sid, 'name': name or prev.get('name') or sid,
                           'role': 'student', 'password': prev.get('password') or sid}
        new_students = list(merged.values())
        removed = []
    else:
        new_students = [{'id': sid, 'name': name or sid, 'role': 'student', 'password': sid}
                        for sid, name in roster]
        removed = sorted(set(old_students) - {s for s, _ in roster})

    renamed = sum(1 for u in new_students
                  if u['id'] in old_students and old_students[u['id']].get('name') != u.get('name'))

    print('\n账号变化：')
    print('  学生 %d -> %d' % (len(old_students), len(new_students)))
    print('  保留非学生账号 %d 个：%s' % (len(kept), [u.get('id') for u in kept]))
    if renamed:
        print('  姓名被名单纠正 %d 人' % renamed)
    if removed:
        print('  移除名单外的旧学生 %d 人：%s%s' % (
            len(removed), removed[:8], ' …' if len(removed) > 8 else ''))

    old_class_files = sorted(f[:-5] for f in os.listdir(STUDENTS_DIR)
                             if f.endswith('.json')) if os.path.isdir(STUDENTS_DIR) else []
    stale = [c for c in old_class_files if c not in by_class]
    if stale:
        print('  清理已不存在的旧班级名单：%s' % stale)

    if args.dry_run:
        print('\n[dry-run] 未写任何文件。')
        return 0

    # 备份账号表，出问题可回滚
    if os.path.exists(USERS_FILE):
        bak = USERS_FILE + '.bak-%s' % time.strftime('%Y%m%d_%H%M%S')
        shutil.copy2(USERS_FILE, bak)
        print('\n已备份账号表 -> %s' % os.path.basename(bak))

    write_json(USERS_FILE, kept + new_students)

    # 班级名单按名单重建
    existing = set(old_class_files)
    for code, ids in by_class.items():
        write_json(os.path.join(STUDENTS_DIR, '%s.json' % code), sorted(ids))
    for code in stale:
        try:
            os.remove(os.path.join(STUDENTS_DIR, '%s.json' % code))
        except OSError:
            pass

    # 班级目录：建新的，清掉旧的空目录（非空目录保留，避免误删业务文件）
    for code in by_class:
        os.makedirs(os.path.join(CLASSES_DIR, code), exist_ok=True)
    for code in stale:
        p = os.path.join(CLASSES_DIR, code)
        if os.path.isdir(p) and not os.listdir(p):
            try:
                os.rmdir(p)
            except OSError:
                pass

    # 班级归属里指向已消失班级的条目一并清掉
    owners_file = os.path.join(DATA_DIR, 'config', 'class_owners.json')
    owners = load_json(owners_file, {})
    if isinstance(owners, dict):
        dropped = [c for c in owners if c not in by_class]
        if dropped:
            for c in dropped:
                owners.pop(c, None)
            write_json(owners_file, owners)
            print('  清理失效的班级归属：%s' % dropped)

    print('\n完成。%d 名学生、%d 个班级已入库。' % (len(new_students), len(by_class)))
    print('提示：学生端登录用学号，密码同学号。')
    return 0


if __name__ == '__main__':
    sys.exit(main())
