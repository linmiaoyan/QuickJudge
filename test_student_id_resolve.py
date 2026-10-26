# -*- coding: utf-8 -*-
"""学号回查逻辑的回归测试：verify app.resolve_student_id()。

背景（务必理解再用）：
  高一学生卷面上**只写学号后四位**（260101 写成 0101）。
  四位号在不同年级之间会撞号——0101 同时对应 260101 / 250101 / 240101，
  所以不能再用「4 位学号前两位是班号」的旧规则推班级，必须回名单反查。

本测试直接读 data/config/users.json（真实名单），因此：
  · 名单没导过时，多数用例会落到「无匹配」，属预期；
  · 断言里用的 260101 等学号是名单首行，换学校后需同步调整。

运行：
    python test_student_id_resolve.py
"""
import app

# (卷面看到的值, 卷面班级号, 期望解析结果, 期望说明是否为空, 说明)
CASES = [
    ('0101', '01', '260101', True, '高一1班01号，班号两位'),
    ('0101', '1', '260101', True, '班号写成一位数也能对上'),
    ('0101', '', '260101', True, '没读到班号也能按年级前缀补全'),
    ('1401', '14', '261401', True, '高一14班01号'),
    ('1448', '14', '261448', True, '高一14班末位座号（该班座号不连续）'),
    ('0101', '14', '260101', False, '班号与后四位矛盾时必须给出说明'),
    ('1499', '', '1499', False, '后四位在名单里不存在'),
    ('260101', '', '260101', True, '完整六位直接命中'),
    ('240101', '', '240101', True, '高三完整学号'),
    ('999999', '', '999999', False, '名单里没有的学号'),
    ('9999', '', '9999', False, '后四位名单里没有'),
    ('123', '', '123', False, '位数异常'),
    ('', '', '', False, '空值'),
]


def main():
    idx = app._student_index()
    print('名单学生数：%d' % len(idx['ids']))
    if not idx['ids']:
        print('！名单为空，请先导入学生名单再跑本测试')
        return 2

    passed = failed = 0
    for seen, cls, expect_id, expect_clean, desc in CASES:
        got, note = app.resolve_student_id(seen, cls)
        # 名单里没有 2601xx 这类学号时，相关用例跳过而非误报
        if expect_id.startswith('26') and expect_id not in idx['ids']:
            print('SKIP  %-8s %-6s %s（名单里没有 %s）' % (seen, cls, desc, expect_id))
            continue
        ok = (got == expect_id) and (bool(note) != expect_clean)
        print('%s  %-8s 班号%-6s -> %-8s %s %s' % (
            'PASS' if ok else 'FAIL', seen or "''", cls or "''", got or "''",
            ('说明：' + note) if note else '已唯一命中', '| ' + desc))
        passed, failed = (passed + 1, failed) if ok else (passed, failed + 1)

    print('\n通过 %d，失败 %d' % (passed, failed))
    return 0 if failed == 0 else 1


if __name__ == '__main__':
    raise SystemExit(main())
