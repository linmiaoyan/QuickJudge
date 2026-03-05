# -*- coding: utf-8 -*-
"""生成模拟数据：班级目录、学生列表、班级中心报告。运行一次即可。"""
import os
import json

DATA_DIR = os.path.join(os.path.dirname(__file__), '..')
CLASSES_DIR = os.path.join(DATA_DIR, 'classes')
STUDENTS_DIR = os.path.join(DATA_DIR, 'students')
CLASS_CENTER_DIR = os.path.join(DATA_DIR, 'class_center')

os.makedirs(CLASSES_DIR, exist_ok=True)
os.makedirs(STUDENTS_DIR, exist_ok=True)
os.makedirs(CLASS_CENTER_DIR, exist_ok=True)

# 1. 班级文件夹结构：高一1班 两个日期，高二3班 一个日期
for cls, dates in [('高一1班', ['2025-03-01', '2025-03-05']), ('高二3班', ['2025-03-04'])]:
    for d in dates:
        path = os.path.join(CLASSES_DIR, cls, d)
        os.makedirs(path, exist_ok=True)

# 2. 学生列表（6位学号）
students = {
    '高一1班': ['230101', '230102', '230103', '230104', '230105'],
    '高二3班': ['220301', '220302', '220303'],
}
for cls, ids in students.items():
    p = os.path.join(STUDENTS_DIR, cls + '.json')
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(ids, f, ensure_ascii=False, indent=2)

# 3. 班级中心：每个学生若干条批阅记录
def make_record(date, ts, filename, file_path, essay_text, report, class_eval):
    return {
        'date': date,
        'timestamp': ts,
        'filename': filename,
        'file_path': file_path,
        'essay_text': essay_text,
        'report': report,
        'class_evaluation': class_eval,
    }

class_eval_1 = """本次作文整体完成度较好。多数同学能围绕主题展开，结构完整。
共同优点：段落清晰，有过渡句。
普遍问题：部分同学高级词汇运用不足，可多积累同义替换。
教学建议：加强应用文格式训练，多做限时写作。"""

class_eval_2 = """本次批阅显示班级写作水平稳定。建议后续加强读后续写的衔接与创意表达训练。"""

essay_sample = """Dear Chris,
I'm glad to tell you about our art class in the park last week. We drew grass, trees and birds. The teacher taught us how to draw an apple tree. I enjoyed it very much. I hope you can join us next time.
Yours,
Li Hua"""

report_sample = """总体评分：12/15
优点：格式正确，要点齐全，能完成基本交际任务。
问题：句式较简单，可适当使用定语从句或状语从句；个别拼写需注意。
改进建议：增加连接词（如 Besides, What's more），丰富句式。"""

class_center = {
    '高一1班': {
        '230101': [
            make_record('2025-03-05', '2025-03-05T10:30:00', 'scan_001.png', '高一1班/2025-03-05/scan_001.png', essay_sample, report_sample, class_eval_1),
            make_record('2025-03-01', '2025-03-01T14:20:00', 'img_101.png', '高一1班/2025-03-01/img_101.png', 'Last week we had an English corner. I talked with two students. We discussed how to learn new words. It was helpful.', '总体评分：11/15。内容完整，表达清楚。可多使用高级词汇。', class_eval_1),
        ],
        '230102': [
            make_record('2025-03-05', '2025-03-05T10:31:00', 'scan_002.png', '高一1班/2025-03-05/scan_002.png', essay_sample, '总体评分：13/15。书写工整，要点全面，有少量语法错误。', class_eval_1),
        ],
        '230103': [
            make_record('2025-03-05', '2025-03-05T10:32:00', 'scan_003.png', '高一1班/2025-03-05/scan_003.png', essay_sample, '总体评分：10/15。需注意时态一致与单词拼写。', class_eval_1),
        ],
    },
    '高二3班': {
        '220301': [
            make_record('2025-03-04', '2025-03-04T09:15:00', 'page1.png', '高二3班/2025-03-04/page1.png', essay_sample, report_sample, class_eval_2),
        ],
        '220302': [
            make_record('2025-03-04', '2025-03-04T09:16:00', 'page2.png', '高二3班/2025-03-04/page2.png', essay_sample, '总体评分：14/15。结构清晰，语言流畅。', class_eval_2),
        ],
    },
}

for cls, students_data in class_center.items():
    p = os.path.join(CLASS_CENTER_DIR, cls + '.json')
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(students_data, f, ensure_ascii=False, indent=2)

print('Mock data generated: classes (高一1班, 高二3班), students, class_center reports.')
