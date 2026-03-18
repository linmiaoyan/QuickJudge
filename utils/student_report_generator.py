# -*- coding: utf-8 -*-
"""
学生个人报告 HTML 生成：按「学生个人报告单」两页版式生成，便于全班合并导出为 PDF。
每生第1页：表头（学校/班级/姓名/学号/作答时间）、题型、得分与评语、批改结果（原文+错误标注或评语全文）。
每生第2页：范文、题目专属语料（可选）。
"""
from typing import Dict, Any, List, Optional


def _escape(s: str) -> str:
    if not s:
        return ''
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;').replace('\n', '<br>\n'))


def _report_styles() -> str:
    return """
.report-page { font-family: "Microsoft YaHei", "SimSun", sans-serif; font-size: 12px; margin: 0 auto; max-width: 210mm; padding: 14px; min-height: 297mm; box-sizing: border-box; }
.report-page h2 { font-size: 14px; margin: 10px 0 6px; border-bottom: 1px solid #333; }
.report-header { display: flex; justify-content: space-between; flex-wrap: wrap; margin-bottom: 12px; font-size: 12px; }
.report-section { margin-bottom: 12px; }
.report-section .label { font-weight: bold; margin-bottom: 4px; }
.report-content { white-space: pre-wrap; line-height: 1.5; }
.report-divider { margin: 16px 0; border-top: 1px dashed #999; }
@media print { .report-page { page-break-after: always; } .report-page:last-child { page-break-after: auto; } }
"""


def generate_one_student_report_html(
    student_id: str,
    student_name: str,
    class_name: str,
    report_text: str,
    school_name: str = '',
    answer_time: str = '',
    task_title: str = '作文练习',
    model_essay: str = '',
    topic_materials: Optional[List[Dict[str, str]]] = None,
    score_text: str = '',
) -> str:
    """
    生成单个学生的两页报告 HTML。
    :param report_text: 批阅报告全文（评语+批改结果，当前系统存为一段文本）
    :param model_essay: 题目范文（可选，全班统一）
    :param topic_materials: 题目专属语料 [{"point": "要点1", "items": [{"text": "语料", "trans": "翻译"}]}]
    :param score_text: 得分说明，如 "得分：8（满分15，最高分12，班级平均分8.5）"
    """
    header_line = f"学校：{school_name or '—'}  班级：{class_name or '—'}  姓名：{student_name or student_id or '—'}  学号：{student_id or '—'}"
    if answer_time:
        header_line += f"\t作答时间：{answer_time}"
    else:
        header_line += "\t作答时间：—"

    # 第1页
    page1 = f'''
<div class="report-page">
<div class="report-header">{_escape(header_line)}</div>
<h2>{_escape(task_title)}</h2>
'''
    if score_text:
        page1 += f'<div class="report-section"><div class="label">作文得分</div><div class="report-content">{_escape(score_text)}</div></div>'
    page1 += '''
<div class="report-section"><div class="label">评语 / 批改结果</div>
<div class="report-content">'''
    page1 += _escape(report_text) if report_text else '（暂无批阅内容）'
    page1 += '</div></div></div>'

    # 第2页：范文 + 题目专属语料
    page2 = f'''
<div class="report-page">
<div class="report-header">{_escape(header_line)}</div>
'''
    if model_essay:
        page2 += '<h2>范文</h2><div class="report-section"><div class="report-content">' + _escape(model_essay) + '</div></div>'
    if topic_materials:
        page2 += '<h2>题目专属语料</h2>'
        for mat in topic_materials:
            point = mat.get('point') or mat.get('name') or ''
            items = mat.get('items') or mat.get('materials') or []
            if point:
                page2 += f'<div class="report-section"><div class="label">{_escape(point)}</div>'
            for it in items:
                text = it.get('text') or it.get('语料') or ''
                trans = it.get('trans') or it.get('翻译') or ''
                if text:
                    page2 += f'<div class="report-content">【语料】{_escape(text)}</div>'
                if trans:
                    page2 += f'<div class="report-content">【翻译】{_escape(trans)}</div>'
            if point:
                page2 += '</div>'
    page2 += '</div>'

    return page1 + page2


def generate_class_report_html(
    results: List[Dict[str, Any]],
    task: Dict[str, Any],
    student_names_map: Dict[str, str],
    school_name: str = '',
    class_name: str = '',
) -> str:
    """
    根据任务批阅结果生成全班学生个人报告合并 HTML。
    :param results: 来自 task_results 的列表，每项含 student_id, report, status 等
    :param task: 任务对象，可选 model_essay, topic_materials
    :param student_names_map: 学号 -> 姓名
    :param school_name: 学校名称（可选）
    :param class_name: 班级名称（用于表头，若空则从 task.class_names 取第一个）
    """
    class_name = class_name or (task.get('class_names') or [''])[0]
    task_title = (task.get('title') or '作文练习').strip()
    model_essay = (task.get('model_essay') or '').strip()
    topic_materials = task.get('topic_materials') or []

    # 按学号排序
    sorted_results = sorted(results, key=lambda r: (r.get('student_id') or '', r.get('filename') or ''))

    full_html = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>学生个人报告 - ''' + _escape(task_title) + '''</title>
<style>
''' + _report_styles() + '''
</style>
</head>
<body>
'''
    for i, rec in enumerate(sorted_results):
        sid = rec.get('student_id') or ''
        name = student_names_map.get(sid, '') or rec.get('student_name', '')
        report_text = rec.get('report') or ''
        score_text = rec.get('score_text') or ''  # 若后续批改写入得分可在此展示
        full_html += generate_one_student_report_html(
            student_id=sid,
            student_name=name,
            class_name=class_name,
            report_text=report_text,
            school_name=school_name,
            answer_time=rec.get('answer_time', ''),
            task_title=task_title,
            model_essay=model_essay,
            topic_materials=topic_materials,
            score_text=score_text,
        )
        if i < len(sorted_results) - 1:
            full_html += '<div class="report-divider"></div>'
    full_html += '</body></html>'
    return full_html
