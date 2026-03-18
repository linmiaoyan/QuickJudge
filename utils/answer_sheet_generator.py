# -*- coding: utf-8 -*-
"""
根据试卷解析结果生成 A4 答题卡 HTML，可转为 PDF。
包含：学号、姓名/班级/日期、选择题涂卡区、主观题书写区、作文区（可带题干/材料）。
"""
from typing import Dict, Any, List, Optional

# 选项列
OPTIONS = ['A', 'B', 'C', 'D']


def _choice_block_html(sections: List[Dict], answers: Dict[int, str], title: str = '答题卡') -> str:
    """生成选择题涂卡区 HTML。"""
    if not sections:
        total = 0
        sections = [{'name': '选择题', 'start': 1, 'end': 0}]
    else:
        total = max(s.get('end', 0) for s in sections)
    if total <= 0:
        return ''

    lines = []
    lines.append('<div class="choice-section">')
    lines.append('<div class="choice-title">选择题</div>')
    lines.append('<div class="choice-grid">')

    for sec in sections:
        start, end = sec.get('start', 1), sec.get('end', 0)
        name = sec.get('name', '')
        if end < start:
            continue
        for n in range(start, end + 1):
            ans = answers.get(n, '')
            row = [f'<span class="q-num">{n}</span>']
            for opt in OPTIONS:
                checked = ' checked' if ans == opt else ''
                row.append(
                    f'<label class="opt"><input type="radio" name="q{n}" value="{opt}"{checked} disabled>'
                    f'<span class="opt-box">{opt}</span></label>'
                )
            lines.append('<div class="choice-row">' + ''.join(row) + '</div>')

    lines.append('</div></div>')
    return '\n'.join(lines)


def _subjective_block_html(sections: List[Dict]) -> str:
    """生成主观题书写区 HTML。"""
    if not sections:
        return ''
    lines = []
    for sec in sections:
        name = sec.get('name', '语法填空/短文改错/书面表达')
        num_lines = int(sec.get('num_lines', 12))
        prompt = (sec.get('prompt') or '').strip()
        lines.append('<div class="subjective-section">')
        lines.append(f'<div class="subjective-title">{name}</div>')
        if prompt:
            lines.append(f'<div class="subjective-prompt"><pre>{_escape_html(prompt)}</pre></div>')
        lines.append('<div class="subjective-lines">')
        for _ in range(num_lines):
            lines.append('<div class="line"></div>')
        lines.append('</div></div>')
    return '\n'.join(lines)


def _escape_html(s: str) -> str:
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def _build_styles() -> str:
    return """
* { box-sizing: border-box; }
body { font-family: "Microsoft YaHei", "SimSun", sans-serif; margin: 16px; font-size: 12px; }
@media print { body { margin: 0; } .no-print { display: none !important; } }
.page { max-width: 210mm; margin: 0 auto; padding: 12px; border: 1px solid #333; min-height: 297mm; }
h1 { text-align: center; font-size: 16px; margin-bottom: 12px; }
.meta { display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 12px; }
.meta label { font-weight: bold; min-width: 56px; }
.meta input { border: none; border-bottom: 1px solid #333; padding: 2px 6px; width: 100px; }
.student-id { margin-bottom: 16px; }
.student-id .label { font-weight: bold; margin-bottom: 6px; }
.student-id .boxes { display: flex; gap: 6px; flex-wrap: wrap; }
.student-id .box { width: 28px; height: 32px; border: 1.5px solid #333; text-align: center; line-height: 28px; font-size: 14px; }
.choice-section { margin-bottom: 16px; }
.choice-title { font-weight: bold; margin-bottom: 8px; }
.choice-grid { border: 1px solid #999; padding: 8px; }
.choice-row { display: flex; align-items: center; gap: 4px; margin-bottom: 4px; }
.choice-row .q-num { width: 22px; font-weight: bold; }
.choice-row .opt { display: flex; align-items: center; cursor: default; }
.choice-row .opt-box { width: 20px; height: 20px; border: 1px solid #333; margin-left: 2px; text-align: center; line-height: 18px; font-size: 11px; }
.subjective-section { margin-top: 14px; border: 1px solid #999; padding: 10px; }
.subjective-title { font-weight: bold; margin-bottom: 6px; }
.subjective-prompt { background: #f5f5f5; padding: 8px; margin-bottom: 8px; font-size: 11px; white-space: pre-wrap; max-height: 200px; overflow: auto; }
.subjective-lines .line { border-bottom: 1px solid #eee; min-height: 24px; }
"""


def generate_answer_sheet_html(
    parsed: Dict[str, Any],
    title: str = '答题卡',
    show_answer_keys: bool = True,
) -> str:
    """
    根据 parse_paper_docx 的返回结果生成完整答题卡 HTML。
    :param parsed: paper_parser.parse_paper_docx 的返回值
    :param title: 答题卡标题（如「限时练2」）
    :param show_answer_keys: 是否在选择题区显示参考答案（打印教师版可 True，学生版 False）
    """
    choice_sections = parsed.get('choice_sections') or []
    choice_answers = parsed.get('choice_answers') if show_answer_keys else {}
    subjective_sections = parsed.get('subjective_sections') or []

    choice_html = _choice_block_html(choice_sections, choice_answers, title)
    subjective_html = _subjective_block_html(subjective_sections)

    html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{_escape_html(title)}</title>
<style>
{_build_styles()}
</style>
</head>
<body>
<div class="page">
<h1>【A4】{_escape_html(title)}</h1>
<div class="meta">
  <label>班级</label><input type="text" placeholder="班级">
  <label>姓名</label><input type="text" placeholder="姓名">
  <label>日期</label><input type="text" placeholder="日期">
</div>
<div class="student-id">
  <div class="label">学号（六位数字，按 8 字形规范书写）</div>
  <div class="boxes">
    <div class="box">8</div><div class="box">8</div><div class="box">8</div>
    <div class="box">8</div><div class="box">8</div><div class="box">8</div>
  </div>
</div>
{choice_html}
{subjective_html}
</div>
</body>
</html>'''
    return html


def html_to_pdf(html: str, output_path: str) -> bool:
    """
    将答题卡 HTML 转为 PDF。需要安装 weasyprint 或 pdfkit。
    :return: 是否成功
    """
    try:
        from weasyprint import HTML
        from weasyprint.text.fonts import FontConfiguration
        font_config = FontConfiguration()
        doc = HTML(string=html)
        doc.write_pdf(output_path, font_config=font_config)
        return True
    except ImportError:
        try:
            import pdfkit
            pdfkit.from_string(html, output_path, options={'encoding': 'UTF-8'})
            return True
        except Exception:
            return False
    except Exception:
        return False
