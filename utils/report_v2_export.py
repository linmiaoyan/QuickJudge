# -*- coding: utf-8 -*-
"""学生报告 v2 的班级批量导出（打印友好 HTML）。

背景：旧的 `student_report_generator.generate_class_report_html` 只认「阅卷任务」那套
`results[filename] = {report, student_id, score_text}` 结构；报告 v2 之后，自动流水线把结果
写进 `data/records/` 的 `essay` 对象（errors / vocab_upgrades / cohesion / model_essay …），
两者对不上，导致用新流水线的班级根本导不出可打印的报告。

这里直接消费 v2 的 essay 结构 + 批次语料，一个学生一页正文、一页语料，浏览器打印即可出 PDF。
不引入第三方依赖，便于在任意机器上跑。
"""
from typing import Any, Dict, List, Optional

_STATUS_CN = {'covered': '已覆盖', 'partial': '部分', 'missing': '缺失'}
_STATUS_COLOR = {'covered': '#1a7f37', 'partial': '#9a6700', 'missing': '#b42318'}


def _e(s: Any) -> str:
    """HTML 转义（含换行）。"""
    if s is None:
        return ''
    s = str(s)
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;').replace('\n', '<br>\n'))


def _score_text(essay: Dict[str, Any], fallback: str = '') -> str:
    sc, mx = essay.get('score'), essay.get('max_score')
    if sc is None:
        return _e(fallback) if fallback else '—'
    try:
        return '%s / %s' % (_num(sc), _num(mx))
    except Exception:
        return _e(fallback) if fallback else '—'


def _num(v: Any) -> str:
    try:
        f = float(v)
        return str(int(f)) if f == int(f) else ('%.1f' % f)
    except (TypeError, ValueError):
        return str(v)


def _styles() -> str:
    return """
@page { size: A4; margin: 14mm 12mm; }
* { box-sizing: border-box; }
body { font-family: "Microsoft YaHei", "PingFang SC", "SimSun", sans-serif;
       font-size: 11.5px; color: #1f2328; margin: 0; background: #f3f4f6; }
.page { max-width: 210mm; min-height: 260mm; margin: 0 auto 14px; padding: 12mm 11mm;
        background: #fff; position: relative; }
@media print { body { background: #fff; } .page { margin: 0; padding: 0; min-height: 0;
              page-break-after: always; border: 0; box-shadow: none; }
              .page:last-child { page-break-after: auto; } }
@media screen { .page { box-shadow: 0 1px 3px rgba(0,0,0,.12); border-radius: 6px; } }
.hd { display: flex; justify-content: space-between; align-items: baseline;
      border-bottom: 2px solid #1f2328; padding-bottom: 6px; margin-bottom: 12px; flex-wrap: wrap; gap: 4px; }
.hd .who { font-size: 12.5px; }
.hd .who b { font-size: 14px; }
.hd .meta { font-size: 10.5px; color: #57606a; }
h1 { font-size: 13px; margin: 0 0 10px; font-weight: 700; }
h2 { font-size: 12px; margin: 14px 0 6px; padding-left: 7px; border-left: 3px solid #1f6feb;
     line-height: 1.5; }
.muted { color: #6e7781; }
.empty { color: #8c959f; font-style: italic; }
table { width: 100%; border-collapse: collapse; table-layout: fixed; }
th, td { border: 1px solid #d0d7de; padding: 4px 6px; vertical-align: top; text-align: left;
         word-break: break-word; }
th { background: #f6f8fa; font-weight: 600; font-size: 10.5px; }
.quote { font-family: Georgia, "Times New Roman", serif; }
.fix { color: #1a7f37; font-family: Georgia, serif; }
.why { color: #57606a; font-size: 10.5px; }
.scorebar { display: flex; gap: 18px; align-items: baseline; flex-wrap: wrap; margin: 4px 0 8px; }
.scorebar .big { font-size: 22px; font-weight: 800; color: #1f6feb; }
.tag { display: inline-block; padding: 1px 7px; border-radius: 10px; font-size: 10px;
       background: #ddf4ff; color: #0969da; margin-left: 6px; }
.before { color: #b42318; text-decoration: line-through; }
.after { color: #1a7f37; font-weight: 600; }
.arrow { color: #8c959f; margin: 0 4px; }
ul.lst { margin: 0; padding-left: 18px; }
ul.lst li { margin: 3px 0; }
.pair { margin: 5px 0; }
.pair .lab { display: inline-block; min-width: 52px; color: #57606a; font-size: 10.5px; }
.essay { background: #f6f8fa; border: 1px solid #d0d7de; border-radius: 5px;
         padding: 9px 11px; line-height: 1.85; font-family: Georgia, "Times New Roman", serif; font-size: 12px; }
.corpus-item { margin: 5px 0 7px; }
.corpus-item .en { font-family: Georgia, serif; font-size: 11.8px; }
.corpus-item .cn { color: #57606a; font-size: 10.8px; }
.corpus-item .use { color: #0969da; font-size: 10.2px; }
.point-title { font-weight: 700; margin: 10px 0 4px; color: #1f2328; }
.foot { position: absolute; bottom: 8mm; left: 11mm; right: 11mm; color: #8c959f;
        font-size: 9.5px; border-top: 1px dashed #d0d7de; padding-top: 4px; }
"""


def _section(title: str, body: str) -> str:
    if not body:
        return ''
    return '<h2>%s</h2>%s' % (_e(title), body)


def _requirements_html(reqs: List[Dict[str, Any]]) -> str:
    if not reqs:
        return ''
    rows = []
    for r in reqs:
        st = str(r.get('status') or 'partial')
        cn = _STATUS_CN.get(st, st)
        color = _STATUS_COLOR.get(st, '#57606a')
        rows.append(
            '<tr><td style="width:38%%">%s</td>'
            '<td style="width:16%%;color:%s;font-weight:600">%s</td>'
            '<td>%s</td></tr>' % (_e(r.get('point')), color, _e(cn), _e(r.get('evidence')) or '<span class="muted">—</span>'))
    return ('<table><thead><tr><th>要点</th><th>覆盖</th><th>原文依据</th></tr></thead>'
            '<tbody>%s</tbody></table>' % ''.join(rows))


def _errors_html(errors: List[Dict[str, Any]]) -> str:
    if not errors:
        return ''
    rows = []
    for i, e in enumerate(errors, 1):
        rows.append(
            '<tr><td style="width:5%%;text-align:center">%d</td>'
            '<td style="width:12%%">%s</td>'
            '<td style="width:40%%" class="quote">%s</td>'
            '<td style="width:43%%"><span class="fix">%s</span>'
            '%s</td></tr>' % (
                i, _e(e.get('type') or '语法'), _e(e.get('quote')) or '<span class="muted">（未定位）</span>',
                _e(e.get('fix')),
                ('<div class="why">%s</div>' % _e(e.get('why'))) if e.get('why') else ''))
    return ('<table><thead><tr><th>#</th><th>类型</th><th>原文</th><th>改为 / 错因</th></tr></thead>'
            '<tbody>%s</tbody></table>' % ''.join(rows))


def _vocab_html(items: List[Dict[str, Any]]) -> str:
    if not items:
        return ''
    out = []
    for v in items:
        out.append('<div class="pair"><span class="lab">替换</span>'
                   '<span class="before">%s</span><span class="arrow">→</span>'
                   '<span class="after">%s</span>%s</div>' % (
                       _e(v.get('original')), _e(v.get('upgrade')),
                       ('<div class="why" style="margin-left:52px">%s</div>' % _e(v.get('sentence'))) if v.get('sentence') else ''))
    return '<div>%s</div>' % ''.join(out)


def _pattern_html(items: List[Dict[str, Any]]) -> str:
    if not items:
        return ''
    out = []
    for p in items:
        extra = ' · '.join(x for x in [p.get('technique'), p.get('added')] if x)
        out.append('<div class="pair">'
                   '<div><span class="lab">原句</span><span class="quote">%s</span></div>'
                   '<div><span class="lab">改写</span><span class="fix">%s</span></div>'
                   '%s%s</div>' % (
                       _e(p.get('original')), _e(p.get('improved')),
                       ('<div class="why" style="margin-left:52px">手法：%s</div>' % _e(extra)) if extra else '',
                       ('<div class="why" style="margin-left:52px">%s</div>' % _e(p.get('why'))) if p.get('why') else ''))
    return '<div>%s</div>' % ''.join(out)


def _cohesion_html(items: List[Dict[str, Any]]) -> str:
    if not items:
        return ''
    out = []
    for c in items:
        before = c.get('before') or '无'
        after = c.get('after') or ''
        out.append('<div class="pair">'
                   '<div><span class="lab">%s</span><span class="before">%s</span>'
                   '<span class="arrow">→</span><span class="after">%s</span></div>'
                   '<div class="why" style="margin-left:52px">%s</div></div>' % (
                       _e(c.get('relation') or '衔接'), _e(before), _e(after), _e(c.get('why'))))
    return '<div>%s</div>' % ''.join(out)


def _content_html(items: List[Dict[str, Any]]) -> str:
    if not items:
        return ''
    out = []
    for c in items:
        out.append('<div class="pair">'
                   '<div><span class="lab">%s</span><span class="before">%s</span></div>'
                   '<div><span class="lab">提升</span><span class="fix">%s</span></div>'
                   '%s</div>' % (
                       _e(c.get('node') or '情节'), _e(c.get('original')), _e(c.get('improved')),
                       ('<div class="why" style="margin-left:52px">%s</div>' % _e(c.get('why'))) if c.get('why') else ''))
    return '<div>%s</div>' % ''.join(out)


def _corpus_html(corpus: Optional[Dict[str, Any]]) -> str:
    if not isinstance(corpus, dict):
        return ''
    points = corpus.get('points') or []
    parts = []
    for p in points:
        items = p.get('items') or []
        if not items:
            continue
        rows = []
        for it in items:
            rows.append('<div class="corpus-item">'
                        '<div class="en">%s</div>'
                        '%s%s</div>' % (
                            _e(it.get('text')),
                            ('<div class="cn">%s</div>' % _e(it.get('trans'))) if it.get('trans') else '',
                            ('<div class="use">用法：%s</div>' % _e(it.get('usage'))) if it.get('usage') else ''))
        parts.append('<div class="point-title">%s</div>%s' % (_e(p.get('point') or '要点'), ''.join(rows)))
    pitfalls = corpus.get('pitfalls') or []
    if pitfalls:
        rows = []
        for p in pitfalls:
            rows.append('<tr><td>%s</td><td class="quote"><span class="before">%s</span></td>'
                        '<td><span class="fix">%s</span></td></tr>' % (
                            _e(p.get('point')), _e(p.get('wrong')), _e(p.get('right'))))
        parts.append('<div class="point-title">易错提醒</div>'
                     '<table><thead><tr><th style="width:22%%">要点</th><th>常见错误</th><th>正确写法</th></tr></thead>'
                     '<tbody>%s</tbody></table>' % ''.join(rows))
    return ''.join(parts)


def _student_pages(rec: Dict[str, Any], name: str, corpus: Optional[Dict[str, Any]],
                   class_name: str, school_name: str = '') -> str:
    essay = rec.get('essay') or {}
    task = essay.get('task') or {}
    sid = rec.get('student_id') or ''
    title = task.get('title') or rec.get('task_title') or '英语作文'
    wc = essay.get('word_count') or {}
    wc_txt = ''
    if wc:
        actual, required = wc.get('actual'), wc.get('required')
        if actual:
            wc_txt = '词数 %s%s' % (_num(actual), (' / 要求 %s' % _num(required)) if required else '')

    header = (
        '<div class="hd"><div class="who">'
        '%s<b>%s</b>　<span class="muted">%s</span>'
        '</div><div class="meta">%s　提交 %s</div></div>' % (
            ('%s · ' % _e(school_name)) if school_name else '',
            _e(name or sid or '未知学生'),
            _e(class_name or rec.get('class_name') or '') + ('　学号 %s' % _e(sid) if sid else ''),
            _e(title), _e(rec.get('submitted_at') or '')))

    scorebar = ('<div class="scorebar"><div><span class="big">%s</span>'
                '<span class="muted"> 分</span></div>'
                '%s%s%s</div>' % (
                    _score_text(essay, rec.get('score_text')),
                    ('<div><span class="muted">等级</span> %s</div>' % _e(essay.get('level'))) if essay.get('level') else '',
                    ('<div class="muted">%s</div>' % _e(wc_txt)) if wc_txt else '',
                    ('<div class="tag">%s</div>' % _e(rec.get('subject_name'))) if rec.get('subject_name') else ''))

    # 字迹读不准
    illeg = essay.get('illegible') or []
    illeg_html = ''
    if illeg:
        rows = ''.join('<tr><td class="quote">%s</td><td>%s</td><td class="why">%s</td></tr>' % (
            _e(u.get('as_read')), _e(u.get('guess')), _e(u.get('why'))) for u in illeg)
        illeg_html = ('<table><thead><tr><th style="width:32%%">识别为</th><th style="width:30%%">可能是</th>'
                      '<th>说明</th></tr></thead><tbody>%s</tbody></table>' % rows)

    comment = ('<div class="essay" style="font-family:inherit;line-height:1.8">%s</div>' % _e(essay.get('comment'))) \
        if essay.get('comment') else ''

    award = essay.get('sentence_award') or {}
    award_html = ''
    if award.get('quote'):
        award_html = ('<div class="pair"><span class="quote">“%s”</span>%s</div>' % (
            _e(award.get('quote')),
            ('<div class="why">%s</div>' % _e(award.get('why'))) if award.get('why') else ''))

    model = essay.get('model_essay') or ''
    model_html = ('<div class="essay">%s</div>%s' % (
        _e(model),
        ('<div class="why" style="margin-top:6px">%s</div>' % _e(essay.get('model_essay_note'))) if essay.get('model_essay_note') else '')) \
        if model else ''

    page1 = (
        '<div class="page">%s<h1>%s</h1>%s'
        '%s%s%s%s%s%s%s%s%s'
        '<div class="foot">QuickJudge 自动批阅 · %s</div></div>' % (
            header, _e(title), scorebar,
            _section('AI 评语', comment),
            _section('字迹读不准（影响得分的识别风险）', illeg_html),
            _section('要点覆盖', _requirements_html(essay.get('requirements_check') or [])),
            _section('语篇批改（原文 → 改为 / 错因）', _errors_html(essay.get('errors') or [])),
            _section('本篇佳句', award_html),
            _section('高级词汇', _vocab_html(essay.get('vocab_upgrades') or [])),
            _section('高级句式', _pattern_html(essay.get('pattern_upgrades') or [])),
            _section('逻辑衔接提升' if not (essay.get('content_upgrade') or []) else '内容提升',
                     _cohesion_html(essay.get('cohesion') or []) or _content_html(essay.get('content_upgrade') or [])),
            _section('个性化范文', model_html),
            _score_text(essay, rec.get('score_text'))))

    corpus_body = _corpus_html(corpus)
    page2 = ''
    if corpus_body:
        page2 = ('<div class="page">%s<h1>题目专属语料（带翻译，便于背诵）</h1>%s'
                 '<div class="foot">语料按题目要点分组 · 先背英文，再对照中文检查理解</div></div>' % (header, corpus_body))
    return page1 + page2


def build_class_report_html(records: List[Dict[str, Any]],
                            names_map: Optional[Dict[str, str]] = None,
                            class_name: str = '',
                            corpus_map: Optional[Dict[str, Dict[str, Any]]] = None,
                            school_name: str = '') -> str:
    """把一个班级的 v2 记录渲染成可打印 HTML。records 为 records/index.json 的子集。"""
    names_map = names_map or {}
    corpus_map = corpus_map or {}
    sorted_recs = sorted(records, key=lambda r: (str(r.get('student_id') or ''), r.get('submitted_at') or ''))
    pages = []
    for rec in sorted_recs:
        sid = str(rec.get('student_id') or '')
        name = names_map.get(sid) or rec.get('student_name') or ''
        corpus = corpus_map.get(str(rec.get('batch_id') or '')) or {}
        pages.append(_student_pages(rec, name, corpus, class_name, school_name))
    title = '学生报告 · %s' % (class_name or '全部班级')
    return (
        '<!DOCTYPE html>\n<html lang="zh-CN"><head><meta charset="UTF-8">'
        '<meta name="viewport" content="width=device-width, initial-scale=1.0">'
        '<title>%s</title><style>%s</style></head><body>%s</body></html>' % (
            _e(title), _styles(), ''.join(pages) or
            '<div class="page"><p class="empty">没有可导出的已批阅记录。</p></div>'))
