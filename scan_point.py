#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""QuickJudge 扫描点代理 —— 只做三件事：驱动本机扫描仪、把图传上去、显示进度。

设计取舍：
  * **只用标准库**。扫描点那台电脑装完 Python 就能跑，不装 Flask、不装 requests、
    不需要 pip（学校机房常常没有外网，多一个依赖就多一道坎）。
  * 本地页面和接口同源（127.0.0.1:5101）。让主机页面来调本机接口会被浏览器的
    私有网络访问策略（PNA）拦掉，这个坑不踩。
  * 一页一个 POST（原始字节），最后 commit 建批。断网只需重传那一页。

用法：
    双击 run_agent.cmd（或 pythonw scan_point.py），浏览器打开 http://127.0.0.1:5101/
"""
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE, 'config.json')
SCAN_ROOT = os.path.join(BASE, 'scans')

DEFAULTS = {
    'host': '',
    'token': '',
    'point_id': '',
    'point_name': '',
    'port': 5101,
    'naps2_path': '',
    'scan': {'driver': '', 'device': '', 'source': 'feeder',
             'pagesize': 'a4', 'dpi': 300, 'bitdepth': 'color'},
}

# 一页一张图，40 页大约 12~40MB。总时长按 40 页彩色 300dpi 估：扫描 100s + 导出 60s，
# 留到 10 分钟；静默 120 秒没输出就认为驱动卡住了。
# 为什么不压到 60 秒：实测这台 TWAIN 扫平板的设备，从"Starting scan"到真正报
# “TWAIN error: NoDS”中间要等 90 秒左右。阈值压在它前面，看到的永远是笼统的
# "卡住了"，真正原因（数据源打不开）反而被吞掉。
SCAN_TOTAL_SECONDS = 600
SCAN_STALL_SECONDS = 120
SCAN_NO_OUTPUT_SECONDS = 45
SCAN_EXPORT_GRACE = 180


# ---------------------------- 配置 ----------------------------

def load_config():
    cfg = {}
    try:
        with open(CONFIG_PATH, encoding='utf-8') as f:
            cfg = json.load(f) or {}
    except (OSError, ValueError):
        cfg = {}
    out = dict(DEFAULTS)
    out.update({k: v for k, v in cfg.items() if v is not None})
    scan = dict(DEFAULTS['scan'])
    scan.update(cfg.get('scan') or {})
    out['scan'] = scan
    return out


def save_config(cfg):
    tmp = CONFIG_PATH + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    os.replace(tmp, CONFIG_PATH)


CFG = load_config()


def find_naps2():
    """找 NAPS2.Console.exe。找到就记进配置，下次不用再找。"""
    p = (CFG.get('naps2_path') or '').strip()
    if p and os.path.isfile(p):
        return p
    cands = [
        r'C:\Program Files\NAPS2\App\NAPS2.Console.exe',
        r'C:\Program Files (x86)\NAPS2\App\NAPS2.Console.exe',
        r'D:\Tools\NAPS2\App\NAPS2.Console.exe',
        os.path.join(BASE, 'NAPS2', 'App', 'NAPS2.Console.exe'),
    ]
    for c in cands:
        if os.path.isfile(c):
            CFG['naps2_path'] = c
            try:
                save_config(CFG)
            except OSError:
                pass
            return c
    w = shutil.which('NAPS2.Console.exe') or shutil.which('naps2.console')
    return w or ''


# ---------------------------- 与主机通讯 ----------------------------

def _opener():
    # 学校网络常常配了系统代理；连内网主机必须绕开，否则请求会被送去代理服务器
    return urllib.request.build_opener(urllib.request.ProxyHandler({}))


def api(path, payload=None, headers=None, method=None, timeout=30, raw=False):
    """调主机接口。返回 (ok, 数据)。数据是解析后的 JSON，或纯文本。"""
    host = (CFG.get('host') or '').rstrip('/')
    if not host:
        return False, '还没有配置主机地址'
    url = host + path
    hdrs = {'X-QJ-Scan-Token': CFG.get('token') or ''}
    hdrs.update(headers or {})
    data = None
    if payload is not None:
        if isinstance(payload, (bytes, bytearray)):
            data = bytes(payload)
        else:
            data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
            hdrs.setdefault('Content-Type', 'application/json')
    req = urllib.request.Request(url, data=data, headers=hdrs, method=method)
    try:
        with _opener().open(req, timeout=timeout) as r:
            body = r.read()
            if raw:
                return True, body
            try:
                return True, json.loads(body.decode('utf-8'))
            except ValueError:
                return True, body.decode('utf-8', 'replace')
    except urllib.error.HTTPError as e:
        try:
            j = json.loads(e.read().decode('utf-8'))
            return False, j.get('error') or ('HTTP %s' % e.code)
        except Exception:
            return False, 'HTTP %s' % e.code
    except Exception as e:
        return False, '连不上主机（%s）。检查这台电脑和服务器是否在同一网络。' % e


# ---------------------------- 扫描 ----------------------------

STATE = {
    'stage': 'idle',      # idle / scanning / uploading / committing / done / error
    'pages': 0,
    'uploaded': 0,
    'batch_id': '',
    'message': '',
    'error': '',
    'started': 0,
}
_STATE_LOCK = threading.Lock()
_SCAN_LOCK = threading.Lock()


def state_set(**kw):
    with _STATE_LOCK:
        STATE.update(kw)


def state_get():
    with _STATE_LOCK:
        return dict(STATE)


def _run_naps2(cmd, out_dir, base_name):
    """跑一次 NAPS2，返回 (ok, 诊断, 文本)。

    和主机 app.py 里那套看门狗同源，但精简过：主机要区分卡纸/脱机/静默，
    扫描点这边只需要"跑完没有、要不要收手"。
    """
    marks = {'last': time.time(), 'n': 0, 'pages': None, 'done': threading.Event()}
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    lines = []

    def _read():
        try:
            for line in iter(proc.stdout.readline, b''):
                lines.append(line)
                marks['last'] = time.time()
                marks['n'] += 1
                low = line.lower()
                m = re.search(rb'(\d+)\s+page\(s\)\s+scanned', low)
                if m:
                    marks['pages'] = int(m.group(1))
                    if marks['pages'] == 0:
                        marks['done'].set()
                # 扫完不等于存完：NAPS2 之后还要 "Exporting image k of N"，
                # 只有 finished saving 才是真的写完了
                if b'finished saving' in low:
                    marks['done'].set()
        except Exception:
            pass
        finally:
            marks['done'].set()

    threading.Thread(target=_read, daemon=True).start()
    started = time.time()
    diag = 'ok'
    while True:
        if marks['done'].wait(0.5):
            diag = 'ok'
            break
        now = time.time()
        if marks['n'] == 0 and now - started > SCAN_NO_OUTPUT_SECONDS:
            diag = 'no_output'
            break
        if marks['n'] > 0 and now - marks['last'] > SCAN_STALL_SECONDS:
            diag = 'stall'
            break
        if now - started > SCAN_TOTAL_SECONDS:
            diag = 'timeout'
            break

    if diag == 'ok':
        # 必须等进程自己退出：目录里出现第 1 张图不代表后面的页存好了
        deadline = time.time() + SCAN_EXPORT_GRACE
        while proc.poll() is None and time.time() < deadline:
            time.sleep(0.3)
    if proc.poll() is None:
        try:
            proc.kill()
        except Exception:
            pass
    try:
        proc.stdout.close()
    except Exception:
        pass
    text = b''.join(lines).decode('utf-8', 'replace').strip()
    return (diag == 'ok'), diag, text


def _collect_pages(out_dir, base_name):
    """NAPS2 多页输出为 scan.1.png / scan.2.png…，单页是 scan.png。"""
    pages = []
    if not os.path.isdir(out_dir):
        return pages
    base = os.path.splitext(base_name)[0]
    for name in os.listdir(out_dir):
        if not name.lower().endswith('.png'):
            continue
        stem = os.path.splitext(name)[0]
        if stem == base or stem.startswith(base + '.'):
            pages.append(os.path.join(out_dir, name))

    def key(p):
        m = re.search(r'\.(\d+)\.png$', p.lower())
        return (1, int(m.group(1))) if m else (0, 0)

    pages.sort(key=key)
    return pages


def explain_failure(diag, text):
    """把 NAPS2/驱动的原始报错翻成老师能看懂的一句话。

    顺序有讲究：NoDS / TWAIN 这类"设备根本没连上"要先判——它们的输出里
    同样带 "0 page(s) scanned"，先按页数判就会说成"没放纸"，老师会去反复
    整理纸张，而真正的问题是 USB 线掉了。
    """
    t = (text or '').lower()
    if 'nodes' in t or 'no data source' in t or 'twain error' in t:
        return ('扫描仪没有连上或没就绪：检查电源和 USB 线，'
                '关掉其它正在使用扫描仪的程序，然后重新点「开始扫描」。')
    if 'maxconnections' in t or 'max connections' in t:
        return '扫描仪连接被占满（上一次扫描没退干净）。等 10 秒再点，或拔插一次 USB。'
    if 'offline' in t or '脱机' in t or 'unavailable' in t or 'not available' in t:
        return '扫描仪脱机：按一下机身电源键唤醒，等指示灯回到就绪状态再试。'
    if diag == 'jam':
        return '扫描仪卡纸：打开进纸盖取出卡住的纸，码齐后重试。'
    if diag == 'no_output':
        return '扫描仪没有响应：确认电源和 USB，或先关掉别的扫描软件。'
    if diag == 'stall':
        return '扫描开始后设备不动了：多半是休眠或卡纸，处理后重试。'
    if diag == 'timeout':
        return '扫描超时：这一叠可能太多页，分两次扫会更稳。'
    if 'no pages' in t or 'no scanned pages' in t or '0 page' in t:
        return '没有进纸：确认纸放进了进纸器，两边导纸板贴住纸边。'
    return '没有扫描到页面。'


def do_scan(opts):
    """一次完整流程：扫描 → 逐页上传 → commit 建批。"""
    try:
        state_set(stage='scanning', pages=0, uploaded=0, batch_id='',
                  message='正在扫描…', error='', started=time.time())
        naps2 = find_naps2()
        if not naps2:
            state_set(stage='error', error='没找到 NAPS2。请在下面「扫描仪」里填 NAPS2.Console.exe 的路径。')
            return
        stamp = time.strftime('%Y%m%d-%H%M%S')
        out_dir = os.path.join(SCAN_ROOT, stamp)
        os.makedirs(out_dir, exist_ok=True)
        base_name = 'scan.png'
        target = os.path.join(out_dir, base_name)

        so = CFG.get('scan') or {}
        cmd = [naps2, '-o', target, '-f', '-v', '--noprofile']
        if so.get('driver'):
            cmd += ['--driver', str(so['driver'])]
        if so.get('device'):
            cmd += ['--device', str(so['device'])]
        for key, flag in (('source', '--source'), ('pagesize', '--pagesize'),
                          ('dpi', '--dpi'), ('bitdepth', '--bitdepth')):
            if so.get(key) not in (None, ''):
                cmd += [flag, str(so[key])]

        ok, diag, text = _run_naps2(cmd, out_dir, base_name)
        pages = _collect_pages(out_dir, base_name)
        state_set(pages=len(pages))

        if not pages:
            detail = ('\n\n技术详情：' + text[-400:]) if text else ''
            state_set(stage='error', error=explain_failure(diag, text) + detail)
            return

        # ---- 上传：一页一个请求，失败重试 2 次，只重传这一页
        session = uuid.uuid4().hex
        state_set(stage='uploading', message='正在上传 %d 页…' % len(pages))
        for i, p in enumerate(pages, 1):
            body = open(p, 'rb').read()
            ok_up = False
            err = ''
            for attempt in (1, 2, 3):
                ok_up, err = api('/api/ingest/page', body, {
                    'X-QJ-Session': session,
                    'X-QJ-Filename': os.path.basename(p),
                    'Content-Type': 'application/octet-stream',
                }, method='POST', timeout=120)
                if ok_up:
                    break
                time.sleep(1.5 * attempt)
            if not ok_up:
                state_set(stage='error', uploaded=i - 1,
                          error='第 %d 页上传失败：%s' % (i, err))
                return
            state_set(uploaded=i, message='已上传 %d / %d 页' % (i, len(pages)))

        # ---- 建批
        state_set(stage='committing', message='正在交给主机处理…')
        payload = {
            'session': session,
            'subject': opts.get('subject') or 'english_essay',
            'task_id': opts.get('task_id') or '',
            'exercise_id': opts.get('exercise_id') or '',
            'teacher_note': opts.get('teacher_note') or '',
            'max_score': opts.get('max_score'),
            'auto': True,
            'collect': bool(opts.get('collect')),
        }
        ok_c, resp = api('/api/ingest/commit', payload, method='POST', timeout=180)
        if not ok_c:
            state_set(stage='error', error='主机拒收了这一批：%s' % resp)
            return
        # 传完就把本地图片删掉：扫描点不是存放处，主机那边已经有一份了
        shutil.rmtree(out_dir, ignore_errors=True)
        batch_id = (resp.get('batch') or {}).get('id') or ''
        note = resp.get('batch_error') or ''
        state_set(stage='done', batch_id=batch_id, message=note or '已完成，等待主机批阅',
                  error='')
    except Exception as e:
        state_set(stage='error', error='扫描流程出错：%s' % e)


def start_scan(opts):
    if not _SCAN_LOCK.acquire(blocking=False):
        return False, '上一次扫描还没结束'
    def _run():
        try:
            do_scan(opts)
        finally:
            _SCAN_LOCK.release()
    threading.Thread(target=_run, daemon=True).start()
    return True, ''


# ---------------------------- 页面 ----------------------------

PAGE = r"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>__POINT__ · 扫描点</title>
<style>
:root{--text:#111827;--text2:#4b5563;--muted:#5f6b7a;--line:#e5e7eb;--primary:#2563eb;
--ok:#047857;--warn:#8a5300;--danger:#b91c1c;}
*{box-sizing:border-box}
body{margin:0;background:#f5f7fa;color:var(--text);font-family:-apple-system,"Segoe UI","Microsoft YaHei",sans-serif;font-size:14px;}
.top{background:#1e293b;color:#fff;padding:14px 20px;display:flex;align-items:center;gap:12px;}
.top h1{font-size:15px;margin:0;font-weight:600;letter-spacing:-.2px;}
.top .host{font-size:12px;opacity:.75;margin-left:auto;}
.wrap{max-width:760px;margin:0 auto;padding:20px 16px 80px;}
.card{background:#fff;border:1px solid var(--line);border-radius:12px;padding:18px;margin-bottom:14px;}
.card h2{font-size:13px;margin:0 0 12px;color:var(--text2);font-weight:600;letter-spacing:.2px;}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;}
label{display:block;font-size:12.5px;color:var(--text2);margin-bottom:5px;}
select,input[type=text],input[type=number],textarea{
width:100%;padding:9px 10px;border:1px solid var(--line);border-radius:8px;font-size:13.5px;
font-family:inherit;color:var(--text);background:#fff;}
select:focus,input:focus,textarea:focus{outline:2px solid #bfdbfe;outline-offset:1px;border-color:#93c5fd;}
textarea{resize:vertical;min-height:56px;}
.row{display:flex;gap:10px;align-items:center;flex-wrap:wrap;}
.btn{padding:11px 18px;border-radius:9px;border:1px solid var(--line);background:#fff;color:var(--text);
font-size:13.5px;font-family:inherit;cursor:pointer;min-height:42px;}
.btn:hover{background:#f9fafb;}
.btn.primary{background:var(--primary);border-color:var(--primary);color:#fff;font-weight:600;font-size:15px;
padding:15px 20px;width:100%;min-height:54px;}
.btn.primary:hover{background:#1d4ed8;}
.btn.primary:disabled{background:#cbd5e1;border-color:#cbd5e1;cursor:not-allowed;}
.status{display:flex;gap:10px;align-items:center;font-size:13px;color:var(--text2);}
.dot{width:9px;height:9px;border-radius:50%;background:#cbd5e1;flex:none;}
.dot.ok{background:#22c55e;} .dot.bad{background:#ef4444;} .dot.run{background:#f59e0b;}
#msg{margin-top:10px;font-size:13px;line-height:1.7;}
#msg.err{color:var(--danger);} #msg.done{color:var(--ok);}
.bar{height:7px;border-radius:99px;background:#eef2f7;overflow:hidden;margin-top:12px;}
.bar>i{display:block;height:100%;width:0;background:var(--primary);transition:width .3s;}
.res{margin-top:12px;font-size:13px;line-height:1.9;color:var(--text2);}
.res b{color:var(--text);}
details summary{cursor:pointer;font-size:12.5px;color:var(--text2);padding:6px 0;min-height:30px;}
.hint{font-size:12px;color:var(--muted);line-height:1.7;margin-top:8px;}
.warnbox{padding:10px 12px;border-radius:8px;background:#fffbeb;border:1px solid #fde68a;color:var(--warn);
font-size:12.5px;line-height:1.7;margin-top:10px;}
</style></head>
<body>
<div class="top">
  <h1>扫描点 · __POINT__</h1>
  <span class="host" id="host"></span>
</div>
<div class="wrap">

  <div class="card">
    <div class="status"><span class="dot" id="dot"></span><span id="connText">正在连接主机…</span></div>
    <div id="msg"></div>
    <div class="bar" id="barWrap" style="display:none"><i id="bar"></i></div>
    <div class="res" id="res" style="display:none"></div>
  </div>

  <div class="card">
    <h2>这一批怎么处理</h2>
    <div class="grid">
      <div><label for="subject">学科</label><select id="subject"></select></div>
      <div id="taskWrap"><label for="task">题目</label><select id="task"></select></div>
      <div id="exWrap"><label for="ex">要点练习（可选）</label><select id="ex"></select></div>
      <div><label for="maxScore">满分</label><input type="number" id="maxScore" min="1" max="200" value="25"></div>
    </div>
    <div style="margin-top:12px"><label for="note">批改备注（可选，整批生效）</label>
      <textarea id="note" maxlength="300" placeholder="例如：本次重点看论证是否完整"></textarea></div>
    <div class="row" style="margin-top:12px">
      <label style="display:flex;align-items:center;gap:7px;margin:0;font-size:13px;color:var(--text)">
        <input type="checkbox" id="collect" style="width:16px;height:16px;margin:0">
        只入库不批阅
      </label>
    </div>
    <div style="margin-top:16px"><button class="btn primary" id="scanBtn">开始扫描</button></div>
    <div class="hint">放好纸再点。扫完会自动上传并交给主机处理，这边不用管。</div>
  </div>

  <div class="card">
    <details>
      <summary>扫描仪设置（一般不用动）</summary>
      <div class="grid" style="margin-top:10px">
        <div style="grid-column:1/-1"><label for="naps2">NAPS2.Console.exe 路径</label>
          <input type="text" id="naps2" placeholder="自动探测；找不到时手填"></div>
        <div><label for="driver">驱动</label><select id="driver">
          <option value="">默认</option><option value="twain">TWAIN</option><option value="wia">WIA</option>
        </select></div>
        <div><label for="device">设备名</label><input type="text" id="device" placeholder="留空用默认设备"></div>
        <div><label for="source">进纸方式</label><select id="source">
          <option value="feeder">进纸器（ADF）</option><option value="glass">平板（盖板）</option>
        </select></div>
        <div><label for="pagesize">纸张</label><select id="pagesize">
          <option value="a4">A4</option><option value="letter">Letter</option>
          <option value="a3">A3</option><option value="auto">自动</option>
        </select></div>
        <div><label for="dpi">分辨率</label><select id="dpi">
          <option value="200">200 dpi</option><option value="300">300 dpi</option><option value="400">400 dpi</option>
        </select></div>
        <div><label for="bitdepth">色彩</label><select id="bitdepth">
          <option value="color">彩色</option><option value="gray">灰度</option><option value="blackwhite">黑白</option>
        </select></div>
      </div>
      <div class="row" style="margin-top:12px"><button class="btn" id="saveCfg">保存设置</button></div>
      <div class="hint">分辨率低于 300 时手写字容易糊，直接影响批阅准确率。</div>
      <div class="warnbox" id="naps2Warn" style="display:none"></div>
    </details>
  </div>
</div>

<script>
(function () {
  var poll = null, batchTimer = null;
  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }

  function setConn(ok, text) {
    $('dot').className = 'dot ' + (ok ? 'ok' : 'bad');
    $('connText').textContent = text;
  }

  function fillSubjects(g) {
    var sel = $('subject'); sel.innerHTML = '';
    (g.subjects || []).forEach(function (s) {
      if (!s.auto_ok) return;
      var o = document.createElement('option');
      o.value = s.key; o.textContent = s.name; o.dataset.max = s.max_score;
      o.dataset.kind = s.kind;
      sel.appendChild(o);
    });
    syncSubject();
  }
  function syncSubject() {
    var o = $('subject').selectedOptions[0];
    if (!o) return;
    var isEssay = o.dataset.kind === 'essay';
    $('taskWrap').style.display = isEssay ? '' : 'none';
    $('exWrap').style.display = isEssay ? 'none' : '';
    if (o.dataset.max && !$('maxScore').dataset.touched) $('maxScore').value = o.dataset.max;
  }

  window.__OPTIONS = null;
  function loadOptions() {
    fetch('/api/options').then(function (r) { return r.json(); }).then(function (j) {
      if (j.error) { setConn(false, j.error); return; }
      setConn(true, '已连接主机 · ' + (j.server_time || ''));
      $('host').textContent = j.host || '';
      window.__OPTIONS = j;
      fillSubjects(j.grading || {});
      var t = $('task'); t.innerHTML = '<option value="">不指定题目</option>';
      (j.tasks || []).forEach(function (x) {
        var o = document.createElement('option'); o.value = x.id; o.textContent = x.title; t.appendChild(o);
      });
      var e = $('ex'); e.innerHTML = '<option value="">不用要点表（按学科标准整体给分）</option>';
      (j.exercises || []).forEach(function (x) {
        var o = document.createElement('option'); o.value = x.id; o.textContent = x.name || x.title || x.id; e.appendChild(o);
      });
    }).catch(function () { setConn(false, '连不上主机'); });
  }

  function loadState() {
    fetch('/api/state').then(function (r) { return r.json(); }).then(function (j) {
      $('naps2').value = j.naps2_path || '';
      var s = j.scan || {};
      if (s.driver) $('driver').value = s.driver;
      $('device').value = s.device || '';
      if (s.source) $('source').value = s.source;
      if (s.pagesize) $('pagesize').value = s.pagesize;
      if (s.dpi) $('dpi').value = String(s.dpi);
      if (s.bitdepth) $('bitdepth').value = s.bitdepth;
      if (!j.naps2_ok) {
        var w = $('naps2Warn');
        w.style.display = ''; w.textContent = '没有找到 NAPS2。先装好 NAPS2，再把 NAPS2.Console.exe 的完整路径填在上面并保存。';
      }
      if (j.progress) renderProgress(j.progress);
    });
  }

  function renderProgress(p) {
    var btn = $('scanBtn');
    var busy = p.stage === 'scanning' || p.stage === 'uploading' || p.stage === 'committing';
    btn.disabled = busy;
    btn.textContent = busy ? '正在处理…' : '开始扫描';
    $('barWrap').style.display = (p.stage === 'idle') ? 'none' : '';
    var pct = 0;
    if (p.stage === 'scanning') { $('bar').style.width = '8%'; }
    else if (p.stage === 'uploading' && p.pages) { pct = Math.round(p.uploaded / p.pages * 92); $('bar').style.width = pct + '%'; }
    else if (p.stage === 'committing') { $('bar').style.width = '96%'; }
    else if (p.stage === 'done') { $('bar').style.width = '100%'; }

    var m = $('msg'); m.className = '';
    if (p.stage === 'error') { m.className = 'err'; m.textContent = p.error || '出错了'; }
    else if (p.stage === 'done') { m.className = 'done'; m.textContent = p.message || '已完成'; }
    else if (p.stage === 'idle') { m.textContent = ''; }
    else { m.textContent = p.message || ''; }

    if (p.batch_id) startBatchPoll(p.batch_id);
  }

  function startBatchPoll(id) {
    if (batchTimer) return;
    var tick = function () {
      fetch('/api/batch/' + id).then(function (r) { return r.json(); }).then(function (j) {
        if (j.error) return;
        var job = j.job || {};
        var total = job.total || job.count || 0, done = job.done || 0;
        var box = $('res'); box.style.display = '';
        var lines = [];
        if (j.collect) lines.push('<b>已登记入库</b>，这一批没有送模型批阅。');
        else {
          lines.push('批阅进度：<b>' + done + ' / ' + total + '</b>' +
            (job.status ? '（' + esc(job.status) + '）' : ''));
          if ((j.results || []).length) {
            var avg = 0, n = 0;
            (j.results || []).forEach(function (r) { if (typeof r.score === 'number') { avg += r.score; n++; } });
            if (n) lines.push('平均分：<b>' + (avg / n).toFixed(1) + '</b>');
          }
          if ((j.failed || []).length) lines.push('失败 <b>' + j.failed.length + '</b> 份，可在后台重试。');
        }
        if (job.status === 'done' || job.status === 'error') {
          box.innerHTML = lines.join('<br>');
          clearInterval(batchTimer); batchTimer = null;
          return;
        }
        box.innerHTML = lines.join('<br>');
      }).catch(function () {});
    };
    batchTimer = setInterval(tick, 2000); tick();
  }

  function scan() {
    var payload = {
      subject: $('subject').value,
      task_id: $('task').value,
      exercise_id: $('ex').value,
      teacher_note: $('note').value,
      max_score: $('maxScore').value || null,
      collect: $('collect').checked
    };
    if ((payload.exercise_id || '') && !payload.collect) {
      // 有要点表时满分以练习为准，主机那边会覆盖，这里不必纠
    }
    $('res').style.display = 'none';
    fetch('/api/scan', { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload) })
      .then(function (r) { return r.json(); })
      .then(function (j) { if (j.error) { $('msg').className = 'err'; $('msg').textContent = j.error; } });
  }

  $('maxScore').addEventListener('input', function () { this.dataset.touched = '1'; });
  $('subject').addEventListener('change', syncSubject);
  $('scanBtn').addEventListener('click', scan);
  $('saveCfg').addEventListener('click', function () {
    fetch('/api/config', { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        naps2_path: $('naps2').value.trim(),
        scan: { driver: $('driver').value, device: $('device').value.trim(),
                source: $('source').value, pagesize: $('pagesize').value,
                dpi: $('dpi').value, bitdepth: $('bitdepth').value }
      })
    }).then(function (r) { return r.json(); }).then(function (j) {
      $('msg').className = j.error ? 'err' : 'done';
      $('msg').textContent = j.error || '设置已保存';
      if (!j.error) { $('naps2Warn').style.display = 'none'; }
    });
  });

  loadState(); loadOptions();
  poll = setInterval(function () {
    fetch('/api/progress').then(function (r) { return r.json(); }).then(renderProgress).catch(function () {});
  }, 1200);
})();
</script>
</body></html>
"""


class Handler(BaseHTTPRequestHandler):
    server_version = 'QuickJudgeScan/1.0'

    def log_message(self, fmt, *args):
        pass    # 不要把每个轮询都打到控制台

    # ---------- 工具 ----------
    def _send(self, code, body, ctype='application/json; charset=utf-8'):
        if isinstance(body, str):
            body = body.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        try:
            self.wfile.write(body)
        except Exception:
            pass

    def _json(self, obj, code=200):
        self._send(code, json.dumps(obj, ensure_ascii=False))

    def _body(self):
        try:
            n = int(self.headers.get('Content-Length') or 0)
        except ValueError:
            n = 0
        return self.rfile.read(n) if n > 0 else b''

    def _json_body(self):
        try:
            return json.loads(self._body().decode('utf-8') or '{}')
        except Exception:
            return {}

    # ---------- 路由 ----------
    def do_GET(self):
        path = urlparse(self.path).path
        if path == '/':
            html = PAGE.replace('__POINT__', (CFG.get('point_name') or '未命名'))
            return self._send(200, html, 'text/html; charset=utf-8')
        if path == '/api/state':
            naps2 = find_naps2()
            return self._json({
                'point_name': CFG.get('point_name'), 'host': CFG.get('host'),
                'naps2_path': naps2, 'naps2_ok': bool(naps2),
                'scan': CFG.get('scan'), 'progress': state_get(),
            })
        if path == '/api/progress':
            return self._json(state_get())
        if path.startswith('/api/batch/'):
            ok, data = api(path)
            return self._json(data if ok else {'error': data}, 200 if ok else 502)
        if path == '/api/options':
            ok, data = api('/api/ingest/options')
            if not ok:
                return self._json({'error': data if isinstance(data, str) else '主机没有响应'}, 502)
            data['host'] = CFG.get('host')
            return self._json(data)
        return self._json({'error': 'not found'}, 404)

    def do_POST(self):
        path = urlparse(self.path).path
        if path == '/api/scan':
            ok, err = start_scan(self._json_body())
            return self._json({'ok': True} if ok else {'error': err}, 200 if ok else 409)
        if path == '/api/config':
            req = self._json_body()
            if 'naps2_path' in req:
                CFG['naps2_path'] = str(req.get('naps2_path') or '').strip()
            if isinstance(req.get('scan'), dict):
                for k in ('driver', 'device', 'source', 'pagesize', 'dpi', 'bitdepth'):
                    if k in req['scan']:
                        CFG['scan'][k] = req['scan'][k]
            if CFG['naps2_path'] and not os.path.isfile(CFG['naps2_path']):
                return self._json({'error': '这个路径下没有 NAPS2.Console.exe'}, 400)
            save_config(CFG)
            return self._json({'ok': True})
        return self._json({'error': 'not found'}, 404)


def main():
    if not CFG.get('host') or not CFG.get('token'):
        print('config.json 里缺 host 或 token —— 请用主机生成的安装包重新解压安装。')
        return 1
    port = int(CFG.get('port') or 5101)
    srv = None
    for p in (port, port + 1, port + 2, port + 3):
        try:
            srv = ThreadingHTTPServer(('127.0.0.1', p), Handler)
            port = p
            break
        except OSError:
            continue
    if srv is None:
        print('5101~5104 都被占用了，关掉占用的程序再试。')
        return 1
    print('QuickJudge 扫描点：%s' % (CFG.get('point_name') or ''))
    print('本机页面：http://127.0.0.1:%d/' % port)
    print('主机：%s' % CFG.get('host'))
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    return 0


if __name__ == '__main__':
    sys.exit(main())
