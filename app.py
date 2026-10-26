from flask import Flask, render_template, request, jsonify, send_from_directory, send_file, session, Response
from flask_cors import CORS
import os
import json
import re
import io
import csv
import zipfile
import uuid
import random
import string
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # dotenv是可选的

# 大模型接口：优先使用 LLM_*（支持 CodingPlan 等 OpenAI 兼容），未设置时回退到 MINIMAX_*
_llm_key = os.getenv("LLM_API_KEY", "").strip() or os.getenv("MINIMAX_API_KEY", "").strip()
_llm_base = os.getenv("LLM_BASE_URL", "").strip() or os.getenv("MINIMAX_BASE_URL", "https://api.minimaxi.com/v1").strip()
_llm_model = os.getenv("LLM_MODEL", "").strip() or os.getenv("MINIMAX_MODEL", "MiniMax-M2.5").strip()
minimax_api_key = _llm_key
minimax_base_url = _llm_base if _llm_base else "https://api.minimaxi.com/v1"
minimax_model = _llm_model if _llm_model else "MiniMax-M2.5"

# 文本任务专用模型：承接纯文本任务（译文、班级汇总等），未配置时回退为主模型。
# 目的：把昂贵的视觉模型留给真正需要"看图"的批阅任务，文本任务用更便宜的模型。
_llm_text_model = os.getenv("LLM_TEXT_MODEL", "").strip()
llm_text_model = _llm_text_model if _llm_text_model else minimax_model

# ===== 识别精度档位 =====
# 试卷字迹差异极大（有的字只有 1.5mm 高、笔画 1 像素），单一模型档位必然有取舍。
# 这里给「看图」任务准备三档，由系统按卷面难度自动挑，老师无需在扫描时做选择题：
#   fast     最便宜，适合印刷体/工整手写
#   standard 默认档（LLM_MODEL）
#   strong   难卷兜底（LLM_MODEL_STRONG），慢 2-4 倍、贵约 2 倍，仅在必要时升档
# LLM_QUALITY_MODE=auto（默认）表示：难卷自动升 strong，普通卷走 standard；
# 也可以固定为 fast/standard/strong，用于对比实验。
_llm_model_fast = os.getenv("LLM_MODEL_FAST", "").strip() or minimax_model
_llm_model_strong = os.getenv("LLM_MODEL_STRONG", "").strip() or "Qwen/Qwen3-VL-32B-Instruct"
LLM_QUALITY_MODE = (os.getenv("LLM_QUALITY_MODE", "auto") or "auto").strip().lower()
if LLM_QUALITY_MODE not in ("auto", "fast", "standard", "strong"):
    LLM_QUALITY_MODE = "auto"


def resolve_vision_model(mode=None, hard=False):
    """按精度档位选视觉模型。mode 为空时用 .env 的 LLM_QUALITY_MODE。

    auto：卷面被判为难卷（字太小/笔画太细）才升 strong，否则 standard。
    """
    m = (mode or LLM_QUALITY_MODE or "auto").strip().lower()
    if m == "fast":
        return _llm_model_fast
    if m == "strong":
        return _llm_model_strong
    if m == "standard":
        return minimax_model
    return _llm_model_strong if hard else minimax_model



def _text_model_kwargs(model_name=None):
    """Qwen3.5 / Qwen3.6 等默认开启思考模式的模型，在文本任务中显式关闭思考，
    避免思考链膨胀导致变慢、输出翻倍与额外计费；其他模型返回空 dict。"""
    name = (model_name or llm_text_model or '').lower()
    if 'qwen3.5' in name or 'qwen3-5' in name or 'qwen3.6' in name or 'qwen3-6' in name:
        return {'extra_body': {'enable_thinking': False}}
    return {}


def _vision_model_kwargs(model_name=None):
    """视觉模型的「关思考」参数只对部分模型有效，需按模型区分（实测结论）：

    · Qwen3-VL-32B-Instruct：默认走思考链，同一张图 141.7s；显式关闭后 6.2s；
    · Qwen3-VL-8B-Instruct / 30B-A3B-Instruct：本身就不思考，传该参数直接 400。
    """
    name = (model_name or minimax_model or '').lower()
    if 'qwen3-vl-32b-instruct' in name or 'thinking' in name:
        return {'extra_body': {'enable_thinking': False}}
    return {}


def _vision_chat(client, **kwargs):
    """调用视觉接口，带上关思考参数；若服务端不认这个参数则自动退回普通调用。"""
    kw = dict(kwargs)
    extra = _vision_model_kwargs(kw.get('model'))
    if extra:
        try:
            return client.chat.completions.create(**kw, **extra)
        except Exception as e:
            msg = str(e).lower()
            if 'enable_thinking' not in msg and 'extra_body' not in msg:
                raise
    return client.chat.completions.create(**kw)
# End block (15-20)

from openai import OpenAI
import base64
from PIL import Image
import io
import logging
import requests
from wsgiref.handlers import format_date_time
from time import mktime
import hashlib
import hmac
from urllib.parse import urlencode

app = Flask(__name__, static_folder='static', template_folder='templates')
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'scan-biyue-dev-secret-key-change-in-production')
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
CORS(app, supports_credentials=True)

# ===== 扫描运行时配置 =====
# 同一台扫描仪同一时刻只允许一次扫描会话：
# 并发扫描会触发 TWAIN "MaxConnections"，并把扫描仪拖入无响应状态（需拔插 USB 才能恢复）。
_SCAN_LOCK = threading.Lock()

# 档案库 / 临时区索引的「读-改-写」互斥锁。
# 老师可以分批上传并让多个批次同时跑，两批同时归档时如果各自 load 再 save，
# 后写的一方会把前一批的记录整体覆盖掉（表现：某个班的学生看不到自己的卷子）。
# 批量写入方（批次归档、临时区清理）必须持有对应锁。
_RECORDS_LOCK = threading.RLock()
_STAGING_LOCK = threading.RLock()
# 单次扫描最长等待秒数；超时后先温和终止、再强制结束进程树，确保扫描仪连接被释放。
SCAN_TIMEOUT_SECONDS = int(os.getenv('SCAN_TIMEOUT_SECONDS', '90') or 90)
# 扫描开始后多久仍"毫无输出"即判定设备无响应（健康的扫描仪会立刻打印 Beginning scan...）
SCAN_NO_OUTPUT_SECONDS = int(os.getenv('SCAN_NO_OUTPUT_SECONDS', '20') or 20)
# 已有输出但随后静默多久，判定为卡住（需大于最慢的一次真实扫描耗时）
SCAN_STALL_SECONDS = int(os.getenv('SCAN_STALL_SECONDS', '60') or 60)
# 扫描逻辑结束后，等待"进程退出或输出文件落盘"的宽限时间
SCAN_FINISH_GRACE_SECONDS = int(os.getenv('SCAN_FINISH_GRACE_SECONDS', '10') or 10)

LOGO_DIR = os.path.join(os.path.dirname(__file__), 'logo')

@app.route('/favicon.ico')
def favicon():
    """返回favicon图标（使用根目录 logo 文件夹中的 ico）"""
    return send_from_directory(LOGO_DIR, 'favicon.ico', mimetype='image/vnd.microsoft.icon')

@app.route('/logo/<path:filename>')
def logo_file(filename):
    """提供 logo 目录下的图片（如 4.png 横向 logo）"""
    return send_from_directory(LOGO_DIR, filename)

# 教师新手教程（静态 HTML，源文件位于项目内「使用说明」目录）
GUIDE_FILE = os.path.join(os.path.dirname(__file__), '使用说明', '教师新手教程.html')

@app.route('/guide')
def teacher_guide():
    """教师新手教程页面；无需登录即可查看，方便新老师上手"""
    if os.path.exists(GUIDE_FILE):
        return send_file(GUIDE_FILE, mimetype='text/html')
    return Response('教程文件未找到：使用说明/教师新手教程.html', status=404,
                    mimetype='text/plain; charset=utf-8')

@app.route('/api/me', methods=['GET'])
def get_current_user():
    """获取当前登录用户；未登录返回 401"""
    if not session.get('user_id'):
        return jsonify({'error': '未登录'}), 401
    return jsonify({
        'user': {
            'id': session.get('user_id'),
            'name': session.get('user_name', ''),
            'role': session.get('user_role', 'teacher')
        }
    })

@app.route('/api/users', methods=['GET'])
def get_users():
    """获取用户列表（id, name, role），不含密码。用于前端展示学号(姓名)等。需登录。"""
    if not session.get('user_id'):
        return jsonify({'error': '未登录'}), 401
    users = load_users()
    out = [{'id': str(u.get('id', '')), 'name': (u.get('name') or '').strip(), 'role': u.get('role', '')} for u in users]
    return jsonify({'users': out})

@app.route('/api/login', methods=['POST'])
def login():
    """登录：body { "id": "学号/工号", "password": "密码" }"""
    data = request.get_json() or {}
    uid = (data.get('id') or data.get('username') or '').strip()
    password = (data.get('password') or '').strip()
    if not uid or not password:
        return jsonify({'error': '请输入账号和密码'}), 400
    users = load_users()
    for u in users:
        if str(u.get('id', '')) == uid and str(u.get('password', '')) == password:
            session['user_id'] = u.get('id')
            session['user_name'] = u.get('name', uid)
            session['user_role'] = u.get('role', 'teacher')
            return jsonify({
                'user': {
                    'id': session['user_id'],
                    'name': session['user_name'],
                    'role': session['user_role']
                }
            })
    return jsonify({'error': '账号或密码错误'}), 401

@app.route('/api/logout', methods=['POST'])
def logout():
    """退出登录"""
    session.clear()
    return jsonify({'msg': '已退出'})


@app.route('/api/change_password', methods=['POST'])
def change_password():
    """当前登录用户自助改密。body { old_password, new_password }。学生、教师、管理员通用。"""
    uid = session.get('user_id')
    if not uid:
        return jsonify({'error': '未登录'}), 401
    data = request.get_json() or {}
    old = (data.get('old_password') or '').strip()
    new = (data.get('new_password') or '').strip()
    if not old or not new:
        return jsonify({'error': '请填写原密码和新密码'}), 400
    if len(new) < 4:
        return jsonify({'error': '新密码至少 4 位'}), 400
    if new == old:
        return jsonify({'error': '新密码不能与原密码相同'}), 400
    users = load_users()
    target = next((u for u in users if str(u.get('id', '')) == str(uid)), None)
    if not target:
        return jsonify({'error': '账号不存在，请联系管理员'}), 404
    if str(target.get('password', '')) != old:
        return jsonify({'error': '原密码不正确'}), 400
    target['password'] = new
    save_users(users)
    return jsonify({'msg': '密码已修改，下次登录请使用新密码'})


def _require_admin():
    """仅管理员可操作，否则返回 True 表示需拦截"""
    if session.get('user_role') != 'admin':
        return True
    return False


@app.route('/api/register_teacher', methods=['POST'])
def register_teacher():
    """教师注册：body { invite_code, id/account, password, name? }。账号建议手机号，密码无限制。"""
    data = request.get_json() or {}
    code = (data.get('invite_code') or data.get('code') or '').strip()
    account = (data.get('id') or data.get('account') or data.get('username') or '').strip()
    password = (data.get('password') or '').strip()
    name = (data.get('name') or account or '').strip()
    if not code:
        return jsonify({'error': '请输入邀请码'}), 400
    if not account:
        return jsonify({'error': '请输入账号（建议使用手机号）'}), 400
    if not password:
        return jsonify({'error': '请输入密码'}), 400
    entries = load_invite_codes()
    entry = next((e for e in entries if e.get('code') == code), None)
    if not entry:
        return jsonify({'error': '邀请码无效'}), 400
    used = int(entry.get('used_count', 0))
    max_uses = int(entry.get('max_uses', 1))
    if used >= max_uses:
        return jsonify({'error': '该邀请码已达使用上限'}), 400
    users = load_users()
    if any(str(u.get('id', '')) == account for u in users):
        return jsonify({'error': '该账号已存在'}), 400
    users.append({'id': account, 'name': name or account, 'role': 'teacher', 'password': password})
    save_users(users)
    entry['used_count'] = used + 1
    save_invite_codes(entries)
    return jsonify({'msg': '注册成功', 'user': {'id': account, 'name': name or account, 'role': 'teacher'}})


@app.route('/api/invite_code/validate', methods=['GET'])
def validate_invite_code():
    """校验邀请码是否有效（未达使用上限）"""
    code = (request.args.get('code') or '').strip()
    if not code:
        return jsonify({'valid': False})
    entries = load_invite_codes()
    entry = next((e for e in entries if e.get('code') == code), None)
    if not entry:
        return jsonify({'valid': False})
    used = int(entry.get('used_count', 0))
    max_uses = int(entry.get('max_uses', 1))
    return jsonify({'valid': used < max_uses})


@app.route('/api/admin/import_students', methods=['POST'])
def admin_import_students():
    """管理员：批量导入学生。可上传 CSV 或使用根目录 namelist.csv（use_source=namelist）。CSV 列：学生姓名, 登录账号。密码=登录账号。"""
    if _require_admin():
        return jsonify({'error': '仅管理员可操作'}), 403
    try:
        users = load_users()
        by_id = {str(u['id']): u for u in users}
        rows = []
        if request.files and request.files.get('file'):
            f = request.files['file']
            content = f.read().decode('utf-8-sig').strip()
            reader = csv.reader(io.StringIO(content))
            next(reader, None)  # skip header
            for row in reader:
                if len(row) >= 2:
                    rows.append((row[0].strip(), row[1].strip()))
        else:
            use = (request.form.get('use_source') or request.get_json(silent=True) or {}).get('use_source')
            if use == 'namelist' and os.path.exists(NAMELIST_CSV):
                with open(NAMELIST_CSV, 'r', encoding='utf-8-sig') as f:
                    reader = csv.reader(f)
                    next(reader, None)
                    for row in reader:
                        if len(row) >= 2:
                            rows.append((row[0].strip(), row[1].strip()))
            else:
                return jsonify({'error': '请上传 CSV 或使用 use_source=namelist 从根目录 namelist.csv 导入'}), 400
        added = 0
        for name, account in rows:
            if not account:
                continue
            sid = str(account)
            if sid not in by_id:
                by_id[sid] = {'id': sid, 'name': name or sid, 'role': 'student', 'password': sid}
                added += 1
            else:
                by_id[sid]['name'] = name or sid
        save_users(list(by_id.values()))
        for name, account in rows:
            if account and len(str(account).strip()) == 6 and str(account).strip().isdigit():
                ensure_class_for_student(str(account).strip(), name or account)
        return jsonify({'msg': f'导入完成，本次新增 {added} 人', 'added': added, 'total_rows': len(rows)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/admin/students', methods=['POST'])
def admin_add_student():
    """教师或管理员：单个添加学生。body { name, id }，密码=id。管理员与教师均可调用。"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    data = request.get_json() or {}
    name = (data.get('name') or '').strip()
    account = (data.get('id') or data.get('account') or '').strip()
    if not account:
        return jsonify({'error': '请填写登录账号（学号）'}), 400
    users = load_users()
    if any(str(u.get('id', '')) == account for u in users):
        return jsonify({'error': '该账号已存在'}), 400
    users.append({'id': account, 'name': name or account, 'role': 'student', 'password': account})
    save_users(users)
    ensure_class_for_student(account, name or account)
    return jsonify({'msg': '添加成功', 'user': {'id': account, 'name': name or account, 'role': 'student'}})


@app.route('/api/admin/reset_password', methods=['POST'])
def admin_reset_student_password():
    """管理员直接重置学生密码。body { student_id, new_password? }，不填新密码就重置为该学号。"""
    if _require_admin():
        return jsonify({'error': '仅管理员可操作'}), 403
    data = request.get_json() or {}
    sid = str(data.get('student_id') or data.get('id') or '').strip()
    new = str(data.get('new_password') or '').strip()
    if not sid:
        return jsonify({'error': '请填写学号'}), 400
    if new and len(new) < 4:
        return jsonify({'error': '新密码至少 4 位'}), 400
    users = load_users()
    target = next((u for u in users if str(u.get('id', '')) == sid), None)
    if not target:
        return jsonify({'error': '找不到账号 %s，请确认学号是否已导入' % sid}), 404
    if target.get('role') != 'student':
        return jsonify({'error': '该账号不是学生，这里只能重置学生密码'}), 400
    target['password'] = new or sid
    save_users(users)
    return jsonify({
        'msg': '已重置',
        'student_id': sid,
        'student_name': target.get('name', ''),
        'new_password': target['password'],
    })


@app.route('/api/admin/invite_codes', methods=['GET'])
def admin_list_invite_codes():
    """管理员：列出所有邀请码"""
    if _require_admin():
        return jsonify({'error': '仅管理员可操作'}), 403
    return jsonify({'codes': load_invite_codes()})


@app.route('/api/admin/invite_codes', methods=['POST'])
def admin_create_invite_code():
    """管理员：生成邀请码。body { max_uses: number, code?: 可选自定义码 }"""
    if _require_admin():
        return jsonify({'error': '仅管理员可操作'}), 403
    data = request.get_json() or {}
    max_uses = int(data.get('max_uses', 1))
    if max_uses < 1:
        max_uses = 1
    custom = (data.get('code') or '').strip()
    if custom:
        code = custom
        entries = load_invite_codes()
        if any(e.get('code') == code for e in entries):
            return jsonify({'error': '该邀请码已存在'}), 400
    else:
        code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=8))
        entries = load_invite_codes()
        while any(e.get('code') == code for e in entries):
            code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=8))
    entries.append({'code': code, 'max_uses': max_uses, 'used_count': 0, 'created_at': datetime.now().isoformat()})
    save_invite_codes(entries)
    return jsonify({'code': code, 'max_uses': max_uses, 'msg': '邀请码已生成'})


# 数据目录结构（集中管理所有数据）
DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')
CONFIG_DIR = os.path.join(DATA_DIR, 'config')
USERS_FILE = os.path.join(CONFIG_DIR, 'users.json')
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(CONFIG_DIR, exist_ok=True)

def load_users():
    """加载用户列表（id, name, role, password）。角色: student / teacher / admin"""
    if not os.path.exists(USERS_FILE):
        return []
    try:
        with open(USERS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return []


def save_users(users):
    """保存用户列表"""
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, ensure_ascii=False, indent=2)


def get_student_names_map():
    """返回学号 -> 姓名 的字典，仅学生角色。用于展示「学号(姓名)」时查姓名，学号存储保持纯净。姓名为空则值为空字符串。"""
    users = load_users()
    return {str(u.get('id', '')): (u.get('name') or '').strip() for u in users if u.get('role') == 'student'}


# 根目录 namelist.csv（学生姓名, 登录账号）
NAMELIST_CSV = os.path.join(os.path.dirname(__file__), 'namelist.csv')
# 邀请码配置
INVITE_CODES_FILE = os.path.join(CONFIG_DIR, 'invite_codes.json')


def load_invite_codes():
    """加载邀请码列表 [{ code, max_uses, used_count, created_at }]"""
    if not os.path.exists(INVITE_CODES_FILE):
        return []
    try:
        with open(INVITE_CODES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return []


def save_invite_codes(entries):
    with open(INVITE_CODES_FILE, 'w', encoding='utf-8') as f:
        json.dump(entries, f, ensure_ascii=False, indent=2)

# 存储识别结果和批阅结果
results_dir = os.path.join(DATA_DIR, 'results')
os.makedirs(results_dir, exist_ok=True)

# 历史记录存储（用于下次打开时恢复）
HISTORY_DIR = os.path.join(DATA_DIR, 'history')
os.makedirs(HISTORY_DIR, exist_ok=True)

# 扫描仪默认目录（根目录，用于读取日期文件夹）
SCAN_DIR = os.path.dirname(__file__)

# 班级管理目录
CLASSES_DIR = os.path.join(DATA_DIR, 'classes')
os.makedirs(CLASSES_DIR, exist_ok=True)

# 学生信息存储
STUDENTS_DIR = os.path.join(DATA_DIR, 'students')
os.makedirs(STUDENTS_DIR, exist_ok=True)

# 班级中心数据存储（记录每个学生的作文和报告）
CLASS_CENTER_DIR = os.path.join(DATA_DIR, 'class_center')
os.makedirs(CLASS_CENTER_DIR, exist_ok=True)

# 临时文件夹（用于存储处理后的图片）
TEMP_DIR = os.path.join(DATA_DIR, 'temp')
os.makedirs(TEMP_DIR, exist_ok=True)

# 导出文件目录
EXPORTS_DIR = os.path.join(DATA_DIR, 'exports')
os.makedirs(EXPORTS_DIR, exist_ok=True)

# 作文素材与教师上传素材目录
MATERIALS_DIR = os.path.join(DATA_DIR, 'materials')
MATERIALS_INDEX_FILE = os.path.join(MATERIALS_DIR, 'index.json')
os.makedirs(MATERIALS_DIR, exist_ok=True)

# 平台内编辑：作文素材使用说明、通用答题卡 HTML、答题卡题型模板（存 CONFIG_DIR）
COMPOSITION_MATERIALS_FILE = os.path.join(CONFIG_DIR, 'composition_materials.txt')
ANSWER_SHEET_HTML_FILE = os.path.join(CONFIG_DIR, 'answer_sheet.html')
ANSWER_SHEET_TEMPLATES_FILE = os.path.join(CONFIG_DIR, 'answer_sheet_templates.json')
# 组卷发布产生的任务列表（发布=创建任务）
TASKS_FILE = os.path.join(CONFIG_DIR, 'tasks.json')
# 按任务归类的试卷图片（扫码或上传时写入，便于批改与学号对应）
TASK_PAPERS_DIR = os.path.join(DATA_DIR, 'task_papers')
os.makedirs(TASK_PAPERS_DIR, exist_ok=True)
# 按任务持久化的批阅结果（答题情况导出用）
TASK_RESULTS_DIR = os.path.join(DATA_DIR, 'task_results')
os.makedirs(TASK_RESULTS_DIR, exist_ok=True)


def _save_task_grading_results(individual_reports):
    """将批阅结果中属于任务试卷的条目按任务持久化到 task_results/<task_id>.json，便于按任务导出答题情况。"""
    import re
    for file_path, report_data in individual_reports.items():
        m = re.match(r'^task/([^/]+)/(.+)$', file_path.strip())
        if not m:
            continue
        task_id, filename = m.group(1), m.group(2)
        result_file = os.path.join(TASK_RESULTS_DIR, f'{task_id}.json')
        data = {'results': {}, 'updated_at': datetime.now().isoformat()}
        if os.path.exists(result_file):
            try:
                with open(result_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if not isinstance(data.get('results'), dict):
                    data['results'] = {}
            except Exception:
                data['results'] = {}
        data['results'][filename] = {
            'student_id': report_data.get('student_id', ''),
            'report': report_data.get('report', ''),
            'status': report_data.get('status', ''),
            'filename': report_data.get('filename', filename),
        }
        data['updated_at'] = datetime.now().isoformat()
        try:
            with open(result_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass


def load_tasks():
    """加载任务列表。每项: id, title, class_names, deadline, student_ids, items, created_at, answer_sheet_html?, status"""
    if not os.path.exists(TASKS_FILE):
        return []
    try:
        with open(TASKS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return []


def save_tasks(tasks):
    """保存任务列表"""
    with open(TASKS_FILE, 'w', encoding='utf-8') as f:
        json.dump(tasks, f, ensure_ascii=False, indent=2)


def decode_qr_from_image(image_bytes):
    """从图片字节中解析二维码内容，返回解码字符串或 None。用于试卷扫码自动归类。"""
    if not image_bytes or len(image_bytes) < 100:
        return None
    try:
        import numpy as np
        import cv2
        arr = np.frombuffer(image_bytes, dtype=np.uint8)
        img = cv2.imdecode(arr, cv2.IMREAD_COLOR)
        if img is None:
            return None
        det = cv2.QRCodeDetector()
        decoded, _, _ = det.detectAndDecode(img)
        if decoded and isinstance(decoded, str) and decoded.strip():
            return decoded.strip()
    except Exception:
        pass
    return None


def _save_paper_to_task(task_id, image_bytes, ext='.png'):
    """将试卷图片写入任务目录，返回保存后的文件名。"""
    task_dir = os.path.join(TASK_PAPERS_DIR, task_id)
    os.makedirs(task_dir, exist_ok=True)
    name = str(uuid.uuid4()) + ext
    path = os.path.join(task_dir, name)
    with open(path, 'wb') as f:
        f.write(image_bytes)
    return name


# 提示词模板存储
PROMPT_TEMPLATE_FILE = os.path.join(CONFIG_DIR, 'prompt_template.txt')
SUBJECT_CONFIG_FILE = os.path.join(CONFIG_DIR, 'subject_config.json')
DEFAULT_PROMPT_TEMPLATE = """请对以下内容进行详细批阅，给出：
1. 总体评分
2. 优点分析
3. 问题指出
4. 改进建议
5. 具体修改意见

内容：
{essay_text}
"""

# ================= 批阅方案（学科 × 题型）=================
# 把「评分依据」从代码里搬出来：每份卷在扫描端选一个方案，批阅时按方案拼提示词。
# 方案里可用的占位符：{max_score}（该块满分）、{subject}（方案名）。
GRADING_CONFIG_FILE = os.path.join(CONFIG_DIR, 'grading_config.json')

DEFAULT_GRADING_CONFIG = {
    'current': 'general',
    'subjects': {
        'general': {
            'name': '通用片段 / 整页',
            'max_score': 100,
            'prompt': (
                '这是一份学生作答的片段，可能只是某几道题，也可能是整页。\n'
                '请先完整识别其中的手写内容，再按学科常规标准评价。\n'
                '评分要点：\n'
                '- 答案或结论是否正确，推导/步骤是否完整；\n'
                '- 表述是否清楚，书写与格式是否规范；\n'
                '- 开放题重点看思路是否成立、论证是否站得住。'
            ),
        },
        'chinese_essay': {
            'name': '语文作文',
            'max_score': 50,
            'prompt': (
                '这是一篇学生手写作文。请按中考/高考作文通用标准评价：\n'
                '评分要点：\n'
                '- 立意：是否明确、是否切题、有无真情实感；\n'
                '- 结构：层次是否清晰，详略是否得当，首尾是否呼应；\n'
                '- 语言：用词是否准确生动，句式是否有变化，有无病句错别字；\n'
                '- 书写：字迹是否工整、卷面是否整洁。'
            ),
        },
        'english_essay': {
            'name': '英语作文',
            'max_score': 25,
            'prompt': (
                'This is a handwritten English composition. 请按英语写作通用标准评价：\n'
                '评分要点：\n'
                '- 内容：要点是否齐全，是否扣题，细节是否充分；\n'
                '- 结构：段落划分、衔接词使用是否合理；\n'
                '- 语言：语法、时态、拼写、用词是否准确，句式是否有变化；\n'
                '- 书写：字迹是否清晰工整。'
            ),
        },
        'math_solution': {
            'name': '数学 / 理科解答题',
            'max_score': 12,
            'prompt': (
                '这是一道（或几道）学生手写的理科解答题。请按步骤给分：\n'
                '评分要点：\n'
                '- 关键步骤是否写出，公式/定理使用是否正确；\n'
                '- 计算过程与中间结果是否正确；\n'
                '- 最终答案是否正确，单位与作答格式是否规范；\n'
                '- 若思路正确但计算失误，应指出失误点，不整体否定。'
            ),
        },
        'dictation': {
            'name': '默写 / 字词 / 填空',
            'max_score': 10,
            'prompt': (
                '这是一份默写、字词或填空类作答。请逐条核对：\n'
                '评分要点：\n'
                '- 逐空/逐条判断正误，指出具体错字、漏字、多字；\n'
                '- 注意形近字、同音字错误，以及标点遗漏；\n'
                '- 在 issues 中逐条列出错误项，写清「原写」与「应为」。'
            ),
        },
    },
}


def load_grading_config():
    """读取批阅方案配置；文件缺失时用内置默认值，并保证 current 指向存在的方案。"""
    cfg = json.loads(json.dumps(DEFAULT_GRADING_CONFIG))  # 深拷贝，避免污染默认值
    try:
        if os.path.exists(GRADING_CONFIG_FILE):
            with open(GRADING_CONFIG_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
            subs = saved.get('subjects')
            if isinstance(subs, dict):
                for key, val in subs.items():
                    if not isinstance(val, dict):
                        continue
                    base = cfg['subjects'].get(key, {})
                    merged = dict(base)
                    merged.update({k: v for k, v in val.items() if v not in (None, '')})
                    merged.setdefault('name', key)
                    merged.setdefault('max_score', 100)
                    merged.setdefault('prompt', '')
                    cfg['subjects'][key] = merged
            if saved.get('current') in cfg['subjects']:
                cfg['current'] = saved['current']
    except Exception as e:
        print('读取批阅方案配置失败，使用默认值: %s' % e)
    return cfg


def save_grading_config(cfg):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(GRADING_CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def _get_grading_subject(key=None):
    """取某个批阅方案；key 为空时用配置里的 current，找不到再回退 general。"""
    cfg = load_grading_config()
    subjects = cfg.get('subjects') or {}
    k = str(key or '').strip() or cfg.get('current') or 'general'
    subj = subjects.get(k) or subjects.get('general') or {}
    out = dict(subj)
    out['key'] = k if k in subjects else 'general'
    out['name'] = out.get('name') or '通用'
    try:
        out['max_score'] = float(out.get('max_score') or 100)
    except (TypeError, ValueError):
        out['max_score'] = 100.0
    out['prompt'] = out.get('prompt') or ''
    return out


def grading_subject_options():
    """给前端用的方案列表（只暴露必要字段）。"""
    cfg = load_grading_config()
    out = []
    for key, val in (cfg.get('subjects') or {}).items():
        out.append({'key': key, 'name': val.get('name') or key,
                    'max_score': val.get('max_score') or 100})
    return {'current': cfg.get('current') or 'general', 'subjects': out}

# PNG压缩质量配置存储
PNG_QUALITY_CONFIG_FILE = os.path.join(CONFIG_DIR, 'png_quality_config.json')
DEFAULT_PNG_QUALITY = 95  # 默认压缩质量（0-100，100为最高质量）

# NAPS2扫描仪配置存储
NAPS2_CONFIG_FILE = os.path.join(CONFIG_DIR, 'naps2_config.json')
# 本机实际安装路径（便携版，2026-10-16 部署）；可在 data/config/naps2_config.json 中覆盖
DEFAULT_NAPS2_PATH = r"D:\Tools\NAPS2\App\NAPS2.Console.exe"

# 扫描输出目录配置
SCAN_OUTPUT_CONFIG_FILE = os.path.join(CONFIG_DIR, 'scan_output_config.json')
# 本机扫描输出目录（2026-10-16 本地部署）；可在 data/config/scan_output_config.json 中修改
DEFAULT_SCAN_OUTPUT_DIR = r"D:\ScanTests"

# 扫描仪高级设置（图像模式、单双面等，与打印机/扫描相关保留）
SCANNER_ADVANCED_CONFIG_FILE = os.path.join(CONFIG_DIR, 'scanner_advanced_config.json')
DEFAULT_SCANNER_IMAGE_MODE = 'color'   # grayscale | color
DEFAULT_SCANNER_SCAN_TYPE = 'double'   # single | double

# 讯飞OCR配置（已停用：讯飞OCR服务已过期，文字识别统一走 llm_ocr_recognize 视觉大模型）
XF_APP_ID = os.getenv("XF_APP_ID", "6c8f82e6")
XF_API_SECRET = os.getenv("XF_API_SECRET", "NGQ2ZTg3NWY3NDkxMTMyYWJlYWQwNTJm")
XF_API_KEY = os.getenv("XF_API_KEY", "73ed4e3a092cc935af22e2c420bd9cd8")
XF_OCR_URL = 'https://api.xf-yun.com/v1/private/sf8e6aca1'

class AssembleHeaderException(Exception):
    def __init__(self, msg):
        self.message = msg

class Url:
    def __init__(self, host, path, schema):
        self.host = host
        self.path = path
        self.schema = schema

def parse_url(requset_url):
    stidx = requset_url.index("://")
    host = requset_url[stidx + 3:]
    schema = requset_url[:stidx + 3]
    edidx = host.index("/")
    if edidx <= 0:
        raise AssembleHeaderException("invalid request url:" + requset_url)
    path = host[edidx:]
    host = host[:edidx]
    return Url(host, path, schema)

def assemble_ws_auth_url(requset_url, method="POST", api_key="", api_secret=""):
    """构建讯飞OCR认证URL"""
    u = parse_url(requset_url)
    host = u.host
    path = u.path
    now = datetime.now()
    date = format_date_time(mktime(now.timetuple()))
    signature_origin = "host: {}\ndate: {}\n{} {} HTTP/1.1".format(host, date, method, path)
    signature_sha = hmac.new(api_secret.encode('utf-8'), signature_origin.encode('utf-8'),
                             digestmod=hashlib.sha256).digest()
    signature_sha = base64.b64encode(signature_sha).decode(encoding='utf-8')
    authorization_origin = "api_key=\"%s\", algorithm=\"%s\", headers=\"%s\", signature=\"%s\"" % (
        api_key, "hmac-sha256", "host date request-line", signature_sha)
    authorization = base64.b64encode(authorization_origin.encode('utf-8')).decode(encoding='utf-8')
    values = {
        "host": host,
        "date": date,
        "authorization": authorization
    }
    return requset_url + "?" + urlencode(values)

def llm_ocr_recognize(image_path_or_bytes, logger=None):
    """
    使用大模型视觉接口（OpenAI 兼容，如硅基流动 Qwen3-VL）进行文字识别。
    替代已停用的讯飞 OCR（2026-10-15 切换）。

    Args:
        image_path_or_bytes: 图片路径（str）、图片字节数据（bytes）或 PIL Image 对象
        logger: 日志记录器（可选）

    Returns:
        str: 识别出的文本内容，失败返回空字符串
    """
    if not minimax_api_key:
        if logger:
            logger.warning("未配置 LLM_API_KEY，无法调用大模型识图")
        return ''
    try:
        # 统一取到 PIL Image
        if isinstance(image_path_or_bytes, str):
            image = Image.open(image_path_or_bytes)
        elif isinstance(image_path_or_bytes, bytes):
            image = Image.open(io.BytesIO(image_path_or_bytes))
        elif hasattr(image_path_or_bytes, 'save'):
            image = image_path_or_bytes
        else:
            raise ValueError("不支持的图片输入类型")

        if image.mode != 'RGB':
            image = image.convert('RGB')

        # 长边限制 2000 像素，与项目其它环节保持一致
        max_size = 2000
        if max(image.size) > max_size:
            ratio = max_size / max(image.size)
            new_size = (int(image.size[0] * ratio), int(image.size[1] * ratio))
            image = image.resize(new_size, Image.Resampling.LANCZOS)

        img_io = io.BytesIO()
        image.save(img_io, format='PNG')
        data_url = 'data:image/png;base64,' + base64.b64encode(img_io.getvalue()).decode('utf-8')

        client = OpenAI(base_url=minimax_base_url, api_key=minimax_api_key, timeout=120.0)
        response = _vision_chat(
            client,
            model=minimax_model,
            messages=[
                {"role": "system", "content": "你是OCR文字识别引擎。只输出图片中识别到的文字，逐字还原，保留原有段落与换行，不翻译、不点评、不添加任何解释性内容。若图中没有文字，输出空。"},
                {"role": "user", "content": [
                    {"type": "image_url", "image_url": {"url": data_url}},
                    {"type": "text", "text": "请识别图片中的全部文字。"}
                ]}
            ],
            temperature=0.1,
            max_tokens=4000
        )
        text = (response.choices[0].message.content or '').strip()
        if logger and text:
            logger.info(f"大模型识图 识别成功，文本长度: {len(text)}")
        return text
    except Exception as e:
        if logger:
            logger.warning(f"大模型识图 识别失败: {e}")
        return ''

def xunfei_ocr_recognize(image_path_or_bytes, logger=None):
    """
    使用讯飞OCR进行文字识别
    
    Args:
        image_path_or_bytes: 图片路径（str）或图片字节数据（bytes）
        logger: 日志记录器（可选）
    
    Returns:
        str: 识别出的文本内容，如果失败返回空字符串
    """
    try:
        # 读取图片数据
        if isinstance(image_path_or_bytes, str):
            # 文件路径
            with open(image_path_or_bytes, "rb") as f:
                image_bytes = f.read()
        elif isinstance(image_path_or_bytes, bytes):
            # 字节数据
            image_bytes = image_path_or_bytes
        else:
            # PIL Image对象，转换为字节
            if hasattr(image_path_or_bytes, 'save'):
                img_io = io.BytesIO()
                image_path_or_bytes.save(img_io, format='PNG')
                image_bytes = img_io.getvalue()
            else:
                if logger:
                    logger.warning("讯飞OCR: 不支持的图片格式")
                return ""
        
        # 检查图片大小（base64编码后不超过10M）
        if len(image_bytes) > 7 * 1024 * 1024:  # 约7MB，留一些余量
            if logger:
                logger.warning(f"讯飞OCR: 图片太大 ({len(image_bytes)} bytes)，跳过")
            return ""
        
        # 构建请求体
        body = {
            "header": {
                "app_id": XF_APP_ID,
                "status": 3
            },
            "parameter": {
                "sf8e6aca1": {
                    "category": "ch_en_public_cloud",
                    "result": {
                        "encoding": "utf8",
                        "compress": "raw",
                        "format": "json"
                    }
                }
            },
            "payload": {
                "sf8e6aca1_data_1": {
                    "encoding": "jpg",
                    "image": str(base64.b64encode(image_bytes), 'UTF-8'),
                    "status": 3
                }
            }
        }
        
        # 构建认证URL
        request_url = assemble_ws_auth_url(XF_OCR_URL, "POST", XF_API_KEY, XF_API_SECRET)
        headers = {'content-type': "application/json", 'host': 'api.xf-yun.com', 'app_id': XF_APP_ID}
        
        # 发送请求
        response = requests.post(request_url, data=json.dumps(body), headers=headers, timeout=30)
        response.raise_for_status()
        
        # 解析响应
        temp_result = json.loads(response.content.decode())
        decoded_text = base64.b64decode(temp_result['payload']['result']['text']).decode()
        
        # 解析JSON结果，提取文本内容
        import re
        try:
            result_json = json.loads(decoded_text)
            text_lines = []
            
            # 处理不同的JSON结构
            if 'pages' in result_json:
                for page in result_json['pages']:
                    if 'lines' in page:
                        for line in page['lines']:
                            if 'content' in line:
                                text_lines.append(line['content'])
            elif 'lines' in result_json:
                for line in result_json['lines']:
                    if 'content' in line:
                        text_lines.append(line['content'])
            elif isinstance(result_json, list):
                for item in result_json:
                    if isinstance(item, dict) and 'content' in item:
                        text_lines.append(item['content'])
            else:
                if 'content' in result_json:
                    text_lines.append(result_json['content'])
                elif 'text' in result_json:
                    text_lines.append(result_json['text'])
            
            # 智能合并文本行
            if text_lines:
                merged_lines = []
                current_line = ""
                
                for line in text_lines:
                    line = line.strip()
                    if not line:
                        if current_line:
                            merged_lines.append(current_line.strip())
                            current_line = ""
                        continue
                    
                    if not current_line:
                        current_line = line
                        continue
                    
                    # 判断是否应该合并
                    should_merge = False
                    if current_line.rstrip().endswith(('.', '!', '?', '。', '！', '？')):
                        should_merge = False
                    elif len(line) <= 5 and not line[0].isupper():
                        should_merge = True
                    elif not current_line.rstrip().endswith(('.', '!', '?', '。', '！', '？', ':', ';')) and not line[0].isupper():
                        should_merge = True
                    elif current_line.rstrip().endswith((',', ';', '，', '；')) and not line[0].isupper():
                        should_merge = True
                    elif line[0].isupper():
                        should_merge = False
                    else:
                        should_merge = True
                    
                    if should_merge:
                        current_line += " " + line
                    else:
                        merged_lines.append(current_line.strip())
                        current_line = line
                
                if current_line:
                    merged_lines.append(current_line.strip())
                
                final_result = '\n'.join(merged_lines).strip()
                # 清理多余的空白字符
                final_result = re.sub(r' +', ' ', final_result)
                final_result = re.sub(r'\n{3,}', '\n\n', final_result)
                final_result = '\n'.join([line.strip() for line in final_result.split('\n')])
            else:
                final_result = decoded_text.strip()
                final_result = re.sub(r' +', ' ', final_result)
                final_result = re.sub(r'\n{3,}', '\n\n', final_result)
                final_result = '\n'.join([line.strip() for line in final_result.split('\n')])
        except (json.JSONDecodeError, KeyError, TypeError):
            # 如果不是JSON格式，直接使用原始文本
            import re
            final_result = decoded_text.strip()
            final_result = re.sub(r' +', ' ', final_result)
            final_result = re.sub(r'\n{3,}', '\n\n', final_result)
            final_result = '\n'.join([line.strip() for line in final_result.split('\n')])
        
        if logger:
            logger.info(f"讯飞OCR 识别成功，文本长度: {len(final_result)}")
        return final_result
        
    except Exception as e:
        if logger:
            logger.warning(f"讯飞OCR 识别失败: {e}")
        return ""

def get_png_quality():
    """获取PNG压缩质量配置"""
    try:
        if os.path.exists(PNG_QUALITY_CONFIG_FILE):
            with open(PNG_QUALITY_CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                quality = config.get('quality', DEFAULT_PNG_QUALITY)
                # 确保质量值在有效范围内
                return max(0, min(100, int(quality)))
    except Exception:
        pass
    return DEFAULT_PNG_QUALITY

def get_naps2_path():
    """获取NAPS2控制台程序路径配置"""
    try:
        if os.path.exists(NAPS2_CONFIG_FILE):
            with open(NAPS2_CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                naps2_path = config.get('naps2_path', DEFAULT_NAPS2_PATH)
                return naps2_path.strip()
    except Exception:
        pass
    return DEFAULT_NAPS2_PATH

def get_scan_output_dir():
    """获取扫描输出目录（带回退，避免因目录不存在或无权限导致扫描直接失败）。

    依次尝试：配置文件指定目录 -> 默认目录 -> 项目内 data/scan_output，
    每个候选目录都做“创建 + 实际写入探测”，第一个可用的即返回。
    """
    candidates = []
    try:
        if os.path.exists(SCAN_OUTPUT_CONFIG_FILE):
            with open(SCAN_OUTPUT_CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
            cfg_dir = (config.get('scan_output_dir') or '').strip()
            if cfg_dir:
                candidates.append(cfg_dir)
    except Exception:
        pass
    candidates.append(DEFAULT_SCAN_OUTPUT_DIR)
    candidates.append(os.path.join(DATA_DIR, 'scan_output'))

    for d in candidates:
        try:
            os.makedirs(d, exist_ok=True)
            probe = os.path.join(d, '.writable_probe')
            with open(probe, 'w', encoding='utf-8') as fp:
                fp.write('ok')
            os.remove(probe)
            return d
        except Exception:
            continue

    fallback = os.path.join(DATA_DIR, 'scan_output')
    os.makedirs(fallback, exist_ok=True)
    return fallback


def initialize_naps2_powershell_function():
    """初始化NAPS2 PowerShell函数"""
    try:
        naps2_path = get_naps2_path()
        
        # 构建PowerShell函数定义脚本
        ps_function_script = f'''
function naps2.console {{
    # 定义NAPS2.Console.exe的完整路径（适配含空格的路径）
    $naps2ConsolePath = "{naps2_path}"
    
    # 检查文件是否存在，避免调用失败
    if (-not (Test-Path -Path $naps2ConsolePath -PathType Leaf)) {{
        Write-Error "错误：未找到NAPS2控制台程序，请检查路径是否正确！路径：$naps2ConsolePath"
        return
    }}
    # 调用程序并传递参数
    & $naps2ConsolePath @args
}}
'''
        
        # 将函数定义写入临时PowerShell脚本文件
        ps_init_file = os.path.join(CONFIG_DIR, 'naps2_init.ps1')
        with open(ps_init_file, 'w', encoding='utf-8') as f:
            f.write(ps_function_script)
        
        # 执行PowerShell脚本以定义函数
        import subprocess
        result = subprocess.run(
            ['powershell', '-ExecutionPolicy', 'Bypass', '-File', ps_init_file],
            capture_output=True,
            text=True,
            encoding='gbk',
            errors='replace',
            timeout=10
        )
        
        if result.returncode != 0:
            print(f"警告：初始化NAPS2 PowerShell函数失败: {result.stderr}")
            return False
        
        print("NAPS2 PowerShell函数初始化成功")
        return True
    except Exception as e:
        print(f"初始化NAPS2 PowerShell函数时出错: {e}")
        return False

# 缓存PNG质量配置，避免频繁读取文件
_cached_png_quality = None
_cached_quality_time = 0
_quality_cache_ttl = 60  # 缓存60秒

# TIF和PDF转换功能已移除 - 扫描仪现在直接生成PNG格式

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/classes')
def get_classes():
    """获取所有班级列表（含学生人数与学生名单，便于班级中心展示与成绩报表联动）"""
    classes = []
    try:
        if os.path.exists(CLASSES_DIR):
            for item in os.listdir(CLASSES_DIR):
                item_path = os.path.join(CLASSES_DIR, item)
                if os.path.isdir(item_path):
                    all_folders = [f for f in os.listdir(item_path)
                                  if os.path.isdir(os.path.join(item_path, f)) and
                                  not f.startswith('.')]
                    student_ids = []
                    students_file = os.path.join(STUDENTS_DIR, f"{item}.json")
                    if os.path.exists(students_file):
                        try:
                            with open(students_file, 'r', encoding='utf-8') as f:
                                raw = json.load(f)
                            student_ids = [str(s) for s in raw] if isinstance(raw, list) else list(raw.keys()) if isinstance(raw, dict) else []
                        except Exception:
                            student_ids = []
                    names_map = get_student_names_map()
                    student_names = {sid: names_map.get(sid, '') for sid in student_ids}
                    classes.append({
                        'name': item,
                        'path': item,
                        'folder_count': len(all_folders),
                        'student_count': len(student_ids),
                        'students': student_ids,
                        'student_names': student_names,
                    })
        classes.sort(key=lambda x: x['name'])
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    return jsonify({'classes': classes})

@app.route('/api/classes', methods=['POST'])
def create_class():
    """创建新班级"""
    data = request.get_json()
    class_name = data.get('name', '').strip()
    
    if not class_name:
        return jsonify({'error': '班级名称不能为空'}), 400
    
    # 清理班级名称，移除非法字符
    class_name = class_name.replace('/', '').replace('\\', '').replace(':', '').replace('*', '').replace('?', '').replace('"', '').replace('<', '').replace('>', '').replace('|', '')
    
    class_path = os.path.join(CLASSES_DIR, class_name)
    if os.path.exists(class_path):
        return jsonify({'error': '班级已存在'}), 400
    
    try:
        os.makedirs(class_path, exist_ok=True)
        return jsonify({'msg': '班级创建成功', 'class_name': class_name})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def student_id_to_class_code(student_id):
    """
    根据**完整**学号推导班级编码。两种学号格式并存：

    · 6 位（当前主用）：前2位年级 + 第3-4位班号 + 第5-6位座号（260101 -> 26级1班01号 -> 班级 2601）；
    · 4 位（2026 级起的老格式，现已退化为「后四位」写法）：前2位班号 + 后2位座号（1426 -> 14班26号 -> 班级 14）。

    ⚠️ 高一学生卷面只写**后四位**（260101 写成 0101），那种输入必须先经
    `resolve_student_id()` 回查成完整学号再调本函数，否则会被当成「14 班」。

    非纯数字或长度不符时返回 None。
    """
    s = str(student_id or '').strip()
    if not s.isdigit():
        return None
    if len(s) == 6:
        return s[:4]
    if len(s) == 4:
        return s[:2]
    return None


_STUDENT_INDEX_CACHE = {'key': None, 'ids': set(), 'names': {}, 'by_suffix': {}}


def _student_index():
    """学生学号索引（按账号表 mtime+size 缓存）：完整学号集合、学号→姓名、后四位→[完整学号]。"""
    try:
        st = os.stat(USERS_FILE)
        key = (st.st_mtime, st.st_size)
    except OSError:
        key = None
    cache = _STUDENT_INDEX_CACHE
    if key is not None and cache['key'] == key:
        return cache
    ids, names, by_suffix = set(), {}, {}
    try:
        for u in load_users():
            if u.get('role') != 'student':
                continue
            sid = str(u.get('id') or '').strip()
            if not sid.isdigit():
                continue
            ids.add(sid)
            names[sid] = (u.get('name') or '').strip()
            if len(sid) >= 4:
                by_suffix.setdefault(sid[-4:], []).append(sid)
    except Exception as e:
        print('构建学号索引失败: %s' % e)
    cache.update({'key': key, 'ids': ids, 'names': names, 'by_suffix': by_suffix})
    return cache


def _four_digit_grade_prefix():
    """哪个年级的卷面只写学号后四位（返回学号前两位的年级码，如 '26'）。

    现实依据：只有当前高一这一届按「后四位」写学号，高二高三写完整六位。
    六位学号 = 年级2位 + 班号2位 + 座号2位，而三个年级的班号、座号完全同构，
    所以**班号无法跨年级消歧**，只能靠这个年级前缀补全。
    学年初换届时改这里（或用环境变量 FOUR_DIGIT_GRADE 覆盖）。
    """
    v = (os.environ.get('FOUR_DIGIT_GRADE') or '').strip()
    return v if len(v) == 2 and v.isdigit() else '26'


def resolve_student_id(observed, class_no=''):
    """把卷面上观测到的学号解析成名单里的完整学号。

    高一学生卷面只写**后四位**（260101 写成 0101）。四位号在三个年级间完全撞号
    （0101 同时对应 240101 / 250101 / 260101），班号也相同，所以不能用班号消歧，
    只能按 `_four_digit_grade_prefix()` 指定的年级补全：26 + 0101 -> 260101。

    返回 (学号, 说明)：
      · 说明为空  → 已在名单中唯一命中，可放心入库建号；
      · 说明非空 → 解析不确定，学号原样返回，由上层标记待确认，**不要**据此建账号。
    """
    raw = re.sub(r'[^\d]', '', str(observed or ''))
    if not raw:
        return '', '未识别到学号'
    idx = _student_index()

    if len(raw) == 6:
        if raw in idx['ids']:
            return raw, ''
        return raw, '学号不在名单中'

    if len(raw) == 4:
        cn = re.sub(r'[^\d]', '', str(class_no or ''))
        cn = cn[-2:].zfill(2) if cn else ''
        # 主路径：后四位前两位本身就是班号，直接补年级前缀
        guess = _four_digit_grade_prefix() + raw
        if guess in idx['ids']:
            if cn and cn != raw[:2]:
                return guess, '班号「%s」与后四位「%s」不一致，已按学号归档' % (cn, raw[:2])
            return guess, ''
        # 兜底：前缀补全不中，再看全表是否只有一个匹配
        cands = sorted(idx['by_suffix'].get(raw, []))
        if len(cands) == 1:
            return cands[0], ''
        if not cands:
            return raw, '后四位在名单中无匹配'
        return raw, '后四位对应多人（%s），需人工确认' % ' / '.join(cands[:4])

    return raw, '学号位数异常（%d 位）' % len(raw)


def ensure_class_for_student(student_id, student_name=None):
    """
    根据学号自动确保班级存在，并把该学生加入班级名单。
    班级编码取学号前4位（如 230205 -> 2302）。若班级不存在则创建目录与学生列表；
    老师也可在界面手动创建班级，本函数只做“不存在则创建”的补充。
    """
    class_code = student_id_to_class_code(student_id)
    if not class_code:
        return
    class_path = os.path.join(CLASSES_DIR, class_code)
    students_file = os.path.join(STUDENTS_DIR, f"{class_code}.json")
    try:
        os.makedirs(class_path, exist_ok=True)
    except Exception:
        pass
    students = []
    if os.path.exists(students_file):
        try:
            with open(students_file, 'r', encoding='utf-8') as f:
                raw = json.load(f)
            if isinstance(raw, list):
                students = [str(x) for x in raw]
            elif isinstance(raw, dict):
                students = list(raw.keys())
            else:
                students = []
        except Exception:
            students = []
    sid = str(student_id)
    if sid not in students:
        students.append(sid)
        students.sort()
        with open(students_file, 'w', encoding='utf-8') as f:
            json.dump(students, f, ensure_ascii=False, indent=2)


def _clean_person_name(name):
    """校验收 OCR 出来的人名：只保留 2-4 个汉字，其余一律视为不可用（返回空串）。

    手写姓名识别噪声很大，宁可留空让老师补，也不要把乱码写进名单。
    """
    s = re.sub(r'[\s.。·,，、]+', '', str(name or ''))
    s = re.sub(r'[^\u4e00-\u9fa5]', '', s)
    return s[:4] if 2 <= len(s) <= 4 else ''


def _ensure_student_registry(student_id, name='', class_no=''):
    """认出学号后自动完成「建班 + 进名单 + 建学生账号」，老师不需要再做任何归档动作。

    两条硬规矩：
      · **学号是主键，姓名以名单为准**。中文手写姓名识别不稳（同一张图会给出多个
        答案），所以 OCR 姓名只在名单里查不到该学号时才作为兜底；
      · **解析不出的学号不建账号**。宁可让这条记录带「待确认」标记，也不要把一个
        猜错的学号写进名单，否则以后每次批阅都会沿用错的。

    返回 (class_code, display_name, resolved_id, note)。note 非空表示解析不确定。
    """
    raw = str(student_id or '').strip()
    sid, note = resolve_student_id(raw, class_no)
    if not sid:
        return '', '', raw, note or '学号为空'
    class_code = student_id_to_class_code(sid)
    if not class_code:
        return '', '', sid, note or '学号位数异常'

    idx = _student_index()
    roster_name = idx['names'].get(sid) or ''
    if sid not in idx['ids']:
        # 不在名单：不建班、不建号，也不拿后四位去推班级（会推出一个不存在的班），
        # 由上层把它归到「待确认」，老师一眼能看到哪些卷子没认准。
        return '', '', sid, note or '学号不在名单中'

    name = _clean_person_name(name)

    # 1) 班级目录 + 名单
    try:
        ensure_class_for_student(sid, roster_name or None)
    except Exception as e:
        print('自动建班失败: %s' % e)

    # 2) 学生账号（密码=学号），让学生端能直接登录看自己的卷子
    try:
        users = load_users()
        hit = next((u for u in users if str(u.get('id', '')) == sid), None)
        if hit is None:
            users.append({'id': sid, 'name': roster_name or name or sid,
                          'role': 'student', 'password': sid})
            save_users(users)
        elif roster_name and str(hit.get('name') or '').strip() != roster_name:
            # 名单是权威来源：账号上若留着早期 OCR 猜的姓名，这里纠正过来
            hit['name'] = roster_name
            save_users(users)
    except Exception as e:
        print('自动建学生账号失败: %s' % e)

    return class_code, roster_name or name, sid, ''


def _task_papers_count(task_id):
    """返回某任务下已收试卷数量（用于阅卷进度展示）"""
    task_dir = os.path.join(TASK_PAPERS_DIR, task_id)
    if not os.path.isdir(task_dir):
        return 0
    return sum(1 for f in os.listdir(task_dir) if os.path.isfile(os.path.join(task_dir, f)) and f[0] != '.')


def _task_graded_count(task_id):
    """返回某任务下已批阅数量（从 task_results 统计）"""
    result_file = os.path.join(TASK_RESULTS_DIR, f'{task_id}.json')
    if not os.path.isfile(result_file):
        return 0
    try:
        with open(result_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return len(data.get('results') or {})
    except Exception:
        return 0


@app.route('/api/tasks', methods=['GET'])
def get_tasks():
    """获取任务列表（组卷发布产生的任务，用于阅卷管理、导出答题卡等）。每项含 papers_count（已收卷数）。"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    tasks = load_tasks()
    # 为每项附加已收卷数，供阅卷进度使用
    for i, t in enumerate(tasks):
        tid = t.get('id')
        if tid:
            tasks[i] = dict(t)
            tasks[i]['papers_count'] = _task_papers_count(tid)
            tasks[i]['graded_count'] = _task_graded_count(tid)
    # 按创建时间倒序
    tasks = sorted(tasks, key=lambda t: t.get('created_at', ''), reverse=True)
    return jsonify({'tasks': tasks})


@app.route('/api/tasks', methods=['POST'])
def create_task():
    """发布=创建任务。body: title, class_names[], deadline?, student_ids?, items[]（试卷篮）, answer_sheet_html?（可选，不传则用当前通用答题卡快照）"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    data = request.get_json() or {}
    title = (data.get('title') or '').strip() or '作文练习'
    class_names = data.get('class_names') or []
    if not isinstance(class_names, list):
        class_names = [class_names] if class_names else []
    class_names = [str(c).strip() for c in class_names if str(c).strip()]
    if not class_names:
        return jsonify({'error': '请至少选择一个班级'}), 400
    deadline = data.get('deadline') or ''
    student_ids = data.get('student_ids') or []
    if not isinstance(student_ids, list):
        student_ids = [student_ids] if student_ids else []
    items = data.get('items') or []
    if not isinstance(items, list):
        items = []
    answer_sheet_html = data.get('answer_sheet_html')
    if answer_sheet_html is None:
        answer_sheet_html = _read_answer_sheet_html()
    subject_cfg = _get_subject_config()
    subject = (data.get('subject') or subject_cfg.get('current') or 'english').strip() or 'english'
    task_id = str(uuid.uuid4())
    created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    task = {
        'id': task_id,
        'title': title,
        'class_names': class_names,
        'deadline': deadline,
        'student_ids': student_ids,
        'items': items,
        'answer_sheet_html': answer_sheet_html,
        'created_at': created_at,
        'status': 'published',
        'subject': subject,
    }
    tasks = load_tasks()
    tasks.append(task)
    save_tasks(tasks)
    return jsonify({'ok': True, 'task': task})


@app.route('/api/tasks/<task_id>', methods=['GET'])
def get_task(task_id):
    """获取单个任务详情（用于导出答题卡、阅卷入口）"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    tasks = load_tasks()
    for t in tasks:
        if t.get('id') == task_id:
            return jsonify(t)
    return jsonify({'error': '任务不存在'}), 404


@app.route('/api/tasks/<task_id>/grading_config', methods=['PUT', 'POST'])
def save_task_grading_config(task_id):
    """保存某任务的阅卷设置（单评/双评等），供「创建阅卷任务」保存时调用"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    data = request.get_json() or {}
    tasks = load_tasks()
    for i, t in enumerate(tasks):
        if t.get('id') == task_id:
            tasks[i] = dict(t)
            tasks[i]['grading_config'] = {
                **(t.get('grading_config') or {}),
                **data,
                'saved_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            }
            save_tasks(tasks)
            return jsonify({'ok': True, 'task_id': task_id})
    return jsonify({'error': '任务不存在'}), 404


def _inject_task_qr_into_html(html, task_id):
    """在答题卡 HTML 中注入任务二维码（编码 task_id），便于打印后扫码自动归类。无 qrcode 库时返回原 HTML。"""
    try:
        import qrcode
        buf = io.BytesIO()
        qrcode.make(task_id, box_size=4, border=1).save(buf, format='PNG')
        buf.seek(0)
        b64 = base64.b64encode(buf.read()).decode('ascii')
        qr_block = (
            '<div class="task-qr" style="margin-top:16px;text-align:center;">'
            '<span style="font-size:12px;color:#666;">扫此码归入本任务</span><br>'
            '<img src="data:image/png;base64,' + b64 + '" alt="task:' + task_id + '" style="width:80px;height:80px;"/>'
            '</div>'
        )
        if '</body>' in html:
            html = html.replace('</body>', qr_block + '\n</body>')
        else:
            html = html + qr_block
    except Exception:
        pass
    return html


@app.route('/api/tasks/<task_id>/answer_sheet')
def get_task_answer_sheet(task_id):
    """获取任务对应的答题卡 HTML（预览或下载，便于打印；含任务二维码便于扫码归类）"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    tasks = load_tasks()
    for t in tasks:
        if t.get('id') == task_id:
            html = t.get('answer_sheet_html') or _read_answer_sheet_html()
            html = _inject_task_qr_into_html(html, task_id)
            if request.args.get('download'):
                return Response(html, mimetype='text/html; charset=utf-8',
                                headers={'Content-Disposition': 'attachment; filename=答题卡_' + task_id[:8] + '.html'})
            return Response(html, mimetype='text/html; charset=utf-8')
    return jsonify({'error': '任务不存在'}), 404


@app.route('/api/papers/upload_auto', methods=['POST'])
def upload_paper_auto():
    """上传一张试卷图片，自动解析图中二维码得到 task_id 并归入该任务；未识别到有效任务二维码时返回错误。支持 multipart 或 JSON base64。"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    image_bytes = None
    ext = '.png'
    if request.files:
        f = request.files.get('file') or request.files.get('image')
        if f and f.filename:
            image_bytes = f.read()
            ext = os.path.splitext(f.filename)[1].lower() or '.png'
            if ext not in ('.png', '.jpg', '.jpeg', '.webp'):
                ext = '.png'
    if not image_bytes and request.get_json(silent=True):
        data = request.get_json()
        b64 = data.get('image') or data.get('base64') or data.get('content')
        if b64:
            if isinstance(b64, str) and ',' in b64 and b64.startswith('data:'):
                b64 = b64.split(',', 1)[-1]
            image_bytes = base64.b64decode(b64)
    if not image_bytes:
        return jsonify({'ok': False, 'error': '请上传图片或提供 base64'}), 400
    decoded = decode_qr_from_image(image_bytes)
    task_ids = {t.get('id') for t in load_tasks() if t.get('id')}
    if not decoded or decoded not in task_ids:
        return jsonify({
            'ok': False,
            'error': '未识别到任务二维码或任务不存在，请使用带任务二维码的答题卡或手动指定任务后上传'
        }), 200
    try:
        filename = _save_paper_to_task(decoded, image_bytes, ext)
        return jsonify({'ok': True, 'task_id': decoded, 'filename': filename})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@app.route('/api/tasks/<task_id>/papers', methods=['POST'])
def upload_task_paper(task_id):
    """扫码或上传时，将一张试卷图片归入指定任务（便于后续按任务批改、学号对应）。支持 multipart 文件或 JSON body 的 base64 图片。"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    tasks = load_tasks()
    if not any(t.get('id') == task_id for t in tasks):
        return jsonify({'error': '任务不存在'}), 404
    task_dir = os.path.join(TASK_PAPERS_DIR, task_id)
    os.makedirs(task_dir, exist_ok=True)
    try:
        if request.files:
            f = request.files.get('file') or request.files.get('image')
            if not f or not f.filename:
                return jsonify({'error': '请上传文件'}), 400
            ext = os.path.splitext(f.filename)[1].lower() or '.png'
            if ext not in ('.png', '.jpg', '.jpeg', '.webp'):
                ext = '.png'
            name = str(uuid.uuid4()) + ext
            path = os.path.join(task_dir, name)
            f.save(path)
        else:
            data = request.get_json() or {}
            b64 = data.get('image') or data.get('base64') or data.get('content')
            if not b64:
                return jsonify({'error': '请提供 file 或 JSON 中的 image/base64'}), 400
            if isinstance(b64, str) and b64.startswith('data:'):
                b64 = b64.split(',', 1)[-1] if ',' in b64 else b64
            raw = base64.b64decode(b64)
            name = str(uuid.uuid4()) + '.png'
            path = os.path.join(task_dir, name)
            with open(path, 'wb') as out:
                out.write(raw)
        return jsonify({'ok': True, 'filename': name, 'task_id': task_id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/tasks/<task_id>/papers', methods=['GET'])
def list_task_papers(task_id):
    """列出某任务下已归类的试卷图片（扫码/上传后的列表，便于批改与结果分析）"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    tasks = load_tasks()
    if not any(t.get('id') == task_id for t in tasks):
        return jsonify({'error': '任务不存在'}), 404
    task_dir = os.path.join(TASK_PAPERS_DIR, task_id)
    if not os.path.isdir(task_dir):
        return jsonify({'task_id': task_id, 'papers': []})
    papers = [f for f in os.listdir(task_dir) if os.path.isfile(os.path.join(task_dir, f)) and f[0] != '.']
    papers.sort()
    return jsonify({'task_id': task_id, 'papers': papers})


@app.route('/api/tasks/<task_id>/papers/<path:filename>')
def get_task_paper_file(task_id, filename):
    """获取任务下某张试卷图片（供批改界面展示）"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    task_dir = os.path.join(TASK_PAPERS_DIR, task_id)
    path = os.path.join(task_dir, filename)
    if not os.path.isfile(path) or not os.path.realpath(path).startswith(os.path.realpath(task_dir)):
        return jsonify({'error': '文件不存在'}), 404
    return send_from_directory(task_dir, filename)


@app.route('/api/tasks/<task_id>/answer_situation')
def get_task_answer_situation(task_id):
    """获取某任务的答题/批阅情况（用于展示或导出）"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    tasks = load_tasks()
    if not any(t.get('id') == task_id for t in tasks):
        return jsonify({'error': '任务不存在'}), 404
    result_file = os.path.join(TASK_RESULTS_DIR, f'{task_id}.json')
    if not os.path.isfile(result_file):
        return jsonify({'task_id': task_id, 'results': [], 'updated_at': None})
    try:
        with open(result_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception:
        return jsonify({'task_id': task_id, 'results': [], 'updated_at': None})
    results = []
    for filename, rec in (data.get('results') or {}).items():
        results.append({
            'filename': filename,
            'student_id': rec.get('student_id', ''),
            'report': rec.get('report', ''),
            'status': rec.get('status', ''),
        })
    results.sort(key=lambda x: (x.get('student_id') or '', x.get('filename', '')))
    return jsonify({'task_id': task_id, 'results': results, 'updated_at': data.get('updated_at')})


@app.route('/api/tasks/<task_id>/export_answer_situation')
def export_task_answer_situation(task_id):
    """导出某任务的答题情况为 CSV 下载"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    tasks = load_tasks()
    if not any(t.get('id') == task_id for t in tasks):
        return jsonify({'error': '任务不存在'}), 404
    result_file = os.path.join(TASK_RESULTS_DIR, f'{task_id}.json')
    data = {}
    if os.path.isfile(result_file):
        try:
            with open(result_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception:
            pass
    results = list((data.get('results') or {}).items())
    results.sort(key=lambda x: (x[1].get('student_id') or '', x[0]))
    import csv
    import io
    out = io.StringIO()
    writer = csv.writer(out)
    writer.writerow(['学号', '文件名', '状态', '批阅报告'])
    if not results:
        writer.writerow(['（暂无批阅数据）', '', '', ''])
    for filename, rec in results:
        report = (rec.get('report') or '').replace('\r\n', '\n').replace('\n', ' ')
        writer.writerow([rec.get('student_id', ''), filename, rec.get('status', ''), report])
    body = out.getvalue()
    try:
        body = body.encode('utf-8-sig')
    except Exception:
        body = body.encode('utf-8')
    from flask import Response
    res = Response(body, mimetype='text/csv; charset=utf-8')
    res.headers['Content-Disposition'] = f'attachment; filename="task_{task_id[:8]}_answer_situation.csv"'
    return res


# 学校/报告配置（用于导出全班报告单表头）
SCHOOL_CONFIG_FILE = os.path.join(CONFIG_DIR, 'school_config.json')


def _get_school_name():
    """读取学校名称，用于学生个人报告表头。"""
    if not os.path.isfile(SCHOOL_CONFIG_FILE):
        return ''
    try:
        with open(SCHOOL_CONFIG_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return (data.get('school_name') or '').strip()
    except Exception:
        return ''


@app.route('/api/tasks/<task_id>/export_class_report')
def export_class_report(task_id):
    """导出某任务下全班学生个人报告（合并为单 HTML 或 PDF），便于一次性打印下发。"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    tasks = load_tasks()
    task = next((t for t in tasks if t.get('id') == task_id), None)
    if not task:
        return jsonify({'error': '任务不存在'}), 404
    result_file = os.path.join(TASK_RESULTS_DIR, f'{task_id}.json')
    data = {}
    if os.path.isfile(result_file):
        try:
            with open(result_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception:
            pass
    results_raw = data.get('results') or {}
    if not results_raw:
        return jsonify({'error': '该任务暂无批阅结果，请先完成智能阅卷后再导出'}), 400
    results = []
    for filename, rec in results_raw.items():
        results.append({
            'filename': filename,
            'student_id': rec.get('student_id', ''),
            'report': rec.get('report', ''),
            'status': rec.get('status', ''),
            'score_text': rec.get('score_text', ''),
            'answer_time': rec.get('answer_time', ''),
        })
    student_names_map = get_student_names_map()
    school_name = _get_school_name()
    class_name = (task.get('class_names') or [''])[0]
    try:
        from utils.student_report_generator import generate_class_report_html
    except ImportError:
        return jsonify({'error': '报告生成模块不可用'}), 500
    html = generate_class_report_html(
        results=results,
        task=task,
        student_names_map=student_names_map,
        school_name=school_name,
        class_name=class_name,
    )
    output_format = (request.args.get('format') or 'html').lower()
    task_title = (task.get('title') or '练习').strip()
    safe_title = re.sub(r'[\\/:*?"<>|]', '-', task_title)[:30]
    date_str = datetime.now().strftime('%Y-%m-%d')
    if output_format == 'pdf':
        try:
            from utils.answer_sheet_generator import html_to_pdf
            pdf_name = f"{class_name}_学生个人报告_{safe_title}_{date_str}.pdf"
            pdf_path = os.path.join(EXPORTS_DIR, str(uuid.uuid4()) + '.pdf')
            if html_to_pdf(html, pdf_path):
                return send_file(
                    pdf_path, as_attachment=True, download_name=pdf_name,
                    mimetype='application/pdf'
                )
        except Exception:
            pass
        return jsonify({
            'error': 'PDF 生成失败。请改用 format=html 下载 HTML 后使用浏览器打印为 PDF，或安装 weasyprint: pip install weasyprint'
        }), 500
    buf = io.BytesIO(html.encode('utf-8'))
    buf.seek(0)
    html_name = f"{class_name}_学生个人报告_{safe_title}_{date_str}.html"
    return send_file(
        buf, as_attachment=True, download_name=html_name,
        mimetype='text/html; charset=utf-8'
    )


def is_file_locked(file_path, max_retries=3, retry_delay=0.5):
    """检查文件是否被锁定（正在使用中）"""
    import time
    for i in range(max_retries):
        try:
            # 尝试以写入模式打开文件，如果文件被锁定会抛出异常
            with open(file_path, 'r+b'):
                return False  # 文件未被锁定
        except (IOError, OSError, PermissionError):
            if i < max_retries - 1:
                time.sleep(retry_delay)
            else:
                return True  # 文件被锁定
    return False

@app.route('/api/organize', methods=['POST'])
def organize_files():
    """将日期文件夹或扫描输出文件夹中的文件按班级归档"""
    import time
    import shutil
    
    data = request.get_json()
    class_name = data.get('class_name', '').strip()
    date_folder = data.get('date_folder', '').strip()
    
    if not class_name or not date_folder:
        return jsonify({'error': '班级名称和文件夹不能为空'}), 400
    
    # 判断是否是扫描输出文件夹
    is_scan_output = date_folder == '__scan_output__'
    if is_scan_output:
        # 从扫描输出目录归档
        source_path = get_scan_output_dir()
        # 使用当前日期作为目标文件夹名称
        target_date_folder = datetime.now().strftime('%Y-%m-%d')
    else:
        # 从日期文件夹归档
        source_path = os.path.join(SCAN_DIR, date_folder)
        target_date_folder = date_folder
    
    if not os.path.exists(source_path):
        return jsonify({'error': '源文件夹不存在'}), 404
    
    class_path = os.path.join(CLASSES_DIR, class_name)
    if not os.path.exists(class_path):
        return jsonify({'error': '班级不存在'}), 404
    
    target_path = os.path.join(class_path, target_date_folder)
    os.makedirs(target_path, exist_ok=True)
    
    try:
        # 对于扫描输出文件夹，总是合并文件（不移动整个文件夹）
        # 如果目标文件夹已存在，合并文件
        if os.path.exists(target_path) or is_scan_output:
            # 移动新文件到目标文件夹
            moved_count = 0
            locked_files = []
            task_ids = {t.get('id') for t in load_tasks() if t.get('id')}

            for f in os.listdir(source_path):
                source_file = os.path.join(source_path, f)
                target_file = os.path.join(target_path, f)
                if os.path.isfile(source_file) and f.lower().endswith(('.jpg', '.jpeg', '.png')):
                    # 若图中含任务二维码，同步归入该任务下（便于按任务批改）
                    try:
                        with open(source_file, 'rb') as rf:
                            raw = rf.read()
                        decoded = decode_qr_from_image(raw)
                        if decoded and decoded in task_ids:
                            ext = os.path.splitext(f)[1].lower() or '.png'
                            if ext not in ('.png', '.jpg', '.jpeg', '.webp'):
                                ext = '.png'
                            _save_paper_to_task(decoded, raw, ext)
                    except Exception:
                        pass
                    if not os.path.exists(target_file):
                        # 检查文件是否被锁定，如果是则重试
                        max_retries = 5
                        retry_delay = 1.0  # 1秒
                        moved = False
                        
                        for attempt in range(max_retries):
                            try:
                                # 检查文件是否被锁定
                                if is_file_locked(source_file, max_retries=1, retry_delay=0.1):
                                    if attempt < max_retries - 1:
                                        time.sleep(retry_delay)
                                        continue
                                    else:
                                        locked_files.append(f)
                                        break
                                
                                shutil.move(source_file, target_file)
                                moved_count += 1
                                moved = True
                                break
                            except PermissionError as pe:
                                error_msg = str(pe)
                                if '32' in error_msg or 'being used' in error_msg.lower() or '另一个程序正在使用此文件' in error_msg:
                                    if attempt < max_retries - 1:
                                        time.sleep(retry_delay)
                                        continue
                                    else:
                                        locked_files.append(f)
                                        break
                                else:
                                    raise
                        
                        if not moved and f not in locked_files:
                            # 最后一次尝试
                            try:
                                shutil.move(source_file, target_file)
                                moved_count += 1
                            except Exception:
                                locked_files.append(f)
                    else:
                        # 如果文件已存在，重命名源文件后移动（避免覆盖）
                        try:
                            # 生成新的文件名（添加数字后缀）
                            name, ext = os.path.splitext(f)
                            counter = 1
                            new_target_file = os.path.join(target_path, f"{name}_{counter}{ext}")
                            while os.path.exists(new_target_file):
                                counter += 1
                                new_target_file = os.path.join(target_path, f"{name}_{counter}{ext}")
                            
                            # 移动文件到新名称
                            max_retries = 5
                            retry_delay = 1.0
                            moved = False
                            
                            for attempt in range(max_retries):
                                try:
                                    # 检查文件是否被锁定
                                    if is_file_locked(source_file, max_retries=1, retry_delay=0.1):
                                        if attempt < max_retries - 1:
                                            time.sleep(retry_delay)
                                            continue
                                        else:
                                            locked_files.append(f)
                                            break
                                    
                                    shutil.move(source_file, new_target_file)
                                    moved_count += 1
                                    moved = True
                                    break
                                except PermissionError as pe:
                                    error_msg = str(pe)
                                    if '32' in error_msg or 'being used' in error_msg.lower() or '另一个程序正在使用此文件' in error_msg:
                                        if attempt < max_retries - 1:
                                            time.sleep(retry_delay)
                                            continue
                                        else:
                                            locked_files.append(f)
                                            break
                                    else:
                                        raise
                            
                            if not moved and f not in locked_files:
                                # 最后一次尝试
                                try:
                                    shutil.move(source_file, new_target_file)
                                    moved_count += 1
                                except Exception:
                                    locked_files.append(f)
                        except Exception as e:
                            # 如果重命名失败，记录到锁定文件列表
                            locked_files.append(f)
            
            # 对于扫描输出文件夹，不删除源文件夹（它是配置的目录，还要继续使用）
            # 对于日期文件夹，如果变空了，可以删除
            remaining_files_count = 0
            if not is_scan_output:
                # 检查源文件夹是否为空（只包含非图片文件或为空），如果为空则删除
                remaining_files = [f for f in os.listdir(source_path) 
                                 if os.path.isfile(os.path.join(source_path, f)) 
                                 and f.lower().endswith(('.jpg', '.jpeg', '.png'))]
                remaining_files_count = len(remaining_files)
                
                if not remaining_files:
                    # 源文件夹中没有图片文件了，尝试删除整个文件夹
                    try:
                        max_retries = 3
                        for attempt in range(max_retries):
                            try:
                                os.rmdir(source_path)  # 只删除空文件夹
                                break
                            except (OSError, PermissionError):
                                if attempt < max_retries - 1:
                                    time.sleep(0.5)
                                # 如果删除失败，不影响归档结果
                                pass
                    except Exception:
                        pass  # 删除失败不影响归档结果
            
            if locked_files:
                locked_count = len(locked_files)
                return jsonify({
                    'error': f'部分文件正在被使用，无法归档。\n\n已成功归档 {moved_count} 个文件。\n有 {locked_count} 个文件无法归档（可能正在转换格式，请稍候再试）：\n' + '\n'.join(locked_files[:5]) + ('\n...' if len(locked_files) > 5 else ''),
                    'partial_success': True,
                    'moved_count': moved_count,
                    'locked_count': locked_count,
                    'locked_files': locked_files,
                    'source_folder_removed': False if is_scan_output else (remaining_files_count == 0)
                }), 200
            else:
                return jsonify({
                    'msg': f'已合并 {moved_count} 个文件到班级文件夹',
                    'source_folder_removed': False if is_scan_output else (remaining_files_count == 0)
                })
        else:
            # 如果是扫描输出文件夹，不应该移动整个文件夹，应该进入合并文件的逻辑
            if is_scan_output:
                # 对于扫描输出文件夹，应该走合并文件的逻辑，不应该到这里
                # 但如果到这里了，说明有问题，返回错误
                return jsonify({'error': '扫描输出文件夹归档逻辑错误'}), 500
            
            # 移动整个文件夹前，先检查文件夹中的文件是否被锁定
            locked_files = []
            for f in os.listdir(source_path):
                file_path = os.path.join(source_path, f)
                if os.path.isfile(file_path) and f.lower().endswith(('.jpg', '.jpeg', '.png')):
                    if is_file_locked(file_path, max_retries=1, retry_delay=0.1):
                        locked_files.append(f)
            
            if locked_files:
                locked_count = len(locked_files)
                return jsonify({
                    'error': f'文件夹中有文件正在被使用，无法归档。\n\n有 {locked_count} 个文件无法归档（可能正在转换格式，请稍候再试）：\n' + '\n'.join(locked_files[:5]) + ('\n...' if len(locked_files) > 5 else ''),
                    'partial_success': False,
                    'locked_count': locked_count,
                    'locked_files': locked_files
                }), 200
            
            # 尝试移动整个文件夹，带重试机制
            max_retries = 5
            retry_delay = 1.0
            for attempt in range(max_retries):
                try:
                    shutil.move(source_path, target_path)
                    # 移动整个文件夹成功，源文件夹已被删除
                    return jsonify({
                        'msg': '文件已归档到班级文件夹',
                        'source_folder_removed': True
                    })
                except PermissionError as pe:
                    error_msg = str(pe)
                    if ('32' in error_msg or 'being used' in error_msg.lower() or 
                        '另一个程序正在使用此文件' in error_msg) and attempt < max_retries - 1:
                        time.sleep(retry_delay)
                        continue
                    else:
                        raise
    except PermissionError as pe:
        error_msg = str(pe)
        if '32' in error_msg or 'being used' in error_msg.lower() or '另一个程序正在使用此文件' in error_msg:
            return jsonify({
                'error': '文件正在被使用，无法归档。\n\n💡 提示：请等待文件处理完成后再尝试归档。\n\n如果问题持续存在，请刷新页面后重试。'
            }), 200
        else:
            return jsonify({'error': f'归档失败：{error_msg}'}), 500
    except Exception as e:
        error_msg = str(e)
        if '32' in error_msg or 'being used' in error_msg.lower() or '另一个程序正在使用此文件' in error_msg:
            return jsonify({
                'error': '文件正在被使用，无法归档。\n\n💡 提示：请等待文件处理完成后再尝试归档。\n\n如果问题持续存在，请刷新页面后重试。'
            }), 200
        return jsonify({'error': f'归档失败：{error_msg}'}), 500

@app.route('/api/folders')
def get_folders():
    """获取所有日期文件夹（扫描仪生成的原始文件夹）和扫描输出文件夹"""
    folders = []
    try:
        # 首先添加扫描输出文件夹（作为特殊文件夹）
        scan_output_dir = get_scan_output_dir()
        if os.path.exists(scan_output_dir) and os.path.isdir(scan_output_dir):
            # 统计扫描输出目录中的图片文件
            image_files = [f for f in os.listdir(scan_output_dir) 
                         if os.path.isfile(os.path.join(scan_output_dir, f)) 
                         and f.lower().endswith(('.jpg', '.jpeg', '.png'))]
            folders.append({
                'name': '扫描文件夹',
                'path': '__scan_output__',  # 使用特殊标识符
                'file_count': len(image_files),
                'is_scan_folder': True
            })
        
        # 然后添加日期格式的文件夹
        for item in os.listdir(SCAN_DIR):
            item_path = os.path.join(SCAN_DIR, item)
            if os.path.isdir(item_path) and not item.startswith('.'):
                # 排除系统文件夹和classes文件夹，以及扫描输出目录（如果它在SCAN_DIR下）
                if item in ['classes', 'results', 'static', 'templates', 'Camera Roll', 'Screenshots', 'QQplayerPic', 'Feedback']:
                    continue
                # 检查是否是日期格式的文件夹
                try:
                    datetime.strptime(item, '%Y-%m-%d')
                    
                    # 统计文件夹中的图片文件
                    image_files = [f for f in os.listdir(item_path) 
                                 if os.path.isfile(os.path.join(item_path, f))
                                 and f.lower().endswith(('.jpg', '.jpeg', '.png'))]
                    folders.append({
                        'name': item,
                        'path': item,
                        'file_count': len(image_files)
                    })
                except ValueError:
                    # 不是日期格式，跳过
                    continue
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
    # 将扫描文件夹放在最前面，其他按日期排序
    scan_folders = [f for f in folders if f.get('is_scan_folder', False)]
    date_folders = [f for f in folders if not f.get('is_scan_folder', False)]
    date_folders.sort(key=lambda x: x['name'], reverse=True)
    folders = scan_folders + date_folders
    
    return jsonify({'folders': folders})

@app.route('/api/class_folders/<class_name>')
def get_class_folders(class_name):
    """获取指定班级下的所有文件夹"""
    class_path = os.path.join(CLASSES_DIR, class_name)
    if not os.path.exists(class_path):
        return jsonify({'error': '班级不存在'}), 404
    
    folders = []
    try:
        for item in os.listdir(class_path):
            item_path = os.path.join(class_path, item)
            if os.path.isdir(item_path) and not item.startswith('.'):
                # 统计文件夹中的图片文件
                image_files = [f for f in os.listdir(item_path) 
                             if os.path.isfile(os.path.join(item_path, f))
                             and f.lower().endswith(('.jpg', '.jpeg', '.png'))]
                
                # 检查是否是日期格式的文件夹
                is_date_folder = False
                try:
                    datetime.strptime(item, '%Y-%m-%d')
                    is_date_folder = True
                except ValueError:
                    pass
                
                folders.append({
                    'name': item,
                    'path': f"{class_name}/{item}",
                    'file_count': len(image_files),
                    'is_date_folder': is_date_folder
                })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
    # 排序：日期文件夹按日期倒序，其他文件夹按名称排序，日期文件夹优先
    date_folders = [f for f in folders if f.get('is_date_folder', False)]
    other_folders = [f for f in folders if not f.get('is_date_folder', False)]
    date_folders.sort(key=lambda x: x['name'], reverse=True)  # 日期倒序
    other_folders.sort(key=lambda x: x['name'])  # 名称正序
    folders = date_folders + other_folders
    
    return jsonify({'folders': folders})


@app.route('/api/class_folder/rename', methods=['POST'])
def rename_class_folder():
    """重命名班级下的日期文件夹"""
    data = request.get_json()
    class_name = data.get('class_name', '').strip()
    old_name = data.get('old_name', '').strip()
    new_name = data.get('new_name', '').strip()

    if not class_name or not old_name or not new_name:
        return jsonify({'error': '班级名称、原文件夹名和新文件夹名不能为空'}), 400

    # 防止路径穿越
    if '/' in new_name or '\\' in new_name:
        return jsonify({'error': '文件夹名称不能包含路径分隔符'}), 400

    class_path = os.path.join(CLASSES_DIR, class_name)
    old_path = os.path.join(class_path, old_name)
    new_path = os.path.join(class_path, new_name)

    if not os.path.exists(old_path):
        return jsonify({'error': '原文件夹不存在'}), 404

    if os.path.exists(new_path):
        return jsonify({'error': '目标文件夹已存在，请更换名称'}), 400

    try:
        os.rename(old_path, new_path)
        return jsonify({
            'success': True,
            'class_name': class_name,
            'old_name': old_name,
            'new_name': new_name
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/files/<path:folder_path>')
def get_files(folder_path):
    """获取指定文件夹中的图片文件（支持班级文件夹路径和扫描输出文件夹）"""
    # 判断是否是扫描输出文件夹（使用特殊标识符）
    if folder_path == '__scan_output__':
        # 扫描输出文件夹
        folder_path_full = get_scan_output_dir()
    elif '/' in folder_path:
        # 班级文件夹路径：class_name/date_folder
        parts = folder_path.split('/', 1)
        if len(parts) == 2:
            class_name, date_folder = parts
            folder_path_full = os.path.join(CLASSES_DIR, class_name, date_folder)
        else:
            folder_path_full = os.path.join(SCAN_DIR, folder_path)
    else:
        # 原始扫描文件夹（日期格式）
        folder_path_full = os.path.join(SCAN_DIR, folder_path)
    
    if not os.path.exists(folder_path_full):
        return jsonify({'error': '文件夹不存在'}), 404
    
    files = []
    try:
        for f in os.listdir(folder_path_full):
            file_path = os.path.join(folder_path_full, f)
            if os.path.isfile(file_path) and f.lower().endswith(('.jpg', '.jpeg', '.png')):
                
                # 检查临时文件夹中是否有处理过的版本
                name, ext = os.path.splitext(f)
                temp_filename = f"{name}_processed{ext}"
                temp_file_path = os.path.join(TEMP_DIR, temp_filename)
                has_processed = os.path.exists(temp_file_path)
                
                file_size = os.path.getsize(file_path)
                file_info = {
                    'name': f,
                    'path': f"{folder_path}/{f}",
                    'size': file_size,
                    'has_processed': has_processed
                }
                
                # 如果有处理过的版本，也返回处理过的文件信息
                if has_processed:
                    temp_file_size = os.path.getsize(temp_file_path)
                    file_info['processed_size'] = temp_file_size
                
                files.append(file_info)
        files.sort(key=lambda x: x['name'])
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
    return jsonify({'files': files})

@app.route('/api/png_quality', methods=['GET'])
def get_png_quality_config():
    """获取PNG压缩质量配置"""
    quality = get_png_quality()
    return jsonify({'quality': quality})

@app.route('/api/png_quality', methods=['POST'])
def set_png_quality_config():
    """设置PNG压缩质量配置"""
    data = request.get_json()
    quality = data.get('quality', DEFAULT_PNG_QUALITY)
    
    # 确保质量值在有效范围内
    quality = max(0, min(100, int(quality)))
    
    try:
        with open(PNG_QUALITY_CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump({'quality': quality}, f)
        return jsonify({'success': True, 'quality': quality})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/naps2_config', methods=['GET'])
def get_naps2_config():
    """获取NAPS2路径配置"""
    naps2_path = get_naps2_path()
    return jsonify({'naps2_path': naps2_path})

@app.route('/api/naps2_config', methods=['POST'])
def set_naps2_config():
    """设置NAPS2路径配置"""
    data = request.get_json()
    naps2_path = data.get('naps2_path', DEFAULT_NAPS2_PATH).strip()
    
    if not naps2_path:
        return jsonify({'error': 'NAPS2路径不能为空'}), 400
    
    try:
        with open(NAPS2_CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump({'naps2_path': naps2_path}, f)
        
        # 重新初始化PowerShell函数
        initialize_naps2_powershell_function()
        
        return jsonify({'success': True, 'naps2_path': naps2_path})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/scan_output_config', methods=['GET'])
def get_scan_output_config():
    """获取扫描输出路径配置"""
    scan_output_dir = get_scan_output_dir()
    return jsonify({'scan_output_dir': scan_output_dir})

@app.route('/api/scan_output_config', methods=['POST'])
def set_scan_output_config():
    """设置扫描输出路径配置"""
    data = request.get_json()
    scan_output_dir = data.get('scan_output_dir', DEFAULT_SCAN_OUTPUT_DIR).strip()
    
    if not scan_output_dir:
        return jsonify({'error': '扫描输出路径不能为空'}), 400
    
    try:
        # 确保路径是有效的目录路径
        if not os.path.isabs(scan_output_dir):
            return jsonify({'error': '扫描输出路径必须是绝对路径'}), 400
        
        # 尝试创建目录以验证路径是否有效
        try:
            os.makedirs(scan_output_dir, exist_ok=True)
        except Exception as e:
            return jsonify({'error': f'无法创建扫描输出目录: {str(e)}'}), 400
        
        with open(SCAN_OUTPUT_CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump({'scan_output_dir': scan_output_dir}, f)
        
        return jsonify({'success': True, 'scan_output_dir': scan_output_dir})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def get_scanner_advanced_config():
    """获取扫描仪高级设置（图像模式、扫描类型）"""
    try:
        if os.path.exists(SCANNER_ADVANCED_CONFIG_FILE):
            with open(SCANNER_ADVANCED_CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return {
                    'image_mode': config.get('image_mode', DEFAULT_SCANNER_IMAGE_MODE),
                    'scan_type': config.get('scan_type', DEFAULT_SCANNER_SCAN_TYPE),
                }
    except Exception:
        pass
    return {'image_mode': DEFAULT_SCANNER_IMAGE_MODE, 'scan_type': DEFAULT_SCANNER_SCAN_TYPE}

@app.route('/api/scanner_advanced_config', methods=['GET'])
def get_scanner_advanced_config_api():
    """获取扫描仪高级设置"""
    config = get_scanner_advanced_config()
    return jsonify(config)

@app.route('/api/scanner_advanced_config', methods=['POST'])
def set_scanner_advanced_config_api():
    """保存扫描仪高级设置（根据实际情况选择单面或双面）"""
    data = request.get_json()
    image_mode = data.get('image_mode', DEFAULT_SCANNER_IMAGE_MODE)
    scan_type = data.get('scan_type', DEFAULT_SCANNER_SCAN_TYPE)
    if image_mode not in ('grayscale', 'color'):
        image_mode = DEFAULT_SCANNER_IMAGE_MODE
    if scan_type not in ('single', 'double'):
        scan_type = DEFAULT_SCANNER_SCAN_TYPE
    try:
        with open(SCANNER_ADVANCED_CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump({'image_mode': image_mode, 'scan_type': scan_type}, f)
        return jsonify({'success': True, 'image_mode': image_mode, 'scan_type': scan_type})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 重命名功能已移除 - 直接读取并显示NAPS2生成的文件（如scan1.1.png, scan1.2.png等）

def get_next_scan_filename(scan_output_dir):
    """获取下一个可用的扫描文件名（格式：scan1.png, scan2.png, ...）"""
    if not os.path.exists(scan_output_dir):
        return 'scan1.png'
    
    # 查找所有 scan*.png 文件
    existing_files = []
    for filename in os.listdir(scan_output_dir):
        if filename.lower().startswith('scan') and filename.lower().endswith('.png'):
            existing_files.append(filename.lower())
    
    # 提取所有数字编号（支持多种格式：scan1.png, scan1.1.png, scan1.2.png等）
    used_numbers = set()
    for filename in existing_files:
        # 匹配 scan1.png 格式
        match = re.match(r'scan(\d+)\.png$', filename)
        if match:
            used_numbers.add(int(match.group(1)))
        # 匹配 scan1.1.png, scan1.2.png 等格式（提取主编号）
        match = re.match(r'scan(\d+)\.\d+\.png$', filename)
        if match:
            used_numbers.add(int(match.group(1)))
    
    # 查找下一个可用的编号
    next_number = 1
    while next_number in used_numbers:
        next_number += 1
    
    return f'scan{next_number}.png'

# ==================== 扫描报错管理（面向教师使用） ====================

def _decode_scan_output(raw):
    """解码 NAPS2.Console 的输出。

    NAPS2 在中文 Windows 下按 GBK 输出中文，若按 UTF-8 解码会变成乱码
    （例如“选定的扫描仪处于脱机状态”会显示为乱码），因此这里按 GBK 优先解码。
    """
    if raw is None:
        return ''
    if isinstance(raw, str):
        raw = raw.encode('utf-8', errors='replace')
    for enc in ('gbk', 'cp936', 'utf-8'):
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode('gbk', errors='replace')


def _humanize_scan_error(text, returncode=None):
    """把 NAPS2/驱动的原始报错翻译成教师能看懂的中文提示。

    返回 {'title': str, 'hints': [str, ...]}
    """
    t = (text or '').lower()

    def has(*keys):
        return any(k in t for k in keys)

    if has('maxconnections', 'twain error', 'max connections'):
        return {
            'title': '扫描仪连接已被占满（被其他程序或未结束的扫描任务占用）',
            'hints': [
                '系统已尝试自动释放被占用的连接',
                '若仍失败：拔掉扫描仪 USB 线，等 5 秒后重新插上，再点击扫描',
                '关闭其他正在使用扫描仪的程序（其它扫描软件、传真和扫描窗口）',
                '等待 10 秒后重试（上一次扫描超时后，扫描仪可能需要一点时间复位）',
            ],
        }
    if '#scan_stall#' in t:
        return {
            'title': '扫描仪无响应（扫描开始后卡住不动）',
            'hints': [
                '按一下扫描仪机身上的电源键唤醒，确认指示灯回到就绪状态后重试',
                '也可运行项目 tools 目录下的 reset_scanner.bat 一键软复位（无需拔线）',
                '关闭其他可能占用扫描仪的程序（其它扫描软件、Windows传真和扫描等）',
                '仍不行：拔掉 USB 线，等 10 秒后插回（建议插主机后置 USB 接口）',
            ],
        }
    if has('脱机', '离线', 'offline', 'is not available'):
        return {
            'title': '扫描仪未响应（显示为脱机/离线）',
            'hints': [
                '扫描仪可能处于休眠：按一下扫描仪机身上的电源键唤醒，等指示灯回到就绪状态',
                '重新插拔 USB 线，优先插在主机后置接口（前置接口供电不足容易导致脱机）',
                '关闭其他可能占用扫描仪的程序（其它扫描软件、Windows传真和扫描等）',
                '若扫描仪机身有模式按键（如“计数”），请退出该模式回到就绪状态',
                '也可运行项目 tools 目录下的 reset_scanner.bat 一键软复位（无需拔线）',
                '以上都不行时，重启电脑后再试',
            ],
        }
    if has('工作进程崩溃', '工作进程', 'crash', '崩溃'):
        return {
            'title': '扫描仪驱动工作进程崩溃（本次未扫描到内容）',
            'hints': [
                '多数情况是玻璃板上没有原稿：请把试卷正面朝下平放、对齐左上角，并盖好盖板后再试',
                '若已放好原稿仍报此错：按一下扫描仪机身上的电源键唤醒设备后重试',
                '仍不行：运行项目 tools 目录下的 reset_scanner.bat 一键软复位（无需拔线）',
                '最后手段：拔掉 USB 线，等 10 秒后插回（建议插主机后置 USB 接口）',
            ],
        }
    if has('no pages', '0 page', 'no scanned pages', '没有扫描到', '未检测到'):
        return {
            'title': '没有扫描到内容（未检测到原稿）',
            'hints': [
                '请把试卷/答题卡正面朝下平放在玻璃板上，并对齐左上角',
                '盖好扫描仪盖板（盖板未压合也可能导致扫描失败）',
                '确认纸张没有超出玻璃板的扫描区域',
                '放好后重新点击扫描',
            ],
        }
    if has('no device', 'no scanners', 'not found', '找不到', '未找到', '无法找到', '未选择'):
        return {
            'title': '未找到扫描仪设备',
            'hints': [
                '检查扫描仪电源是否打开、USB 线是否插紧',
                '打开“设备管理器”，确认“图像设备”下能看到扫描仪',
                '重新插拔 USB 线后重试，必要时重启电脑',
            ],
        }
    if has('jam', '卡纸'):
        return {
            'title': '扫描仪卡纸',
            'hints': [
                '打开盖板或进纸口，轻轻取出卡住的纸张',
                '把纸张整理平整（避免褶皱、卷边）后重新扫描',
            ],
        }
    if has('busy', 'in use', 'locked', '占用'):
        return {
            'title': '扫描仪正被其他程序占用',
            'hints': [
                '关闭其他正在使用扫描仪的程序（包括后台打开过的扫描窗口）',
                '等待几秒后重新点击扫描',
            ],
        }
    if has('driver', '驱动'):
        return {
            'title': '扫描仪驱动异常',
            'hints': [
                '在「扫描仪设置」中切换驱动后重试（本机实测推荐 TWAIN 驱动）',
                '确认扫描仪驱动已正确安装',
            ],
        }
    if has('timeout', '超时'):
        return {
            'title': '扫描超时（长时间未完成）',
            'hints': [
                '检查扫描仪电源与 USB 连接是否稳定',
                '降低扫描分辨率（如 300dpi 改为 200dpi）后重试',
                '重启扫描仪（断电 10 秒后重新上电）再试',
            ],
        }
    return {
        'title': '扫描失败',
        'hints': [
            '确认扫描仪电源已打开、USB 线插紧',
            '按一下扫描仪电源键唤醒设备后重试',
            '把原稿平放在玻璃板上并盖好盖板',
        ],
    }


def _cleanup_stale_scanner_processes():
    """清理残留的 NAPS2 控制台进程。

    扫描如果被中断（超时、服务重启等），NAPS2.Console 可能变成孤儿进程并一直
    占用扫描仪连接，导致之后每次扫描都报 "TWAIN error: MaxConnections"。
    这里在每次扫描前做一次清理，保证连接可用；不影响 NAPS2 图形界面（NAPS2.exe）。
    """
    import subprocess as _sp
    for name in ('NAPS2.Console.exe', 'NAPS2.Worker.exe'):
        try:
            _sp.run(['taskkill', '/F', '/T', '/IM', name],
                    capture_output=True, timeout=10)
        except Exception:
            continue


def _list_scan_devices(naps2_path, driver=None):
    """列出 NAPS2 当前可识别的扫描设备，用于扫描失败时的诊断提示。"""
    import subprocess as _sp
    for drv in ([driver] if driver else ['twain', 'wia']):
        try:
            r = _sp.run([naps2_path, '--listdevices', '--driver', drv],
                        capture_output=True, timeout=6)
            out = _decode_scan_output(r.stdout).strip()
            if out:
                return drv, [ln.strip() for ln in out.splitlines() if ln.strip()]
        except Exception:
            continue
    return None, []


def _collect_scan_pages(scan_output_dir, expected_filename):
    """收集一次扫描实际生成的全部页面文件。

    NAPS2 多页扫描会输出 scan1.1.png / scan1.2.png ...（单页为 scan1.png），
    这里按本次的基准编号收集所有页面并按页码排序，避免多页只取到一页。
    """
    base = os.path.splitext(expected_filename)[0]
    pages = []
    if not os.path.isdir(scan_output_dir):
        return pages
    for filename in os.listdir(scan_output_dir):
        if not filename.lower().endswith('.png'):
            continue
        stem = os.path.splitext(filename)[0]
        if stem == base or stem.startswith(base + '.'):
            pages.append(os.path.join(scan_output_dir, filename))

    def _sort_key(p):
        m = re.search(r'\.(\d+)\.png$', p.lower())
        return (1, int(m.group(1))) if m else (0, 0)

    pages.sort(key=_sort_key)
    return pages


def get_naps2_scan_options():
    """读取扫描参数（分辨率/纸张/色彩/来源）。

    重要：NAPS2 自带的默认档位是 100dpi + Letter，扫试卷时二维码与文字都会糊，
    因此这里给出适合试卷的推荐默认值，并允许在 naps2_config.json 中覆盖：
        {"source": "glass", "pagesize": "a4", "dpi": 300, "bitdepth": "color"}
    """
    opts = {'source': 'glass', 'pagesize': 'a4', 'dpi': 300, 'bitdepth': 'color'}
    try:
        if os.path.exists(NAPS2_CONFIG_FILE):
            with open(NAPS2_CONFIG_FILE, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            for key in list(opts.keys()):
                val = cfg.get(key)
                if val not in (None, ''):
                    opts[key] = val
    except Exception:
        pass
    return opts


def get_naps2_scan_target():
    """读取扫描驱动与设备名配置（可选）。

    返回 (driver, device)；未配置的项为 None，此时沿用 NAPS2 默认/图形界面选定的设备。
    """
    try:
        if os.path.exists(NAPS2_CONFIG_FILE):
            with open(NAPS2_CONFIG_FILE, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            driver = (cfg.get('driver') or '').strip() or None
            device = (cfg.get('device') or '').strip() or None
            return driver, device
    except Exception:
        pass
    return None, None


def _terminate_then_kill(proc):
    """结束扫描子进程，并清理可能残留的扫描工作进程。

    顺序：先请求结束 -> 等待 8 秒 -> 仍在则强制结束整棵进程树 -> 清理残留进程。
    这是保证"扫描仪连接被释放"的关键：残留进程会一直占着设备，
    导致之后每次扫描都报 TWAIN MaxConnections 或直接无响应。
    """
    if proc is None:
        return
    try:
        proc.terminate()
        proc.wait(timeout=3)
    except Exception:
        pass
    if proc.poll() is None:
        try:
            proc.kill()
            proc.wait(timeout=3)
        except Exception:
            pass
    _cleanup_stale_scanner_processes()


def _is_transient_scan_failure(returncode, text):
    """判断扫描失败是否属于"重试一次往往就好"的瞬时故障。

    瞬时故障：驱动/设备被占用、扫描仪休眠导致脱机、TWAIN 连接计数占满。
    而"没放纸"（返回码 0 + 0 页）属于操作问题，不做重试，直接提示更快。
    """
    if returncode not in (0, None):
        return True
    t = (text or '').lower()
    for kw in ('maxconnections', 'max connections', 'offline', '脱机', 'busy',
               'in use', 'occupied', '占用', 'unavailable', 'not available'):
        if kw in t:
            return True
    return False


def _run_scanner_reset(timeout=60):
    """调用 tools/reset_scanner.ps1 做一次扫描仪软复位（脚本内部自行提权）。

    返回 (ok: bool, detail: str)。
    """
    import subprocess as _sp
    import tempfile
    import time as _t

    script = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          'tools', 'reset_scanner.ps1')
    if not os.path.exists(script):
        return False, '未找到复位脚本：%s' % script

    log_path = os.path.join(tempfile.gettempdir(), 'qj_scanner_reset.log')
    try:
        if os.path.exists(log_path):
            os.remove(log_path)
    except Exception:
        pass

    try:
        _sp.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass',
                 '-File', script, '-LogPath', log_path],
                capture_output=True, timeout=timeout)
    except Exception as e:
        return False, '调用复位脚本失败：%s' % e

    # 提权后的副本是异步执行的，轮询日志判断最终结果
    for _ in range(40):
        _t.sleep(1)
        try:
            if os.path.exists(log_path):
                detail = io.open(log_path, encoding='utf-8', errors='replace').read()
                if 'RESULT=OK' in detail:
                    return True, detail.strip()
                if 'RESULT=FAIL' in detail:
                    return False, detail.strip()
        except Exception:
            continue
    return False, '复位脚本未在预期时间内返回结果（可能被系统策略拦截）'


@app.route('/api/scan', methods=['POST'])
def start_scan():
    """开始扫描（同一时刻只允许一次扫描，避免并发把扫描仪占死）"""
    import subprocess
    import time as _time

    # 并发保护：同一台扫描仪同一时刻只能有一个扫描会话。
    # 连续快速点击、或多人同时点击，会让 TWAIN 报 MaxConnections，
    # 并把设备拖入"无响应"状态（届时只能拔插 USB 才能恢复）。
    if not _SCAN_LOCK.acquire(blocking=False):
        return jsonify({
            'error': '已有一次扫描正在进行中，请等本次扫描结束后再点击。\n\n'
                     '（同一台扫描仪同一时刻只能扫描一次，'
                     '连续快速点击会让扫描仪被占用而卡住。）'
        }), 429

    try:
        # 扫描输出目录与目标文件名
        scan_output_dir = get_scan_output_dir()
        next_filename = get_next_scan_filename(scan_output_dir)
        scan_output_path = os.path.join(scan_output_dir, next_filename)

        naps2_path = get_naps2_path()
        if not os.path.exists(naps2_path):
            return jsonify({
                'error': '未找到 NAPS2 扫描程序，无法扫描。\n当前配置路径：%s\n'
                         '请在「扫描仪设置」中配置正确的 NAPS2.Console.exe 路径。' % naps2_path
            }), 400

        # 读取驱动/设备配置（可选）；未配置时沿用 NAPS2 默认设备
        naps2_driver, naps2_device = get_naps2_scan_target()
        # 读取扫描参数（分辨率/纸张/色彩/来源）；未配置时用适合试卷的推荐值
        scan_opts = get_naps2_scan_options()

        # 直接调用 NAPS2.Console.exe（不经 PowerShell 中转，避免中文报错被二次转码成乱码）。
        # --noprofile：不沿用图形界面最近使用的配置档案，
        #   避免档案里带有「自动保存/等待输入」等设置时，命令行扫描卡住不返回。
        scan_cmd = [naps2_path, '-o', scan_output_path, '-f', '-v', '--noprofile']
        if naps2_driver:
            scan_cmd += ['--driver', naps2_driver]
        if naps2_device:
            scan_cmd += ['--device', naps2_device]
        # 显式指定扫描参数：NAPS2 默认档位（100dpi/Letter）扫不清试卷二维码，
        # 必须明确指定，不能依赖默认档位或图形界面的历史设置。
        for _key, _flag in (('source', '--source'), ('pagesize', '--pagesize'),
                            ('dpi', '--dpi'), ('bitdepth', '--bitdepth')):
            _val = scan_opts.get(_key)
            if _val not in (None, ''):
                scan_cmd += [_flag, str(_val)]

        def _run_once():
            """执行一次扫描，返回 (returncode, 输出文本, 诊断码)。

            诊断码：ok / no_output（设备毫无回应）/ stall（启动后卡住）/ timeout。

            实测结论：NAPS2 的输出不缓冲，且故障时会在打印完
            "N page(s) scanned." 之后迟迟不退出。因此这里盯着输出流：
              · 出现 "N page(s) scanned." -> 扫描逻辑已结束：
                  N=0 -> 立即收尾；N>0 -> 等进程退出或输出文件落盘（宽限 10 秒），
                  只要文件生成即算成功（进程"假死"不影响结果）
              · EOF -> 正常结束
              · 20 秒内毫无输出 -> 设备无响应
              · 有输出后静默超过 45 秒 -> 判定卡住
              · 总时长超过 90 秒 -> 判定超时
            """
            import threading as _th
            import time as _tw

            _cleanup_stale_scanner_processes()
            print('执行扫描命令: %s' % ' '.join(scan_cmd), flush=True)
            proc = subprocess.Popen(scan_cmd, stdout=subprocess.PIPE,
                                    stderr=subprocess.STDOUT)
            buf = []
            done = _th.Event()
            marks = {'last': _tw.time(), 'n': 0, 'pages': None}

            def _read():
                try:
                    for line in iter(proc.stdout.readline, b''):
                        buf.append(line)
                        marks['last'] = _tw.time()
                        marks['n'] += 1
                        m = re.search(rb'(\d+)\s+page\(s\)\s+scanned', line.lower())
                        if m:
                            marks['pages'] = int(m.group(1))
                            done.set()
                except Exception:
                    pass
                finally:
                    done.set()

            _th.Thread(target=_read, daemon=True).start()

            started = _tw.time()
            diag = 'ok'
            while True:
                if done.wait(0.5):
                    diag = 'ok'
                    break
                now = _tw.time()
                if marks['n'] == 0 and now - started > SCAN_NO_OUTPUT_SECONDS:
                    diag = 'no_output'
                    break
                if marks['n'] > 0 and now - marks['last'] > SCAN_STALL_SECONDS:
                    diag = 'stall'
                    break
                if now - started > SCAN_TIMEOUT_SECONDS:
                    diag = 'timeout'
                    break

            if diag == 'ok':
                pages = marks.get('pages')
                if pages == 0:
                    # 没有内容需要保存，直接收尾
                    _terminate_then_kill(proc)
                else:
                    # 扫到了内容：等进程退出，或等输出文件落盘后收尾
                    deadline = _tw.time() + SCAN_FINISH_GRACE_SECONDS
                    while _tw.time() < deadline:
                        if proc.poll() is not None:
                            break
                        if _collect_scan_pages(scan_output_dir, next_filename):
                            break
                        _tw.sleep(0.4)
                    if proc.poll() is None:
                        _terminate_then_kill(proc)
            else:
                _terminate_then_kill(proc)

            try:
                proc.stdout.close()
            except Exception:
                pass
            return proc.returncode, _decode_scan_output(b''.join(buf)).strip(), diag

        returncode = None
        combined_output = ''
        generated_pages = []
        timed_out = False
        unresponsive = False

        for attempt in (1, 2):
            if attempt == 2:
                # 首次失败：清掉可能残留的扫描进程，稍等后再试一次（扫描仪休眠唤醒常需第二次）
                _cleanup_stale_scanner_processes()
                _time.sleep(1.5)
                print('首次扫描未成功，自动重试一次...', flush=True)
            returncode, combined_output, diag = _run_once()
            timed_out = (diag == 'timeout')
            unresponsive = diag in ('no_output', 'stall')
            if timed_out:
                break
            generated_pages = _collect_scan_pages(scan_output_dir, next_filename)
            if generated_pages:
                break
            # 以下情况不重试，直接给出结论（重试只会白等一轮）：
            #  · 设备毫无回应（unresponsive）：重试也打不开设备
            #  · 已经明确扫到 0 页（no_content）：这是确定性结果，不是瞬时故障
            #  · 不属于"被占用/脱机"等瞬时故障：例如没放纸等操作问题
            _out_low = (combined_output or '').lower()
            _no_content = ('0 page(s) scanned' in _out_low
                           or 'no scanned pages' in _out_low)
            if attempt == 1 and (unresponsive or _no_content
                                 or not _is_transient_scan_failure(returncode, combined_output)):
                break

        print('扫描命令返回码: %s' % returncode, flush=True)
        if combined_output:
            print('扫描命令输出: %s' % combined_output, flush=True)

        if timed_out or unresponsive or not generated_pages:
            if timed_out:
                err_source = 'timeout'
            elif diag == 'stall':
                err_source = '#scan_stall#'
            elif diag == 'no_output':
                err_source = '脱机'
            else:
                err_source = combined_output
            info = _humanize_scan_error(err_source, returncode)
            hint_lines = '\n'.join('%d. %s' % (i + 1, h) for i, h in enumerate(info['hints']))
            # 设备已经无响应时不再去枚举设备（枚举也会卡住，白等十几秒）
            if unresponsive:
                driver_used, devices = None, []
            else:
                driver_used, devices = _list_scan_devices(naps2_path, naps2_driver)
            if devices:
                device_text = ('\n\n系统可识别的扫描设备（%s 驱动）：\n' % driver_used
                               + '\n'.join('· %s' % d for d in devices))
            elif unresponsive:
                device_text = ''
            else:
                device_text = '\n\n未能列出任何扫描设备，请检查扫描仪电源与 USB 连接。'
            if diag == 'stall':
                detail = ('扫描已经开始，但之后 %d 秒内没有任何进展（驱动工作进程很可能已出错）。'
                          % SCAN_STALL_SECONDS)
            elif diag == 'no_output':
                detail = ('扫描仪在 %d 秒内没有任何回应（连"开始扫描"都没有打印出来），'
                          '说明当前无法与设备通讯。' % SCAN_NO_OUTPUT_SECONDS)
            else:
                detail = combined_output or '（扫描仪无任何输出，可能未响应或已休眠）'
            # 排查步骤里若已提到一键软复位，就不再重复追加
            extra = ('' if any('reset_scanner' in h for h in info['hints']) else
                     '\n\n也可运行项目 tools 目录下的 reset_scanner.bat 一键软复位'
                     '（无需拔线），仍无效再拔插 USB。')
            return jsonify({
                'error': ('%s\n\n请按以下步骤排查：\n%s%s%s\n\n技术详情：%s'
                          % (info['title'], hint_lines, extra, device_text, detail)),
                'returncode': returncode,
                'raw_output': combined_output,
                'output_dir': scan_output_dir,
            }), 500

        return jsonify({
            'success': True,
            'message': '扫描成功，共 %d 页' % len(generated_pages),
            'output_path': generated_pages[0],
            'output_paths': generated_pages,
            'filename': os.path.basename(generated_pages[0]),
            'page_count': len(generated_pages),
        })

    except subprocess.TimeoutExpired:
        info = _humanize_scan_error('timeout')
        hint_lines = '\n'.join('%d. %s' % (i + 1, h) for i, h in enumerate(info['hints']))
        return jsonify({'error': '%s\n\n请按以下步骤排查：\n%s' % (info['title'], hint_lines)}), 500
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        print('扫描失败: %s\n%s' % (e, error_trace), flush=True)
        msg = str(e)
        if 'WinError 5' in msg or '拒绝访问' in msg or 'PermissionError' in msg:
            friendly = ('扫描失败：保存扫描文件的目录没有写入权限。\n\n'
                        '请检查「扫描仪设置」中的扫描输出目录，'
                        '确认该目录存在且当前账户有读写权限，或改选一个可写的目录后重试。')
        elif 'FileNotFoundError' in msg or 'No such file' in msg:
            friendly = ('扫描失败：找不到指定的文件或目录。\n\n'
                        '请检查「扫描仪设置」中的 NAPS2 路径与扫描输出目录是否正确。')
        else:
            friendly = ('扫描失败，请检查扫描仪电源、USB 连接与纸张摆放后重试。\n\n'
                        '技术详情：%s' % msg)
        return jsonify({'error': friendly, 'raw_error': msg}), 500
    finally:
        # 无论成功失败，都清掉可能残留的扫描进程，避免其继续占用扫描仪
        try:
            _cleanup_stale_scanner_processes()
        finally:
            _SCAN_LOCK.release()


@app.route('/api/scan/reset', methods=['POST'])
def reset_scanner_device():
    """扫描仪软复位：重启图像采集服务，释放被占用的扫描仪连接。

    适用：反复提示「脱机 / 被占用 / 扫描仪未响应」时，
    在拔插 USB 之前先点这个（本机实测约 2~10 秒完成，无需拔线）。
    """
    if _SCAN_LOCK.locked():
        return jsonify({'error': '正在扫描中，无法复位。请等本次扫描结束后再试。'}), 409
    ok, detail = _run_scanner_reset()
    if ok:
        return jsonify({'success': True,
                        'message': '扫描仪已复位，请重新点击扫描。',
                        'detail': detail})
    return jsonify({
        'error': ('扫描仪软复位未成功。\n\n'
                  '请拔掉扫描仪 USB 线，等 10 秒后插回'
                  '（建议插主机后置 USB 接口），再点击扫描。'),
        'detail': detail,
    }), 500


@app.route('/api/import_local_images', methods=['POST'])
def import_local_images():
    """导入本地图片到扫描输出目录（制卡/阅卷 - 本地文件入口）"""
    from werkzeug.utils import secure_filename
    if 'files' not in request.files and not request.files.getlist('files'):
        return jsonify({'error': '请选择要导入的图片文件'}), 400
    scan_output_dir = get_scan_output_dir()
    os.makedirs(scan_output_dir, exist_ok=True)
    saved = 0
    for key in request.files:
        for f in request.files.getlist(key):
            if not f or not f.filename:
                continue
            ext = os.path.splitext(secure_filename(f.filename))[1].lower()
            if ext not in ('.jpg', '.jpeg', '.png'):
                continue
            base = datetime.now().strftime('%Y%m%d_%H%M%S')
            name = f"import_{base}_{saved}{ext}"
            path = os.path.join(scan_output_dir, name)
            try:
                f.save(path)
                saved += 1
            except Exception:
                pass
    if saved == 0:
        return jsonify({'error': '没有可保存的图片文件（仅支持 jpg/jpeg/png）'}), 400
    return jsonify({'success': True, 'count': saved})

@app.route('/api/check_conversion/<path:folder_path>')
def check_conversion_status(folder_path):
    """检查文件夹中的文件状态（已废弃，保留用于兼容性）"""
    try:
        # 判断是班级文件夹还是原始扫描文件夹
        parts = folder_path.split('/')
        if len(parts) == 2:
            # 班级文件夹：class_name/date_folder
            class_name, date_folder = parts[0], parts[1]
            folder_path_full = os.path.join(CLASSES_DIR, class_name, date_folder)
        else:
            # 原始扫描文件夹：date_folder
            folder_name = parts[0]
            folder_path_full = os.path.join(SCAN_DIR, folder_name)
        
        if not os.path.exists(folder_path_full):
            return jsonify({'error': '文件夹不存在'}), 404
        
        # 统计PNG文件
        png_files = [f for f in os.listdir(folder_path_full) 
                    if os.path.isfile(os.path.join(folder_path_full, f)) and f.lower().endswith('.png')]
        
        return jsonify({
            'total_tif': 0,
            'total_pdf': 0,
            'total': 0,
            'converted': 0,
            'remaining': 0,
            'remaining_files': []
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/check_new_files/<path:folder_path>')
def check_new_files(folder_path):
    """检查文件夹中是否有新文件（用于轮询检测）"""
    try:
        # 判断是否是扫描输出文件夹
        if folder_path == '__scan_output__':
            # 扫描输出文件夹
            folder_path_full = get_scan_output_dir()
        else:
            # 判断是班级文件夹还是原始扫描文件夹
            parts = folder_path.split('/')
            if len(parts) == 2:
                # 班级文件夹：class_name/date_folder
                class_name, date_folder = parts[0], parts[1]
                folder_path_full = os.path.join(CLASSES_DIR, class_name, date_folder)
            else:
                # 原始扫描文件夹：date_folder
                folder_name = parts[0]
                folder_path_full = os.path.join(SCAN_DIR, folder_name)
        
        if not os.path.exists(folder_path_full):
            return jsonify({'error': '文件夹不存在'}), 404
        
        # 获取所有图片文件及其修改时间（包括PDF）
        files_info = []
        for f in os.listdir(folder_path_full):
            file_path = os.path.join(folder_path_full, f)
            if os.path.isfile(file_path) and f.lower().endswith(('.jpg', '.jpeg', '.png')):
                
                stat_info = os.stat(file_path)
                files_info.append({
                    'name': f,
                    'path': file_path,
                    'modified_time': stat_info.st_mtime,
                    'size': stat_info.st_size
                })
        
        # 按修改时间排序（最新的在前）
        files_info.sort(key=lambda x: x['modified_time'], reverse=True)
        
        # 返回文件列表和最新文件的修改时间
        latest_time = files_info[0]['modified_time'] if files_info else 0
        
        return jsonify({
            'files': [f['name'] for f in files_info],
            'file_count': len(files_info),
            'latest_modified_time': latest_time,
            'has_new_files': len(files_info) > 0
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/files/batch_delete', methods=['POST'])
def batch_delete_files():
    """批量删除文件"""
    data = request.get_json()
    file_paths = data.get('file_paths', [])
    
    if not file_paths or not isinstance(file_paths, list):
        return jsonify({'error': '请提供要删除的文件路径列表'}), 400
    
    deleted_count = 0
    failed_files = []
    folders_to_check = set()  # 记录需要检查的文件夹路径（班级文件夹）
    
    for file_path in file_paths:
        try:
            parts = file_path.split('/')
            filename = parts[-1]
            
            # 判断是扫描输出文件夹、班级文件夹还是原始扫描文件夹
            if parts[0] == '__scan_output__':
                # 扫描输出文件夹：__scan_output__/filename
                file_path_full = os.path.join(get_scan_output_dir(), filename)
                folder_path_full = None  # 扫描输出文件夹不检查
            elif len(parts) == 3:
                # 班级文件夹：class_name/date_folder/filename
                class_name, date_folder = parts[0], parts[1]
                file_path_full = os.path.join(CLASSES_DIR, class_name, date_folder, filename)
                folder_path_full = os.path.join(CLASSES_DIR, class_name, date_folder)
                folders_to_check.add(folder_path_full)  # 记录需要检查的文件夹
            elif len(parts) == 2:
                # 原始扫描文件夹：date_folder/filename
                folder_name = parts[0]
                file_path_full = os.path.join(SCAN_DIR, folder_name, filename)
                folder_path_full = None  # 扫描文件夹不检查
            else:
                failed_files.append({'file': file_path, 'error': '无效的文件路径'})
                continue
            
            # 检查文件是否存在
            if not os.path.exists(file_path_full):
                failed_files.append({'file': file_path, 'error': '文件不存在'})
                continue
            
            # 删除文件
            try:
                os.remove(file_path_full)
                
                # 检查是否有处理过的版本（临时文件）
                name, ext = os.path.splitext(filename)
                processed_filename = f"{name}_processed{ext}"
                temp_processed_path = os.path.join(TEMP_DIR, processed_filename)
                if os.path.exists(temp_processed_path):
                    try:
                        os.remove(temp_processed_path)
                    except Exception:
                        pass  # 忽略删除临时文件失败
                
                deleted_count += 1
            except PermissionError:
                failed_files.append({'file': file_path, 'error': '文件正在被使用，无法删除'})
            except Exception as e:
                failed_files.append({'file': file_path, 'error': f'删除失败: {str(e)}'})
        except Exception as e:
            failed_files.append({'file': file_path, 'error': f'处理失败: {str(e)}'})
    
    # 检查班级文件夹是否为空，如果为空则删除
    deleted_folders = []
    for folder_path_full in folders_to_check:
        if os.path.exists(folder_path_full):
            try:
                # 检查文件夹中是否还有图片文件
                remaining_files = [f for f in os.listdir(folder_path_full) 
                                 if os.path.isfile(os.path.join(folder_path_full, f)) 
                                 and f.lower().endswith(('.jpg', '.jpeg', '.png'))]
                if not remaining_files:
                    # 文件夹为空，尝试删除
                    try:
                        os.rmdir(folder_path_full)  # 只删除空文件夹
                        deleted_folders.append(os.path.basename(folder_path_full))
                    except Exception:
                        pass  # 删除失败不影响结果
            except Exception:
                pass  # 检查失败不影响结果
    
    result = {
        'success': True,
        'deleted_count': deleted_count,
        'total_count': len(file_paths),
        'failed_count': len(failed_files)
    }
    
    if deleted_folders:
        result['deleted_folders'] = deleted_folders
    
    if failed_files:
        result['failed_files'] = failed_files[:10]  # 只返回前10个失败的文件
    
    return jsonify(result)

@app.route('/api/file/<path:file_path>', methods=['DELETE'])
def delete_file(file_path):
    """删除指定的图片文件"""
    try:
        parts = file_path.split('/')
        filename = parts[-1]
        folder_path_full = None  # 用于记录文件夹路径（如果是班级文件夹）
        
        # 判断是扫描输出文件夹、班级文件夹还是原始扫描文件夹
        if parts[0] == '__scan_output__':
            # 扫描输出文件夹：__scan_output__/filename
            file_path_full = os.path.join(get_scan_output_dir(), filename)
        elif len(parts) == 3:
            # 班级文件夹：class_name/date_folder/filename
            class_name, date_folder = parts[0], parts[1]
            file_path_full = os.path.join(CLASSES_DIR, class_name, date_folder, filename)
            folder_path_full = os.path.join(CLASSES_DIR, class_name, date_folder)
        elif len(parts) == 2:
            # 原始扫描文件夹：date_folder/filename
            folder_name = parts[0]
            file_path_full = os.path.join(SCAN_DIR, folder_name, filename)
        else:
            return jsonify({'error': '无效的文件路径'}), 400
        
        # 检查文件是否存在
        if not os.path.exists(file_path_full):
            return jsonify({'error': '文件不存在'}), 404
        
        # 删除文件
        try:
            os.remove(file_path_full)
            
            # 检查是否有处理过的版本（临时文件）
            name, ext = os.path.splitext(filename)
            processed_filename = f"{name}_processed{ext}"
            temp_processed_path = os.path.join(TEMP_DIR, processed_filename)
            if os.path.exists(temp_processed_path):
                try:
                    os.remove(temp_processed_path)
                except Exception:
                    pass  # 忽略删除临时文件失败
            
            # 如果是班级文件夹，检查文件夹是否为空，如果为空则删除
            folder_deleted = False
            if folder_path_full and os.path.exists(folder_path_full):
                try:
                    # 检查文件夹中是否还有图片文件
                    remaining_files = [f for f in os.listdir(folder_path_full) 
                                     if os.path.isfile(os.path.join(folder_path_full, f)) 
                                     and f.lower().endswith(('.jpg', '.jpeg', '.png'))]
                    if not remaining_files:
                        # 文件夹为空，尝试删除
                        try:
                            os.rmdir(folder_path_full)  # 只删除空文件夹
                            folder_deleted = True
                        except Exception:
                            pass  # 删除失败不影响结果
                except Exception:
                    pass  # 检查失败不影响结果
            
            result = {'success': True, 'message': '文件删除成功'}
            if folder_deleted:
                result['folder_deleted'] = True
            return jsonify(result)
        except PermissionError:
            return jsonify({'error': '文件正在被使用，无法删除'}), 403
        except Exception as e:
            return jsonify({'error': f'删除文件失败: {str(e)}'}), 500
            
    except Exception as e:
        return jsonify({'error': f'删除文件失败: {str(e)}'}), 500

@app.route('/api/image/<path:file_path>')
def get_image(file_path):
    """获取图片文件（支持班级文件夹路径、扫描输出文件夹和临时文件夹）"""
    parts = file_path.split('/')
    
    # 检查是否是临时文件
    if len(parts) >= 2 and parts[0] == 'temp':
        filename = parts[-1]
        temp_path = os.path.join(TEMP_DIR, filename)
        if os.path.exists(temp_path):
            return send_from_directory(TEMP_DIR, filename)
        else:
            return jsonify({'error': '临时文件不存在'}), 404

    # 任务试卷图片：task/<task_id>/<filename>
    if len(parts) >= 3 and parts[0] == 'task':
        task_id = parts[1]
        filename = parts[-1]
        folder_path = os.path.join(TASK_PAPERS_DIR, task_id)
        full_path = os.path.join(folder_path, filename)
        if os.path.isfile(full_path) and os.path.realpath(full_path).startswith(os.path.realpath(folder_path)):
            return send_from_directory(folder_path, filename)
        return jsonify({'error': '文件不存在'}), 404
    
    if len(parts) >= 2:
        filename = parts[-1]
        # 判断是扫描输出文件夹、班级文件夹还是原始扫描文件夹
        if parts[0] == '__scan_output__':
            # 扫描输出文件夹：__scan_output__/filename
            folder_path = get_scan_output_dir()
        elif len(parts) == 3:
            # 班级文件夹：class_name/date_folder/filename
            class_name, date_folder = parts[0], parts[1]
            folder_path = os.path.join(CLASSES_DIR, class_name, date_folder)
        else:
            # 原始扫描文件夹：date_folder/filename
            folder_name = parts[0]
            folder_path = os.path.join(SCAN_DIR, folder_name)
        
        # 检查临时文件夹中是否有processed版本
        name, ext = os.path.splitext(filename)
        processed_filename = f"{name}_processed{ext}"
        processed_path = os.path.join(TEMP_DIR, processed_filename)
        
        # 优先返回临时文件夹中的processed版本，如果不存在则返回原图
        if os.path.exists(processed_path):
            return send_from_directory(TEMP_DIR, processed_filename)
        elif os.path.exists(os.path.join(folder_path, filename)):
            return send_from_directory(folder_path, filename)
    
    return jsonify({'error': '文件不存在'}), 404

def save_to_class_center(individual_reports, class_evaluation, logger=None):
    """
    将批阅结果保存到班级中心
    
    Args:
        individual_reports: 个人报告字典 {file_path: report_data}
        class_evaluation: 班级整体评价
        logger: 日志记录器（可选）
    """
    try:
        current_date = datetime.now().strftime('%Y-%m-%d')
        timestamp = datetime.now().isoformat()
        
        # 按班级组织数据
        class_data = {}  # {class_name: {student_id: [records]}}
        
        for file_path, report_data in individual_reports.items():
            if report_data.get('status') != 'success':
                continue
            
            student_id = report_data.get('student_id', '')
            if not student_id or len(student_id) != 6 or not student_id.isdigit():
                continue
            
            # 查找学生所属班级
            class_name = find_class_by_student_id(student_id, logger)
            if not class_name:
                continue
            
            # 初始化班级数据结构
            if class_name not in class_data:
                class_data[class_name] = {}
            if student_id not in class_data[class_name]:
                class_data[class_name][student_id] = []
            
            # 创建记录
            record = {
                'date': current_date,
                'timestamp': timestamp,
                'filename': report_data.get('filename', ''),
                'file_path': file_path,
                'essay_text': report_data.get('essay_text', ''),  # 作文原文
                'report': report_data.get('report', ''),
                'class_evaluation': class_evaluation
            }
            
            class_data[class_name][student_id].append(record)
        
        # 保存到班级中心文件
        for class_name, students_data in class_data.items():
            class_center_file = os.path.join(CLASS_CENTER_DIR, f"{class_name}.json")
            
            # 读取现有数据
            if os.path.exists(class_center_file):
                with open(class_center_file, 'r', encoding='utf-8') as f:
                    existing_data = json.load(f)
            else:
                existing_data = {}
            
            # 合并新数据
            for student_id, records in students_data.items():
                if student_id not in existing_data:
                    existing_data[student_id] = []
                existing_data[student_id].extend(records)
            
            # 按时间戳排序（最新的在前）
            for student_id in existing_data:
                existing_data[student_id].sort(key=lambda x: x.get('timestamp', ''), reverse=True)
            
            # 保存
            with open(class_center_file, 'w', encoding='utf-8') as f:
                json.dump(existing_data, f, ensure_ascii=False, indent=2)
            
            if logger:
                logger.info(f"已保存 {len(students_data)} 个学生的记录到班级 {class_name} 的中心")
    
    except Exception as e:
        if logger:
            logger.error(f"保存到班级中心失败: {e}")

def find_class_by_student_id(student_id, logger=None):
    """
    根据学号查找对应的班级
    
    Args:
        student_id: 6位数字学号
        logger: 日志记录器（可选）
    
    Returns:
        str: 班级名称，如果未找到返回 None
    """
    if not student_id or len(student_id) != 6 or not student_id.isdigit():
        return None
    
    try:
        # 遍历所有班级的学生列表文件
        for class_file in os.listdir(STUDENTS_DIR):
            if not class_file.endswith('.json'):
                continue
            
            class_name = class_file[:-5]  # 移除 .json 后缀
            students_file = os.path.join(STUDENTS_DIR, class_file)
            
            try:
                with open(students_file, 'r', encoding='utf-8') as f:
                    students = json.load(f)
                
                # 检查学号是否在列表中（支持 list 或 dict）
                if isinstance(students, list) and student_id in students:
                    if logger:
                        logger.info(f"找到学号 {student_id} 对应的班级: {class_name}")
                    return class_name
                if isinstance(students, dict) and student_id in students:
                    if logger:
                        logger.info(f"找到学号 {student_id} 对应的班级: {class_name}")
                    return class_name
            except Exception as e:
                if logger:
                    logger.warning(f"读取班级 {class_name} 的学生列表失败: {e}")
                continue
        
        if logger:
            logger.info(f"未找到学号 {student_id} 对应的班级")
        return None
    except Exception as e:
        if logger:
            logger.error(f"查找班级时出错: {e}")
        return None

_FULLWIDTH_DIGITS = str.maketrans('０１２３４５６７８９', '0123456789')


def _squash_digit_gaps(s):
    """把分格书写、被空格/点线切开的数字合并：'1 4 3 1' / '1.4.3.1' / '1|4|3|1' -> '1431'。

    手写学号常常一格一个数字，OCR 会原样带出分隔符，不合并就永远匹配不到 4 位学号。
    """
    s = str(s or '').translate(_FULLWIDTH_DIGITS)
    prev = None
    while prev != s:
        prev = s
        s = re.sub(r'(?<=\d)[\s.·|:：\-_/\\]+(?=\d)', '', s)
    return s


def extract_student_id(text):
    """从 OCR 文本中提取学号，兼容 4 位（新版）与 6 位（旧版）。

    优先看「学号」标签后面的内容（最可靠）；没有再退化为扫描前几行里的独立数字。
    """
    if not text:
        return None
    raw = str(text).translate(_FULLWIDTH_DIGITS)

    # 1) 带标签：学号 / 學號 / 考号 / Student ID / ID
    label = re.search(r'(?:学\s*号|學\s*號|考\s*号|Student\s*ID|SID|ID)\s*[：:]?\s*([0-9\s.·|_\-]{3,24})',
                      raw, re.IGNORECASE)
    if label:
        cand = _squash_digit_gaps(label.group(1))
        m = re.search(r'(?<!\d)(\d{6}|\d{4})(?!\d)', cand)
        if m:
            return m.group(1)

    # 2) 前 6 行里的独立数字（先合并分格书写，再匹配）
    for line in raw.split('\n')[:6]:
        cand = _squash_digit_gaps(line)
        for m in re.finditer(r'(?<!\d)(\d{6}|\d{4})(?!\d)', cand):
            v = m.group(1)
            # 4 位数字容易被年份、日期、分值误命中，跳过明显像年份/日期的
            if len(v) == 4 and (v[:2] in ('19', '20') or v in ('0000',)):
                continue
            return v

    return None

def translate_essay_to_chinese(text, logger=None):
    """将识别出的英文作文译为中文，供识别结果旁展示译文。调用模型时在识别结果旁输出译文。"""
    if not text or not text.strip() or not minimax_api_key:
        return ''
    try:
        client = OpenAI(base_url=minimax_base_url, api_key=minimax_api_key)
        response = client.chat.completions.create(
            model=llm_text_model,
            messages=[
                {"role": "system", "content": "你是翻译助手。将用户给出的英文作文或文本翻译成中文，保持段落与语气，不要批注或点评，只输出译文。"},
                {"role": "user", "content": text.strip()}
            ],
            temperature=0.3,
            max_tokens=3000,
            **_text_model_kwargs()
        )
        out = (response.choices[0].message.content or '').strip()
        if logger and out:
            logger.info("译文已生成")
        return out
    except Exception as e:
        if logger:
            logger.warning(f"译文生成失败: {e}")
        return ''

@app.route('/api/my_reports', methods=['GET'])
def get_my_reports():
    """学生端：获取当前登录学生本人的历次批阅记录"""
    if session.get('user_role') != 'student':
        return jsonify({'error': '仅学生可查看'}), 403
    student_id = session.get('user_id', '')
    if not student_id:
        return jsonify({'records': []})
    records = []
    try:
        for class_file in os.listdir(STUDENTS_DIR):
            if not class_file.endswith('.json'):
                continue
            class_name = class_file[:-5]
            students_file = os.path.join(STUDENTS_DIR, class_file)
            with open(students_file, 'r', encoding='utf-8') as f:
                students = json.load(f)
            if not isinstance(students, list) or student_id not in students:
                continue
            center_file = os.path.join(CLASS_CENTER_DIR, f"{class_name}.json")
            if not os.path.exists(center_file):
                continue
            with open(center_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            for r in data.get(student_id, []):
                r = dict(r)
                r['class_name'] = class_name
                records.append(r)
        records.sort(key=lambda x: x.get('timestamp', ''), reverse=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    return jsonify({'records': records})

@app.route('/api/students/<class_name>', methods=['GET'])
def get_students(class_name):
    """获取班级学生列表（students 为纯净学号列表；student_names 为学号->姓名，供展示「学号(姓名)」用）"""
    students_file = os.path.join(STUDENTS_DIR, f"{class_name}.json")
    if os.path.exists(students_file):
        with open(students_file, 'r', encoding='utf-8') as f:
            raw = json.load(f)
        student_ids = [str(s) for s in raw] if isinstance(raw, list) else list(raw.keys()) if isinstance(raw, dict) else []
        names_map = get_student_names_map()
        student_names = {sid: names_map.get(sid, '') for sid in student_ids}
        return jsonify({'students': student_ids, 'student_names': student_names})
    return jsonify({'students': [], 'student_names': {}})

@app.route('/api/students/<class_name>', methods=['POST'])
def save_students(class_name):
    """保存班级学生列表（学号为6位数字）"""
    data = request.get_json()
    students = data.get('students', [])
    
    # 验证学号格式（6位数字）
    invalid_ids = []
    for student_id in students:
        if not isinstance(student_id, str) or not student_id.isdigit() or len(student_id) != 6:
            invalid_ids.append(student_id)
    
    if invalid_ids:
        return jsonify({'error': f'以下学号格式不正确（应为6位数字）：{", ".join(invalid_ids)}'}), 400
    
    students_file = os.path.join(STUDENTS_DIR, f"{class_name}.json")
    
    with open(students_file, 'w', encoding='utf-8') as f:
        json.dump(students, f, ensure_ascii=False, indent=2)
    
    return jsonify({'msg': '学生列表已保存'})

@app.route('/api/ocr', methods=['POST'])
def ocr_recognize():
    """OCR文字识别"""
    data = request.get_json()
    files = data.get('files', [])
    
    if not files:
        return jsonify({'error': '请选择文件'}), 400
    
    # 配置日志
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    
    try:
        results = {}
        for file_path in files:
            parts = file_path.split('/')
            filename = parts[-1]
            # 判断是班级文件夹还是原始扫描文件夹
            if len(parts) >= 3 and parts[0] == 'task':
                # 任务文件夹：task/task_id/filename
                task_id = parts[1]
                folder_path = os.path.join(TASK_PAPERS_DIR, task_id)
            elif len(parts) >= 2 and parts[0] == '__scan_output__':
                folder_path = get_scan_output_dir()
            elif len(parts) == 3:
                # 班级文件夹：class_name/date_folder/filename
                class_name, date_folder = parts[0], parts[1]
                folder_path = os.path.join(CLASSES_DIR, class_name, date_folder)
            else:
                # 原始扫描文件夹：date_folder/filename
                folder_name = parts[0]
                folder_path = os.path.join(SCAN_DIR, folder_name)
            
            # 优先检查临时文件夹中是否有processed版本（总是使用处理过的版本，如果存在）
            name, ext = os.path.splitext(filename)
            processed_filename = f"{name}_processed{ext}"
            processed_path = os.path.join(TEMP_DIR, processed_filename)
            original_path = os.path.join(folder_path, filename)
            
            # 记录所有尝试的路径（用于错误日志）
            checked_paths = []
            
            # 确定使用的文件路径（优先使用处理过的版本）
            full_path = None
            use_processed = False
            if os.path.exists(processed_path):
                # 使用临时文件夹中的processed版本
                full_path = processed_path
                use_processed = True
                logger.info(f"OCR识别 - 使用处理过的版本: {processed_path}")
            elif os.path.exists(original_path):
                full_path = original_path
                logger.info(f"OCR识别 - 使用原始文件: {original_path}")
            else:
                checked_paths.append(f"处理过的版本: {processed_path} (不存在)")
                checked_paths.append(f"原始文件: {original_path} (不存在)")
                # 如果原文件不存在，检查TIF/PNG转换或PDF/PNG转换
                if original_path.lower().endswith(('.tif', '.tiff')):
                    png_path = os.path.splitext(original_path)[0] + '.png'
                    checked_paths.append(f"PNG转换文件: {png_path}")
                    if os.path.exists(png_path):
                        full_path = png_path
                        new_filename = os.path.basename(png_path)
                        file_path = '/'.join(parts[:-1] + [new_filename]) if len(parts) > 1 else new_filename
                        logger.info(f"OCR识别 - 使用PNG转换文件: {png_path}")
                    else:
                        error_msg = f"文件不存在。原始路径: {file_path}\n尝试的路径:\n  - {processed_path} (不存在)\n  - {original_path} (不存在)\n  - {png_path} (不存在)\n文件夹路径: {folder_path}"
                        logger.error(f"OCR识别失败 - {error_msg}")
                        results[file_path] = {'error': '文件不存在'}
                        continue
                elif original_path.lower().endswith('.pdf'):
                    # PDF可能已转换为单页或多页PNG
                    base_name = os.path.splitext(original_path)[0]
                    png_path = base_name + '.png'  # 单页PNG
                    png_path_page1 = base_name + '_page1.png'  # 多页PNG第一页
                    checked_paths.append(f"PNG转换文件（单页）: {png_path}")
                    checked_paths.append(f"PNG转换文件（多页第一页）: {png_path_page1}")
                    if os.path.exists(png_path):
                        full_path = png_path
                        new_filename = os.path.basename(png_path)
                        file_path = '/'.join(parts[:-1] + [new_filename]) if len(parts) > 1 else new_filename
                        logger.info(f"OCR识别 - 使用PDF转换的PNG文件（单页）: {png_path}")
                    elif os.path.exists(png_path_page1):
                        full_path = png_path_page1
                        new_filename = os.path.basename(png_path_page1)
                        file_path = '/'.join(parts[:-1] + [new_filename]) if len(parts) > 1 else new_filename
                        logger.info(f"OCR识别 - 使用PDF转换的PNG文件（多页第一页）: {png_path_page1}")
                    else:
                        error_msg = f"文件不存在。原始路径: {file_path}\n尝试的路径:\n  - {processed_path} (不存在)\n  - {original_path} (不存在)\n  - {png_path} (不存在)\n  - {png_path_page1} (不存在)\n文件夹路径: {folder_path}"
                        logger.error(f"OCR识别失败 - {error_msg}")
                        results[file_path] = {'error': '文件不存在'}
                        continue
                else:
                    error_msg = f"文件不存在。原始路径: {file_path}\n尝试的路径:\n  - {processed_path} (不存在)\n  - {original_path} (不存在)\n文件夹路径: {folder_path}"
                    logger.error(f"OCR识别失败 - {error_msg}")
                    results[file_path] = {'error': '文件不存在'}
                    continue
            
            if not full_path:
                error_msg = f"文件不存在。原始路径: {file_path}\n尝试的路径:\n  - {processed_path} (不存在)\n  - {original_path} (不存在)\n文件夹路径: {folder_path}"
                logger.error(f"OCR识别失败 - {error_msg}")
                results[file_path] = {'error': '文件不存在'}
                continue
            
            try:
                # 如果文件是TIF格式且存在，先转换为PNG
                # TIF转换功能已移除，扫描仪现在直接生成PNG格式
                
                # 读取图片
                image = Image.open(full_path)
                
                # TIF格式处理：确保图片模式正确
                # 某些TIF可能是灰度图或带透明通道，需要转换为RGB
                if image.mode in ('RGBA', 'LA', 'P'):
                    # 如果有透明通道，先转换为RGB
                    background = Image.new('RGB', image.size, (255, 255, 255))
                    if image.mode == 'P':
                        image = image.convert('RGBA')
                    if image.mode in ('RGBA', 'LA'):
                        background.paste(image, mask=image.split()[-1] if image.mode == 'RGBA' else None)
                        image = background
                    else:
                        image = image.convert('RGB')
                elif image.mode not in ('RGB', 'L'):
                    image = image.convert('RGB')
                
                # 如果是灰度图，转换为RGB（某些OCR模型需要RGB）
                if image.mode == 'L':
                    image = image.convert('RGB')
                
                # 图像预处理：提高OCR识别率
                # 1. 如果图片太大，适当缩小（保持长边不超过2000像素）
                max_size = 2000
                if max(image.size) > max_size:
                    ratio = max_size / max(image.size)
                    new_size = (int(image.size[0] * ratio), int(image.size[1] * ratio))
                    image = image.resize(new_size, Image.Resampling.LANCZOS)
                
                # OCR 识别：使用大模型视觉识图（原讯飞OCR已过期停用）
                result = None
                text_from_result = ''

                try:
                    logger.info(f"使用大模型识图识别文件: {filename}")
                    # 将PIL Image转换为字节数据
                    img_io = io.BytesIO()
                    image.save(img_io, format='PNG')
                    img_bytes = img_io.getvalue()

                    text_from_result = llm_ocr_recognize(img_bytes, logger)
                    if text_from_result and text_from_result.strip():
                        result = {'text': text_from_result}
                        logger.info(f"大模型识图 识别成功，文本长度: {len(text_from_result)}")
                    else:
                        logger.warning("大模型识图 返回空文本")
                except Exception as ocr_error:
                    logger.error(f"大模型识图 识别失败: {ocr_error}")
                    text_from_result = ''
                    result = None

                # 处理识别返回的文本
                text = text_from_result if text_from_result else ''
                
                # 尝试从文本中提取学号
                student_id = extract_student_id(text)
                
                logger.info(f"Extracted text length: {len(text)}, student_id: {student_id}")
                
                # 调用模型生成译文，识别结果旁输出译文
                translation = ''
                if text and text.strip():
                    try:
                        translation = translate_essay_to_chinese(text, logger)
                    except Exception as te:
                        if logger:
                            logger.warning(f"译文生成失败: {te}")
                
                # 如果有学号，尝试自动归类到班级文件夹
                auto_organized = False
                auto_organize_info = None
                if student_id:
                    # 查找包含该学号的班级
                    class_name = find_class_by_student_id(student_id, logger)
                    if class_name:
                        # 获取当前日期作为文件夹名
                        date_folder = datetime.now().strftime('%Y-%m-%d')
                        # 尝试自动归类
                        try:
                            target_path = os.path.join(CLASSES_DIR, class_name, date_folder)
                            os.makedirs(target_path, exist_ok=True)
                            
                            # 移动文件到班级文件夹
                            target_file = os.path.join(target_path, filename)
                            if not os.path.exists(target_file):
                                import shutil
                                shutil.move(full_path, target_file)
                                auto_organized = True
                                auto_organize_info = {
                                    'class_name': class_name,
                                    'date_folder': date_folder,
                                    'new_path': f"{class_name}/{date_folder}/{filename}"
                                }
                                logger.info(f"文件已自动归类到班级 {class_name} 的 {date_folder} 文件夹")
                            else:
                                logger.info(f"目标文件已存在，跳过自动归类: {target_file}")
                        except Exception as org_error:
                            logger.warning(f"自动归类失败: {org_error}")
                
                results[file_path] = {
                    'text': text,
                    'translation': translation,
                    'student_id': student_id,
                    'status': 'success',
                    'auto_organized': auto_organized,
                    'auto_organize_info': auto_organize_info
                }
                
                # 如果使用了临时文件夹中的processed版本，识别完成后删除临时文件
                if use_processed and os.path.exists(processed_path):
                    try:
                        os.remove(processed_path)
                        logger.info(f"已删除临时文件: {processed_path}")
                    except Exception as del_error:
                        logger.warning(f"删除临时文件失败: {del_error}")
            except Exception as e:
                error_msg = f"OCR识别处理失败: {str(e)}\n文件路径: {file_path}\n使用的完整路径: {full_path if full_path else '未确定'}\n尝试的路径:\n  - 处理过的版本: {processed_path} ({'存在' if os.path.exists(processed_path) else '不存在'})\n  - 原始文件: {original_path} ({'存在' if os.path.exists(original_path) else '不存在'})\n文件夹路径: {folder_path}"
                logger.error(f"OCR识别异常 - {error_msg}")
                results[file_path] = {
                    'error': f'OCR识别失败: {str(e)}',
                    'status': 'error'
                }
                # 即使识别失败，如果使用了临时文件，也尝试删除
                if use_processed and os.path.exists(processed_path):
                    try:
                        os.remove(processed_path)
                    except:
                        pass
        
        # 保存识别结果
        result_file = os.path.join(results_dir, f"ocr_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
        with open(result_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        
        # 保存到历史记录（用于下次打开时恢复）
        history_file = os.path.join(HISTORY_DIR, f"history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
        history_data = {
            'type': 'ocr',
            'result_file': result_file,
            'results': results,
            'timestamp': datetime.now().isoformat()
        }
        with open(history_file, 'w', encoding='utf-8') as f:
            json.dump(history_data, f, ensure_ascii=False, indent=2)
        
        return jsonify({'results': results, 'result_file': result_file, 'history_file': history_file})
    
    except Exception as e:
        import traceback
        error_msg = str(e)
        error_trace = traceback.format_exc()
        logger.error(f"OCR识别异常: {error_msg}\n{error_trace}")
        return jsonify({'error': f'OCR识别失败: {error_msg}'}), 500

def _get_subject_config():
    """读取学科配置，用于扩展多学科时按学科选提示词等。当前仅 english。"""
    try:
        if os.path.exists(SUBJECT_CONFIG_FILE):
            with open(SUBJECT_CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return {'current': 'english', 'subjects': {'english': {'name': '英语', 'prompt_file': 'prompt_template.txt', 'paper_preset': 'english'}}}

def get_prompt_template(subject=None):
    """获取提示词模板。subject 为空时使用 subject_config 中的 current 对应学科的 prompt_file。"""
    cfg = _get_subject_config()
    sid = (subject or cfg.get('current') or 'english').strip() or 'english'
    subjects = cfg.get('subjects') or {}
    prompt_file = (subjects.get(sid) or {}).get('prompt_file') or 'prompt_template.txt'
    path = os.path.join(CONFIG_DIR, prompt_file)
    try:
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                return f.read().strip()
    except Exception:
        pass
    if os.path.exists(PROMPT_TEMPLATE_FILE):
        try:
            with open(PROMPT_TEMPLATE_FILE, 'r', encoding='utf-8') as f:
                return f.read().strip()
        except Exception:
            pass
    return DEFAULT_PROMPT_TEMPLATE

def save_prompt_template(template):
    """保存提示词模板"""
    try:
        with open(PROMPT_TEMPLATE_FILE, 'w', encoding='utf-8') as f:
            f.write(template)
        return True
    except Exception:
        return False

@app.route('/api/prompt_template', methods=['GET'])
def get_prompt_template_api():
    """获取提示词模板"""
    template = get_prompt_template()
    return jsonify({'template': template})

@app.route('/api/prompt_template', methods=['POST'])
def save_prompt_template_api():
    """保存提示词模板"""
    data = request.get_json()
    template = data.get('template', '').strip()
    
    if not template:
        return jsonify({'error': '提示词模板不能为空'}), 400
    
    if save_prompt_template(template):
        return jsonify({'success': True, 'message': '提示词模板已保存'})
    else:
        return jsonify({'error': '保存失败'}), 500

@app.route('/api/prompt_template', methods=['DELETE'])
def delete_prompt_template_api():
    """删除提示词模板文件，恢复默认值"""
    try:
        if os.path.exists(PROMPT_TEMPLATE_FILE):
            os.remove(PROMPT_TEMPLATE_FILE)
        return jsonify({'success': True, 'template': DEFAULT_PROMPT_TEMPLATE})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def _resolve_file_path_to_full(file_path):
    """将前端传的 file_path 解析为本地完整路径（仅原图，不查 temp processed）。返回 (full_path, error_msg)，失败时 full_path 为 None。"""
    parts = file_path.split('/')
    filename = parts[-1]
    if len(parts) >= 3 and parts[0] == 'task':
        task_id = parts[1]
        folder_path = os.path.join(TASK_PAPERS_DIR, task_id)
    elif len(parts) >= 2 and parts[0] == '__scan_output__':
        folder_path = get_scan_output_dir()
    elif len(parts) == 3:
        class_name, date_folder = parts[0], parts[1]
        folder_path = os.path.join(CLASSES_DIR, class_name, date_folder)
    else:
        folder_name = parts[0]
        folder_path = os.path.join(SCAN_DIR, folder_name)
    original_path = os.path.join(folder_path, filename)
    if os.path.exists(original_path):
        return (original_path, None)
    if original_path.lower().endswith(('.tif', '.tiff')):
        png_path = os.path.splitext(original_path)[0] + '.png'
        if os.path.exists(png_path):
            return (png_path, None)
    return (None, f'文件不存在: {file_path}')

@app.route('/api/vision_grade', methods=['POST'])
def vision_grade():
    """一步方案：直接看图识别+批阅（使用 MiniMax 多模态，不依赖讯飞 OCR）"""
    data = request.get_json()
    files = data.get('files', [])
    custom_prompt_template = data.get('prompt_template', None)
    save_to_class_center_flag = data.get('save_to_class_center', True)
    if not files:
        return jsonify({'error': '请选择文件'}), 400
    if not minimax_api_key:
        return jsonify({'error': '未配置 LLM_API_KEY 或 MINIMAX_API_KEY，请在 .env 中设置'}), 500
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    try:
        # 视觉批阅需上传图片+长文生成，使用较长超时；连接超时多为网络/服务响应慢，非接口方式问题
        client = OpenAI(base_url=minimax_base_url, api_key=minimax_api_key, timeout=120.0)
        prompt_template = custom_prompt_template or get_prompt_template()
        # 一步提示：要求模型看图后按格式输出识别文本与批阅报告
        vision_instruction = (
            '请仔细观察图片中的手写作文内容，完成以下两项并严格按格式回复：\n'
            '1. 【识别文本】\n（此处仅输出从图片中识别出的全部手写文字，不要添加其他说明）\n'
            '2. 【批阅报告】\n（此处输出你对这篇作文的批阅报告，包含评语与建议）\n'
            '若能从图中或文字中看出学号（如6位数字），请在识别文本中保留。'
        )
        user_prompt = prompt_template.replace('{essay_text}', '[请根据图片中的作文内容进行识别并批阅]') + '\n\n' + vision_instruction
        individual_reports = {}
        for file_path in files:
            full_path, err = _resolve_file_path_to_full(file_path)
            if err or not full_path:
                individual_reports[file_path] = {'error': err or '文件不存在', 'status': 'error'}
                continue
            filename = os.path.basename(file_path)
            try:
                image = Image.open(full_path)
                if image.mode in ('RGBA', 'LA', 'P'):
                    background = Image.new('RGB', image.size, (255, 255, 255))
                    if image.mode == 'P':
                        image = image.convert('RGBA')
                    if image.mode in ('RGBA', 'LA'):
                        background.paste(image, mask=image.split()[-1] if image.mode == 'RGBA' else None)
                        image = background
                    else:
                        image = image.convert('RGB')
                elif image.mode not in ('RGB', 'L'):
                    image = image.convert('RGB')
                if image.mode == 'L':
                    image = image.convert('RGB')
                max_size = 2000
                if max(image.size) > max_size:
                    ratio = max_size / max(image.size)
                    new_size = (int(image.size[0] * ratio), int(image.size[1] * ratio))
                    image = image.resize(new_size, Image.Resampling.LANCZOS)
                img_io = io.BytesIO()
                image.save(img_io, format='PNG')
                b64 = base64.b64encode(img_io.getvalue()).decode('utf-8')
                data_url = f'data:image/png;base64,{b64}'
                last_err = None
                for attempt in range(2):
                    try:
                        response = client.chat.completions.create(
                            model=minimax_model,
                            messages=[
                                {"role": "system", "content": "你是一位经验丰富的教师，擅长识别手写作文并批阅。"},
                                {"role": "user", "content": [
                                    {"type": "image_url", "image_url": {"url": data_url}},
                                    {"type": "text", "text": user_prompt}
                                ]}
                            ],
                            temperature=0.25,
                            max_tokens=3000
                        )
                        raw = (response.choices[0].message.content or '').strip()
                        last_err = None
                        break
                    except (Exception) as e:
                        last_err = e
                        if attempt == 0 and (getattr(e, 'status_code', None) == 'timeout' or 'timeout' in str(type(e).__name__).lower() or 'timed out' in str(e).lower()):
                            continue
                        raise
                if last_err is not None:
                    raise last_err
                essay_text = ''
                report = raw
                if '【识别文本】' in raw and '【批阅报告】' in raw:
                    try:
                        _, rest = raw.split('【识别文本】', 1)
                        essay_part, report_part = rest.split('【批阅报告】', 1)
                        essay_text = essay_part.strip()
                        report = report_part.strip()
                    except Exception:
                        essay_text = raw
                else:
                    essay_text = raw
                student_id = extract_student_id(essay_text) or filename
                individual_reports[file_path] = {
                    'report': report,
                    'filename': filename,
                    'student_id': student_id,
                    'essay_text': essay_text,
                    'status': 'success'
                }
            except Exception as e:
                logger.exception(f"vision_grade 单文件失败: {file_path}")
                individual_reports[file_path] = {'error': str(e), 'status': 'error'}
        if save_to_class_center_flag:
            summary_prompt = "请将以下学生作文批阅报告进行精简汇总，提取关键信息：\n\n"
            for fp, report_data in individual_reports.items():
                if report_data.get('status') == 'success':
                    summary_prompt += f"\n【{report_data.get('student_id', fp)}】\n{report_data.get('report', '')}\n"
            summary_prompt += "\n\n请基于以上个人报告，生成一份班级整体评价报告，包括：\n1. 整体水平评估\n2. 共同优点\n3. 普遍问题\n4. 教学建议\n5. 好词好句：从学生作文或批阅点评中摘录 5-10 条佳句，单独作为一小节，每行一条，格式为「- 句子」"
            try:
                class_response = client.chat.completions.create(
                    model=llm_text_model,
                    messages=[
                        {"role": "system", "content": "你是一位经验丰富的教师，擅长分析班级整体学习情况，并能从作文中提炼好词好句。"},
                        {"role": "user", "content": summary_prompt}
                    ],
                    temperature=0.7,
                    max_tokens=6000,
                    **_text_model_kwargs()
                )
                class_evaluation = class_response.choices[0].message.content or ''
            except Exception as e:
                class_evaluation = f"生成班级评价时出错: {str(e)}"
        else:
            class_evaluation = ''
        result_data = {'individual_reports': individual_reports, 'class_evaluation': class_evaluation, 'timestamp': datetime.now().isoformat()}
        result_file = os.path.join(results_dir, f"grade_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
        with open(result_file, 'w', encoding='utf-8') as f:
            json.dump(result_data, f, ensure_ascii=False, indent=2)
        if save_to_class_center_flag:
            save_to_class_center(individual_reports, class_evaluation, logger)
        _save_task_grading_results(individual_reports)
        history_file = os.path.join(HISTORY_DIR, f"history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
        with open(history_file, 'w', encoding='utf-8') as f:
            json.dump({'type': 'grade', 'result_file': result_file, 'individual_reports': individual_reports, 'class_evaluation': class_evaluation, 'timestamp': datetime.now().isoformat()}, f, ensure_ascii=False, indent=2)
        return jsonify({
            'individual_reports': individual_reports,
            'class_evaluation': class_evaluation,
            'result_file': result_file,
            'history_file': history_file
        })
    except Exception as e:
        logger.exception("vision_grade 异常")
        return jsonify({'error': str(e)}), 500


@app.route('/api/vision_class_report', methods=['POST'])
def vision_class_report():
    """根据已有个人批阅报告生成班级整体评价（不重新识别/批阅）"""
    data = request.get_json() or {}
    individual_reports = data.get('individual_reports', {})
    save_to_class_center_flag = data.get('save_to_class_center', False)
    if not individual_reports:
        return jsonify({'error': '请先完成识别并批阅，再生成班级报告'}), 400
    if not minimax_api_key:
        return jsonify({'error': '未配置 LLM_API_KEY 或 MINIMAX_API_KEY'}), 500
    logger = logging.getLogger(__name__)
    try:
        client = OpenAI(base_url=minimax_base_url, api_key=minimax_api_key, timeout=60.0)
        summary_prompt = "请将以下学生作文批阅报告进行精简汇总，提取关键信息：\n\n"
        for fp, report_data in individual_reports.items():
            if report_data.get('status') == 'success':
                summary_prompt += f"\n【{report_data.get('student_id', fp)}】\n{report_data.get('report', '')}\n"
        summary_prompt += "\n\n请基于以上个人报告，生成一份班级整体评价报告，包括：\n1. 整体水平评估\n2. 共同优点\n3. 普遍问题\n4. 教学建议\n5. 好词好句：从学生作文或批阅点评中摘录 5-10 条佳句，单独作为一小节，每行一条，格式为「- 句子」"
        class_response = client.chat.completions.create(
            model=llm_text_model,
            messages=[
                {"role": "system", "content": "你是一位经验丰富的教师，擅长分析班级整体学习情况，并能从作文中提炼好词好句。"},
                {"role": "user", "content": summary_prompt}
            ],
            temperature=0.7,
            max_tokens=6000,
            **_text_model_kwargs()
        )
        class_evaluation = class_response.choices[0].message.content or ''
        if save_to_class_center_flag:
            save_to_class_center(individual_reports, class_evaluation, logger)
        return jsonify({'class_evaluation': class_evaluation})
    except Exception as e:
        logger.exception("vision_class_report 异常")
        return jsonify({'error': str(e)}), 500


@app.route('/api/grade', methods=['POST'])
def grade_essays():
    """批阅作文"""
    data = request.get_json()
    ocr_results = data.get('ocr_results', {})
    student_mapping = data.get('student_mapping', {})  # {file_path: student_id}
    custom_prompt_template = data.get('prompt_template', None)  # 自定义提示词模板
    save_to_class_center_flag = data.get('save_to_class_center', True)  # 是否保存到班级中心，默认True
    
    if not ocr_results:
        return jsonify({'error': '没有识别结果'}), 400
    if not minimax_api_key:
        return jsonify({'error': '未配置 LLM_API_KEY 或 MINIMAX_API_KEY，请在 .env 中设置'}), 500
    
    # 配置日志
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    
    try:
        client = OpenAI(base_url=minimax_base_url, api_key=minimax_api_key)
        
        # 为每篇作文生成个人报告
        individual_reports = {}
        for file_path, ocr_data in ocr_results.items():
            # 检查OCR识别状态
            if ocr_data.get('status') != 'success':
                error_msg = ocr_data.get('error', 'OCR识别失败')
                individual_reports[file_path] = {
                    'error': f'OCR识别失败: {error_msg}',
                    'status': 'error'
                }
                continue
            
            # 检查文本是否为空
            text = ocr_data.get('text', '').strip()
            if not text:
                individual_reports[file_path] = {
                    'error': 'OCR识别成功但文本为空，可能是图片中没有文字或图片质量不佳。请检查图片是否清晰，是否包含文字内容。',
                    'status': 'error'
                }
                continue
            
            essay_text = ocr_data['text']
            filename = os.path.basename(file_path)
            
            # 获取学生学号（优先使用手动映射，其次使用OCR识别的）
            student_id = student_mapping.get(file_path) or ocr_data.get('student_id') or filename
            
            # 构建批阅提示词（使用自定义模板或默认模板）
            prompt_template = custom_prompt_template or get_prompt_template()
            # 替换模板中的占位符
            prompt = prompt_template.replace('{essay_text}', essay_text)
            
            try:
                response = client.chat.completions.create(
                    model=minimax_model,
                    messages=[
                        {"role": "system", "content": "你是一位经验丰富的教师，擅长批阅作文和总结报告。"},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.25,
                    max_tokens=3000
                )
                
                report = response.choices[0].message.content
                individual_reports[file_path] = {
                    'report': report,
                    'filename': filename,
                    'student_id': student_id,
                    'essay_text': ocr_data.get('text', ''),
                    'status': 'success'
                }
            except Exception as e:
                individual_reports[file_path] = {
                    'error': str(e),
                    'status': 'error'
                }
        
        # 只有在需要保存到班级中心时才生成班级评价
        if save_to_class_center_flag:
            # 生成精简汇总
            summary_prompt = "请将以下学生作文批阅报告进行精简汇总，提取关键信息：\n\n"
            for file_path, report_data in individual_reports.items():
                if report_data.get('status') == 'success':
                    student_id = report_data.get('student_id', '')
                    filename = report_data.get('filename', file_path)
                    report = report_data.get('report', '')
                    # 使用学号标识，如果没有学号则使用文件名
                    identifier = f"学号：{student_id}" if student_id and student_id != filename else filename
                    summary_prompt += f"\n【{identifier}】\n{report}\n"
            
            summary_prompt += "\n\n请基于以上个人报告，生成一份班级整体评价报告，包括：\n1. 整体水平评估\n2. 共同优点\n3. 普遍问题\n4. 教学建议\n5. 好词好句：从学生作文或批阅点评中摘录 5-10 条佳句，单独作为一小节，每行一条，格式为「- 句子」"
            
            # 获取班级评价
            try:
                class_response = client.chat.completions.create(
                    model=llm_text_model,
                    messages=[
                        {"role": "system", "content": "你是一位经验丰富的教师，擅长分析班级整体学习情况。"},
                        {"role": "user", "content": summary_prompt}
                    ],
                    temperature=0.7,
                    max_tokens=6000,
                    **_text_model_kwargs()
                )
                class_evaluation = class_response.choices[0].message.content
            except Exception as e:
                class_evaluation = f"生成班级评价时出错: {str(e)}"
        else:
            # 不需要生成班级评价时，设置为空
            class_evaluation = ""
        
        # 保存批阅结果到results目录
        result_data = {
            'individual_reports': individual_reports,
            'class_evaluation': class_evaluation,
            'timestamp': datetime.now().isoformat()
        }
        
        result_file = os.path.join(results_dir, f"grade_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
        with open(result_file, 'w', encoding='utf-8') as f:
            json.dump(result_data, f, ensure_ascii=False, indent=2)
        
        # 根据参数决定是否保存到班级中心（按学生和班级组织）
        if save_to_class_center_flag:
            save_to_class_center(individual_reports, class_evaluation, logger)
        
        # 保存到历史记录（用于下次打开时恢复）
        history_file = os.path.join(HISTORY_DIR, f"history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
        history_data = {
            'type': 'grade',
            'result_file': result_file,
            'individual_reports': individual_reports,
            'class_evaluation': class_evaluation,
            'timestamp': datetime.now().isoformat()
        }
        with open(history_file, 'w', encoding='utf-8') as f:
            json.dump(history_data, f, ensure_ascii=False, indent=2)
        
        return jsonify({
            'individual_reports': individual_reports,
            'class_evaluation': class_evaluation,
            'result_file': result_file,
            'history_file': history_file
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/class_center/<class_name>', methods=['GET'])
def get_class_center(class_name):
    """获取班级中心数据"""
    try:
        class_center_file = os.path.join(CLASS_CENTER_DIR, f"{class_name}.json")
        
        if not os.path.exists(class_center_file):
            return jsonify({'students': {}})
        
        with open(class_center_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        return jsonify({'students': data})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/class_center/<class_name>/<student_id>', methods=['GET'])
def get_student_records(class_name, student_id):
    """获取指定学生的所有记录"""
    try:
        class_center_file = os.path.join(CLASS_CENTER_DIR, f"{class_name}.json")
        
        if not os.path.exists(class_center_file):
            return jsonify({'records': []})
        
        with open(class_center_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        records = data.get(student_id, [])
        # 按时间戳排序（最新的在前）
        records.sort(key=lambda x: x.get('timestamp', ''), reverse=True)
        
        return jsonify({'records': records})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/history', methods=['GET'])
def get_history():
    """获取历史记录列表"""
    try:
        history_files = []
        if os.path.exists(HISTORY_DIR):
            for filename in os.listdir(HISTORY_DIR):
                if filename.startswith('history_') and filename.endswith('.json'):
                    file_path = os.path.join(HISTORY_DIR, filename)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                        history_files.append({
                            'filename': filename,
                            'type': data.get('type', 'unknown'),
                            'timestamp': data.get('timestamp', ''),
                            'result_file': data.get('result_file', '')
                        })
                    except Exception as e:
                        continue
        
        # 按时间戳排序（最新的在前）
        history_files.sort(key=lambda x: x.get('timestamp', ''), reverse=True)
        return jsonify({'history': history_files})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/history/<filename>', methods=['GET'])
def get_history_detail(filename):
    """获取历史记录详情"""
    try:
        file_path = os.path.join(HISTORY_DIR, filename)
        if not os.path.exists(file_path):
            return jsonify({'error': '历史记录不存在'}), 404
        
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/export_txt', methods=['POST'])
def export_to_txt():
    """导出多个学号的报告为TXT文档"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': '请求数据为空'}), 400
            
        student_ids = data.get('student_ids', [])  # 学号列表（如果为空且export_all为True，则导出所有）
        class_name = data.get('class_name', '').strip()  # 班级名称
        export_all = data.get('export_all', False)  # 是否导出所有报告
        
        if not class_name:
            return jsonify({'error': '请选择班级'}), 400
        
        # 读取班级中心数据
        class_center_file = os.path.join(CLASS_CENTER_DIR, f"{class_name}.json")
        if not os.path.exists(class_center_file):
            return jsonify({'error': '班级数据不存在', 'class_name': class_name}), 404
        
        try:
            with open(class_center_file, 'r', encoding='utf-8') as f:
                class_data = json.load(f)
        except json.JSONDecodeError:
            return jsonify({'error': '班级数据文件格式错误'}), 500
        except Exception as e:
            return jsonify({'error': f'读取班级数据失败: {str(e)}'}), 500
        
        # 如果选择导出所有，则获取所有学号
        if export_all:
            student_ids = list(class_data.keys())
        
        # 收集所有学生的报告
        all_reports = []
        for student_id in student_ids:
            if student_id in class_data:
                records = class_data[student_id]
                # 按时间排序（最新的在前）
                records.sort(key=lambda x: x.get('timestamp', ''), reverse=True)
                all_reports.append({
                    'student_id': student_id,
                    'records': records
                })
        
        if not all_reports:
            error_msg = '未找到相关报告'
            if not export_all and student_ids:
                error_msg += f'（请求的学号：{", ".join(student_ids)}）'
            return jsonify({'error': error_msg, 'class_name': class_name}), 404
        
        # 生成TXT文档内容
        lines = []
        lines.append('=' * 60)
        lines.append(f'{class_name} 班级作文批阅报告汇总')
        lines.append('=' * 60)
        lines.append('')
        lines.append(f'生成时间：{datetime.now().strftime("%Y年%m月%d日 %H:%M:%S")}')
        if export_all:
            lines.append(f'包含学号：全部学号（共 {len(student_ids)} 个）')
        else:
            lines.append(f'包含学号：{", ".join(student_ids)}')
        lines.append(f'共 {len(all_reports)} 位学生，{sum(len(r["records"]) for r in all_reports)} 份报告')
        lines.append('')
        lines.append('=' * 60)
        lines.append('')
        
        # 为每个学生添加报告
        for student_report in all_reports:
            student_id = student_report['student_id']
            records = student_report['records']
            
            # 学生标题
            lines.append('-' * 60)
            lines.append(f'学号：{student_id}')
            lines.append('-' * 60)
            lines.append('')
            
            for idx, record in enumerate(records, 1):
                # 报告标题
                lines.append(f'第 {idx} 次批阅 - {record.get("date", "")}')
                lines.append('')
                
                # 作文原文
                if record.get('essay_text'):
                    lines.append('【作文原文】')
                    lines.append('-' * 40)
                    lines.append(record['essay_text'])
                    lines.append('-' * 40)
                    lines.append('')
                
                # 批阅报告
                if record.get('report'):
                    lines.append('【批阅报告】')
                    lines.append('-' * 40)
                    lines.append(record['report'])
                    lines.append('-' * 40)
                    lines.append('')
                
                # 添加分隔线（除了最后一条记录）
                if idx < len(records):
                    lines.append('')
            
            # 添加分隔线（除了最后一个学生）
            if student_report != all_reports[-1]:
                lines.append('')
                lines.append('=' * 60)
                lines.append('')
        
        # 保存TXT文档
        export_dir = EXPORTS_DIR
        os.makedirs(export_dir, exist_ok=True)
        
        export_filename = f"{class_name}_报告汇总_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        export_path = os.path.join(export_dir, export_filename)
        
        with open(export_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))
        
        return jsonify({
            'success': True,
            'filename': export_filename,
            'path': export_path,
            'download_url': f'/api/download/{export_filename}'
        })
    
    except Exception as e:
        import traceback
        return jsonify({'error': f'导出失败: {str(e)}\n{traceback.format_exc()}'}), 500

@app.route('/api/download/<filename>', methods=['GET'])
def download_file(filename):
    """下载导出的文件"""
    try:
        export_dir = EXPORTS_DIR
        file_path = os.path.join(export_dir, filename)
        
        if not os.path.exists(file_path):
            return jsonify({'error': '文件不存在'}), 404
        
        return send_from_directory(export_dir, filename, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def _require_teacher():
    """要求老师或管理员，否则返回 403"""
    role = session.get('user_role')
    if role not in ('teacher', 'admin'):
        return True  # 返回 True 表示需要拦截
    return False


def _load_materials_index():
    """加载素材索引"""
    if not os.path.exists(MATERIALS_INDEX_FILE):
        return []
    try:
        with open(MATERIALS_INDEX_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return []


def _save_materials_index(entries):
    with open(MATERIALS_INDEX_FILE, 'w', encoding='utf-8') as f:
        json.dump(entries, f, ensure_ascii=False, indent=2)


# ---------- 平台内编辑：作文素材、通用答题卡 ----------
DEFAULT_COMPOSITION_MATERIALS = """# 作文素材使用说明

本素材供作文练习使用。请将题目或阅读材料发放给学生，学生可在通用答题卡上书写作文（务必在作文上方写清6位学号），写完后扫描或拍照上传至 QuickJudge 进行批阅。

## 使用流程
1. 在平台内编辑并打印「通用答题卡」。
2. 将本次练习的题目/材料发给学生。
3. 学生在答题卡上写清学号并书写作文。
4. 扫描或拍照后，在「智能阅卷」中识别并批阅。

## 注意事项
- 学号请写6位数字，便于系统识别。
- 字迹清晰、拍照端正有利于识别准确率。
"""

DEFAULT_ANSWER_SHEET_HTML = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>通用作文答题卡</title>
<style>
* { box-sizing: border-box; }
body { font-family: "Microsoft YaHei", sans-serif; margin: 20px; max-width: 800px; margin-left: auto; margin-right: auto; }
h1 { text-align: center; font-size: 1.4rem; margin-bottom: 24px; }
.sheet { border: 2px solid #333; padding: 24px; min-height: 100vh; }
.meta { margin-bottom: 20px; display: flex; gap: 24px; flex-wrap: wrap; }
.meta label { font-weight: bold; min-width: 80px; }
.meta input { border: none; border-bottom: 1px solid #333; padding: 4px 8px; flex: 1; min-width: 120px; }
.student-id { margin-bottom: 24px; }
.student-id .label { font-weight: bold; margin-bottom: 8px; }
.student-id .hint { font-size: 12px; color: #666; margin-top: 4px; }
.student-id .boxes { display: flex; gap: 8px; }
.student-id .box { width: 36px; height: 44px; border: 2px solid #333; text-align: center; font-size: 1.2rem; line-height: 40px; position: relative; }
.student-id .box .writing-guide { position: absolute; inset: 0; display: flex; align-items: center; justify-content: center; font-size: 1.4rem; color: #ddd; pointer-events: none; font-family: "Microsoft YaHei", sans-serif; }
.writing-area { border: 1px solid #999; min-height: 400px; padding: 12px; line-height: 2; }
.writing-area .lines { border-bottom: 1px solid #eee; min-height: 32px; }
@media print { body { margin: 0; } .sheet { box-shadow: none; } }
</style>
</head>
<body>
<div class="sheet">
<h1>通用作文答题卡</h1>
<div class="meta">
  <label>班级</label><input type="text" placeholder="班级名称">
  <label>姓名</label><input type="text" placeholder="姓名">
  <label>日期</label><input type="text" placeholder="日期">
</div>
<div class="student-id">
  <div class="label">学号（六位数字，请按 8 字形规范书写，便于系统识别）</div>
  <div class="boxes">
    <div class="box"><span class="writing-guide">8</span></div><div class="box"><span class="writing-guide">8</span></div><div class="box"><span class="writing-guide">8</span></div><div class="box"><span class="writing-guide">8</span></div><div class="box"><span class="writing-guide">8</span></div><div class="box"><span class="writing-guide">8</span></div>
  </div>
  <div class="hint">每格按 8 字形书写一位数字，示例：230101</div>
</div>
<div class="writing-area">
  <div class="label" style="font-weight:bold; margin-bottom:8px;">作文书写区</div>
  <div class="lines"></div><div class="lines"></div><div class="lines"></div><div class="lines"></div><div class="lines"></div>
  <div class="lines"></div><div class="lines"></div><div class="lines"></div><div class="lines"></div><div class="lines"></div>
  <div class="lines"></div><div class="lines"></div><div class="lines"></div><div class="lines"></div><div class="lines"></div>
</div>
</div>
</body>
</html>
"""


def _read_composition_materials():
    if os.path.exists(COMPOSITION_MATERIALS_FILE):
        with open(COMPOSITION_MATERIALS_FILE, 'r', encoding='utf-8') as f:
            return f.read()
    return DEFAULT_COMPOSITION_MATERIALS


def _read_answer_sheet_html():
    if os.path.exists(ANSWER_SHEET_HTML_FILE):
        with open(ANSWER_SHEET_HTML_FILE, 'r', encoding='utf-8') as f:
            return f.read()
    return DEFAULT_ANSWER_SHEET_HTML


@app.route('/api/config/composition_materials', methods=['GET'])
def get_composition_materials():
    """获取作文素材使用说明（平台内编辑用）"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    return jsonify({'content': _read_composition_materials()})


@app.route('/api/config/composition_materials', methods=['POST', 'PUT'])
def save_composition_materials():
    """保存作文素材使用说明"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    try:
        data = request.get_json() or {}
        content = data.get('content', '')
        with open(COMPOSITION_MATERIALS_FILE, 'w', encoding='utf-8') as f:
            f.write(content)
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/config/answer_sheet', methods=['GET'])
def get_answer_sheet():
    """获取通用答题卡 HTML（平台内编辑/预览用）"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    return jsonify({'content': _read_answer_sheet_html()})


@app.route('/api/config/answer_sheet', methods=['POST', 'PUT'])
def save_answer_sheet():
    """保存通用答题卡 HTML"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    try:
        data = request.get_json() or {}
        content = data.get('content', '')
        with open(ANSWER_SHEET_HTML_FILE, 'w', encoding='utf-8') as f:
            f.write(content)
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/config/answer_sheet/preview')
def preview_answer_sheet():
    """预览/打印：返回答题卡 HTML 页面（新窗口打开用）"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    html = _read_answer_sheet_html()
    return Response(html, mimetype='text/html; charset=utf-8')


@app.route('/api/generate/composition_materials', methods=['GET'])
def generate_composition_materials():
    """生成并下载作文素材包（ZIP，使用平台内已编辑的说明）"""
    try:
        content = _read_composition_materials()
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr('作文素材使用说明.txt', content.encode('utf-8'))
            zf.writestr('阅读材料模板.txt', '请在此处填写或粘贴本次练习的阅读材料、题目要求等。\n\n'.encode('utf-8'))
        buf.seek(0)
        return send_file(buf, as_attachment=True, download_name='作文素材.zip', mimetype='application/zip')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/generate/answer_sheet', methods=['GET'])
def generate_answer_sheet():
    """下载通用答题卡（使用平台内已编辑的 HTML）"""
    html = _read_answer_sheet_html()
    buf = io.BytesIO(html.encode('utf-8'))
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name='通用答题卡.html', mimetype='text/html; charset=utf-8')


def _load_answer_sheet_templates():
    """加载答题卡题型模板（按学科预设）。"""
    try:
        if os.path.exists(ANSWER_SHEET_TEMPLATES_FILE):
            with open(ANSWER_SHEET_TEMPLATES_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return {}


@app.route('/api/config/answer_sheet_templates', methods=['GET'])
def list_answer_sheet_templates():
    """获取答题卡题型模板列表（学科预设），用于「按题型模板生成答题卡」。"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    raw = _load_answer_sheet_templates()
    templates = [{'id': sid, 'name': (data.get('name') or sid)} for sid, data in raw.items() if isinstance(data, dict)]
    return jsonify({'templates': templates})


@app.route('/api/config/answer_sheet_templates/<subject_id>', methods=['GET'])
def get_answer_sheet_template(subject_id):
    """获取指定学科的答题卡模板详情（选择题区块、主观题区块）。"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    raw = _load_answer_sheet_templates()
    data = raw.get(subject_id) if isinstance(raw, dict) else None
    if not data or not isinstance(data, dict):
        return jsonify({'error': '模板不存在'}), 404
    return jsonify({
        'id': subject_id,
        'name': data.get('name', subject_id),
        'choice_sections': data.get('choice_sections') or [],
        'subjective_sections': data.get('subjective_sections') or [],
    })


@app.route('/api/generate/answer_sheet_from_template', methods=['GET'])
def generate_answer_sheet_from_template():
    """按题型模板生成答题卡。参数：subject（学科预设 id）、title（标题）、format=html|pdf、show_answers=0|1。"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    subject_id = (request.args.get('subject') or request.args.get('preset') or '').strip() or 'english'
    title = (request.args.get('title') or '').strip() or '答题卡'
    output_format = (request.args.get('format') or 'html').lower()
    show_answers = request.args.get('show_answers', '0') not in ('0', 'false', 'no')
    raw = _load_answer_sheet_templates()
    data = raw.get(subject_id) if isinstance(raw, dict) else None
    if not data or not isinstance(data, dict):
        return jsonify({'error': '模板不存在，请选择有效学科预设'}), 404
    try:
        from utils.answer_sheet_generator import generate_answer_sheet_html, html_to_pdf
    except ImportError:
        return jsonify({'error': '答题卡生成模块不可用'}), 500
    parsed = {
        'choice_sections': data.get('choice_sections') or [],
        'choice_answers': {},
        'subjective_sections': data.get('subjective_sections') or [],
    }
    html = generate_answer_sheet_html(parsed, title=title, show_answer_keys=show_answers)
    safe_title = re.sub(r'[\\/:*?"<>|]', '-', title)
    if output_format == 'pdf':
        pdf_name = f"【A4】答题卡-{safe_title}.pdf"
        pdf_path = os.path.join(EXPORTS_DIR, pdf_name)
        os.makedirs(EXPORTS_DIR, exist_ok=True)
        if html_to_pdf(html, pdf_path):
            return send_file(pdf_path, as_attachment=True, download_name=pdf_name, mimetype='application/pdf')
        return jsonify({'error': 'PDF 生成失败，请改用 format=html 或安装 weasyprint'}), 500
    buf = io.BytesIO(html.encode('utf-8'))
    buf.seek(0)
    html_name = f"【A4】答题卡-{safe_title}.html"
    return send_file(buf, as_attachment=True, download_name=html_name, mimetype='text/html; charset=utf-8')


def _llm_parse_paper_structure(full_text: str) -> dict:
    """调用大模型从试卷正文中抽取结构，返回可覆盖/补充的 JSON。"""
    if not full_text or not minimax_api_key:
        return {}
    prompt = """你是一位试卷分析助手。请根据下面这份试卷/答案的正文，严格按 JSON 输出以下信息（不要输出其他内容）：
1. choice_sections: 选择题区块数组，每项 { "name": "听力/阅读理解/完形填空等", "start": 起始题号, "end": 结束题号 }
2. choice_answers: 选择题答案对象，题号为键，"A"/"B"/"C"/"D"为值，如 {"1":"A","2":"B"}
3. subjective_sections: 主观题区块数组，每项 { "name": "语法填空/短文改错/书面表达", "num_questions": 题数, "num_lines": 建议书写行数, "prompt": "题干或材料（书面表达必填，其他可空）" }
4. essay_prompt: 仅当有书面表达/作文时，提取题干或阅读材料全文，用于附在答题卡上

若某部分无法从正文推断，可省略或给空数组/空对象。只输出一个合法 JSON，不要 markdown 代码块包裹。"""
    try:
        client = OpenAI(base_url=minimax_base_url, api_key=minimax_api_key, timeout=60.0)
        response = client.chat.completions.create(
            model=llm_text_model,
            messages=[
                {"role": "user", "content": prompt + "\n\n" + full_text[:2500]}
            ],
            max_tokens=4000,
            **_text_model_kwargs()
        )
        content = (response.choices[0].message.content or "").strip()
        # 去掉可能的 markdown 代码块
        if content.startswith("```"):
            content = re.sub(r"^```\w*\n?", "", content)
            content = re.sub(r"\n?```\s*$", "", content)
        return json.loads(content)
    except Exception:
        return {}


@app.route('/api/generate/answer_sheet_from_paper', methods=['POST'])
def generate_answer_sheet_from_paper():
    """根据上传的 Word 试卷/答案文档生成答题卡。支持学科（如英语）选择题、主观题、作文题干。可返回 HTML 或 PDF。"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    try:
        from utils.paper_parser import parse_paper_docx, parse_paper_docx_with_llm
        from utils.answer_sheet_generator import generate_answer_sheet_html, html_to_pdf
    except ImportError as e:
        return jsonify({'error': f'缺少依赖: {e}。请安装 python-docx: pip install python-docx'}), 500

    file = request.files.get('file') or request.files.get('docx')
    if not file or not file.filename or not file.filename.lower().endswith('.docx'):
        return jsonify({'error': '请上传 .docx 格式的试卷或答案文档'}), 400

    title = (request.form.get('title') or request.form.get('sheet_title') or '').strip()
    if not title:
        # 从文件名推断，如「高三英语限时练（二）答案」->「限时练2」
        base = os.path.splitext(file.filename)[0]
        m = re.search(r'限时练\s*[（(]?(\d+)[)）]?|练\s*(\d+)|第\s*(\d+)', base)
        if m:
            title = '限时练' + (m.group(1) or m.group(2) or m.group(3) or '')
        else:
            title = base[:20] if base else '答题卡'

    use_llm = request.form.get('use_llm', 'false').lower() in ('true', '1', 'yes')
    output_format = (request.form.get('format') or 'html').lower()
    show_answers = request.form.get('show_answers', 'true').lower() in ('true', '1', 'yes')

    os.makedirs(EXPORTS_DIR, exist_ok=True)
    safe_name = str(uuid.uuid4()) + '_' + (file.filename or 'paper.docx')
    save_path = os.path.join(EXPORTS_DIR, safe_name)
    file.save(save_path)

    try:
        subject_cfg = _get_subject_config()
        sid = subject_cfg.get('current') or 'english'
        preset = ((subject_cfg.get('subjects') or {}).get(sid) or {}).get('paper_preset') or 'english'
        if use_llm and minimax_api_key:
            def llm_callback(text):
                return _llm_parse_paper_structure(text)
            parsed = parse_paper_docx_with_llm(save_path, llm_callback, preset=preset)
        else:
            parsed = parse_paper_docx(save_path, preset=preset)
        if parsed.get('error'):
            return jsonify({'error': parsed['error'], 'parsed': parsed}), 400

        html = generate_answer_sheet_html(parsed, title=title, show_answer_keys=show_answers)

        if output_format == 'pdf':
            pdf_name = f"【A4】答题卡-{title}.pdf"
            pdf_name = re.sub(r'[\\/:*?"<>|]', '-', pdf_name)
            pdf_path = os.path.join(EXPORTS_DIR, pdf_name)
            if html_to_pdf(html, pdf_path):
                return send_file(
                    pdf_path, as_attachment=True, download_name=pdf_name,
                    mimetype='application/pdf'
                )
            return jsonify({
                'error': 'PDF 生成失败。请安装 weasyprint: pip install weasyprint，或改用 format=html 下载 HTML 后用浏览器打印为 PDF'
            }), 500

        buf = io.BytesIO(html.encode('utf-8'))
        buf.seek(0)
        html_name = f"【A4】答题卡-{title}.html"
        html_name = re.sub(r'[\\/:*?"<>|]', '-', html_name)
        return send_file(
            buf, as_attachment=True, download_name=html_name,
            mimetype='text/html; charset=utf-8'
        )
    finally:
        if os.path.isfile(save_path):
            try:
                os.remove(save_path)
            except Exception:
                pass


@app.route('/api/generate/answer_sheet_from_paper/preview')
def preview_answer_sheet_from_paper():
    """预览：通过 GET 传入本地已生成的 HTML 文件名（仅用于 PDF 失败时降级）。"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    path = request.args.get('path') or request.args.get('name')
    if not path or '..' in path:
        return jsonify({'error': '无效路径'}), 400
    file_path = os.path.join(EXPORTS_DIR, path)
    if not os.path.isfile(file_path) or not path.endswith('.html'):
        return jsonify({'error': '文件不存在'}), 404
    with open(file_path, 'r', encoding='utf-8') as f:
        return Response(f.read(), mimetype='text/html; charset=utf-8')


@app.route('/api/materials', methods=['GET'])
def list_materials():
    """教师端：列出所有自定义素材（仅 teacher/admin）"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    entries = _load_materials_index()
    return jsonify({'materials': entries})


@app.route('/api/materials', methods=['POST'])
def add_material():
    """教师端：添加素材（上传文件或创建文本素材）"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    try:
        entries = _load_materials_index()
        # 上传文件
        if request.files:
            f = request.files.get('file')
            if not f or not f.filename:
                return jsonify({'error': '请选择文件'}), 400
            ext = os.path.splitext(f.filename)[1] or '.bin'
            safe_name = (re.sub(r'[^\w\s\-\.]', '', f.filename[:50]) or 'file') + ext
            mid = str(uuid.uuid4())[:8]
            filename = mid + '_' + safe_name
            filepath = os.path.join(MATERIALS_DIR, filename)
            f.save(filepath)
            entry = {'id': mid, 'name': f.filename, 'filename': filename, 'created_at': datetime.now().isoformat()}
            entries.append(entry)
            _save_materials_index(entries)
            return jsonify({'material': entry})
        # 创建文本素材（JSON body: title, content）
        data = request.get_json() or {}
        title = (data.get('title') or data.get('name') or '未命名素材').strip() or '未命名素材'
        content = data.get('content') or ''
        mid = str(uuid.uuid4())[:8]
        filename = mid + '_' + title[:30] + '.txt'
        filename = re.sub(r'[^\w\s\-\.]', '', filename) or mid + '.txt'
        filepath = os.path.join(MATERIALS_DIR, filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        entry = {'id': mid, 'name': title, 'filename': filename, 'created_at': datetime.now().isoformat()}
        entries.append(entry)
        _save_materials_index(entries)
        return jsonify({'material': entry})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/materials/<material_id>', methods=['GET'])
def download_material(material_id):
    """下载单个素材文件"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    entries = _load_materials_index()
    entry = next((e for e in entries if e.get('id') == material_id), None)
    if not entry:
        return jsonify({'error': '素材不存在'}), 404
    path = os.path.join(MATERIALS_DIR, entry['filename'])
    if not os.path.exists(path):
        return jsonify({'error': '文件不存在'}), 404
    return send_file(path, as_attachment=True, download_name=entry.get('name', entry['filename']))


@app.route('/api/materials/<material_id>', methods=['DELETE'])
def delete_material(material_id):
    """删除素材"""
    if _require_teacher():
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    entries = _load_materials_index()
    entry = next((e for e in entries if e.get('id') == material_id), None)
    if not entry:
        return jsonify({'error': '素材不存在'}), 404
    path = os.path.join(MATERIALS_DIR, entry['filename'])
    if os.path.exists(path):
        try:
            os.remove(path)
        except Exception:
            pass
    entries = [e for e in entries if e.get('id') != material_id]
    _save_materials_index(entries)
    return jsonify({'msg': '已删除'})


@app.route('/api/class_center/<class_name>/summary', methods=['GET'])
def get_class_summary(class_name):
    """获取班级中心摘要信息"""
    try:
        class_center_file = os.path.join(CLASS_CENTER_DIR, f"{class_name}.json")
        
        if not os.path.exists(class_center_file):
            return jsonify({
                'total_students': 0,
                'total_records': 0,
                'latest_date': None,
                'students': []
            })
        
        with open(class_center_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        summary = {
            'total_students': len(data),
            'total_records': sum(len(records) for records in data.values()),
            'latest_date': None,
            'students': []
        }
        
        # 计算每个学生的统计信息
        all_dates = []
        for student_id, records in data.items():
            student_summary = {
                'student_id': student_id,
                'record_count': len(records),
                'latest_date': None,
                'first_date': None
            }
            
            if records:
                dates = [r.get('date', '') for r in records if r.get('date')]
                if dates:
                    dates.sort(reverse=True)
                    student_summary['latest_date'] = dates[0]
                    student_summary['first_date'] = dates[-1]
                    all_dates.extend(dates)
            
            summary['students'].append(student_summary)
        
        # 班级最新日期
        if all_dates:
            all_dates.sort(reverse=True)
            summary['latest_date'] = all_dates[0]
        
        # 按记录数排序
        summary['students'].sort(key=lambda x: x['record_count'], reverse=True)
        
        return jsonify(summary)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/class_center/<class_name>/reports', methods=['GET'])
def get_class_reports_by_date(class_name):
    """获取班级按日期分组的报告列表"""
    try:
        class_center_file = os.path.join(CLASS_CENTER_DIR, f"{class_name}.json")
        
        if not os.path.exists(class_center_file):
            return jsonify({'reports': []})
        
        with open(class_center_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 按日期分组，每个日期对应一次批阅报告
        reports_by_date = {}  # {date: {class_evaluation, students: [student_id], timestamp}}
        
        for student_id, records in data.items():
            for record in records:
                date = record.get('date', '')
                if not date:
                    continue
                
                if date not in reports_by_date:
                    reports_by_date[date] = {
                        'date': date,
                        'class_evaluation': record.get('class_evaluation', ''),
                        'timestamp': record.get('timestamp', ''),
                        'students': set()
                    }
                
                reports_by_date[date]['students'].add(student_id)
        
        # 转换为列表并排序
        reports = []
        for date, report_data in reports_by_date.items():
            reports.append({
                'date': date,
                'class_evaluation': report_data['class_evaluation'],
                'timestamp': report_data['timestamp'],
                'student_count': len(report_data['students']),
                'students': list(report_data['students'])
            })
        
        # 按日期排序（最新的在前）
        reports.sort(key=lambda x: x.get('timestamp', x.get('date', '')), reverse=True)
        
        return jsonify({'reports': reports})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/class_center/<class_name>/reports/<date>', methods=['GET'])
def get_class_report_detail(class_name, date):
    """获取指定日期的班级报告详情（包含该日期所有学生的记录）"""
    try:
        class_center_file = os.path.join(CLASS_CENTER_DIR, f"{class_name}.json")
        
        if not os.path.exists(class_center_file):
            return jsonify({'students': {}})
        
        with open(class_center_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 获取该日期所有学生的记录
        report_students = {}
        class_evaluation = None
        
        for student_id, records in data.items():
            date_records = [r for r in records if r.get('date') == date]
            if date_records:
                report_students[student_id] = date_records
                # 获取班级评价（从第一条记录中获取）
                if not class_evaluation and date_records[0].get('class_evaluation'):
                    class_evaluation = date_records[0].get('class_evaluation')
        
        return jsonify({
            'date': date,
            'class_evaluation': class_evaluation,
            'students': report_students
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/class_center/<class_name>/grades', methods=['GET'])
def get_class_grades(class_name):
    """成绩报表：按学号返回分数列表。后续与批阅结果联动，当前返回空列表。格式 [{ student_id, score?, score_text?, note? }, ...]"""
    try:
        # 可选：从班级中心或成绩存储中汇总各学号最新分数，当前无分数字段则返回空
        grades = []
        # 后续可在此读取批阅结果，按 student_id 聚合分数后填入 grades
        return jsonify({'grades': grades})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/process_image', methods=['POST'])
def process_image():
    """处理图片区域选择，保存处理后的图片"""
    try:
        from werkzeug.utils import secure_filename
        
        if 'file' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400
        
        file = request.files['file']
        original_path = request.form.get('original_path', '')
        regions_json = request.form.get('regions', '[]')
        
        if file.filename == '':
            return jsonify({'error': '文件名为空'}), 400
        
        # 解析区域信息
        try:
            regions = json.loads(regions_json)
        except json.JSONDecodeError:
            regions = []
        
        # 读取上传的图片
        from PIL import Image
        import io
        
        image_bytes = file.read()
        processed_image = Image.open(io.BytesIO(image_bytes))
        
        # 确保是RGB模式
        if processed_image.mode != 'RGB':
            processed_image = processed_image.convert('RGB')
        
        # 保存到临时文件夹（不再保存到原文件夹）
        # 生成临时文件名（基于原始文件名）
        if original_path:
            parts = original_path.split('/')
            original_filename = parts[-1]
        else:
            original_filename = secure_filename(file.filename)
        
        name, ext = os.path.splitext(original_filename)
        new_filename = f"{name}_processed{ext}"
        save_path = os.path.join(TEMP_DIR, new_filename)
        
        # 如果临时文件夹中已存在同名文件，先删除
        if os.path.exists(save_path):
            os.remove(save_path)
        
        # 保存处理后的图片到临时文件夹
        processed_image.save(save_path, quality=95)
        
        # 返回临时文件信息（不返回路径，因为前端不需要知道具体路径）
        return_path = f"temp/{new_filename}"
        
        return jsonify({
            'success': True,
            'message': '图片处理成功',
            'path': return_path,
            'filename': new_filename,
            'regions': regions
        })
    
    except Exception as e:
        import traceback
        error_msg = str(e)
        error_trace = traceback.format_exc()
        print(f"处理图片失败: {error_msg}\n{error_trace}")
        return jsonify({'error': f'处理图片失败: {error_msg}'}), 500

# ============================================================================
# 三端分离：扫描端（Station）+ 用户端（Portal）
# ----------------------------------------------------------------------------
# 扫描端 /station ：仅本机设备可操作。负责扫描仪控制；扫描结果先进「临时区」，
#                  在临时区内完成 学号矫正 / 区域分割 / 旋转 / 归属 等整理后再提交入库。
# 用户端 /portal  ：任意设备浏览器访问。教师查看本班已扫描内容与批阅状态并生成评语，
#                  学生查看自己的卷子图片与 AI 评语。
# 原有 /         ：教师完整后台（组卷、阅卷、班级中心）保持不变。
# ============================================================================

import shutil
import socket

STAGING_DIR = os.path.join(DATA_DIR, 'staging')
STAGING_IMAGES_DIR = os.path.join(STAGING_DIR, 'images')
STAGING_INDEX_FILE = os.path.join(STAGING_DIR, 'index.json')
RECORDS_DIR = os.path.join(DATA_DIR, 'records')
RECORDS_IMAGES_DIR = os.path.join(RECORDS_DIR, 'images')
RECORDS_INDEX_FILE = os.path.join(RECORDS_DIR, 'index.json')
CLASS_OWNERS_FILE = os.path.join(CONFIG_DIR, 'class_owners.json')

for _d in (STAGING_DIR, STAGING_IMAGES_DIR, RECORDS_DIR, RECORDS_IMAGES_DIR):
    os.makedirs(_d, exist_ok=True)

# 调试开关：设为 1 时允许非本机设备操作扫描端（默认关闭，扫描端仅本机可用）
STATION_ALLOW_REMOTE = (os.getenv('STATION_ALLOW_REMOTE', '') or '').strip().lower() in ('1', 'true', 'yes', 'on')
_LOCAL_ADDRS = {'127.0.0.1', '::1', 'localhost', '::ffff:127.0.0.1'}
PORTAL_PORT = int(os.getenv('PORTAL_PORT', '5001') or '5001')
IMAGE_EXTS = ('.png', '.jpg', '.jpeg', '.webp', '.bmp')


# ---------------------------- 通用小工具 ----------------------------

def _is_local_request():
    """扫描端写操作闸门：判断请求是否来自本机。"""
    if STATION_ALLOW_REMOTE:
        return True
    addr = (request.remote_addr or '').strip()
    return addr in _LOCAL_ADDRS


def _now_str():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _read_json(path, default):
    if not os.path.isfile(path):
        return default
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception:
        return default
    if isinstance(default, list):
        return data if isinstance(data, list) else default
    if isinstance(default, dict):
        return data if isinstance(data, dict) else default
    return data


def _write_json(path, data):
    """原子写入 JSON，避免写入中途被打断导致索引损坏。"""
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def load_staging():
    return _read_json(STAGING_INDEX_FILE, [])


def save_staging(items):
    _write_json(STAGING_INDEX_FILE, items)


def load_records():
    return _read_json(RECORDS_INDEX_FILE, [])


def save_records(items):
    _write_json(RECORDS_INDEX_FILE, items)


def load_class_owners():
    return _read_json(CLASS_OWNERS_FILE, {})


def save_class_owners(data):
    _write_json(CLASS_OWNERS_FILE, data)


def _lan_portal_urls():
    """列出本机可供其它设备访问的用户端地址，供扫描端展示。"""
    ips = set()
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            s.connect(('8.8.8.8', 80))
            ips.add(s.getsockname()[0])
        finally:
            s.close()
    except Exception:
        pass
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None):
            if info[0] == socket.AF_INET:
                ips.add(info[4][0])
    except Exception:
        pass
    urls = []
    for ip in sorted(ips):
        if ip.startswith('127.') or ip.startswith('169.254.'):
            continue
        urls.append('http://%s:%d/portal' % (ip, PORTAL_PORT))
    return urls


def _load_classes_roster():
    """返回 [{name, students:[学号], student_names:{学号:姓名}}]，供扫描端选择归属。"""
    classes = []
    try:
        if os.path.isdir(CLASSES_DIR):
            names_map = get_student_names_map()
            for item in sorted(os.listdir(CLASSES_DIR)):
                p = os.path.join(CLASSES_DIR, item)
                if not os.path.isdir(p) or item.startswith('.'):
                    continue
                raw = _read_json(os.path.join(STUDENTS_DIR, item + '.json'), [])
                ids = [str(s) for s in raw] if isinstance(raw, list) else (
                    [str(k) for k in raw.keys()] if isinstance(raw, dict) else [])
                classes.append({
                    'name': item,
                    'students': ids,
                    'student_names': {i: names_map.get(i, '') for i in ids},
                })
    except Exception:
        pass
    return classes


def _open_image(path):
    """打开图片并统一转为 RGB，兼容透明通道与调色板模式。"""
    img = Image.open(path)
    if img.mode == 'P':
        img = img.convert('RGBA')
    if img.mode in ('RGBA', 'LA'):
        bg = Image.new('RGB', img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[-1])
        img = bg
    elif img.mode != 'RGB':
        img = img.convert('RGB')
    return img


def _transform_image(img, rotation=0, crop=None):
    """按旋转角度与归一化裁剪框处理图片。crop 为 {x,y,w,h}（0~1 相对坐标）。"""
    try:
        rot = int(rotation or 0) % 360
    except Exception:
        rot = 0
    if rot in (90, 180, 270):
        img = img.rotate(-rot, expand=True)
    if isinstance(crop, dict):
        try:
            w, h = img.size
            x = max(0.0, min(1.0, float(crop.get('x', 0))))
            y = max(0.0, min(1.0, float(crop.get('y', 0))))
            cw = max(0.01, min(1.0 - x, float(crop.get('w', 1))))
            ch = max(0.01, min(1.0 - y, float(crop.get('h', 1))))
            box = (int(x * w), int(y * h), int((x + cw) * w), int((y + ch) * h))
            if box[2] - box[0] > 8 and box[3] - box[1] > 8:
                img = img.crop(box)
        except Exception:
            pass
    return img


def _norm_region(region, index=None):
    """校验并归一化一个区域框，返回 {id,label,x,y,w,h} 或 None。

    区域框用 0~1 的相对坐标，与图片分辨率无关，换 dpi 后依然有效。
    """
    if not isinstance(region, dict):
        return None
    try:
        x = max(0.0, min(1.0, float(region.get('x', 0))))
        y = max(0.0, min(1.0, float(region.get('y', 0))))
        w = max(0.0, min(1.0, float(region.get('w', 0))))
        h = max(0.0, min(1.0, float(region.get('h', 0))))
    except (TypeError, ValueError):
        return None
    w = min(w, 1.0 - x)
    h = min(h, 1.0 - y)
    if w <= 0.005 or h <= 0.005:
        return None
    rid = str(region.get('id') or '').strip()
    if not rid:
        rid = 'r%d' % ((index + 1) if index is not None else 1)
    label = str(region.get('label') or '').strip()
    if not label:
        label = '区域%s' % ((index + 1) if index is not None else 1)
    return {'id': rid, 'label': label[:40],
            'x': round(x, 4), 'y': round(y, 4), 'w': round(w, 4), 'h': round(h, 4)}


def _clean_regions(regions):
    """批量清洗区域数组：丢弃无效框，按有效顺序补齐 id/label，保持用户顺序。

    幂等：已带 id/label 的区域再次清洗时保持原值，保证档案里存的与批阅时读的一致。
    """
    out = []
    if not isinstance(regions, list):
        return out
    for r in regions:
        nr = _norm_region(r, len(out))
        if nr:
            out.append(nr)
    return out[:24]


def _item_regions(item):
    """取条目的有效区域：优先 regions，其次兼容旧的单个 crop；都没有则返回 []（整页）。"""
    if not isinstance(item, dict):
        return []
    regions = _clean_regions(item.get('regions'))
    if regions:
        return regions
    crop = _norm_region(item.get('crop'), 0)
    if crop:
        return [dict(crop, id='r1', label='选区')]
    return []


def _find_region(regions, region_id):
    for r in regions or []:
        if str(r.get('id')) == str(region_id):
            return r
    return None


def _staging_image_path(item):
    """解析临时区条目对应的真实图片路径，并做目录逃逸校验。"""
    rel = (item or {}).get('image') or ''
    p = os.path.join(STAGING_DIR, rel)
    try:
        if os.path.isfile(p) and os.path.realpath(p).startswith(os.path.realpath(STAGING_IMAGES_DIR)):
            return p
    except Exception:
        pass
    return None


def _records_image_path(rec):
    """解析档案记录对应的真实图片路径，并做目录逃逸校验。"""
    rel = (rec or {}).get('image') or ''
    p = os.path.join(RECORDS_DIR, rel)
    try:
        if os.path.isfile(p) and os.path.realpath(p).startswith(os.path.realpath(RECORDS_IMAGES_DIR)):
            return p
    except Exception:
        pass
    return None


def _find_by_id(items, item_id):
    for it in items:
        if it.get('id') == item_id:
            return it
    return None


def _run_scanner(scan_output_dir, next_filename):
    """调用 NAPS2 扫描，返回 (ok, message, [页面路径...])。

    与后台 /api/scan 完全对齐的三件事（2026-10-19 修正）：
      1) 显式传 --source/--pagesize/--dpi/--bitdepth：NAPS2 默认档位是 100dpi+Letter，
         扫试卷时手写字会糊，直接影响后续视觉模型识别与评分准确率；
      2) 加 --noprofile：避免图形界面里残留的「等待输入/自动保存」档案让命令行扫描卡住；
      3) 复用模块级 _SCAN_LOCK：同一台扫描仪同时只能有一次扫描会话，否则 TWAIN 会报
         MaxConnections 并把设备拖入「无响应」状态（只能拔插 USB 恢复）。
    """
    import subprocess
    naps2_path = get_naps2_path()
    if not os.path.exists(naps2_path):
        return False, ('未找到 NAPS2 扫描程序，无法扫描。当前配置路径：' + str(naps2_path) +
                       '。请在后台「文件管理」的扫描设置中配置正确路径。'), []
    if not _SCAN_LOCK.acquire(blocking=False):
        return False, ('已有一次扫描正在进行中，请等本次扫描结束后再点击。\n\n'
                       '（同一台扫描仪同时只能扫描一次，连续快速点击会让设备被占用而卡住。）'), []
    try:
        os.makedirs(scan_output_dir, exist_ok=True)
        scan_output_path = os.path.join(scan_output_dir, next_filename)
        driver, device = get_naps2_scan_target()
        scan_opts = get_naps2_scan_options()
        cmd = [naps2_path, '-o', scan_output_path, '-f', '-v', '--noprofile']
        if driver:
            cmd += ['--driver', driver]
        if device:
            cmd += ['--device', device]
        for key, flag in (('source', '--source'), ('pagesize', '--pagesize'),
                          ('dpi', '--dpi'), ('bitdepth', '--bitdepth')):
            val = scan_opts.get(key)
            if val not in (None, ''):
                cmd += [flag, str(val)]
        print('执行扫描命令: %s' % ' '.join(cmd), flush=True)

        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        try:
            raw_out, raw_err = proc.communicate(timeout=180)
        except subprocess.TimeoutExpired:
            # 超时必须结束子进程，否则 NAPS2 会一直占用扫描仪，导致后续扫描持续失败
            try:
                proc.kill()
                proc.communicate(timeout=10)
            except Exception:
                pass
            return False, '扫描超时，请检查扫描仪连接或重试。', []
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass
            raise
        return _finish_scan(proc, scan_output_dir, next_filename, raw_out, raw_err)
    finally:
        try:
            _SCAN_LOCK.release()
        except RuntimeError:
            pass


def _finish_scan(proc, scan_output_dir, next_filename, raw_out, raw_err):
    """扫描进程结束后统一收尾：解码输出、收集页面、生成中文排查提示。"""

    combined = (_decode_scan_output(raw_out) + '\n' + _decode_scan_output(raw_err)).strip()
    print('扫描命令返回码: %s' % proc.returncode)
    if combined:
        print('扫描命令输出: %s' % combined)

    pages = _collect_scan_pages(scan_output_dir, next_filename)
    if proc.returncode != 0 or not pages:
        info = _humanize_scan_error(combined, proc.returncode)
        hints = '\n'.join('%d. %s' % (i + 1, h) for i, h in enumerate(info['hints']))
        driver, _device = get_naps2_scan_target()
        drv_used, devices = _list_scan_devices(get_naps2_path(), driver)
        if devices:
            dev_text = ('\n\n系统可识别的扫描设备（%s 驱动）：\n' % drv_used) + '\n'.join('· ' + d for d in devices)
        else:
            dev_text = '\n\n未能列出任何扫描设备，请检查扫描仪电源与 USB 连接。'
        detail = combined or '（扫描仪无任何输出，可能未响应或已休眠）'
        return False, '%s\n\n请按以下步骤排查：\n%s%s\n\n技术详情：%s' % (info['title'], hints, dev_text, detail), []
    return True, '', pages


def _new_staging_item(image_path, source='scan', original_name=''):
    """把一张图片登记进临时区，返回新条目（图片会被移动到临时区目录）。"""
    item_id = uuid.uuid4().hex
    ext = os.path.splitext(image_path)[1].lower()
    if ext not in IMAGE_EXTS:
        ext = '.png'
    dest_name = item_id + ext
    dest_path = os.path.join(STAGING_IMAGES_DIR, dest_name)
    if os.path.abspath(image_path) != os.path.abspath(dest_path):
        shutil.move(image_path, dest_path)
    item = {
        'id': item_id,
        'image': 'images/' + dest_name,
        'original_name': original_name or os.path.basename(image_path),
        'source': source,
        'created_at': _now_str(),
        'student_id': '',
        'student_id_guess': '',
        'class_name': '',
        'task_id': '',
        'subject': '',
        'crop': None,
        'regions': [],
        'rotation': 0,
        'ocr_text': '',
        'note': '',
        'status': 'staged',
    }
    # 入临时区就做一次本地卷面体检（毫秒级、不花钱），
    # 老师扫完立刻能看到「这张字太小 / 这张偏糊」，不必等批阅结果。
    q = image_quality_metrics(dest_path)
    if q:
        item['quality'] = q
    return item


def _student_id_candidates(guess, class_name=''):
    """依据 OCR 猜测给出学号候选：精确命中优先，其次同班同前缀的近似学号。

    4 位与 6 位学号都支持；同班判定按「学号前缀长度」（4 位取前 2 位、6 位取前 4 位）。
    """
    if not guess:
        return []
    g = str(guess).strip()
    prefix_len = 4 if len(g) == 6 else (2 if len(g) == 4 else 0)
    cands = []
    for c in _load_classes_roster():
        if class_name and c['name'] != class_name:
            continue
        for sid in c['students']:
            name = (c.get('student_names') or {}).get(sid, '')
            if sid == g:
                cands.append({'student_id': sid, 'class_name': c['name'], 'name': name, 'score': 100})
            elif prefix_len and len(sid) == len(g) and sid[:prefix_len] == g[:prefix_len]:
                diff = sum(1 for a, b in zip(sid, g) if a != b)
                cands.append({'student_id': sid, 'class_name': c['name'], 'name': name,
                              'score': 90 - diff, 'diff_positions': diff})
    cands.sort(key=lambda x: (-x['score'], x['student_id']))
    return cands[:8]


def _commit_staging_item(item, student_id, class_name, task_id, note, crop, rotation,
                         regions=None, subject='', extra=None):
    """把临时区条目正式入库到档案库。返回 (ok, record_or_error)。

    regions 为区域数组（一页多块，如小练习里的某几道题）；为空表示整页当作一份。
    只要 regions 非空，crop 就不再参与入库（区域优先）。
    extra 为附加字段（batch_id / essay / topic / auto 等），合并进记录，供门户与流水线使用。
    """
    src = _staging_image_path(item)
    if not src:
        return False, '图片文件缺失，无法提交'
    student_id = (student_id or '').strip()
    class_name = (class_name or '').strip()
    if student_id and len(re.sub(r'[^\d]', '', student_id)) == 4:
        # 高一学生卷面/手工录入可能只有后四位，能唯一回查就换成完整学号
        full_sid, sid_note = resolve_student_id(student_id, '')
        if not sid_note:
            student_id = full_sid
    if not class_name and student_id:
        class_name = student_id_to_class_code(student_id) or ''
    if not class_name:
        return False, '请先指定班级（或填写学号，系统会按学号自动归档：6 位学号前 4 位是班级；高一可只填后四位，系统按名单回查）'

    clean_regions = _clean_regions(regions if regions is not None else item.get('regions'))
    # 区域优先：有区域时整页原样保存，裁切在各区域取图时按需进行，便于日后调整
    effective_crop = None if clean_regions else crop

    rec_id = uuid.uuid4().hex
    ext = os.path.splitext(src)[1].lower() or '.png'
    dest_name = rec_id + ext
    dest_path = os.path.join(RECORDS_IMAGES_DIR, dest_name)
    try:
        img = _open_image(src)
        _transform_image(img, rotation=rotation, crop=effective_crop).save(dest_path)
    except Exception as e:
        return False, '图片保存失败：%s' % e
    try:
        os.remove(src)
    except OSError:
        pass

    tasks = load_tasks()
    task = next((t for t in tasks if t.get('id') == task_id), None)
    names = get_student_names_map()
    # 后台线程里没有请求上下文，session 不可用，必须容错
    try:
        role, uid = session.get('user_role'), session.get('user_id')
    except Exception:
        role, uid = None, None
    owner = ''
    if role in ('teacher', 'admin'):
        owner = str(uid or '')
    subj_key = str(subject or '').strip()
    subj = _get_grading_subject(subj_key)
    rec = {
        'id': rec_id,
        'image': 'images/' + dest_name,
        'class_name': class_name,
        'student_id': student_id,
        'student_name': names.get(student_id, '') if student_id else '',
        'task_id': task_id or '',
        'task_title': (task or {}).get('title', ''),
        'subject': subj_key,
        'subject_name': subj.get('name', ''),
        'max_score': subj.get('max_score'),
        'crop': effective_crop,
        'regions': clean_regions,
        'rotation': int(rotation or 0) % 360,
        'note': note or '',
        'source': item.get('source', 'scan'),
        'created_at': item.get('created_at') or _now_str(),
        'submitted_at': _now_str(),
        'owner': owner,
        'status': 'pending',
        'ai_comment': '',
        'score_text': '',
        'essay_text': '',
        'grading': None,
        'graded_at': '',
    }
    if isinstance(extra, dict) and extra:
        rec.update({k: v for k, v in extra.items() if k not in ('id', 'image')})
    with _RECORDS_LOCK:
        records = load_records()
        records.insert(0, rec)
        save_records(records)

    if owner:
        owners = load_class_owners()
        lst = owners.get(class_name) or []
        if owner not in lst:
            lst.append(owner)
        owners[class_name] = lst
        save_class_owners(owners)

    # 若归属到某个阅卷任务，同步一份到该任务的试卷目录，兼容原有阅卷流程
    if task_id:
        try:
            task_dir = os.path.join(TASK_PAPERS_DIR, task_id)
            os.makedirs(task_dir, exist_ok=True)
            shutil.copyfile(dest_path, os.path.join(task_dir, dest_name))
        except Exception:
            pass
    return True, rec


# ---------------------------- 扫描端页面与接口 ----------------------------

@app.route('/station')
def station_page():
    """扫描端工作台页面（设计上仅本机使用）。"""
    return render_template('station.html')


@app.route('/api/station/context')
def station_context():
    """扫描端初始化信息：是否本机、扫描配置、班级名单、任务列表、局域网门户地址。"""
    is_local = _is_local_request()
    user = None
    if session.get('user_id'):
        user = {'id': session.get('user_id'), 'name': session.get('user_name', ''),
                'role': session.get('user_role', '')}
    tasks = [{'id': t.get('id'), 'title': t.get('title'), 'class_names': t.get('class_names') or []}
             for t in load_tasks() if t.get('id')]
    return jsonify({
        'is_local': is_local,
        'remote_addr': request.remote_addr,
        'scan_output_dir': get_scan_output_dir(),
        'naps2_path': get_naps2_path(),
        'classes': _load_classes_roster(),
        'tasks': tasks,
        'grading': grading_subject_options(),
        'user': user,
        'portal_urls': _lan_portal_urls(),
        'staging_count': len(load_staging()),
    })


@app.route('/api/station/scan', methods=['POST'])
def station_scan():
    """点击扫描：驱动扫描仪出图（支持一次多页），随后立即自动起批。

    默认 auto=1：扫完直接进入「识别表头 → 自动分班 → 批阅 → 入库」，老师不需要再做任何操作；
    传 auto=0 则退化为旧流程（只进临时区，由老师手动整理后提交）。
    """
    if not _is_local_request():
        return jsonify({'error': '扫描端仅限本机设备操作（当前为远程访问）'}), 403
    req = request.get_json(silent=True) or {}
    auto = req.get('auto', True) not in (False, 0, '0', 'false', 'False')
    subject = str(req.get('subject') or 'english_essay').strip() or 'english_essay'
    task_id = str(req.get('task_id') or '').strip()
    try:
        scan_output_dir = get_scan_output_dir()
        next_filename = get_next_scan_filename(scan_output_dir)
        ok, msg, pages = _run_scanner(scan_output_dir, next_filename)
        if not ok:
            return jsonify({'error': msg}), 500
        with _STAGING_LOCK:
            items = load_staging()
            added = []
            for p in pages:
                item = _new_staging_item(p, source='scan', original_name=os.path.basename(p))
                items.insert(0, item)
                added.append(item)
            save_staging(items)
        resp = {'ok': True, 'items': added, 'added': added, 'count': len(added),
                'filename': os.path.basename(pages[0]) if pages else '', 'auto': auto}
        if auto and added:
            b = _start_auto_batch([it['id'] for it in added], subject=subject, task_id=task_id)
            resp['batch'] = {'id': b['id'], 'job': b.get('job')}
        return jsonify(resp)
    except Exception as e:
        import traceback
        print('扫描失败: %s\n%s' % (e, traceback.format_exc()))
        return jsonify({'error': '扫描失败：%s' % e}), 500


@app.route('/api/station/jobs/<batch_id>')
def station_job(batch_id):
    """查询自动批阅进度（扫描端轮询）。"""
    if not _is_local_request():
        return jsonify({'error': '扫描端仅限本机设备操作（当前为远程访问）'}), 403
    b = _load_batch(batch_id)
    if not b:
        return jsonify({'error': '批次不存在'}), 404
    return jsonify({
        'id': b['id'],
        'job': b.get('job') or {},
        'topic': b.get('topic') or {},
        'results': b.get('results') or [],
        'record_ids': b.get('record_ids') or [],
        'failed': b.get('failed') or [],
        'corpus_ready': bool((b.get('corpus') or {}).get('points')
                             or (b.get('corpus') or {}).get('phrases')),
        'awards': (b.get('awards') or {}).get('awards') or [],
    })


@app.route('/api/station/jobs/<batch_id>/retry', methods=['POST'])
def station_job_retry(batch_id):
    """重跑某个批次（例如模型临时超时）。"""
    if not _is_local_request():
        return jsonify({'error': '扫描端仅限本机设备操作（当前为远程访问）'}), 403
    b = _load_batch(batch_id)
    if not b:
        return jsonify({'error': '批次不存在'}), 404
    t = threading.Thread(target=_auto_pipeline, args=(batch_id,), daemon=True)
    t.start()
    return jsonify({'ok': True, 'id': batch_id})


@app.route('/api/station/import', methods=['POST'])
def station_import():
    """本机导入图片进临时区（无扫描仪或补扫时使用）。

    传 auto=1（表单字段或 ?auto=1）时，导入后直接起自动批阅，与扫描同一条流水线。
    """
    if not _is_local_request():
        return jsonify({'error': '扫描端仅限本机设备操作（当前为远程访问）'}), 403
    files = request.files.getlist('files') or ([request.files['file']] if 'file' in request.files else [])
    if not files:
        return jsonify({'error': '请选择要导入的图片'}), 400
    auto = str(request.form.get('auto') or request.args.get('auto') or '').lower() in ('1', 'true', 'yes')
    subject = str(request.form.get('subject') or request.args.get('subject') or 'english_essay').strip()
    task_id = str(request.form.get('task_id') or request.args.get('task_id') or '').strip()
    items = load_staging()
    added = []
    tmp_dir = os.path.join(STAGING_DIR, '_incoming')
    os.makedirs(tmp_dir, exist_ok=True)
    try:
        with _STAGING_LOCK:
            for f in files:
                if not f or not f.filename:
                    continue
                ext = os.path.splitext(f.filename)[1].lower()
                if ext not in IMAGE_EXTS:
                    ext = '.png'
                tmp_path = os.path.join(tmp_dir, uuid.uuid4().hex + ext)
                f.save(tmp_path)
                item = _new_staging_item(tmp_path, source='upload', original_name=f.filename)
                items.insert(0, item)
                added.append(item)
            save_staging(items)
    except Exception as e:
        return jsonify({'error': '导入失败：%s' % e}), 500
    finally:
        try:
            if os.path.isdir(tmp_dir) and not os.listdir(tmp_dir):
                os.rmdir(tmp_dir)
        except OSError:
            pass
    resp = {'ok': True, 'added': added, 'count': len(added), 'auto': auto}
    if auto and added:
        b = _start_auto_batch([it['id'] for it in added], subject=subject, task_id=task_id)
        resp['batch'] = {'id': b['id'], 'job': b.get('job')}
    return jsonify(resp)


@app.route('/api/station/staging', methods=['GET'])
def station_staging_list():
    """临时区列表（仅本机可读，避免学生/远程端看到未整理的原图）。"""
    if not _is_local_request():
        return jsonify({'error': '扫描端仅限本机设备操作（当前为远程访问）'}), 403
    items = sorted(load_staging(), key=lambda x: x.get('created_at', ''), reverse=True)
    # 只返回仍能读到图片的条目，避免死链
    alive = []
    for it in items:
        it = dict(it)
        it['has_image'] = bool(_staging_image_path(it))
        # 老条目没有卷面体检数据：读列表时补算一次（本地算法，毫秒级），保证界面一致
        if not it.get('quality') and it['has_image']:
            q = image_quality_metrics(_staging_image_path(it))
            if q:
                it['quality'] = q
        alive.append(it)
    return jsonify({'items': alive, 'count': len(alive)})


@app.route('/api/station/staging/<item_id>/image')
def station_staging_image(item_id):
    """临时区图片。

    默认原图；?rotated=1 仅旋转；?processed=1 旋转+（单 crop 时）裁剪；
    ?region=<区域id> 输出该区域的裁剪图（一页多区域时使用）。
    """
    if not _is_local_request():
        return jsonify({'error': '扫描端仅限本机设备操作（当前为远程访问）'}), 403
    item = _find_by_id(load_staging(), item_id)
    if not item:
        return jsonify({'error': '临时区条目不存在'}), 404
    path = _staging_image_path(item)
    if not path:
        return jsonify({'error': '图片文件缺失'}), 404
    region_id = (request.args.get('region') or '').strip()
    if region_id:
        region = _find_region(_clean_regions(item.get('regions')), region_id)
        if not region:
            return jsonify({'error': '区域不存在'}), 404
        try:
            img = _transform_image(_open_image(path), item.get('rotation', 0), region)
            buf = io.BytesIO()
            img.save(buf, format='PNG')
            return Response(buf.getvalue(), mimetype='image/png')
        except Exception as e:
            return jsonify({'error': '区域取图失败：%s' % e}), 500
    if request.args.get('processed'):
        try:
            # 有区域时整页不再裁切（裁切交给 ?region=），但仍应用旋转
            crop = None if _clean_regions(item.get('regions')) else item.get('crop')
            img = _transform_image(_open_image(path), item.get('rotation', 0), crop)
            buf = io.BytesIO()
            img.save(buf, format='PNG')
            return Response(buf.getvalue(), mimetype='image/png')
        except Exception:
            pass
    if request.args.get('rotated'):
        try:
            img = _transform_image(_open_image(path), item.get('rotation', 0), None)
            buf = io.BytesIO()
            img.save(buf, format='PNG')
            return Response(buf.getvalue(), mimetype='image/png')
        except Exception:
            pass
    return send_from_directory(STAGING_IMAGES_DIR, os.path.basename(path))


@app.route('/api/station/staging/<item_id>', methods=['POST'])
def station_staging_update(item_id):
    """更新临时区条目：学号、班级、任务、裁剪区域、旋转、备注。"""
    if not _is_local_request():
        return jsonify({'error': '扫描端仅限本机设备操作（当前为远程访问）'}), 403
    data = request.get_json(silent=True) or {}
    items = load_staging()
    item = _find_by_id(items, item_id)
    if not item:
        return jsonify({'error': '临时区条目不存在'}), 404
    if 'student_id' in data:
        sid = str(data.get('student_id') or '').strip()
        item['student_id'] = sid
        if sid and not (data.get('class_name') or item.get('class_name')):
            item['class_name'] = student_id_to_class_code(sid) or item.get('class_name', '')
    if 'class_name' in data:
        item['class_name'] = str(data.get('class_name') or '').strip()
    if 'task_id' in data:
        item['task_id'] = str(data.get('task_id') or '').strip()
    if 'note' in data:
        item['note'] = str(data.get('note') or '')
    if 'rotation' in data:
        try:
            item['rotation'] = int(data.get('rotation') or 0) % 360
        except Exception:
            pass
    if 'crop' in data:
        crop = data.get('crop')
        if isinstance(crop, dict) and all(k in crop for k in ('x', 'y', 'w', 'h')):
            try:
                item['crop'] = {k: max(0.0, min(1.0, float(crop[k]))) for k in ('x', 'y', 'w', 'h')}
            except Exception:
                item['crop'] = None
        else:
            item['crop'] = None
    if 'regions' in data:
        item['regions'] = _clean_regions(data.get('regions'))
        # 一旦有了多区域，旧的单 crop 自动让位，避免两套坐标互相干扰
        if item['regions']:
            item['crop'] = None
    if 'subject' in data:
        item['subject'] = str(data.get('subject') or '').strip()
    save_staging(items)
    return jsonify({'ok': True, 'item': item})


@app.route('/api/station/staging/<item_id>/rotate', methods=['POST'])
def station_staging_rotate(item_id):
    """顺时针旋转 90 度（在临时区内累积）。"""
    if not _is_local_request():
        return jsonify({'error': '扫描端仅限本机设备操作（当前为远程访问）'}), 403
    data = request.get_json(silent=True) or {}
    step = int(data.get('step', 90))
    items = load_staging()
    item = _find_by_id(items, item_id)
    if not item:
        return jsonify({'error': '临时区条目不存在'}), 404
    item['rotation'] = (int(item.get('rotation', 0)) + step) % 360
    save_staging(items)
    return jsonify({'ok': True, 'item': item})


@app.route('/api/station/staging/<item_id>/detect_id', methods=['POST'])
def station_detect_id(item_id):
    """识别学号：调用视觉模型读图，抽取 6 位学号并给出候选（含名单校正）。"""
    if not _is_local_request():
        return jsonify({'error': '扫描端仅限本机设备操作（当前为远程访问）'}), 403
    data = request.get_json(silent=True) or {}
    items = load_staging()
    item = _find_by_id(items, item_id)
    if not item:
        return jsonify({'error': '临时区条目不存在'}), 404
    path = _staging_image_path(item)
    if not path:
        return jsonify({'error': '图片文件缺失'}), 404
    try:
        img = _open_image(path)
        img = _transform_image(img, item.get('rotation', 0), None)
        text = llm_ocr_recognize(img) or ''
    except Exception as e:
        return jsonify({'error': '识别失败：%s' % e}), 500
    guess_seen = extract_student_id(text) or ''
    # 高一卷面只有后四位，这里先尝试回查成完整学号，再据完整学号推班级
    guess, guess_note = resolve_student_id(guess_seen, item.get('class_name') or '')
    if guess_note:
        guess = guess_seen  # 解析不确定就保留原样，交给老师确认
    item['ocr_text'] = text[:4000]
    item['student_id_guess'] = guess
    item['student_id_seen'] = guess_seen
    item['sid_note'] = guess_note
    if guess and not guess_note:
        item['student_id'] = guess
        if not item.get('class_name'):
            item['class_name'] = student_id_to_class_code(guess) or ''
    save_staging(items)
    if not text:
        return jsonify({'ok': False, 'error': '未能识别出文字，请检查图片或手动填写学号', 'item': item,
                        'candidates': []})
    return jsonify({'ok': True, 'student_id': guess, 'student_id_seen': guess_seen,
                    'sid_note': guess_note, 'ocr_text': item['ocr_text'],
                    'item': item, 'candidates': _student_id_candidates(guess, item.get('class_name', ''))})


@app.route('/api/station/handwriting', methods=['POST'])
def station_handwriting():
    """字迹体检（与改卷无关）：评估一份手写卷「机器读起来顺不顺」。

    入参任选：{"item_id": 临时区条目} / {"record_id": 已归档记录}。
    返回：本地量化指标 + 模型逐行转写（读不准的词标出）+ 可读性等级 A~D。
    """
    if not _is_local_request():
        return jsonify({'error': '扫描端仅限本机设备操作（当前为远程访问）'}), 403
    data = request.get_json(silent=True) or {}
    path, ref = None, None

    iid = str(data.get('item_id') or '').strip()
    rid = str(data.get('record_id') or '').strip()
    if iid:
        item = _find_by_id(load_staging(), iid)
        if not item:
            return jsonify({'error': '临时区条目不存在'}), 404
        path = _staging_image_path(item)
        ref = item
    elif rid:
        rec = next((r for r in load_records() if r.get('id') == rid), None)
        if not rec:
            return jsonify({'error': '记录不存在'}), 404
        path = _records_image_path(rec)
        ref = rec
    if not path:
        return jsonify({'error': '图片文件缺失'}), 404

    out, err = assess_handwriting(path, model=data.get('model') or None)
    if err and not out.get('local'):
        return jsonify({'ok': False, 'error': err}), 500
    # 顺便把结果记进条目/记录，避免重复开销
    if isinstance(ref, dict):
        try:
            if iid:
                items = load_staging()
                for it in items:
                    if it.get('id') == iid:
                        it['handwriting'] = {k: out.get(k) for k in
                                             ('score', 'grade', 'grade_label', 'unsure_ratio', 'tips', 'local')}
                save_staging(items)
            else:
                brief = {k: out.get(k) for k in
                         ('score', 'grade', 'grade_label', 'unsure_ratio', 'tips', 'local')}
                recs = load_records()
                hit = False
                for r in recs:
                    if r.get('id') == rid:
                        r['handwriting'] = brief
                        hit = True
                if hit:
                    save_records(recs)
        except Exception:
            pass
    return jsonify(dict(out, ok=True, error=err or ''))


@app.route('/api/station/staging/<item_id>', methods=['DELETE'])
def station_staging_delete(item_id):
    """丢弃临时区条目（连同图片一并删除）。"""
    if not _is_local_request():
        return jsonify({'error': '扫描端仅限本机设备操作（当前为远程访问）'}), 403
    items = load_staging()
    item = _find_by_id(items, item_id)
    if not item:
        return jsonify({'error': '临时区条目不存在'}), 404
    path = _staging_image_path(item)
    if path:
        try:
            os.remove(path)
        except OSError:
            pass
    save_staging([it for it in items if it.get('id') != item_id])
    return jsonify({'ok': True, 'id': item_id})


@app.route('/api/station/staging/<item_id>/submit', methods=['POST'])
def station_submit(item_id):
    """把一张整理好的试卷从临时区提交入库（此后用户端可见）。"""
    if not _is_local_request():
        return jsonify({'error': '扫描端仅限本机设备操作（当前为远程访问）'}), 403
    data = request.get_json(silent=True) or {}
    items = load_staging()
    item = _find_by_id(items, item_id)
    if not item:
        return jsonify({'error': '临时区条目不存在'}), 404
    ok, payload = _commit_staging_item(
        item,
        data.get('student_id', item.get('student_id', '')),
        data.get('class_name', item.get('class_name', '')),
        data.get('task_id', item.get('task_id', '')),
        data.get('note', item.get('note', '')),
        data.get('crop', item.get('crop')),
        data.get('rotation', item.get('rotation', 0)),
        data.get('regions', item.get('regions')),
        data.get('subject', item.get('subject', '')),
    )
    if not ok:
        return jsonify({'error': payload}), 400
    save_staging([it for it in items if it.get('id') != item_id])
    return jsonify({'ok': True, 'record': payload})


@app.route('/api/station/staging/submit_all', methods=['POST'])
def station_submit_all():
    """批量提交：仅提交已填学号（或已选班级）的条目，其余留在临时区。"""
    if not _is_local_request():
        return jsonify({'error': '扫描端仅限本机设备操作（当前为远程访问）'}), 403
    data = request.get_json(silent=True) or {}
    default_task = str(data.get('task_id') or '').strip()
    default_class = str(data.get('class_name') or '').strip()
    items = load_staging()
    submitted, skipped, remain = [], [], []
    for item in items:
        sid = str(item.get('student_id') or '').strip()
        cls = str(item.get('class_name') or '').strip() or default_class or (student_id_to_class_code(sid) or '')
        if not cls and not sid:
            skipped.append({'id': item.get('id'), 'reason': '未填写学号或班级'})
            remain.append(item)
            continue
        ok, payload = _commit_staging_item(
            item, sid, cls, str(item.get('task_id') or '').strip() or default_task,
            item.get('note', ''), item.get('crop'), item.get('rotation', 0),
            item.get('regions'), item.get('subject', ''))
        if ok:
            submitted.append(payload)
        else:
            skipped.append({'id': item.get('id'), 'reason': payload})
            remain.append(item)
    save_staging(remain)
    return jsonify({'ok': True, 'submitted': submitted, 'skipped': skipped,
                    'submitted_count': len(submitted), 'remaining': len(remain)})


@app.route('/api/station/staging/clear', methods=['POST'])
def station_staging_clear():
    """清空临时区（全部丢弃，不可恢复）。"""
    if not _is_local_request():
        return jsonify({'error': '扫描端仅限本机设备操作（当前为远程访问）'}), 403
    items = load_staging()
    for item in items:
        path = _staging_image_path(item)
        if path:
            try:
                os.remove(path)
            except OSError:
                pass
    save_staging([])
    return jsonify({'ok': True, 'cleared': len(items)})


# ---------------------------- 用户端页面与接口 ----------------------------

def _teacher_visible_classes():
    """返回当前教师可见的班级集合；None 表示不限（管理员，或尚未建立归属数据）。"""
    role = session.get('user_role')
    uid = str(session.get('user_id') or '')
    if role == 'admin':
        return None
    owners = load_class_owners()
    owned = {c for c, lst in owners.items() if uid and uid in (lst or [])}
    owned |= {c for c, lst in owners.items() if not lst}
    return owned or None


def _filter_records_for_session(records, class_name='', student_id='', task_id='', status=''):
    role = session.get('user_role')
    uid = str(session.get('user_id') or '')
    if role == 'student':
        records = [r for r in records if str(r.get('student_id') or '') == uid]
    else:
        scope = _teacher_visible_classes()
        if scope is not None:
            records = [r for r in records if (r.get('class_name') or '') in scope]
    if class_name:
        records = [r for r in records if (r.get('class_name') or '') == class_name]
    if student_id:
        records = [r for r in records if (r.get('student_id') or '') == student_id]
    if task_id:
        records = [r for r in records if (r.get('task_id') or '') == task_id]
    if status:
        records = [r for r in records if (r.get('status') or 'pending') == status]
    return records


def _public_record(rec):
    out = dict(rec)
    # 姓名以当前名单为准，若记录里没有则实时补全（名单更新后可自愈）
    sid = str(out.get('student_id') or '')
    if sid and not out.get('student_name'):
        out['student_name'] = get_student_names_map().get(sid, '')
    out['image_url'] = '/api/portal/records/%s/image' % rec.get('id')
    out['essay_image_url'] = '/api/portal/records/%s/image?processed=1' % rec.get('id')

    grading = out.get('grading') if isinstance(out.get('grading'), dict) else None
    items = (grading or {}).get('items') or []
    regions = _clean_regions(out.get('regions'))
    if regions:
        # 一页多区域：前端只消费 pieces，避免在模板里反复对齐坐标与评分
        by_id = {str(i.get('id')): i for i in items}
        pieces = []
        for r in regions:
            it = by_id.get(str(r['id']), {})
            pieces.append({
                'id': r['id'],
                'label': r['label'],
                'score': it.get('score'),
                'max_score': it.get('max_score'),
                'level': it.get('level', ''),
                'comment': it.get('comment', ''),
                'text': it.get('text', ''),
                'strengths': it.get('strengths') or [],
                'issues': it.get('issues') or [],
                'image_url': '/api/portal/records/%s/image?region=%s' % (rec.get('id'), r['id']),
            })
        out['pieces'] = pieces
    else:
        out['pieces'] = []
    out['score'] = (grading or {}).get('total')
    out['max_total'] = (grading or {}).get('max_total')
    out['level'] = (grading or {}).get('level', '')
    if 'score_text' not in out:
        out['score_text'] = ''
    # 英语作文：结构化学情 + 批次级材料（话题语料、同学佳句）
    essay = out.get('essay') if isinstance(out.get('essay'), dict) else None
    out['essay'] = essay
    bid = str(out.get('batch_id') or '')
    if bid:
        b = _load_batch(bid)
        if b:
            out['batch'] = {
                'id': bid,
                'topic': b.get('topic') or {},
                'corpus': b.get('corpus') or {},
                'awards': (b.get('awards') or {}).get('awards') or [],
                'award_standard': (b.get('awards') or {}).get('standard', ''),
            }
    return out


GRADE_JSON_RULES = (
    '严格只输出一个 JSON 对象（不要 Markdown 代码块，不要任何解释文字）：\n'
    '{\n'
    '  "text": "图片中识别出的全部文字/算式，保留换行；没有文字则为空字符串",\n'
    '  "score": 数字或 null,\n'
    '  "max_score": 数字,\n'
    '  "level": "优秀|良好|合格|待改进",\n'
    '  "strengths": ["具体亮点1", "具体亮点2"],\n'
    '  "issues": ["具体问题1", "具体问题2"],\n'
    '  "comment": "给学生本人的一段话"\n'
    '}\n'
    '评分按满分 {max_score} 分计。若图中内容无法判断（空白、严重模糊、与题目无关），'
    'score 置为 null，并在 issues 中写清原因。\n'
    'comment 写法：先用一句话肯定亮点，再指出 1-2 处具体可改进之处，最后给 1-3 条可执行建议；'
    '语气温和、具体、有建设性，200 字以内；不要出现「系统」「AI」「学号」「分数」「模型」等字样。'
)


def _fmt_num(v):
    """把 8.0 显示成 8，把 8.5 保留一位小数。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    return str(int(f)) if f == int(f) else ('%.1f' % f)


def _level_from_ratio(ratio):
    if ratio is None:
        return ''
    if ratio >= 0.85:
        return '优秀'
    if ratio >= 0.70:
        return '良好'
    if ratio >= 0.60:
        return '合格'
    return '待改进'


def _parse_grade_json(raw, default_max_score):
    """解析模型返回的评分 JSON。容错：代码块包裹、前后废话、尾逗号、纯文本回退。"""
    s = (raw or '').strip()
    if not s:
        return None
    if s.startswith('```'):
        s = re.sub(r'^```[a-zA-Z]*\s*', '', s)
        s = re.sub(r'\s*```\s*$', '', s).strip()
    data = None
    start, end = s.find('{'), s.rfind('}')
    if start >= 0 and end > start:
        blob = s[start:end + 1]
        try:
            data = json.loads(blob)
        except Exception:
            try:
                data = json.loads(re.sub(r',\s*([}\]])', r'\1', blob))
            except Exception:
                data = None
    if not isinstance(data, dict):
        # 模型没按契约返回：把原文当识别文本与评语，分数留空，仍然可用
        return {'text': s, 'score': None, 'max_score': default_max_score, 'level': '',
                'strengths': [], 'issues': [], 'comment': s[:800], 'parsed': False}

    def _num(v):
        if isinstance(v, bool) or v is None:
            return None
        if isinstance(v, (int, float)):
            return float(v)
        if isinstance(v, str):
            m = re.search(r'-?\d+(?:\.\d+)?', v)
            if m:
                try:
                    return float(m.group())
                except ValueError:
                    return None
        return None

    def _strlist(v):
        if isinstance(v, str):
            v = [x for x in re.split(r'[\n；;]', v) if x.strip()]
        if not isinstance(v, list):
            return []
        out = []
        for x in v:
            t = str(x).strip().lstrip('-•· ').strip()
            if t:
                out.append(t[:300])
        return out[:8]

    try:
        mx = float(default_max_score or 100)
    except (TypeError, ValueError):
        mx = 100.0
    mx_model = _num(data.get('max_score'))
    if mx_model and 0 < mx_model <= 1000:
        mx = mx_model
    score = _num(data.get('score'))
    if score is not None:
        score = max(0.0, min(mx, score))

    level = str(data.get('level') or '').strip()
    if level not in ('优秀', '良好', '合格', '待改进'):
        level = _level_from_ratio((score / mx) if (score is not None and mx) else None)

    strengths = _strlist(data.get('strengths'))
    issues = _strlist(data.get('issues'))
    comment = str(data.get('comment') or '').strip()
    if not comment:
        bits = []
        if strengths:
            bits.append('亮点：' + '；'.join(strengths[:2]))
        if issues:
            bits.append('可改进：' + '；'.join(issues[:2]))
        comment = ' '.join(bits)[:800]

    return {
        'text': str(data.get('text') or '').strip(),
        'score': score,
        'max_score': mx,
        'level': level,
        'strengths': strengths,
        'issues': issues,
        'comment': comment,
        'parsed': True,
    }


def _vision_grade_one(image, subject_key=None, max_score=None, extra_prompt=None):
    """对「一块作答」调用视觉模型，返回 (结果 dict, 错误信息)。

    image 可以是 PIL.Image 或图片路径；结果含 text/score/max_score/level/strengths/issues/comment。
    subject_key 决定用哪个批阅方案（学科×题型），extra_prompt 可临时覆盖方案提示词。
    """
    if not minimax_api_key:
        return None, '未配置 LLM_API_KEY / MINIMAX_API_KEY，请在 .env 中设置后重试'
    try:
        img = _open_image(image) if isinstance(image, str) else image
        if img.mode != 'RGB':
            img = img.convert('RGB')
        max_size = 2000
        if max(img.size) > max_size:
            ratio = max_size / max(img.size)
            img = img.resize((int(img.size[0] * ratio), int(img.size[1] * ratio)), Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        data_url = 'data:image/png;base64,' + base64.b64encode(buf.getvalue()).decode('utf-8')

        subj = _get_grading_subject(subject_key)
        try:
            full = float(max_score or subj.get('max_score') or 100)
        except (TypeError, ValueError):
            full = 100.0
        base_prompt = (extra_prompt or '').strip() or subj.get('prompt') or ''
        base_prompt = base_prompt.replace('{max_score}', _fmt_num(full)).replace('{subject}', subj.get('name', ''))
        user_prompt = base_prompt + '\n\n' + GRADE_JSON_RULES.replace('{max_score}', _fmt_num(full))

        client = OpenAI(base_url=minimax_base_url, api_key=minimax_api_key, timeout=150.0)
        response = _vision_chat(
            client,
            model=minimax_model,
            messages=[
                {"role": "system", "content": "你是一位经验丰富的教师，擅长识别手写作答，并按评分标准给出准确、具体、对学生有帮助的评价。"},
                {"role": "user", "content": [
                    {"type": "image_url", "image_url": {"url": data_url}},
                    {"type": "text", "text": user_prompt}
                ]}
            ],
            temperature=0.2,
            max_tokens=3000
        )
        raw = (response.choices[0].message.content or '').strip()
        parsed = _parse_grade_json(raw, full)
        if not parsed:
            return None, '模型未返回有效结果，请重试'
        return parsed, ''
    except Exception as e:
        return None, '评阅失败：%s' % e


# ==================== 英语作文专项批阅（结构化学情）====================
# 与通用契约 GRADE_JSON_RULES 的差别：作文要输出可执行的「要点覆盖 / 语言错误 / 句子改写 / 佳句候选」，
# 学生端才能看到「哪里错了、应该怎么写」，而不是一句笼统评语。
# 一次视觉调用同时读「表头」与「全文」：表头用于自动分班，全文用于批阅，不额外增加调用次数。

ESSAY_MAX_SCORE = 25
ESSAY_HEADER_RECT = {'x': 0.035, 'y': 0.024, 'w': 0.90, 'h': 0.100}

ESSAY_PROMPT_TEMPLATE = '''你正在批阅一份中国高一学生的英语作文（手写扫描件）。

图片说明：第 1 张是整页试卷；若还有第 2 张，那是同一页顶部「班级 / 姓名 / 学号」栏的放大特写，读学号以特写为准。

请完成以下六件事，然后严格按 JSON 输出。

一、读表头
识别班级号、姓名、学号。学号为 4 位数字：前 2 位是班级号，后 2 位是班内座号。
手写数字常常一格一个（如 1 4 3 1），必须合并成一个 4 位数。姓氏姓名按字辨认，认不出就留空，不要猜。

二、识别正文
忠实转写学生手写的英文作文：保留换行；原样保留拼写与语法错误（不要顺手改对），因为后面要据此定位错误。
题干、注意事项、印刷体说明不要写进 text，只写学生手写的作文部分。

二·补、把「你自己没把握的地方」交出来（不要给书写打分）
手写卷子一定有看不清的字。凡是你只能靠上下文猜的词，转写时可照猜的写进 text，
但必须同时列进 illegible，让老师知道哪几处不可靠：
  as_read 你最终写成了什么 / guess 你猜它本来想写什么 / where 位置（第几段第几行）/
  why 为什么难读（连笔、笔画粘连、涂改覆盖、字太小、压线等）。
只报真正犹豫的；看得清就别报。全部看得清就返回空数组。
严禁在 why 里写「字迹清晰」却仍然列出来——自相矛盾的条目会被系统直接丢弃。
不要评价字迹美丑（"字丑"不是错误），只报影响识别的地方。

三、要点核对（防「要点偏离」）
按题目要求逐条判断：covered（写到了且清楚）/ partial（提到但不完整或不准确）/ missing（没写或写偏）。
evidence 摘录学生原文片段作为依据。若学生写了大量与题目无关的内容，相关要点判 missing，
并在 issues 里明确指出「偏离题意」。
注意：requirements 只写「内容要点」（该写什么），词数要求、书写要求、格式要求都不是要点，不要列进去。

四、语篇改错（这些条目会被直接画到学生原文上）
逐条列出真实存在的语言错误。要求：
· type 只能取以下中文标签之一：拼写 / 大小写 / 冠词 / 介词 / 单复数 / 所有格 / 固定搭配 /
  用词 / 语态 / 时态 / 动词不定式 / 句子成分残缺 / 表达冗余 / 表达混乱；
  若本题是读后续写，另可用：分词 / 动名词 / 连词 / 代词 / 词性误用 / 并列结构 / 从句 / 标点 /
  逗号连接两个完整句子。
· quote 必须从 text 里原样复制，不超过 8 个词（按空格数算，超长的会被系统舍弃），并且在 text 中只出现一次——
  系统靠它在原文上划掉错处、贴上改法。做不到唯一就换一段更长的片段。
· fix 只写改后的写法，不要写「应改为」；why 用一句话说明错在哪条规则上（这是给学生的错因）。
· 同一处不要重复报；拼写和语法重合时只报一条，取更具体的那一类。
· 有几处报几处，不要凑数；改前改后一模一样的不许出现。错误很多时最多 8 条，优先报影响理解的。
· 完全正确的作文返回空数组。

五、表达升级（回答「这几句怎么写得更好」）
· vocab_upgrades：3-4 条。挑「没有语法错误、但用词平」的表达（不是错句，错句归 errors）。
  original 从 text 原样摘（可以是词组）；upgrade 给更准或更地道的表达，不要生僻词；
  sentence 给出把升级项放回句子后的完整句子，必须保留学生原意与原有信息（不要顺带塞进别的新信息）；
  why 一句话说明为什么更好（搭配更地道 / 语义更准 / 语域更合适）。
· pattern_upgrades：2 条句式升级（最多 3 条）。
  original 从 text 原样摘一句；improved 必须同时「换句型 + 补内容」；
  technique 写手法名（非谓语作状语 / 定语从句 / 名词性从句 / 强调句 / 倒装 / 独立主格 / 平行结构 等）；
  added 写补了什么内容（如「补出结果：给师生一个喘息的机会」）——只换同义词不许算升级；
  why 一句话说明效果。
· cohesion：2 条逻辑衔接升级，只看衔接手段，不修语法（语法问题归 errors）。
  original 摘原句；improved 改写后必须真的出现衔接手段；
  relation 只能取：并列 / 递进 / 转折 / 因果 / 举例 / 对比 / 总结 / 目的；
  before 写原句用的衔接词（没有就写「无」），after 写改后用的衔接词，
  优先从这些常用词里选：however / therefore / as a result / moreover / furthermore / in addition /
  for example / by contrast / above all / to begin with / more importantly / with this in mind；
  所选的词必须真的表达 relation 说的那种关系，别硬套；
  why 点明这两句之间到底是什么关系（学生常常用错关系，比如该转折却用了 Thus）。
  **只报「该有却没有」或「关系用错」两类**；原句衔接词本来就用对了、改后只是换个近义词
  （如 Moreover → Furthermore），不算升级，不要报。宁可一条都不给，也不要凑。
· content_upgrade：仅当本题是「读后续写」时填 2-3 条（情节推进 / 细节 / 情感层次），字段同 cohesion，
  但把 relation 换成 node（如「第一段承接」「第二段转折」）；不是读后续写必须返回空数组 []。
· sentence_award：本篇最出彩的一句（只在确实存在时填写，否则 quote 留空）。

五·补、个性化范文
· model_essay：贴着这个学生写的「升级版」，不是通用范文。
  保留他的要点顺序、观点和素材——地名、活动名、人物一律沿用原文，不许换成别的景点或活动；
  把上面 vocab_upgrades / pattern_upgrades / cohesion 的升级真的用进去；
  **词数是硬指标**：题目给了词数就落在要求的 0.9–1.1 倍之内；题目没给就与原文词数相当（原文 ±15%）。
  宁可少写一个次要细节，也不要超词数——超词数的范文会把学生教坏；
  分段与原文一致，称呼与落款照抄原文。
· model_essay_note：一两句话，说明保留了什么、动了几处、还差什么（诚实说清这版为什么还不是满分）。
  若学生原文把事实写错了（如把杭州写成北京），范文里改对，并在 note 里点明这一点。
· word_count：actual 数一下 text 的实写词数；required 从题干里找「80 词左右」这类要求填数字，
  题干没写就填 null。

五、评分
满分 {max_score} 分。权重：内容要点 40%、语言准确性 35%、结构衔接 15%、书写与卷面 10%。
要点缺失或偏题必须显著扣分，不要给「安慰分」。
必须独立判断每篇的质量：不同学生能力差异明显，得分应当拉开；除非两篇确实水平相当，否则不要给相同分数。

输出（只输出 JSON，不要 Markdown 代码块，不要任何解释）：
{
  "student": {"class_no": "", "name": "", "student_id": ""},
  "task": {"title": "一句话概括本题写作任务", "requirements": ["要点1", "要点2", "要点3"]},
  "text": "学生作文全文",
  "illegible": [{"as_read": "你写成的样子", "guess": "猜的原词", "where": "第2段第1行", "why": "连笔"}],
  "requirements_check": [{"point": "要点", "status": "covered", "evidence": "原文片段"}],
  "errors": [{"type": "拼写", "quote": "原文片段", "fix": "改后写法", "why": "错在哪条规则"}],
  "vocab_upgrades": [{"original": "原文里的表达", "upgrade": "更准的表达", "sentence": "放进句子后的完整句子", "why": "为什么更好"}],
  "pattern_upgrades": [{"original": "原句", "improved": "改写句", "technique": "手法名", "added": "补了什么内容", "why": "为什么更好"}],
  "cohesion": [{"original": "原句", "improved": "改写句", "relation": "递进", "before": "无", "after": "More importantly", "why": "两句是什么关系"}],
  "content_upgrade": [],
  "model_essay": "个性化范文全文（保留称呼与落款）",
  "model_essay_note": "保留了什么、动了几处、还差什么",
  "word_count": {"actual": 0, "required": null},
  "sentence_award": {"quote": "本篇最出彩的一句", "why": "好在哪里"},
  "score": "本次得分（填数字）", "max_score": {max_score}, "level": "优秀|良好|合格|待改进",
  "strengths": ["具体亮点"],
  "issues": ["具体问题"],
  "comment": "给学生本人的一段话"
}

注意：上面 JSON 里的 "本次得分（填数字）" 与 "优秀|良好|合格|待改进" 都只是占位说明，
必须替换成你实际判定的数值与等级，不能原样输出。

comment 要求：先肯定亮点，再指出 1-2 处最值得改的地方（对应上面的 errors / pattern_upgrades / cohesion，
不要复述整张清单），最后给 1-3 条可执行建议；180 字以内，语气温和具体；不要出现「AI」「系统」「模型」「学号」等字样。

题目要求（若此处为空，说明这是本批第一份，请从图片中的题干自行提取并写入 task.requirements）：
{requirements}
'''

CORPUS_PROMPT = '''你是高中英语写作教研员。下面是一道英语写作任务的题干与要点。

请为学生准备一份「题目专属语料包」，全批同学共用。要求：

1. points：按题干要点分组，每个要点给 2-3 条语料（至少 2 条）。
   · point 原样抄题干要点（不要改写、不要合并）；
   · 若要点里混进了「词数」「格式」「书写」这类非内容要求，直接跳过，不要为它生成语料；
   · items 每条包含：
     text  一个可直接背诵的英文句子或短语，不超过 22 词，高一学生看得懂，不要生僻词；
     trans 与 text 逐句对应的中文翻译（必须给——学生靠它背诵）；
     usage 十个字以内说明用在哪，如「开头表态后补理由」。
2. 同一组里的句子要有分工：分别承担「表态 / 理由 / 举例或建议」，不要同义反复。
3. pitfalls：3 条中国学生在该话题下最容易犯的错误，每条给 wrong → right，并用 point 标明属于哪个要点。
4. 不要写「万能模板句」。每条都必须扣住本题的具体内容（人物、地点、活动名称），
   换一道题就不该能直接用。

只输出 JSON：
{"points":[{"point":"要点原文","items":[{"text":"...","trans":"...","usage":"..."}]}],
 "pitfalls":[{"point":"要点","wrong":"...","right":"..."}]}'''

AWARD_PROMPT = '''下面是一个班同一道英语写作题中，若干篇作文里各自被判定为「本篇最出彩」的句子。

请以高一英语教师的口吻，挑出最值得全班学习的 3-5 句。评判标准：
（1）语法正确、表意清晰，没有硬伤；
（2）略微超出高一平均水平——不是靠生僻词，而是句式或表达有意识
     （倒装、非谓语、定语从句、平行结构、比喻、精准的动词选择等）；
（3）确实是学生自己写的，不是模板套话。

对每句给出：why（具体点出用了什么句式或手法）、takeaway（其他同学可以怎么借鉴）。
不足 3 句时有多少给多少；完全没有达标句子就返回空数组。

只输出 JSON：
{"awards":[{"student_id":"原文中的学号","quote":"原句","why":"...","takeaway":"..."}],
 "standard":"一句话说明本次评选口径"}'''

BATCHES_DIR = os.path.join(DATA_DIR, 'batches')
CORPUS_CACHE_FILE = os.path.join(CONFIG_DIR, 'corpus_cache.json')
os.makedirs(BATCHES_DIR, exist_ok=True)


def _img_to_data_url(img, max_size=2400):
    """PIL 图片 -> data URL，长边压到 max_size，控制 token。

    max_size 的取值经过实测：A4/200dpi 扫描件是 1657×2322，旧值 2000 会把它压到
    1427×2000（等效 172dpi），手写小字（笔画仅 1 像素）先被自己压糊一次。
    2400 能保住 200dpi 原尺寸，300dpi 扫描件压到 2400 也有约 204dpi；
    代价是单页视觉 token 约 +30%，换来的是不主动丢信息。
    """
    if img.mode != 'RGB':
        img = img.convert('RGB')
    if max(img.size) > max_size:
        r = max_size / max(img.size)
        img = img.resize((int(img.size[0] * r), int(img.size[1] * r)), Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    return 'data:image/png;base64,' + base64.b64encode(buf.getvalue()).decode('utf-8')


# ---------------------------- 卷面质量体检（本地、零成本） ----------------------------

_QUALITY_TIP = {
    'small': '字号偏小',
    'thin': '笔画过细',
    'blur': '图像偏糊',
    'tilt': '页面倾斜',
    'light': '墨色偏淡',
}


def image_quality_metrics(path):
    """用本地算法量化「这张卷子机器好不好读」，不调模型、不花钱、几毫秒出结果。

    核心指标：
      · char_mm   等效字高（毫米）—— 手写英文字母高度低于 1.8mm 时，200dpi 下只剩
                  1 个像素的笔画，视觉模型必须靠猜，是所有识别错误的源头；
      · stroke_px 笔画宽度（像素）—— 1 像素意味着已经贴着采样极限；
      · sharp     拉普拉斯方差，判断扫描/拍摄是否失焦；
      · slant     文字行倾斜角度。

    返回 dict；任何环节失败都返回 None（调用方按「无数据」处理，不影响批阅）。
    """
    try:
        import numpy as np
        import cv2
    except ImportError:
        return None
    try:
        im = cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_GRAYSCALE)
        if im is None:
            return None
        h, w = im.shape[:2]
        if min(h, w) < 200:
            return None

        sharp = float(cv2.Laplacian(im, cv2.CV_64F).var())
        bw = cv2.adaptiveThreshold(im, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY_INV, 31, 12)
        ink = float((bw > 0).mean())

        dist = cv2.distanceTransform(bw, cv2.DIST_L2, 5)
        vals = dist[bw > 0]
        stroke = float(np.percentile(vals, 50)) if vals.size else 0.0

        n, labels, stats, _ = cv2.connectedComponentsWithStats(bw, 8)
        hs, area = stats[1:, cv2.CC_STAT_HEIGHT], stats[1:, cv2.CC_STAT_AREA]
        keep = (area >= 6) & (area <= 20000)
        char_h = float(np.median(hs[keep])) if keep.any() else 0.0
        specks = int((area < 12).sum())          # 噪点 / 细小笔画断裂
        blobs = int((area > 40000).sum())        # 大面积涂黑 / 阴影

        edges = cv2.Canny(im, 50, 150)
        lines = cv2.HoughLinesP(edges, 1, np.pi / 180, 120,
                                minLineLength=w // 3, maxLineGap=8)
        slant = 0.0
        if lines is not None:
            seg = lines.reshape(-1, 4) if getattr(lines, 'ndim', 2) == 3 else lines
            angs = [np.degrees(np.arctan2(y2 - y1, x2 - x1)) for x1, y1, x2, y2 in seg]
            angs = [a for a in angs if abs(a) < 12]
            if angs:
                slant = float(np.median(angs))

        # 有效 dpi：按 A4 宽 210mm 反推（扫描件几乎都是 A4）
        dpi = w / 8.27
        char_mm = char_h / dpi * 25.4 if dpi else 0.0

        flags = []
        # 阈值来自本批 5 张真实难卷的实测：
        #   200dpi 下手写笔画普遍就是 1 像素（stroke_px=1.0），字高 1.5~2.2mm，
        #   所以「笔画细」不能当异常信号，否则整批都会被判成难卷、全部升档，白花钱。
        # 这里只把「真的小到接近看不清 / 明显失焦 / 墨色过淡」当成难卷。
        if 0 < char_mm < 1.7:
            flags.append('small')
        if stroke and stroke < 0.9:
            flags.append('thin')
        if sharp < 1000:
            flags.append('blur')
        if abs(slant) > 3:
            flags.append('tilt')
        if ink < 0.012:
            flags.append('light')

        # 难卷判定（决定要不要升到精读档）：只在信息可能真的不够时触发，
        # 正常 200dpi 试卷不升档——实测 32B 对这类卷子的提升远不如预期。
        hard = bool(char_mm and char_mm < 1.4) or 'blur' in flags or 'light' in flags
        return {
            'dpi': int(round(dpi)),
            'char_mm': round(char_mm, 2),
            'char_px': int(round(char_h)),
            'stroke_px': round(stroke, 2),
            'sharp': int(round(sharp)),
            'ink_pct': round(ink * 100, 2),
            'slant': round(slant, 1),
            'specks': specks,
            'blobs': blobs,
            'flags': flags,
            'labels': [_QUALITY_TIP[f] for f in flags],
            'hard': hard,
            'level': 'hard' if hard else ('warn' if flags else 'ok'),
        }
    except Exception:
        return None


def quality_badge(q):
    """把质量指标压成一句话，给老师看。"""
    if not q:
        return ''
    if q.get('labels'):
        return '、'.join(q['labels'])
    return '卷面清晰'


HANDWRITING_PROMPT = '''这是学生手写英语答卷的扫描件。请只做一件事：判断「这份字迹对机器识别有多友好」。

不要评分、不要改错、不要评价美丑（字丑不是问题，读不准才是问题）。

要求：
1. 逐行转写学生手写部分的前 8 行英文，原样保留拼写错误，不要顺手改对。
2. 凡是你读不准、只能靠上下文猜的词，在转写里写成 [你猜的词?]，并且必须同时列进 unsure。
3. 真的看得清就不许列，不要为凑数编造。全部清楚时 unsure 返回空数组。

只输出 JSON（不要代码块）：
{"lines": ["第1行", "..."],
 "total_words": 你转写的总词数,
 "unsure": [{"word": "读不准的词", "why": "连笔/字太小/笔画粘连/涂改覆盖/压线", "conf": "低或中"}],
 "size": "偏大/正常/偏小/极小",
 "note": "一句话说明这份字迹卡在哪"}
'''


def assess_handwriting(path, model=None):
    """评估一份手写卷的「机器可读性」。返回 (dict, 错误信息)。

    设计上刻意不让模型给「字迹美丑」打 1-5 分——实测同图反复打分会在 3/4 之间漂移，
    而且风险点全是「1和7形似」这类套话。这里改用一个可核对的信号：
    让它逐行转写，把读不准的词标 `[词?]` 并列入 unsure，
    再用「不确定词占比」算出稳定的可读性等级，并叠加本地量化指标。
    """
    q = image_quality_metrics(path)
    if not minimax_api_key:
        return {'local': q}, '未配置 LLM_API_KEY'
    try:
        img = _open_image(path)
        use_model = model or resolve_vision_model(hard=bool((q or {}).get('hard')))
        client = OpenAI(base_url=minimax_base_url, api_key=minimax_api_key, timeout=180.0)
        resp = _vision_chat(
            client, model=use_model,
            messages=[
                {"role": "system", "content": "你是手写体识别专家，只报告「读不准」的地方，绝不编造不确定项。"},
                {"role": "user", "content": [
                    {'type': 'image_url', 'image_url': {'url': _img_to_data_url(img)}},
                    {'type': 'text', 'text': HANDWRITING_PROMPT},
                ]},
            ],
            temperature=0.1, max_tokens=1500)
        data = _loads_lenient((resp.choices[0].message.content or '').strip())
        if not isinstance(data, dict):
            return {'local': q}, '模型未返回有效 JSON'

        unsure = []
        for u in (data.get('unsure') or []):
            if not isinstance(u, dict):
                continue
            w = str(u.get('word') or '').strip()
            # 只收「词/短语」级：模型偶尔会把整句丢进来（"We are sure to have a great time"），
            # 那不是识别不确定，而是它在做阅读判断，收进来会把比例算高。
            if not w or len(w.split()) > 3:
                continue
            unsure.append({'word': w[:60], 'why': str(u.get('why') or '').strip()[:40],
                           'conf': str(u.get('conf') or '').strip()[:8]})
        lines = [str(x).strip()[:400] for x in (data.get('lines') or []) if str(x).strip()]
        total = data.get('total_words')
        try:
            total = int(total)
        except (TypeError, ValueError):
            total = sum(len(re.findall(r"[A-Za-z']+", l)) for l in lines)
        total = max(total, 1)
        ratio = len(unsure) / total

        # 等级阈值按实测校准：本批两张难卷不确定率都是 6.9%，
        # 若按「>5% 即很难」会被判 D，但实际上模型转写基本正确、只是个别地名没把握，
        # 所以 8% 以内归「偏难」，超过 12% 才算「很难」。
        score = 100.0
        if ratio > 0.12:
            score -= 50
        elif ratio > 0.08:
            score -= 35
        elif ratio > 0.03:
            score -= 20
        elif ratio > 0.01:
            score -= 8
        flags = set((q or {}).get('flags') or [])
        # 本地指标的权重刻意给足：不确定率是模型自报的，同一张图两次跑会差 1 个百分点；
        # 「字高只有 1.5mm」这类事实是稳定的，应该让它把等级拉开。
        for f, cut in (('small', 20), ('thin', 6), ('blur', 18), ('light', 6), ('tilt', 4)):
            if f in flags:
                score -= cut
        score = max(0, min(100, score))
        grade, label = ('A', '清晰') if score >= 85 else \
                       ('B', '较清晰') if score >= 65 else \
                       ('C', '偏难') if score >= 40 else ('D', '很难')

        return {
            'local': q,
            'ai': {'lines': lines, 'size': str(data.get('size') or '').strip()[:8],
                   'note': str(data.get('note') or '').strip()[:200]},
            'unsure': unsure[:15],
            'total_words': total,
            'unsure_ratio': round(ratio * 100, 1),
            'score': round(score),
            'grade': grade,
            'grade_label': label,
            'model': use_model,
            'tips': (['不确定词占 %.1f%%（%d/%d）' % (ratio * 100, len(unsure), total)]
                     if unsure else ['模型逐行转写未遇到读不准的词'])
                    + [_QUALITY_TIP[f] for f in (q or {}).get('flags') or []],
        }, ''
    except Exception as e:
        return {'local': q}, '评估失败：%s' % e



def _loads_lenient(raw):
    """从模型输出里抠出 JSON：容忍代码块围栏、前后废话、尾逗号。"""
    s = (raw or '').strip()
    if not s:
        return None
    if s.startswith('```'):
        s = re.sub(r'^```[a-zA-Z]*\s*', '', s)
        s = re.sub(r'\s*```\s*$', '', s).strip()
    start, end = s.find('{'), s.rfind('}')
    if start < 0 or end <= start:
        return None
    blob = s[start:end + 1]
    for cand in (blob, re.sub(r',\s*([}\]])', r'\1', blob)):
        try:
            data = json.loads(cand)
            if isinstance(data, dict):
                return data
        except Exception:
            continue
    return None


def _llm_json(prompt, system, max_tokens=1600, temperature=0.2):
    """纯文本任务：调文本模型并解析 JSON。返回 (dict, 错误信息)。"""
    if not minimax_api_key:
        return None, '未配置 LLM_API_KEY'
    try:
        client = OpenAI(base_url=minimax_base_url, api_key=minimax_api_key, timeout=90.0)
        resp = client.chat.completions.create(
            model=llm_text_model,
            messages=[{"role": "system", "content": system},
                      {"role": "user", "content": prompt}],
            temperature=temperature,
            max_tokens=max_tokens,
            **_text_model_kwargs()
        )
        raw = (resp.choices[0].message.content or '').strip()
        data = _loads_lenient(raw)
        if data is None:
            return None, '模型返回的不是有效 JSON：' + raw[:160]
        return data, ''
    except Exception as e:
        return None, '调用失败：%s' % e


def _as_text_list(v, limit=12):
    if isinstance(v, str):
        v = [x for x in re.split(r'[\n；;]', v) if x.strip()]
    if not isinstance(v, list):
        return []
    out = []
    for x in v:
        t = str(x).strip().lstrip('-•· ').strip()
        if t:
            out.append(t[:400])
    return out[:limit]


def _clip(v, n):
    return str(v or '').strip()[:n]


def _locate_quote(text, quote):
    """在原文里找回模型引用的片段，返回原文中的真实子串（找不到返回 ''）。

    模型给的 quote 常与原文差在大小写、空格、弯直引号上；直接 indexOf 会定位失败，
    整条错误就画不到原文上。这里三级回退：原文整串 → 忽略大小写 → 忽略空白/引号差异。
    """
    q = (quote or '').strip()
    if not q or not text:
        return ''
    if q in text:
        return q
    i = text.lower().find(q.lower())
    if i >= 0:
        return text[i:i + len(q)]
    compact = re.sub(r'[\s\u2018\u2019\u201c\u201d"\'`]+', '', q)
    if not compact:
        return ''
    pat = r'[\s\u2018\u2019\u201c\u201d"\'`]*'.join(re.escape(c) for c in compact)
    m = re.search(pat, text, re.I)
    return m.group(0) if m else ''


def _norm_essay_result(data, max_score=ESSAY_MAX_SCORE):
    """把模型返回的原始 JSON 归一化成前端可直接消费的结构。"""
    if not isinstance(data, dict):
        return None
    try:
        mx = float(data.get('max_score') or max_score or ESSAY_MAX_SCORE)
        if not (0 < mx <= 1000):
            mx = float(max_score or ESSAY_MAX_SCORE)
    except (TypeError, ValueError):
        mx = float(max_score or ESSAY_MAX_SCORE)

    def _num(v):
        if isinstance(v, bool) or v is None:
            return None
        if isinstance(v, (int, float)):
            return float(v)
        m = re.search(r'-?\d+(?:\.\d+)?', str(v))
        return float(m.group()) if m else None

    score = _num(data.get('score'))
    if score is not None:
        score = max(0.0, min(mx, score))
    level = str(data.get('level') or '').strip()
    if level not in ('优秀', '良好', '合格', '待改进'):
        level = _level_from_ratio((score / mx) if (score is not None and mx) else None)

    stu = data.get('student') if isinstance(data.get('student'), dict) else {}
    sid = re.sub(r'[^\d]', '', str(stu.get('student_id') or ''))
    cls = re.sub(r'[^\d]', '', str(stu.get('class_no') or ''))
    name = _clean_person_name(stu.get('name'))

    status_map = {'covered': 'covered', 'partial': 'partial', 'missing': 'missing',
                  '已覆盖': 'covered', '部分': 'partial', '缺失': 'missing', '未写': 'missing'}
    reqs = []
    for r in (data.get('requirements_check') or []):
        if not isinstance(r, dict):
            continue
        st = status_map.get(str(r.get('status') or '').strip().lower(),
                            status_map.get(str(r.get('status') or '').strip(), 'partial'))
        reqs.append({'point': str(r.get('point') or '').strip()[:200],
                     'status': st,
                     'evidence': str(r.get('evidence') or '').strip()[:300]})

    errors, _seen_err = [], set()
    text_raw = str(data.get('text') or '').strip()
    for e in (data.get('errors') or []):
        if not isinstance(e, dict):
            continue
        q = _clip(e.get('quote'), 200)
        f = _clip(e.get('fix'), 200)
        if not q and not f:
            continue
        # 「改了等于没改」的条目是模型凑数，丢掉
        if q and f and re.sub(r'\s+', '', q).lower() == re.sub(r'\s+', '', f).lower():
            continue
        # 尽量把 quote 换成原文里的真实写法，前端才能逐条画到原文上
        hit = _locate_quote(text_raw, q) if q else ''
        if hit:
            q = hit
        sig = (re.sub(r'\s+', ' ', q).lower(), re.sub(r'\s+', ' ', f).lower())
        if sig in _seen_err:      # 同一处重复报（拼写 + 语法各报一次）只留第一条
            continue
        _seen_err.add(sig)
        ty = _clip(e.get('type'), 16) or '语法'
        if ty.isascii():          # 兼容模型仍返回英文枚举的情况
            ty = ty.lower()
        errors.append({'type': ty, 'quote': q, 'fix': f, 'why': _clip(e.get('why'), 200)})
    # 按在原文中出现的先后排序：前端是从左往右画的，乱序会互相遮蔽
    errors.sort(key=lambda x: text_raw.find(x['quote']) if x['quote'] and x['quote'] in text_raw else 9999)
    errors = errors[:14]

    rewrites = []
    for w in (data.get('rewrites') or []):
        if not isinstance(w, dict):
            continue
        o, i = str(w.get('original') or '').strip(), str(w.get('improved') or '').strip()
        if not o or not i:
            continue
        rewrites.append({'original': o[:240], 'improved': i[:240],
                         'focus': str(w.get('focus') or '').strip()[:40]})

    # ---- 表达升级三块（v2）：词汇 / 句式 / 衔接 ----
    # 三块都是「原文原句 → 升级后」，只换同义词等于没升级，这类条目直接丢掉。
    def _same(a, b):
        return re.sub(r'[\s\W]+', '', a or '').lower() == re.sub(r'[\s\W]+', '', b or '').lower()

    vocab = []
    for v in (data.get('vocab_upgrades') or []):
        if not isinstance(v, dict):
            continue
        o, up = _clip(v.get('original'), 160), _clip(v.get('upgrade'), 160)
        if not o or not up or _same(o, up):
            continue
        vocab.append({'original': o, 'upgrade': up,
                      'sentence': _clip(v.get('sentence'), 320),
                      'why': _clip(v.get('why'), 160)})

    patterns = []
    for p in (data.get('pattern_upgrades') or []):
        if not isinstance(p, dict):
            continue
        o, im = _clip(p.get('original'), 320), _clip(p.get('improved'), 320)
        if not o or not im or _same(o, im):
            continue
        patterns.append({'original': o, 'improved': im,
                         'technique': _clip(p.get('technique'), 30),
                         'added': _clip(p.get('added'), 80),
                         'why': _clip(p.get('why'), 160)})

    cohesion = []
    for c in (data.get('cohesion') or []):
        if not isinstance(c, dict):
            continue
        o, im = _clip(c.get('original'), 320), _clip(c.get('improved'), 320)
        if not o or not im or _same(o, im):
            continue
        cohesion.append({'original': o, 'improved': im,
                         'relation': _clip(c.get('relation'), 12),
                         'before': _clip(c.get('before'), 30) or '无',
                         'after': _clip(c.get('after'), 40),
                         'why': _clip(c.get('why'), 160)})

    # 读后续写专用：按情节节点提升内容（非续写应为空）
    content_up = []
    for c in (data.get('content_upgrade') or []):
        if not isinstance(c, dict):
            continue
        o, im = _clip(c.get('original'), 320), _clip(c.get('improved'), 320)
        if not o or not im or _same(o, im):
            continue
        content_up.append({'original': o, 'improved': im,
                           'node': _clip(c.get('node'), 20) or '情节推进',
                           'added': _clip(c.get('added'), 80),
                           'why': _clip(c.get('why'), 160)})

    model_essay = _clip(data.get('model_essay'), 4000)
    model_note = _clip(data.get('model_essay_note'), 300)
    # 别把 note 截成半句话：退到最后一个句读
    if model_note and model_note[-1] not in '。！？.!?':
        cut = max(model_note.rfind(c) for c in '。；！？.;!?')
        if cut >= 40:
            model_note = model_note[:cut + 1]
    wc = data.get('word_count') if isinstance(data.get('word_count'), dict) else {}
    actual_wc = _num(wc.get('actual'))
    if not actual_wc and text_raw:
        actual_wc = len(re.findall(r"[A-Za-z][A-Za-z'-]*", text_raw))
    req_wc = _num(wc.get('required'))

    aw = data.get('sentence_award') if isinstance(data.get('sentence_award'), dict) else {}
    award = {'quote': str(aw.get('quote') or '').strip()[:300],
             'why': str(aw.get('why') or '').strip()[:200]}
    if len(award['quote']) < 10:
        award = {'quote': '', 'why': ''}

    task = data.get('task') if isinstance(data.get('task'), dict) else {}
    comment = str(data.get('comment') or '').strip()

    # 模型自报的「读不准」清单：比让它给书写打 1-5 分可靠得多。
    # 实测同一张卷子反复打分会在 3/4 之间漂移，但「哪几个词没把握」是稳定且可核对的。
    illegible = []
    for u in (data.get('illegible') or []):
        if not isinstance(u, dict):
            continue
        ar = str(u.get('as_read') or '').strip()
        gs = str(u.get('guess') or '').strip()
        why = str(u.get('why') or '').strip()
        if not ar and not gs:
            continue
        # 自相矛盾的条目：原因写「字迹清晰」却列进「读不准」，是模型凑数
        if '清晰' in why and ar.lower() == gs.lower():
            continue
        illegible.append({'as_read': ar[:80], 'guess': gs[:80],
                          'where': str(u.get('where') or '').strip()[:40],
                          'why': why[:60]})

    return {
        'student': {'class_no': cls[:2], 'student_id': sid if len(sid) in (4, 6) else '', 'name': name},
        'task': {'title': str(task.get('title') or '').strip()[:200],
                 'requirements': _as_text_list(task.get('requirements'), 10)},
        'text': text_raw,
        'illegible': illegible[:12],
        'requirements_check': reqs[:10],
        'errors': errors,
        'rewrites': rewrites[:5],
        'vocab_upgrades': vocab[:6],
        'pattern_upgrades': patterns[:4],
        'cohesion': cohesion[:4],
        'content_upgrade': content_up[:4],
        'model_essay': model_essay,
        'model_essay_note': model_note,
        'word_count': {'actual': int(actual_wc or 0), 'required': int(req_wc) if req_wc else None},
        'sentence_award': award,
        'score': score, 'max_score': mx, 'level': level,
        'strengths': _as_text_list(data.get('strengths'), 6),
        'issues': _as_text_list(data.get('issues'), 8),
        'comment': comment[:1200] or ' '.join(
            ['亮点：' + '；'.join(_as_text_list(data.get('strengths'), 2))] if data.get('strengths') else []),
    }


def _essay_words(s):
    """数英文词数：字母开头、可带连字符或撇号，标点不算。"""
    return len(re.findall(r"[A-Za-z][A-Za-z'’-]*", s or ''))


def _fit_model_essay(essay, note, required):
    """范文超出题目词数太多时，用文本模型压一次。

    提示词里已经写了「必须落在 0.9–1.1 倍」，但实测 8B 仍会写成一倍半到两倍长；
    超词数的范文会把「凑字数」的习惯教给学生，所以这里做一次兜底压缩。
    压缩失败就原样返回——宁长不丢，不要在批阅主流程里抛错。
    """
    if not essay or not required:
        return essay, note
    try:
        req = float(required)
    except (TypeError, ValueError):
        return essay, note
    if req < 20 or _essay_words(essay) <= req * 1.35:
        return essay, note
    prompt = ('题目要求这篇英语作文约 %d 词，下面这版写了 %d 词，太长了。请重写一版 **不超过 %d 词** 的范文。\n\n'
              '必须保留：称呼、落款，以及学生写到的每一个地点 / 活动（可以合并进同一句，如 A, B and C）。\n'
              '必须删掉：重复的形容词、空泛的抒情、同一个意思的第二次表达（如先 "peaceful and unique" 再说 "quite special"）。\n'
              '结构照这个来：称呼 / 一句说明写信目的 / 二三句按行程列出地点 / 一句期待回复 / 落款。\n'
              '词数是硬上限，超过 %d 词就是不合格。只输出英文全文，不要任何说明。\n\n%s'
              % (req, _essay_words(essay), int(req * 1.1), int(req * 1.1), essay))
    try:
        client = OpenAI(base_url=minimax_base_url, api_key=minimax_api_key, timeout=90.0)
        resp = client.chat.completions.create(
            model=llm_text_model,
            messages=[{"role": "system", "content": "你是高中英语教师，擅长在不丢失要点的前提下把作文压到规定词数。"},
                      {"role": "user", "content": prompt}],
            temperature=0.2, max_tokens=900, **_text_model_kwargs())
        new = (resp.choices[0].message.content or '').strip()
    except Exception:
        return essay, note
    if not new or _essay_words(new) >= _essay_words(essay):
        return essay, note
    tail = '（范文已按题目要求压到约 %d 词，删掉的是重复铺陈。）' % _essay_words(new)
    return new, ((note or '') + ' ' + tail).strip()


# ---------------------------- 并行批阅与重试（实测参数，别随手调） ----------------------------
# 实测（2026-10-26，temp/bench_concurrency.py，40 张规模推演）：硅基流动侧墙钟在 ~105s / ~153s
# 呈阶跃而不是线性增长，说明瓶颈是服务端排队，不是本地线程。6 路时 40 张约 12 分钟；
# 8–12 路只再快 2–3 分钟，却把单张尾延迟从 105s 拉到 153s（老师会以为卡死），所以默认 6 路。
LLM_BATCH_WORKERS = max(1, min(16, int(os.getenv('LLM_BATCH_WORKERS', '6') or 6)))
# 多个批次同时跑时共用这一个闸门：3 个班各开 6 路 = 18 个在飞请求，必然触发限流
LLM_MAX_CONCURRENCY = max(1, min(32, int(os.getenv('LLM_MAX_CONCURRENCY', '8') or 8)))
_LLM_SEM = threading.BoundedSemaphore(LLM_MAX_CONCURRENCY)

# 送图长边上限：2400 能保住 200dpi 原尺寸；难卷抬到 3300。
# 注意：只有原图长边 > 2400（即扫描 dpi ≥ 300）时抬 cap 才真正生效，
# 200dpi 扫出来的 1657px 图抬了也没用——分辨率这条必须先改扫描端 dpi。
IMG_MAX_DEFAULT = max(1200, int(os.getenv('LLM_IMG_MAX_DEFAULT', '2400') or 2400))
IMG_MAX_HARD = max(1200, int(os.getenv('LLM_IMG_MAX_HARD', '3300') or 3300))

_LLM_RETRY_HINTS = ('429', 'rate limit', 'too many requests', 'timeout', 'timed out',
                    'temporarily', 'overloaded', 'busy', '502', '503', '504',
                    'connection', 'reset by peer', 'json decode', '20015')


def _llm_retry(fn, tries=3, base_delay=2.0):
    """调用模型并重试「值得重试」的失败，返回 (结果, 错误信息)。

    并发后失败几乎都是 429 / 超时 / 网关抖动，退避重试 2–3 次基本都能过；
    key 错误、参数错误这类确定性失败不重试，避免白花钱。
    """
    last = ''
    for i in range(max(1, tries)):
        try:
            with _LLM_SEM:
                return fn(), ''
        except Exception as e:
            last = str(e)
            low = last.lower()
            if i >= tries - 1 or not any(h in low for h in _LLM_RETRY_HINTS):
                break
            time.sleep(base_delay * (2 ** i))
    return None, last


def _vision_grade_essay(page, requirements='', max_score=ESSAY_MAX_SCORE, model=None,
                        quality=None, img_max_size=None):
    """批阅一页英语作文：一次调用同时读表头（学号/班级/姓名）与全文。返回 (dict, 错误信息)。

    model        显式指定视觉模型；为空时按 `.env` 的精度档位决定，
                 并把 quality（本地卷面体检结果）传进去用于「难卷自动升档」。
    quality      image_quality_metrics() 的返回，可为 None。
    img_max_size 送图长边上限；为空时普通卷 2400、难卷 IMG_MAX_HARD（3300）。
    """
    if not minimax_api_key:
        return None, '未配置 LLM_API_KEY，请在 .env 中设置'
    try:
        hard = bool((quality or {}).get('hard'))
        use_model = model or resolve_vision_model(hard=hard)
        page_max = img_max_size or (IMG_MAX_HARD if hard else IMG_MAX_DEFAULT)
        img = _open_image(page) if isinstance(page, str) else page
        if img.mode != 'RGB':
            img = img.convert('RGB')
        parts = [{'type': 'image_url', 'image_url': {'url': _img_to_data_url(img, page_max)}}]
        try:
            head = _transform_image(img, 0, ESSAY_HEADER_RECT)
            if head.size[0] >= 60 and head.size[1] >= 14:
                parts.append({'type': 'image_url', 'image_url': {'url': _img_to_data_url(head, 1800)}})
        except Exception:
            pass
        prompt = ESSAY_PROMPT_TEMPLATE.replace('{max_score}', _fmt_num(max_score))
        prompt = prompt.replace('{requirements}', requirements or '（本批第一份，请自行提取）')
        parts.append({'type': 'text', 'text': prompt})

        client = OpenAI(base_url=minimax_base_url, api_key=minimax_api_key, timeout=240.0)

        def _call():
            resp = _vision_chat(
                client,
                model=use_model,
                messages=[
                    {"role": "system", "content": "你是经验丰富的高中英语教师，读手写扫描件准确，评分严格但有建设性。"},
                    {"role": "user", "content": parts}
                ],
                temperature=0.2,
                max_tokens=6000
            )
            # 用量日志：并发调参和成本核算都靠它，别删
            u = getattr(resp, 'usage', None)
            if u is not None:
                print('[essay] %s in=%s out=%s' % (
                    use_model, getattr(u, 'prompt_tokens', '?'), getattr(u, 'completion_tokens', '?')),
                    flush=True)
            return resp

        # 服务端偶发 20015「JSON decode error」（大请求体被网关截断）与限流/超时，退避重试
        t_call = time.time()
        resp, cerr = _llm_retry(_call, tries=3)
        if resp is None:
            return None, '批阅失败：%s' % cerr
        raw = (resp.choices[0].message.content or '').strip()
        print('[essay] 耗时 %.1fs（%s）' % (time.time() - t_call, use_model), flush=True)
        data = _loads_lenient(raw)
        if data is None:
            return None, '模型未返回有效 JSON'
        out = _norm_essay_result(data, max_score)
        if not out:
            return None, '模型返回结构异常'
        if not out.get('text') and not out.get('errors'):
            return None, '未能识别出作文内容（可能是空白页或图片过暗）'
        # 范文超词数兜底：压缩一次，并把实际词数写进 word_count.model 供前端提示
        wc = out.get('word_count') or {}
        out['model_essay'], out['model_essay_note'] = _fit_model_essay(
            out.get('model_essay'), out.get('model_essay_note'), wc.get('required'))
        wc['model'] = _essay_words(out.get('model_essay'))
        out['word_count'] = wc
        out['_model'] = use_model
        return out, ''
    except Exception as e:
        return None, '批阅失败：%s' % e


# ---------------------------- 批次：题干 / 语料 / 佳句 ----------------------------

def _batch_path(batch_id):
    return os.path.join(BATCHES_DIR, re.sub(r'[^a-zA-Z0-9]', '', str(batch_id or '')) + '.json')


_BATCH_MEM = {}


def _load_batch(batch_id):
    """读批次文件（按 mtime 缓存，避免记录列表里重复读盘）。"""
    p = _batch_path(batch_id)
    try:
        mt = os.path.getmtime(p)
    except OSError:
        return None
    hit = _BATCH_MEM.get(p)
    if hit and hit[0] == mt:
        return hit[1]
    try:
        with open(p, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception:
        return None
    _BATCH_MEM[p] = (mt, data)
    return data


def _save_batch(b):
    if not isinstance(b, dict) or not b.get('id'):
        return
    os.makedirs(BATCHES_DIR, exist_ok=True)
    p = _batch_path(b['id'])
    tmp = p + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(b, f, ensure_ascii=False, indent=2)
    os.replace(tmp, p)
    _BATCH_MEM.pop(p, None)


def _topic_key(topic):
    """话题指纹：题干 + 要点规范化后取散列，同一话题跨批次复用语料。"""
    if not isinstance(topic, dict):
        return ''
    raw = (str(topic.get('title') or '') + '|' + '|'.join(topic.get('requirements') or []))
    raw = re.sub(r'\s+', '', raw)
    if len(raw) < 6:
        return ''
    return hashlib.sha1(raw.encode('utf-8')).hexdigest()[:12]


def _bigrams(text):
    s = re.sub(r'[\s，。、；：,.;:!?！？"“”\'（）()\[\]]+', '', str(text or ''))
    return {s[i:i + 2] for i in range(len(s) - 1)}


def _topic_similarity(a, b):
    """字符二元组 Jaccard 相似度：题目措辞略有出入时也能认出是同一题。"""
    A, B = _bigrams(a), _bigrams(b)
    if not A or not B:
        return 0.0
    return len(A & B) / float(len(A | B))


def _topic_text(d):
    """取用于相似度比较的文本；同时兼容「话题对象」与「语料缓存条目」两种结构。"""
    if not isinstance(d, dict):
        return ''
    return (str(d.get('title') or d.get('topic') or '') + ' '
            + ' '.join(d.get('requirements') or []))


def _load_corpus_cache():
    if not os.path.exists(CORPUS_CACHE_FILE):
        return {}
    try:
        with open(CORPUS_CACHE_FILE, 'r', encoding='utf-8') as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def _save_corpus_cache(cache):
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        with open(CORPUS_CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print('写入语料缓存失败: %s' % e)


CORPUS_VERSION = 2


def _norm_corpus(data):
    """把模型返回的语料归一化成「按要点分组 + 逐条翻译」的结构。

    v2 之前是 phrases/sentences/patterns 三堆按语言形式分类，学生不知道用在哪，
    而且 sentences 有中文大意、phrases 没有翻译；现在统一成 points[].items[]，
    每条都必须有 text + trans，usage 可选。
    """
    if not isinstance(data, dict):
        return None
    points = []
    for p in (data.get('points') or []):
        if not isinstance(p, dict):
            continue
        items = []
        for it in (p.get('items') or []):
            if isinstance(it, str):          # 容错：模型偶尔只给一个字符串数组
                it = {'text': it}
            if not isinstance(it, dict):
                continue
            t = str(it.get('text') or it.get('en') or '').strip()
            if not t:
                continue
            items.append({'text': t[:240],
                          'trans': str(it.get('trans') or it.get('翻译') or '').strip()[:240],
                          'usage': str(it.get('usage') or it.get('scene') or '').strip()[:40]})
        point = str(p.get('point') or p.get('name') or '').strip()[:120]
        if items:
            points.append({'point': point, 'items': items[:4]})
    pitfalls = []
    for p in (data.get('pitfalls') or []):
        if not isinstance(p, dict):
            continue
        w = str(p.get('wrong') or '').strip()
        r = str(p.get('right') or '').strip()
        if not w and not r:
            continue
        pitfalls.append({'point': str(p.get('point') or '').strip()[:60],
                         'wrong': w[:200], 'right': r[:200]})
    if not points:
        return None
    return {'points': points[:6], 'pitfalls': pitfalls[:5]}


def _get_topic_corpus(tkey, topic, force=False):
    """题目语料：同题只生成一次，之后全批复用（省 token 的关键）。

    先按指纹精确命中；未命中时再做一次相似度匹配（同一道题在不同批次里，
    模型概括出的标题措辞往往不同，只靠散列会重复生成）。
    缓存只认 v2 结构：老缓存没有翻译、也没按要点分组，命中它等于让学生看旧版。
    """
    cache = _load_corpus_cache()
    if not force:
        hit = cache.get(tkey) if tkey else None
        if isinstance(hit, dict) and hit.get('version') == CORPUS_VERSION:
            return hit, ''
        want = _topic_text(topic)
        best_key, best_score = '', 0.0
        for k, v in cache.items():
            if not isinstance(v, dict) or v.get('version') != CORPUS_VERSION:
                continue
            score = _topic_similarity(want, _topic_text(v))
            if score > best_score:
                best_key, best_score = k, score
        if best_key and best_score >= 0.55:
            reused = dict(cache[best_key])
            reused['reused_from'] = best_key
            return reused, ''
    if not isinstance(topic, dict):
        return {}, '缺少话题信息'
    prompt = ('题目：%s\n要点：\n%s' % (
        topic.get('title') or '（未提供标题）',
        '\n'.join('· ' + str(r) for r in (topic.get('requirements') or [])) or '（未提供）'))
    data, err = _llm_json(prompt, CORPUS_PROMPT, max_tokens=2000)
    if err:
        return {}, err
    norm = _norm_corpus(data)
    if not norm:
        return {}, '语料结构异常（模型未按要点分组返回）'
    corpus = dict(norm)
    corpus.update({
        'version': CORPUS_VERSION,
        'topic': topic.get('title') or '',
        'requirements': topic.get('requirements') or [],
        'generated_at': _now_str(),
    })
    if tkey:
        cache[tkey] = corpus
        _save_corpus_cache(cache)
    return corpus, ''


def _pick_awards(essays, level_hint='高一'):
    """同学佳句：全批批完后一次调用横向评选，不逐篇调用。"""
    pool = []
    for e in essays:
        essay = e.get('essay') or {}
        aw = essay.get('sentence_award') or {}
        q = str(aw.get('quote') or '').strip()
        if len(q) < 10:
            continue
        stu = essay.get('student') or {}
        pool.append({'student_id': stu.get('student_id') or '',
                     'name': stu.get('name') or '',
                     'quote': q, 'why': (aw.get('why') or '')[:120]})
    if not pool:
        return {'awards': [], 'standard': '', 'pool_size': 0}
    payload = json.dumps({'level': level_hint, 'candidates': pool}, ensure_ascii=False)
    data, err = _llm_json(payload, AWARD_PROMPT, max_tokens=1200, temperature=0.3)
    if err:
        return {'awards': [], 'standard': '', 'pool_size': len(pool), 'error': err}
    awards = []
    for a in (data.get('awards') or []):
        if not isinstance(a, dict):
            continue
        q = str(a.get('quote') or '').strip()
        if not q:
            continue
        sid = re.sub(r'[^\d]', '', str(a.get('student_id') or ''))
        hit = next((p for p in pool if p['quote'][:20] == q[:20]), None)
        awards.append({'student_id': sid or (hit or {}).get('student_id', ''),
                       'name': (hit or {}).get('name', ''),
                       'quote': q[:300],
                       'why': str(a.get('why') or '').strip()[:240],
                       'takeaway': str(a.get('takeaway') or '').strip()[:200]})
    return {'awards': awards[:6], 'standard': str(data.get('standard') or '').strip()[:200],
            'pool_size': len(pool)}


# ---------------------------- 自动流水线：扫描 → 出结果 ----------------------------
_JOB_LOCK = threading.Lock()


def _new_batch(item_ids, subject='english_essay', task_id='', note=''):
    b = {
        'id': uuid.uuid4().hex[:16],
        'created_at': _now_str(),
        'subject': subject,
        'task_id': task_id or '',
        'note': note or '',
        'item_ids': list(item_ids or []),
        'topic': {},
        'corpus': {},
        'awards': {'awards': []},
        'record_ids': [],
        'job': {'status': 'queued', 'stage': 'queued', 'total': len(item_ids or []),
                'done': 0, 'message': '已排队，等待开始…'},
    }
    _save_batch(b)
    return b


def _auto_pipeline(batch_id):
    """零操作流水线：读表头 → 自动分班 → 逐篇批阅 → 话题语料 → 同学佳句 → 入库。"""
    b = _load_batch(batch_id)
    if not b:
        return

    def bump(**kw):
        job = b.setdefault('job', {})
        job.update({k: v for k, v in kw.items() if v is not None})
        job['updated_at'] = _now_str()
        _save_batch(b)

    try:
        live = {it['id']: it for it in load_staging()}
        pages = [live[i] for i in b.get('item_ids', []) if i in live]
        bump(status='running', stage='reading', done=0,
             total=len(pages), message='正在识别表头与正文…')
        if not pages:
            raise RuntimeError('临时区里找不到这批扫描件（可能已被清空）')

        # ---------------- 批阅：预热 1 张拿题干 → 其余并行 → 失败补跑 ----------------
        # 题干是本批共享的（第一份产出、其余复用），所以不能一上来全并行：
        # 先串行预热 1 张（只等一张），拿到 requirements 再并行剩下的。
        # 进度写盘只发生在主线程：_save_batch 是「读-改-写」，多线程一起写会丢更新。
        n = len(pages)
        results = [None] * n
        requirements_text = ''
        state = {'done': 0}

        def _grade_one(idx):
            """批一张，异常统一折成 error，返回 (idx, 结果)。"""
            it = pages[idx]
            path = _staging_image_path(it)
            if not path:
                return idx, {'item_id': it['id'], 'error': '图片文件缺失',
                             'quality': it.get('quality'), 'sec': 0}
            t0 = time.time()
            res, err = _vision_grade_essay(path, requirements_text, ESSAY_MAX_SCORE,
                                           quality=it.get('quality'))
            sec = round(time.time() - t0, 1)
            if err:
                return idx, {'item_id': it['id'], 'error': err,
                             'quality': it.get('quality'), 'sec': sec}
            return idx, {'item_id': it['id'], 'essay': res,
                         'quality': it.get('quality'), 'sec': sec}

        # ① 预热：最多试前 3 张，拿到题干就停（正常只等 1 张的量）
        for idx in range(min(3, n)):
            bump(message='正在预热第 %d 张（先确定本批题干）…' % (idx + 1))
            _, r = _grade_one(idx)
            results[idx] = r
            state['done'] += 1
            if r.get('essay') and not b.get('topic'):
                task = (r['essay'].get('task') or {})
                reqs = task.get('requirements') or []
                if reqs or task.get('title'):
                    b['topic'] = {'title': task.get('title', ''), 'requirements': reqs}
                    b['topic']['key'] = _topic_key(b['topic'])
            bump(done=state['done'])
            if b.get('topic'):
                break

        requirements_text = '\n'.join(
            '· ' + str(x) for x in ((b.get('topic') or {}).get('requirements') or []))

        # ② 其余并行
        rest = [i for i in range(n) if results[i] is None]
        if rest:
            workers = max(1, min(LLM_BATCH_WORKERS, len(rest)))
            bump(message='正在并行批阅剩余 %d 份（%d 路并行）…' % (len(rest), workers))
            with ThreadPoolExecutor(max_workers=workers, thread_name_prefix='grade') as ex:
                futs = {ex.submit(_grade_one, i): i for i in rest}
                for fut in as_completed(futs):
                    idx = futs[fut]
                    try:
                        _, r = fut.result()
                    except Exception as e:
                        it = pages[idx]
                        r = {'item_id': it['id'], 'error': '批阅异常：%s' % e,
                             'quality': it.get('quality'), 'sec': 0}
                    results[idx] = r
                    state['done'] += 1
                    bump(done=state['done'],
                         message='已批阅 %d/%d 份（%d 路并行）…' % (state['done'], n, workers))

        # ③ 失败补跑一次：并发下的失败多为限流/超时，不该让老师手动补
        bad = [i for i, r in enumerate(results)
               if r and r.get('error') and '缺失' not in str(r.get('error'))]
        if bad and len(bad) < n:
            bump(message='有 %d 份失败，正在补跑…' % len(bad))
            for idx in bad:
                _, r = _grade_one(idx)
                if r.get('essay'):
                    results[idx] = r

        results = [r for r in results if r]
        ok = [r for r in results if r.get('essay')]
        if not ok:
            raise RuntimeError('本批全部识别失败，请检查模型配置或扫描清晰度')

        # 话题语料（跨批缓存）
        tkey = _topic_key(b.get('topic'))
        if b.get('topic'):
            b['topic']['key'] = tkey
        bump(stage='corpus', message='正在生成话题语料…')
        corpus, cerr = _get_topic_corpus(tkey, b.get('topic'))
        b['corpus'] = corpus or {}
        if cerr:
            b['corpus_error'] = cerr

        # 同学佳句：按班级分组评选。老师分批上传时一批里可能混了两个班，
        # 跨班混着挑「佳句」没有意义，学生端也只该看到自己班的。
        bump(stage='awards', message='正在评选同学佳句…')
        by_class = {}
        for r in ok:
            sid_r = str(((r.get('essay') or {}).get('student') or {}).get('student_id') or '')
            cls_r = (student_id_to_class_code(sid_r) if sid_r else '') or '未分班'
            by_class.setdefault(cls_r, []).append(r)
        b['classes'] = sorted(by_class)
        if len(by_class) <= 1:
            b['awards'] = _pick_awards(ok)
        else:
            all_awards, pool_all, per = [], 0, {}
            for cls_r, rs in by_class.items():
                a = _pick_awards(rs)
                per[cls_r] = {'pool': a.get('pool_size') or 0,
                              'awards': len(a.get('awards') or [])}
                all_awards.extend(a.get('awards') or [])
                pool_all += a.get('pool_size') or 0
            b['awards'] = {'awards': all_awards[:8],
                           'standard': '本批跨 %d 个班，佳句按班分别评选' % len(by_class),
                           'pool_size': pool_all, 'by_class': per}

        # 分数雷同是模型「抄示例」的典型信号，明确提示老师抽查，而不是假装正常
        scores = [(r['essay'] or {}).get('score') for r in ok]
        scores = [s for s in scores if isinstance(s, (int, float))]
        if len(scores) >= 3 and len(set(scores)) == 1:
            b['warning'] = '本批 %d 份得分完全相同（%s），建议抽查或重新批阅' % (len(scores), _fmt_num(scores[0]))

        # 入库
        bump(stage='committing', message='正在归档到用户端…')
        commit_ids, failed = [], []
        resolved_map = {}
        for r in results:
            it = live.get(r.get('item_id'))
            if not it:
                continue
            essay = r.get('essay') or {}
            stu = essay.get('student') or {}
            sid_seen = str(stu.get('student_id') or '').strip() or str(it.get('student_id') or '').strip()
            # 卷面四位号（高一写后四位）靠班级号回查完整学号；学号是主键，姓名以名单为准
            cls, resolved_name, sid, sid_note = _ensure_student_registry(
                sid_seen, stu.get('name') or '', stu.get('class_no') or '')
            if not cls:
                # 学号没落到名单上的卷子统一进「待确认」，不要按后四位硬推班级
                cls = '待确认' if sid_note else str(stu.get('class_no') or '').strip()
            resolved_map[r.get('item_id')] = (sid, sid_note, resolved_name)
            extra = {
                'batch_id': b['id'],
                'essay': essay,
                'topic': b.get('topic') or {},
                'quality': r.get('quality') or {},
                'auto': True,
                'student_name': resolved_name or '',
                'essay_text': essay.get('text') or '',
                'ai_comment': essay.get('comment') or '',
                'score_text': ('%s/%s' % (_fmt_num(essay.get('score')), _fmt_num(essay.get('max_score')))
                               if essay.get('score') is not None else ''),
                'grading': {
                    'items': [essay] if essay else [],
                    'total': essay.get('score'),
                    'max_total': essay.get('max_score') or ESSAY_MAX_SCORE,
                    'level': essay.get('level', ''),
                    'subject': b.get('subject') or 'english_essay',
                    'subject_name': '英语作文',
                    'graded_at': _now_str(),
                },
                'status': 'graded' if essay else 'pending',
                'graded_at': _now_str() if essay else '',
                # 学号没能唯一落到名单上时，这份卷子必须人工确认归属
                'needs_review': (not sid or bool(sid_note)) and bool(essay),
                'sid_note': sid_note,
                'student_id_seen': sid_seen,
            }
            okc, rec = _commit_staging_item(
                it, sid, cls, b.get('task_id') or '', '自动批阅', None, 0,
                None, b.get('subject') or 'english_essay', extra)
            if okc:
                commit_ids.append(rec['id'])
            else:
                failed.append({'item_id': it['id'], 'reason': str(rec)})

        # 从临时区移除已入库的条目（多批同时跑时必须持锁，否则会互相覆盖索引）
        done_ids = {r['item_id'] for r in results if r.get('item_id')}
        with _STAGING_LOCK:
            save_staging([it for it in load_staging() if it.get('id') not in done_ids])

        b['record_ids'] = commit_ids
        b['results'] = []
        for r in results:
            stu_r = (r.get('essay') or {}).get('student') or {}
            rsid, rnote, rname = resolved_map.get(r.get('item_id'), ('', '', ''))
            b['results'].append({
                'item_id': r.get('item_id'),
                # 落库学号是回查后的完整学号；seen 是卷面上真实看到的（高一为后四位）
                'student_id': rsid or stu_r.get('student_id', ''),
                'student_id_seen': stu_r.get('student_id', ''),
                'name': rname or stu_r.get('name', ''),
                'sid_note': rnote,
                'score': (r.get('essay') or {}).get('score'),
                'level': (r.get('essay') or {}).get('level', ''),
                'model': (r.get('essay') or {}).get('_model', ''),
                'quality': r.get('quality') or {},
                'illegible': len((r.get('essay') or {}).get('illegible') or []),
                'sec': r.get('sec', 0),
                'error': r.get('error', ''),
            })
        b['failed'] = failed
        done_msg = '已完成 %d 份批阅，可直接在用户端查看' % len(commit_ids)
        # 卷面体检汇总：让老师知道哪几张值得人工过一眼，而不是盲目相信分数
        hard_n = sum(1 for r in results if (r.get('quality') or {}).get('hard'))
        illeg_n = sum(len((r.get('essay') or {}).get('illegible') or []) for r in ok)
        notes = []
        if hard_n:
            notes.append('%d 张卷面偏难' % hard_n)
        if illeg_n:
            notes.append('%d 处字迹读不准' % illeg_n)
        if notes:
            done_msg += '（%s，已在结果里标出）' % '、'.join(notes)
        if b.get('warning'):
            done_msg += '（' + b['warning'] + '）'
        bump(status='done', stage='done', done=len(pages), message=done_msg)
    except Exception as e:
        import traceback
        print('自动流水线失败: %s\n%s' % (e, traceback.format_exc()))
        bump(status='error', stage='error', message=str(e))


def _start_auto_batch(item_ids, subject='english_essay', task_id=''):
    """建批 + 起后台线程。返回 batch（job 已在写盘）。"""
    with _JOB_LOCK:
        b = _new_batch(item_ids, subject=subject, task_id=task_id)
        t = threading.Thread(target=_auto_pipeline, args=(b['id'],), daemon=True)
        t.start()
    return b


def _grade_essay_record(rec, extra_prompt=None):
    """对一条已入库的作文记录重新批阅（教师手动「重新生成」走这里）。"""
    path = _records_image_path(rec)
    if not path:
        return None, '图片文件缺失'
    reqs = ''
    b = _load_batch(rec.get('batch_id') or '')
    if b and isinstance(b.get('topic'), dict):
        reqs = '\n'.join('· ' + str(r) for r in (b['topic'].get('requirements') or []))
    essay, err = _vision_grade_essay(path, reqs, ESSAY_MAX_SCORE)
    if err:
        return None, err
    rec['essay'] = essay
    rec['essay_text'] = essay.get('text') or ''
    rec['ai_comment'] = essay.get('comment') or ''
    rec['score_text'] = ('%s/%s' % (_fmt_num(essay.get('score')), _fmt_num(essay.get('max_score')))
                         if essay.get('score') is not None else '')
    rec['grading'] = {'items': [essay], 'total': essay.get('score'),
                      'max_total': essay.get('max_score') or ESSAY_MAX_SCORE,
                      'level': essay.get('level', ''),
                      'subject': 'english_essay', 'subject_name': '英语作文',
                      'graded_at': _now_str()}
    rec['status'] = 'graded'
    rec['graded_at'] = _now_str()
    return rec['grading'], ''


def _grade_record(rec, extra_prompt=None):
    """对一条档案执行批阅：按区域逐块评分，汇总总分与等级。返回 (grading, 错误信息)。

    区域规则：
      · 未设区域 → 整页当作一块（沿用旧的单选区 crop 也视为一块）；
      · 设了 N 个区域（如小练习里的某几道题）→ 每块独立评分，
        满分按「方案的每块满分」计，合计满分 = 每块满分 × 块数。
    """
    # 英语作文走专项契约（要点覆盖 / 语言错误 / 句子改写），而不是通用评分
    if str(rec.get('subject') or '') == 'english_essay':
        return _grade_essay_record(rec, extra_prompt)
    path = _records_image_path(rec)
    if not path:
        return None, '图片文件缺失'
    try:
        base = _transform_image(_open_image(path), rec.get('rotation', 0), None)
    except Exception as e:
        return None, '图片读取失败：%s' % e

    regions = _item_regions(rec)
    if regions:
        pieces = [{'id': r['id'], 'label': r['label'], 'rect': r} for r in regions]
    else:
        pieces = [{'id': 'r1', 'label': '整页', 'rect': None}]

    subj = _get_grading_subject(rec.get('subject'))
    try:
        max_score = float(rec.get('max_score') or subj.get('max_score') or 100)
    except (TypeError, ValueError):
        max_score = 100.0

    items = []
    for p in pieces:
        img = _transform_image(base, 0, p['rect']) if p.get('rect') else base
        res, err = _vision_grade_one(img, rec.get('subject'), max_score, extra_prompt)
        if err:
            return None, ('%s：%s' % (p['label'], err)) if len(pieces) > 1 else err
        res['id'] = p['id']
        res['label'] = p['label']
        items.append(res)

    scored = [i for i in items if isinstance(i.get('score'), (int, float))]
    total = round(sum(i['score'] for i in scored), 1) if scored else None
    max_total = round(max_score * len(items), 1)
    ratio = (total / max_total) if (total is not None and max_total) else None

    grading = {
        'items': items,
        'total': total,
        'max_total': max_total,
        'level': _level_from_ratio(ratio),
        'subject': rec.get('subject') or '',
        'subject_name': subj.get('name', ''),
        'graded_at': _now_str(),
    }
    rec['grading'] = grading
    rec['score_text'] = (_fmt_num(total) + '/' + _fmt_num(max_total)) if total is not None else ''
    rec['essay_text'] = '\n\n'.join((i.get('text') or '') for i in items).strip()
    if len(items) == 1:
        rec['ai_comment'] = items[0].get('comment', '')
    else:
        rec['ai_comment'] = '\n\n'.join('【%s】%s' % (i['label'], i.get('comment', '')) for i in items)
    rec['status'] = 'graded' if (rec['ai_comment'] or total is not None) else 'pending'
    rec['graded_at'] = _now_str()
    return grading, ''


@app.route('/api/grading_subjects', methods=['GET'])
def get_grading_subjects():
    """批阅方案列表（学科 × 题型），供扫描端选择、供设置页编辑。"""
    return jsonify(grading_subject_options())


@app.route('/api/grading_subjects', methods=['POST', 'PUT'])
def save_grading_subjects():
    """保存批阅方案。body: {current?, subjects?: {key: {name,max_score,prompt}}}。"""
    if session.get('user_role') not in ('teacher', 'admin'):
        return jsonify({'error': '仅教师或管理员可修改批阅方案'}), 403
    data = request.get_json(silent=True) or {}
    cfg = load_grading_config()
    subs = data.get('subjects')
    if isinstance(subs, dict):
        for key, val in subs.items():
            key = str(key).strip()
            if not key or not isinstance(val, dict):
                continue
            cur = dict(cfg['subjects'].get(key) or {})
            if 'name' in val:
                cur['name'] = str(val.get('name') or key)[:60]
            if 'max_score' in val:
                try:
                    cur['max_score'] = max(0.0, float(val.get('max_score')))
                except (TypeError, ValueError):
                    pass
            if 'prompt' in val:
                cur['prompt'] = str(val.get('prompt') or '')
            if not cur:
                continue
            cur.setdefault('name', key)
            cur.setdefault('max_score', 100)
            cur.setdefault('prompt', '')
            cfg['subjects'][key] = cur
    cur_key = str(data.get('current') or '').strip()
    if cur_key in cfg['subjects']:
        cfg['current'] = cur_key
    try:
        save_grading_config(cfg)
    except Exception as e:
        return jsonify({'error': '保存失败：%s' % e}), 500
    return jsonify({'ok': True, **grading_subject_options()})


@app.route('/portal')
def portal_page():
    """用户端门户：任意设备浏览器访问，按登录角色呈现教师/学生视图。"""
    return render_template('portal.html')


@app.route('/api/portal/me', methods=['GET'])
def portal_me():
    """当前登录用户信息（含姓名）。"""
    if not session.get('user_id'):
        return jsonify({'logged_in': False})
    return jsonify({
        'logged_in': True,
        'user': {'id': session.get('user_id'), 'name': session.get('user_name', ''),
                 'role': session.get('user_role', 'teacher')},
    })


@app.route('/api/portal/overview', methods=['GET'])
def portal_overview():
    """门户首页数据：教师=各班已扫描汇总；学生=本人记录汇总。"""
    if not session.get('user_id'):
        return jsonify({'error': '未登录'}), 401
    role = session.get('user_role', 'teacher')
    uid = str(session.get('user_id') or '')
    records = load_records()

    if role == 'student':
        mine = [r for r in records if str(r.get('student_id') or '') == uid]
        return jsonify({
            'role': 'student',
            'total': len(mine),
            'graded': len([r for r in mine if r.get('status') == 'graded']),
            'records': [_public_record(r) for r in mine],
        })

    scoped = _filter_records_for_session(records)
    names_map = get_student_names_map()
    owners = load_class_owners()
    groups = {}
    for r in scoped:
        c = r.get('class_name') or '未归档'
        g = groups.setdefault(c, {'class_name': c, 'total': 0, 'graded': 0, 'pending': 0,
                                  'students': set(), 'last_at': ''})
        g['total'] += 1
        if r.get('status') == 'graded':
            g['graded'] += 1
        else:
            g['pending'] += 1
        if r.get('student_id'):
            g['students'].add(str(r['student_id']))
        if (r.get('submitted_at') or '') > g['last_at']:
            g['last_at'] = r.get('submitted_at') or ''
    group_list = []
    for c, g in groups.items():
        group_list.append({
            'class_name': c,
            'total': g['total'],
            'graded': g['graded'],
            'pending': g['pending'],
            'student_count': len(g['students']),
            'last_at': g['last_at'],
            'claimed_by_me': uid in (owners.get(c) or []),
            'owners': owners.get(c) or [],
        })
    group_list.sort(key=lambda x: x['last_at'], reverse=True)

    # 可认领的班级（有名字但尚无归属）——用于教师认领自己的班
    all_classes = [c['name'] for c in _load_classes_roster()]
    unclaimed = [c for c in all_classes if not (owners.get(c) or [])]
    return jsonify({
        'role': role,
        'is_admin': role == 'admin',
        'total': len(scoped),
        'graded': len([r for r in scoped if r.get('status') == 'graded']),
        'pending': len([r for r in scoped if r.get('status') != 'graded']),
        'classes': group_list,
        'all_classes': all_classes,
        'unclaimed_classes': unclaimed,
        'student_names': names_map,
    })


@app.route('/api/portal/records', methods=['GET'])
def portal_records():
    """按条件列出已扫描档案（教师视图）。"""
    if not session.get('user_id'):
        return jsonify({'error': '未登录'}), 401
    records = _filter_records_for_session(
        load_records(),
        class_name=(request.args.get('class_name') or '').strip(),
        student_id=(request.args.get('student_id') or '').strip(),
        task_id=(request.args.get('task_id') or '').strip(),
        status=(request.args.get('status') or '').strip(),
    )
    records = sorted(records, key=lambda r: r.get('submitted_at') or '', reverse=True)
    return jsonify({'records': [_public_record(r) for r in records], 'count': len(records)})


@app.route('/api/portal/records/<rec_id>/image')
def portal_record_image(rec_id):
    """档案图片。

    默认原图；?processed=1 返回旋转（单 crop 时含裁剪）后的整页；
    ?region=<区域id> 返回该区域的裁剪图（一页多区域时使用）。
    """
    if not session.get('user_id'):
        return jsonify({'error': '未登录'}), 401
    rec = _find_by_id(load_records(), rec_id)
    if not rec:
        return jsonify({'error': '记录不存在'}), 404
    if not _filter_records_for_session([rec]):
        return jsonify({'error': '无权查看该记录'}), 403
    path = _records_image_path(rec)
    if not path:
        return jsonify({'error': '图片文件缺失'}), 404
    region_id = (request.args.get('region') or '').strip()
    if region_id:
        region = _find_region(_clean_regions(rec.get('regions')), region_id)
        if not region:
            return jsonify({'error': '区域不存在'}), 404
        try:
            img = _transform_image(_open_image(path), rec.get('rotation', 0), region)
            buf = io.BytesIO()
            img.save(buf, format='PNG')
            return Response(buf.getvalue(), mimetype='image/png')
        except Exception as e:
            return jsonify({'error': '区域取图失败：%s' % e}), 500
    if request.args.get('processed'):
        try:
            crop = None if _clean_regions(rec.get('regions')) else rec.get('crop')
            img = _transform_image(_open_image(path), rec.get('rotation', 0), crop)
            buf = io.BytesIO()
            img.save(buf, format='PNG')
            return Response(buf.getvalue(), mimetype='image/png')
        except Exception:
            pass
    return send_from_directory(RECORDS_IMAGES_DIR, os.path.basename(path))


@app.route('/api/portal/records/<rec_id>/grade', methods=['POST'])
def portal_record_grade(rec_id):
    """执行 AI 批阅（教师）：按批阅方案逐块识别、评分并产出面向学生的评语。"""
    if session.get('user_role') not in ('teacher', 'admin'):
        return jsonify({'error': '仅教师或管理员可生成评语'}), 403
    records = load_records()
    rec = _find_by_id(records, rec_id)
    if not rec:
        return jsonify({'error': '记录不存在'}), 404
    if not _filter_records_for_session([rec]):
        return jsonify({'error': '无权操作该记录'}), 403
    path = _records_image_path(rec)
    if not path:
        return jsonify({'error': '图片文件缺失'}), 404
    data = request.get_json(silent=True) or {}
    grading, err = _grade_record(rec, data.get('prompt_template'))
    if err:
        return jsonify({'error': err}), 500
    save_records(records)
    return jsonify({'ok': True, 'record': _public_record(rec), 'grading': grading})


@app.route('/api/portal/records/<rec_id>/comment', methods=['POST'])
def portal_record_comment(rec_id):
    """教师填写/修改评语（可覆盖 AI 评语），或调整学号归属。"""
    if session.get('user_role') not in ('teacher', 'admin'):
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    data = request.get_json(silent=True) or {}
    records = load_records()
    rec = _find_by_id(records, rec_id)
    if not rec:
        return jsonify({'error': '记录不存在'}), 404
    if not _filter_records_for_session([rec]):
        return jsonify({'error': '无权操作该记录'}), 403
    if 'student_id' in data:
        sid = str(data.get('student_id') or '').strip()
        rec['student_id'] = sid
        rec['student_name'] = get_student_names_map().get(sid, '') if sid else ''
        if sid and not rec.get('class_name'):
            rec['class_name'] = student_id_to_class_code(sid) or ''
    if 'class_name' in data:
        rec['class_name'] = str(data.get('class_name') or '').strip()
    if 'ai_comment' in data:
        rec['ai_comment'] = str(data.get('ai_comment') or '')
        rec['status'] = 'graded' if rec['ai_comment'].strip() else 'pending'
        rec['graded_at'] = _now_str()
    if 'score_text' in data:
        # 允许教师直接改写分数（如「8/10」），不再依赖模型
        rec['score_text'] = str(data.get('score_text') or '').strip()
        rec['status'] = 'graded' if (rec.get('ai_comment') or rec['score_text']) else 'pending'
    if 'note' in data:
        rec['note'] = str(data.get('note') or '')
    save_records(records)
    return jsonify({'ok': True, 'record': _public_record(rec)})


@app.route('/api/portal/records/<rec_id>', methods=['DELETE'])
def portal_record_delete(rec_id):
    """删除一条档案（仅教师/管理员）。"""
    if session.get('user_role') not in ('teacher', 'admin'):
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    records = load_records()
    rec = _find_by_id(records, rec_id)
    if not rec:
        return jsonify({'error': '记录不存在'}), 404
    if not _filter_records_for_session([rec]):
        return jsonify({'error': '无权操作该记录'}), 403
    path = _records_image_path(rec)
    if path:
        try:
            os.remove(path)
        except OSError:
            pass
    save_records([r for r in records if r.get('id') != rec_id])
    return jsonify({'ok': True, 'id': rec_id})


@app.route('/api/portal/classes/claim', methods=['POST'])
def portal_class_claim():
    """教师认领/放弃班级归属，用于「只看到自己班级」。"""
    if session.get('user_role') not in ('teacher', 'admin'):
        return jsonify({'error': '仅教师或管理员可操作'}), 403
    data = request.get_json(silent=True) or {}
    class_name = str(data.get('class_name') or '').strip()
    action = str(data.get('action') or 'claim').strip()
    if not class_name:
        return jsonify({'error': '请指定班级'}), 400
    uid = str(session.get('user_id') or '')
    owners = load_class_owners()
    lst = owners.get(class_name) or []
    if action == 'release':
        lst = [x for x in lst if x != uid]
    else:
        if uid and uid not in lst:
            lst.append(uid)
    if lst:
        owners[class_name] = lst
    else:
        owners.pop(class_name, None)
    save_class_owners(owners)
    return jsonify({'ok': True, 'class_name': class_name, 'owners': owners.get(class_name) or []})


@app.route('/api/portal/export.csv')
def portal_export_csv():
    """按当前筛选导出档案清单 CSV（教师用）。"""
    if not session.get('user_id'):
        return jsonify({'error': '未登录'}), 401
    records = _filter_records_for_session(
        load_records(),
        class_name=(request.args.get('class_name') or '').strip(),
        task_id=(request.args.get('task_id') or '').strip(),
        status=(request.args.get('status') or '').strip(),
    )
    records = sorted(records, key=lambda r: (r.get('class_name') or '', r.get('student_id') or ''))
    names_map = get_student_names_map()
    out = io.StringIO()
    writer = csv.writer(out)
    writer.writerow(['班级', '学号', '姓名', '任务', '批阅方案', '得分', '等级', '状态', '提交时间', '评语'])
    for r in records:
        sid = str(r.get('student_id') or '')
        writer.writerow([
            r.get('class_name', ''), sid, r.get('student_name') or names_map.get(sid, ''),
            r.get('task_title', ''), r.get('subject_name', ''), r.get('score_text', ''),
            (r.get('grading') or {}).get('level', '') if isinstance(r.get('grading'), dict) else '',
            '已评' if r.get('status') == 'graded' else '待评',
            r.get('submitted_at', ''),
            (r.get('ai_comment') or '').replace('\r\n', '\n').replace('\n', ' '),
        ])
    body = out.getvalue().encode('utf-8-sig')
    res = Response(body, mimetype='text/csv; charset=utf-8')
    res.headers['Content-Disposition'] = 'attachment; filename="portal_records.csv"'
    return res


if __name__ == '__main__':
    # 初始化NAPS2 PowerShell函数（可选，主要用于测试）
    print("正在初始化NAPS2 PowerShell函数...")
    initialize_naps2_powershell_function()

    # 服务固定绑 0.0.0.0，局域网内其它设备才能访问。
    # 但调试器不能跟着开：Werkzeug 的交互式调试终端一旦暴露在网段上，
    # 等同于给所有人一个 Python 执行入口。要热重载时临时设 QJ_DEBUG=1。
    debug_mode = os.environ.get('QJ_DEBUG', '0').strip() == '1'
    app.config['TEMPLATES_AUTO_RELOAD'] = True      # 改模板不用重启（页面刷新即生效）

    print("启动 Flask 应用：监听 0.0.0.0:%d（局域网内可访问）" % PORTAL_PORT)
    print("  本机访问：  http://127.0.0.1:%d/" % PORTAL_PORT)
    for _u in _lan_portal_urls():
        print("  局域网访问：" + _u)
    if debug_mode:
        print("  ⚠ 调试模式已开启（QJ_DEBUG=1），请勿在公共网络下使用。")
    app.run(host='0.0.0.0', port=PORTAL_PORT, debug=debug_mode, use_reloader=debug_mode)

