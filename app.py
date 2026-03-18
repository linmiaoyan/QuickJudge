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

LOGO_DIR = os.path.join(os.path.dirname(__file__), 'logo')

@app.route('/favicon.ico')
def favicon():
    """返回favicon图标（使用根目录 logo 文件夹中的 ico）"""
    return send_from_directory(LOGO_DIR, 'favicon.ico', mimetype='image/vnd.microsoft.icon')

@app.route('/logo/<path:filename>')
def logo_file(filename):
    """提供 logo 目录下的图片（如 4.png 横向 logo）"""
    return send_from_directory(LOGO_DIR, filename)

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
DEFAULT_PROMPT_TEMPLATE = """请对以下英语作文进行详细批阅，给出：
1. 总体评分
2. 优点分析
3. 问题指出（语法、拼写、结构等）
4. 改进建议
5. 具体修改意见

作文内容：
{essay_text}
"""

# PNG压缩质量配置存储
PNG_QUALITY_CONFIG_FILE = os.path.join(CONFIG_DIR, 'png_quality_config.json')
DEFAULT_PNG_QUALITY = 95  # 默认压缩质量（0-100，100为最高质量）

# NAPS2扫描仪配置存储
NAPS2_CONFIG_FILE = os.path.join(CONFIG_DIR, 'naps2_config.json')
DEFAULT_NAPS2_PATH = r"D:\Download\NAPS2\App\NAPS2.Console.exe"

# 扫描输出目录配置
SCAN_OUTPUT_CONFIG_FILE = os.path.join(CONFIG_DIR, 'scan_output_config.json')
DEFAULT_SCAN_OUTPUT_DIR = r"C:\Users\65789\Pictures\scan"

# 扫描仪高级设置（图像模式、单双面等，与打印机/扫描相关保留）
SCANNER_ADVANCED_CONFIG_FILE = os.path.join(CONFIG_DIR, 'scanner_advanced_config.json')
DEFAULT_SCANNER_IMAGE_MODE = 'color'   # grayscale | color
DEFAULT_SCANNER_SCAN_TYPE = 'double'   # single | double

# 讯飞OCR配置（从环境变量读取，如果没有则使用默认值）
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
    """获取扫描输出目录配置"""
    try:
        if os.path.exists(SCAN_OUTPUT_CONFIG_FILE):
            with open(SCAN_OUTPUT_CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                scan_output_dir = config.get('scan_output_dir', DEFAULT_SCAN_OUTPUT_DIR)
                scan_output_dir = scan_output_dir.strip()
                # 确保目录存在
                os.makedirs(scan_output_dir, exist_ok=True)
                return scan_output_dir
    except Exception:
        pass
    # 默认路径，确保目录存在
    os.makedirs(DEFAULT_SCAN_OUTPUT_DIR, exist_ok=True)
    return DEFAULT_SCAN_OUTPUT_DIR

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
    根据学号推导班级编码。规则：前2位=年份，第3-4位=班号，第5-6位=学号。
    例如 230205 -> 23年2班5号 -> 班级编码 2302。
    仅当学号为6位数字时返回前4位，否则返回 None。
    """
    s = str(student_id or '').strip()
    if len(s) != 6 or not s.isdigit():
        return None
    return s[:4]


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

@app.route('/api/scan', methods=['POST'])
def start_scan():
    """开始扫描"""
    try:
        import subprocess
        
        # 获取扫描输出路径（使用配置的路径）
        scan_output_dir = get_scan_output_dir()
        
        # 获取下一个可用的文件名（格式：scan1.png, scan2.png, ...）
        next_filename = get_next_scan_filename(scan_output_dir)
        scan_output_path = os.path.join(scan_output_dir, next_filename)
        
        # 获取NAPS2路径配置
        naps2_path = get_naps2_path()
        
        # 检查NAPS2程序是否存在
        if not os.path.exists(naps2_path):
            return jsonify({
                'error': f'NAPS2程序不存在，请检查路径是否正确。\n当前路径：{naps2_path}\n请在设置中配置正确的路径。'
            }), 400
        
        # 构建PowerShell命令
        # 使用单引号包裹路径，避免反斜杠转义问题
        # 将路径中的单引号转义为两个单引号（PowerShell的转义方式）
        naps2_path_escaped = naps2_path.replace("'", "''")
        scan_output_path_escaped = scan_output_path.replace("'", "''")
        
        ps_command = f'''
$naps2ConsolePath = '{naps2_path_escaped}'
if (-not (Test-Path -Path $naps2ConsolePath -PathType Leaf)) {{
    Write-Error "错误：未找到NAPS2控制台程序"
    exit 1
}}
& $naps2ConsolePath -o '{scan_output_path_escaped}'
'''
        
        # 执行PowerShell命令
        print(f"执行扫描命令，输出路径: {scan_output_path}")
        result = subprocess.run(
            ['powershell', '-ExecutionPolicy', 'Bypass', '-Command', ps_command],
            capture_output=True,
            text=True,
            timeout=120  # 120秒超时
        )
        
        # 打印调试信息
        print(f"扫描命令返回码: {result.returncode}")
        if result.stdout:
            print(f"扫描命令输出: {result.stdout}")
        if result.stderr:
            print(f"扫描命令错误: {result.stderr}")
        
        if result.returncode != 0:
            error_msg = result.stderr or result.stdout or '扫描失败，未知错误'
            return jsonify({
                'error': f'扫描失败：{error_msg}',
                'returncode': result.returncode
            }), 500
        
        # 检查输出文件是否生成（使用绝对路径）
        scan_output_path_normalized = os.path.normpath(scan_output_path)
        print(f"检查文件是否存在: {scan_output_path_normalized}")
        print(f"文件是否存在: {os.path.exists(scan_output_path_normalized)}")
        
        # 如果期望的文件不存在，查找NAPS2实际生成的文件（可能是scan1.1.png, scan1.2.png等格式）
        if not os.path.exists(scan_output_path_normalized):
            import time
            scan_start_time = time.time() - 10  # 扫描开始前10秒（给一些余量）
            
            if os.path.exists(scan_output_dir):
                files_in_dir = os.listdir(scan_output_dir)
                print(f"扫描输出目录中的文件: {files_in_dir}")
                
                # 查找最近修改的scan*.png文件（可能是NAPS2自动生成的格式，如scan1.1.png）
                latest_file = None
                latest_time = 0
                for filename in files_in_dir:
                    if filename.lower().startswith('scan') and filename.lower().endswith('.png'):
                        file_path = os.path.join(scan_output_dir, filename)
                        file_time = os.path.getmtime(file_path)
                        if file_time > latest_time and file_time > scan_start_time:
                            latest_time = file_time
                            latest_file = filename
                
                # 如果找到了新生成的文件，直接使用它（不重命名）
                if latest_file:
                    scan_output_path_normalized = os.path.join(scan_output_dir, latest_file)
                    next_filename = latest_file
                    print(f"找到NAPS2生成的文件: {latest_file}")
                else:
                    # 如果没有找到新文件，返回警告
                    return jsonify({
                        'warning': f'扫描命令执行完成，但未找到输出文件。\n期望文件: {scan_output_path_normalized}\n目录中的文件: {", ".join(files_in_dir[:10])}\n请检查扫描仪是否正常工作。',
                        'expected_path': scan_output_path_normalized,
                        'files_in_dir': files_in_dir[:10]
                    }), 200
        
        return jsonify({
            'success': True,
            'message': '扫描成功',
            'output_path': scan_output_path_normalized,
            'filename': os.path.basename(scan_output_path_normalized)
        })
        
    except subprocess.TimeoutExpired:
        return jsonify({'error': '扫描超时，请检查扫描仪连接或重试'}), 500
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        print(f"扫描失败: {e}\n{error_trace}")
        return jsonify({'error': f'扫描失败: {str(e)}'}), 500

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

def extract_student_id(text):
    """从OCR文本中提取学号（6位数字），无需前缀"""
    if not text:
        return None
    
    # 学号格式：6位数字，无需前缀
    # 优先匹配文本开头的6位数字（最可能的位置）
    # 也支持带标识的格式（兼容性）
    patterns = [
        r'^[^\d]*(\d{6})',  # 文本开头附近的6位数字（忽略前面的非数字字符）
        r'学号[：:]\s*(\d{6})',  # 学号：123456（兼容格式）
        r'Student\s*ID[：:]\s*(\d{6})',  # Student ID: 123456（兼容格式）
        r'ID[：:]\s*(\d{6})',  # ID: 123456（兼容格式）
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            student_id = match.group(1).strip()
            # 验证是否为6位数字
            if len(student_id) == 6 and student_id.isdigit():
                return student_id
    
    # 如果没找到，尝试提取文本前几行的6位数字（独立数字，不依赖前缀）
    lines = text.split('\n')[:5]  # 只看前5行
    for line in lines:
        # 查找独立的6位数字（前后是空格、换行或标点）
        # 优先匹配行首的6位数字
        line_stripped = line.strip()
        if line_stripped:
            # 检查是否是行首的6位数字
            if re.match(r'^\d{6}(?:\s|$|[^\d])', line_stripped):
                match = re.search(r'^(\d{6})', line_stripped)
                if match:
                    return match.group(1)
            # 查找行中独立的6位数字
            numbers = re.findall(r'(?<!\d)\d{6}(?!\d)', line)  # 前后都不是数字的6位数字
            if numbers:
                return numbers[0]
    
    return None

def translate_essay_to_chinese(text, logger=None):
    """将识别出的英文作文译为中文，供识别结果旁展示译文。调用模型时在识别结果旁输出译文。"""
    if not text or not text.strip() or not minimax_api_key:
        return ''
    try:
        client = OpenAI(base_url=minimax_base_url, api_key=minimax_api_key)
        response = client.chat.completions.create(
            model=minimax_model,
            messages=[
                {"role": "system", "content": "你是翻译助手。将用户给出的英文作文或文本翻译成中文，保持段落与语气，不要批注或点评，只输出译文。"},
                {"role": "user", "content": text.strip()}
            ],
            temperature=0.3
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
                
                # OCR 识别：使用讯飞OCR
                result = None
                text_from_result = ''

                # 使用讯飞OCR进行识别
                try:
                    logger.info(f"使用讯飞OCR识别文件: {filename}")
                    # 将PIL Image转换为字节数据
                    img_io = io.BytesIO()
                    image.save(img_io, format='PNG')
                    img_bytes = img_io.getvalue()
                    
                    text_from_result = xunfei_ocr_recognize(img_bytes, logger)
                    if text_from_result and text_from_result.strip():
                        result = {'text': text_from_result}
                        logger.info(f"讯飞OCR 识别成功，文本长度: {len(text_from_result)}")
                    else:
                        logger.warning("讯飞OCR 返回空文本")
                except Exception as xf_error:
                    logger.error(f"讯飞OCR 识别失败: {xf_error}")
                    text_from_result = ''
                    result = None

                # 处理讯飞OCR返回的文本
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
                            temperature=0.7
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
                    model=minimax_model,
                    messages=[
                        {"role": "system", "content": "你是一位经验丰富的教师，擅长分析班级整体学习情况，并能从作文中提炼好词好句。"},
                        {"role": "user", "content": summary_prompt}
                    ],
                    temperature=0.7
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
            model=minimax_model,
            messages=[
                {"role": "system", "content": "你是一位经验丰富的教师，擅长分析班级整体学习情况，并能从作文中提炼好词好句。"},
                {"role": "user", "content": summary_prompt}
            ],
            temperature=0.7
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
                    temperature=0.7
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
                    model=minimax_model,
                    messages=[
                        {"role": "system", "content": "你是一位经验丰富的教师，擅长分析班级整体学习情况。"},
                        {"role": "user", "content": summary_prompt}
                    ],
                    temperature=0.7
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
        client = OpenAI(base_url=minimax_base_url, api_key=minimax_api_key, timeout=30.0)
        response = client.chat.completions.create(
            model=minimax_model,
            messages=[
                {"role": "user", "content": prompt + "\n\n" + full_text[:2500]}
            ],
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

if __name__ == '__main__':
    # 初始化NAPS2 PowerShell函数（可选，主要用于测试）
    print("正在初始化NAPS2 PowerShell函数...")
    initialize_naps2_powershell_function()
    print("启动Flask应用...")
    app.run(host='0.0.0.0', port=5001, debug=True)

