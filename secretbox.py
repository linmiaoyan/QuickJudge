"""本地密钥保险箱：用 Windows DPAPI 把 API token 加密成「只有本机本账户能解开」的密文。

为什么要有这个模块
------------------
`.env` 里的 `LLM_API_KEY` / `FLASK_SECRET_KEY` 是明文。任何拿到这台电脑的人——包括以后
来接手设备的第三方厂家——只要打开项目目录就能读走学校的模型额度；拿到 `FLASK_SECRET_KEY`
甚至能自己伪造管理员 Cookie。这是实打实的泄漏面，不是理论风险。

这里用 Windows 自带的 DPAPI（`CryptProtectData`）把这两类值加密后集中存到
`data/config/secrets.dpapi.json`（`data/` 已在 .gitignore 里）：

  · 加密密钥不在项目里，而是当前 Windows 用户在系统里的主密钥（LSA 保管，由账户口令保护）。
    项目目录被整个拷走，也只是拿到一堆读不了的密文；
  · 密文绑定 **本机 + 本账户 + 一段应用熵**（`APP_ENTROPY`）。换机器、换账户登录，
    `CryptUnprotectData` 直接失败——这正是「设备被别的厂家接管」时要的效果；
  · 服务照常无人值守启动：解密不需要人输口令，启服务的账户自己就能解开。

边界（必须如实告知使用方）
--------------------------
1. **DPAPI 保护的是「账户」，不是「人」。** 同一个 Windows 账户下的任何进程都能解开。
   所以交接设备时如果连账户一起交出去，这一层就没了。正确做法见 §迁移；
2. 进程内存里密钥仍是明文（调大模型必须如此）。能读进程内存的人本来也就能读密钥；
3. 非 Windows 平台没有 DPAPI。本模块不会"降级成明文偷偷省事"，而是明确报错，
   让 `tools/seal_secrets.py --status` 一眼看出没加密。

迁移 / 换机器
-------------
`secrets.dpapi.json` 拷到别的机器（或换 Windows 账户后）是**解不开**的，这是设计而非故障。
换手时二选一：
  · 新账户登录后重新灌一次：`python tools/seal_secrets.py --set LLM_API_KEY`
    （会让服务商后台的 token 重填一遍，旧密文随即作废）；
  · 干脆删掉 `secrets.dpapi.json`，让新接手方用他们自己的 token。
"""

import base64
import ctypes
import json
import os
import sys

try:                      # 只在 Windows 上存在
    from ctypes import wintypes
except Exception:         # pragma: no cover
    wintypes = None

VERSION = 1
BACKEND = 'dpapi-user'
STORE_NAME = 'secrets.dpapi.json'
# 应用熵：把密文再绑到这个应用上，别的程序即使拿到用户主密钥也不能直接解开这段 blob
APP_ENTROPY = b'QuickJudge/secrets/v1'
# 需要保护的键。其他键（模型名、base_url、并发数）不是秘密，留在 .env 里方便人改。
SEALED_KEYS = ('LLM_API_KEY', 'MINIMAX_API_KEY', 'FLASK_SECRET_KEY')

_CRYPTPROTECT_UI_FORBIDDEN = 0x01

# 启动过程中攒下的告警，统一由 startup_report() 打出来
_WARNINGS = []
# 每个键的解密失败原因（供 status / 诊断使用）
_ERRORS = {}
# 记录哪些键是从明文 .env 读到的（说明还没迁移）
_FROM_PLAINTEXT = set()


class SecretError(Exception):
    """保险箱打不开（多半是换了 Windows 账户 / 换了机器 / 文件损坏）。"""


# --------------------------------------------------------------------------- 路径

def project_root():
    return os.path.dirname(os.path.abspath(__file__))


def store_path():
    """密文位置。允许用 QJ_SECRETS_PATH 指到别处（比如放到受控目录）。"""
    custom = os.environ.get('QJ_SECRETS_PATH', '').strip()
    if custom:
        return os.path.abspath(custom)
    return os.path.join(project_root(), 'data', 'config', STORE_NAME)


# --------------------------------------------------------------------------- DPAPI

if wintypes is not None and sys.platform == 'win32':
    class _BLOB(ctypes.Structure):
        _fields_ = [('cbData', wintypes.DWORD),
                    ('pbData', ctypes.POINTER(ctypes.c_char))]

    _crypt32 = ctypes.WinDLL('crypt32', use_last_error=True)
    _kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)

    _crypt32.CryptProtectData.argtypes = [
        ctypes.POINTER(_BLOB), wintypes.LPCWSTR, ctypes.POINTER(_BLOB),
        ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(_BLOB)]
    _crypt32.CryptProtectData.restype = wintypes.BOOL
    _crypt32.CryptUnprotectData.argtypes = [
        ctypes.POINTER(_BLOB), ctypes.POINTER(wintypes.LPWSTR), ctypes.POINTER(_BLOB),
        ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(_BLOB)]
    _crypt32.CryptUnprotectData.restype = wintypes.BOOL
    _kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    _kernel32.LocalFree.restype = ctypes.c_void_p


def dpapi_available():
    return sys.platform == 'win32' and wintypes is not None


def _dpapi(protect, data, entropy=APP_ENTROPY):
    """跑一次 CryptProtectData / CryptUnprotectData，返回处理后的字节。"""
    if not dpapi_available():
        raise SecretError('当前平台不是 Windows，没有 DPAPI 可用')
    in_buf = ctypes.create_string_buffer(data, len(data))
    ent_buf = ctypes.create_string_buffer(entropy, len(entropy))
    blob_in = _BLOB(len(data), ctypes.cast(in_buf, ctypes.POINTER(ctypes.c_char)))
    blob_ent = _BLOB(len(entropy), ctypes.cast(ent_buf, ctypes.POINTER(ctypes.c_char)))
    blob_out = _BLOB()
    if protect:
        ok = _crypt32.CryptProtectData(ctypes.byref(blob_in), 'QuickJudge secret',
                                       ctypes.byref(blob_ent), None, None,
                                       _CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(blob_out))
    else:
        ok = _crypt32.CryptUnprotectData(ctypes.byref(blob_in), None,
                                         ctypes.byref(blob_ent), None, None,
                                         _CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(blob_out))
    if not ok:
        err = ctypes.get_last_error()
        raise SecretError('DPAPI %s失败（GetLastError=%s）' % ('加密' if protect else '解密', err))
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        _kernel32.LocalFree(ctypes.cast(blob_out.pbData, ctypes.c_void_p))


def encrypt_text(value):
    """明文 → 'dpapi-user:<base64>'。空值不加密。"""
    if not value:
        raise SecretError('空值不需要加密')
    return '%s:%s' % (BACKEND, base64.b64encode(_dpapi(True, value.encode('utf-8'))).decode('ascii'))


def decrypt_text(blob):
    """'dpapi-user:<base64>' → 明文。解不开抛 SecretError（原因分得清才好排查）。"""
    blob = str(blob or '')
    if not blob.startswith(BACKEND + ':'):
        raise SecretError('密文格式不认识（应以 %s: 开头）' % BACKEND)
    try:
        raw = base64.b64decode(blob.split(':', 1)[1])
    except Exception as e:
        raise SecretError('密文不是合法 base64：%s' % e)
    try:
        return _dpapi(False, raw).decode('utf-8')
    except SecretError as e:
        raise SecretError('%s —— 通常是换了 Windows 账户或换了机器（密文绑定本机本账户）' % e)


# --------------------------------------------------------------------------- 保险箱读写

def load_store(path=None):
    """读密文文件；不存在返回空壳，损坏则直接报错（不能当成"没有密钥"静默继续）。"""
    p = path or store_path()
    if not os.path.exists(p):
        return {'v': VERSION, 'backend': BACKEND, 'items': {}}
    try:
        with open(p, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        raise SecretError('密文文件损坏，读不出来：%s（%s）' % (p, e))
    if not isinstance(data, dict) or not isinstance(data.get('items'), dict):
        raise SecretError('密文文件结构不对：%s' % p)
    return data


def save_store(store, path=None):
    """原子写回，并把文件权限收到当前用户（Windows 上尽量窄，失败不致命）。"""
    p = path or store_path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    store = dict(store or {})
    store['v'] = VERSION
    store['backend'] = BACKEND
    tmp = p + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(store, f, ensure_ascii=False, indent=2)
    os.replace(tmp, p)
    _tighten_acl(p)
    return p


def _tighten_acl(path):
    """只把读权限留给当前用户与 SYSTEM/管理员组。

    密文本身没有主密钥就是废纸，这一步是"少一层顺手可读"的加固，不是安全边界；
    失败（比如没有权限改 ACL）只提示，不影响功能。
    """
    if sys.platform != 'win32':
        try:
            os.chmod(path, 0o600)
        except OSError:
            pass
        return
    try:
        import subprocess
        user = os.environ.get('USERNAME') or ''
        if not user:
            return
        subprocess.run(['icacls', path, '/inheritance:r',
                        '/grant:r', '%s:(R,W)' % user,
                        '/grant:r', 'SYSTEM:(F)',
                        '/grant:r', 'Administrators:(F)'],
                       capture_output=True, timeout=15)
    except Exception:
        pass


def set_secret(name, value, path=None):
    """写入/更新一项密文。value 为空表示删除该项。"""
    store = load_store(path)
    items = store.setdefault('items', {})
    if value:
        items[name] = encrypt_text(value)
    else:
        items.pop(name, None)
    save_store(store, path)
    return name in items


def has_secret(name, path=None):
    try:
        return name in (load_store(path).get('items') or {})
    except SecretError:
        return False


def get_secret(name, default=''):
    """按「密文 → 明文环境变量 → default」取值。这是 app.py 唯一该用的入口。

    密文优先：即便 .env 里还残留同名明文（迁移没删干净），也以保险箱为准，
    避免"以为加密了，其实一直读的是明文"。
    """
    try:
        items = load_store().get('items') or {}
        if name in items:
            try:
                val = decrypt_text(items[name]).strip()
                if val:
                    return val
            except SecretError as e:
                _ERRORS[name] = str(e)
                _WARNINGS.append('%s 的密文解不开：%s' % (name, e))
    except SecretError as e:
        _ERRORS[name] = str(e)
        _WARNINGS.append(str(e))
    plain = (os.environ.get(name) or '').strip()
    if plain:
        _FROM_PLAINTEXT.add(name)
        return plain
    return default


def status(path=None):
    """体检信息：不返回任何明文，只回答"加没加密、解不解得开"。"""
    p = path or store_path()
    out = {'path': p, 'exists': os.path.exists(p), 'backend': BACKEND,
           'dpapi': dpapi_available(), 'items': {}, 'errors': {}}
    try:
        store = load_store(p)
    except SecretError as e:
        out['errors']['__store__'] = str(e)
        return out
    for name, blob in (store.get('items') or {}).items():
        item = {'sealed': True, 'bytes': len(str(blob))}
        try:
            val = decrypt_text(blob)
            item['ok'] = True
            item['length'] = len(val)
            item['fingerprint'] = fingerprint(val)
        except SecretError as e:
            item['ok'] = False
            item['error'] = str(e)
        out['items'][name] = item
    return out


def fingerprint(value):
    """指纹：sha256 前 8 位。用来比对"解出来的还是不是原来那个 token"，又不泄漏明文。"""
    import hashlib
    return hashlib.sha256((value or '').encode('utf-8')).hexdigest()[:8]


def startup_report():
    """启动时打一行体检结论。没配加密也别吵，但要说清楚。"""
    lines = []
    st = status()
    sealed_ok = [k for k, v in st['items'].items() if v.get('ok')]
    sealed_bad = [k for k, v in st['items'].items() if not v.get('ok')]
    if sealed_ok:
        lines.append('密钥保险箱：已加密 %d 项（%s）' % (len(sealed_ok), '、'.join(sealed_ok)))
    if sealed_bad:
        lines.append('⚠ 密钥保险箱：%s 解密失败 —— %s' % (
            '、'.join(sealed_bad), st['items'][sealed_bad[0]].get('error', '')))
    if _FROM_PLAINTEXT:
        lines.append('⚠ 仍在用明文 .env：%s，建议运行 python tools/seal_secrets.py --apply'
                     % '、'.join(sorted(_FROM_PLAINTEXT)))
    bad_prefix = tuple(k + ' 的密文解不开' for k in sealed_bad)
    for w in _WARNINGS:
        # 上面那行已经说过"哪一项解不开"，别同一件事再说一遍
        if w in lines or w.startswith(bad_prefix):
            continue
        lines.append('⚠ %s' % w)
    if not st['exists'] and not _FROM_PLAINTEXT:
        lines.append('密钥保险箱：未启用（只用环境变量）')
    return '\n'.join('  ' + x for x in lines)
