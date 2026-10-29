#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把 .env 里的 API token 搬进加密保险箱（Windows DPAPI）。

用法（在项目根目录执行，用项目自带的 python）：

    python tools/seal_secrets.py                  # 体检：加密了哪些、能不能解开
    python tools/seal_secrets.py --apply          # 把 .env 里的明文密钥搬进保险箱
    python tools/seal_secrets.py --set LLM_API_KEY        # 换 token（不回显、不进历史）
    python tools/seal_secrets.py --unseal LLM_API_KEY     # 临时看明文（换机器前导出用）
    python tools/seal_secrets.py --verify         # 密文解出来的值 vs .env 明文，比对指纹
    python tools/seal_secrets.py --list-plaintext # 找出磁盘上还有明文密钥的文件

设计要点见 secretbox.py 顶部注释。这个工具不做任何"自动降级成明文"的事：
解不开就报错，让人去处理，而不是悄悄把明文写回 .env。
"""

import argparse
import getpass
import glob
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import secretbox  # noqa: E402

ROOT = secretbox.project_root()
ENV_PATH = os.path.join(ROOT, '.env')
ENV_LINE = re.compile(r'^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$')


def _read_env_lines(path=ENV_PATH):
    with open(path, 'r', encoding='utf-8', errors='replace') as f:
        return f.read().splitlines()


def _parse(lines, keys=None):
    """取出 KEY=value（去引号）。keys 为空表示全部。"""
    out = {}
    for line in lines:
        if line.strip().startswith('#'):
            continue
        m = ENV_LINE.match(line)
        if not m:
            continue
        k, v = m.group(1), m.group(2).strip()
        if v[:1] in ('"', "'") and v[-1:] == v[:1] and len(v) >= 2:
            v = v[1:-1]
        if keys and k not in keys:
            continue
        out[k] = v
    return out


def _mask(v):
    return '空' if not v else '%d 字符 · 指纹 %s' % (len(v), secretbox.fingerprint(v))


def cmd_status(_args):
    st = secretbox.status()
    print('保险箱文件：%s' % st['path'])
    print('  存在：%s    DPAPI 可用：%s' % ('是' if st['exists'] else '否',
                                        '是' if st['dpapi'] else '否'))
    if '__store__' in st['errors']:
        print('  ✗ 读不出来：%s' % st['errors']['__store__'])
    if st['items']:
        print('  已加密：')
        for name, it in sorted(st['items'].items()):
            if it.get('ok'):
                print('    ✓ %-18s %d 字符 · 指纹 %s' % (name, it.get('length', 0),
                                                      it.get('fingerprint')))
            else:
                print('    ✗ %-18s 解不开：%s' % (name, it.get('error')))
    else:
        print('  已加密：（无）')

    plain = _parse(_read_env_lines(), secretbox.SEALED_KEYS) if os.path.exists(ENV_PATH) else {}
    plain = {k: v for k, v in plain.items() if v}
    if plain:
        print('.env 里的明文密钥（建议 --apply 收进保险箱）：')
        for k, v in sorted(plain.items()):
            print('    ! %-18s %s' % (k, _mask(v)))
    else:
        print('.env 里没有明文密钥了。')
    return 0


def cmd_apply(_args):
    if not os.path.exists(ENV_PATH):
        print('✗ 找不到 %s' % ENV_PATH)
        return 2
    lines = _read_env_lines()
    plain = {k: v for k, v in _parse(lines, secretbox.SEALED_KEYS).items() if v}
    if not plain:
        print('没有需要迁移的明文密钥。')
        return 0

    for name, value in plain.items():
        secretbox.set_secret(name, value)
        print('  ✓ 已加密 %-18s %s' % (name, _mask(value)))
        # 立刻解一遍，确认存进去的东西真的能解开（宁可现在失败，也别等重启后才发现）
        back = secretbox.get_secret(name)
        if secretbox.fingerprint(back) != secretbox.fingerprint(value):
            print('  ✗ %s 回读校验失败，已中止，.env 未改动' % name)
            return 3

    # 逐行改：把这些键换成一行注释，其余内容（含注释、顺序）原样保留
    out, done = [], set()
    for line in lines:
        m = ENV_LINE.match(line)
        if m and m.group(1) in plain and not m.group(1) in done:
            k = m.group(1)
            out.append('# %s 已加密存入 data/config/secrets.dpapi.json'
                       '（明文不再落盘；查看/更换：python tools/seal_secrets.py --status）' % k)
            done.add(k)
        else:
            out.append(line)

    # 先写临时文件再替换；不生成 .env.bak（那等于又留一份明文）
    tmp = ENV_PATH + '.tmp'
    with open(tmp, 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(out) + '\n')
    check = _parse(out)
    for name in plain:
        if check.get(name):
            print('  ✗ 改写后 %s 仍是明文，已中止' % name)
            os.remove(tmp)
            return 3
    os.replace(tmp, ENV_PATH)
    print('\n✓ .env 已改写：%d 个明文密钥换成注释。其余配置项未动。' % len(done))
    print('  下一步：重启服务（python app.py 或 启动服务.bat）让进程用上密文。')
    _warn_backups()
    return 0


def cmd_set(args):
    name = args.set
    if not name:
        print('✗ 用法：--set LLM_API_KEY')
        return 2
    print('输入 %s 的新值（不回显，不会进命令历史；直接回车＝删除该项）：' % name)
    value = getpass.getpass('> ').strip()
    if not value:
        secretbox.set_secret(name, '')
        print('✓ 已删除 %s 的密文' % name)
        return 0
    secretbox.set_secret(name, value)
    back = secretbox.get_secret(name)
    if secretbox.fingerprint(back) != secretbox.fingerprint(value):
        print('✗ 回读校验失败')
        return 3
    print('✓ 已加密写入 %s（%s）' % (name, _mask(value)))
    # 顺手把 .env 里的同名明文删掉，否则等于没换
    if os.path.exists(ENV_PATH):
        lines = _read_env_lines()
        if {k: v for k, v in _parse(lines, [name]).items() if v}:
            print('  ! .env 里还有同名明文，运行 --apply 清掉')
    return 0


def cmd_unseal(args):
    value = secretbox.get_secret(args.unseal)
    if not value:
        print('✗ 保险箱里没有 %s，也没有对应的环境变量' % args.unseal)
        return 2
    print('%s = %s' % (args.unseal, value))
    print('\n（这一行是明文，用完请清屏；换机器时把它填进新机器的 '
          'python tools/seal_secrets.py --set %s 即可）' % args.unseal)
    return 0


def cmd_verify(_args):
    ok = True
    st = secretbox.status()
    for name, it in (st['items'] or {}).items():
        if it.get('ok'):
            print('  ✓ %-18s 解开了，%d 字符 · 指纹 %s' % (name, it['length'], it['fingerprint']))
        else:
            ok = False
            print('  ✗ %-18s %s' % (name, it.get('error')))
    if os.path.exists(ENV_PATH):
        plain = {k: v for k, v in _parse(_read_env_lines(), secretbox.SEALED_KEYS).items() if v}
        for name, value in plain.items():
            sealed = (st['items'] or {}).get(name)
            if sealed and sealed.get('ok') and sealed['fingerprint'] != secretbox.fingerprint(value):
                ok = False
                print('  ✗ %-18s .env 明文与密文不一致（哪份才是当前在用的？）' % name)
    print('\n结论：%s' % ('全部正常' if ok else '有问题，见上面 ✗'))
    return 0 if ok else 1


def _plaintext_files():
    """找出磁盘上还存着明文密钥的文件（.env 与各种备份副本）。"""
    hits = []
    pats = ['.env', '.env.*', '*.env']
    seen = set()
    for pat in pats:
        for p in glob.glob(os.path.join(ROOT, pat)):
            if not os.path.isfile(p) or p in seen:
                continue
            seen.add(p)
            try:
                lines = _read_env_lines(p)
            except OSError:
                continue
            found = {k: v for k, v in _parse(lines, secretbox.SEALED_KEYS).items() if v}
            if found:
                hits.append((p, found))
    return hits


def _warn_backups():
    hits = [(p, f) for p, f in _plaintext_files() if os.path.abspath(p) != os.path.abspath(ENV_PATH)]
    if not hits:
        return
    print('\n⚠ 以下副本仍然含明文密钥，交接前请处理掉：')
    for p, found in hits:
        print('    %s（%s）' % (os.path.relpath(p, ROOT), '、'.join(sorted(found))))


def cmd_list_plaintext(_args):
    hits = _plaintext_files()
    if not hits:
        print('磁盘上没有找到含明文密钥的 .env 文件。')
        return 0
    print('含明文密钥的文件：')
    for p, found in hits:
        print('  %s' % os.path.relpath(p, ROOT))
        for k, v in sorted(found.items()):
            print('      %-18s %s' % (k, _mask(v)))
    print('\n建议：1) --apply 收进保险箱；2) 到服务商后台把旧 token 吊销重发；'
          '\n      3) 删掉这些 .env.bak 副本（吊销之后它们就是废纸）。')
    return 0


def cmd_shred_backups(args):
    """删掉含明文密钥的副本（不含 .env 本身）。

    只删副本：.env 是活的配置，删了服务就起不来了。删之前先把里面的密钥也收进保险箱，
    万一以后要找回旧 token 还有据可查——但请记住：真正让旧 token 失效的只有服务商后台吊销。
    """
    hits = [(p, f) for p, f in _plaintext_files()
            if os.path.abspath(p) != os.path.abspath(ENV_PATH)]
    if not hits:
        print('没有含明文密钥的副本。')
        return 0
    print('将要删除这些副本（.env 本身不动）：')
    for p, found in hits:
        print('  %s（%s）' % (os.path.relpath(p, ROOT), '、'.join(sorted(found))))
    if not args.yes:
        print('\n这是预演。确认无误后加 --yes 真正删除：'
              '\n  python tools/seal_secrets.py --shred-backups --yes'
              '\n（建议先在服务商后台吊销这几个旧 token，删文件并不等于吊销）')
        return 0
    for p, found in hits:
        for name, value in found.items():
            try:
                secretbox.set_secret(name + '_OLD', value)   # 以 _OLD 后缀存档，不进启动读取路径
            except Exception as e:
                print('  ! %s 归档失败：%s' % (name, e))
        try:
            os.remove(p)
            print('  ✓ 已删除 %s' % os.path.relpath(p, ROOT))
        except OSError as e:
            print('  ✗ 删除 %s 失败：%s' % (p, e))
    print('\n旧值已以 <名字>_OLD 存进保险箱（--unseal LLM_API_KEY_OLD 可取），'
          '\n但请到服务商后台吊销旧 token —— 文件删了不等于 token 失效。')
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description='QuickJudge 密钥保险箱工具（Windows DPAPI）')
    ap.add_argument('--status', action='store_true', help='体检（默认动作）')
    ap.add_argument('--apply', action='store_true', help='把 .env 明文密钥搬进保险箱')
    ap.add_argument('--set', metavar='NAME', help='录入/更换一项密钥（交互式，不回显）')
    ap.add_argument('--unseal', metavar='NAME', help='打印一项明文（迁移用，慎用）')
    ap.add_argument('--verify', action='store_true', help='校验密文能否解开且与明文一致')
    ap.add_argument('--list-plaintext', action='store_true', help='列出磁盘上仍含明文密钥的文件')
    ap.add_argument('--shred-backups', action='store_true', help='删除含明文密钥的 .env 副本（先预演）')
    ap.add_argument('--yes', action='store_true', help='配合 --shred-backups：真正执行删除')
    a = ap.parse_args(argv)

    if not secretbox.dpapi_available():
        print('✗ 当前系统没有 Windows DPAPI，无法加密。请在 Windows 上执行本工具。')
        return 4

    if a.apply:
        return cmd_apply(a)
    if a.set:
        return cmd_set(a)
    if a.unseal:
        return cmd_unseal(a)
    if a.verify:
        return cmd_verify(a)
    if a.list_plaintext:
        return cmd_list_plaintext(a)
    if a.shred_backups:
        return cmd_shred_backups(a)
    return cmd_status(a)


if __name__ == '__main__':
    sys.exit(main())
