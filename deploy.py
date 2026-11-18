# -*- coding: utf-8 -*-
"""Выкладка пульта на linko.leepr.uz под пользователем pult (без root).
python3 pult_server/deploy.py "что изменено"   — заберёт правки с GitHub, отправит изменённые файлы, зафиксирует их в git, отправит на GitHub и перезапустит сервис."""
import os, sys, hashlib, paramiko
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__))); R = '/opt/linko-pult'
MAP = [('pult_server/app.py', 'app.py'), ('pult_server/backup.py', 'backup.py'), ('pult_server/deploy.py', 'deploy.py')] + [('pult_server/web/' + f, 'web/' + f) for f in ('index.html', 'style.css', 'app.js', 'shim.js', 'boot.js')]
MAP += [(f, 'calc/' + f) for f in ('cfg_cut.py', 'cls.py', 'style_lib.py', 'extract_oct31.py', 'extract_oct31_mkt.py', 'extract_oct31_churn.py', 'extract_oct31_calls.py',
        'gen_oct31_clients.py', 'gen_oct31_multi.py', 'gen_oct31_top20.py', 'gen_oct31_cost.py', 'gen_oct31_metrics.py', 'gen_oct31_mrr.py', 'gen_oct31_churn.py', 'gen_oct31_calls.py', 'extract_revenue_cash.py', 'gen_revenue_cash.py')]
MAP += [('pult/' + f, 'calc/pult/' + f) for f in ('refresh.py', 'build_data.py', 'extract_staff.py', 'odoo_db.py')]
msg = sys.argv[1] if len(sys.argv) > 1 else 'Обновление пульта'
c = paramiko.SSHClient(); c.set_missing_host_key_policy(paramiko.AutoAddPolicy())
c.connect('89.31.29.22', username='pult', key_filename=os.path.join(ROOT, 'pult', 'deploy_key'), timeout=20, look_for_keys=False, allow_agent=False)
def sh(cmd):
    i, o, e = c.exec_command(cmd, timeout=120); out = o.read().decode(); err = e.read().decode()
    return out + (('\n' + err) if err.strip() else '')
def sums():
    return dict(l.split('  ', 1)[::-1] for l in sh('cd %s && sha256sum %s 2>/dev/null' % (R, ' '.join(r for _, r in MAP))).split('\n') if '  ' in l)
lsum = lambda p: hashlib.sha256(open(p, 'rb').read()).hexdigest() if os.path.exists(p) else None
before = sums()
print(sh('cd %s && git pull -q --ff-only origin main 2>&1 | tail -3' % R).strip() or 'GitHub: новых правок нет')
after = sums(); s = c.open_sftp(); conflicts, pulled, sent = [], [], []
for loc, rem in MAP:
    p = os.path.join(ROOT, loc)
    if before.get(rem) != after.get(rem):            # файл поменяли на GitHub
        if lsum(p) in (before.get(rem), after.get(rem)): s.get('%s/%s' % (R, rem), p); pulled.append(rem)
        else: conflicts.append(rem)
if conflicts:
    s.close(); c.close()
    raise SystemExit('СТОП: эти файлы изменены и на GitHub, и в папке проекта — сначала свести правки: %s' % conflicts)
for loc, rem in MAP:
    p = os.path.join(ROOT, loc)
    if os.path.exists(p) and lsum(p) != after.get(rem):
        s.put(p, '%s/%s' % (R, rem)); sent.append(rem)
s.close()
if pulled: print('забрано с GitHub в папку проекта:', pulled)
print('отправлено:', sent or 'ничего не изменилось')
if sent:
    trailer = '\n\nCo-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>\nClaude-Session: https://claude.ai/code/session_01VdFm2FEr6VgqW7aE1pqVUg'
    print(sh("cd %s && git add -A && git -c user.name=Claude -c user.email=pult@linko.local commit -q -F - <<'M'\n%s%s\nM\ngit log --oneline | head -1" % (R, msg, trailer)))
print(sh('cd %s && git push -q origin main 2>&1; git status -sb | head -1' % R).strip())
if 'app.py' in sent + pulled: print(sh('sudo /usr/bin/systemctl restart linko-pult && sleep 2 && curl -s -o /dev/null -w "сервис %{http_code}\\n" http://127.0.0.1:8082/healthz'))
c.close()
