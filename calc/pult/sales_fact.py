# -*- coding: utf-8 -*-
"""Факт продаж для раздела «Цели и планы → Планы продаж» (только чтение из учётной системы). Печатает JSON.
Правило (Азиз, 26.11.2026), всё в MRR, по продуктам (поле «Тип подписки»: SFA, POS, AI ассистент, Perfect store):
  продажи месяца = новый MRR подписок, созданных в этом месяце (событие «создание» журнала MRR)
                 + расширения − сокращения тех подписок, которым на дату изменения меньше 6 месяцев с создания.
Отток новой подписки в первые 6 месяцев — справочно, из продаж не вычитается. Январь 2026 — перенос базы в учётную систему, не продажи.
USD — по 12 000 сум/$."""
import json, sys, os, datetime, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__))); sys.path.insert(1, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from odoo_db import OdooDB
from cls import PROD_SQL, AMT_SQL
START = datetime.date(2026, 2, 1); WIN = 6
def addm(d, n):
    y, m = divmod(d.month - 1 + n, 12); y += d.year; m += 1
    import calendar
    return d.replace(year=y, month=m, day=min(d.day, calendar.monthrange(y, m)[1]))
db = OdooDB(); q = db.query
rows = q(f"""select l.order_id o, so.name, coalesce(rp.name, '') client, {PROD_SQL} prod, l.event_type et, l.event_date d, {AMT_SQL} a,
   coalesce(up.name, '') usr
 from sale_order_log l join sale_order so on so.id=l.order_id left join res_partner rp on rp.id=so.partner_id
 left join res_users u on u.id=coalesce(so.user_id, l.user_id) left join res_partner up on up.id=u.partner_id
 where l.event_type in ('0_creation','1_expansion','15_contraction','2_churn') order by l.event_date, l.id""")
db.close()
born = {}
for r in rows:
    if r['et'] == '0_creation': born.setdefault(r['o'], r['d'])
PN = {'SFA': 'SFA', 'POS': 'POS', 'AI': 'AI ассистент', 'PS': 'Perfect store'}
fact = collections.defaultdict(lambda: collections.defaultdict(lambda: {'new': 0.0, 'exp': 0.0, 'con': 0.0, 'churn': 0.0, 'n_new': 0}))
det = collections.defaultdict(list)
KIND = {'0_creation': 'new', '1_expansion': 'exp', '15_contraction': 'con', '2_churn': 'churn'}
for r in rows:
    p = r['prod']
    if p not in PN: continue
    b = born.get(r['o']); d = r['d']
    if not b or b < START or d < START: continue           # подписки, созданные до февраля 2026 (база), — не новые продажи
    if r['et'] != '0_creation' and d >= addm(b, WIN): continue   # изменения после 6 месяцев с создания — уже не продажа
    k = KIND[r['et']]; ym = d.strftime('%Y-%m'); a = float(r['a'] or 0)
    f = fact[ym][p]; f[k] += a
    if k == 'new': f['n_new'] += 1
    det[ym].append({'name': r['name'], 'client': r['client'], 'prod': p, 'kind': k, 'a': round(a, 2), 'date': str(d), 'born': str(b), 'user': r['usr']})
out = {'computed': datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'), 'start': str(START), 'window': WIN,
       'fact': {ym: {p: dict(v, total=v['new'] + v['exp'] + v['con']) for p, v in ps.items()} for ym, ps in sorted(fact.items())},
       'details': {ym: sorted(v, key=lambda x: (x['prod'], x['kind'], -abs(x['a']))) for ym, v in det.items()}}
print(json.dumps(out, ensure_ascii=False))
