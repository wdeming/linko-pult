# -*- coding: utf-8 -*-
"""Расходы из реестра Odoo «Расходы» (x_custom_expenses + строки): дата оплаты, период, к которому относится расход, вид и статья.
Сумма = сумма строк; если в расходе указан курс (оплата в долларах) — пересчёт по 12 000 сум/$, как во всём комплекте."""
import json, re
from odoo_db import OdooDB
from cfg_cut import *
db = OdooDB(); K = 12000
rows = db.query("""
select h.id, h.x_studio_date pay, h.x_studio_start_date d1, h.x_studio_end_date d2, coalesce(h.x_studio_rate, 0) rate,
  coalesce(ap.name->>'ru_RU', ap.name->>'en_US', 'Без вида') plan, coalesce(aa.name->>'ru_RU', aa.name->>'en_US', 'Без статьи') cat,
  coalesce(aj.name->>'ru_RU', aj.name->>'en_US', '') kassa, h.x_studio_notes nt,
  (select coalesce(sum(l.x_studio_sum), 0) from x_custom_expenses_line_dc8d4 l where l.x_custom_expenses_id = h.id) lsum
from x_custom_expenses h
left join account_analytic_plan ap on ap.id = h.x_studio_accountanalyticplan
left join account_analytic_account aa on aa.id = h.x_studio_accountanalyticaccount
left join account_journal aj on aj.id = h.x_studio_journal
where h.x_active is not false and h.x_studio_date is not null""")
clean = lambda s: re.sub(r'\s+', ' ', re.sub(r'<[^>]+>|&nbsp;', ' ', s or '')).strip()
out = []
for r in rows:
    fx = r['rate'] > 1
    d1, d2 = r['d1'] or r['pay'], r['d2'] or r['d1'] or r['pay']
    if d2 < d1: d1, d2 = d2, d1
    out.append({'id': r['id'], 'pay': str(r['pay']), 'd1': str(d1), 'd2': str(d2), 'plan': r['plan'], 'cat': r['cat'], 'kassa': r['kassa'],
                'usd': float(r['lsum']) if fx else 0.0, 'rate': float(r['rate']), 'amt': float(r['lsum']) * (K if fx else 1),
                'amt_odoo': float(r['lsum']) * (r['rate'] if r['rate'] > 0 else 1), 'note': clean(r['nt'])[:200]})
json.dump(out, open(J('exp'), 'w'), ensure_ascii=False)
print('расходов %d, сумма %.1f млн, из них в долларах $%.0f' % (len(out), sum(x['amt'] for x in out) / 1e6, sum(x['usd'] for x in out)))
