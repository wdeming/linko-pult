# -*- coding: utf-8 -*-
"""Поступления денег (account.payment, проведённые) с 01.02.2026 по срез: по месяцам, продукту сервера, способу оплаты.
Платежи в валюте пересчитаны по курсу 12 000 сум/$; комиссии Payme/Click/Paynet — исходящие платежи клиентам."""
import json
from odoo_db import OdooDB
from cfg_cut import *
db = OdooDB(); K = 12000
rows = db.query(f"""
select to_char(am.date,'YYYY-MM') m, p.payment_type pt, coalesce(p.x_studio_server_product, '') prod,
  coalesce(p.x_studio_payment_type, '') meth, coalesce(aj.name->>'ru_RU', aj.name->>'en_US') jr, c.name cur,
  count(*) n, sum(case when c.name='UZS' then p.amount_company_currency_signed
                       else (case when p.payment_type='inbound' then 1 else -1 end) * p.amount * {K} end) s,
  sum(p.amount_company_currency_signed) s_odoo, sum(case when c.name<>'UZS' then p.amount else 0 end) fx
from account_payment p join account_move am on am.id=p.move_id join account_journal aj on aj.id=am.journal_id
join res_currency c on c.id=p.currency_id
where am.state='posted' and am.date >= date '{P1}' and am.date <= date '{CUT}' and not p.is_internal_transfer
group by 1,2,3,4,5,6""")
out = [{'m': r['m'], 'pt': r['pt'], 'prod': r['prod'], 'meth': r['meth'], 'jr': r['jr'], 'cur': r['cur'], 'n': r['n'],
        's': float(r['s']), 's_odoo': float(r['s_odoo']), 'fx': float(r['fx'])} for r in rows]
json.dump(out, open(J('cash'), 'w'), ensure_ascii=False)
tin = sum(r['s'] for r in out if r['pt'] == 'inbound'); tout = sum(r['s'] for r in out if r['pt'] == 'outbound')
print('поступило %.1f млн, комиссии %.1f млн, платежей %d, в валюте $%.1f' % (tin / 1e6, tout / 1e6, sum(r['n'] for r in out), sum(r['fx'] for r in out)))
