# -*- coding: utf-8 -*-
"""Актуальные данные «на сегодня» для верхнего блока пульта (только чтение из Odoo). Печатает JSON.
MRR — нарастающий итог журнала изменений MRR по сегодняшний день; движения, поступления и счета — с 1-го числа месяца по сегодня.
Валюта — по 12 000 сум/$."""
import json, datetime, calendar, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__))); sys.path.insert(1, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from odoo_db import OdooDB
from cls import PROD_SQL, AMT_SQL
K = 12000
today = datetime.date.today(); m0 = today.replace(day=1); prev_end = m0 - datetime.timedelta(days=1)
p0 = prev_end.replace(day=1); p_same = p0.replace(day=min(today.day, calendar.monthrange(p0.year, p0.month)[1]))
db = OdooDB(); q = db.query
def byprod(sql):
    return {r['prod']: float(r['s'] or 0) for r in q(sql)}
LOG = f"select {PROD_SQL} prod, sum({AMT_SQL}) s from sale_order_log l join sale_order so on so.id=l.order_id where "
mrr_now = byprod(LOG + f"l.event_date <= date '{today}' group by 1")
mrr_prev = byprod(LOG + f"l.event_date <= date '{prev_end}' group by 1")
mov = {}
for r in q(f"""select {PROD_SQL} prod, l.event_type et, sum({AMT_SQL}) s from sale_order_log l join sale_order so on so.id=l.order_id
               where l.event_date >= date '{m0}' and l.event_date <= date '{today}' group by 1,2"""):
    mov.setdefault(r['et'], {})[r['prod']] = float(r['s'] or 0)
subs = {r['prod']: [r['n'], r['c']] for r in q(f"""select {PROD_SQL} prod, count(*) n, count(distinct so.partner_id) c from sale_order so
          where so.is_subscription and so.subscription_state in ('3_progress','4_paused') group by 1""")}
clients = q("""select count(distinct partner_id) c from sale_order where is_subscription and subscription_state in ('3_progress','4_paused')""")[0]['c']
CASH = f"""select coalesce(sum(case when p.payment_type='inbound' then (case when c.name='UZS' then p.amount_company_currency_signed else p.amount*{K} end) end),0) i,
  coalesce(-sum(case when p.payment_type='outbound' then (case when c.name='UZS' then p.amount_company_currency_signed else -p.amount*{K} end) end),0) o
  from account_payment p join account_move am on am.id=p.move_id join res_currency c on c.id=p.currency_id
  where am.state='posted' and not p.is_internal_transfer and am.date >= date '%s' and am.date <= date '%s'"""
cash = q(CASH % (m0, today))[0]; cash_p = q(CASH % (p0, p_same))[0]
INV = f"""select coalesce(sum(case when c.name='UZS' then aml.price_subtotal else aml.price_subtotal*{K} end),0) s
  from account_move_line aml join account_move am on am.id=aml.move_id join res_currency c on c.id=aml.currency_id
  where aml.parent_state='posted' and am.move_type in ('out_invoice','out_refund') and aml.display_type='product' and aml.date >= date '%s' and aml.date <= date '%s'"""
inv = float(q(INV % (m0, today))[0]['s']); inv_p = float(q(INV % (p0, p_same))[0]['s'])
db.close()
P = ('SFA', 'POS')
print(json.dumps({'today': str(today), 'month_start': str(m0), 'prev_end': str(prev_end), 'prev_same': str(p_same),
  'mrr': {p: mrr_now.get(p, 0) for p in P}, 'mrr_prev': {p: mrr_prev.get(p, 0) for p in P},
  'mov': {et: {p: v.get(p, 0) for p in P} for et, v in mov.items()},
  'subs': {p: subs.get(p, [0, 0])[0] for p in P}, 'clients': clients,
  'cash_in': float(cash['i']), 'cash_comm': float(cash['o']), 'cash_in_prev': float(cash_p['i']),
  'inv': inv, 'inv_prev': inv_p}, ensure_ascii=False))
