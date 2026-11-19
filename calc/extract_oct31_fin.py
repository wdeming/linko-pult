# -*- coding: utf-8 -*-
"""Финансы для банка: строки счетов с периодом услуги (признание выручки), остатки денежных счетов по месяцам,
дебиторка, непривязанные платежи, журнал MRR по каждой подписке по месяцам и лицензии. Только чтение Odoo."""
import json
from odoo_db import OdooDB
from cls import PROD_SQL
from cfg_cut import *
db = OdooDB(); K = 12000
out = {}
# 1. строки счетов клиентам (без НДС), доллары по 12 000
out['inv'] = [{'d': str(r['d']), 'pid': r['pid'], 'a': float(r['a'] or 0), 'd1': str(r['d1']) if r['d1'] else None, 'd2': str(r['d2']) if r['d2'] else None}
    for r in db.query(f"""select m.invoice_date d, l.product_id pid, l.deferred_start_date d1, l.deferred_end_date d2,
       (case when m.move_type='out_refund' then -1 else 1 end) * (case when m.currency_id=1 then l.price_subtotal*{K} else l.price_subtotal end) a
     from account_move_line l join account_move m on m.id=l.move_id
     where m.state='posted' and m.move_type in ('out_invoice','out_refund') and l.display_type='product'""")]   # все счета: выручка периода до среза признаётся и по счетам, выставленным позже
# 2. денежные счета: обороты по месяцам (без «Касса Азиз» — долг учредителям, и FOTON)
out['cash_acc'] = [{'code': r['code'], 'name': r['jr'], 'cur': 'USD' if r['cur'] == 1 else 'UZS', 'm': r['m'], 'bal': float(r['bal'] or 0), 'amc': float(r['amc'] or 0)}
    for r in db.query(f"""select a.code, a.currency_id cur, (select string_agg(distinct coalesce(j.name->>'ru_RU', j.name->>'en_US'), ', ') from account_journal j
         where j.default_account_id=a.id and j.type in ('bank','cash')) jr, to_char(l.date,'YYYY-MM') m, sum(l.balance) bal, sum(l.amount_currency) amc
     from account_account a join account_move_line l on l.account_id=a.id and l.parent_state='posted'
     where a.id in (select default_account_id from account_journal where type in ('bank','cash')) and a.code not in ('101506','101507') and l.date <= date '{CUT}'
     group by 1,2,3,4""")]
# 3. дебиторка: открытые счета, выставленные до среза (остаток на сегодня)
out['ar'] = [{'due': str(r['due']), 'res': float(r['res'] or 0)} for r in db.query(f"""select coalesce(invoice_date_due, invoice_date) due,
     case when currency_id=1 then amount_residual*{K} else amount_residual_signed end res
   from account_move where move_type='out_invoice' and state='posted' and payment_state in ('not_paid','partial') and invoice_date <= date '{CUT}'""")]
r = db.query(f"""select count(*) n, coalesce(sum(case when p.currency_id=1 then p.amount*{K} else p.amount_company_currency_signed end),0) s
   from account_payment p join account_move m on m.id=p.move_id where m.state='posted' and p.payment_type='inbound' and not p.is_internal_transfer
   and not p.is_reconciled and m.date <= date '{CUT}'""")[0]
out['unmatched'] = {'n': r['n'], 's': float(r['s'])}
# 4. журнал MRR по подпискам по месяцам + клиент, продукт, лицензии
out['log'] = [{'o': r['o'], 'm': r['m'], 'et': r['et'], 'a': float(r['a'])} for r in db.query(f"""select l.order_id o, to_char(l.event_date,'YYYY-MM') m, l.event_type et,
     sum(case when l.currency_id=1 then l.amount_signed*{K} else l.amount_signed end) a from sale_order_log l where l.event_date <= date '{CUT}' group by 1,2,3""")]
out['subs'] = {str(r['id']): {'p': r['partner_id'], 'prod': r['prod'], 'lic': float(r['lic'] or 0), 'kas': float(r['kas'] or 0)} for r in db.query(f"""
   select so.id, so.partner_id, {PROD_SQL} prod,
     (select sum(sol.product_uom_qty) from sale_order_line sol where sol.order_id=so.id and sol.product_id in (25,26,41,42)) lic,
     (select sum(sol.product_uom_qty) from sale_order_line sol where sol.order_id=so.id and sol.product_id in (24,40,43,32)) kas
   from sale_order so where so.id in (select distinct order_id from sale_order_log)""")}
json.dump(out, open(J('fin'), 'w'), ensure_ascii=False)
print('строк счетов', len(out['inv']), '| с периодом', sum(1 for x in out['inv'] if x['d1']), '| счетов дебиторки', len(out['ar']), '| непривязанных платежей', out['unmatched'],
      '| подписок', len(out['subs']), '| записей журнала', len(out['log']))
