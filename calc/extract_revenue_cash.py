# -*- coding: utf-8 -*-
"""Выручка по дате поступления денег: входящие платежи клиентов, разложенные по продуктам через зачтённые счета."""
import json
from odoo_db import OdooDB
from cfg_cut import *
db = OdooDB()
PAY = f"""select p.id pay, am.date d, p.amount_company_currency_signed amt, coalesce(p.x_studio_server_product,'') sp, aml.id lid
  from account_payment p join account_move am on am.id=p.move_id
  join account_move_line aml on aml.move_id=am.id join account_account aa on aa.id=aml.account_id and aa.account_type='asset_receivable'
  where am.state='posted' and p.payment_type='inbound' and p.partner_type='customer' and am.date >= date '{P1}' and am.date < date '{P2}'"""
rows = db.query(f"""
with pay as ({PAY}),
parts as (select pay.pay, pay.d, apr.amount a, il.move_id inv from pay join account_partial_reconcile apr on apr.credit_move_id=pay.lid
          join account_move_line il on il.id=apr.debit_move_id join account_move im on im.id=il.move_id and im.move_type='out_invoice' and im.state='posted'),
lines as (select move_id inv, product_id pid, -balance bal, quantity q from account_move_line where display_type='product' and move_id in (select inv from parts)),
invt as (select inv, sum(bal) tb from lines group by inv having sum(bal) <> 0)
select to_char(parts.d,'YYYY-MM') m, lines.pid, sum(parts.a*lines.bal/invt.tb) amt, sum(parts.a*lines.q/invt.tb) qty, count(distinct parts.pay) n
from parts join lines using(inv) join invt using(inv) group by 1,2""")
bypid = {'%s|%s' % (r['m'], r['pid']): [float(r['amt']), float(r['qty'] or 0), r['n']] for r in rows}
un = db.query(f"""
with pay as ({PAY}),
alloc as (select pay.pay, sum(apr.amount) a from pay join account_partial_reconcile apr on apr.credit_move_id=pay.lid
          join account_move_line il on il.id=apr.debit_move_id join account_move im on im.id=il.move_id and im.move_type='out_invoice' and im.state='posted'
          join (select move_id from account_move_line where display_type='product' group by move_id having sum(balance) <> 0) ok on ok.move_id=im.id
          group by 1)
select to_char(pay.d,'YYYY-MM') m, pay.sp, sum(pay.amt - coalesce(alloc.a,0)) amt, count(*) filter (where pay.amt - coalesce(alloc.a,0) > 0.5) n
from pay left join alloc on alloc.pay=pay.pay group by 1,2""")
unapplied = {'%s|%s' % (r['m'], r['sp'] or 'Нет'): [float(r['amt']), r['n']] for r in un}
cin = {r['m']: [float(r['s']), r['n']] for r in db.query(f"select to_char(d,'YYYY-MM') m, sum(amt) s, count(*) n from ({PAY}) x group by 1")}
out = {r['m']: float(r['s']) for r in db.query(f"""select to_char(am.date,'YYYY-MM') m, sum(p.amount_company_currency_signed) s from account_payment p join account_move am on am.id=p.move_id
  where am.state='posted' and p.payment_type='outbound' and am.date >= date '{P1}' and am.date < date '{P2}' group by 1""")}
accr = {r['m']: float(r['s']) for r in db.query(f"""select to_char(aml.date,'YYYY-MM') m, sum(-aml.balance) s from account_move_line aml join account_move am on am.id=aml.move_id
  where aml.parent_state='posted' and am.move_type in ('out_invoice','out_refund') and aml.display_type='product' and aml.date >= date '{P1}' and aml.date < date '{P2}' group by 1""")}
accr_pid = {'%s|%s' % (r['m'], r['pid']): float(r['s']) for r in db.query(f"""select to_char(aml.date,'YYYY-MM') m, aml.product_id pid, sum(-aml.balance) s from account_move_line aml join account_move am on am.id=aml.move_id
  where aml.parent_state='posted' and am.move_type in ('out_invoice','out_refund') and aml.display_type='product' and aml.date >= date '{P1}' and aml.date < date '{P2}' group by 1,2""")}
names = {str(r['id']): r['n'] for r in db.query("select pp.id, coalesce(pt.name->>'ru_RU', pt.name->>'en_US') n from product_product pp join product_template pt on pt.id=pp.product_tmpl_id")}
json.dump({'bypid': bypid, 'unapplied': unapplied, 'cin': cin, 'out': out, 'accr': accr, 'accr_pid': accr_pid, 'names': names}, open(J('revcash'), 'w'), ensure_ascii=False)
tot_alloc = sum(v[0] for v in bypid.values()); tot_un = sum(v[0] for v in unapplied.values()); tot_in = sum(v[0] for v in cin.values())
print('поступило %.1f млн = по счетам %.1f + без счёта %.1f (сходится: %s) | исходящие %.1f | начислено %.1f' % (tot_in/1e6, tot_alloc/1e6, tot_un/1e6, abs(tot_in-tot_alloc-tot_un) < 1, sum(out.values())/1e6, sum(accr.values())/1e6))
