# -*- coding: utf-8 -*-
"""Обзвоны клиентов (x_nps): сырая выгрузка в формате файла «Обзвоны клиентов — сырая выгрузка из Odoo»."""
import json, re, datetime as dt
from odoo_db import OdooDB
from cfg_cut import *
db = OdooDB()
LAB = {}
for r in db.query("""select f.name fld, s.value, coalesce(s.name->>'ru_RU', s.name->>'en_US') lab
  from ir_model_fields_selection s join ir_model_fields f on f.id=s.field_id where f.model='x_nps'"""):
    LAB[(r['fld'], r['value'])] = r['lab']
# устаревшее поле статуса выводится в исходных (en_US) подписях — как в прежних выгрузках
for r in db.query("""select s.value, coalesce(s.name->>'en_US', s.name->>'ru_RU') lab from ir_model_fields_selection s
  join ir_model_fields f on f.id=s.field_id where f.model='x_nps' and f.name='x_studio_selection_field_8r5_1j56e5rdi'"""):
    LAB[('x_studio_selection_field_8r5_1j56e5rdi', r['value'])] = r['lab']
DEP = {'Департамент продаж POS': 'Продажи POS', 'Департамент продаж SFA': 'Продажи SFA',
       'Отдел по работе с ключевыми клиентами': 'Ключевые клиенты', 'Департамент разработки': 'Разработка',
       'Отдел тех. поддержки': 'Тех. поддержка', 'Департамента по административным вопросам': 'Департамент по адм. вопросам',
       'Департамент по работе с партнерами': 'Партнёры'}
CR = {'Клиент отказался от использования системы по собственным причинам.': 'Клиент отказался по собственным причинам',
      'Несоответствие системы бизнес-процессам клиента.': 'Система не соответствует бизнес-процессам клиента',
      'Обнаружены проблемы, связанные с продуктом (IT).': 'Проблемы с продуктом (IT)',
      'Недостатки в процессе онбординга.': 'Недостатки онбординга',
      'Возникли трудности при взаимодействии с технической поддержкой.': 'Трудности при взаимодействии с техподдержкой',
      'Ошибки, допущенные на этапе продажи.': 'Ошибки на этапе продажи'}
ROUTE = {'Онбординг': 'Отдел внедрения (онбординг)', 'Тех. поддержка': 'Техподдержка', 'Продукт (IT)': 'Разработка / IT',
         'Продажа': 'Отдел продаж', 'Проблема в клиенте': 'На стороне клиента (не направлялась)'}
# дата закрытия из истории (смена этапа на Решено/Отмена)
closed = {}
for r in db.query("""select m.res_id rid, max(m.date)::date d from mail_tracking_value tv
  join ir_model_fields f on f.id=tv.field_id join mail_message m on m.id=tv.mail_message_id
  where m.model='x_nps' and f.name='x_studio_selection_field_8r5_1j56e5rdi' and tv.new_value_char in ('Решено','Отмена') group by 1"""):
    closed[r['rid']] = r['d']
rows = db.query(f"""
select n.id, n.x_studio_date d, coalesce(rp.complete_name, rp.name) partner, coalesce(n.x_name->>'ru_RU', n.x_name->>'en_US') x_name, st.name st_name, co.code cc,
  coalesce(nullif(n.x_studio_server,''), nullif(n.x_studio_server_sale,''), nullif(ls.x_name->>'ru_RU','')) srv, so.name so_name, n.x_studio_reason reason, n.x_studio_first_problem fp,
  up.name resp, coalesce(hd.name->>'ru_RU', hd.name->>'en_US') resp_dep,
  sn.x_name->>'ru_RU' stage, n.x_studio_selection_field_8r5_1j56e5rdi st2, n.x_studio_cancel_reason cr,
  n.x_studio_ball ball, n.x_studio_problem problem, n.x_studio_comment comment, n.x_studio_action action,
  cp.name creator, n.create_date::date cd, n.write_date::date wd,
  coalesce(nullif(n.x_studio_phone,''), nullif(n.x_studio_no_client_phone,'')) phone, n.x_active act
from x_nps n
left join res_partner rp on rp.id=n.x_studio_paertner
left join res_country_state st on st.id=n.x_studio_state left join res_country co on co.id=st.country_id
left join sale_order so on so.id=n.x_studio_sale_order
left join res_users ru on ru.id=n.x_studio_resp_user left join res_partner up on up.id=ru.partner_id
left join lateral (select e.department_id from hr_employee e where e.user_id=ru.id order by e.active desc, e.id desc limit 1) he on true
left join hr_department hd on hd.id=he.department_id
left join x_state_nps sn on sn.id=n.x_state_nps
left join x_last_sale_date ls on ls.id=n.x_studio_last_sale_date
left join res_users cu on cu.id=n.create_uid left join res_partner cp on cp.id=cu.partner_id
where n.x_studio_date is null or n.x_studio_date <= date '{CUT}'
order by n.x_studio_date nulls last, n.id""")
def clean(t):
    if t is None: return None
    t = re.sub(r'<[^>]+>', ' ', t) if '<' in t else t
    return t
out = []
for r in rows:
    reason = r['reason']
    res = 'Нет данных' if not reason else ('Всё ок' if reason == 'Все хорошо' else ('Не дозвонились' if reason == 'Не дозвонились' else 'Проблема'))
    stage = r['stage']; cl = closed.get(r['id']) if stage in ('Решено', 'Отмена') else None
    days = None
    if stage in ('Решено', 'Отмена') and r['d']:
        end = cl or r['wd']
        days = max((end - r['d']).days, 0) if end else None
    ph = re.sub(r'\D', '', r['phone'] or '')
    out.append([r['id'], str(r['d']) if r['d'] else None, r['d'].strftime('%Y-%m') if r['d'] else None,
        r['partner'] or r['x_name'], ('%s (%s)' % (r['st_name'], r['cc'])) if r['st_name'] else None,
        'SFA' if r['so_name'] else 'POS', r['srv'], r['so_name'], res,
        LAB.get(('x_studio_reason', reason), reason), r['fp'], ROUTE.get(r['fp']),
        r['resp'], DEP.get(r['resp_dep'], r['resp_dep']), stage, LAB.get(('x_studio_selection_field_8r5_1j56e5rdi', r['st2']), r['st2']),
        CR.get(LAB.get(('x_studio_cancel_reason', r['cr']), r['cr']), LAB.get(('x_studio_cancel_reason', r['cr']), r['cr'])), r['ball'] or None,
        r['problem'], r['comment'], r['action'], r['creator'], str(r['cd']) if r['cd'] else None,
        str(cl) if cl else None, str(r['wd']) if r['wd'] else None, days,
        ('***' + ph[-4:]) if len(ph) >= 4 else None, 'да' if r['act'] else 'нет'])
json.dump(out, open(J('calls'), 'w'), ensure_ascii=False)
print('записей', len(out), '| без даты', sum(1 for o in out if not o[1]), '| период', out[0][1], '—', max(o[1] for o in out if o[1]))
