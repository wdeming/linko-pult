# -*- coding: utf-8 -*-
import json
from odoo_db import OdooDB
from cfg_cut import *
db = OdooDB()
rows = db.query(f"""select coalesce(d.name->>'ru_RU', d.name->>'en_US') dep, count(*) n, sum(coalesce(e.x_studio_salary_1,0)) sal,
   count(*) filter (where not e.active) gone, string_agg(case when not e.active then e.name || ' (' || to_char(e.departure_date,'DD.MM.YYYY') || ')' end, ', ' order by e.departure_date) gone_names
 from hr_employee e left join hr_department d on d.id=e.department_id
 where (e.active or e.departure_date > date '{CUT}') and e.create_date::date <= date '{CUT}' and e.company_id=1 group by 1 order by 3 desc""")
json.dump([{'dep': r['dep'], 'n': r['n'], 'sal': float(r['sal'] or 0), 'gone': r['gone'], 'gone_names': r['gone_names'] or ''} for r in rows], open(J('fot'), 'w'), ensure_ascii=False)
print('штат', sum(r['n'] for r in rows))
