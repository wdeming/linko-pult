# -*- coding: utf-8 -*-
"""История для финансов пульта: 2025 — файл финансиста PL_last_v.xlsx (учёт по оплатам), янв–мар 2026 — управленческий пакет (P&L PL_Linko, ДДС).
python3 build_fin_hist.py [PL_last_v.xlsx] [пакет.xlsx] [выход fin_hist.json]  — результат кладётся на сервер в /opt/linko-pult/data/fin_hist.json"""
import sys, json, datetime, collections
import openpyxl
from fin_groups import *
a = sys.argv[1:] + [None] * 3
F1 = a[0] or 'PL_last_v.xlsx'; F2 = a[1] or 'Linko_управленческий_пакет_Янв-Июл_2026.xlsx'; OUT = a[2] or 'fin_hist.json'
num = lambda v: float(v) if isinstance(v, (int, float)) else 0.0
H = {}
# ---------- 2025: лист P&L файла финансиста ----------
ws = openpyxl.load_workbook(F1, data_only=True)['P&L']
cols = {}
for c in range(4, ws.max_column + 1):
    d = ws.cell(2, c).value
    if isinstance(d, datetime.datetime) and d.year == 2025 and ws.cell(1, c).value == 'Actual': cols[c] = d.strftime('%Y-%m')
rows = {}
for r in range(3, ws.max_row + 1):
    lab = ws.cell(r, 2).value
    if lab: rows.setdefault(str(lab).strip(), r)
SKIP = {'Услуги связи и интернет', 'Профессиональные услуги (консалтинг, юристы, аудит)', 'IT услуги'}
LEAF_START = rows['Затраты на персонал']; LEAF_END = rows['Итого операционные расходы']
unk = set()
for c, m in cols.items():
    v = lambda lab: num(ws.cell(rows[lab], c).value)
    h = {'src': 'fin', 'rev': {'Подписки SFA': v('SFA выручка') - v('SFA Доработки (сумма)') - v('SFA Обучение (сумма)'), 'Подписки POS': v('POS выручка'), 'Разовые работы и прочее': v('SFA Доработки (сумма)') + v('SFA Обучение (сумма)') + v('Tap выручка') + v('Perfect Store выручка')},
         'exp': collections.defaultdict(float), 'chk': {'rev': v('Итого выручка'), 'opex': v('Итого операционные расходы'), 'ebitda': v('EBITDA')}}
    for r in range(LEAF_START, LEAF_END):
        lab = ws.cell(r, 2).value; unit = ws.cell(r, 3).value
        if not lab or unit != 'сум' or lab in SKIP or str(lab).startswith('Итого'): continue
        lab = str(lab).strip()
        if lab not in CAT: unk.add(lab)
        h['exp'][group(lab)] += num(ws.cell(r, c).value)
    tot = sum(h['exp'].values())
    assert abs(tot - h['chk']['opex']) < 1, (m, tot, h['chk']['opex'])
    assert abs(sum(h['rev'].values()) - h['chk']['rev']) < 1, (m, h['rev'], h['chk']['rev'])
    h['exp'] = dict(h['exp']); H[m] = h
# финансирование учредителями (лист CF_отчет): сен 2024 – сен 2025
ws = openpyxl.load_workbook(F1, data_only=True)['CF_отчет']; FUND = {}; CHARTER = {}
for r in range(1, 10):
    lab = ws.cell(r, 2).value
    if lab in ('Финансирование от учредителей', 'Вклад в устав'):
        for c in range(4, 17):
            hd = str(ws.cell(2, c).value or ''); val = num(ws.cell(r, c).value)
            if val: (FUND if lab.startswith('Фин') else CHARTER)[hd] = val
# ---------- янв–мар 2026: управленческий пакет ----------
wb = openpyxl.load_workbook(F2, data_only=True)
def codes(ws):
    res = {}
    for r in range(3, ws.max_row + 1):
        lab = ws.cell(r, 1).value; cm = str(ws.cell(r, 10).value or ''); code = cm.split(' ')[0] if cm[:1] in 'PCOTF' and '-' in cm[:6] else None
        vals = {'2026-%02d' % (i + 1): num(ws.cell(r, 2 + i).value) for i in range(7)}
        if code: res[code] = vals
        if lab: res.setdefault(str(lab).strip(), vals)
    return res
P = codes(wb['P&L Янв-Июл']); D = codes(wb['ДДС Янв-Июл'])
EXPC = {'C-051': G_HOST, 'C-052': G_TEL, 'C-053': G_IT, 'O-063': G_MKT, 'O-064': G_STAFF, 'O-065': G_RENT, 'O-066': G_OUT, 'O-067': G_BANK, 'O-068': G_OTH, 'T-091': G_TAX}
PACKCHK = {}
for i in range(7):
    m = '2026-%02d' % (i + 1)
    PACKCHK[m] = {'rev': P['Итого выручка'][m], 'ebitda': P['EBITDA'][m], 'net': P['ЧИСТАЯ ПРИБЫЛЬ'][m], 'in': D['Итого поступления'][m], 'end': D['Остаток на конец месяца'][m]}
    if i >= 3: continue
    capex = -D['CF-301'][m]
    h = {'src': 'pack', 'rev': {'Подписки SFA': P['P-011'][m], 'Подписки POS': P['P-013'][m], 'Разовые работы и прочее': P['P-012'][m] + P['P-017'][m] + P['P-018'][m] + P['P-020'][m]},
         'exp': {g: 0.0 for g in GROUPS}, CAPEX: capex, 'chk': PACKCHK[m]}
    for code, g in EXPC.items(): h['exp'][g] += -P[code][m]
    h['exp'][G_OTH] -= capex          # техника в пакете сидит в «Прочих OPEX» — выносим в капрасходы (по ДДС, CF-301)
    assert abs(sum(h['rev'].values()) - P['Итого выручка'][m]) < 0.01
    h['cash'] = {'open': D['Остаток на начало месяца'][m], 'in': D['Итого поступления'][m], 'pay': -D['Итого выплаты'][m] + capex, 'loans': D['CF-401'][m],
                 'adj': D['CF-501'][m], 'end': D['Остаток на конец месяца'][m], 'in_sfa': D['CF-101'][m], 'in_pos': D['CF-102'][m]}
    H[m] = h
for m, h in H.items():           # пакет в млн сум — переводим в сумы
    if h['src'] == 'pack':
        for k in ('rev', 'exp'): h[k] = {kk: v * 1e6 for kk, v in h[k].items()}
        h[CAPEX] *= 1e6; h['cash'] = {k: v * 1e6 for k, v in h['cash'].items()}; h['chk'] = {k: v * 1e6 for k, v in h['chk'].items()}
    else:
        h[CAPEX] = h['exp'].pop(CAPEX, 0.0); h[TAP] = h['exp'].pop(TAP, 0.0)
        for g in GROUPS: h['exp'].setdefault(g, 0.0)
PACKCHK = {m: {k: v * 1e6 for k, v in x.items()} for m, x in PACKCHK.items()}
BS = {}
ws = wb['Баланс 31.07']
for r in range(3, ws.max_row + 1):
    code, lab, val = ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value
    if isinstance(val, (int, float)) and (code or lab): BS[code or lab] = {'lab': lab, 'v': val * 1e6}
json.dump({'months': H, 'fund': {k: v for k, v in FUND.items()}, 'charter': CHARTER, 'pack_chk': PACKCHK, 'pack_bs': BS,
           'files': {'fin': F1.split('/')[-1], 'pack': F2.split('/')[-1]}, 'built': datetime.date.today().isoformat()}, open(OUT, 'w'), ensure_ascii=False, indent=0)
print('месяцев', len(H), sorted(H)); print('неизвестные статьи', unk)
for m in sorted(H):
    h = H[m]; rv = sum(h['rev'].values()); ex = sum(h['exp'].values())
    print(m, h['src'], 'выручка %.1f расходы %.1f EBITDA %.1f капрасх %.1f tap %.1f' % (rv / 1e6, ex / 1e6, (rv - ex) / 1e6, h.get(CAPEX, 0) / 1e6, h.get(TAP, 0) / 1e6), '| у источника EBITDA %.1f' % (h['chk']['ebitda'] / 1e6))
print('фин. учредителей', FUND, CHARTER, sum(FUND.values()) / 1e6)
