# -*- coding: utf-8 -*-
"""MRR и когортное удержание — в формате исходного файла (4 листа на формулах)."""
import json, copy, collections, openpyxl
from cfg_cut import *
TPL = '_шаблоны_форматов/tpl_mrr.xlsx'
d = json.load(open(J('mrr'))); c = json.load(open(J('const')))
ET = ('0_creation', '1_expansion', '15_contraction', '2_churn')
mov = collections.defaultdict(float)
for r in d['moves']: mov[(r['prod'], r['m'], r['et'])] += r['amt']
mrr_end = c['MRR']
net = {p: sum(mov[(p, m, e)] for m in MON for e in ET) for p in ('POS', 'SFA')}
mrr_start = {p: mrr_end[p] - net[p] for p in ('POS', 'SFA')}
sp = lambda v: f"{v:,.0f}".replace(',', ' ')
NOTE = ('Источник: Odoo, sale.order.log (движение MRR) и sale.order (факт по подпискам). Срез %s, выгрузка %s. '
        'Валютные подписки пересчитаны по курсу 12 000 сум/$ (управленческий курс).' % (CUT_RU, EXPORT_RU))
HEAL = ('Замечание к сравнению с прежними версиями: при повторном открытии закрытой подписки Odoo удаляет событие оттока из журнала. '
        'Поэтому отток последних месяцев, измеренный сразу после конца месяца, завышен и со временем уменьшается '
        '(основная часть возвратов происходит в первые 2–3 недели). Все месяцы в этом файле измерены по состоянию журнала на %s.' % EXPORT_RU)
wb = openpyxl.load_workbook(TPL)
def styles(ws, r): return [copy.copy(ws.cell(r, k)._style) for k in range(1, ws.max_column + 1)]
def wipe(ws, r0):
    blank = styles(ws, 2)
    for r in range(r0, ws.max_row + 40):
        for k in range(1, len(blank) + 1):
            ws.cell(r, k).value = None; ws.cell(r, k)._style = copy.copy(blank[k - 1])
        ws.row_dimensions[r].height = None
def put(ws, r, vals, st):
    for k in range(1, len(st) + 1):
        ws.cell(r, k)._style = copy.copy(st[k - 1])
    for k, v in enumerate(vals, 1):
        if v is not None: ws.cell(r, k).value = v
def notes(ws, r, texts, st):
    for t in texts:
        put(ws, r, [t], st); r += 1

# ================= лист 1 =================
ws = wb['MRR по продуктам']
s_sec, s_m, s_tot, s_co, s_note = styles(ws, 4), styles(ws, 5), styles(ws, 12), styles(ws, 24), styles(ws, 26)
wipe(ws, 4)
r = 4; tot_rows = {}
for prod in ('POS', 'SFA'):
    put(ws, r, ['ПРОДУКТ ' + prod], s_sec); r += 1; f0 = r
    for i, m in enumerate(MON):
        put(ws, r, ['%s 2026' % RUS[m], mrr_start[prod] if i == 0 else '=H%d' % (r - 1),
                    mov[(prod, m, '0_creation')], mov[(prod, m, '1_expansion')], mov[(prod, m, '15_contraction')], mov[(prod, m, '2_churn')],
                    '=C%d+D%d+E%d+F%d' % (r, r, r, r), '=B%d+G%d' % (r, r), '=G%d/B%d' % (r, r)], s_m); r += 1
    l0 = r - 1
    put(ws, r, ['ИТОГО %s, %s 2026' % (prod, PERIOD), '=B%d' % f0] + ['=SUM(%s%d:%s%d)' % (k, f0, k, l0) for k in 'CDEFG']
        + ['=H%d' % l0, '=G%d/B%d' % (r, r)], s_tot); tot_rows[prod] = r; r += 2
put(ws, r, ['ИТОГО КОМПАНИЯ, MRR на %s' % CUT_RU, None, None, None, None, None, None, '=H%d+H%d' % (tot_rows['POS'], tot_rows['SFA'])], s_co); r += 2
notes(ws, r, [NOTE,
  'Период — с февраля 2026. Январь исключён: в нём в Odoo переносили действующую базу (у SFA 318 записей на 431 млн одним месяцем), это не продажи и не движение MRR.',
  '«MRR на конец» месяца — нарастающий итог журнала изменений MRR в Odoo (sale.order.log: создания, расширения, сокращения, отток) по последний день месяца, как график «Изменение MRR» в Odoo; USD по 12 000. «MRR на начало» февраля — тот же итог на 31.01.2026: POS %s, SFA %s.' % (sp(mrr_start['POS']), sp(mrr_start['SFA'])),
  '«Сокращения» и «Расширения» — изменения суммы действующих подписок (downgrade/upgrade), «Отток» — закрытие подписки целиком.',
  HEAL], s_note)
# ================= лист 2 =================
ws = wb['GRR и NRR помесячно']
s_sec, s_m, s_note = styles(ws, 4), styles(ws, 5), styles(ws, 22)
wipe(ws, 4)
r = 4; GN = {}
for prod in ('POS', 'SFA'):
    put(ws, r, ['ПРОДУКТ ' + prod], s_sec); r += 1; run = mrr_start[prod]
    for m in MON:
        ch, co, ex = mov[(prod, m, '2_churn')], mov[(prod, m, '15_contraction')], mov[(prod, m, '1_expansion')]
        put(ws, r, ['%s 2026' % RUS[m], run, ch, co, ex, '=B%d+C%d+D%d' % (r, r, r), '=IF(B%d>0,F%d/B%d,"")' % (r, r, r),
                    '=F%d+E%d' % (r, r), '=IF(B%d>0,H%d/B%d,"")' % (r, r, r)], s_m)
        GN[(prod, m)] = {'start': run, 'grr': (run + ch + co) / run, 'nrr': (run + ch + co + ex) / run}
        run += sum(mov[(prod, m, e)] for e in ET); r += 1
    r += 1
notes(ws, r, [NOTE,
  'GRR (Gross Revenue Retention) = (MRR на начало − отток − сокращения) / MRR на начало. Расширения не учитываются, потолок 100%.',
  'NRR (Net Revenue Retention) = (MRR на начало − отток − сокращения + расширения) / MRR на начало. Больше 100% означает, что рост действующей базы перекрывает потери.',
  'Новые продажи в расчёт удержания НЕ входят — метрика показывает, что происходит с уже имеющейся базой.',
  'Февраль–март — первые месяцы после переноса базы в Odoo, там ещё много технических правок сумм; наиболее чистое окно — апрель–%s 2026.' % RUS[LAST_M].lower()], s_note)
# ================= лист 3 =================
ws = wb['Когорты Фев–Авг 2026']; ws.title = 'Когорты %s 2026' % PERIOD
ws['A1'] = ws['A1'].value.replace('Фев–Авг', PERIOD)
s_m, s_tot, s_note = styles(ws, 4), styles(ws, 11), styles(ws, 13)
wipe(ws, 4)
coh = {(x['coh'], x['prod']): x for x in d['cohorts']}
fact = {(x['coh'], x['prod']): x['mrr_now'] - x['aft'] for x in d['cohort_fact']}
cev = collections.defaultdict(dict)
for x in d['cohort_events']: cev[(x['coh'], x['prod'])][x['et']] = x['amt']
r = 4
for i, m in enumerate(MON):
    v = ['%s 2026' % RUS[m]]
    for prod, L in (('POS', 'BCEF'), ('SFA', 'HIKL')):
        k = (m, prod)
        v += [coh[k]['n'], coh[k]['alive'], '=IF(%s%d>0,%s%d/%s%d,"")' % (L[0], r, L[1], r, L[0], r),
              cev[k].get('0_creation', 0), fact[k], '=IF(%s%d>0,%s%d/%s%d,"")' % (L[2], r, L[3], r, L[2], r)]
    v.append(NM - 1 - i)
    put(ws, r, v, s_m); r += 1
l0 = r - 1
put(ws, r, ['ИТОГО %s 2026' % PERIOD, '=SUM(B4:B%d)' % l0, '=SUM(C4:C%d)' % l0, '=C%d/B%d' % (r, r), '=SUM(E4:E%d)' % l0, '=SUM(F4:F%d)' % l0, '=F%d/E%d' % (r, r),
            '=SUM(H4:H%d)' % l0, '=SUM(I4:I%d)' % l0, '=I%d/H%d' % (r, r), '=SUM(K4:K%d)' % l0, '=SUM(L4:L%d)' % l0, '=L%d/K%d' % (r, r)], s_tot); r += 2
notes(ws, r, [NOTE,
  'Когорта — месяц первого контракта подписки (first_contract_date). «Живых» — подписка не была закрыта по состоянию на %s (статусы «в работе» и «на паузе»).' % CUT_RU,
  '«Первоначальный MRR» — сумма событий «создание подписки» по этой когорте. NRR = MRR сейчас / первоначальный MRR: сколько денег осталось от когорты с учётом апсейлов и сокращений.',
  '«Месяцев с начала» — сколько месяцев когорта прожила к %s. Сравнивать когорты между собой корректно только при одинаковом возрасте: у когорты последнего месяца удержание почти 100%% просто потому, что она новая.' % CUT_RU,
  'MRR валютных подписок пересчитан в сумы по курсу 12 000 — расхождения с общим MRR нет.',
  'Январь 2026 исключён из отчёта (методологическое решение менеджмента): в этом месяце базу переносили в Odoo и данные неполные.'], s_note)
# ================= лист 4 =================
ws = wb['Когорты — GRR и NRR']
ws['A1'] = ws['A1'].value.replace('Фев–Авг', PERIOD)
s_m, s_note = styles(ws, 4), styles(ws, 11)
wipe(ws, 4)
Q = [('Фев–Мар 2026', ['2026-02', '2026-03'])]
for q, ms in ((2, [4, 5, 6]), (3, [7, 8, 9]), (4, [10, 11, 12])):
    have = ['2026-%02d' % x for x in ms if '2026-%02d' % x in MON]
    if not have: continue
    lab = '%d кв. 2026' % q
    if len(have) < 3: lab += ' (%s)' % ('–'.join([RUS_SH[have[0]], RUS_SH[have[-1]]]) if len(have) > 1 else RUS_SH[have[0]])
    Q.append((lab, have))
r = 4; COHQ = {}; maxdiff = {'POS': 0, 'SFA': 0}
for lab, ms in Q:
    for prod in ('POS', 'SFA'):
        a = collections.defaultdict(float); f = 0.0
        for m in ms:
            for e, v in cev[(m, prod)].items(): a[e] += v
            f += fact[(m, prod)]
        put(ws, r, [lab, prod, a['0_creation'], a['1_expansion'], a['15_contraction'], a['2_churn'], '=C%d+D%d+E%d+F%d' % (r, r, r, r), f,
                    '=H%d-G%d' % (r, r), '=(C%d+E%d+F%d)/C%d' % (r, r, r, r), '=H%d/C%d' % (r, r)], s_m)
        COHQ['%s|%s' % (lab, prod)] = {'init': a['0_creation'], 'now': f, 'grr': (a['0_creation'] + a['15_contraction'] + a['2_churn']) / a['0_creation'], 'nrr': f / a['0_creation']}
        maxdiff[prod] = max(maxdiff[prod], abs(f - sum(a[e] for e in ET))); r += 1
r += 1
last_lab, last_ms = Q[-1]
notes(ws, r, [NOTE,
  'Первоначальный MRR — сумма событий «создание подписки» по когорте. Расширения / сокращения / отток — все последующие события этих же подписок.',
  'NRR = MRR сейчас / первоначальный MRR. Показывает, сколько денег осталось от когорты с учётом апсейлов.',
  'GRR = (первоначальный + сокращения + отток) / первоначальный. Может уходить ниже нуля: если подписку сначала расширили, а потом сократили, сокращение считается от увеличенной суммы.',
  'Расхождение между расчётом и фактом — подписки, у которых в логах нет события создания (переносы, ручные правки). По POS — до %s млн на когорту, по SFA — до %s млн.' % (('%.1f' % (maxdiff['POS'] / 1e6)).replace('.', ','), ('%.1f' % (maxdiff['SFA'] / 1e6)).replace('.', ',')),
  'Взят период с февраля 2026: январь — месяц переноса базы в Odoo, более ранние периоды логами движения MRR не покрыты.',
  'Когорта «%s» — самая молодая (%s), её удержание пока не сравнивать с остальными.' % (last_lab, 'ей меньше месяца' if len(last_ms) == 1 else 'ей меньше %d месяцев' % len(last_ms))], s_note)
fn = OUTP('MRR_и_когортное_удержание_2026_%s.xlsx' % SUF)
wb.save(fn)
json.dump({'start': mrr_start, 'end': mrr_end, 'mov': {'%s|%s|%s' % k: v for k, v in mov.items()}, 'GN': {'%s|%s' % k: v for k, v in GN.items()},
           'COHQ': COHQ, 'coh': {'%s|%s' % k: dict(v, init=cev[k].get('0_creation', 0), now=fact[k]) for k, v in coh.items()}}, open(J('res_mrr'), 'w'), ensure_ascii=False)
print(fn)
for p in ('POS', 'SFA'): print(p, 'старт', sp(mrr_start[p]), 'net', sp(net[p]), 'конец', sp(mrr_end[p]))
