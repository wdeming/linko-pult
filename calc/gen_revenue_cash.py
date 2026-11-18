# -*- coding: utf-8 -*-
"""Выручка по продуктам и типам — по дате поступления денег (кэшфло), в формате файла комплекта."""
import json, copy, collections, openpyxl
from openpyxl.utils import get_column_letter as L
from cfg_cut import *
d = json.load(open(J('revcash'), encoding='utf-8'))
TYPES = [('SFA', 'Подписка SFA', 'SFA full, SFA part (в т.ч. валютные счета)', 'Регулярная (recurring)'),
         ('POS', 'Подписка POS', 'POS Подписка, POS бизнес, POS interprice, POS Ritm', 'Регулярная (recurring)'),
         ('EDU', 'Обучение', 'Услуга обучения POS, обучение SFA, обучение AI', 'Разовая при подключении'),
         ('OTH', 'Прочее', 'Perfect store, Производство, доработки SFA/POS, AI ассистент, прочие позиции', 'Разовая / проектная'),
         ('UN', 'Без привязки к счёту', 'Поступления, ещё не зачтённые ни к одному счёту: авансы, оплаты без счёта', 'Не определён')]
CAT = {25: 'SFA', 26: 'SFA', 41: 'SFA', 42: 'SFA', 24: 'POS', 40: 'POS', 32: 'POS', 43: 'POS', 30: 'EDU', 37: 'EDU', 33: 'EDU'}
LAB = {25: 'SFA full', 26: 'SFA part', 41: 'SFA full USD (в сумах)', 42: 'SFA part USD (в сумах)', 24: 'POS Подписка', 40: 'POS бизнес', 32: 'POS Подписка (Ritm)', 43: 'POS interprice',
       30: 'Услуга обучения POS', 37: 'Обучение AI', 33: 'Услуга обучения SFA', 35: 'Perfect store', 38: 'Доработка SFA', 36: 'Производство', 39: 'Доработка POS', 44: 'AI ассистент'}
ORDER = [25, 26, 41, 42, 24, 40, 32, 43, 30, 37, 33, 35, 38, 36, 39, 44]
cat = lambda pid: CAT.get(pid, 'OTH')
pidk = lambda k: None if k.split('|')[1] == 'None' else int(k.split('|')[1])
P = collections.defaultdict(lambda: [0.0, 0.0, 0]); A = collections.defaultdict(float); Mt = collections.defaultdict(float); Ma = collections.defaultdict(float)
for k, (amt, qty, n) in d['bypid'].items():
    pid, m = pidk(k), k.split('|')[0]
    if m not in MON: continue
    P[pid][0] += amt; P[pid][1] += qty; P[pid][2] += n; Mt[(cat(pid), m)] += amt
for k, v in d['accr_pid'].items():
    pid, m = pidk(k), k.split('|')[0]
    if m in MON: A[cat(pid)] += v; Ma[(cat(pid), m)] += v
U = collections.defaultdict(lambda: [0.0, 0])
for k, (amt, n) in d['unapplied'].items():
    m, sp = k.split('|')
    if m in MON: U[sp][0] += amt; U[sp][1] += n; Mt[('UN', m)] += amt
G = {t[0]: sum(v[0] for p, v in P.items() if cat(p) == t[0]) for t in TYPES[:4]}; G['UN'] = sum(v[0] for v in U.values())
CIN = sum(d['cin'].get(m, [0])[0] for m in MON); OUT = sum(d['out'].get(m, 0) for m in MON); ACC = sum(d['accr'].get(m, 0) for m in MON)
assert abs(sum(G.values()) - CIN) < 1, (sum(G.values()), CIN)
wb = openpyxl.load_workbook('_шаблоны_форматов/tpl_revenue.xlsx')
def st(ws, r, n): return [copy.copy(ws.cell(r, c)._style) for c in range(1, n + 1)]
def wipe(ws, r0, n):
    blank = st(ws, 2, n)
    for r in range(r0, ws.max_row + 30):
        for c in range(1, n + 1): ws.cell(r, c).value = None; ws.cell(r, c)._style = copy.copy(blank[c - 1])
def put(ws, r, vals, sty):
    for c in range(1, len(sty) + 1): ws.cell(r, c)._style = copy.copy(sty[c - 1])
    for c, v in enumerate(vals, 1):
        if v is not None: ws.cell(r, c).value = v
SRC = ('Источник: Odoo, account.payment — проведённые входящие платежи клиентов по дате поступления денег, %s–%s. Продукт и тип определены по счетам, '
       'к которым платёж зачтён (сумма платежа делится пропорционально строкам счёта). Выгрузка %s.' % ('01.02.2026', CUT_RU, EXPORT_RU))
# ---------- Выручка по типам ----------
ws = wb['Выручка по типам']; n = 6
s_h, s_r, s_t, s_sec, s_h2, s_r2, s_n = st(ws, 3, n), st(ws, 4, n), st(ws, 8, n), st(ws, 10, n), st(ws, 11, n), st(ws, 12, n), st(ws, 17, n)
wipe(ws, 3, n)
ws['A1'] = 'Выручка по продуктам и типам — по дате поступления денег, %s' % PERIOD_LONG
put(ws, 3, ['Тип выручки', 'Что входит', 'Поступило, сум', 'Доля', 'Справочно: по дате счёта, сум', 'Характер выручки'], s_h)
r = 4
for k, name, what, kind in TYPES:
    put(ws, r, [name, what, G[k], '=C%d/$C$9' % r, A.get(k) if k != 'UN' else None, kind], s_r); r += 1
put(ws, r, ['ИТОГО поступило от клиентов', None, '=SUM(C4:C8)', '=C9/C9', '=SUM(E4:E8)', None], s_t); r += 2
put(ws, r, ['Группировка по характеру выручки'], s_sec); r += 1
put(ws, r, ['Характер', 'Поступило, сум', 'Доля'], s_h2); r += 1
for lab, f in (('Регулярная выручка (подписки)', '=C4+C5'), ('Разовая выручка (обучение и проекты)', '=C6+C7'), ('Без привязки к счёту', '=C8')):
    put(ws, r, [lab, f, '=B%d/$C$9' % r], s_r2); r += 1
r += 1
put(ws, r, ['Сверка с кэшфло Odoo (account.payment)'], s_sec); r += 1
put(ws, r, ['Показатель', 'Сумма, сум', 'Комментарий'], s_h2); r += 1
r0 = r
for lab, v, cm in (('Поступило от клиентов', '=C9', 'Входящие платежи клиентов'),
                   ('Исходящие платежи', OUT, 'Комиссии Payme, Click, Paynet и возвраты; в кэшфло Odoo — колонка «Нет»'),
                   ('Итого как в кэшфло Odoo', '=B%d+B%d' % (r0, r0 + 1), 'Должно совпадать с итогом пивота cashflow за тот же период'),
                   ('Справочно: выставлено счетов (по дате счёта)', ACC, 'Начисленная выручка; прежний способ расчёта этого файла'),
                   ('Разница: поступило минус выставлено', '=B%d-B%d' % (r0, r0 + 3), 'Сдвиг оплаты относительно даты счёта, авансы и неоплаченные счета')):
    put(ws, r, [lab, v, cm], s_r2); r += 1
r += 1
for t in (SRC, 'Выручка = поступившие от клиентов деньги в месяце поступления (кассовый метод), до вычета комиссий платёжных систем. Валютные поступления — в сумах по курсу проведения.',
          '«Без привязки к счёту» — деньги пришли, но ещё не зачтены ни к одному счёту: авансы под будущие счета или оплата без счёта. После зачёта они перейдут в свой тип.',
          'Месяц %s учтён полностью — по %s включительно.' % (LAST_RU.lower(), CUT_RU)):
    put(ws, r, [t], s_n); r += 1
# ---------- По продуктам ----------
ws = wb['По продуктам']; n = 8
s_h, s_r, s_t, s_n = st(ws, 3, n), st(ws, 4, n), st(ws, 21, n), st(ws, 23, n)
wipe(ws, 3, n)
ws['A1'] = 'Выручка по отдельным продуктам — по дате поступления денег, %s' % PERIOD_LONG
put(ws, 3, ['Тип', 'Продукт', 'Оплат (платежей)', 'Количество', 'Поступило, сум', 'Доля в типе', 'Доля во всех поступлениях', 'Средняя оплата за единицу'], s_h)
TN = {t[0]: t[1] for t in TYPES}
pids = [p for p in ORDER if p in P] + sorted([p for p in P if p not in ORDER and p is not None]) + ([None] if None in P else [])
rows = [(TN[cat(p)], LAB.get(p) or (d['names'].get(str(p)) if p else 'Без продукта в строке'), P[p][2], P[p][1], P[p][0], G[cat(p)]) for p in pids if abs(P[p][0]) > 0.5]
for sp, lab in (('SFA', 'Без счёта — клиенты SFA'), ('POS', 'Без счёта — клиенты POS'), ('Нет', 'Без счёта — сервер не указан')):
    if sp in U: rows.append((TN['UN'], lab, U[sp][1], None, U[sp][0], G['UN']))
r = 4
for t, lab, cnt, q, s, g in rows:
    put(ws, r, [t, lab, cnt, round(q, 1) if q is not None else None, s, (s / g) if g else 0, '=E%d/$E$%d' % (r, 4 + len(rows)), '=IF(AND(ISNUMBER(D%d),D%d>0),E%d/D%d,"")' % (r, r, r, r)], s_r); r += 1
put(ws, r, [None, 'ИТОГО', None, None, '=SUM(E4:E%d)' % (r - 1), None, '=E%d/$E$%d' % (r, r), None], s_t); r += 2
for t in (SRC, '«Оплат» — число платежей, в которых есть этот продукт; один платёж может покрывать несколько продуктов. «Количество» — лицензии или услуги пропорционально оплаченной доле счёта.',
          'Валютные счета (SFA USD) теперь в сумах: в кассовом учёте сумма берётся из платежа.'):
    put(ws, r, [t], s_n); r += 1
# ---------- Помесячно ----------
ws = wb['Помесячно']; n = max(12, NM + 3)
s_h, s_r, s_t, s_sec, s_p, s_n = st(ws, 3, 12), st(ws, 4, 12), st(ws, 8, 12), st(ws, 10, 12), st(ws, 12, 12), st(ws, 17, 12)
def ext(sty):  # растянуть стили на число месяцев
    mid = sty[1]; return [sty[0]] + [copy.copy(mid) for _ in range(NM)] + [sty[10], sty[11]]
s_h, s_r, s_t, s_p = ext(s_h), ext(s_r), ext(s_t), ext(s_p); s_tp = ext(st(ws, 12, 12))
wipe(ws, 3, n)
ws['A1'] = 'Выручка по типам помесячно — по дате поступления денег'
cK, cS = 2 + NM, 3 + NM; LK, LS = L(cK), L(cS); LE = L(1 + NM)
put(ws, 3, ['Тип выручки'] + ['%s 2026' % RUS[m] for m in MON] + ['ИТОГО', 'Доля'], s_h)
r = 4
for k, name, *_ in TYPES:
    put(ws, r, [name] + [Mt.get((k, m), 0.0) for m in MON] + ['=SUM(B%d:%s%d)' % (r, LE, r), '=%s%d/$%s$9' % (LK, r, LK)], s_r); r += 1
put(ws, 9, ['ИТОГО поступило'] + ['=SUM(%s4:%s8)' % (L(c), L(c)) for c in range(2, cK)] + ['=SUM(B9:%s9)' % LE, '=%s9/%s9' % (LK, LK)], s_t)
put(ws, 10, ['Исходящие платежи (комиссии)'] + [d['out'].get(m, 0.0) for m in MON] + ['=SUM(B10:%s10)' % LE, None], s_r)
put(ws, 11, ['Итого как в кэшфло Odoo'] + ['=%s9+%s10' % (L(c), L(c)) for c in range(2, cK)] + ['=SUM(B11:%s11)' % LE, None], s_t)
put(ws, 12, ['Справочно: выставлено счетов'] + [d['accr'].get(m, 0.0) for m in MON] + ['=SUM(B12:%s12)' % LE, None], s_r)
put(ws, 14, ['Доля типа в поступлениях месяца'], s_sec)
put(ws, 15, ['Тип выручки'] + ['%s 2026' % RUS[m] for m in MON], s_h[:1 + NM])
r = 16
for i, (k, name, *_) in enumerate(TYPES):
    put(ws, r, [name] + ['=%s%d/%s$9' % (L(c), 4 + i, L(c)) for c in range(2, cK)], s_p[:1 + NM]); r += 1
r += 1
for t in (SRC, '«Итого как в кэшфло Odoo» совпадает с пивотом account.payment (дата, статус «проведён»): поступления минус исходящие платежи.',
          'Прошлые месяцы могут меняться: когда авансы зачитывают к счетам, деньги переходят из «Без привязки к счёту» в свой тип.'):
    put(ws, r, [t], s_n); r += 1
for c in range(2, cK): ws.column_dimensions[L(c)].width = 15
ws.column_dimensions[LK].width = 16; ws.column_dimensions[LS].width = 9
fn = OUTP('Выручка_по_продуктам_и_типам_%s_2026_%s.xlsx' % (PERIOD.replace('–', '-'), SUF))
wb.save(fn)
print(fn, '| поступило %.1f млн: ' % (CIN / 1e6) + ', '.join('%s %.1f' % (t[1], G[t[0]] / 1e6) for t in TYPES), '| исходящие %.1f | как в пивоте %.1f' % (OUT / 1e6, (CIN + OUT) / 1e6))
