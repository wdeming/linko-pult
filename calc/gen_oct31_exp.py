# -*- coding: utf-8 -*-
"""Расходы: по дате оплаты (кэшфло) и за период, к которому относятся (Excel) + сравнение с поступлениями и реестр."""
import json, collections, datetime, calendar, openpyxl
from openpyxl.styles import Font, Alignment
from style_lib import *
from cfg_cut import *
d = json.load(open(J('exp'), encoding='utf-8')); cash = json.load(open(J('cash'), encoding='utf-8'))
L = openpyxl.utils.get_column_letter; D = datetime.date.fromisoformat
START = '2026-04'                        # с апреля расходы в Odoo ведутся помесячно; до этого — сводные записи 30–31.03
MS = [m for m in MON if m >= START]; EARLY = 'Ранее (до апреля 2026)'
COLS = [EARLY] + MS
bucket = lambda m: EARLY if m < START else m
def split(x):
    """Сумма расхода по месяцам периода — пропорционально дням."""
    a, b = D(x['d1']), D(x['d2']); n = (b - a).days + 1; res = collections.defaultdict(float); cur = a
    while cur <= b:
        me = min(b, cur.replace(day=calendar.monthrange(cur.year, cur.month)[1])); k = (me - cur).days + 1
        res[cur.strftime('%Y-%m')] += x['amt'] * k / n; cur = me + datetime.timedelta(days=1)
    return res
paid = [x for x in d if x['pay'] <= CUT]
CA = collections.defaultdict(float); PE = collections.defaultdict(float); FUT = collections.defaultdict(float); LATE = collections.defaultdict(float)
PREPAID = collections.defaultdict(float)   # оплачено до среза за месяцы после среза
for x in paid: CA[(bucket(x['pay'][:7]), x['plan'], x['cat'])] += x['amt']
for x in d:
    for m, v in split(x).items():
        if m > LAST_M:
            FUT[(x['plan'], x['cat'])] += v
            if x['pay'] <= CUT: PREPAID[(x['plan'], x['cat'])] += v
        else:
            PE[(bucket(m), x['plan'], x['cat'])] += v
            if m >= START and x['pay'][:7] > m: LATE[m] += v
plans = collections.defaultdict(float)
for (m, p, c), v in list(CA.items()) + list(PE.items()): plans[p] += v
PLANS = sorted(plans, key=lambda p: -plans[p])
CATS = {p: sorted({c for (_, pp, c) in list(CA) + list(PE) if pp == p} | {c for (pp, c) in FUT if pp == p},
                  key=lambda c: -sum(v for (m, pp, cc), v in list(CA.items()) + list(PE.items()) if pp == p and cc == c)) for p in PLANS}
wb = openpyxl.Workbook(); wb.remove(wb.active)
def block(ws, A, extra=None):
    """Строки: вид расхода (подытог) и статьи; колонки: ранее, месяцы, итого (+ доп. колонка)."""
    n = len(COLS); cT = L(2 + n); r = 4; subs = []
    hd = ['Вид / статья расхода'] + [c if c == EARLY else '%s 2026' % RUS[c] for c in COLS] + ['ИТОГО ' + ('по %s' % CUT_RU)]
    if extra: hd.append(extra[0])
    header(ws, 3, hd)
    for p in PLANS:
        cats = CATS[p]; r0 = r; subs.append(r0)
        vals = [p] + ['=SUM(%s%d:%s%d)' % (L(c), r0 + 1, L(c), r0 + len(cats)) for c in range(2, 3 + n)]
        if extra: vals.append('=SUM(%s%d:%s%d)' % (L(3 + n), r0 + 1, L(3 + n), r0 + len(cats)))
        datarow(ws, r, vals, fill=GREY, fmts=[None] + [NUM] * (n + 2), bold=True, blue_inputs=False); r += 1
        for c in cats:
            vals = ['    ' + c] + [round(A.get((m, p, c), 0)) or None for m in COLS] + ['=SUM(B%d:%s%d)' % (r, L(1 + n), r)]
            if extra: vals.append(round(extra[1].get((p, c), 0)) or None)
            datarow(ws, r, vals, fmts=[None] + [NUM] * (n + 2)); r += 1
    vals = ['ИТОГО'] + ['=%s' % '+'.join('%s%d' % (L(c), s) for s in subs) for c in range(2, 3 + n + (1 if extra else 0))]
    totalrow(ws, r, vals, fmts=[None] + [NUM] * (n + 2))
    widths(ws, {'A': 44, **{L(c): 15 for c in range(2, 4 + n)}}); ws.freeze_panes = 'B4'
    return r
# ---------- лист 1: кэшфло ----------
ws = wb.create_sheet('Расходы по дате оплаты')
title(ws, 'A1', 'Расходы по дате оплаты (кэшфло), %s (на %s)' % (PERIOD_LONG, CUT_RU))
r = block(ws, CA)
notes(ws, r + 2, [
  'Источник: Odoo, реестр «Расходы» (x_custom_expenses) — все расходы с датой оплаты по %s включительно. Выгрузка %s.' % (CUT_RU, EXPORT_RU),
  'Месяц — по «Дате расхода» (когда деньги ушли), независимо от того, за какой период расход.',
  'Сумма — сумма строк расхода. Расходы в долларах (с курсом в карточке) пересчитаны по 12 000 сум/$, как во всём комплекте.',
  'Помесячный учёт расходов в Odoo начат в апреле 2026: в колонке «%s» — сводные записи 30–31.03.2026 («расходы за период», в основном «Прочие расходы»).' % EARLY])
# ---------- лист 2: за период ----------
ws = wb.create_sheet('Расходы за период')
title(ws, 'A1', 'Расходы за период, к которому они относятся, %s (на %s)' % (PERIOD_LONG, CUT_RU))
r = block(ws, PE, ('Оплачено авансом за месяцы после %s' % LAST_RU_L, PREPAID))
notes(ws, r + 2, [
  'Месяц — по периоду расхода («Период начало» – «Период конец»): если период захватывает несколько месяцев, сумма делится пропорционально дням.',
  'Расход без периода (не «оплата за период») относится к месяцу оплаты.',
  'Сюда входят и расходы, оплаченные после %s (по дату выгрузки %s), если их период относится к месяцам до среза — например, оклад за %s, выплаченный в следующем месяце.' % (CUT_RU, EXPORT_RU, LAST_RU_L),
  'Последняя колонка — оплачено до %s за месяцы после среза (подписки, аренда авансом); в итог по месяцам не входит.' % CUT_RU,
  'Последний месяц может быть неполным: расходы за него, которые ещё не оплачены и не внесены в Odoo, сюда не попали.'])
# ---------- лист 3: сравнение ----------
ws = wb.create_sheet('Сравнение и денежный поток')
title(ws, 'A1', 'Расходы по оплате и за период, поступления и денежный поток по месяцам (на %s)' % CUT_RU)
header(ws, 3, ['Месяц', 'Расходы по дате оплаты', 'Расходы за период', 'Разница (оплата − период)', 'Расходов за период, оплачено позже',
               'Поступило денег', 'Комиссии Payme / Click / Paynet', 'Чистый денежный поток (поступления − комиссии − расходы)'])
inb = collections.defaultdict(float); com = collections.defaultdict(float)
for x in cash:
    if x['pt'] == 'inbound': inb[x['m']] += x['s']
    else: com[x['m']] += -x['s']
r = 4; CMP = {}
for m in MS:
    ca = sum(v for (mm, _, _), v in CA.items() if mm == m); pe = sum(v for (mm, _, _), v in PE.items() if mm == m)
    CMP[m] = {'ca': ca, 'pe': pe, 'late': LATE[m], 'in': inb[m], 'com': com[m]}
    datarow(ws, r, ['%s 2026' % RUS[m], round(ca), round(pe), '=B%d-C%d' % (r, r), round(LATE[m]), round(inb[m]), round(-com[m]), '=F%d+G%d-B%d' % (r, r, r)], fmts=[None] + [NUM] * 7); r += 1
totalrow(ws, r, ['ИТОГО апр–%s' % LAST_SH.lower()] + ['=SUM(%s4:%s%d)' % (L(c), L(c), r - 1) for c in range(2, 9)], fmts=[None] + [NUM] * 7)
notes(ws, r + 2, ['Поступления и комиссии — из отчёта «Поступления денег» (account.payment по дате оплаты).',
                  'Чистый денежный поток — деньги, которые пришли, минус комиссии платёжных систем и расходы по дате оплаты. Без учёта переводов между кассами.',
                  'Разница «оплата − период»: плюс — заплатили больше, чем относится к месяцу (авансы, долги прошлых месяцев); минус — часть расходов месяца оплачена позже.'])
widths(ws, {'A': 18, **{L(c): 19 for c in range(2, 9)}}); ws.freeze_panes = 'B4'
for c in range(1, 9): ws.cell(3, c).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ws.row_dimensions[3].height = 46
# ---------- лист 4: реестр ----------
ws = wb.create_sheet('Реестр расходов')
title(ws, 'A1', 'Реестр расходов из Odoo (выгрузка %s)' % EXPORT_RU)
header(ws, 3, ['ID в Odoo', 'Дата оплаты', 'Период с', 'Период по', 'Вид расхода', 'Статья', 'Касса', 'Сумма, сум', 'В долларах', 'Курс в Odoo', 'Примечание'])
r = 4
for x in sorted(d, key=lambda x: (x['pay'], x['id'])):
    datarow(ws, r, [x['id'], D(x['pay']), D(x['d1']), D(x['d2']), x['plan'], x['cat'], x['kassa'], x['amt'], x['usd'] or None, x['rate'] or None, x['note']],
            fmts=[None, 'DD.MM.YYYY', 'DD.MM.YYYY', 'DD.MM.YYYY', None, None, None, NUM, '#,##0.00', NUM, None], blue_inputs=False); r += 1
totalrow(ws, r, ['ИТОГО', None, None, None, None, None, None, '=SUM(H4:H%d)' % (r - 1), '=SUM(I4:I%d)' % (r - 1)], fmts=[None] * 7 + [NUM, '#,##0.00'])
widths(ws, {'A': 10, 'B': 12, 'C': 12, 'D': 12, 'E': 26, 'F': 30, 'G': 18, 'H': 16, 'I': 12, 'J': 11, 'K': 60})
ws.freeze_panes = 'A4'; ws.auto_filter.ref = 'A3:K%d' % (r - 1)
fn = OUTP('Расходы_кэшфло_и_за_период_%s_2026_%s.xlsx' % (PERIOD.replace('–', '-'), SUF))
wb.save(fn)
k3 = lambda A: {'%s|%s|%s' % kk: v for kk, v in A.items()}
json.dump({'CA': k3(CA), 'PE': k3(PE), 'CMP': CMP, 'PLANS': PLANS, 'PREPAID': sum(PREPAID.values()), 'FUT': sum(FUT.values()), 'EARLY': EARLY}, open(J('res_exp'), 'w'), ensure_ascii=False)
print(fn, '| оплачено', round(sum(CA.values()) / 1e6, 1), 'за период', round(sum(PE.values()) / 1e6, 1), 'аванс', round(sum(PREPAID.values()) / 1e6, 1))
for m in MS: print(m, {k: round(v / 1e6, 1) for k, v in CMP[m].items()})
