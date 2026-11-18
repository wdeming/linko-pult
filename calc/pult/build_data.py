# -*- coding: utf-8 -*-
"""Собирает данные пульта из pult/work/<срез>/: out/metrics.json, out/series.json, out/bundle_<срез>.json (Excel в base64)."""
import os, json, glob, base64, collections, datetime
P = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(P); OUT = os.path.join(P, 'out'); os.makedirs(OUT, exist_ok=True)
TAG = {8: 'aug31', 9: 'sep30', 10: 'oct31', 11: 'nov30', 12: 'dec31'}
RM = ['', 'янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек']
ml = lambda m: '%s %s' % (RM[int(m[5:7])], m[2:4])
cuts = sorted(os.path.basename(d) for d in glob.glob(os.path.join(P, 'work', '20*')) if os.path.exists(os.path.join(d, 'stamp.json')))
PP, F = 'POS', 'SFA'
DOCS = {'mrr': 'MRR_и_когортное_удержание_2026_на_%s.xlsx', 'churn': 'Отток_подписок_причины_%s_2026_на_%s.xlsx', 'clients': 'Клиентская_база_продукты_сегменты_на_%s.xlsx',
        'multi': 'Клиенты_с_несколькими_подписками_на_%s.xlsx', 'top20': 'Концентрация_выручки_и_условия_топ20_на_%s.xlsx', 'cost': 'Себестоимость_поддержки_внедрения_инфраструктуры_на_%s.xlsx',
        'metrics': 'Ключевые_метрики_SaaS_на_%s_с_разовой.xlsx', 'calls': 'Обзвоны_клиентов_сырая_выгрузка_Odoo_на_%s.xlsx',
        'cash': 'Поступления_денег_%s_2026_на_%s.xlsx', 'exp': 'Расходы_кэшфло_и_за_период_%s_2026_на_%s.xlsx'}
MANUAL = {'org': 'Оргструктура_Linktrade_на_%s.xlsx', 'revenue': 'Выручка_по_продуктам_и_типам_%s_2026_на_%s.xlsx', 'debts': 'Все_долги_на_%s_ЧЕРНОВИК.xlsx'}
metrics, series = [], []
def revenue(cut, ddmm, per):
    import openpyxl
    months = ['2026-%02d' % i for i in range(2, int(cut[5:7]) + 1)]
    for c2 in sorted([c for c in cuts if c >= cut], reverse=True):   # берём самый свежий файл выручки, где есть нужные месяцы
        fn = os.path.join(ROOT, MANUAL['revenue'] % ('Фев-' + RM[int(c2[5:7])].capitalize(), '%s.%s' % (c2[8:10], c2[5:7])))
        if not os.path.exists(fn): continue
        ws = openpyxl.load_workbook(fn)['Помесячно']; res = {}
        try:
            for r in range(4, 8):
                res[ws.cell(r, 1).value] = {m: float(ws.cell(r, 2 + i).value) for i, m in enumerate(months)}
            return res
        except (TypeError, ValueError): continue
    return None
def debts(ddmm):
    fn = os.path.join(ROOT, MANUAL['debts'] % ddmm)
    if not os.path.exists(fn): return None
    import openpyxl
    ws = openpyxl.load_workbook(fn)['Долги']; res = []
    for r in range(2, 20):
        a, c, d = ws.cell(r, 1).value, ws.cell(r, 3).value, ws.cell(r, 4).value
        if a and a != 'ИТОГО' and isinstance(c, (int, float)): res.append((str(a).strip(), float(c), float(d or 0)))
    return res
for cut in cuts:
    dn = os.path.join(P, 'work', cut); t = TAG[int(cut[5:7])]; m = cut[:7]
    L = lambda n: json.load(open(os.path.join(dn, 'x_%s_%s.json' % (t, n)), encoding='utf-8'))
    ddmm = '%s.%s' % (cut[8:10], cut[5:7]); per = 'Фев-' + RM[int(cut[5:7])].capitalize()
    R, K, M, C, base, const, calls, fot, surv = L('res_metrics'), L('cost'), L('res_mrr'), L('res_churn'), L('base'), L('const'), L('res_calls'), L('fot'), L('survival')
    stamp = json.load(open(os.path.join(dn, 'stamp.json')))
    def add(doc, key, label, value, kind='n', good='', prod=''):
        metrics.append({'id': '%s.%s' % (cut, key), 'cut': cut, 'doc': doc, 'key': key, 'label': label, 'value': None if value is None else round(float(value), 6), 'kind': kind, 'good': good, 'prod': prod})
    def ser(chart, x, s, v):
        series.append({'id': '%s.%s.%s.%s' % (cut, chart, x, s), 'cut': cut, 'chart': chart, 'x': x, 's': s, 'v': None if v is None else round(float(v), 6)})
    mov = lambda e, p: M['mov'].get('%s|%s|%s' % (p, m, e), 0)
    tm = R['MRR'][PP] + R['MRR'][F]; tmj = tm
    # --- сводка и MRR
    add('mrr', 'mrr_total', 'MRR всего', tm, 'm', 'up'); add('mrr', 'mrr_sfa', 'MRR SFA', R['MRR'][F], 'm', 'up', F); add('mrr', 'mrr_pos', 'MRR POS', R['MRR'][PP], 'm', 'up', PP)
    add('mrr', 'subs', 'Подписок в базе', len(base), 'n', 'up'); add('mrr', 'subs_sfa', 'Подписок SFA', R['SRV'][F], 'n', 'up', F); add('mrr', 'subs_pos', 'Подписок POS', R['SRV'][PP], 'n', 'up', PP)
    # потоки MRR: всего, SFA, POS
    for key, lab, e, sign, good in (('new', 'Новые продажи за месяц', '0_creation', 1, 'up'), ('exp', 'Расширения за месяц', '1_expansion', 1, 'up'),
                                    ('con', 'Сокращения за месяц', '15_contraction', -1, 'down'), ('chu', 'Отток за месяц', '2_churn', -1, 'down')):
        add('mrr', key + '_total', lab + ', всего', sign * (mov(e, F) + mov(e, PP)), 'm', good)
        for p in (F, PP): add('mrr', key + '_' + p, '%s, %s' % (lab, p), sign * mov(e, p), 'm', good, p)
    st = {p: M['GN']['%s|%s' % (p, m)]['start'] for p in (F, PP)}; S0 = st[F] + st[PP]
    lost = sum(mov('2_churn', p) + mov('15_contraction', p) for p in (F, PP)); gain = sum(mov('1_expansion', p) for p in (F, PP))
    add('mrr', 'grr_total', 'GRR месяца, всего', (S0 + lost) / S0, 'p', 'up')
    for p in (F, PP): add('mrr', 'grr_' + p, 'GRR месяца, %s' % p, R['GRR'][p], 'p', 'up', p)
    add('mrr', 'nrr_total', 'NRR месяца, всего', (S0 + lost + gain) / S0, 'p', 'up')
    for p in (F, PP): add('mrr', 'nrr_' + p, 'NRR месяца, %s' % p, R['NRR'][p], 'p', 'up', p)
    for p in (F, PP):
        for x in R['series'][p]:
            ser('mrr_end', ml(x['m']), p, x['end'])
            for k, nm in (('new', 'Новые'), ('exp', 'Расширения'), ('con', 'Сокращения'), ('chu', 'Отток')): ser('flow_' + p, ml(x['m']), nm, x[k])
        for k, g in M['GN'].items():
            if k.startswith(p + '|'): ser('retention', ml(k.split('|')[1]), 'GRR ' + p, g['grr']); ser('retention', ml(k.split('|')[1]), 'NRR ' + p, g['nrr'])
        for a in surv[p]: ser('survival', '%d мес' % a[0], p, a[3])
    for mm in sorted({k.split('|')[1] for k in M['GN']}):
        g = [M['GN'].get('%s|%s' % (p, mm)) for p in (F, PP)]
        if all(g):
            s0 = sum(x['start'] for x in g)
            ser('retention', ml(mm), 'GRR всего', sum(x['start'] * x['grr'] for x in g) / s0); ser('retention', ml(mm), 'NRR всего', sum(x['start'] * x['nrr'] for x in g) / s0)
    sv = {}
    for p in (F, PP):
        for a in surv[p]: sv.setdefault(a[0], [0, 0]); sv[a[0]][0] += a[1]; sv[a[0]][1] += a[2]
    for age, (al, n) in sorted(sv.items()):
        if n: ser('survival', '%d мес' % age, 'Всего', al / n)
    # --- отток
    cats = collections.defaultdict(float)
    for k, v in C['cats'].items(): cats[k.split('|')[0]] += v[1]
    tc = sum(cats.values())
    d2 = {p: C['DYN']['%s|%s' % (p, m)] for p in (F, PP)}
    add('churn', 'closed_total', 'Закрыто подписок за месяц, всего', d2[F]['n'] + d2[PP]['n'], 'n', 'down')
    for p in (F, PP): add('churn', 'closed_' + p, 'Закрыто подписок за месяц, %s' % p, d2[p]['n'], 'n', 'down', p)
    add('churn', 'churnpct_total', 'Отток MRR за месяц, всего', (d2[F]['loss'] + d2[PP]['loss']) / (d2[F]['base_mrr'] + d2[PP]['base_mrr']), 'p', 'down')
    for p in (F, PP): add('churn', 'churnpct_' + p, 'Отток MRR за месяц, %s' % p, d2[p]['loss'] / d2[p]['base_mrr'], 'p', 'down', p)
    add('churn', 'newn_total', 'Новых подписок за месяц, всего', d2[F]['new'] + d2[PP]['new'], 'n', 'up')
    for p in (F, PP): add('churn', 'newn_' + p, 'Новых подписок за месяц, %s' % p, d2[p]['new'], 'n', 'up', p)
    for p in (F, PP):
        for k, d in C['DYN'].items():
            if k.startswith(p + '|'):
                ser('churn_closed', ml(k.split('|')[1]), p, d['n']); ser('churn_new', ml(k.split('|')[1]), p, d['new'])
                ser('churn_pct', ml(k.split('|')[1]), p, d['loss'] / d['base_mrr'])
    for mm in sorted({k.split('|')[1] for k in C['DYN']}):
        g = [C['DYN'].get('%s|%s' % (p, mm)) for p in (F, PP)]
        if all(g): ser('churn_pct', ml(mm), 'Всего', sum(x['loss'] for x in g) / sum(x['base_mrr'] for x in g))
    add('churn', 'closed_cum', 'Закрыто с 01.02 накопленно', C['tot'][PP][0] + C['tot'][F][0], 'n', 'down'); add('churn', 'lost_cum', 'Потерянный MRR с 01.02', C['tot'][PP][1] + C['tot'][F][1], 'm', 'down')
    for k, v in sorted(cats.items(), key=lambda x: -x[1]): ser('churn_cats', k, 'Доля в потерянном MRR', v / tc)
    # --- клиенты, мульти, концентрация
    cl = collections.defaultdict(lambda: {'mrr': 0.0, 'prods': set(), 'n': 0, 'subs': []})
    for b in base:
        c = cl[b['partner_id']]; c['mrr'] += b['mrr']; c['prods'].add(b['prod']); c['n'] += 1; c['subs'].append(b); c['name'] = b['pname']
    seg = {s: [sum(1 for c in cl.values() if f(c)), sum(c['mrr'] for c in cl.values() if f(c))] for s, f in
           (('Только SFA', lambda c: c['prods'] == {F}), ('Только POS', lambda c: c['prods'] == {PP}), ('Оба продукта', lambda c: len(c['prods']) == 2))}
    add('clients', 'clients', 'Клиентов всего', len(cl), 'n', 'up'); add('clients', 'cl_sfa', 'Только SFA', seg['Только SFA'][0], 'n', 'up', F); add('clients', 'cl_pos', 'Только POS', seg['Только POS'][0], 'n', 'up', PP)
    add('clients', 'cl_both', 'Оба продукта', seg['Оба продукта'][0], 'n', 'up'); add('clients', 'arpc', 'Средний MRR на клиента', tm / len(cl), 'm', 'up')
    sp = const['SPLIT']; add('clients', 'paused_n', 'Подписок на паузе', sp[PP]['4_paused'][0] + sp[F]['4_paused'][0], 'n', 'down'); add('clients', 'paused_mrr', 'MRR на паузе', sp[PP]['4_paused'][1] + sp[F]['4_paused'][1], 'm', 'down')
    add('clients', 'lic_sfa', 'Лицензий SFA', R['LIC'][F], 'n', 'up', F); add('clients', 'lic_pos', 'Лицензий POS', R['LIC'][PP], 'n', 'up', PP)
    for s, v in seg.items(): ser('seg_clients', s, 'Клиентов', v[0]); ser('seg_mrr', s, 'MRR', v[1])
    mu = [c for c in cl.values() if c['n'] >= 2]; mm = sum(c['mrr'] for c in mu)
    add('multi', 'mu_clients', 'Клиентов с 2+ подписками', len(mu), 'n', 'up'); add('multi', 'mu_subs', 'Подписок у них', sum(c['n'] for c in mu), 'n', 'up'); add('multi', 'mu_mrr', 'MRR этих клиентов', mm, 'm', 'up'); add('multi', 'mu_share', 'Доля в MRR базы', mm / tm, 'p', '')
    for c in sorted(mu, key=lambda c: -c['mrr'])[:12]: ser('multi_top', c['name'][:28], 'MRR', c['mrr'])
    rk = sorted(cl.values(), key=lambda c: -c['mrr']); acc = []; s = 0
    for c in rk: s += c['mrr']; acc.append(s)
    for k in (1, 5, 10, 20, 50, 100):
        add('top20', 'top%d' % k, 'Доля топ-%d в MRR' % k, acc[k - 1] / tm, 'p', 'down'); ser('top_share', 'Топ-%d' % k, 'Доля в MRR', acc[k - 1] / tm)
    add('top20', 'p80', '80% MRR набирается на клиенте №', next(i + 1 for i, a in enumerate(acc) if a >= 0.8 * tm), 'n', 'up')
    cd = datetime.date.fromisoformat(cut)
    risk = [c for c in rk[:20] if any(x['end_date'] and datetime.date.fromisoformat(x['end_date']) <= cd + datetime.timedelta(days=14) for x in c['subs']) or any(x['st'] == '4_paused' for x in c['subs'])]
    add('top20', 'risk_n', 'Клиентов топ-20 под риском', len(risk), 'n', 'down'); add('top20', 'risk_mrr', 'MRR топ-20 под риском', sum(c['mrr'] for c in risk), 'm', 'down')
    for c in rk[:20]: ser('top_clients', c['name'][:28], 'MRR', c['mrr'])
    # --- себестоимость
    ct = K['COGS'][PP] + K['COGS'][F]
    add('cost', 'cost_total', 'Себестоимость в месяц', ct, 'm', 'down'); add('cost', 'cost_pct', 'Себестоимость, % от MRR', ct / tmj, 'p', 'down'); add('cost', 'cost_sfa', 'На SFA', K['COGS'][F], 'm', 'down', F); add('cost', 'cost_pos', 'На POS', K['COGS'][PP], 'm', 'down', PP)
    for k, nm in (('SUP', 'Техподдержка'), ('ONP', 'Onboarding POS'), ('ONS', 'Onboarding SFA'), ('KA', 'Ключевые клиенты'), ('HOST_AVG', 'Хостинг (среднее)')):
        add('cost', 'c_' + k, nm, K[k], 'm', 'down'); ser('cost_parts', nm, 'Сумма в месяц', K[k])
    add('cost', 'host_m', 'Оплаты хостинга за месяц', K['HOST_M'].get(m), 'm', 'down')
    for mmm, v in K['HOST_M'].items(): ser('hosting', ml(mmm), 'Оплаты хостинга', v)
    # --- ключевые метрики
    for lab, k, kind, good in (('ARPA', 'ARPA', 'm', 'up'), ('Gross Margin по подписке', 'GM_SUB', 'p', 'up'), ('Отток по деньгам в месяц', 'CH_CLOSED', 'p', 'down'), ('Срок жизни, мес', 'LIFE', 'd1', 'up'),
                                ('LTV', 'LTV', 'm', 'up'), ('CAC с внедрением', 'CAC', 'm', 'down'), ('LTV : CAC', 'LTVCAC', 'd2', 'up'), ('Payback, мес', 'PAYBACK', 'd1', 'down')):
        for p in (F, PP): add('metrics', '%s_%s' % (k.lower(), p), '%s, %s' % (lab, p), R[k][p], kind, good, p)
    for p in (F, PP):
        ser('unit', 'LTV', p, R['LTV'][p]); ser('unit', 'CAC', p, R['CAC'][p])
    # --- обзвоны
    b = calls['bym'].get(m, {}); n = sum(b.values()) or 1
    add('calls', 'calls_n', 'Звонков за месяц', sum(b.values()), 'n', 'up'); add('calls', 'calls_pr', 'Из них «Проблема»', b.get('Проблема', 0), 'n', 'down'); add('calls', 'calls_nd', 'Из них «Не дозвонились»', b.get('Не дозвонились', 0), 'n', 'down')
    add('calls', 'calls_share', 'Доля проблемных', b.get('Проблема', 0) / n, 'p', 'down'); add('calls', 'calls_cum', 'Записей накопленно', calls['N'], 'n', '')
    for mmm, v in calls['bym'].items():
        for k in ('Всё ок', 'Проблема', 'Не дозвонились'): ser('calls_m', ml(mmm), k, v.get(k, 0))
    # --- поступления денег
    cp = os.path.join(dn, 'x_%s_res_cash.json' % t)
    if os.path.exists(cp):
        CS = json.load(open(cp, encoding='utf-8')); PCS = collections.defaultdict(float); MCS = collections.defaultdict(float)
        for k2, v in CS['P'].items(): PCS[tuple(k2.split('|'))] += v
        for k2, v in CS['M'].items(): MCS[tuple(k2.split('|'))] += v
        cin = sum(v for (mm, _), v in PCS.items() if mm == m); com = CS['C'].get(m, 0)
        add('cash', 'cash_in', 'Поступило за месяц, всего', cin, 'm', 'up')
        add('cash', 'cash_SFA', 'Поступило за месяц, SFA', PCS.get((m, 'SFA'), 0), 'm', 'up', F)
        add('cash', 'cash_POS', 'Поступило за месяц, POS', PCS.get((m, 'POS'), 0), 'm', 'up', PP)
        add('cash', 'cash_comm', 'Комиссии Payme / Click / Paynet за месяц', com, 'm', 'down')
        add('cash', 'cash_net', 'Чистые поступления за месяц', cin - com, 'm', 'up')
        add('cash', 'cash_comm_pct', 'Комиссия, % от поступлений', com / cin if cin else None, 'p', 'down')
        add('cash', 'cash_cum', 'Поступило с 01.02 накопленно', sum(PCS.values()), 'm', 'up')
        months = sorted({mm for mm, _ in PCS})
        prods = [p for p in CS['PR'] if any(PCS.get((mm, p)) for mm in months)]
        for mm in months:
            for pr in prods: ser('cash_prod', ml(mm), pr, PCS.get((mm, pr), 0))
            for me in CS['METH']: ser('cash_meth', ml(mm), me, MCS.get((mm, me), 0))
            ser('cash_comm', ml(mm), 'Комиссии', CS['C'].get(mm, 0))
    # --- расходы: по дате оплаты и за период
    ep = os.path.join(dn, 'x_%s_res_exp.json' % t)
    if os.path.exists(ep):
        EX = json.load(open(ep, encoding='utf-8')); A3 = lambda D: {tuple(k3.split('|')): v for k3, v in D.items()}
        CA, PE = A3(EX['CA']), A3(EX['PE']); cm = EX['CMP'].get(m, {})
        add('exp', 'exp_cash', 'Расходы за месяц по дате оплаты', cm.get('ca'), 'm', 'down')
        add('exp', 'exp_period', 'Расходы за месяц за период', cm.get('pe'), 'm', 'down')
        add('exp', 'exp_late', 'Из расходов за период оплачено позже', cm.get('late'), 'm', '')
        add('exp', 'exp_prepaid', 'Оплачено авансом за следующие месяцы', EX['PREPAID'], 'm', '')
        if cm: add('exp', 'flow_net', 'Чистый денежный поток за месяц', cm['in'] - cm['com'] - cm['ca'], 'm', 'up')
        for mm in sorted(EX['CMP']):
            for pl in EX['PLANS']:
                ser('exp_cash', ml(mm), pl, sum(v for (a, b, _), v in CA.items() if a == mm and b == pl))
                ser('exp_period', ml(mm), pl, sum(v for (a, b, _), v in PE.items() if a == mm and b == pl))
            c2 = EX['CMP'][mm]
            ser('exp_flow', ml(mm), 'Поступило за вычетом комиссий', c2['in'] - c2['com']); ser('exp_flow', ml(mm), 'Расходы по дате оплаты', c2['ca'])
            ser('exp_flow', ml(mm), 'Чистый денежный поток', c2['in'] - c2['com'] - c2['ca'])
        cats = collections.defaultdict(lambda: [0.0, 0.0])
        for (a, _, c3), v in CA.items():
            if a == m: cats[c3][0] += v
        for (a, _, c3), v in PE.items():
            if a == m: cats[c3][1] += v
        for c3, (va, vb) in sorted(cats.items(), key=lambda z: -max(z[1])):
            if abs(va) + abs(vb) >= 1: ser('exp_cat', c3[:40], 'По дате оплаты', va); ser('exp_cat', c3[:40], 'За период', vb)
    # --- оргструктура
    add('org', 'staff', 'Штат, человек', sum(x['n'] for x in fot), 'n', ''); add('org', 'payroll', 'Оклады на руки в месяц', sum(x['sal'] for x in fot), 'm', 'down')
    add('org', 'payroll_share', 'Оклады, % от MRR', sum(x['sal'] for x in fot) / tmj, 'p', 'down')
    for x in fot: ser('staff_dep', (x['dep'] or '—')[:30], 'Человек', x['n']); ser('payroll_dep', (x['dep'] or '—')[:30], 'Оклады', x['sal'])
    # --- выручка и долги (из файлов комплекта, обновляются вручную)
    rv = revenue(cut, ddmm, per)
    if rv:
        tot = sum(v[m] for v in rv.values()); add('revenue', 'rev_total', 'Выручка за месяц', tot, 'm', 'up')
        for k, v in rv.items():
            add('revenue', 'rev_' + str(list(rv).index(k)), k, v[m], 'm', 'up')
            for mmm, val in v.items(): ser('revenue_m', ml(mmm), k, val)
    db = debts(ddmm)
    if db:
        add('debts', 'debt_total', 'Операционные долги (расчётно)', sum(c - d for _, c, d in db), 'm', 'down'); add('debts', 'debt_paid', 'Оплачено за два месяца', sum(d for _, c, d in db), 'm', 'up')
        for nm, c, d in db: ser('debts', nm, 'Было', c); ser('debts', nm, 'Остаток (расчётно)', c - d)
    # --- Excel
    files = {}
    for doc, pat in list(DOCS.items()) + list(MANUAL.items()):
        fn = pat % ((per, ddmm) if pat.count('%s') == 2 else ddmm)
        src = os.path.join(dn if doc in DOCS else ROOT, fn)
        if os.path.exists(src): files[doc] = {'name': fn, 'b64': base64.b64encode(open(src, 'rb').read()).decode(), 'at': stamp.get(doc) or datetime.datetime.utcfromtimestamp(os.path.getmtime(src)).strftime('%Y-%m-%dT%H:%M:%SZ')}
    json.dump({'cut': cut, 'files': files}, open(os.path.join(OUT, 'bundle_%s.json' % cut), 'w'), ensure_ascii=False)
    for doc, f in files.items(): metrics.append({'id': '%s.file_%s' % (cut, doc), 'cut': cut, 'doc': doc, 'key': '_file', 'label': f['name'], 'value': None, 'kind': 'file', 'good': f['at'], 'prod': ''})
# изменение к предыдущему срезу
by = {(r['cut'], r['key']): r for r in metrics if r['key'] != '_file'}
for r in metrics:
    if r['key'] == '_file': continue
    i = cuts.index(r['cut']); pv = by.get((cuts[i - 1], r['key'])) if i > 0 else None
    r['prev'] = pv['value'] if pv else None
    r['delta'] = None if (pv is None or pv['value'] is None or r['value'] is None) else round(r['value'] - pv['value'], 6)
code = open(os.path.abspath(__file__), encoding='utf-8').read()
meta = [{'label': 'build_data.py', 'value': code, 'type': 'code', 'lang': 'python'}, {'label': 'Источник', 'value': 'Odoo linktrade, только чтение; расчёт по скриптам комплекта data room'}]
json.dump({'rows': metrics, 'meta': meta}, open(os.path.join(OUT, 'metrics.json'), 'w'), ensure_ascii=False)
json.dump({'rows': series, 'meta': meta[1:]}, open(os.path.join(OUT, 'series.json'), 'w'), ensure_ascii=False)
print('срезы', cuts, '| показателей', len(metrics), '| точек', len(series), '|', {f: os.path.getsize(os.path.join(OUT, f)) for f in os.listdir(OUT)})
