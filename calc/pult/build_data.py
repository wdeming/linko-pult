# -*- coding: utf-8 -*-
"""Собирает данные пульта из pult/work/<срез>/: out/metrics.json, out/series.json, out/bundle_<срез>.json (Excel в base64)."""
import os, json, glob, base64, collections, datetime
P = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(P); OUT = os.path.join(P, 'out'); os.makedirs(OUT, exist_ok=True)
TAG = {8: 'aug31', 9: 'sep30', 10: 'oct31', 11: 'nov30', 12: 'dec31'}
RM = ['', 'янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек']
ml = lambda m: '%s %s' % (RM[int(m[5:7])], m[2:4])
cuts = sorted(os.path.basename(d) for d in glob.glob(os.path.join(P, 'work', '20*')) if os.path.exists(os.path.join(d, 'stamp.json')))
PP, F = 'POS', 'SFA'
# продукты по полю «Тип подписки» (с 26.11.2026): основные SFA / POS + допродукты SFA — AI ассистент и Perfect store (отдельные подписки)
ALLP = (F, PP, 'AI', 'PS'); ADD = ('AI', 'PS'); PN = {F: 'SFA', PP: 'POS', 'AI': 'AI ассистент', 'PS': 'Perfect store'}; FAMP = {F: F, 'AI': F, 'PS': F, PP: PP}
REVK = ('Подписки SFA', 'Подписки POS', 'Подписки AI ассистент', 'Подписки Perfect store', 'Разовые работы и прочее')
DOCS = {'mrr': 'MRR_и_когортное_удержание_2026_на_%s.xlsx', 'churn': 'Отток_подписок_причины_%s_2026_на_%s.xlsx', 'clients': 'Клиентская_база_продукты_сегменты_на_%s.xlsx',
        'multi': 'Клиенты_с_несколькими_подписками_на_%s.xlsx', 'top20': 'Концентрация_выручки_и_условия_топ20_на_%s.xlsx', 'cost': 'Себестоимость_поддержки_внедрения_инфраструктуры_на_%s.xlsx',
        'metrics': 'Ключевые_метрики_SaaS_на_%s_с_разовой.xlsx', 'calls': 'Обзвоны_клиентов_сырая_выгрузка_Odoo_на_%s.xlsx',
        'cash': 'Поступления_денег_%s_2026_на_%s.xlsx', 'exp': 'Расходы_кэшфло_и_за_период_%s_2026_на_%s.xlsx', 'fin': 'Финансы_для_банка_%s_2026_на_%s.xlsx'}
MANUAL = {'org': 'Оргструктура_Linktrade_на_%s.xlsx', 'revenue': 'Выручка_по_продуктам_и_типам_%s_2026_на_%s.xlsx'}
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
for cut in cuts:
    dn = os.path.join(P, 'work', cut); t = TAG[int(cut[5:7])]; m = cut[:7]
    L = lambda n: json.load(open(os.path.join(dn, 'x_%s_%s.json' % (t, n)), encoding='utf-8'))
    ddmm = '%s.%s' % (cut[8:10], cut[5:7]); per = 'Фев-' + RM[int(cut[5:7])].capitalize()
    R, K, M, C, base, const, calls, fot, surv = L('res_metrics'), L('cost'), L('res_mrr'), L('res_churn'), L('base'), L('const'), L('res_calls'), L('fot'), L('survival')
    stamp = json.load(open(os.path.join(dn, 'stamp.json')))
    def add(doc, key, label, value, kind='n', good='', prod=''):
        metrics.append({'id': '%s.%s' % (cut, key), 'cut': cut, 'doc': doc, 'key': key, 'label': label, 'value': None if value is None else round(float(value), 6), 'kind': kind, 'good': good, 'prod': prod})
    def ser(chart, x, s, v):
        series.append({'id': '%s.%s.%s.%s' % (cut, chart, x, s), 'cut': cut, 'chart': chart, 'x': x, 's': s, 'v': None if v is None else v if isinstance(v, str) else round(float(v), 6)})
    mov = lambda e, p: M['mov'].get('%s|%s|%s' % (p, m, e), 0)
    MR_ = {p: R['MRR'].get(p, 0.0) for p in ALLP}; tm = sum(MR_.values()); tmj = tm
    # --- сводка и MRR
    add('mrr', 'mrr_total', 'MRR всего', tm, 'm', 'up'); add('mrr', 'mrr_sfa', 'MRR SFA', R['MRR'][F], 'm', 'up', F); add('mrr', 'mrr_pos', 'MRR POS', R['MRR'][PP], 'm', 'up', PP)
    for p in ADD:
        if p in R['MRR']: add('mrr', 'mrr_' + p.lower(), 'MRR ' + PN[p], R['MRR'][p], 'm', 'up', p)
    _sp = const['SPLIT']
    add('mrr', 'mrr_paused', 'в т.ч. MRR подписок на паузе (входит в MRR всего)', sum(_sp[p]['4_paused'][1] for p in ALLP if p in _sp), 'm', '')
    add('mrr', 'subs_paused', 'в т.ч. подписок на паузе (входят в базу)', sum(_sp[p]['4_paused'][0] for p in ALLP if p in _sp), 'n', '')
    STRU = {'3_progress': 'в работе', '4_paused': 'на паузе', '6_churn': 'закрыта', '5_renewed': 'продлена', '1_draft': 'черновик', '2_renewal': 'продление', '7_upsell': 'допродажа'}
    for e in sorted(const.get('JERR') or [], key=lambda e: -abs(e['log'] - e['sub'])):
        lab = '%s · %s' % (e['name'], STRU.get(e['st'], e['st'] or ''))
        ser('mrr_jerr', lab, 'По подписке', e['sub']); ser('mrr_jerr', lab, 'В журнале Odoo', e['log']); ser('mrr_jerr', lab, 'Разница', e['log'] - e['sub'])
    add('mrr', 'subs', 'Подписок в базе', len(base), 'n', 'up'); add('mrr', 'subs_sfa', 'Подписок SFA', R['SRV'][F], 'n', 'up', F); add('mrr', 'subs_pos', 'Подписок POS', R['SRV'][PP], 'n', 'up', PP)
    for p in ADD:
        if p in R['SRV']: add('mrr', 'subs_' + p.lower(), 'Подписок ' + PN[p], R['SRV'][p], 'n', 'up', p)
    # потоки MRR: всего, SFA, POS
    for key, lab, e, sign, good in (('new', 'Новые продажи за месяц', '0_creation', 1, 'up'), ('exp', 'Расширения за месяц', '1_expansion', 1, 'up'),
                                    ('con', 'Сокращения за месяц', '15_contraction', -1, 'down'), ('chu', 'Отток за месяц', '2_churn', -1, 'down')):
        add('mrr', key + '_total', lab + ', всего', sign * sum(mov(e, p) for p in ALLP), 'm', good)
        for p in ALLP:
            if p in (F, PP) or (key == 'new' and p in R['MRR']): add('mrr', key + '_' + p, '%s, %s' % (lab, PN[p]), sign * mov(e, p), 'm', good, p)
    _gn = {p: M['GN'].get('%s|%s' % (p, m)) for p in ALLP}; S0 = sum(g['start'] for g in _gn.values() if g)
    lost = sum(mov('2_churn', p) + mov('15_contraction', p) for p in ALLP if _gn[p]); gain = sum(mov('1_expansion', p) for p in ALLP if _gn[p])
    add('mrr', 'grr_total', 'GRR месяца, всего', (S0 + lost) / S0, 'p', 'up')
    for p in (F, PP): add('mrr', 'grr_' + p, 'GRR месяца, %s' % p, R['GRR'][p], 'p', 'up', p)
    add('mrr', 'nrr_total', 'NRR месяца, всего', (S0 + lost + gain) / S0, 'p', 'up')
    for p in (F, PP): add('mrr', 'nrr_' + p, 'NRR месяца, %s' % p, R['NRR'][p], 'p', 'up', p)
    for p in ALLP:
        for x in R['series'].get(p, []):
            ser('mrr_end', ml(x['m']), PN[p], x['end'])
            if p in (F, PP):
                for k, nm in (('new', 'Новые'), ('exp', 'Расширения'), ('con', 'Сокращения'), ('chu', 'Отток')): ser('flow_' + p, ml(x['m']), nm, x[k])
    for p in (F, PP):
        for k, g in M['GN'].items():
            if k.startswith(p + '|'): ser('retention', ml(k.split('|')[1]), 'GRR ' + p, g['grr']); ser('retention', ml(k.split('|')[1]), 'NRR ' + p, g['nrr'])
        for a in surv[p]: ser('survival', '%d мес' % a[0], p, a[3])
    for mm in sorted({k.split('|')[1] for k in M['GN']}):
        g = [x for x in (M['GN'].get('%s|%s' % (p, mm)) for p in ALLP) if x]
        if M['GN'].get('%s|%s' % (F, mm)) and M['GN'].get('%s|%s' % (PP, mm)):
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
        c = cl[b['partner_id']]; c['mrr'] += b['mrr']; c['prods'].add(FAMP.get(b['prod'], b['prod'])); c['n'] += 1 if b['prod'] not in ADD else 0; c['subs'].append(b); c['name'] = b['pname']   # допродукты — в линейке SFA, в счёт подписок не идут
    seg = {s: [sum(1 for c in cl.values() if f(c)), sum(c['mrr'] for c in cl.values() if f(c))] for s, f in
           (('Только SFA', lambda c: c['prods'] == {F}), ('Только POS', lambda c: c['prods'] == {PP}), ('Оба продукта', lambda c: len(c['prods']) == 2))}
    add('clients', 'clients', 'Клиентов всего', len(cl), 'n', 'up'); add('clients', 'cl_sfa', 'Только SFA', seg['Только SFA'][0], 'n', 'up', F); add('clients', 'cl_pos', 'Только POS', seg['Только POS'][0], 'n', 'up', PP)
    add('clients', 'cl_both', 'Оба продукта', seg['Оба продукта'][0], 'n', 'up'); add('clients', 'arpc', 'Средний MRR на клиента', tm / len(cl), 'm', 'up')
    sp = const['SPLIT']; add('clients', 'paused_n', 'Подписок на паузе', sum(sp[p]['4_paused'][0] for p in ALLP if p in sp), 'n', 'down'); add('clients', 'paused_mrr', 'MRR на паузе', sum(sp[p]['4_paused'][1] for p in ALLP if p in sp), 'm', 'down')
    add('clients', 'lic_sfa', 'Лицензий SFA', R['LIC'][F], 'n', 'up', F); add('clients', 'lic_pos', 'Лицензий POS', R['LIC'][PP], 'n', 'up', PP)
    for p in ADD:
        if p in R['MRR']: add('clients', 'lic_' + p.lower(), 'Лицензий ' + PN[p], R['LIC'].get(p, 0), 'n', 'up', p)
    for s, v in seg.items(): ser('seg_clients', s, 'Клиентов', v[0]); ser('seg_mrr', s, 'MRR', v[1])
    jout = sum((const.get('MRR_JOUT') or {}).values())
    if abs(jout) > 1: ser('seg_mrr', 'Расхождения журнала Odoo', 'MRR', jout)
    add('clients', 'cl_mrr', 'MRR клиентов всего (по журналу Odoo)', tm, 'm', 'up')
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
    CP = (('SUP', 'Техподдержка'), ('ONP', 'Onboarding POS'), ('ONS', 'Onboarding SFA'), ('KA', 'Ключевые клиенты'), ('HOST_AVG', 'Хостинг'))
    for k, nm in CP:
        add('cost', 'c_' + k, nm + ' за месяц', K[k], 'm', 'down'); ser('cost_parts', nm, 'За месяц', K[k])
    add('cost', 'host_m', 'Оплаты хостинга за месяц (по дате оплаты)', (K.get('HOST_PAY_M') or K['HOST_M']).get(m), 'm', 'down')
    for mmm, v in K['HOST_M'].items():
        ser('hosting', ml(mmm), 'За период (в себестоимости)', v)
        if K.get('HOST_PAY_M'): ser('hosting', ml(mmm), 'По дате оплаты', K['HOST_PAY_M'].get(mmm, 0))
    if K.get('FOT_M'):
        for mmm in K['HOST_M']:
            for k, nm in CP[:4]: ser('cost_m', ml(mmm), nm, K['FOT_M'][k].get(mmm, 0))
            ser('cost_m', ml(mmm), 'Хостинг', K['HOST_M'][mmm])
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
    if 'base_clients' in calls:
        add('calls', 'calls_cov', 'Клиентов базы, с кем поговорили за 90 дней', calls['reached90'] / calls['base_clients'] if calls['base_clients'] else None, 'p', 'up')
        add('calls', 'calls_cov_n', 'Поговорили за 90 дней, клиентов', calls['reached90'], 'n', 'up')
        add('calls', 'calls_open', 'Проблем из обзвонов не решено', calls['open_problems'], 'n', 'down')
        add('calls', 'calls_scored', 'Звонков с оценкой клиента (для NPS) за месяц', calls['scored_month'], 'n', 'up')
        for k2, v in sorted(calls['ops'].items(), key=lambda z: -z[1]): ser('calls_ops', k2[:30], 'Звонков за месяц', v)
        for k2, v in sorted(calls['open_by_route'].items(), key=lambda z: -z[1]): ser('calls_open', k2[:40], 'Не решено', v)
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
    # --- финансы для банка: P&L, деньги, выручка, сегменты
    fp = os.path.join(dn, 'x_%s_res_fin.json' % t)
    if os.path.exists(fp):
        FN = json.load(open(fp, encoding='utf-8')); MSx = FN['MS']; PLm = FN['PL'].get(m); RECx = FN['REC']
        rkx = lambda k, mm: RECx.get('%s|%s' % (k, mm), 0)
        def pl_row(lab, f, pct=False):
            tot_n = tot_d = 0.0
            for mm in MSx:
                v = f(mm)
                if pct: n_, d_ = v; ser('pl_tab', lab, ml(mm), n_ / d_ if d_ else None); tot_n += n_; tot_d += d_
                else: ser('pl_tab', lab, ml(mm), v); tot_n += v
            ser('pl_tab', lab, 'Итого', (tot_n / tot_d if tot_d else None) if pct else tot_n)
        PLd = FN['PL']
        for k_ in REVK:
            if k_ in ('Подписки SFA', 'Подписки POS', 'Разовые работы и прочее') or any(kk.startswith(k_ + '|') for kk in RECx): pl_row(k_, lambda mm, k_=k_: rkx(k_, mm))
        pl_row('ВЫРУЧКА', lambda mm: PLd[mm]['rev'])
        pl_row('ФОТ поддержки, внедрения, ключевых клиентов', lambda mm: FN['COGS_FOT'][mm]); pl_row('Хостинг и инфраструктура', lambda mm: FN['HOST'][mm])
        pl_row('СЕБЕСТОИМОСТЬ', lambda mm: PLd[mm]['cogs']); pl_row('ВАЛОВАЯ ПРИБЫЛЬ', lambda mm: PLd[mm]['gp'])
        pl_row('Валовая маржа, %', lambda mm: (PLd[mm]['gp'], PLd[mm]['rev']), True)
        for lab, v in FN['OPX'].items(): pl_row(lab.split(' (')[0] if lab.startswith('Персонал') or lab.startswith('Комиссии') else lab, lambda mm, v=v: v[mm])
        pl_row('ОПЕРАЦИОННЫЕ РАСХОДЫ', lambda mm: PLd[mm]['opex']); pl_row('EBITDA', lambda mm: PLd[mm]['ebitda'])
        pl_row('EBITDA, % выручки', lambda mm: (PLd[mm]['ebitda'], PLd[mm]['rev']), True)
        pl_row('Капитальные расходы (не в EBITDA)', lambda mm: FN.get('CAPEX', {}).get(mm, 0))
        if FN.get('BR'):     # мостик: выручка по периодам услуги → как в Odoo (по дате счёта, вся сумма сразу)
            BRx = FN['BR']
            pl_row('+ рост оплаченного вперёд за будущие месяцы', lambda mm: BRx[mm]['ddef']); pl_row('+ счета месяца за услуги прошлых месяцев', lambda mm: BRx[mm]['late_in'])
            pl_row('− услуги месяца, счёт выставлен позже', lambda mm: -BRx[mm]['late_out']); pl_row('ВЫРУЧКА КАК В ODOO', lambda mm: PLd[mm]['rev_odoo'])
            pl_row('EBITDA при выручке как в Odoo', lambda mm: PLd[mm]['ebitda_odoo'])
        for mm in MSx:
            for k2, nm in (('rev', 'Выручка'), ('cogs', 'Себестоимость'), ('opex', 'Операционные расходы'), ('ebitda', 'EBITDA')): ser('pl_chart', ml(mm), nm, PLd[mm][k2])
        if PLm:
            add('fin_pl', 'f_rev', 'Выручка за месяц (по дате счёта)', PLm['rev'], 'm', 'up'); add('fin_pl', 'f_gm', 'Валовая маржа', PLm['gp'] / PLm['rev'] if PLm['rev'] else None, 'p', 'up')
            add('fin_pl', 'f_opex', 'Операционные расходы за месяц', PLm['opex'], 'm', 'down'); add('fin_pl', 'f_ebitda', 'EBITDA за месяц', PLm['ebitda'], 'm', 'up')
            add('fin_pl', 'f_ebitda_pct', 'EBITDA, % выручки', PLm['ebitda'] / PLm['rev'] if PLm['rev'] else None, 'p', 'up')
            add('fin_pl', 'f_ebitda_cum', 'EBITDA с апреля, накопленно', sum(PLd[mm]['ebitda'] for mm in MSx), 'm', 'up')
            if 'rev_odoo' in PLm: add('fin_pl', 'f_rev_odoo', 'Справочно: выручка как в Odoo за месяц (по дате счёта)', PLm['rev_odoo'], 'm', 'up')
        CSx = FN['CS']; cm = CSx.get(m); MCx = FN.get('MC') or MSx
        if FN.get('MC'):
            for mm in MCx: ser('cash_tab', 'Источник', ml(mm), CSx[mm].get('src'))
        for lab, k2 in (('Остаток денег на начало месяца', 'open'), ('Поступления от клиентов', 'in'), ('Комиссии платёжных систем', 'com'), ('Расходы по дате оплаты', 'exp'),
                        ('ОПЕРАЦИОННЫЙ ДЕНЕЖНЫЙ ПОТОК', 'op'), ('Займы получены', 'lt'), ('Займы возвращены', 'lp'), ('Прочие движения и расхождения учёта', 'oth'),
                        ('Остаток денег на конец месяца', 'end'), ('Долг по займам на конец месяца', 'debt')):
            sg = -1 if k2 in ('com', 'exp', 'lp') else 1
            for mm in MCx: ser('cash_tab', lab, ml(mm), (CSx[mm][k2] or 0) * sg if CSx[mm][k2] is not None else None)
        for mm in MCx:
            ser('cash_flow', ml(mm), 'Операционный поток', CSx[mm]['op']); ser('cash_flow', ml(mm), 'Займы, нетто', CSx[mm]['lt'] - CSx[mm]['lp'])
            ser('cash_bal', ml(mm), 'Деньги на конец месяца', CSx[mm]['end'])
        for a, v in sorted(FN['ACC'].items(), key=lambda z: -z[1]): ser('cash_acc', a[:40], 'Остаток', v)
        if cm:
            add('fin_cash', 'f_cash', 'Деньги на конец месяца', cm['end'], 'm', 'up'); add('fin_cash', 'f_op', 'Операционный денежный поток за месяц', cm['op'], 'm', 'up')
            add('fin_cash', 'f_avg3', 'Средний операционный поток за 3 месяца', FN['AVG3'], 'm', 'up')
            add('fin_cash', 'f_runway', 'Запас денег, месяцев', cm['end'] / -FN['AVG3'] if FN['AVG3'] < 0 else None, 'd1', 'up')
            if FN.get('LOANS_OK'):
                add('fin_cash', 'f_loans_m', 'Займы получены за месяц', cm['lt'], 'm', ''); add('fin_cash', 'f_debt', 'Долг по займам на конец месяца', cm['debt'], 'm', 'down')
                add('fin_cash', 'f_loans_cum', 'Займы с апреля, нетто', sum(CSx[mm]['lt'] - CSx[mm]['lp'] for mm in MSx), 'm', '')
        for mm in FN['MR']:
            ser('rev_lines', ml(mm), 'MRR', FN['MRR'][mm]); ser('rev_lines', ml(mm), 'Выручка', sum(RECx.get('%s|%s' % (k, mm), 0) for k in REVK))
            ser('rev_lines', ml(mm), 'Получено', FN['IN'][mm])
            ser('rev_def', ml(mm), 'Оплачено вперёд', FN['DEF'][mm])
            ser('rev_sfa', ml(mm), 'Выручка SFA', rkx('Подписки SFA', mm))
        for b, v in FN['AR'].items(): ser('ar_age', b, 'Сумма', v)
        HI = FN.get('HIST')      # история с 2025: файл финансиста, пакет, Odoo
        if HI and HI['M']:
            for lab, kind, vals, tot in HI['ROWS']:
                for mm in HI['M']: ser('hist_tab', lab, ml(mm), vals.get(mm))
                for yk in HI['Y']: ser('hist_tab', lab, yk, tot.get(yk))
            HR = {r_[0]: r_ for r_ in HI['ROWS']}
            for mm in HI['M']:
                ser('hist_chart', ml(mm), 'Выручка', HR['ВЫРУЧКА'][2][mm]); ser('hist_chart', ml(mm), 'Расходы', HR['РАСХОДЫ ВСЕГО'][2][mm]); ser('hist_chart', ml(mm), 'EBITDA', HR['EBITDA'][2][mm])
            for yk in HI['Y']:
                y4 = yk.split()[1].rstrip(',')
                add('fin_hist', 'h_rev_' + y4, 'Выручка, ' + yk.replace('Итого ', ''), HR['ВЫРУЧКА'][3].get(yk), 'm', 'up')
                add('fin_hist', 'h_ebitda_' + y4, 'EBITDA, ' + yk.replace('Итого ', ''), HR['EBITDA'][3].get(yk), 'm', 'up')
            add('fin_hist', 'h_capex_tap', 'Техника и Tap 2025 — вынесены под EBITDA', (HR['Капрасходы: техника и мебель (не в EBITDA)'][3].get(HI['Y'][0]) or 0) + (HR['Финансирование Tap (не в EBITDA)'][3].get(HI['Y'][0]) or 0), 'm', '')
        if FN.get('BAL'):
            BV = {}
            for lab, v, src in FN['BAL'] + [(None, None, None)] + FN['BREF']:
                if lab is None: continue
                ser('bal_tab', lab, 'Сумма', v); ser('bal_tab', lab, 'Источник и как посчитано', src); BV[lab] = v
            add('fin_bal', 'b_cash', 'Деньги', BV['Денежные средства'], 'm', 'up'); add('fin_bal', 'b_liab', 'Обязательства всего', BV['ИТОГО ОБЯЗАТЕЛЬСТВА'], 'm', 'down')
            add('fin_bal', 'b_loans', 'из них займы', BV['Займы'], 'm', 'down'); add('fin_bal', 'b_def', 'Справочно: оплачено вперёд за будущие месяцы', BV.get('Справочно: оплачено клиентами вперёд за будущие месяцы', BV.get('Предоплаты клиентов за будущие месяцы')), 'm', '')
            add('fin_bal', 'b_eq', 'Капитал и накопленный результат', BV['Капитал и накопленный результат'], 'm', 'up')
            for row in FN.get('SRCS', []):
                ser('src_tab', row[0], 'Источник', row[1]); ser('src_tab', row[0], 'Учёт выручки', row[2]); ser('src_tab', row[0], 'Как посчитано', row[3])
            for x in FN.get('CHK', []):
                for col, k2 in (('Источник А', 'a_lab'), ('Сумма А', 'a'), ('Источник Б', 'b_lab'), ('Сумма Б', 'b'), ('Разница', 'diff'), ('Что взято и почему', 'note')): ser('src_check', x['what'], col, x[k2])
        add('fin_rev', 'f_mrr', 'MRR на конец месяца', FN['MRR'][m], 'm', 'up'); add('fin_rev', 'f_rec', 'Выручка за месяц (по дате счёта)', sum(rkx(k, m) for k in REVK), 'm', 'up')
        add('fin_rev', 'f_in', 'Получено за месяц', FN['IN'][m], 'm', 'up')
        if m >= '2026-10': add('fin_rev', 'f_mrr0', 'MRR на 01.01.2026 (данные менеджмента)', 440e6, 'm', ''); add('fin_rev', 'f_mrr_ytd', 'Рост MRR с начала года', FN['MRR'][m] / 440e6 - 1, 'p', 'up')
        add('fin_rev', 'f_def', 'Оплачено клиентами вперёд за будущие месяцы', FN['DEF'][m], 'm', '')
        add('fin_rev', 'f_ar', 'Неоплаченные счета (дебиторка)', sum(FN['AR'].values()), 'm', 'down'); add('fin_rev', 'f_ar90', 'из них просрочены больше 90 дней', FN['AR']['91–180 дней'] + FN['AR']['больше 180 дней'], 'm', 'down')
        add('fin_rev', 'f_unm', 'Поступления, не привязанные к счетам', FN['UNM']['s'], 'm', 'down')
        SGx = FN['SEG']
        for sname in FN['SEGS'] + ['Всего']:
            g = SGx[sname]
            for col, v in (('Клиентов', g['n']), ('MRR', g['mrr']), ('Доля MRR', g['mrr'] / SGx['Всего']['mrr'] if SGx['Всего']['mrr'] else None), ('NRR с 31.01', g['nrr']), ('GRR с 31.01', g['grr']),
                           ('NRR за год (пересчёт)', g['nrr_y']), ('Ушли клиентов с 31.01', g['lost']), ('Потери MRR за 3 мес.', g['loss3'])):
                ser('seg_tab', sname, col, v)
            if sname != 'Всего': ser('seg_nrr', sname, 'NRR с 31.01', g['nrr']); ser('seg_nrr', sname, 'GRR с 31.01', g['grr'])
        for cm2, cd2 in FN['COH'].items():
            lab = ml(cm2); ser('coh_money', lab, 'Подписок', cd2['n']); ser('coh_money', lab, 'MRR старта', cd2['start'])
            for i, v in enumerate(cd2['k']): ser('coh_money', lab, 'М%d' % i, v)
        tg = SGx['Всего']
        add('fin_seg', 'f_nrr', 'NRR с 31.01, всего', tg['nrr'], 'p', 'up'); add('fin_seg', 'f_grr', 'GRR с 31.01, всего', tg['grr'], 'p', 'up')
        add('fin_seg', 'f_nrr_y', 'NRR за год (пересчёт с %d мес.)' % FN['NMON'], tg['nrr_y'], 'p', 'up'); add('fin_seg', 'f_lost', 'Ушли клиентов с 31.01', tg['lost'], 'p', 'down')
        big = [s2 for s2 in FN['SEGS'] if '51+' in s2 or '21–50' in s2]
        if big: add('fin_seg', 'f_big', 'Доля MRR клиентов от 21 агента', sum(SGx[s2]['mrr'] for s2 in big) / tg['mrr'] if tg['mrr'] else None, 'p', '')
        if FN.get('RET'):
            RT = FN['RET']; nm_ = FN['NMON']
            for lab in ('SFA', 'POS', 'Всего'):
                x = RT[lab]
                for col, v in (('Клиентов на 31.01', x['n0']), ('MRR на 31.01', x['mrr0']), ('MRR этих клиентов на срез', x['mrr_now']), ('NRR за %d мес.' % nm_, x['nrr']), ('GRR за %d мес.' % nm_, x['grr']),
                               ('NRR за год (пересчёт)', x['nrr_y']), ('GRR за год (пересчёт)', x['grr_y']), ('Ушли клиентов', x['lost'])):
                    ser('ret_prod', lab, col, v)
            for i, c3 in enumerate(RT['Всего']['curve']):
                lab = 'янв 26' if c3['m'] == '2026-01' else ml(c3['m'])
                ser('ret_curve', lab, 'NRR всего', c3['nrr']); ser('ret_curve', lab, 'GRR всего', c3['grr'])
                ser('ret_curve', lab, 'NRR SFA', RT['SFA']['curve'][i]['nrr']); ser('ret_curve', lab, 'NRR POS', RT['POS']['curve'][i]['nrr'])
            for mm, x in FN['MONTHLY'].items():
                ser('ret_con', ml(mm), 'Колебания (вернули позже)', x['con_back']); ser('ret_con', ml(mm), 'Настоящие сокращения', x['con_real']); ser('ret_con', ml(mm), 'Отток', x['churn'])
                ser('ret_month', ml(mm), 'GRR месяца', x['grr']); ser('ret_month', ml(mm), 'GRR без колебаний', x['grr_adj']); ser('ret_month', ml(mm), 'NRR месяца', x['nrr'])
            tv = RT['Всего']
            add('fin_ret', 'r_nrr', 'NRR с 31.01, %d мес., всего' % nm_, tv['nrr'], 'p', 'up'); add('fin_ret', 'r_grr', 'GRR с 31.01, %d мес., всего' % nm_, tv['grr'], 'p', 'up')
            add('fin_ret', 'r_nrr_y', 'NRR за год (пересчёт), всего', tv['nrr_y'], 'p', 'up'); add('fin_ret', 'r_grr_y', 'GRR за год (пересчёт), всего', tv['grr_y'], 'p', 'up')
            add('fin_ret', 'r_nrr_sfa', 'NRR за год (пересчёт), SFA', RT['SFA']['nrr_y'], 'p', 'up', F); add('fin_ret', 'r_nrr_pos', 'NRR за год (пересчёт), POS', RT['POS']['nrr_y'], 'p', 'up', PP)
            add('fin_ret', 'r_conback', 'Сокращения, которые потом вернулись', FN['CON_B'] / FN['CON_T'] if FN['CON_T'] else None, 'p', 'up')
    # --- оргструктура
    add('org', 'staff', 'Штат, человек', sum(x['n'] for x in fot), 'n', ''); add('org', 'payroll', 'Оклады на руки в месяц', sum(x['sal'] for x in fot), 'm', 'down')
    add('org', 'payroll_share', 'Оклады, % от MRR', sum(x['sal'] for x in fot) / tmj, 'p', 'down')
    gone = sum(x.get('gone', 0) for x in fot)
    add('org', 'staff_gone', 'Из них уволены после среза (на сегодня)', gone, 'n', ''); add('org', 'staff_now', 'Из этого штата работают сегодня', sum(x['n'] for x in fot) - gone, 'n', '')
    for x in fot:
        ser('staff_dep', (x['dep'] or '—')[:30], 'На срез', x['n']); ser('payroll_dep', (x['dep'] or '—')[:30], 'Оклады', x['sal'])
        if 'gone' in x: ser('staff_dep', (x['dep'] or '—')[:30], 'Уволены после среза', x['gone']); ser('staff_dep', (x['dep'] or '—')[:30], 'Работают сегодня', x['n'] - x['gone'])
    # --- выручка и долги (из файлов комплекта, обновляются вручную)
    rv = revenue(cut, ddmm, per)
    if rv:
        tot = sum(v[m] for v in rv.values()); add('revenue', 'rev_total', 'Выручка за месяц', tot, 'm', 'up')
        for k, v in rv.items():
            add('revenue', 'rev_' + str(list(rv).index(k)), k, v[m], 'm', 'up')
            for mmm, val in v.items(): ser('revenue_m', ml(mmm), k, val)
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
