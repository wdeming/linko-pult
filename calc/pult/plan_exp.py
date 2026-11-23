# -*- coding: utf-8 -*-
"""Плановые расходы на 6 месяцев (вызывается из plan.py). Только чтение Odoo.
1) Плановые платежи — реестр «Плановые расходы» (x_exp_plan, активные): «Дата след. оплаты», сумма, валюта, повторяющийся (= ежемесячно) или разовый.
   Сверка с фактом — реестр «Расходы» (x_custom_expenses): если по статье уже есть оплата за этот период (период расхода покрывает дату
   платежа или оплата была в 25 дней до неё и позже) — платёж считается оплаченным и сдвигается на следующий месяц (оплата вперёд тоже так
   учитывается). Платёж с датой в прошлом без оплаты — просрочен и ставится в текущий месяц.
2) ФОТ — ведомость текущего месяца (x_wages): оклады + KPI; неоплаченный остаток текущего месяца — в следующем месяце; дальше оклады
   платятся по факту прошлых месяцев: доля в своём месяце (аванс) и остаток в следующем. Долги по ЗП прошлых периодов (остаток «Баланс»
   в ведомости) — отдельной колонкой. Квартальный KPI — по последнему факту, в месяце после конца квартала.
3) Расходы, которых нет в плановом реестре (сервера Ритм, Proxima, обеды, маркетинг…) — среднее факта по дате оплаты за 3 полных месяца."""
import re, datetime, calendar, collections, os, sqlite3
K = 12000; RUB = 150.0
CUR_RATE = {'USD': K, 'RUB': RUB, 'UZS': 1}
FOT_CATS = {'Оклад', 'KPI', 'Материальная помощь'}
CAPEX = {'Техника и мебель'}
# сопоставление плановых платежей с фактом: (статьи, регулярное выражение по описанию, сумма ±%)
RULES = [
    (r'аренда офиса', {'Аренда'}, None), (r'корпоративные номера', None, r'uztelecom|корпоратив|номер'), (r'admixer', None, r'admixer'),
    (r'playmobile', None, r'play ?mobile|плеймоб'), (r'servorcore|servercore', None, r'servercore|servorcore|серверкор'),
    (r'бухгалтерское обслуживание', {'Аутсорс'}, r'бухгалтер(?!.*ritm)'), (r'бухгалтерия аутсорс ritm', {'Аутсорс'}, r'бухгалтер'),
    (r'amo crm', None, r'\bamo'), (r'сервер\w* rit|ritm|ритм', {'Хостинг и инфраструктура'}, r'ритм|ritm'), (r'sarkor', None, r'sarkor|саркор'), (r'атто', {'Дорожные расходы (ATTO карты)', 'Дорожные расходы (АТТО карты)'}, None),
    (r'интеренет|интернет', {'Интернет'}, None), (r'online pbx', None, r'pbx'), (r'appstore', None, r'app ?store|apple'), (r'selectel', None, r'selectel'),
    (r'webstorm', None, r'webstorm'), (r'pycharm', None, r'pycharm'), (r'^ssl', None, r'\bssl'), (r'cloud?a?flare|clouadflare', None, r'cloudflare|clouadflare'),
    (r'sentry', None, r'sentry'), (r'monthly visa', None, r'visa|виза'), (r'hetzner', None, r'hetzner'), (r'canva', None, r'canva'),
    (r'chat ?gpt', None, r'gpt'), (r'салфет', {'Хозтовары'}, r'салфет|бумаг'), (r'налоги зп', {'Налоги ЗП'}, None), (r'^it park', {'IT Park'}, None),
    (r'налог на аренду', None, r'налог\w* (на )?аренд'), (r'коммунальн', {'Обслуживание офиса'}, r'свет|газ|вод|коммунал|электр'),
    (r'cloud google', {'Google CLOUD'}, None), (r'google workspace', None, r'workspace'), (r'cloude для всех|anthropic claude', {'CLAUDE'}, None),
    (r'расходы офиса', {'Бакалея', 'Канцтовары'}, None)]
def addm(d, n):
    y, m = divmod(d.month - 1 + n, 12); y += d.year; m += 1
    return datetime.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))
def build(q, TODAY, CUR_END, HM, ml, C_CUR, FC):
    H_END = addm(datetime.date.fromisoformat(HM[-1] + '-01'), 1) - datetime.timedelta(days=1)
    col = lambda d: C_CUR if d <= CUR_END else ml(d.strftime('%Y-%m'))
    clean = lambda s: re.sub(r'\s+', ' ', re.sub(r'<[^>]+>|&nbsp;', ' ', s or '')).strip()
    # ---- факт: реестр «Расходы»
    ACT = []
    for r in q("""select h.id, h.x_studio_date pay, h.x_studio_start_date d1, h.x_studio_end_date d2, coalesce(h.x_studio_rate, 0) rate, h.x_studio_notes nt,
          coalesce(aa.name->>'ru_RU', aa.name->>'en_US', '') cat, (select coalesce(sum(l.x_studio_sum), 0) from x_custom_expenses_line_dc8d4 l where l.x_custom_expenses_id = h.id) s
        from x_custom_expenses h left join account_analytic_account aa on aa.id = h.x_studio_accountanalyticaccount
        where h.x_active is not false and h.x_studio_date is not null and h.x_studio_date >= date '2026-04-01'"""):
        d1 = r['d1'] or r['pay']; d2 = r['d2'] or d1
        ACT.append({'pay': r['pay'], 'd1': min(d1, d2), 'd2': max(d1, d2), 'cat': r['cat'], 'note': clean(r['nt']).lower(), 'amt': float(r['s']) * (K if r['rate'] > 1 else 1), 'used': False})
    # ---- плановые платежи
    ITEMS = q("""select p.id, p.x_name nm, p.x_studio_date d, p.x_studio_payment_sum s, p.x_studio_repeat_payment rep, c.name cur, coalesce(st.x_name->>'ru_RU', st.x_name->>'en_US', '') stage
        from x_exp_plan p left join res_currency c on c.id = p.x_studio_currency_id left join x_exp_plan_stage st on st.id = p.x_studio_stage_id
        where p.x_active is not false and coalesce(p.x_studio_payment_sum, 0) > 0""")
    SCH = collections.defaultdict(float); OVR = collections.defaultdict(float); DET = []
    for it in ITEMS:
        nm = (it['nm'] if isinstance(it['nm'], str) else (it['nm'] or {}).get('ru_RU') or (it['nm'] or {}).get('en_US') or '').strip()
        amt = float(it['s']) * CUR_RATE.get(it['cur'] or 'UZS', 1)
        rule = next(((cats, rx) for pat, cats, rx in RULES if re.search(pat, nm.lower())), (None, None))
        def matches(a):
            cats, rx = rule
            if cats is None and rx is None: return False
            if cats and a['cat'] not in cats: return False
            if rx and not re.search(rx, a['note']): return False
            return True
        mine = [a for a in ACT if matches(a)]
        for a in mine: a['used'] = True
        due = it['d'] or TODAY.replace(day=1)
        paid_to = max((a['d2'] for a in mine if a['d2'] > a['d1']), default=None)        # оплачено вперёд по периоду
        free = sorted(mine, key=lambda a: a['pay'])
        def covered(dd):         # период оплаты покрывает дату или оплата была в пределах 25 дней от неё (каждая оплата закрывает один платёж)
            if any(a['d1'] <= dd <= a['d2'] and a['d2'] > a['d1'] for a in mine): return True
            for a in free:
                if abs((a['pay'] - dd).days) <= 25: free.remove(a); return True
            return False
        if it['rep']:
            for _ in range(24):
                if due <= H_END and covered(due): due = addm(due, 1)
                else: break
        elif covered(due): due = None
        row = {'name': nm, 'amt': amt, 'cur': it['cur'] or 'UZS', 'src_amt': float(it['s']), 'rep': bool(it['rep']), 'plan_date': str(it['d'] or ''), 'next': None, 'over': 0.0, 'n_fact': len(mine),
               'last_pay': str(max((a['pay'] for a in mine), default='') or ''), 'paid_to': str(paid_to or ''), 'm': {}}
        d = due
        while d and d <= H_END:
            if d < TODAY:
                if d >= addm(TODAY, -3): OVR[nm] += amt; row['over'] += amt; SCH[('Просроченные плановые платежи — оплатить сейчас', C_CUR)] += amt; row['m'][C_CUR] = row['m'].get(C_CUR, 0) + amt
            else:
                c = col(d); SCH[('Плановые платежи (реестр «Плановые расходы»)', c)] += amt; row['m'][c] = row['m'].get(c, 0) + amt; row['next'] = row['next'] or str(d)
            if not it['rep']: break
            d = addm(d, 1)
        DET.append(row)
    # ---- ФОТ по ведомостям
    W = q("""select w.id, w.x_name nm, w.x_studio_date d from x_wages w where w.x_studio_date is not null and w.x_studio_date <= date %s order by w.x_studio_date desc limit 4""", (TODAY,))
    WL = {w['id']: q("""select l.x_studio_salary s, l.x_studio_kpi k, l.x_studio_balance b from x_wages_line_9898e l where l.x_wages_id = %s""", (w['id'],)) for w in W}
    cur_w = W[0]; prev_w = W[1:]
    S = sum(float(l['s'] or 0) for l in WL[cur_w['id']])
    KPI_M = sum(sum(float(l['k'] or 0) for l in WL[w['id']]) for w in prev_w) / max(1, len(prev_w))     # ежемесячный KPI в ведомостях
    OLD = sum(max(float(l['b'] or 0), 0) for l in WL[cur_w['id']])
    # оплачено по ведомости текущего месяца — из реестра по статье «Оклад» за период текущего месяца
    m0 = TODAY.replace(day=1)
    paid_cur = sum(a['amt'] for a in ACT if a['cat'] == 'Оклад' and a['d1'] >= m0 and a['d1'] <= CUR_END)
    REM = max(S - paid_cur, 0.0)
    # доля оклада, выплаченная в своём месяце (аванс), — по трём прошлым полным месяцам
    sh = []
    for i in (1, 2, 3):
        pm = addm(m0, -i); pe = addm(pm, 1) - datetime.timedelta(days=1)
        tot = sum(a['amt'] for a in ACT if a['cat'] == 'Оклад' and pm <= a['d1'] <= pe)
        same = sum(a['amt'] for a in ACT if a['cat'] == 'Оклад' and pm <= a['d1'] <= pe and pm <= a['pay'] <= pe)
        if tot: sh.append(same / tot)
    A0 = sum(sh) / len(sh) if sh else 0.35
    QK = [a for a in ACT if a['cat'] == 'KPI']
    QKPI = max((sum(a['amt'] for a in QK if a['pay'].strftime('%Y-%m') == m) for m in {a['pay'].strftime('%Y-%m') for a in QK}), default=0.0)   # последний квартальный KPI
    FOT = collections.defaultdict(float)
    cols = [C_CUR] + [ml(m) for m in HM]
    for i, c in enumerate(cols):
        if i == 0: continue                                    # остаток текущего месяца: аванс уже выплачен
        FOT[('Оклады', c)] += (REM if i == 1 else (1 - A0) * S) + A0 * S
        FOT[('KPI ежемесячный (в ведомости)', c)] += KPI_M
        mm = int(HM[i - 1][5:7])
        if mm in (1, 4, 7, 10): FOT[('KPI квартальный (по последнему факту)', c)] += QKPI
    # ---- вне плана: факт по дате оплаты за 3 полных месяца, чего нет в плановом реестре и не ФОТ
    P3 = [addm(m0, -i).strftime('%Y-%m') for i in (3, 2, 1)]
    import fin_groups as FG
    UNP = collections.defaultdict(float); UNN = collections.defaultdict(float)
    for a in ACT:
        if a['used'] or a['cat'] in FOT_CATS or a['cat'] == 'Налоги ЗП' or a['pay'].strftime('%Y-%m') not in P3: continue
        g = 'Техника и мебель' if a['cat'] in CAPEX else FG.group(a['cat'])
        UNP[g] += a['amt'] / 3
        key = re.sub(r'\d[\d .,]*', '', a['note'])[:40].strip() or a['cat'].lower()
        UNN[(a['cat'], key)] += a['amt'] / 3
    share = lambda c: ((CUR_END - TODAY).days + 1) / calendar.monthrange(TODAY.year, TODAY.month)[1] if c == C_CUR else 1.0
    TAB = collections.OrderedDict()
    for k in ('Плановые платежи (реестр «Плановые расходы»)', 'Просроченные плановые платежи — оплатить сейчас'):
        TAB[k] = {c: SCH[(k, c)] for c in cols}
    for k in ('Оклады', 'KPI ежемесячный (в ведомости)', 'KPI квартальный (по последнему факту)'):
        TAB[k] = {c: FOT[(k, c)] for c in cols}
    for g, v in sorted(UNP.items(), key=lambda z: -z[1]):
        TAB['Вне плана: %s' % g] = {c: v * share(c) for c in cols}
    TAB['ВСЕГО ПЛАН РАСХОДОВ'] = {c: sum(TAB[k][c] for k in TAB) for c in cols}
    for v in TAB.values(): v['Итого за горизонт'] = sum(v[c] for c in cols)
    TAB['Долги по ЗП прошлых периодов (в ведомости)'] = {'Долги прошлых периодов': OLD}
    TAB['Просроченные плановые платежи — оплатить сейчас']['Долги прошлых периодов'] = sum(OVR.values())
    top = sorted(([c, n, v] for (c, n), v in UNN.items() if v >= 300e3), key=lambda z: -z[2])[:25]
    info = {'S': S, 'paid_cur': paid_cur, 'REM': REM, 'A0': A0, 'KPI_M': KPI_M, 'QKPI': QKPI, 'OLD': OLD, 'wages': cur_w['nm'] if isinstance(cur_w['nm'], str) else (cur_w['nm'] or {}).get('ru_RU') or (cur_w['nm'] or {}).get('en_US') or '', 'P3': P3, 'RUB': RUB}
    return TAB, DET, top, info
