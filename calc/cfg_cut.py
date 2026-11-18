# -*- coding: utf-8 -*-
"""Единые параметры среза комплекта data room.
Срез по умолчанию — 31.10.2026. Для пересчёта на другую дату: переменная окружения LINKO_CUT=ГГГГ-ММ-ДД
(и LINKO_OUT — папка для результатов, чтобы не затереть выпущенные файлы)."""
import datetime, os, calendar
CUT      = os.environ.get('LINKO_CUT', '2026-10-31')
CUT_D    = datetime.date.fromisoformat(CUT)
assert CUT_D.day == calendar.monthrange(CUT_D.year, CUT_D.month)[1], 'срез должен быть концом месяца'
CUT_RU   = CUT_D.strftime('%d.%m.%Y')
CUT_UTC  = CUT + ' 19:00:00'       # конец дня по Ташкенту (UTC+5) в UTC — для истории статусов
P1       = '2026-02-01'
_n       = (CUT_D.replace(day=1) + datetime.timedelta(days=32)).replace(day=1)
P2       = _n.isoformat()
EXPORT_RU = '18.11.2026'           # дата выгрузки
_T = {8: 'aug31', 9: 'sep30', 10: 'oct31', 11: 'nov30', 12: 'dec31'}
TAG      = _T[CUT_D.month]
SUF      = 'на_' + CUT_D.strftime('%d.%m')
OUTDIR   = os.environ.get('LINKO_OUT', '.')
_R = ['', 'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
_RG = ['', 'января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']
_RS = ['', 'Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек']
MON = ['2026-%02d' % m for m in range(2, CUT_D.month + 1)]          # с февраля по месяц среза
RUS = {m: _R[int(m[5:])] for m in MON}
RUS_GEN = {m: _RG[int(m[5:])] for m in MON}
RUS_SH = {m: _RS[int(m[5:])] for m in MON}
NM = len(MON)
LAST_M, PREV_M = MON[-1], MON[-2]
LAST_RU, LAST_RU_L, LAST_SH = RUS[LAST_M], RUS[LAST_M].lower(), RUS_SH[LAST_M]
PERIOD = 'Фев–' + LAST_SH                       # «Фев–Окт»
PERIOD_LONG = 'Февраль–%s 2026' % LAST_RU       # «Февраль–Октябрь 2026»
PREV_END = CUT_D.replace(day=1) - datetime.timedelta(days=1)
PREV_END_RU = PREV_END.strftime('%d.%m.%Y')     # конец предыдущего месяца
J    = lambda name: os.path.join(OUTDIR, 'x_%s_%s.json' % (TAG, name))
OUTP = lambda fn: os.path.join(OUTDIR, fn)
