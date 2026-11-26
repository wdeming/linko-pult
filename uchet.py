# -*- coding: utf-8 -*-
"""Замена «Odoo» на «учётная система (компании)» во всех текстах, которые видит пользователь пульта (26.11.2026)."""
import re
_RULES = [(re.compile(r'Odoo linktrade \(odoo\.linko\.uz\)|Odoo \(odoo\.linko\.uz\)|odoo\.linko\.uz'), 'учётная система компании'),
          (re.compile(r'Odoo linktrade'), 'учётная система компании'),
          (re.compile(r'(журнал[а-яё]* изменений MRR) (?:в )?Odoo'), r'\1 учётной системы'),
          (re.compile(r'\bВ Odoo'), 'В учётной системе'), (re.compile(r'(?<![А-Яа-яЁё])в Odoo'), 'в учётной системе'),
          (re.compile(r'(?<![А-Яа-яЁё])из Odoo'), 'из учётной системы'), (re.compile(r'(?<![А-Яа-яЁё])по Odoo'), 'по учётной системе'),
          (re.compile(r'(?<![А-Яа-яЁё])с Odoo'), 'с учётной системой')]
def _rest(m):
    s, i = m.string, m.start(); pre = s[max(0, i - 2):i]
    if i == 0 or pre.endswith(('. ', '\n', '>')): return 'Учётная система'
    if pre.endswith(('— ', ': ', '(', '«', '/ ')): return 'учётная система'
    return 'учётной системы'          # после существительного: «Данные Odoo», «реестра Odoo», «MRR Odoo»
def uchet(s):
    if not isinstance(s, str) or 'Odoo' not in s and 'odoo.linko' not in s: return s
    for rx, rp in _RULES: s = rx.sub(rp, s)
    return re.sub(r'Odoo', _rest, s)
def walk(o):
    if isinstance(o, str): return uchet(o)
    if isinstance(o, list): return [walk(x) for x in o]
    if isinstance(o, dict): return {k: walk(v) for k, v in o.items()}
    return o
def xlsx(body):
    import io, openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(body)); ch = False
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and ('Odoo' in c.value or 'odoo.linko' in c.value) and not c.value.startswith('='): c.value = uchet(c.value); ch = True
    if not ch: return body
    out = io.BytesIO(); wb.save(out); return out.getvalue()
