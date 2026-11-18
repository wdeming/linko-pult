# -*- coding: utf-8 -*-
"""Единый стиль отчётов Linktrade (как в комплекте data room)."""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

NUM  = '#,##0;\\(#,##0\\);\\-'
PCT  = '0.0%'
DEC  = '0.0'
DEC2 = '0.00'

HDR_FILL  = PatternFill('solid', fgColor='FF1F3864')
BLUE      = PatternFill('solid', fgColor='FFD9E2F3')   # POS
GREEN     = PatternFill('solid', fgColor='FFE2EFDA')   # SFA
YELLOW    = PatternFill('solid', fgColor='FFFFF2CC')   # смешанное / выделение
TOTAL     = PatternFill('solid', fgColor='FFFFD966')   # итого
GREY      = PatternFill('solid', fgColor='FFF2F2F2')
THIN      = Border(*[Side(style='thin', color='FFBFBFBF')] * 4)
INPUT_FONT_COLOR = 'FF0000FF'                          # синим — введённые значения

def title(ws, cell, text):
    ws[cell] = text
    ws[cell].font = Font(bold=True, size=14)

def section(ws, cell, text):
    ws[cell] = text
    ws[cell].font = Font(bold=True, size=12)

def header(ws, row, labels, col0=1):
    for i, t in enumerate(labels):
        c = ws.cell(row, col0 + i)
        c.value = t
        c.font = Font(bold=True, size=10, color='FFFFFFFF')
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

def datarow(ws, row, values, fill=None, fmts=None, bold=False, col0=1, blue_inputs=True, wrap=None):
    for i, v in enumerate(values):
        c = ws.cell(row, col0 + i)
        c.value = v
        c.border = THIN
        if fill is not None:
            c.fill = fill
        fmt = (fmts[i] if fmts and i < len(fmts) else None)
        if fmt:
            c.number_format = fmt
        is_input = isinstance(v, (int, float)) and blue_inputs
        c.font = Font(bold=bold, size=11 if bold else 10,
                      color=(INPUT_FONT_COLOR if is_input and not bold else None))
        if isinstance(v, str) and not v.startswith('='):
            c.alignment = Alignment(horizontal='left', wrap_text=bool(wrap))
        else:
            c.alignment = Alignment(horizontal='right' if not isinstance(v, str) or v.startswith('=') else 'left')

def totalrow(ws, row, values, fmts=None, col0=1):
    datarow(ws, row, values, fill=TOTAL, fmts=fmts, bold=True, col0=col0, blue_inputs=False)

def notes(ws, row, lines):
    for i, t in enumerate(lines):
        c = ws.cell(row + i, 1)
        c.value = t
        c.font = Font(size=9, italic=True, color='FF666666')
    return row + len(lines)

def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w
