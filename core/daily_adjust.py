# -*- coding: utf-8 -*-
import io
import openpyxl
from openpyxl.utils import get_column_letter

MAX_COL = 33  # A through AG

# Row 1: group label start columns (1-indexed)
_GROUP_STARTS = {
    5:  'Newinstall',
    13: 'Branding+boosting',
    18: 'Search',
    22: 'Reattribution',
    28: 'Active',
}

# Row 2: column headers (33 items, index 0 = col A)
_COL_HEADERS = [
    'Day', 'Date', 'Total Spent', 'Total Imp',
    'Spent', 'CPI', 'Install', 'Imp', 'Click', 'CTR', 'CVR', 'IR',
    'Spent', 'Imp', 'View', 'CPV', 'CPM',
    'Spent', 'Imp', 'Click', 'CPC',
    'Spent', 'Imp', 'click', 'cpc', 'CPR', 'Reattributions',
    'Spent', 'Imp', 'click', 'cpc', 'CPR', 'Reattributions',
]

# Within-row derived formulas: col_idx → rhs template (use {r} for row number)
_DERIVED = {
    3:  'E{r}+M{r}+R{r}+V{r}+AB{r}',   # C: Total Spent
    4:  'H{r}+N{r}+S{r}+W{r}+AC{r}',   # D: Total Imp
    6:  'E{r}/G{r}',                     # F: CPI
    10: 'I{r}/H{r}',                     # J: CTR
    11: 'G{r}/I{r}',                     # K: CVR
    12: 'G{r}/H{r}',                     # L: IR
    16: 'M{r}/O{r}',                     # P: CPV
    17: 'M{r}/N{r}*1000',               # Q: CPM
    21: 'R{r}/T{r}',                     # U: CPC
    25: 'V{r}/X{r}',                     # Y: cpc (Reattr)
    26: 'V{r}/AA{r}',                    # Z: CPR (Reattr)
    31: 'AB{r}/AD{r}',                   # AE: cpc (Active)
    32: 'AB{r}/AG{r}',                   # AF: CPR (Active)
}


def _eeu_formula(col_idx, row):
    cl = get_column_letter(col_idx)
    r  = row
    if col_idx == 1: return f'=AZ!A{r}'
    if col_idx == 2: return f'=AZ!B{r}'
    if col_idx in _DERIVED: return '=' + _DERIVED[col_idx].format(r=r)

    az  = f"AZ!{cl}{r}"
    eeu = f"'EEU（KZ、KG、BY）'!{cl}{r}"
    eo  = f"EEU_Others!{cl}{r}"
    kz  = f"KZ!{cl}{r}"

    return f'={az}+{eeu}+{eo}+{kz}'


def _sa_formula(col_idx, row):
    cl = get_column_letter(col_idx)
    r  = row
    if col_idx == 1: return f'=SA!A{r}'
    if col_idx == 2: return f'=SA!B{r}'
    if col_idx in _DERIVED: return '=' + _DERIVED[col_idx].format(r=r)
    return f'=SA!{cl}{r}+SA_IOS!{cl}{r}'


def _build_total_sheet(ws, a1_label, formula_fn, ref_ws):
    # Row 1: merged label + group headers
    ws.merge_cells('A1:D1')
    ws['A1'] = a1_label
    for col_idx, label in _GROUP_STARTS.items():
        ws.cell(1, col_idx, label)

    # Row 2: column headers
    for i, header in enumerate(_COL_HEADERS, 1):
        ws.cell(2, i, header)

    # Data rows: same rows as reference sheet
    for r in range(3, ref_ws.max_row + 1):
        if ref_ws.cell(r, 2).value is None:
            continue
        for col_idx in range(1, MAX_COL + 1):
            ws.cell(r, col_idx, formula_fn(col_idx, r))


def process_eeu(file_bytes: bytes) -> bytes:
    """
    1. Rename 'EEU' sheet → 'EEU（KZ、KG、BY）'
    2. Delete 'EEU TOTAL' sheet
    3. Insert 'EEU其他 TOTAL' at position 0 with formulas aggregating AZ + EEU（KZ、KG、BY） + EEU_Others + KZ
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))

    if 'EEU' in wb.sheetnames:
        wb['EEU'].title = 'EEU（KZ、KG、BY）'

    if 'EEU TOTAL' in wb.sheetnames:
        wb.remove(wb['EEU TOTAL'])

    ws_new = wb.create_sheet('EEU其他 TOTAL', 0)
    _build_total_sheet(ws_new, 'CIS TOTAL', _eeu_formula, wb['AZ'])

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def process_mena(file_bytes: bytes) -> bytes:
    """
    Add 'SA TOTAL' sheet at end, aggregating SA + SA_IOS.
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))

    ws_new = wb.create_sheet('SA TOTAL')
    _build_total_sheet(ws_new, 'SA TOTAL', _sa_formula, wb['SA'])

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
