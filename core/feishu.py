# -*- coding: utf-8 -*-
import re
import datetime
import requests

APP_ID      = 'cli_a9745bcb34b99cce'
APP_SECRET  = 'RMLCaKRp6oyqzU5AXv8imfK8kFsgfOKC'
SHEET_TOKEN = 'HUuWs1nQmh96nttM07vcDSFSnUg'

REGION_SHEETS = {
    'EEU':             '5DkwrP',
    'SEA':             'PMHf3k',
    'WEU/NA':          'DQsMUI',
    'MENA':            '3716mM',
    'GLOBAL/SA/LATAM': 'yr1tqZ',
    'AF':              '3RhFeU',
}

DEFAULT_REMARK = '继续保持观察'


def get_token() -> str:
    r = requests.post(
        'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal',
        json={'app_id': APP_ID, 'app_secret': APP_SECRET},
        timeout=15,
    )
    r.raise_for_status()
    d = r.json()
    if d.get('code') != 0:
        raise RuntimeError(f'Feishu auth failed: {d}')
    return d['tenant_access_token']


def read_sheet_values(token: str, sheet_id: str, range_str: str, fallback_on_error: bool = False) -> list:
    url = (f'https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/'
           f'{SHEET_TOKEN}/values/{sheet_id}!{range_str}')
    for attempt in range(3):
        try:
            r = requests.get(url, headers={'Authorization': f'Bearer {token}'}, timeout=90)
            r.raise_for_status()
            return r.json().get('data', {}).get('valueRange', {}).get('values', [])
        except Exception as e:
            if attempt == 2:
                if fallback_on_error:
                    return []
                raise
            import time; time.sleep(2)


def _to_float(v) -> float:
    if v is None:
        return 0.0
    s = str(v)
    if s.startswith('=') or s.upper().startswith('SUBTOTAL') or s.upper().startswith('IF('):
        return 0.0
    # strip commas (locale formatting)
    s = s.replace(',', '')
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def _to_str(v) -> str:
    return str(v).strip() if v is not None else ''


def load_all_data() -> tuple:
    """
    Read all campaign data from all 6 Feishu region sheets.

    Per sheet, reads 3 column ranges:
      A:U  (21 cols) — identifiers + basic metrics
      V:AT (25 cols) — extended metrics (CPR, CVR, UU, Freq, Impress, Reg, Login)
      AW   ( 1 col ) — AI建议

    Returns:
      all_campaigns: list of dicts, one per campaign row (E != 'Total')
      country_total_ai: {(a_region, b_country, c_category): ai_text}
    """
    token = get_token()
    all_campaigns: list = []
    country_total_ai: dict = {}

    for _sheet_key, sid in REGION_SHEETS.items():
        rows_au  = read_sheet_values(token, sid, 'A1:U200')
        rows_vat = read_sheet_values(token, sid, 'V1:AT200', fallback_on_error=True)
        rows_aw  = read_sheet_values(token, sid, 'AW1:AW200')

        n = min(len(rows_au), len(rows_aw))

        for i in range(1, n):
            row    = rows_au[i]
            ext    = rows_vat[i] if i < len(rows_vat) else []
            aw_row = rows_aw[i]

            a_region   = _to_str(row[0]  if len(row) >  0 else None)
            b_country  = _to_str(row[1]  if len(row) >  1 else None)
            c_category = _to_str(row[2]  if len(row) >  2 else None)
            d_code     = _to_str(row[3]  if len(row) >  3 else None)
            e_name     = _to_str(row[4]  if len(row) >  4 else None)
            ai         = _to_str(aw_row[0] if aw_row else None)

            if not a_region or not c_category:
                continue

            if e_name == 'Total':
                key = (a_region, b_country, c_category)
                if key not in country_total_ai:
                    country_total_ai[key] = ai
                continue

            if not d_code or not e_name:
                continue

            # V:AT column offsets within rows_vat (V=0, W=1, ..., AT=24):
            #  W(1)=CPR_reg_y  X(2)=CPR_reg_t
            #  Z(4)=CPR_login_y  AA(5)=CPR_login_t
            #  AC(7)=CVR_y   AD(8)=CVR_t
            #  AF(10)=UU_y   AG(11)=UU_t
            #  AI(13)=Freq_y AJ(14)=Freq_t
            #  AL(16)=Impress_y AM(17)=Impress_t
            #  AO(19)=Reg_y  AP(20)=Reg_t
            #  AR(22)=Login_y  AS(23)=Login_t

            def _ext(idx):
                return _to_float(ext[idx] if len(ext) > idx else None)

            all_campaigns.append({
                'a_region':   a_region,
                'b_country':  b_country,
                'c_category': c_category,
                'd_code':     d_code,
                'e_name':     e_name,
                'budget':     _to_float(row[5]  if len(row) >  5 else None),
                'budget_t':   _to_float(row[6]  if len(row) >  6 else None),
                'cost_y':     _to_float(row[7]  if len(row) >  7 else None),
                'cost_t':     _to_float(row[8]  if len(row) >  8 else None),
                'clicks_y':   _to_float(row[10] if len(row) > 10 else None),
                'clicks_t':   _to_float(row[11] if len(row) > 11 else None),
                'install_y':  _to_float(row[13] if len(row) > 13 else None),
                'install_t':  _to_float(row[14] if len(row) > 14 else None),
                'cpi_y':      _to_float(row[19] if len(row) > 19 else None),
                'cpi_t':      _to_float(row[20] if len(row) > 20 else None),
                # extended
                'uu_y':       _ext(10),
                'uu_t':       _ext(11),
                'impress_y':  _ext(16),
                'impress_t':  _ext(17),
                'reg_y':      _ext(19),
                'reg_t':      _ext(20),
                'login_y':    _ext(22),
                'login_t':    _ext(23),
                'ai':         ai,
            })

    return all_campaigns, country_total_ai


def get_campaign_period() -> tuple:
    """
    Read spreadsheet metainfo, find the sheet named like 'Campaign对比 09-11 17点',
    and return (date_str, period) e.g. ('2026-09-11', '17').
    Falls back to (None, None) if not found.
    """
    try:
        token = get_token()
        url = (f'https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/'
               f'{SHEET_TOKEN}/metainfo')
        r = requests.get(url, headers={'Authorization': f'Bearer {token}'}, timeout=15)
        r.raise_for_status()
        sheets = r.json().get('data', {}).get('sheets', [])
        pat = re.compile(r'Campaign对比\s+(\d{1,2}-\d{1,2})\s+(\d{1,2})点')
        for sheet in sheets:
            m = pat.search(sheet.get('title', ''))
            if m:
                md = m.group(1).split('-')
                month_day = f'{int(md[0]):02d}-{int(md[1]):02d}'
                hour = int(m.group(2))
                period = '17' if hour >= 13 else '09'
                year = datetime.date.today().year
                date_str = f'{year}-{month_day}'
                return date_str, period
    except Exception:
        pass
    return None, None
