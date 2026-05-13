"""
process_data.py – Processer alle Excel-filer fra Finanstilsynet/Norges Bank
og genererer JSON-filer til bankanalyse-dashboardet.

Kjøres én gang (eller når nye filer lastes ned):
    python scripts/process_data.py

Legg Excel-filene i mappen 'excel_data/' (eller juster UPLOADS-stien under).
"""

import pandas as pd
import numpy as np
import json
import os
import sys
from pathlib import Path

# ── Stier ──────────────────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).parent
ROOT = SCRIPT_DIR.parent
UPLOADS = ROOT / 'excel_data'          # Legg Excel-filene her
OUT = ROOT / 'data'
OUT.mkdir(exist_ok=True)

# ── 1. Norges Bank styringsrente (skrapt direkte fra NB 2026-05-13) ────────
# Kilde: https://www.norges-bank.no/en/topics/monetary-policy/policy-rate/
NB_RATE_DECISIONS = [
    ("2000-04-13", 5.75), ("2000-06-15", 6.25), ("2000-08-10", 6.75), ("2000-09-21", 7.00),
    ("2001-01-10", 7.00), ("2001-02-21", 7.00), ("2001-10-31", 7.00), ("2001-12-12", 6.50),
    ("2002-01-23", 6.50), ("2002-04-10", 6.50), ("2002-05-22", 6.50), ("2002-07-03", 7.00),
    ("2002-08-07", 7.00), ("2002-10-30", 7.00), ("2002-12-11", 6.50),
    ("2003-01-22", 6.00), ("2003-03-05", 5.50), ("2003-04-30", 5.00), ("2003-06-25", 4.00),
    ("2003-08-13", 3.00), ("2003-09-17", 2.50), ("2003-10-29", 2.50), ("2003-12-17", 2.25),
    ("2004-01-28", 2.00), ("2004-03-11", 1.75),
    ("2005-02-02", 1.75), ("2005-05-25", 1.75), ("2005-09-21", 2.00), ("2005-12-14", 2.25),
    ("2006-01-25", 2.25), ("2006-03-16", 2.50), ("2006-04-26", 2.50), ("2006-06-29", 2.75),
    ("2006-09-27", 3.00), ("2006-11-01", 3.25), ("2006-12-13", 3.50),
    ("2007-01-24", 3.75), ("2007-04-25", 4.00), ("2007-06-27", 4.50), ("2007-08-15", 4.75),
    ("2007-10-31", 5.00), ("2007-12-12", 5.25),
    ("2008-01-23", 5.25), ("2008-06-25", 5.75), ("2008-10-15", 5.25), ("2008-10-29", 4.75),
    ("2008-12-17", 3.00),
    ("2009-02-04", 2.50), ("2009-03-25", 2.00), ("2009-05-06", 1.50), ("2009-06-17", 1.25),
    ("2009-10-28", 1.50), ("2009-12-16", 1.75),
    ("2010-02-03", 1.75), ("2010-03-24", 1.75), ("2010-05-05", 2.00),
    ("2011-03-16", 2.00), ("2011-05-12", 2.25), ("2011-12-14", 1.75),
    ("2012-03-14", 1.50),
    ("2014-10-23", 1.50), ("2014-12-11", 1.25),
    ("2015-06-18", 1.00), ("2015-09-24", 0.75), ("2015-12-17", 0.75),
    ("2016-03-17", 0.50),
    ("2018-09-20", 0.75),
    ("2019-03-21", 1.00), ("2019-08-15", 1.25), ("2019-09-19", 1.50),
    ("2020-03-13", 1.00), ("2020-03-20", 0.25), ("2020-05-07", 0.00),
    ("2021-09-23", 0.25), ("2021-12-16", 0.50),
    ("2022-01-20", 0.50), ("2022-03-24", 0.75), ("2022-05-05", 0.75),
    ("2022-06-23", 1.25), ("2022-08-18", 1.75), ("2022-09-22", 2.25),
    ("2022-11-03", 2.50), ("2022-12-15", 2.75),
    ("2023-01-19", 2.75), ("2023-03-23", 3.00), ("2023-05-04", 3.25),
    ("2023-06-22", 3.75), ("2023-08-17", 4.00), ("2023-11-02", 4.25), ("2023-12-14", 4.50),
    ("2024-01-25", 4.50),
    ("2025-01-23", 4.50), ("2025-08-14", 4.25), ("2025-09-18", 4.00), ("2025-12-18", 4.00),
    ("2026-01-22", 4.00), ("2026-03-26", 4.00), ("2026-05-07", 4.25),
]

# Fremskriving: NB MPR 1/2026 (mars 2026) – intervallmidtpunkt
NB_FORECAST = [
    ("2026-12-31", 4.375),
    ("2027-12-31", 4.00),
    ("2028-12-31", 3.50),
    ("2029-12-31", 3.00),
]

def build_annual_rates():
    dates = pd.date_range('2000-01-01', '2026-12-31', freq='D')
    rate_series = pd.Series(index=dates, dtype=float)
    rate_series.iloc[0] = 6.50
    for ds, rate in NB_RATE_DECISIONS:
        d = pd.Timestamp(ds)
        if d >= dates[0] and d <= dates[-1]:
            rate_series.loc[d:] = rate
    rate_series.ffill(inplace=True)
    rate_series.bfill(inplace=True)
    annual = {}
    for yr in range(2000, 2027):
        s = rate_series[rate_series.index.year == yr]
        if len(s) > 0:
            annual[yr] = round(float(s.mean()), 4)
    return annual

ANNUAL_RATES = build_annual_rates()

# ── 2. Parser: Sparebankenes årsregnskaper 2015-2023 ───────────────────────
def parse_sparebank_tab3(filepath, year):
    try:
        df = pd.read_excel(filepath, sheet_name='Tab3', header=None)
    except Exception as e:
        print(f"  Tab3 error in {filepath.name}: {e}")
        return []

    # Finn header-rad (inneholder 'Banks' eller 'LEI')
    hrow = None
    for i in range(12):
        row = df.iloc[i, :].astype(str).str.strip()
        if row.isin(['Banks', 'LEI']).any():
            hrow = i
            break
    if hrow is None:
        return []

    headers = [str(h).replace('\n', ' ').lower().strip() for h in df.iloc[hrow, :].tolist()]

    col = lambda *terms: next((ci for ci, h in enumerate(headers)
                               if all(t in h for t in terms)), None)

    col_roe          = col('return on equity') or col('egenkapitalrentabilitet')
    col_cost         = col('operating expenses as %') or col('driftskostnader i %')
    col_asset_g      = col('total assets (%)') or col('forvaltnings')
    col_loan_g       = col('gross loans') or col('utlån til kunder (%)')
    col_dep_g        = col('deposits from customers (%)')
    col_dep_cov      = col('deposits from customers as %') or col('innskudd fra kunder i %')
    col_cet1         = col('core capital  (cet1)') or col('ren kjerne')
    col_t1           = col('tier 1- capital') or col('kjernekapital')
    col_lev          = col('leverage') or col('uvektet')
    col_ppm          = col('operating profit before losses') or col('driftsresultat')
    col_npm          = col('profit for the year') or col('resultat for regnskaps')
    col_rwa          = col('risk weighted assets') or col('beregningsgrunnlag')
    col_emp          = col('employees') or col('årsverk')

    # Finn første datrad
    data_start = hrow + 1
    for i in range(hrow + 1, min(hrow + 5, len(df))):
        v = str(df.iloc[i, 0]).strip()
        if v.isdigit():
            data_start = i
            break

    records = []
    for i in range(data_start, len(df)):
        row = df.iloc[i]
        idx = str(row.iloc[0]).strip()
        bank = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else None
        if not bank or bank in ('nan', 'NaN', 'Banks', 'Banker', 'None', ''):
            continue
        if not idx or not idx.isdigit():
            continue

        def get(c): 
            if c is None: return None
            v = row.iloc[c]
            return float(v) if pd.notna(v) and v != '' else None

        rec = {
            'bank': bank, 'year': year,
            'lei': str(row.iloc[2]) if pd.notna(row.iloc[2]) else None,
            'roe': get(col_roe),
            'cost_ratio': get(col_cost),
            'asset_growth': get(col_asset_g),
            'loan_growth': get(col_loan_g),
            'deposit_growth': get(col_dep_g),
            'deposit_coverage': get(col_dep_cov),
            'cet1': get(col_cet1),
            'tier1': get(col_t1),
            'leverage': get(col_lev),
            'pre_provision_margin': get(col_ppm),
            'net_profit_margin': get(col_npm),
            'rwa': get(col_rwa),
            'employees': get(col_emp),
        }
        records.append(rec)
    return records


def parse_sparebank_tab1(filepath, year):
    try:
        df = pd.read_excel(filepath, sheet_name='Tab1', header=None)
    except:
        return {}

    hrow = None
    for i in range(12):
        row = df.iloc[i, :].astype(str)
        if row.str.contains('Average total assets|Gjennomsnittlig forvaltnings', na=False).any():
            hrow = i
            break
    if hrow is None:
        return {}

    headers = [str(h).replace('\n', ' ').lower().strip() for h in df.iloc[hrow, :].tolist()]
    col_assets = next((ci for ci, h in enumerate(headers) if 'average total assets' in h or 'gjennomsnittlig' in h), None)
    col_income = next((ci for ci, h in enumerate(headers) if h in ('income', 'inntekter') or 'inn-' in h), None)
    col_costs  = next((ci for ci, h in enumerate(headers) if h in ('expenses', 'kostnader') or 'kost-' in h), None)

    data_start = hrow + 1
    for i in range(hrow + 1, min(hrow + 4, len(df))):
        if str(df.iloc[i, 0]).strip().isdigit():
            data_start = i
            break

    nim_data = {}
    for i in range(data_start, len(df)):
        row = df.iloc[i]
        if not str(row.iloc[0]).strip().isdigit():
            continue
        bank = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else None
        if not bank or bank in ('nan', 'NaN', 'None', 'Banks', ''):
            continue

        avg_assets = float(row.iloc[col_assets]) if col_assets and pd.notna(row.iloc[col_assets]) else None
        ii = float(row.iloc[col_income]) if col_income and pd.notna(row.iloc[col_income]) else None
        ie = float(row.iloc[col_costs]) if col_costs and pd.notna(row.iloc[col_costs]) else None

        nim = None
        net_ii = None
        if avg_assets and avg_assets > 0 and ii is not None and ie is not None:
            net_ii = ii - ie
            nim = net_ii / avg_assets

        nim_data[bank] = {
            'avg_assets': avg_assets,
            'interest_income': ii,
            'interest_expense': ie,
            'net_interest_income': net_ii,
            'nim': nim,
        }
    return nim_data


def parse_sparebank_tab2(filepath, year):
    try:
        df = pd.read_excel(filepath, sheet_name='Tab2', header=None)
    except:
        return {}

    hrow = None
    for i in range(12):
        row = df.iloc[i, :].astype(str)
        if row.str.contains('Total assets|Forvaltnings-kapital', na=False).any():
            hrow = i
            break
    if hrow is None:
        return {}

    headers = [str(h).replace('\n', ' ').lower().strip() for h in df.iloc[hrow, :].tolist()]
    col_ta   = next((ci for ci, h in enumerate(headers) if 'total assets' in h or 'forvaltnings-kapital' in h), None)
    col_gl   = next((ci for ci, h in enumerate(headers) if 'gross loans to customers' in h or 'netto utlån til kunder' in h), None)
    col_dep  = next((ci for ci, h in enumerate(headers) if 'deposits from customers' in h or 'innskudd fra kunder' in h), None)
    col_eq   = next((ci for ci, h in enumerate(headers) if 'equity' in h or 'egenkapital' in h), None)

    data_start = hrow + 1
    for i in range(hrow + 1, min(hrow + 4, len(df))):
        if str(df.iloc[i, 0]).strip().isdigit():
            data_start = i
            break

    bs = {}
    for i in range(data_start, len(df)):
        row = df.iloc[i]
        if not str(row.iloc[0]).strip().isdigit():
            continue
        bank = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else None
        if not bank or bank in ('nan', 'NaN', 'None', 'Banks', ''):
            continue

        def g(c): return float(row.iloc[c]) if c and pd.notna(row.iloc[c]) else None
        bs[bank] = {
            'total_assets': g(col_ta),
            'gross_loans': g(col_gl),
            'deposits': g(col_dep),
            'equity': g(col_eq),
        }
    return bs


# ── 3. Parser: Analytikerrapport 2024/2025 ───────────────────────────────
METRIC_MAP_2024 = {
    'SUM EIENDELER': 'total_assets',
    'NETTO RENTEINNTEKTER': 'net_interest_income',
    'Renteinntekter og lignende inntekter': 'interest_income',
    'Rentekostnader og lignende kostnader': 'interest_expense',
    'SUM DRIFTSINNTEKTER': 'total_income',
    'SUM DRIFTSKOSTNADER': 'total_costs',
    'RESULTAT FØR SKATT': 'pre_tax_profit',
    'SUM INNSKUDD OG ANDRE INNLÅN FRA KUNDER': 'deposits',
    'Brutto utlån til Personmarkedet ': 'retail_loans',
    'Brutto utlån til næringsmarkedet ': 'corporate_loans',
    'Sum brutto utlån til og fordringer på kunder (inkl. opptjente ikke betalte renter)': 'gross_loans',
    'NETTO UTLÅN TIL KUNDER': 'net_loans',
    'Kredittap på utlån, garantier og rentebærende verdipapirer': 'loan_losses',
    'Ren kjernekapitaldekning': 'cet1',
    'Kjernekapitaldekning': 'tier1',
    'Kapitaldekning': 'capital_adequacy',
    'Uvektet kjernekapitalandel': 'leverage',
    'Samlet beregningsgrunnlag (RWA)': 'rwa',
    'SUM EGENKAPITAL': 'equity',
    'Antall årsverk': 'employees',
    'Lønn og personalkostnader': 'staff_costs',
    'Andre driftskostnader': 'other_costs',
    'Brutto utlån til kunder, trinn 3 *)': 'stage3_loans',
    'TOTALRESULTAT HITTIL I ÅR': 'total_result',
    'LCR': 'lcr',
    'NSFR': 'nsfr',
    'Forretningskapital': 'business_capital',
}


def parse_analytikerrapport(filepath, year):
    try:
        xl = pd.ExcelFile(filepath)
        sheet = next((s for s in xl.sheet_names if 'Analytiker' in s), xl.sheet_names[0])
        df = pd.read_excel(filepath, sheet_name=sheet, header=None)
    except Exception as e:
        print(f"  Error {filepath.name}: {e}")
        return []

    # Finn rad med banknavn
    bank_row = None
    for i in range(5, 15):
        row = df.iloc[i, :].astype(str)
        if 'Bankens navn' in row.values:
            bank_row = i
            break
    if bank_row is None:
        bank_row = 9

    bank_names = df.iloc[bank_row, 3:].tolist()
    records = []

    for bi, bname in enumerate(bank_names):
        if pd.isna(bname) or str(bname).strip() in ('nan', ''):
            continue
        col_idx = 3 + bi
        if col_idx >= df.shape[1]:
            break

        rec = {'bank': str(bname).strip(), 'year': year, 'lei': None}

        # Lei fra rad 8
        lei_row = df.iloc[8, 3:].tolist()
        if bi < len(lei_row) and pd.notna(lei_row[bi]):
            rec['lei'] = str(lei_row[bi]).strip()

        for ri in range(bank_row + 1, len(df)):
            label = df.iloc[ri, 2]
            if pd.isna(label):
                continue
            label_str = str(label).strip()
            if label_str in METRIC_MAP_2024:
                key = METRIC_MAP_2024[label_str]
                val = df.iloc[ri, col_idx]
                rec[key] = float(val) if pd.notna(val) and val != '' else None

        # Beregnede nøkkeltall
        ta  = rec.get('total_assets')
        ni  = rec.get('net_interest_income')
        tc  = rec.get('total_costs')
        ti  = rec.get('total_income')
        ptp = rec.get('pre_tax_profit')
        eq  = rec.get('equity')
        gl  = rec.get('gross_loans')
        dep = rec.get('deposits')
        ll  = rec.get('loan_losses')
        ii  = rec.get('interest_income')
        ie  = rec.get('interest_expense')

        if ta and ta > 0 and ni:
            rec['nim'] = ni / ta
        if tc and ti and ti != 0:
            rec['cost_ratio'] = abs(tc) / abs(ti)
        if ptp and eq and eq > 0:
            rec['roe'] = ptp / eq
        if dep and gl and gl > 0:
            rec['deposit_coverage'] = dep / gl
        if ll and gl and gl > 0:
            rec['loss_rate'] = abs(ll) / gl
        if ii and ie:
            rec['net_interest_income'] = rec.get('net_interest_income') or (ii - ie)
        records.append(rec)
    return records


# ── 4. Kjør prosessering ─────────────────────────────────────────────────
def main():
    all_records = []

    sparebank_files = {
        2015: 'sparebankenes-arsregnskaper---annual-accounts-for-savings-banks-2015__3_.xlsx',
        2016: 'Sparebankenes_årsregnskaper_2016__4_.xlsx',
        2017: 'Sparebankenes_årsregnskaper_2017__3_.xlsx',
        2018: 'Sparebankenes_årsregnskaper_2018__3_.xlsx',
        2019: 'Sparebankenes_årsregnskaper_2019__3_.xlsx',
        2020: 'Sparebankenes_årsregnskaper_2020__3_.xlsx',
        2021: 'Sparebankenes_årsregnskaper_2021__3_.xlsx',
        2022: 'Sparebankenes_årsregnskaper_2022__3_.xlsx',
        2023: 'Sparebankenes_årsregnskaper_2023__7_.xlsx',
    }

    for year, fname in sparebank_files.items():
        fpath = UPLOADS / fname
        if not fpath.exists():
            print(f"  ⚠️  Mangler: {fname}")
            continue
        print(f"  Prosesserer {year}...", end=' ')
        tab3 = parse_sparebank_tab3(fpath, year)
        nim  = parse_sparebank_tab1(fpath, year)
        bs   = parse_sparebank_tab2(fpath, year)
        for rec in tab3:
            b = rec['bank']
            if b in nim:
                for k, v in nim[b].items():
                    if v is not None and rec.get(k) is None:
                        rec[k] = v
            if b in bs:
                for k, v in bs[b].items():
                    if v is not None and rec.get(k) is None:
                        rec[k] = v
        all_records.extend(tab3)
        print(f"{len(tab3)} banker")

    for year, fname in [
        (2024, 'Regnskapstall_for_norske_banker_for_året_2024__2_.xlsx'),
        (2025, 'Regnskapstall_for_norske_banker_for_året_2025.xlsx'),
    ]:
        fpath = UPLOADS / fname
        if not fpath.exists():
            print(f"  ⚠️  Mangler: {fname}")
            continue
        print(f"  Prosesserer {year}...", end=' ')
        recs = parse_analytikerrapport(fpath, year)
        all_records.extend(recs)
        print(f"{len(recs)} banker")

    if not all_records:
        print("❌ Ingen data funnet. Sjekk at Excel-filene er i excel_data/")
        sys.exit(1)

    df = pd.DataFrame(all_records)
    print(f"\n📊 Totalt {len(df)} rader, {df['bank'].nunique()} unike banker, "
          f"år {sorted(df['year'].unique().tolist())}")

    # Normaliser prosentverdier (decimal-form: 0.08 = 8%)
    pct_cols = ['roe', 'cost_ratio', 'nim', 'cet1', 'tier1', 'capital_adequacy',
                'leverage', 'deposit_coverage', 'loss_rate', 'asset_growth',
                'loan_growth', 'deposit_growth', 'pre_provision_margin', 'net_profit_margin']
    for col in pct_cols:
        if col not in df.columns:
            continue
        vals = df[col].dropna()
        if len(vals) == 0:
            continue
        # Hvis >70% av verdiene er >1, er de sannsynligvis i prosentform (ikke desimal)
        if (vals.abs() > 1).sum() / len(vals) > 0.7:
            df[col] = df[col] / 100.0

    # Segmentering basert på siste tilgjengelige forvaltningskapital
    latest_ta = df.sort_values('year').groupby('bank')['total_assets'].last()
    def seg(ta):
        if pd.isna(ta): return 'unknown'
        return 'large' if ta >= 50_000_000 else 'medium' if ta >= 5_000_000 else 'small'
    df['segment'] = df['bank'].map(lambda b: seg(latest_ta.get(b)))

    # Legg til styringsrente
    df['policy_rate'] = df['year'].map(ANNUAL_RATES)

    # Serialisering
    def safe(v):
        if isinstance(v, float) and (np.isnan(v) or np.isinf(v)): return None
        if isinstance(v, (np.integer,)): return int(v)
        if isinstance(v, (np.floating,)): return float(v)
        return v

    records_out = [{k: safe(v) for k, v in row.items()} for _, row in df.iterrows()]

    with open(OUT / 'banks_data.json', 'w', encoding='utf-8') as f:
        json.dump(records_out, f, ensure_ascii=False, separators=(',', ':'))

    rate_data = {
        'annual': ANNUAL_RATES,
        'decisions': NB_RATE_DECISIONS,
        'forecast': NB_FORECAST,
        'source': 'Norges Bank (skrapt 2026-05-13)',
        'current_rate': 4.25,
        'current_date': '2026-05-07',
    }
    with open(OUT / 'policy_rates.json', 'w') as f:
        json.dump(rate_data, f, indent=2)

    summary = {
        'total_records': len(records_out),
        'years': sorted(df['year'].unique().tolist()),
        'bank_count': int(df['bank'].nunique()),
        'banks': sorted(df['bank'].unique().tolist()),
        'segments': {k: int(v) for k, v in df.groupby('segment')['bank'].nunique().items()},
        'columns': df.columns.tolist(),
        'generated': '2026-05-13',
    }
    with open(OUT / 'summary.json', 'w', encoding='utf-8') as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    print(f"\n✅ Eksportert til data/")
    print(f"   banks_data.json    – {len(records_out)} poster")
    print(f"   policy_rates.json  – {len(NB_RATE_DECISIONS)} rentebeslutninger (2000–2026)")
    print(f"   summary.json")


if __name__ == '__main__':
    main()
