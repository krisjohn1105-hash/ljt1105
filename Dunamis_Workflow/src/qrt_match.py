import json
import pandas as pd
from datetime import datetime
from pathlib import Path
from api import fetch_eod_report

# from config import ENF_USERNAME, ENF_PASSWORD

ENF_USERNAME = 'ops@dunamiscap.com'
ENF_PASSWORD = 'Kingduna10!'

# Run end of day at 5.30 pm HKT
fetch_eod_report(ENF_USERNAME, ENF_PASSWORD)


# Date
today = datetime.today().strftime('%m%d%y')          # e.g. 060126 (used for logging only)
today_qrt = datetime.today().strftime('%Y%m%d')      # e.g. 20260601 (QRT filename date prefix)

# Variables needed to match: QTY; SettleCCY; FX; Net Notional; Commission/taxes; SettleDate; PRICE?
# Keys: Fund; Account; Custodian/Broker; Ticker; TxnType; Asset;

# Load data ------------------------------------------------------------------------
# Set base directory using pathlib (OS-compatible path handling)
BASE_DIR = Path(__file__).resolve().parent
QRT_DIR = BASE_DIR.parent / 'data' / 'input' / 'QRT'


def find_qrt_file(qrt_dir: Path, date_str: str) -> Path:
    """Locate today's QRT trade file.

    QRT delivers files named like ``Qube_Dunamis_UAT_Trades_<YYYYMMDDhhmmss>.csv``.
    We first try to match today's date (YYYYMMDD); if none is found we fall back to
    the most recently modified CSV in the folder so the reconciliation still runs.
    """
    csvs = sorted(qrt_dir.glob('*.csv'))
    if not csvs:
        raise FileNotFoundError(f'No QRT CSV files found in {qrt_dir}')

    todays = [p for p in csvs if date_str in p.name]
    if todays:
        # Filename timestamp sorts chronologically -> last is the latest run for the day
        return sorted(todays)[-1]

    latest = max(csvs, key=lambda p: p.stat().st_mtime)
    print('-' * 20, f'No QRT file for {date_str}; using latest: {latest.name}', '-' * 20)
    return latest


qrt_path = find_qrt_file(QRT_DIR, today_qrt)
print('-' * 20, f'Using QRT file: {qrt_path.name}', '-' * 20)

with open(BASE_DIR / "EOD_Trade.json", "r") as f:
    data = json.load(f)

# QRT trade file (CSV)
qrt_df = pd.read_csv(qrt_path)

## Then convert to DataFrame as usual
enf_df = pd.json_normalize(data['rows'])

## Columns I need
enf_df = enf_df[['bBYellow.value', 'rIC.value', 'tradeTransactionType.value', 'notionalQuantity.value', 'tradePrice.value',
                 'tradePrice.currency',
                 'instrumentDescription.value', 'tradePrice.quoteType', 'tradePrice.quotationFormat',
                 'instrumentSubtype.value', 'tradeDate.value', 'settleDate.value', 'tradeCurrency.value',
                 'settleCurrency.value', 'actualTradeToSettleFXRate.value', 'grossTradeCommisions.value',
                 'grossTradeCommisions.currency', 'grossFees.value', 'grossTaxes.value',
                 'netTradePrice.value', 'netTradePrice.currency',
                 'tradingNotionalNetProceeds.value',
                 ]]

# Formatting column names: tradingNotionalNetProceeds.currency
enf_df.columns = enf_df.apply(lambda x: x.name.split('.')[0])  # Change column names

# ------CONFIG----------------------------------------------------------------------------

# Keys for comparison
key_enf = ['rIC', 'tradeTransactionType', 'settleCurrency']

# Variables for comparison Qty, net notional
column_list = ['notionalQuantity', 'tradingNotionalNetProceeds']

# Group by different
agg_map = {'notionalQuantity': 'sum',
           'tradingNotionalNetProceeds': 'sum'}

# ---------------------------------------------------------------------------------------

# Transform QRT variables to Enfusion variables
qrt_dict = {'RIC': 'rIC',
            'Side': 'tradeTransactionType',
            'Settle Currency': 'settleCurrency',
            'Trade Quantity': 'notionalQuantity',
            'Net Consideration': 'tradingNotionalNetProceeds'}

# Rename QRT columns and keep only the ones we need
qrt_df = qrt_df.rename(columns=qrt_dict)
qrt_df = qrt_df[key_enf + column_list].copy()

# Format Txn type
# Enfusion uses Buy / Sell / Sell Short / Buy to Cover -> collapse to Buy / Sell
enf_df['tradeTransactionType'] = enf_df['tradeTransactionType'].replace({'Buy to Cover': "Buy", 'Sell Short': 'Sell'})
# QRT uses single-letter Side codes -> Buy / Sell
qrt_df['tradeTransactionType'] = qrt_df['tradeTransactionType'].replace(
    {'B': 'Buy', 'S': 'Sell', 'SS': 'Sell', 'BC': 'Buy'})

# Normalize KEY columns
for col in key_enf:
    try:
        enf_df[col] = enf_df[col].astype(str).str.strip()
        qrt_df[col] = qrt_df[col].astype(str).str.strip()
    except:
        print(col, 'the problems')

# Format comparison variables (absolute value -> direction lives in the key, not the sign)
for col in column_list:
    enf_df[col] = pd.to_numeric(enf_df[col], errors='coerce').abs()
    qrt_df[col] = pd.to_numeric(qrt_df[col], errors='coerce').abs()

# Aggregate (add the comparing variables together)
enf_df = enf_df.groupby(key_enf, as_index=False).agg(agg_map)
qrt_df = qrt_df.groupby(key_enf, as_index=False).agg(agg_map)

# Composite Keys
enf_df["_key"] = list(zip(*[enf_df[col] for col in key_enf]))
keys_enf = set(enf_df["_key"])

qrt_df["_key"] = list(zip(*[qrt_df[col] for col in key_enf]))
keys_qrt = set(qrt_df["_key"])

# Generate Missing rows
missing_in_qrt = enf_df[enf_df["_key"].isin(keys_enf - keys_qrt)][key_enf].copy()
missing_in_qrt["issue"] = "In enf_df, missing in qrt_df"
missing_in_enf = qrt_df[qrt_df["_key"].isin(keys_qrt - keys_enf)][key_enf].copy()
missing_in_enf["issue"] = "In qrt_df, missing in enf_df"
missing_report = pd.concat([missing_in_qrt, missing_in_enf], ignore_index=True)

if missing_report.empty:
    print("-" * 20, "No missing report", "-" * 20)
else:
    print("-" * 20, 'Missing report', "-" * 20)
    print(missing_report)

# Index by composite key
enf_idx = enf_df.set_index(key_enf)
qrt_idx = qrt_df.set_index(key_enf)

# Compare values for common keys
common_keys = keys_enf & keys_qrt
diff_rows = []

for key in common_keys:
    for col in column_list:
        val_1 = enf_idx.loc[key, col] if col in enf_idx.columns else "COLUMN NOT FOUND"
        val_2 = qrt_idx.loc[key, col] if col in qrt_idx.columns else "COLUMN NOT FOUND"
        if val_1 != val_2:
            try:
                diff = val_1 - val_2
            except TypeError:
                diff = "N/A"
            diff_rows.append({
                "key": key,
                "variable": col,
                "val_enf": val_1,
                "val_qrt": val_2,
                "Difference": diff
            })

diff_report = pd.DataFrame(diff_rows)

# Filter -> Tolerance Level; need to set different tolerance levels
tolerance = 1.0
if not diff_report.empty:
    diff_report['Difference'] = pd.to_numeric(diff_report['Difference'], errors='coerce')
    diff_report = diff_report[(diff_report["Difference"].isna()) | (diff_report["Difference"].abs() >= tolerance)]

# Error Handling
if diff_report.empty:
    print("-" * 20, "No differences found", "-" * 20)
else:
    for index in diff_report.index:
        print("-" * 20, "Differences found", "-" * 20)
        print(diff_report.loc[index])
