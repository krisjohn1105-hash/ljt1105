import sys
import json
import pandas as pd
from datetime import datetime
from pathlib import Path

# Make local-module imports work no matter where the folder is copied to or
# which directory the script is launched from (add this file's folder to path).
BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from api import fetch_eod_report
# from config import ENF_USERNAME, ENF_PASSWORD
# from cleaner import normalize

ENF_USERNAME = 'ops@dunamiscap.com'
ENF_PASSWORD = 'Kingduna10!'

# Run end of day at 5.30 pm HKT.
# If there is no network / the credentials are unavailable, keep going with the
# EOD_Trade.json that was already fetched previously so the script still runs.
try:
    fetch_eod_report(ENF_USERNAME, ENF_PASSWORD)
except Exception as exc:
    print(f"[WARN] Could not fetch EOD report ({exc}); using existing EOD_Trade.json")


# Date
today = datetime.today().strftime('%m%d%y')

# Variables needed to match: QTY; SettleCCY; FX; Net Notional; Commission/taxes; SettleDate; PRICE?

# Keys: Fund; Account; Custondian/Broker; Ticker; TxnType; Asset;

# Load data ------------------------------------------------------------------------
# Set base directory using pathlib (OS-compatible path handling).
# BASE_DIR is defined at the top of the file (this file's folder).
DATA_DIR = BASE_DIR.parent / 'data' / 'input' / 'pre-trade'

EOD_TRADE_PATH = BASE_DIR / "EOD_Trade.json"
if not EOD_TRADE_PATH.exists():
    raise FileNotFoundError(
        f"{EOD_TRADE_PATH} not found. Run api.py to fetch the EOD trade report first."
    )
with open(EOD_TRADE_PATH, "r") as f:
    data = json.load(f)


def load_pre_trade_excel(directory, prefix, date_str):
    """Load the pre-trade file for `date_str`.

    What if one of the files is missing, make sure it still runs:
    fall back to the most recent available file, and if none exist return an
    empty DataFrame instead of crashing.
    """
    preferred = directory / f'{prefix} - {date_str}.xls'
    if preferred.exists():
        return pd.read_excel(preferred)

    candidates = sorted(
        directory.glob(f'{prefix} - *.xls'),
        key=lambda p: p.stat().st_mtime,
    )
    if candidates:
        print(f"[WARN] {preferred.name} not found; using {candidates[-1].name} instead")
        return pd.read_excel(candidates[-1])

    print(f"[WARN] No files matching '{prefix} - *.xls' in {directory}; skipping")
    return pd.DataFrame()


ms_df = load_pre_trade_excel(DATA_DIR, 'Pre allocation Korea Stocks', today)
ms_futures = load_pre_trade_excel(DATA_DIR, 'Pre allocation Korea Futures', today)

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
# key_ms = ['Ric', 'Buy/Sell', 'Listing Ccy', 'Swap Ccy']

# Variables for comparison Qty, net notional
column_list = ['notionalQuantity', 'tradingNotionalNetProceeds']

# Group by different
agg_map = {'notionalQuantity': 'sum',
           'tradingNotionalNetProceeds': 'sum'}

# ---------------------------------------------------------------------------------------

# Transform variables to Enfusion variables
ms_dict = {'Buy/Sell': 'tradeTransactionType', 'Stock Quantity': 'notionalQuantity',
           'Net Notional in Swap Ccy': 'tradingNotionalNetProceeds', 'Ric': 'rIC', 'Swap Ccy': 'settleCurrency',
           'Futures Contracts Quantity': 'notionalQuantity'}

# Combine MS File and rename columns
ms_df = ms_df.rename(columns=ms_dict)
ms_futures = ms_futures.rename(columns=ms_dict)
ms_df = pd.concat([ms_df, ms_futures], ignore_index=True) # Change this later

# Normalize KEY columns
for col in key_enf:
    try:
        enf_df[col] = enf_df[col].astype(str).str.strip()
        ms_df[col] = ms_df[col].astype(str).str.strip()
    except:
        print(col, 'the problems')


# Format Quantity
enf_df['notionalQuantity'] = enf_df['notionalQuantity'].abs()  # Absolute value
enf_df['tradingNotionalNetProceeds'] = enf_df['tradingNotionalNetProceeds'].abs()  # Absolute value


# Format Txn type
enf_df['tradeTransactionType'] = enf_df['tradeTransactionType'].replace({'Buy to Cover': "Buy", 'Sell Short': 'Sell'})

# Aggregate (add the comparing variables together)
enf_df = enf_df.groupby(key_enf, as_index=False).agg(agg_map)
ms_df = ms_df.groupby(key_enf, as_index=False).agg(agg_map)

# Composite Keys
enf_df["_key"] = list(zip(*[enf_df[col] for col in key_enf]))  ## revisit this later
keys_enf = set(enf_df["_key"])

ms_df["_key"] = list(zip(*[ms_df[col] for col in key_enf]))  ## revisit this later
keys_ms = set(ms_df["_key"])

# Generate Missing rows
missing_in_ms = enf_df[enf_df["_key"].isin(keys_enf - keys_ms)][key_enf].copy()  # revisit later
missing_in_ms["issue"] = "In enf_df, missing in ms_df"
missing_in_enf = ms_df[ms_df["_key"].isin(keys_ms - keys_enf)][key_enf].copy()
missing_in_enf = missing_in_enf.rename(columns=dict(zip(key_enf, key_enf)))
missing_in_enf["issue"] = "In ms_df, missing in enf_df"
missing_report = pd.concat([missing_in_ms, missing_in_enf], ignore_index=True)

if missing_report.empty:
    print("-"*20,"No missing report","-"*20)
else:
    print("-"*20,'Missing report',"-"*20)
    print(missing_report)

# Index by composite key
enf_idx = enf_df.set_index(key_enf)
ms_idx = ms_df.set_index(key_enf)

# Compare values for common keys
common_keys = keys_enf & keys_ms
diff_rows = []

for key in common_keys:
    for col in column_list:
        val_1 = enf_idx.loc[key, col] if col in enf_idx.columns else "COLUMN NOT FOUND"
        val_2 = ms_idx.loc[key, col] if col in ms_idx.columns else "COLUMN NOT FOUND"
        if val_1 != val_2:
            try:
                diff = val_1 - val_2
            except TypeError:
                diff = "N/A"
            diff_rows.append({
                "key": key,
                "variable": col,
                "val_enf": val_1,
                "val_ms": val_2,
                "Difference": diff
            })

# Build the report with explicit columns so it is valid even when there are no
# differing rows (otherwise an empty frame has no 'Difference' column).
diff_report = pd.DataFrame(
    diff_rows,
    columns=["key", "variable", "val_enf", "val_ms", "Difference"],
)

# Filter -> Tolerance Level; need to set different tolerance levels
tolerance = 1.0
diff_report['Difference'] = pd.to_numeric(diff_report['Difference'], errors='coerce')
diff_report = diff_report[(diff_report["Difference"].isna()) | (diff_report["Difference"].abs() >= tolerance)]

# Error Handling
if diff_report.empty:
    print("-"*20,"No differences found","-"*20)
else:
    for index in diff_report.index:
        print("-"*20,"Differences found","-"*20)
        print(diff_report.loc[index])

