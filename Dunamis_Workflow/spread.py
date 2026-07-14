# import msal
import json
import pandas as pd
import datetime as dt
import os
import glob

with open("spread.json", "r") as f:
    data = json.load(f)

# Data Cleaning required
files = glob.glob('C:/Users/David/Desktop/EOD/Today/Spread/*.xlsx')
files_ms = glob.glob('C:/Users/David/Desktop/EOD/Today/Spread/MS/*.csv')
latest_file = max(files, key=os.path.getmtime)
latest_file_ms = max(files_ms, key=os.path.getmtime)
today = dt.datetime.today().strftime('%Y%m%d')


# MS Files
ms_df = pd.read_csv(latest_file_ms)
ms_df = ms_df[ms_df['Quantity'] < 0] # Short spreads
ms_df = ms_df[['Trade Date', 'Stock', 'Description', 'Spread']]
ms_df_today = ms_df[ms_df['Trade Date'] == '3/20/2026']
ms_df_today['Stock'] = ms_df_today['Stock'].apply(lambda x: x[:6] + " KS Equity")
ms_df_today = ms_df_today.set_index('Stock')

# Remove duplicates
if ms_df_today.index.duplicated().any():
    ms_df_today = ms_df_today[~ms_df_today.index.duplicated(keep='first')]
    print("Duplicates removed!")


# Enfusion Files
enf_spread = pd.json_normalize(data['rows'])
enf_spread = enf_spread[~enf_spread['financingCurrentResetSpread.value'].isin([0, 55])]
#enf_spread['positionBBYellow.value'] = enf_spread['positionBBYellow.value'].apply(lambda x: x[:6])
enf_spread = enf_spread.drop_duplicates(subset=['positionBBYellow.value'], keep='first')
enf_spread.set_index('positionBBYellow.value', inplace=True)


# Merge
comparison = pd.DataFrame({'MS': ms_df_today['Spread'], "Enf": enf_spread['financingCurrentResetSpread.value']})

# Split into same vs different
same = comparison[comparison['MS'] == comparison['Enf']]
diff = comparison[comparison['MS'] != comparison['Enf']]

print(f"Matching rows: {len(same)}")
print(f"Different rows: {len(diff)}")
print("\n=== Differences ===")
print(diff)

exit()
# Create Quote
endpoint = '/api/marketdata/quotes/import'

json_msg = [{
            "instrumentId": 90783745,(TRS Spread override id)
            "assetMeasure": "MarketPrice",
            "date": "2024-11-15",
            "quoteSet": "Enfusion- Default",
            "quoteSource": "Internal",
            "bid": "250",
            "ask": "250",
            "last": "250"
    }
]
parameters = None
parameters = {
    "ignoreInvalidRequests": False,
    "allowDuplicateQuotes": True
}
