import json
import requests
from pathlib import Path

# Set base directory using pathlib (OS-compatible path handling)
BASE_DIR = Path(__file__).resolve().parent

# from config import PRELUDE_LS

def change_spread(username, password):
    session = requests.Session()
    # host = 'webservices.enfusionsystems.com'  # this means pointing to PROD
    host = 'webservices-qa.enfusionsystems.com'  #this is QA
    endpoint = '/auth/authentication/generateSecureAPIToken'
    rest_url = 'https://' + host + endpoint
    res = session.get(rest_url, auth=(username, password))
    res.raise_for_status()
    token = res.text

    # put token in request header
    session.headers["Authorization"] = "Bearer " + token

    # Get Instrument Id
    endpoint_report = '/api/marketdata/quotes/import'

    json_msg = [{
        "instrumentId": 90783745, # (TRS Spread override id)
            "assetMeasure": "MarketPrice",
    "date": "2024-11-15",
    "quoteSet": "Enfusion- Default",
    "quoteSource": "Internal",
    "bid": "250",
    "ask": "250",
    "last": "250"
    }
    ]
    params = {
        "ignoreInvalidRequests": False,
        "allowDuplicateQuotes": True
    }
    response = session.get('https://' + host + endpoint_report, params=params)
    if response.ok:
        print("Success")
    else:
        print("failed")

    return



def fetch_deal_id(username, password):
    session = requests.Session()
    host = 'webservices.enfusionsystems.com'  # this means pointing to PROD
    # host = 'webservices-qa.enfusionsystems.com'  #this is QA
    endpoint = '/auth/authentication/generateSecureAPIToken'
    rest_url = 'https://' + host + endpoint
    res = session.get(rest_url, auth=(username, password))
    res.raise_for_status()
    token = res.text

    # put token in request header
    session.headers["Authorization"] = "Bearer " + token
    endpoint_report = '/api/reports'


    params = {
        "report": "shared/03 Portfolio Management/04 Deal Id Positions.ppr",
        "includeMetaData": "false",
        "includeTotals": "false"
    }
    response_report = session.get('https://' + host + endpoint_report, params=params).json()
    with open(BASE_DIR / "basket_enfusion.json", "w") as f:
        json.dump(response_report, f, indent=2)
    print("----------------Fetched EOD Trade Report----------------")
    return


def fetch_eod_report(username, password):
    session = requests.Session()
    host = 'webservices.enfusionsystems.com'  # this means pointing to PROD
    # host = 'webservices-qa.enfusionsystems.com'  #this is QA
    endpoint = '/auth/authentication/generateSecureAPIToken'
    rest_url = 'https://' + host + endpoint
    res = session.get(rest_url, auth=(username, password))
    res.raise_for_status()
    token = res.text

    # put token in request header
    session.headers["Authorization"] = "Bearer " + token

    # Get Instrument Id
    endpoint_instrument = '/api/instrument/search/fast'
    endpoint_report = '/api/reports'

    params = {
        "report": "shared/01 Operation/Daily Task/B) EOD Trade Blotter/EOD Trade Blotter - Bulk.trb",
        "includeMetaData": "false",
        "includeTotals": "false"
    }

    # response = session.get('https://' + host + endpoint_instrument, headers={"Authorization": "Bearer " + token}, params=params)
    # print(response.json())

    response_report = session.get('https://' + host + endpoint_report, params=params).json()

    with open(BASE_DIR / "EOD_Trade.json", "w") as f:
        json.dump(response_report, f, indent=2)
    print("----------------Fetched Deal Id Report----------------")
    return


# Only run when this module is executed directly, not when it is imported.
# (Importing this module elsewhere must not trigger a network call.)
if __name__ == "__main__":
    fetch_eod_report('ops@dunamiscap.com', 'Kingduna10!')






