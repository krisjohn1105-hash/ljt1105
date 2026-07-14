import json
import requests
import pandas as pd
#from Ipython.display import display


username = 'ops@dunamiscap.com'
password = 'Kingduna10!'

session = requests.Session()
host = 'webservices.enfusionsystems.com'  #this means pointing to PROD
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
    "report": "shared/TRS Spread Report.ppr",
    "includeMetaData": "false",
    "includeTotals": "false"
}

#response = session.get('https://' + host + endpoint_instrument, headers={"Authorization": "Bearer " + token}, params=params)
#print(response.json())

response_report = session.get('https://' + host + endpoint_report, params=params).json()
print(response_report)

with open("spread.json", "w") as f:
    json.dump(response_report, f, indent=2)
enf_df = pd.DataFrame(response_report)

print(enf_df)