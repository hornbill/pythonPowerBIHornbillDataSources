# PowerBIHistoricReport.py
# v1.0.0 02/01/2020
# This script will retrieve the CSV data from a historic report on a Hornbill instance, and output the data in a Power BI consumable dataframe

import pandas as pd
import requests
import sys
import time
import io
import os

apiKey = ''           # This is the (case sensitive) API key to use to authenticate the API calls against the Hornbill instance
instanceId = ''   # This is the (case sensitive) ID of your Hornbill instance
reportId = 1                    # This is the ID of the report, which can be taken from the report URL in the admin console
reportRunId = 21               # This is the ID of the report run instance, which can be taken from the report history in the admin console
useXLSX = False                 # This specifies whether to use the XLSX output from your target Hornbill Report, rather than the default CSV

# Get API endpoint
URL = 'https://files.hornbill.com/instances/{instanceId}/zoneinfo'.format(instanceId=instanceId)
try:
    zoneInfo = requests.get(url = URL).json()["zoneinfo"]
    endpoint = zoneInfo["endpoint"] + "xmlmc/"
    davEndpoint = zoneInfo["endpoint"] + "dav/"
    if "apiEndpoint" in zoneInfo:
        endpoint = zoneInfo["apiEndpoint"]
        davEndpoint = zoneInfo["davEndpoint"]
except requests.exceptions.RequestException as e:
    sys.exit('Unexpected response when attempting to retrieve Hornbill Zone Information: ' + e)

def invokeXmlmc(service, method, params):
    xmlmcEndpoint = endpoint + "{service}/?method={method}" 
    URL = xmlmcEndpoint.format(service=service, method=method)
    headers = {
        'Content-type': 'application/json',
        'Accept': 'application/json',
        'Authorization': 'ESP-APIKEY ' + apiKey
    }
    methodCall={
            '@service':service,
            '@method':method,
            'params':params
    }
    try:
        response = requests.post(url=URL, json=methodCall, headers=headers)
        if 200 >= response.status_code <= 299:
            return response.json()
        else:
            return {
                '@status': False,
                'state': {
                    'error': 'Unexpected status returned from call to {service}::{method}: {statusCode}'.format(service=service, method=method, statusCode=response.status_code)
                }
            } 
    except requests.exceptions.RequestException as e:
        return {
            '@status': False,
            'state': {
                'error': 'Unexpected response from call to {service}::{method}: {errorString}'.format(service=service, method=method, errorString=e)
            }
        }

reportRunDetails = invokeXmlmc('reporting', 'reportRunGetStatus', {'runId': reportRunId})

if useXLSX == True:
    xlsURL = davEndpoint + "reports/{reportId}/{reportLink}"
    for file in reportRunDetails['params']['files']:
        if file['type'] == 'xlsx':
            fileName = file['name']
    URL = xlsURL.format(reportId=reportId,reportLink=fileName)
    getExcel = requests.get(url=URL, headers={'Authorization':'ESP-APIKEY ' + apiKey})
    open(fileName, 'wb').write(getExcel.content)
    df = pd.read_excel(fileName)
    if os.path.exists(fileName):
        os.remove(fileName)
else:
    csvURL = davEndpoint + "reports/{reportId}/{reportLink}"
    URL = csvURL.format(reportId=reportId,reportLink=reportRunDetails['params']['reportRun']['csvLink'])
    getCSV = requests.get(url=URL, headers={'Authorization':'ESP-APIKEY ' + apiKey})
    df = pd.read_csv(io.StringIO(getCSV.text))

print(df)