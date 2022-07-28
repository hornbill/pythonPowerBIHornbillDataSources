# PowerBIReport.py
# v1.0.0 02/01/2020
# This script will run a report on a Hornbill instance, and retrieve the data in a Power BI consumable dataframe

import pandas as pd
import requests
import sys
import time
import io
import os

apiKey = 'yourapikey'           # This is the (case sensitive) API key to use to authenticate the API calls against the Hornbill instance
instanceId = 'yourinstanceid'   # This is the (case sensitive) ID of your Hornbill instance
reportId = 6                    # This is the ID of the report, which can be taken from the report URL in the admin console
suspendSeconds = 1              # This is the number of seconds to wait inbetween the checking for report run completion 
deleteReportInstance = True     # This defines whether the report run instance is deleted once the report data has been retrieved
useXLSX = False                 # This specifies whether to use the XLSX output from your target Hornbill Report, rather than the default CSV

# Get API endpoint
URL = 'https://files.hornbill.com/instances/{instanceId}/zoneinfo'.format(instanceId=instanceId)
try:
    endpoint = requests.get(url = URL).json()["zoneinfo"]["endpoint"]
except requests.exceptions.RequestException as e:
    sys.exit('Unexpected response when attempting to retrieve Hornbill Zone Information: ' + e)

def invokeXmlmc(service, method, params):
    xmlmcEndpoint = endpoint + "xmlmc/{service}/?method={method}" 
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
    
# Kick off report
reportJson = {
    'reportId': reportId
}
reportDetails = invokeXmlmc('reporting', 'reportRun', reportJson)
if reportDetails['@status'] != True:
    sys.exit(reportDetails['state']['error'])

reportRunJson = {'runId': reportDetails['params']['runId']}

reportComplete = False
reportSuccess = False

while reportComplete == False:
    time.sleep(suspendSeconds)
    reportRunDetails = invokeXmlmc('reporting', 'reportRunGetStatus', reportRunJson)
    if reportRunDetails['@status'] != True:
        sys.exit(reportRunDetails['state']['error'])

    if reportRunDetails['params']['reportRun']['status'] == 'completed':
        reportComplete = True
        reportSuccess = True
    elif (reportRunDetails['params']['reportRun']['status'] == 'failed') or (reportRunDetails['params']['reportRun']['status'] == 'aborted'):
        reportComplete = True

if reportSuccess == False:
    sys.exit('Report run was not successful: ' + reportRunDetails['statusMessage'])

if useXLSX == True:
    xlsURL = endpoint + "dav/reports/{reportId}/{reportLink}"
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
    csvURL = endpoint + "dav/reports/{reportId}/{reportLink}"
    URL = csvURL.format(reportId=reportId,reportLink=reportRunDetails['params']['reportRun']['csvLink'])
    getCSV = requests.get(url=URL, headers={'Authorization':'ESP-APIKEY ' + apiKey})
    df = pd.read_csv(io.StringIO(getCSV.text))

if deleteReportInstance == True:
    deleteReport = invokeXmlmc('reporting', 'reportRunDelete', reportRunJson)

print(df)