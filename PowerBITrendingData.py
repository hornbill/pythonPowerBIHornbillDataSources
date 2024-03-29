# PowerBITrendingData.py
# v1.0.0 02/01/2020
# This script will retrieve trending data from a measure on a Hornbill instance, and output the data in a Power BI consumable dataframe

from pandas.io.json import json_normalize
import requests
import sys
import time
import io
import json

apiKey = 'yourapikey'           # This is the (case sensitive) API key to use to authenticate the API calls against the Hornbill instance
instanceId = 'yourinstanceid'   # This is the (case sensitive) ID of your Hornbill instance
measureId = 11                  # This is the ID of the measure, which can be taken from the measure URL in the admin console

# Get API endpoint
URL = 'https://files.hornbill.com/instances/{instanceId}/zoneinfo'.format(instanceId=instanceId)
try:
    zoneInfo = requests.get(url = URL).json()["zoneinfo"]
    endpoint = zoneInfo["endpoint"] 
    if "apiEndpoint" in zoneInfo:
        endpoint = zoneInfo["apiEndpoint"]
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

measureAPIJson = {
    'measureId': measureId,
    'returnMeasureValue': 'true',
    'returnMeasureTrendData': 'true'
}
measureData = invokeXmlmc('reporting', 'measureGetInfo', measureAPIJson)
df = json_normalize(measureData['params']['trendValue'])
print(df)