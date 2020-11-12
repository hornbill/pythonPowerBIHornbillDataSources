# [Power BI](https://powerbi.microsoft.com/) [Python 3](https://www.python.org/) Data Sources for Hornbill Reporting And Trend Engine

## Overview

These example scripts have been provided to enable Power BI administrators to build reports and dashboards using Hornbill Reporting and Advanced Analytics Trend Engine data as their data source(s).

## Dependencies

The scripts have been written in [Python 3](https://www.python.org/), and were developed using the following:

- [Power BI Desktop build 2.75.5649.861 64-bit (November 2019)](https://powerbi.microsoft.com/)
- [Python 3.8](https://docs.python.org/3/whatsnew/3.8.html)

The following packages are required dependencies, and can be installed using the Python Package Installer (pip):

- [requests](https://pypi.org/project/requests/)
- [pandas](https://pypi.org/project/pandas/)
- [xlrd](https://pypi.org/project/xlrd/)

## Configuration used in all scripts

Each script requires the following variables to be set (all case-sensitive):

- `apiKey` - This is an API key generated against a user account on the Hornbill Administration Console, where the user account has sufficient access to run reports and access trending data.
- `instanceId` - This is the (case sensitive) name of the instance to connect to

## Scripts

### PowerBIReport.py

This script will:

- Run a pre-defined report on the Hornbill instance;
- Wait for the report to complete;
- Retrieve the report CSV data and present back as a dataframe called df, which can then be consumed by PowerBI.

Script Variables:

- `reportId`: The ID (Primary Key, INT) of the  report to be run;
- `suspendSeconds`: The number of seconds the script should wait between checks to see if the report is complete;
- `deleteReportInstance`: a boolean value to determine if, once the report is run on Hornbill and the data has been pulled in to PowerBI, whether the historic report run instance should be removed from your Hornbill report;
- `useXLSX`: False = the script will use the CSV output from your report; True = the script will use the XLSX output from your report. NOTE: XLSX output will need to be enabled within the Output Formats > Additional Data Formats section of your report in Hornbill.

### PowerBIHistoricReport.py

This script will:

- Retrieve a historic report CSV from your Hornbill instance;
- Present the report data back as a dataframe called df, which can then be consumed by PowerBI.

Script Variables:

- `reportId`: The ID (Primary Key, INT) of the  report to be run;
- `reportRunId`: The Run ID (INT) of a historic run of the above report ID;
- `useXLSX`: False = the script will use the CSV output from your report; True = the script will use the XLSX output from your report. NOTE: XLSX output will need to be enabled within the Output Formats > Additional Data Formats section of your report in Hornbill.

### PowerBITrendingData.py

This script will:

- Run the reporting::measureGetInfo API  against your Hornbill instance, with a given measure ID (Primary Key, INT);
- Build a table containing all Trend Value entries for the selected measure;
- Present the trend data back as a dataframe called df, which can then be consumed by PowerBI.

Script Variable:

- `measureId`: The ID (Primary Key, INT) of the measure to return trend data from.

Outputs:
As the response parameters from the Trending Engine is fixed (unlike the Reporting engine, which has user-specified column outputs), the output for this report will always consist  of the following columns:

- `value`: the value of the trend sample;
- `sampleId`: the ID of the sample;
- `sampleTime`: the time & date that the sample was taken;
- `dateRange.from`: the start date of the sample snapshot;
- `dateRange.to`: the end date of the sample snapshot;

## Power BI Notes

Please see the [Power BI Documentation](https://docs.microsoft.com/en-us/power-bi/desktop-python-scripts) for more information about using Python with Power BI.

These scripts have been designed to be used as data sources only, and not as the source of Python visuals within Power BI. Which is not to say they couldn't be used in your Python visuals, with a little extra code and the [matplotlib](https://pypi.org/project/matplotlib/) library!
