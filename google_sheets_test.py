#!/usr/bin/env python3


"""
File:      google_sheets_test.py
Author:    Joshua Tice
Date:      2018-12-01

This script is a template for uploading a pandas Dataframe to a Google
Spreadsheet

Usage Notes
-----------
See the following resources for Google Sheet and project folder setup:
- https://gspread.readthedocs.io/en/latest/api.html
- https://medium.com/@monipip3/take-your-job-to-the-next-level-with-python-
  google-sheets-d18a39b815ab

Requirements
------------
gspread
oauth2client
pandas
"""


import argparse
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd


SCOPE = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive',
]
CREDENTIAL_FILE = 'client_secret.json'
GOOGLE_WORKBOOK = 'python_test'


def fetch_args():
    """Fetch arguments from the command line"""
    parser = argparse.ArgumentParser()
    parser.add_argument('-v', '--verbose',
                        help="Prints status when running script", action='store_true')
    args = parser.parse_args()

    return args


def create_client(scope=SCOPE, credential_file=CREDENTIAL_FILE):
    """Instantiates a client for interacting with Google Sheets

    Parameters
    ----------
    scope : str or list of str
        The scope that the client is allowed access to
    credentials : str
        Path to json file with stored credentials
    """

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        credential_file, scope)
    client = gspread.authorize(creds)

    return client


def upload_df_to_google(df, google_workbook=GOOGLE_WORKBOOK, verbose=False):
    """Takes a pandas dataframe and appends the data to a Google sheet

    Parameters
    ----------
    df : pandas DataFrame
        Data frame to be uploaded to Google spreadsheet
    google_sheet : str
        The name of the Google spreadsheet to upload data to
    """

    client = create_client()
    sheet = client.open(google_workbook).sheet1
    if verbose:
        print("Client authenticated...")

    # Make sure the google sheet can accommodate the dataframe by adding
    # extra rows
    sheet.add_rows(len(df))
    if verbose:
        print("Adding rows...")

    for i, row in df.iterrows():
        sheet.append_row(list(row))
    if verbose:
        print("Rows successfully added!...")


def view_sheet(google_workbook=GOOGLE_WORKBOOK):
    """Print the current values in the Google worksheet

    Parameters
    ----------
    google_workbook : str
        The name of the Google workbook to print (Sheet 1)
    """

    client = create_client()
    sheet = client.open(google_workbook).sheet1

    print("Current Google sheet values:")
    rows = sheet.get_all_values()
    for row in rows:
        print(row)


if __name__ == '__main__':
    df = pd.DataFrame({
        'a': [1, 1, 1],
        'b': [2, 2, 2],
        'c': [3, 3, 3],
    })

    args = fetch_args()
    upload_df_to_google(df, verbose=args.verbose)
    view_sheet()
    if args.verbose:
        print('DONE!')
