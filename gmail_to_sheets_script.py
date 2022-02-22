from __future__ import print_function

import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from base64 import urlsafe_b64decode
import re
import gspread

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive']


def authenticate():
    """Authentication. This is from gmail API quickstart.py."""
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials_desktop.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    client = gspread.authorize(creds)  # initialize gspread client
    return creds, client


def get_messages(creds, client):
    """
    Get messages from gmail. In my case I want messages with the words 'order on täidetud allpool toodud tingimustel'
    (the order has been filled with details below).
    """
    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        after = get_last_entry_date(client)
        q = f'order on täidetud allpool toodud tingimustel after:{after}'  # take last entry date from sheets. Only
        # take emails which were received after that
        result = service.users().messages().list(userId='me', q=q).execute()
        base_messages = []
        if 'messages' in result:
            base_messages.extend(result['messages'])
        while 'nextPageToken' in result:
            page_token = result['nextPageToken']
            result = service.users().messages().list(userId='me', q=q, pageToken=page_token).execute()
            if 'messages' in result:
                messages.extend(result['messages'])
        result = []
        for msg in base_messages:
            info = service.users().messages().get(userId='me', id=msg['id'], format='full').execute()
            payload = info['payload']
            body = payload['body']
            data = body.get("data")
            if data:
                text = urlsafe_b64decode(data).decode()  # apparently the body of the email has to be decoded first
                date = re.findall("Orderi kuupäev: {3,} \\d{4}-\\d{2}-\\d{2}", text)[0][
                       -10:]  # regex to find date of purchase
                result.append([info['payload']['headers'][-4]['value'], text, date])
        return result

    except HttpError as error:
        print(f'An error occurred: {error}')


def process_messages(raw_messages):
    """Return list where elements are lists as: [date, ticker, price, amount]."""
    sell = "müügiorder"  # this string is included if the order is a sell order.
    final_list = []
    for msg in raw_messages:
        order_info, text, date = msg
        temp = date.split("-")
        date = f"{temp[-1]}.{temp[1]}.{temp[0]}"
        coef = 1
        if sell in order_info:  # id the order is a sell order then amount has to be a negative number
            coef = -1
        amount = int(float(re.findall('Kogus:  {3,}\\d+[.]\\d+', text)[0].split()[1])) * coef  # regex to find amount
        price = float(re.findall("Hind:  {3,}\\d+[.]\\d+", text)[0].split()[1])  # regex ro find price
        ticker = re.findall("bol: {3,} [A-Z0-9]+", text)[0].split()[1]  # regex to find ticker
        final_list.append([
            date, ticker, price, amount
        ])
    return reversed(final_list)


def get_last_entry_date(client):  # check when the last stock order was.
    sheet = get_sheet('all', client)
    row_nr = len(sheet.col_values(1))
    if row_nr <= 1:
        return "2019/01/01"
    date = sheet.row_values(row_nr)[0]
    temp = date.split(".")
    return f"{temp[-1]}/{temp[1]}/{int(temp[0]) + 1}"  # return date one day after last stock order. Otherwise some
    # would be repeated.


def write_to_sheets(lines, client):
    sheet = get_sheet('all', client)
    prev_line_year = ""
    year_sheet = get_sheet("2022", client)
    for line in lines:
        row_nr = len(sheet.col_values(1)) + 1
        sheet.insert_row(line, row_nr)
        year = line[0].split(".")[-1]
        if year != prev_line_year:
            year_sheet = get_sheet(year, client)
            prev_line_year = year
        row_nr_year = len(year_sheet.col_values(1)) + 1
        year_sheet.insert_row(line, row_nr_year)


def get_sheet(year, client):
    # this could be done with a dictionary in a much nicer way but for some reason it didn't work
    # when I tried it like client.open(sheets_dict['2019']).sheet1
    if year == "all":
        return client.open('aktsiad').sheet1
    if year == "2019":
        return client.open('aktsiad2019').sheet1
    if year == "2020":
        return client.open('aktsiad2020').sheet1
    if year == "2021":
        return client.open('aktsiad2021').sheet1
    if year == "2022":
        return client.open('aktsiad2022').sheet1
    if year == "2023":
        return client.open('aktsiad2023').sheet1


if __name__ == '__main__':
    creds_, client_ = authenticate()  # all authentication stuff
    messages = get_messages(creds_, client_)  # get initial messages
    processed_messages = process_messages(
        messages)  # get messages distilled down to [[date1, ticker1, price1, amount1], ..]
    write_to_sheets(processed_messages, client_)  # add new info to sheets
