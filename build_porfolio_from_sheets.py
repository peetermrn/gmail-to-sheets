import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import gspread

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive']


class Stock:
    def __init__(self, ticker):
        self.ticker = ticker
        self.ind_stocks = []

    def add_order_info(self, price, amount):
        if amount > 0:
            [self.ind_stocks.append(price) for _ in range(amount)]
        else:
            self.ind_stocks = self.ind_stocks[abs(amount):]

    def get_average_price(self):
        return round(sum(self.ind_stocks) / len(self.ind_stocks), 3)

    def is_empty(self):
        return len(self.ind_stocks) == 0

    def __repr__(self):
        return f"{self.ticker}, {len(self.ind_stocks)}, {self.get_average_price()}"


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
    return client


def build_portfolio(client):
    """Take all stock orders and build current portfolio from them."""
    sheet = client.open('aktsiad').sheet1
    rows = sheet.get_all_records()  # this breaks if some cells are empty strings
    # explanation on why breaks: https://github.com/burnash/gspread/issues/1007
    port_dict = {}
    portfolio = []
    nr = 0
    for row in rows:
        nr += 1
        if row['aktsia'] in port_dict:
            stock = port_dict[row['aktsia']]
        else:
            stock = Stock(row['aktsia'])
            port_dict[row['aktsia']] = stock
        stock.add_order_info(float(row['hind']), int(row['kogus']))
    [portfolio.append(s) for s in port_dict.values() if not s.is_empty()]
    return portfolio


if __name__ == '__main__':
    client_ = authenticate()
    portfolio_ = build_portfolio(client_)
    print(portfolio_)
