import re
from models.portfolio import Portfolio, PortfolioItem
from exporter.portfolio.portfolio_exporter import PortfolioExporter
from utils.cell_functions import find_cell, cell_value
from utils.email_sender import send_mail
from utils.upload_ops import ftp_upload, get_users
from utils.upload_ops import get_user_email

INVALID_TICKERS = dict()


def get_tickers(worksheet, direction):
    folio = dict()
    folio_type = 'LONG Pfolio ($M)' if direction == 'long' else 'Short Pfolio ($M)'
    row_offset = 3 if direction == 'long' else 2
    range_length = 11 if direction == 'long' else 13
    cell_address = find_cell(worksheet, folio_type)
    if cell_address:
        row, col = cell_address
        target_row, target_col = row + row_offset, col + 1
        for i in range(range_length):
            val = cell_value(worksheet, target_row, target_col)
            if val and val != "":
                ticker = valid_ticker(val)
                if ticker:
                    folio[ticker] = cell_value(worksheet, target_row, target_col+5), \
                                    cell_value(worksheet, target_row, target_col+7)
                else:
                    INVALID_TICKERS[target_row + 1] = val
            target_row += 1
    return folio


def valid_ticker(ticker):
    sr = re.search(r"\s[A-Za-z]{2}\s(Equity)$", ticker)
    if sr is None:
        sr = re.search(r"^[\w{1}]+\s(Equity)$", ticker)
        if sr is None:
            return None
        else:
            return ticker.strip().replace(' ', ' US ')
    else:
        return ticker


def get_analyst(worksheet):
    name = cell_value(worksheet, 5, 1)
    user_data = get_users()
    if name:
        for key, val in user_data.items():
            if name.strip().lower() in val:
                return key
    return ""


class PortfolioParserV2(object):

    def __init__(self, worksheet, input_file):
        self.long_tickers = get_tickers(worksheet, "long")
        self.short_tickers = get_tickers(worksheet, "short")
        self.invalid_tickers = INVALID_TICKERS
        self.analyst = get_analyst(worksheet)
        self.output_file_long = ''
        self.output_file_long_name = ''
        self.output_file_short = ''
        self.output_file_short_name = ''
        self.input_file = input_file

    def save_and_generate_files(self):
        short_list = []
        long_list = []
        for stock, (weight, roc) in self.short_tickers.items():
            portfolio_item = PortfolioItem(stock_code=stock, weight=weight, reason_for_change=roc)
            short_list.append(portfolio_item)
        for stock, (weight, roc) in self.long_tickers.items():
            portfolio_item = PortfolioItem(stock_code=stock, weight=weight, reason_for_change=roc)
            long_list.append(portfolio_item)
        portfolio = Portfolio(analyst=self.analyst, shorts=short_list, longs=long_list,
                              file_path='/var/www/portfolio/{}'.format(self.input_file))
        portfolio.save()
        self.generate_upload_file()

    def generate_upload_file(self):
        exporter = PortfolioExporter(self.long_tickers, self.short_tickers)
        self.output_file_long, self.output_file_long_name = exporter.export(self.analyst, 'long')
        self.output_file_short, self.output_file_short_name = exporter.export(self.analyst, 'short')

    def send_email(self):
        analyst_email = get_user_email(self.analyst)
        recipients = ["datascience@auroim.com"]
        if analyst_email is not None:
            recipients.append(analyst_email)
        send_mail.delay("ppal@auroim.com",
                        recipients,
                        "Best Ideas Published for {}".format(self.analyst),
                        self.get_body(),
                        [self.output_file_long,
                         self.output_file_short,
                         '/var/www/portfolio/{}'.format(self.input_file)],
                        username="ppal@auroim.com",
                        password="AuroOct2016")
        global INVALID_TICKERS
        INVALID_TICKERS = dict()

    def ftp_upload(self):
        ftp_upload.delay(self.output_file_long, self.output_file_long_name)
        ftp_upload.delay(self.output_file_short, self.output_file_short_name)

    def get_body(self):
        if len(self.invalid_tickers) > 0:
            return """Please find the Best Ideas files published by {} \n\n Some Invalid Tickers Found in file {} \n {}""". \
                format(self.analyst, self.input_file, self.get_errors())
        else:
            return """Please find the Best Ideas files published by {} \n\n No Invalid Securities""". \
                format(self.analyst)

    def get_errors(self):
        err_arr = []
        for key, val in self.invalid_tickers.items():
            err_arr.append("{} : Row Number {}".format(val, key))
        return "\n".join(err_arr)
