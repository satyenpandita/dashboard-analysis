import re
from models.portfolio import Portfolio, PortfolioItem
from exporter.portfolio.portfolio_exporter import PortfolioExporter
from exporter.portfolio.portfolio_eps_exporter import PortfolioEPSExporter
from exporter.portfolio.portfolio_live_tp_exporter import PortfolioLiveTPExporter
from exporter.portfolio.tp_override_exporter import TPOverrideExporter
from utils.cell_functions import find_cell, cell_value
from utils.email_sender import send_mail
from utils.upload_ops import ftp_upload, get_users
from utils.file_helper import portfolio_upload_util

INVALID_TICKERS = dict()


def check_is_new(worksheet):
    val = cell_value(worksheet, 0, 0)
    if val is not None and val != "" and val == 'new':
        return True
    else:
        return False


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
                ticker = val
                if ticker:
                    folio[ticker] = dict(weight=cell_value(worksheet, target_row, target_col + 5),
                                         roc=cell_value(worksheet, target_row, target_col + 7),
                                         base_tp_1yr=cell_value(worksheet, target_row, target_col + 13) if cell_value(worksheet, target_row, target_col + 13) != "" else None,
                                         bear_tp_1yr=cell_value(worksheet, target_row, target_col + 14) if cell_value(worksheet, target_row, target_col + 14) != "" else None,
                                         base_con_1yr=cell_value(worksheet, target_row, target_col + 15)*100 if cell_value(worksheet, target_row, target_col + 15) != "" else None,
                                         bear_con_1yr=cell_value(worksheet, target_row, target_col + 16)*100 if cell_value(worksheet, target_row, target_col + 16) != "" else None,
                                         base_tp_3yr=cell_value(worksheet, target_row, target_col + 27) if cell_value(worksheet, target_row, target_col + 27) != "" else None,
                                         bear_tp_3yr=cell_value(worksheet, target_row, target_col + 28) if cell_value(worksheet, target_row, target_col + 28) != "" else None,
                                         base_con_3yr=(worksheet, target_row, target_col + 29)*100 if cell_value(worksheet, target_row, target_col + 29) != "" else None,
                                         bear_con_3yr=(worksheet, target_row, target_col + 30)*100 if cell_value(worksheet, target_row, target_col + 30) != "" else None,
                                         base_eps_1yr=cell_value(worksheet, target_row, target_col + 57) if cell_value(worksheet, target_row, target_col + 57) != "" else None,
                                         base_multiple_1yr=cell_value(worksheet, target_row, target_col + 58) if cell_value(worksheet, target_row, target_col + 58) != "" else None,
                                         bear_eps_1yr=cell_value(worksheet, target_row, target_col + 59) if cell_value(worksheet, target_row, target_col + 59) != "" else None,
                                         bear_multiple_1yr=cell_value(worksheet, target_row, target_col + 60) if cell_value(worksheet, target_row, target_col + 60) != "" else None,
                                         valuation_str=""
                                         )
                else:
                    INVALID_TICKERS[target_row + 1] = val
            target_row += 1
    return folio


def is_aa_portfolio(worksheet):
    analyst = get_analyst(worksheet)
    if analyst == 'AA':
        return True
    return False


def get_tickers_new(worksheet, direction):
    aa_portfolio = is_aa_portfolio(worksheet)
    folio = dict()
    folio_type = 'LONG Pfolio ($M)' if direction == 'long' else 'Short Pfolio ($M)'
    row_offset = 3 if direction == 'long' else 2
    cell_address = find_cell(worksheet, folio_type)
    if cell_address:
        row, col = cell_address
        target_row, target_col = row + row_offset, col + 1
        while True:
            ticker = cell_value(worksheet, target_row, target_col)
            weight = cell_value(worksheet, target_row, target_col + 5)
            if ticker == "END":
                break
            elif ticker and ticker != "" and weight != "" and weight >= 0 and not aa_portfolio:
                folio[ticker] = dict(weight=weight,
                                     name=cell_value(worksheet, target_row, target_col - 1),
                                     roc=cell_value(worksheet, target_row, target_col + 7),
                                     base_tp_1yr=cell_value(worksheet, target_row, target_col + 13) if cell_value(worksheet, target_row, target_col + 13) != "" else None,
                                     bear_tp_1yr=cell_value(worksheet, target_row, target_col + 14) if cell_value(worksheet, target_row, target_col + 14) != "" else None,
                                     base_con_1yr=cell_value(worksheet, target_row, target_col + 15)*100 if cell_value(worksheet, target_row, target_col + 15) != "" else None,
                                     bear_con_1yr=cell_value(worksheet, target_row, target_col + 16)*100 if cell_value(worksheet, target_row, target_col + 16) != "" else None,
                                     valuation_str=cell_value(worksheet, target_row, target_col + 20),
                                     base_tp_3yr=cell_value(worksheet, target_row, target_col + 28) if cell_value(worksheet, target_row, target_col + 28) != "" else None,
                                     bear_tp_3yr=cell_value(worksheet, target_row, target_col + 29) if cell_value(worksheet, target_row, target_col + 29) != "" else None,
                                     base_con_3yr=cell_value(worksheet, target_row, target_col + 30)*100 if cell_value(worksheet, target_row, target_col + 30) != "" else None,
                                     bear_con_3yr=cell_value(worksheet, target_row, target_col + 31)*100 if cell_value(worksheet, target_row, target_col + 31) != "" else None,
                                     base_eps_1yr=cell_value(worksheet, target_row, target_col + 58) if cell_value(worksheet, target_row, target_col + 58) != "" else None,
                                     base_multiple_1yr=cell_value(worksheet, target_row, target_col + 59) if cell_value(worksheet, target_row, target_col + 59) != "" else None,
                                     bear_eps_1yr=cell_value(worksheet, target_row, target_col + 60) if cell_value(worksheet, target_row, target_col + 60) != "" else None,
                                     bear_multiple_1yr=cell_value(worksheet, target_row, target_col + 61) if cell_value(worksheet, target_row, target_col + 61) != "" else None,
                                     is_live=False,
                                     is_overrride=False if not aa_portfolio else (cell_value(worksheet, target_row, target_col + 62) == "YES" if cell_value(worksheet, target_row, target_col + 62) != "" else False)
                                     )
            elif ticker and ticker != "" and weight == "":
                folio[ticker] = dict(weight=None,
                                     name=cell_value(worksheet, target_row, target_col - 1),
                                     roc=cell_value(worksheet, target_row, target_col + 7),
                                     base_tp_1yr=cell_value(worksheet, target_row, target_col + 13) if cell_value(worksheet, target_row, target_col + 13) != "" else None,
                                     bear_tp_1yr=cell_value(worksheet, target_row, target_col + 14) if cell_value(worksheet, target_row, target_col + 14) != "" else None,
                                     base_con_1yr=cell_value(worksheet, target_row, target_col + 15)*100 if cell_value(worksheet, target_row, target_col + 15) != "" else None,
                                     bear_con_1yr=cell_value(worksheet, target_row, target_col + 16)*100 if cell_value(worksheet, target_row, target_col + 16) != "" else None,
                                     valuation_str=cell_value(worksheet, target_row, target_col + 20),
                                     base_tp_3yr=cell_value(worksheet, target_row, target_col + 28) if cell_value(worksheet, target_row, target_col + 28) != "" else None,
                                     bear_tp_3yr=cell_value(worksheet, target_row, target_col + 29) if cell_value(worksheet, target_row, target_col + 29) != "" else None,
                                     base_con_3yr=cell_value(worksheet, target_row, target_col + 30)*100 if cell_value(worksheet, target_row, target_col + 30) != "" else None,
                                     bear_con_3yr=cell_value(worksheet, target_row, target_col + 31)*100 if cell_value(worksheet, target_row, target_col + 31) != "" else None,
                                     base_eps_1yr=cell_value(worksheet, target_row, target_col + 58) if cell_value(worksheet, target_row, target_col + 58) != "" else None,
                                     base_multiple_1yr=cell_value(worksheet, target_row, target_col + 59) if cell_value(worksheet, target_row, target_col + 59) != "" else None,
                                     bear_eps_1yr=cell_value(worksheet, target_row, target_col + 60) if cell_value(worksheet, target_row, target_col + 60) != "" else None,
                                     bear_multiple_1yr=cell_value(worksheet, target_row, target_col + 61) if cell_value(worksheet, target_row, target_col + 61) != "" else None,
                                     is_live=True,
                                     is_overrride=False if not aa_portfolio else (cell_value(worksheet, target_row, target_col + 62) == "YES" if cell_value(worksheet, target_row, target_col + 62) != "" else False)
                                     )
            else:
                INVALID_TICKERS[target_row + 1] = ticker
            target_row += 1
    return folio


def valid_ticker(ticker):
    sr = re.search(r"\s[A-Za-z]{2}\s(equity)$", ticker.lower())
    if sr is None:
        sr = re.search(r"^[\w{1}]+\s(equity)$", ticker.lower())
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
        self.is_new = check_is_new(worksheet)
        self.long_tickers = get_tickers_new(worksheet, "long") if self.is_new else get_tickers(worksheet, "long")
        self.short_tickers = get_tickers_new(worksheet, "short") if self.is_new else get_tickers(worksheet, "short")
        self.invalid_tickers = INVALID_TICKERS
        self.analyst = get_analyst(worksheet)
        self.output_file_long = ''
        self.output_file_long_name = ''
        self.output_file_short = ''
        self.output_file_short_name = ''
        self.input_file = input_file
        self.eps_file = ''
        self.eps_file_name = ''
        self.live_tp_file = ''
        self.live_tp_file_name = ''
        self.live_tp_override_file = ''
        self.live_tp_override_file_name = ''

    def save_and_generate_files(self):
        short_list = []
        long_list = []
        for stock, data in self.short_tickers.items():
            portfolio_item = PortfolioItem(stock_code=stock, weight=data['weight'], reason_for_change=data['roc'],
                                           base_tp_1yr=data['base_tp_1yr'], bear_tp_1yr=data['bear_tp_1yr'],
                                           base_con_1yr=data['base_con_1yr'], bear_con_1yr=data['bear_con_1yr'],
                                           base_tp_3yr=data['base_tp_3yr'], bear_tp_3yr=data['bear_tp_3yr'],
                                           base_con_3yr=data['base_con_3yr'], bear_con_3yr=data['bear_con_3yr'],
                                           base_eps_1yr=data['base_eps_1yr'], bear_eps_1yr=data['bear_eps_1yr'],
                                           base_multiple_1yr=data['base_multiple_1yr'], name=data['name'],
                                           bear_multiple_1yr=data['bear_multiple_1yr'],
                                           valuation_str=data['valuation_str'], is_live=data['is_live'],
                                           is_override=data['is_override'])
            short_list.append(portfolio_item)
        for stock, data in self.long_tickers.items():
            portfolio_item = PortfolioItem(stock_code=stock, weight=data['weight'], reason_for_change=data['roc'],
                                           base_tp_1yr=data['base_tp_1yr'], bear_tp_1yr=data['bear_tp_1yr'],
                                           base_con_1yr=data['base_con_1yr'], bear_con_1yr=data['bear_con_1yr'],
                                           base_tp_3yr=data['base_tp_3yr'], bear_tp_3yr=data['bear_tp_3yr'],
                                           base_con_3yr=data['base_con_3yr'], bear_con_3yr=data['bear_con_3yr'],
                                           base_eps_1yr=data['base_eps_1yr'], bear_eps_1yr=data['bear_eps_1yr'],
                                           base_multiple_1yr=data['base_multiple_1yr'], name=data['name'],
                                           bear_multiple_1yr=data['bear_multiple_1yr'],
                                           valuation_str=data['valuation_str'], is_live=data['is_live'],
                                           is_override=data['is_override'])
            long_list.append(portfolio_item)
        portfolio = Portfolio(analyst=self.analyst, shorts=short_list, longs=long_list,
                              file_path='/var/www/portfolio/{}'.format(self.input_file))
        portfolio.save()
        self.generate_upload_file()

    def generate_upload_file(self):
        exporter = PortfolioExporter(self.long_tickers, self.short_tickers)
        eps_exporter = PortfolioEPSExporter(self.long_tickers, self.short_tickers)
        live_exporter = PortfolioLiveTPExporter(self.long_tickers, self.short_tickers)
        self.output_file_long, self.output_file_long_name = exporter.export(self.analyst, 'long')
        self.output_file_short, self.output_file_short_name = exporter.export(self.analyst, 'short')
        self.eps_file, self.eps_file_name = eps_exporter.export(self.analyst)
        self.live_tp_file, self.live_tp_file_name = live_exporter.export(self.analyst)
        if self.analyst == "AA":
            live_override_exporter = TPOverrideExporter(self.long_tickers, self.short_tickers)
            self.live_tp_override_file, self.live_tp_override_file_name = live_override_exporter.export(self.analyst)
        portfolio_upload_util(self.analyst)

    def send_email(self):
        recipients = ["ppal@auroim.com"]
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
