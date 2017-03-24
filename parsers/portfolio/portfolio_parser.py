from utils.cell_functions import find_cell, cell_value
from exporter.portfolio.portfolio_exporter import PortfolioExporter
from utils.ftp_upload import ftp_upload, get_users
from utils.email_sender import send_mail


def get_tickers(worksheet, direction):
    folio = dict()
    cell_address = find_cell(worksheet, "Ticker", row_offset=0)
    if cell_address:
        row, col = cell_address
        next_row = row
        short = False
        while True:
            next_row += 1
            val = cell_value(worksheet, next_row, col)
            if 'total' in val.lower():
                if direction == 'long':
                    break
                elif direction == 'short':
                    if not short:
                        short = True
                        continue
                    else:
                        break
            if direction == 'long' or (direction == 'short' and short):
                folio[cell_value(worksheet, next_row, col)] = cell_value(worksheet, next_row, col+1), \
                                                              cell_value(worksheet, next_row, col+2)
    return folio


def get_analyst(worksheet):
    name = cell_value(worksheet, 6, 2)
    user_data = get_users()
    if name:
        for key, val in user_data.items():
            if name.lower() in val:
                return key
    return ""


class PortfolioParser(object):

    def __init__(self, worksheet):
        self.long_tickers = get_tickers(worksheet, "long")
        self.short_tickers = get_tickers(worksheet, "short")
        self.analyst = get_analyst(worksheet)
        self.output_file_long = ''
        self.output_file_long_name = ''
        self.output_file_short = ''
        self.output_file_short_name = ''

    def generate_upload_file(self, filename):
        exporter = PortfolioExporter(self.long_tickers, self.short_tickers)
        self.output_file_long, self.output_file_long_name = exporter.export(self.analyst, 'long')
        self.output_file_short, self.output_file_short_name = exporter.export(self.analyst, 'short')

    def send_email(self):
        send_mail("ppal@auroim.com", ["ppal@auroim.com", "bmukherjee@auroim.com"], "Best Ideas Published",
                  "Please file the file attached published by {}".format(self.analyst),
                  [self.output_file_long, self.output_file_short], username="ppal@auroim.com", password="AuroOct2016")

    def ftp_upload(self):
        output_path = self.output_file_long if self.output_file_long else r'uploaded_files/output/portfolio.xlsx'
        ftp_upload(output_path, self.output_file_long_name)
        output_path = self.output_file_short if self.output_file_short else r'uploaded_files/output/portfolio.xlsx'
        ftp_upload(output_path, self.output_file_short_name)
