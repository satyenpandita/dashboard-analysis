from utils.cell_functions import find_cell, cell_value
from exporter.portfolio.portfolio_exporter import PortfolioExporter
from utils.ftp_upload import ftp_upload


def get_tickers(worksheet, direction):
    folio = dict()
    cell_address = find_cell(worksheet, "Ticker", row_offset=8)
    if cell_address:
        row, col = cell_address
        next_row = row
        short = False
        while True:
            next_row += 1
            print(next_row)
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
                folio[cell_value(worksheet, next_row, col)] = cell_value(worksheet, next_row, col+1)
    return folio


class PortfolioParser(object):

    def __init__(self, worksheet):
        self.long_tickers = get_tickers(worksheet, "long")
        self.short_tickers = get_tickers(worksheet, "short")
        self.output_file = ''

    def generate_upload_file(self):
        exporter = PortfolioExporter(self.long_tickers, self.short_tickers)
        self.output_file = exporter.export()

    def ftp_upload(self):
        output_path = self.output_file if self.output_file else r'uploaded_files/output/portfolio.xlsx'
        ftp_upload(output_path)
