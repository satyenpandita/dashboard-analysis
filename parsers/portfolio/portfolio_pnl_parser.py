import xlrd
import xlsxwriter
from utils.cell_functions import find_cell, cell_value
from models.portfolio.portfolio_pnl import PortfolioPnl
from models.portfolio.portfolio import Portfolio
from utils.upload_ops import s3_upload_file


ROW_OFFSET = 3


class PortfolioPnlParser(object):

    def __init__(self, workbook):
        self.workbook = workbook

    def parse_save_data(self):
        for sheet in self.workbook.sheets():
            long_pnl = []
            short_pnl = []
            analyst = sheet.name.upper()
            for direction in ['LONG', "SHORT"]:
                key = ".{}_{} INDEX".format(analyst.upper(), direction)
                cell_address = find_cell(sheet, key)
                if cell_address:
                    row, col = cell_address
                    start_row = row + ROW_OFFSET
                    while start_row < sheet.nrows:
                        if cell_value(sheet, start_row, col) not in [None, '']:
                            pnl_date = xlrd.xldate.xldate_as_datetime(sheet.cell(start_row, col).value, self.workbook.datemode)
                            mtd_pct_return = cell_value(sheet, start_row, col+1) if cell_value(sheet, start_row, col+1) != "" else None
                            ytd_pct_return = cell_value(sheet, start_row, col+2) if cell_value(sheet, start_row, col+2) != "" else None
                            market_value = cell_value(sheet, start_row, col+3) if cell_value(sheet, start_row, col+3) != "" else None
                            last_price = cell_value(sheet, start_row, col+4) if cell_value(sheet, start_row, col+4) != "" else None
                            mtd_usd_return = cell_value(sheet, start_row, col+5) if cell_value(sheet, start_row, col+5)!= "" else None
                            ytd_usd_return = cell_value(sheet, start_row, col+6) if cell_value(sheet, start_row, col+6)!= "" else None
                            start_row += 1
                            portfolio_pnl = PortfolioPnl(pnl_date=pnl_date, mtd_pct_return=mtd_pct_return,
                                                         ytd_pct_return=ytd_pct_return, market_value=market_value,
                                                         last_price=last_price, mtd_usd_return=mtd_usd_return,
                                                         ytd_usd_return=ytd_usd_return)
                            if direction == "LONG":
                                long_pnl.append(portfolio_pnl)
                            else:
                                short_pnl.append(portfolio_pnl)
                        else:
                            break
            portfolio = Portfolio.objects.get(analyst=analyst)
            n_pf = Portfolio(analyst=portfolio.analyst, shorts=portfolio.shorts, longs=portfolio.longs,
                             file_path=portfolio.file_path, long_pnl=long_pnl, short_pnl=short_pnl)
            n_pf.save()

    def export(self):
        resp = []
        portfolios = Portfolio.objects.all()
        for portfolio in portfolios:
            file_name = "analyst_pnl_{}.xlsx".format(portfolio.analyst)
            output_path = "/var/www/output/"
            complete_path = output_path+file_name
            try:
                export_book = xlsxwriter.Workbook(complete_path)
                worksheet = export_book.add_worksheet("Sheet 1")
                merge_format = export_book.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'border': 1})
                date_format = export_book.add_format({'num_format': 'mmm-yy'})
                worksheet.write('A1', 'Analyst', merge_format)
                worksheet.write('B1', 'Pnl Date', merge_format)
                worksheet.write('C1', 'MTD Pct Long', merge_format)
                worksheet.write('D1', 'YTD Pct Long', merge_format)
                worksheet.write('E1', 'Market Value Long', merge_format)
                worksheet.write('F1', 'Last Price Long', merge_format)
                worksheet.write('G1', 'MTD USD Long', merge_format)
                worksheet.write('H1', 'YTD USD Long', merge_format)

                worksheet.write('I1', 'MTD Pct Short', merge_format)
                worksheet.write('J1', 'YTD Pct Short', merge_format)
                worksheet.write('K1', 'Market Value Short', merge_format)
                worksheet.write('L1', 'Last Price Short', merge_format)
                worksheet.write('M1', 'MTD USD Short', merge_format)
                worksheet.write('N1', 'YTD USD Short', merge_format)
                count = 2

                ts = portfolio.get_pnl_ts()
                for item in ts:
                    if bool(item):
                        worksheet.write('A{}'.format(count), portfolio.analyst)
                        worksheet.write_datetime('B{}'.format(count), item['pnl_date'], date_format)
                        worksheet.write('C{}'.format(count), item['mtd_pct_return_long'])
                        worksheet.write('D{}'.format(count), item['ytd_pct_return_long'])
                        worksheet.write('E{}'.format(count), item['market_value_long'])
                        worksheet.write('F{}'.format(count), item['last_price_long'])
                        worksheet.write('G{}'.format(count), item['mtd_usd_return_long'])
                        worksheet.write('H{}'.format(count), item['ytd_usd_return_long'])
                        worksheet.write('I{}'.format(count), item['mtd_pct_return_short'])
                        worksheet.write('J{}'.format(count), item['ytd_pct_return_short'])
                        worksheet.write('K{}'.format(count), item['market_value_short'])
                        worksheet.write('L{}'.format(count), item['last_price_short'])
                        worksheet.write('M{}'.format(count), item['mtd_usd_return_short'])
                        worksheet.write('N{}'.format(count), item['ytd_usd_return_short'])
                        count += 1
            except Exception as e:
                print(str(e))
            finally:
                export_book.close()
                s3_upload_file.delay(complete_path, 'pnl', file_name)
                resp.append(file_name)
        return resp
