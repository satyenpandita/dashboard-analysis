import xlsxwriter
import datetime


class PortfolioEPSExporter(object):

    def __init__(self, longs, shorts):
        self.workbook = None
        self.longs = longs
        self.shorts = shorts

    def export(self, analyst):
        if analyst == "AA":
            filename = "aa override eps ideas {}.xlsx".format(analyst)
        else:
            filename = "eps aim best ideas {}.xlsx".format(analyst)
        output_path = "/var/www/output/{}".format(filename)
        self.workbook = xlsxwriter.Workbook(output_path)
        try:
            self.__write_headers('Sheet1')
            self.__write_data('Sheet1')
        except Exception as e:
            print(str(e))
        finally:
            self.workbook.close()
            return output_path, filename

    def __write_headers(self, sheet):
        worksheet = self.workbook.add_worksheet(sheet)
        merge_format = self.workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        worksheet.write('A1', 'Stock Code', merge_format)
        worksheet.write('B1', 'Rel Period', merge_format)
        worksheet.write('C1', 'EPS Base', merge_format)
        worksheet.write('D1', 'EPS Bear', merge_format)
        worksheet.write('E1', 'Date', merge_format)

    def __write_data(self, sheet):
        portfolio = {**self.longs, **self.shorts}
        worksheet = self.workbook.get_worksheet_by_name(sheet)
        percentage_format = self.workbook.add_format()
        percentage_format.set_num_format(0x0a)
        count = 2

        for stock, data in portfolio.items():
            worksheet.write('A{}'.format(count), stock)
            worksheet.write('B{}'.format(count), "2FY")
            worksheet.write('C{}'.format(count), data['base_eps_1yr'])
            worksheet.write('D{}'.format(count), data['bear_eps_1yr'])
            now = datetime.datetime.now()
            worksheet.write('E{}'.format(count), now.strftime('%m/%d/%y'))
            count += 1


