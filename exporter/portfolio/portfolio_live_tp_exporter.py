import xlsxwriter
import datetime


class PortfolioLiveTPExporter(object):

    def __init__(self, longs, shorts):
        self.workbook = None
        self.longs = longs
        self.shorts = shorts

    def export(self, analyst):
        filename = "portfolio live tps {}.xlsx".format(analyst)
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
        worksheet.write('B1', 'Date', merge_format)
        worksheet.write('C1', 'Base Tp 1yr', merge_format)
        worksheet.write('D1', 'Bear Tp 1yr', merge_format)
        worksheet.write('E1', 'Base Con 1yr', merge_format)
        worksheet.write('F1', 'Bear Con 1yr', merge_format)
        worksheet.write('G1', 'Base Tp 3yr', merge_format)
        worksheet.write('H1', 'Bear Tp 3yr', merge_format)
        worksheet.write('I1', 'Base Con 3yr', merge_format)
        worksheet.write('J1', 'Bear Con 3yr', merge_format)
        worksheet.write('K1', 'Base Multiple 1yr', merge_format)
        worksheet.write('L1', 'Bear Multiple 1yr', merge_format)
        worksheet.write('M1', 'Valuation Str', merge_format)

    def __write_data(self, sheet):
        portfolio = {**self.longs, **self.shorts}
        worksheet = self.workbook.get_worksheet_by_name(sheet)
        percentage_format = self.workbook.add_format()
        percentage_format.set_num_format(0x0a)
        count = 2

        for stock, data in portfolio.items():
            worksheet.write('A{}'.format(count), stock)
            worksheet.write('B{}'.format(count), datetime.datetime.now().strftime('%m/%d/%y'))
            worksheet.write('C{}'.format(count), data['base_tp_1yr'])
            worksheet.write('D{}'.format(count), data['bear_tp_1yr'])
            worksheet.write('E{}'.format(count), data['base_con_1yr'])
            worksheet.write('F{}'.format(count), data['bear_con_1yr'])
            worksheet.write('G{}'.format(count), data['base_tp_3yr'])
            worksheet.write('H{}'.format(count), data['bear_tp_3yr'])
            worksheet.write('I{}'.format(count), data['base_con_3yr'])
            worksheet.write('J{}'.format(count), data['bear_con_3yr'])
            worksheet.write('K{}'.format(count), data['base_multiple_1yr'])
            worksheet.write('L{}'.format(count), data['bear_multiple_1yr'])
            worksheet.write('M{}'.format(count), data['valuation_str'])
            count += 1

