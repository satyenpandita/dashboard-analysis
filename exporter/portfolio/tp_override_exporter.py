import xlsxwriter
import datetime


class TPOverrideExporter(object):

    def __init__(self, longs, shorts):
        self.workbook = None
        self.longs = longs
        self.shorts = shorts

    def export(self, analyst):
        if analyst == "AA":
            filename = "live tp override {}".format(analyst)
        else:
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

    def __write_data(self, sheet):
        portfolio = {**self.longs, **self.shorts}
        worksheet = self.workbook.get_worksheet_by_name(sheet)
        percentage_format = self.workbook.add_format()
        percentage_format.set_num_format(0x0a)
        count = 2

        for stock, data in portfolio.items():
            if data['is_override']:
                worksheet.write('A{}'.format(count), stock)
                worksheet.write('B{}'.format(count), datetime.datetime.now().strftime('%m/%d/%y'))
                worksheet.write('C{}'.format(count), data['base_tp_1yr'])
                worksheet.write('D{}'.format(count), data['bear_tp_1yr'])
                worksheet.write('E{}'.format(count), data['base_con_1yr'])
                worksheet.write('F{}'.format(count), data['bear_con_1yr'])
                count += 1

