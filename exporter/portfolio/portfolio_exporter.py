import xlsxwriter
import datetime


class PortfolioExporter(object):

    def __init__(self, longs, shorts):
        self.workbook = None
        self.longs = longs
        self.shorts = shorts

    def export(self, analyst, direction):
        filename = "aim best ideas {} {}.xlsx".format(analyst, direction)
        output_path = "/var/www/output/{}".format(filename)
        self.workbook = xlsxwriter.Workbook(output_path)
        try:
            self.__write_headers('Sheet1')
            self.__write_data('Sheet1', direction, analyst)
        except Exception as e:
            print(str(e))
        finally:
            self.workbook.close()
            return output_path, filename

    def __write_headers(self, sheet):
        worksheet = self.workbook.add_worksheet(sheet)
        merge_format = self.workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        worksheet.write('A1', 'Portfolio Name', merge_format)
        worksheet.write('B1', 'Security ', merge_format)
        worksheet.write('C1', 'Weight', merge_format)
        worksheet.write('D1', 'Date', merge_format)
        worksheet.write('E1', 'Reason For Change', merge_format)
        worksheet.write('F1', 'CDE Weight', merge_format)
        worksheet.write('G1', 'CDE Portfolio Name', merge_format)
        worksheet.write('H1', 'Base Tp 1yr', merge_format)
        worksheet.write('I1', 'Bear Tp 1yr', merge_format)
        worksheet.write('J1', 'Base Con 1yr', merge_format)
        worksheet.write('K1', 'Bear Con 1yr', merge_format)
        worksheet.write('L1', 'Base Tp 3yr', merge_format)
        worksheet.write('M1', 'Bear Tp 3yr', merge_format)
        worksheet.write('N1', 'Base Con 3yr', merge_format)
        worksheet.write('O1', 'Bear Con 3yr', merge_format)
        worksheet.write('P1', 'Base Multiple 1yr', merge_format)
        worksheet.write('Q1', 'Bear Multiple 1yr', merge_format)
        worksheet.write('R1', 'Valuation Str', merge_format)

    def __write_data(self, sheet, direction, analyst):
        worksheet = self.workbook.get_worksheet_by_name(sheet)
        percentage_format = self.workbook.add_format()
        percentage_format.set_num_format(0x0a)
        count = 2

        if direction == 'long':
            for stock, data in self.longs.items():
                if not data['is_live']:
                    worksheet.write('A{}'.format(count), 'AIM BEST IDEAS {} LONG'.format(analyst))
                    worksheet.write('B{}'.format(count), stock)
                    worksheet.write('C{}'.format(count), data['weight']*100)
                    now = datetime.datetime.now()
                    worksheet.write('D{}'.format(count), now.strftime('%m/%d/%y'))
                    worksheet.write('E{}'.format(count), data['roc'])
                    worksheet.write('F{}'.format(count), data['weight']*100)
                    worksheet.write('G{}'.format(count), 'AIM BEST IDEAS {} LONG'.format(analyst))
                    worksheet.write('H{}'.format(count), data['base_tp_1yr'])
                    worksheet.write('I{}'.format(count), data['bear_tp_1yr'])
                    worksheet.write('J{}'.format(count), data['base_con_1yr'])
                    worksheet.write('K{}'.format(count), data['bear_con_1yr'])
                    worksheet.write('L{}'.format(count), data['base_tp_3yr'])
                    worksheet.write('M{}'.format(count), data['bear_tp_3yr'])
                    worksheet.write('N{}'.format(count), data['base_con_3yr'])
                    worksheet.write('O{}'.format(count), data['bear_con_3yr'])
                    worksheet.write('P{}'.format(count), data['base_multiple_1yr'])
                    worksheet.write('Q{}'.format(count), data['bear_multiple_1yr'])
                    worksheet.write('R{}'.format(count), data['valuation_str'])
                    count += 1
        elif direction == 'short':
            for stock, data in self.shorts.items():
                if not data['is_live']:
                    worksheet.write('A{}'.format(count), 'AIM BEST IDEAS {} SHORT'.format(analyst))
                    worksheet.write('B{}'.format(count), stock)
                    worksheet.write('C{}'.format(count), -data['weight']*100)
                    now = datetime.datetime.now()
                    worksheet.write('D{}'.format(count), now.strftime('%m/%d/%y'))
                    worksheet.write('E{}'.format(count), data['roc'])
                    worksheet.write('F{}'.format(count), -data['weight']*100)
                    worksheet.write('G{}'.format(count), 'AIM BEST IDEAS {} SHORT'.format(analyst))
                    worksheet.write('H{}'.format(count), data['base_tp_1yr'])
                    worksheet.write('I{}'.format(count), data['bear_tp_1yr'])
                    worksheet.write('J{}'.format(count), data['base_con_1yr'])
                    worksheet.write('K{}'.format(count), data['bear_con_1yr'])
                    worksheet.write('L{}'.format(count), data['base_tp_3yr'])
                    worksheet.write('M{}'.format(count), data['bear_tp_3yr'])
                    worksheet.write('N{}'.format(count), data['base_con_3yr'])
                    worksheet.write('O{}'.format(count), data['bear_con_3yr'])
                    worksheet.write('P{}'.format(count), data['base_multiple_1yr'])
                    worksheet.write('Q{}'.format(count), data['bear_multiple_1yr'])
                    worksheet.write('R{}'.format(count), data['valuation_str'])
                    count += 1

