import xlsxwriter
import datetime


class PortfolioExporter(object):

    def __init__(self, longs, shorts):
        self.workbook = None
        self.longs = longs
        self.shorts = shorts

    def export(self):
        output_path = r'uploaded_files/output/portfolio.xlsx'
        self.workbook = xlsxwriter.Workbook(output_path)
        try:
            self.__write_headers('Sheet1')
            self.__write_data('Sheet1')
        except Exception as e:
            print(str(e))
        finally:
            self.workbook.close()
            return output_path

    def __write_headers(self, sheet):
        worksheet = self.workbook.add_worksheet(sheet)
        merge_format = self.workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        worksheet.write('A1', 'Portfolio Name', merge_format)
        worksheet.write('B1', 'Security ', merge_format)
        worksheet.write('C1', 'Weight', merge_format)
        worksheet.write('D1', 'Date', merge_format)

    def __write_data(self, sheet):
        worksheet = self.workbook.get_worksheet_by_name(sheet)
        percentage_format = self.workbook.add_format()
        percentage_format.set_num_format(0x0a)
        count = 2
        for stock, weight in self.longs.items():
            worksheet.write('A{}'.format(count), 'AURO IM MC')
            worksheet.write('B{}'.format(count), stock)
            worksheet.write('C{}'.format(count), weight, percentage_format)
            now = datetime.datetime.now()
            worksheet.write('D{}'.format(count), now.strftime('%d-%m-%y'))
            count += 1

        for stock, weight in self.shorts.items():
            worksheet.write('A{}'.format(count), 'AURO IM MC')
            worksheet.write('B{}'.format(count), stock)
            worksheet.write('C{}'.format(count), -weight, percentage_format)
            now = datetime.datetime.now()
            worksheet.write('D{}'.format(count), now.strftime('%d-%m-%y'))
            count += 1

