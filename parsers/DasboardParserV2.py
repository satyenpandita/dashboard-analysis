import xlwings as xw

class DashboardParserV2(object):

    def __init__(self, filepath):
        wb = xw.Book(filepath)
        worksheet = wb.sheets[0]
        print("Read the file")