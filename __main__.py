import os
from xlrd import open_workbook


from parsers.DashboardParser import DashboardParser


    # if len(sys.argv) < 2:
    #     print("Please supply a valid package name")
    # else:
if __name__ == '__main__':
    workbook = open_workbook('dashboard-analysis/data/dashboard.xlsx')
    worksheet = workbook.sheet_by_index(0)
    dparser = DashboardParser(worksheet)
    dparser.save_dashboard()
