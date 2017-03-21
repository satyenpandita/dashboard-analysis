import sys
import os
from xlrd import open_workbook
from parsers.DashboardParserV2 import DashboardParserV2
from parsers.portfolio.portfolio_parser import PortfolioParser
from exporter.report_card_exporters import ReportCardExporter
from exporter.exporter import Exporter

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Please supply valid options")
    elif sys.argv[1] == 'parse':
        for file in os.listdir('uploaded_files/dashboard'):
            if 'xls' in file[-4:]:
                print(file)
                workbook = open_workbook('uploaded_files/dashboard/'+file)
                worksheet = workbook.sheet_by_index(0)
                dparser = DashboardParserV2(worksheet)
                dparser.save_dashboard()
    elif sys.argv[1] == 'reportcard':
        exporter = ReportCardExporter()
        exporter.export_report_card()
    elif sys.argv[1] == 'consolidated':
        exporter = Exporter()
        exporter.export()
    elif sys.argv[1] == 'portfolio':
        for file in os.listdir('uploaded_files/portfolio'):
            if 'xls' in file[-4:]:
                print(file)
                workbook = open_workbook('uploaded_files/portfolio/'+file)
                worksheet = workbook.sheet_by_index(0)
                parser = PortfolioParser(worksheet)
                parser.generate_upload_file(file)
                #parser.ftp_upload()

