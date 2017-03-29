import sys
import os
from xlrd import open_workbook
from parsers.DashboardParserV2 import DashboardParserV2
from parsers.portfolio.portfolio_parser import PortfolioParser
from exporter.report_card_exporters import ReportCardExporter
from exporter.exporter import Exporter
from utils.ftp_upload import ftp_upload

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
                parser = PortfolioParser(worksheet, file)
                parser.generate_upload_file(file.split(".")[0])
                #parser.ftp_upload()
    elif sys.argv[1] == 'publish' and len(sys.argv) == 3:
        for file in os.listdir('uploaded_files/output'):
            if 'xls' in file[-4:]:
                print(file)
                if sys.argv[2] in file.lower():
                    ftp_upload("uploaded_files/output/{}".format(file), file)



