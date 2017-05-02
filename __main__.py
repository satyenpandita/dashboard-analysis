import sys
import os
from xlrd import open_workbook
from parsers.DashboardParserV3 import DashboardParserV3
from parsers.portfolio.portfolio_parser_v2 import PortfolioParserV2
from exporter.report_card_exporters import ReportCardExporter
from exporter.exporter import Exporter
from utils.upload_ops import ftp_upload
from utils.upload_ops import s3_upload

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Please supply valid options")
    elif sys.argv[1] == 'parse':
        for file in os.listdir('uploaded_files/dashboard'):
            if 'xls' in file[-4:]:
                print(file)
                workbook = open_workbook('uploaded_files/dashboard/'+file)
                dparser = DashboardParserV3(workbook)
                dparser.save_dashboard()
    elif sys.argv[1] == 'reportcard':
        exporter = ReportCardExporter()
        exporter.export_report_card()
    elif sys.argv[1] == 'consolidated':
        exporter = Exporter()
        exporter.export()
    elif sys.argv[1] == 's3':
        s3_upload('ONDK US Equity_03-04-17.xlsx')
    elif sys.argv[1] == 'portfolio':
        for file in os.listdir('uploaded_files/portfolio'):
            if 'xls' in file[-4:]:
                print(file)
                workbook = open_workbook('uploaded_files/portfolio/'+file)
                worksheet = workbook.sheet_by_index(0)
                parser = PortfolioParserV2(worksheet, file)
                parser.generate_upload_file(file.split(".")[0])
    elif sys.argv[1] == 'publish' and len(sys.argv) == 3:
        for file in os.listdir('uploaded_files/output'):
            if 'xls' in file[-4:]:
                print(file)
                if sys.argv[2] in file.lower():
                    ftp_upload.delay("uploaded_files/output/{}".format(file), file)



