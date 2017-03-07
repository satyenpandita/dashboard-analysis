import sys
import os
from xlrd import open_workbook
from parsers.DashboardParserV2 import DashboardParserV2
from exporter.report_card_exporters import ReportCardExporter
from exporter.consolidated_exporter import ConsolidatedExporter
from exporter.consolidated_exporter_v2 import ConsolidatedExporterV2
from exporter.consolidated_exporter_v3 import ConsolidatedExporterV3
from utils.name_diff import get_name_list
from utils.cell_functions import find_cell

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Please supply valid options")
    elif sys.argv[1] == 'parse':
        for file in os.listdir('data'):
            if 'xls' in file[-4:]:
                print(file)
                workbook = open_workbook('data/'+file)
                worksheet = workbook.sheet_by_index(0)
                dparser = DashboardParserV2(worksheet)
                dparser.save_dashboard()
    elif sys.argv[1] == 'diff':
        name_list = []
        for file in os.listdir('data'):
            if 'xls' == file[-3:] and '~' != file[0]:
                try:
                    workbook = open_workbook('data/'+file, formatting_info=True)
                    temp_list = get_name_list(workbook, 1)
                    for x in temp_list:
                        if x not in name_list:
                            name_list.append(x)
                except NotImplementedError:
                    print("Formatting not implemented for this file. Convert xlsx to xls")
        print(name_list)
    elif sys.argv[1] == 'reportcard':
        exporter = ReportCardExporter()
        exporter.export_report_card()
    elif sys.argv[1] == 'consolidated2':
        exporter = ConsolidatedExporterV3()
        exporter.export_report_card()
    elif sys.argv[1] == 'consolidated':
        exporter = ConsolidatedExporter()
        exporter.export_report_card()
    elif sys.argv[1] == 'cell_function':
        for file in os.listdir('data'):
            if 'xls' in file[-4:]:
                print(file)
                workbook = open_workbook('data/'+file)
                worksheet = workbook.sheet_by_index(0)
                find_cell(worksheet, 'Leverage and Returns')
