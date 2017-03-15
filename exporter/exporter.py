import xlsxwriter
from exporter.consolidated_exporter_daily import ConsolidatedExporterDaily
from exporter.consolidated_exporter_fiscal import ConsolidatedExporterFiscal


class Exporter:
    def __init__(self):
        self.workbook = None

    def export(self):
        self.workbook = xlsxwriter.Workbook("uploaded_files/output/upload.xlsx")
        try:
            self.workbook = ConsolidatedExporterDaily.export(self.workbook, 'Daily')
            self.workbook = ConsolidatedExporterFiscal.export(self.workbook, 'Fiscal')
        except Exception as e:
            print(str(e))
        finally:
            self.workbook.close()