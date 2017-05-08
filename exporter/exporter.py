import xlsxwriter
from utils.upload_ops import ftp_upload
from utils.email_sender import send_mail
from exporter.consolidated_exporter_daily import ConsolidatedExporterDaily
from exporter.consolidated_exporter_fiscal import ConsolidatedExporterFiscal


class Exporter:
    def __init__(self):
        self.workbook_fiscal_base = None
        self.workbook_fiscal_bear = None
        self.workbook_fiscal_bull = None
        self.workbook_daily1 = None
        self.workbook_daily2 = None

    def export(self, stock_code=None):
        self.workbook_fiscal_base = xlsxwriter.Workbook("/var/www/output/fiscal_base.xlsx")
        self.workbook_fiscal_bear = xlsxwriter.Workbook("/var/www/output/fiscal_bear.xlsx")
        self.workbook_fiscal_bull = xlsxwriter.Workbook("/var/www/output/fiscal_bull.xlsx")
        self.workbook_daily1 = xlsxwriter.Workbook("/var/www/output/daily1.xlsx")
        self.workbook_daily2 = xlsxwriter.Workbook("/var/www/output/daily2.xlsx")

        # self.workbook_fiscal_base = xlsxwriter.Workbook("uploaded_files/output/fiscal_base.xlsx")
        # self.workbook_fiscal_bear = xlsxwriter.Workbook("uploaded_files/output/fiscal_bear.xlsx")
        # self.workbook_fiscal_bull = xlsxwriter.Workbook("uploaded_files/output/fiscal_bull.xlsx")
        # self.workbook_daily1 = xlsxwriter.Workbook("uploaded_files/output/daily1.xlsx")
        # self.workbook_daily2 = xlsxwriter.Workbook("uploaded_files/output/daily2.xlsx")
        try:
            self.workbook_daily1 = ConsolidatedExporterDaily.export(self.workbook_daily1, 'Daily1', stock_code)
        finally:
            self.workbook_daily1.close()
        try:
            self.workbook_daily2 = ConsolidatedExporterDaily.export(self.workbook_daily2, 'Daily2', stock_code)
        finally:
            self.workbook_daily2.close()
        try:
            self.workbook_fiscal_base = ConsolidatedExporterFiscal.export(self.workbook_fiscal_base, 'Fiscal Base',
                                                                          'base', stock_code)
        finally:
            self.workbook_fiscal_base.close()
        try:
            self.workbook_fiscal_bull = ConsolidatedExporterFiscal.export(self.workbook_fiscal_bull, 'Fiscal Bull',
                                                                          'bull', stock_code)
        finally:
            self.workbook_fiscal_bull.close()
        try:
            self.workbook_fiscal_bear = ConsolidatedExporterFiscal.export(self.workbook_fiscal_bear, 'Fiscal Bear',
                                                                          'bear', stock_code)
        finally:
            self.workbook_fiscal_bear.close()

    def ftp_upload(self):
        ftp_upload.delay(self.workbook_daily1.filename, "daily1.xlsx")
        ftp_upload.delay(self.workbook_daily2.filename, "daily2.xlsx")
        ftp_upload.delay(self.workbook_fiscal_base.filename, "fiscal_base.xlsx")
        ftp_upload.delay(self.workbook_fiscal_bear.filename, "fiscal_bear.xlsx")
        ftp_upload.delay(self.workbook_fiscal_bull.filename, "fiscal_bull.xlsx")

    def send_email(self):
        send_mail.delay("ppal@auroim.com",
                        ["datascience@auroim.com"],
                        "Dashbord Published",
                        "Dashbord Published",
                        [self.workbook_daily1.filename,
                         self.workbook_daily2.filename,
                         self.workbook_fiscal_base.filename,
                         self.workbook_fiscal_bear.filename,
                         self.workbook_fiscal_bull.filename],
                        username="ppal@auroim.com",
                        password="AuroOct2016")

    def send_email_me(self):
        send_mail.delay("ppal@auroim.com",
                        ["ppal@auroim.com"],
                        "Dashbord Published",
                        "Dashbord Published",
                        [self.workbook_daily1.filename,
                         self.workbook_daily2.filename,
                         self.workbook_fiscal_base.filename,
                         self.workbook_fiscal_bear.filename,
                         self.workbook_fiscal_bull.filename],
                        username="ppal@auroim.com",
                        password="AuroOct2016")

    def export_and_upload(self, stock_code=None):
        self.export(stock_code)
        self.ftp_upload()
        self.send_email()
        return "Files Generated Upload Queued"
