from config.mongo_config import db
import xlsxwriter
from models.Dashboard import Dashboard


class ReportCardExporter(object):

    def __init__(self):
        self.workbook = None

    def export_report_card(self):
        self.workbook = xlsxwriter.Workbook("aim_workplan.xlsx")
        self.__write_headers()
        self.__write_data()
        self.workbook.close()

    def __write_headers(self):
        worksheet = self.workbook.add_worksheet('Sheet1')
        merge_format = self.workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        worksheet.merge_range('D1:D3', 'Model in Dropbox', merge_format)
        worksheet.merge_range('E1:E3', 'Auroville Dashboard', merge_format)
        worksheet.merge_range('F1:H1', 'Target Price', merge_format)
        worksheet.merge_range('F2:F3', 'Base: 1-yr TP', merge_format)
        worksheet.merge_range('G2:G3', 'Base: 3-yr TP', merge_format)
        worksheet.merge_range('H2:H3', 'Bear: 1-yr TP', merge_format)
        worksheet.merge_range('I1:J1', 'Operating model', merge_format)
        worksheet.write('K1','Qual', merge_format)
        worksheet.merge_range('K2:K3','Most LikeLy OutCome', merge_format)
        worksheet.merge_range('L1:N1', 'Data Science', merge_format)
        worksheet.merge_range('L2:L3', 'KPI 1', merge_format)
        worksheet.merge_range('M2:M3', 'KPI 2', merge_format)
        worksheet.merge_range('N2:N3', 'KPI 3', merge_format)
        worksheet.merge_range('O1:O3', 'KPI 3', merge_format)

    def __write_data(self):
        row = 4
        data = self.__get_grouped_analyst_data()
        worksheet = self.workbook.get_worksheet_by_name('Sheet1')
        err_format = self.workbook.add_format({'bg_color': '#FF0000'})
        corr_format = self.workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter'})
        for analyst, data in data.items():
            worksheet.merge_range(row, 0, row, 1, analyst)
            row += 1
            for idx, dashboard in enumerate(data):
                dsh = Dashboard(dashboard)
                self.__populate_data(dsh, worksheet, corr_format, err_format, row)
                row += 1
            row += 1

    @staticmethod
    def __populate_data(dsh, worksheet, corr_format, err_format, row):
        worksheet.write(row, 0, dsh.direction_char())
        worksheet.write(row, 1, dsh.company, corr_format)
        worksheet.write(row, 2, dsh.stock_code, corr_format)
        worksheet.write(row, 3, dsh.model_present(), corr_format)
        worksheet.write(row, 4, dsh.dashboard_present(), corr_format)
        if dsh.base_1yr_present():
            worksheet.write(row, 5, u'\u2713', corr_format)
        else:
            worksheet.write(row, 5, '', err_format)

        if dsh.base_3yr_present():
            worksheet.write(row, 6, u'\u2713', corr_format)
            worksheet.write(row, 14, dsh.target_price['base']['pt_3year'])
        else:
            worksheet.write(row, 6, '', err_format)

        if dsh.bear_1yr_present():
            worksheet.write(row, 7, u'\u2713', corr_format)
        else:
            worksheet.write(row, 7, '', err_format)

        if dsh.likely_outcome:
            worksheet.write(row, 10, u'\u2713', corr_format)
        else:
            worksheet.write(row, 10, '', err_format)

        if dsh.kpi1_present():
            worksheet.write(row, 11, dsh.get_kpi('kpi1'), corr_format)
        else:
            worksheet.write(row, 11, '', err_format)

        if dsh.kpi2_present():
            worksheet.write(row, 12, dsh.get_kpi('kpi2'), corr_format)
        else:
            worksheet.write(row, 12, '', err_format)

        if dsh.kpi3_present():
            worksheet.write(row, 13, dsh.get_kpi('kpi3'), corr_format)
        else:
            worksheet.write(row, 13, '', err_format)

    @staticmethod
    def __get_grouped_analyst_data():
        data = dict()
        for dashboard in db.dashboards.find({}):
            if dashboard['analyst_primary'] not in data.keys():

                data[dashboard['analyst_primary']] = [] + [dashboard]
            else:
                data[dashboard['analyst_primary']] = data[dashboard['analyst_primary']] + [dashboard]
        return data