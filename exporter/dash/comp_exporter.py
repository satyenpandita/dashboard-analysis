import xlsxwriter
from config.mongo_config import db
from models.CumulativeDashBoard import CumulativeDashBoard
from models.DashboardV2 import DashboardV2
from utils.email_sender import send_mail


def export_comps():
    comps_dict = {}
    workbook_comps = xlsxwriter.Workbook("/var/www/output/comps.xlsx")
    data = db.cumulative_dashboards.find({})
    for idx, dashboard in enumerate(data):
        cum_dsh = CumulativeDashBoard.from_dict(dashboard)
        dsh_init = getattr(cum_dsh, 'base')
        if dsh_init is not None:
            dsh = DashboardV2(dsh_init)
            comps_dict[dsh.stock_code] = dsh.tam.get('key_comps', None)
    write_excel(workbook_comps, comps_dict)
    send_mail.delay("ppal@auroim.com", ["ppal@auroim.com"], "Comps", "Comps Data", ["/var/www/output/comps.xlsx"],
                    username="ppal@auroim.com", password="AuroOct2016")
    return "Success"


def colnum_string(n):
    div = n
    string = ""
    while div > 0:
        module = (div - 1) % 26
        string = chr(65 + module) + string
        div = int((div - module) / 26)
    return string


def write_excel(workbook, data):
    worksheet = workbook.add_worksheet("Data")

    # Headers
    merge_format = workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    worksheet.write('A1', 'Stock Code ', merge_format)
    worksheet.write('B1', 'Comp1', merge_format)
    worksheet.write('C1', 'Comp2', merge_format)
    worksheet.write('D1', 'Comp3', merge_format)

    #Data
    count = 2
    for key, val in data.items():
        worksheet.write("A{}".format(count), key)
        for idx, comp in enumerate(val):
            worksheet.write("{}{}".format(colnum_string(idx+2), count), comp)
        count += 1
    workbook.close()
