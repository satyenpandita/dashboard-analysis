import xlsxwriter
from datetime import datetime
from utils.email_sender import send_mail
from models.CumulativeDashBoard import CumulativeDashBoard

FILE_PATH = "/var/www/DashPublishReport_{}.xlsx".format(datetime.now().strftime('%d%B%y'))


def generate_and_send():
    file_path = generate_publish_report()
    send_mail.delay("ppal@auroim.com",
                    "ppal@auroim.com",
                    "Dashboard Publish Report",
                    "",
                    [file_path],
                    username="ppal@auroim.com",
                    password="AuroOct2016")


def generate_publish_report():
    workbook = xlsxwriter.Workbook(FILE_PATH)
    sheet = workbook.add_worksheet("Report")
    sheet.write("A1", "Ticker")
    sheet.write("B1", "Date")

    data = CumulativeDashBoard.dashboard_publish_report()
    for idx, (tic, pub_date) in enumerate(data):
        sheet.write("A{}".format(2+idx), tic)
        sheet.write("B{}".format(2+idx), pub_date)
    workbook.close()
    return FILE_PATH
