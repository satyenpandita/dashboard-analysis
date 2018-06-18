import smtplib, os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders
from config.celery import app
from utils.upload_ops import get_user_from_stock, get_user_email


@app.task()
def send_mail(send_from, send_to, subject, text, files=[], server="smtp.office365.com", port=587, username='',
              password='', is_tls=True):
    try:
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = COMMASPACE.join(send_to)
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = subject
        msg.attach(MIMEText(text))

        for f in files:
            part = MIMEBase('application', "octet-stream")
            part.set_payload( open(f,"rb").read() )
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="{0}"'.format(os.path.basename(f)))
            msg.attach(part)

        smtp = smtplib.SMTP(server, port)
        if is_tls: smtp.starttls()
        smtp.login(username, password)
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.quit()
    except Exception as e:
        print(str(e))
        send_mail.retry(exc=e, max_retries=2, countdown=10)


@app.task()
def send_dashboard_email(exporter_file_list, stock_code):
    subject = "Dashboard Published"
    recipients = ["datascience@auroim.com"]
    if stock_code:
        subject = "Dashboard Published for {}".format(stock_code)
        email = get_user_from_stock(stock_code)
        if email is not None:
            recipients.append(email)

    send_mail("ppal@auroim.com",
              recipients,
              subject,
              "Dashbord Published",
              exporter_file_list,
              username="ppal@auroim.com",
              password="AuroOct2016")


@app.task()
def best_ideas_notification_email(analyst):
    recipients = ["ppal@auroim.com"]
    # email = get_user_email(analyst)
    # if email:
    #     recipients.append(email)
    send_mail("ppal@auroim.com",
              recipients,
              "Best Ideas Uploaded to Bloomberg",
              "Hi {}, \n Your best Ideas have been uploaded to Bloomberg. Please refresh your monitors and check",
              files=[],
              username="ppal@auroim.com",
              password="AuroOct2016")