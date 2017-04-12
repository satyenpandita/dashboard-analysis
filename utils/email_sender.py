import smtplib, os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders
from config.celery import app


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
        smtp.login(username,password)
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.quit()
    except Exception as e:
        print(str(e))
        send_mail.retry(exc=e, max_retries=2, countdown=10)