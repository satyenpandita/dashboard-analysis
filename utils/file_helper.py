import os
from utils.upload_ops import ftp_upload
from utils.email_sender import best_ideas_notification_email


def handle_uploaded_file(complete_path, file):
    try:
        with open(complete_path, 'wb+') as destination:
            for chunk in file.chunks():
                destination.write(chunk)
        return True
    except Exception as e:
        return False


def portfolio_upload_util(analyst):
    file_found = False
    for file in os.listdir('/var/www/output'):
        if 'xls' in file[-4:] and analyst in file:
            file_found = True
            ftp_upload.delay("/var/www/output/{}".format(file), file)
    if file_found:
        best_ideas_notification_email.delay(analyst)
    return True