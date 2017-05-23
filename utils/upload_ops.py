import os
import paramiko
from boto.s3.key import Key
import boto
from config.celery import app
from constants import secrets, users


@app.task()
def ftp_upload(file_path, file_name):
    try:
        username, password = get_ftp_credentials()
        transport = paramiko.Transport(('sftp.bloomberg.com', 22))
        transport.connect(username=username, password=password)
        sftp = paramiko.SFTPClient.from_transport(transport)
        local_path = os.path.abspath(file_path)
        print(local_path)
        rdata = sftp.put(local_path, "/{}".format(file_name), confirm=True)
        print(rdata)
        sftp.close()
        transport.close()
        return "Upload Successful"
    except FileNotFoundError as e:
        return str(e)
    except Exception as e:
        return str(e)


@app.task()
def s3_upload(file):
    complete_path = '/var/www/dashboard/originals/{}'.format(file)
    conn = boto.connect_s3()
    conn.host = "s3-us-west-2.amazonaws.com"
    try:
        stock = (file.split('_')[0]).split()[0]
        bucket = conn.get_bucket("aimdashboards", validate=False)
        key = "{}/{}".format(stock, file)
        k = Key(bucket)
        k.key = key
        print(k)
        k.set_contents_from_filename(complete_path)
    except Exception as e:
        print(str(e))
    finally:
        conn.close()
        # os.remove(complete_path)


def get_ftp_credentials():
    credentials = secrets.FTP_CREDENTIALS
    return credentials['username'], credentials['password']


def get_users():
    data = users.USERS
    return data


def get_user_email(user):
    data = users.USERS_EMAILS
    try:
        return data[user]
    except KeyError:
        return None


def get_user_from_stock(stock):
    from config.mongo_config import db
    from models.DashboardV2 import DashboardV2
    doc = db.cumulative_dashboards.find_one({"stock_code" : stock})
    if doc is not None:
        dashboard = DashboardV2(doc['base'])
        return get_user_email(dashboard.analyst_primary)
    else:
        return None
