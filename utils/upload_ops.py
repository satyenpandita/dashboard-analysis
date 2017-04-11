import json
import os
import paramiko
from boto.s3.connection import S3Connection
from boto.s3.key import Key
import boto
from config.celery import app


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
    complete_path = 'uploaded_files/dashboard/originals/{}'.format(file)
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
    with open(os.path.abspath("constants/secrets.json")) as data_file:
        data = json.load(data_file)
        credentials = data["ftp_credentials"]
        return credentials['username'], credentials['password']


def get_users():
    with open(os.path.abspath("constants/secrets.json")) as data_file:
        data = json.load(data_file)
        return data
