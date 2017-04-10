import paramiko
import json
import os
from boto.s3.connection import S3Connection
from boto.s3.key import Key


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


def s3_upload(file):
    complete_path = 'uploaded_files/dashboard/originals/{}'.format(file)
    key, secret, host = get_s3_credentials()
    conn = S3Connection(key, secret, host=host)
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
    with open("constants/secrets.json") as data_file:
        data = json.load(data_file)
        credentials = data["ftp_credentials"]
        return credentials['username'], credentials['password']


def get_s3_credentials():
    with open("constants/secrets.json") as data_file:
        data = json.load(data_file)
        credentials = data["s3_credentials"]
        return credentials['s3_key'], credentials['s3_secret'], credentials['host']


def get_users():
    with open("constants/users.json") as data_file:
        data = json.load(data_file)
        return data
