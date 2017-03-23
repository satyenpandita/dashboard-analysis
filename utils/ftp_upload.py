import paramiko
import json
import os


def ftp_upload(file_path, file_name):
    #try:
        username, password = get_credentials()
        transport = paramiko.Transport(('sftp.bloomberg.com', 22))
        transport.connect(username=username, password=password)
        sftp = paramiko.SFTPClient.from_transport(transport)
        local_path = os.path.abspath(file_path)
        print(local_path)
        rdata = sftp.put(local_path, "/{}".format(file_name), confirm=True)
        print(rdata)
        sftp.close()
        transport.close()
        print('Upload done.')
    #except FileNotFoundError as e:
    #    print(str(e))
    #except Exception as e:
    #    print(str(e))


def get_credentials():
    with open("constants/secrets.json") as data_file:
        data = json.load(data_file)
        credentials = data["ftp_credentials"]
        return credentials['username'], credentials['password']


def get_users():
    with open("constants/users.json") as data_file:
        data = json.load(data_file)
        return data
