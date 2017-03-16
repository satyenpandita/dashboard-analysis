import paramiko
import json
from os.path import abspath, exists

def ftp_upload(file_path):
    try:
        username, password = get_credentials()
        transport = paramiko.Transport(('sftp.bloomberg.com', 22))
        transport.connect(username=username, password=password)
        sftp = paramiko.SFTPClient.from_transport(transport)
        sftp.put(file_path, "/", confirm=True)
        sftp.close()
        transport.close()
        print('Upload done.')
    except FileNotFoundError as e:
        print(str(e))
    except Exception as e:
        print(str(e))


def get_credentials():
    with open("utils/secrets.json") as data_file:
        data = json.load(data_file)
        credentials = data["ftp_credentials"]
        return credentials['username'], credentials['password']