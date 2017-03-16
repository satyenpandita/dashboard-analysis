from ftplib import FTP


def ftp_upload(file_path):
    try:
        output_path = file_path
        session = FTP('sftp.bloomberg.com', 'u97335997', 'rn]dT[27OQOE=-k8')
        file = open(output_path, 'rb')  # file to send
        session.storbinary('portfolio.xlsx', file)  # send the file
        file.close()  # close file and FTP
        session.quit()
    except FileNotFoundError as e:
        print(str(e))
    except Exception as e:
        print(str(e))