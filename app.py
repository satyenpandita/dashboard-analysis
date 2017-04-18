from flask import Flask, request, jsonify
from exporter.exporter import Exporter
from parsers.DashboardParserV2 import DashboardParserV2
from parsers.DashboardParserV3 import DashboardParserV3
from parsers.portfolio.portfolio_parser import PortfolioParser
from parsers.portfolio.portfolio_parser_v2 import PortfolioParserV2
from xlrd import open_workbook
from werkzeug.contrib.fixers import ProxyFix
from utils.upload_ops import ftp_upload
from utils.upload_ops import s3_upload
import os
from config.mongo_config import db
from models.CumulativeDashBoard import CumulativeDashBoard
from models.DashboardV2 import DashboardV2

app = Flask(__name__)


@app.route('/')
def index():
    return "Hello, World!!!!!!!!!!!"


@app.route('/dashboard', methods=['POST'])
def dashboard():
    file = request.files['uploadfile']
    complete_name = '/var/www/dashboard/{}'.format(file.filename)
    file.save(complete_name)
    workbook = open_workbook(complete_name)
    worksheet = workbook.sheet_by_index(0)
    dparser = DashboardParserV2(worksheet)
    dparser.save_dashboard()
    return jsonify({'file': file.filename}), 201


@app.route('/archive', methods=['POST'])
def dashboard_archive():
    file = request.files['uploadfile']
    complete_name = '/var/www/dashboard/originals/{}'.format(file.filename)
    file.save(complete_name)
    s3_upload.delay(file.filename)
    return jsonify({'file': file.filename}), 201


@app.route('/dashboard2', methods=['POST'])
def dashboard2():
    file = request.files['uploadfile']
    complete_name = '/var/www/dashboard/{}'.format(file.filename)
    file.save(complete_name)
    workbook = open_workbook(complete_name)
    dparser = DashboardParserV3(workbook)
    dparser.save_dashboard()
    exporter = Exporter()
    exporter.export()
    # exporter.ftp_upload()
    return jsonify({'file': file.filename}), 201


@app.route('/portfolio', methods=['POST'])
def portfolio():
    file = request.files['uploadfile']
    app.logger.info(file.filename)
    complete_name = '/var/www/portfolio/{}'.format(file.filename)
    file.save(complete_name)
    workbook = open_workbook(complete_name)
    worksheet = workbook.sheet_by_index(0)
    app.logger.info("Parser Starting")
    parser = PortfolioParser(worksheet, file.filename)
    parser.generate_upload_file(file.filename)
    app.logger.info("Email Start")
    parser.send_email.delay()
    app.logger.info("Email End")
    return jsonify({'file': file.filename}), 201


@app.route('/portfolio2', methods=['POST'])
def portfolio2():
    file = request.files['uploadfile']
    app.logger.info(file.filename)
    complete_name = '/var/www/portfolio/{}'.format(file.filename)
    file.save(complete_name)
    workbook = open_workbook(complete_name)
    worksheet = workbook.sheet_by_index(0)
    app.logger.info("Parser Starting")
    parser = PortfolioParserV2(worksheet, file.filename)
    parser.generate_upload_file(file.filename)
    app.logger.info("Email Start")
    parser.send_email()
    app.logger.info("Email End")
    return jsonify({'file': file.filename}), 201


@app.route('/portfolio_upload', methods=['GET'])
def portfolio_upload():
    if 'analyst' in request.args:
        analyst = request.args['analyst']
        for file in os.listdir('/var/www/output'):
            if 'xls' in file[-4:] and analyst in file:
                app.logger.info(file)
                ftp_upload.delay("/var/www/output/{}".format(file), file)
        return jsonify({'success': 'Tasks Queued'}), 201
    else:
        return jsonify({'error': 'No Analyst name'})


@app.route('/dashboard_upload', methods=['GET'])
def dashboard_upload():
    exporter = Exporter()
    exporter.export()
    exporter.ftp_upload()
    exporter.send_email()
    return jsonify({'response': "Files Generated and Tasks queued"}), 201


@app.route('/dashboard_email', methods=['GET'])
def dashboard_email():
    exporter = Exporter()
    exporter.export()
    exporter.send_email_me()
    return jsonify({'response': "Files Generated and Tasks queued"}), 201


@app.route('/dashboard_generate', methods=['GET'])
def dashboard_generate():
    exporter = Exporter()
    exporter.export()
    return jsonify({'response': "Files Generated"}), 201


@app.route('/dashboard_upload_only', methods=['GET'])
def dashboard_upload_only():
    ftp_upload.delay("/var/www/output/fiscal_base.xlsx", "fiscal_base.xlsx")
    ftp_upload.delay("/var/www//output/fiscal_bear.xlsx", "fiscal_bear.xlsx")
    ftp_upload.delay("/var/www//output/fiscal_bull.xlsx", "fiscal_bull.xlsx")
    ftp_upload.delay("/var/www//output/daily1.xlsx", "daily1.xlsx")
    ftp_upload.delay("/var/www//output/daily2.xlsx", "daily2.xlsx")
    return jsonify({'response': "Upload Queued"}), 201


@app.route('/migration_old', methods=['GET'])
def migration_old():
    for idx, cum_dsh in enumerate(db.cumulative_dashboards.find({})):
        cum_dsh = CumulativeDashBoard.from_dict(cum_dsh)
        dsh_base = DashboardV2(cum_dsh.base)
        dsh_bull = DashboardV2(cum_dsh.bull)
        dsh_bear = DashboardV2(cum_dsh.bear)
        dsh_base.old = False
        dsh_bear.old = False
        dsh_bull.old = False
        updated_cum_dsh = CumulativeDashBoard(cum_dsh.stock_code, dsh_base, dsh_bull,dsh_bear)
        updated_cum_dsh.save()
    return jsonify({'response': "Upload Queued"}), 201


app.wsgi_app = ProxyFix(app.wsgi_app)
if __name__ == '__main__':
    app.run()