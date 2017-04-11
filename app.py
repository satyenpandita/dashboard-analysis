from flask import Flask, request, jsonify
from exporter.exporter import Exporter
from parsers.DashboardParserV2 import DashboardParserV2
from parsers.DashboardParserV3 import DashboardParserV3
from parsers.portfolio.portfolio_parser import PortfolioParser
from parsers.portfolio.portfolio_parser_v2 import PortfolioParserV2
from xlrd import open_workbook
from werkzeug.contrib.fixers import ProxyFix
import logging
from logging.handlers import RotatingFileHandler
from utils.upload_ops import ftp_upload
from utils.upload_ops import s3_upload
import os

app = Flask(__name__)


@app.route('/')
def index():
    return "Hello, World!"


@app.route('/dashboard', methods=['POST'])
def dashboard():
    file = request.files['uploadfile']
    complete_name = 'uploaded_files/dashboard/{}'.format(file.filename)
    file.save(complete_name)
    workbook = open_workbook(complete_name)
    worksheet = workbook.sheet_by_index(0)
    dparser = DashboardParserV2(worksheet)
    dparser.save_dashboard()
    return jsonify({'file': file.filename}), 201


@app.route('/archive', methods=['POST'])
def dashboard_archive():
    file = request.files['uploadfile']
    complete_name = 'uploaded_files/dashboard/originals/{}'.format(file.filename)
    file.save(complete_name)
    s3_upload.delay(file.filename)
    return jsonify({'file': file.filename}), 201


@app.route('/dashboard2', methods=['POST'])
def dashboard2():
    file = request.files['uploadfile']
    complete_name = 'uploaded_files/dashboard/{}'.format(file.filename)
    file.save(complete_name)
    workbook = open_workbook(complete_name)
    dparser = DashboardParserV3(workbook)
    dparser.save_dashboard()
    exporter = Exporter()
    exporter.export()
    exporter.ftp_upload()
    return jsonify({'file': file.filename}), 201


@app.route('/portfolio', methods=['POST'])
def portfolio():
    file = request.files['uploadfile']
    app.logger.info(file.filename)
    complete_name = 'uploaded_files/portfolio/{}'.format(file.filename)
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
    complete_name = 'uploaded_files/portfolio/{}'.format(file.filename)
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
        response_dict = dict()
        analyst = request.args['analyst']
        for file in os.listdir('uploaded_files/output'):
            if 'xls' in file[-4:] and analyst in file:
                app.logger.info(file)
                res = ftp_upload.delay("uploaded_files/output/{}".format(file), file)
                response_dict[file] = res
        return jsonify(response_dict), 201
    else:
        return jsonify({'error':'No Analyst name'})


app.wsgi_app = ProxyFix(app.wsgi_app)
if __name__ == '__main__':
    app.run()