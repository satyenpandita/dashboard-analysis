from flask import Flask, request, jsonify, send_from_directory
from exporter.exporter import Exporter
from parsers.DashboardParserV2 import DashboardParserV2
from parsers.DashboardParserV3 import DashboardParserV3
from parsers.portfolio.portfolio_parser import PortfolioParser
from parsers.portfolio.portfolio_pnl_parser import PortfolioPnlParser
from parsers.portfolio.portfolio_parser_v2 import PortfolioParserV2
from xlrd import open_workbook, XLRDError
from werkzeug.contrib.fixers import ProxyFix
from utils.upload_ops import ftp_upload
from utils.upload_ops import s3_upload
import logging
from config.mongo_config import db
from models.CumulativeDashBoard import CumulativeDashBoard
from models.DashboardV2 import DashboardV2
from utils.error_handlers import handle_500
from utils.file_helper import portfolio_upload_util

app = Flask(__name__, static_folder='static')


@app.before_first_request
def setup_logging():
    if not app.debug:
        # In production mode, add log handler to sys.stderr.
        app.logger.addHandler(logging.StreamHandler())
        app.logger.setLevel(logging.INFO)


@app.route('/')
def index():
    app.logger.info("Hello")
    return "Hello, World!!!!!!!!!!!"


@app.route('/robots.txt')
def static_from_root():
    return send_from_directory(app.static_folder, request.path[1:])


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
    s3_upload.delay(complete_name)
    return jsonify({'file': file.filename}), 201


@app.route('/dashboard2', methods=['POST'])
def dashboard2():
    try:
        file = request.files['uploadfile']
        complete_name = '/var/www/dashboard/{}'.format(file.filename)
        file.save(complete_name)
        workbook = open_workbook(complete_name)
        dparser = DashboardParserV3(workbook)
        dparser.save_dashboard(file=complete_name)
        exporter = Exporter()
        exporter.export_and_upload(dparser.stock_code)
        return jsonify({'file': file.filename}), 201
    except Exception as e:
        app.logger.error("Publish Failed for file : {}".format(file.filename))
        handle_500(e, file)
        return jsonify({'file': file.filename}), 500


@app.route('/portfolio', methods=['POST'])
def portfolio():
    file = request.files['uploadfile']
    app.logger.error(file.filename)
    app.logger.error("File Mime {}".format(file.content_type))
    try:
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
    except XLRDError as e:
        app.logger.info("XLRD Error {}".format(str(e)))
        app.logger.info("File Mime {}".format(file.content_type))
    except Exception as e:
        app.logger.error("Publish Failed for file : {}".format(file.filename))
        handle_500(e, file)


@app.route('/portfolio2', methods=['POST'])
def portfolio2():
    file = request.files['uploadfile']
    app.logger.info(file.filename)
    app.logger.info("File Mime {}".format(file.content_type))
    try:
        complete_name = '/var/www/portfolio/{}'.format(file.filename)
        file.save(complete_name)
        workbook = open_workbook(complete_name)
        worksheet = workbook.sheet_by_index(0)
        app.logger.info("Parser Starting")
        parser = PortfolioParserV2(worksheet, file.filename)
        parser.save_and_generate_files()
        app.logger.info("Email Start")
        parser.send_email()
        app.logger.info("Email End")
        return jsonify({'file': file.filename}), 201
    except XLRDError as e:
        app.logger.error("XLRD Error {}".format(str(e)))
        app.logger.error("File Mime {}".format(file.content_type))
        return jsonify({'file': file.filename}), 500
    except Exception as e:
        app.logger.error("Publish Failed for file : {}".format(file.filename))
        app.logger.error("File Mime {}".format(file.content_type))
        app.logger.error("Error Occured".format(str(e)))
        handle_500(e, file)
        return jsonify({'file': file.filename}), 500


@app.route('/portfolio_pnl', methods=['POST'])
def portfolio_pnl():
    file = request.files['uploadfile']
    app.logger.info(file.filename)
    app.logger.info("File Mime {}".format(file.content_type))
    try:
        complete_name = '/var/www/portfolio/{}'.format(file.filename)
        file.save(complete_name)
        app.logger.info("Parser Starting")
        workbook = open_workbook(complete_name)
        parser = PortfolioPnlParser(workbook)
        parser.parse_save_data()
        app.logger.info("Export Started")
        resp = parser.export()
        app.logger.info("Export Ended")
        return jsonify({'file': resp}), 201
    except XLRDError as e:
        app.logger.error("XLRD Error {}".format(str(e)))
        app.logger.error("File Mime {}".format(file.content_type))
        return jsonify({'file': file.filename}), 500
    except Exception as e:
        app.logger.error("Publish Failed for file : {}".format(file.filename))
        app.logger.error("File Mime {}".format(file.content_type))
        handle_500(e, file)
        return jsonify({'file': file.filename}), 500


@app.route('/portfolio_upload', methods=['GET'])
def portfolio_upload():
    if 'analyst' in request.args:
        analyst = request.args['analyst']
        portfolio_upload_util(analyst)
        return jsonify({'success': 'Tasks Queued'}), 201
    else:
        return jsonify({'error': 'No Analyst name'})


@app.route('/dashboard_upload', methods=['GET'])
def dashboard_upload():
    exporter = Exporter()
    exporter.export_and_upload()
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


@app.route('/portfolio', methods=['GET'])
def portfolio_json():
    from models.portfolio.portfolio import Portfolio
    portfolios = Portfolio.objects.to_json()
    return portfolios, 200


@app.route('/comps', methods=['GET'])
def comps():
    from exporter.dash.comp_exporter import export_comps
    status = export_comps()
    return status, 200


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


@app.route('/publish_report', methods=['GET'])
def publish_report():
    from exporter.reports import publish_report
    publish_report.generate_and_send()
    return jsonify({'response': "Report Generated and sent"}), 201


@app.errorhandler(500)
def server_error(e):
    handle_500(e)


app.wsgi_app = ProxyFix(app.wsgi_app)
if __name__ == '__main__':
    app.run(debug=False)