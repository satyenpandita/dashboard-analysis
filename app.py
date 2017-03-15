from flask import Flask, request, jsonify
from parsers.DashboardParserV2 import DashboardParserV2
from xlrd import open_workbook
from werkzeug.contrib.fixers import ProxyFix

app = Flask(__name__)


@app.route('/')
def index():
    return "Hello, World!"


@app.route('/dashboard', methods=['POST'])
def dashboard_update():
    file = request.files['uploadfile']
    complete_name = 'uploaded_files/{}'.format(file.filename)
    file.save(complete_name)
    workbook = open_workbook(complete_name)
    worksheet = workbook.sheet_by_index(0)
    dparser = DashboardParserV2(worksheet)
    dparser.save_dashboard()
    return jsonify({'file': file.filename}), 201

app.wsgi_app = ProxyFix(app.wsgi_app)
if __name__ == '__main__':
    app.run()