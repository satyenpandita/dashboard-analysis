from os import environ as env
from flask import Flask, request, jsonify, abort
from parsers.DashboardParserV2 import DashboardParserV2
from xlrd import open_workbook
from tornado.wsgi import WSGIContainer
from tornado.httpserver import HTTPServer
from tornado.ioloop import IOLoop

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


if __name__ == '__main__':
    http_server = HTTPServer(WSGIContainer(app))
    http_server.listen(env.get('app.port', 5000))
    IOLoop.instance().start()