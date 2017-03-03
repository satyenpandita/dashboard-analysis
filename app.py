from flask import Flask, request, jsonify, abort

app = Flask(__name__)


@app.route('/')
def index():
    return "Hello, World!"


@app.route('/dashboard', methods=['POST'])
def dashboard_update():
    file = request.files['uploadfile']
    print(file.filename)
    file.save('uploaded_files/{}'.format(file.filename))
    return jsonify({'file': file.filename}), 201


if __name__ == '__main__':
    app.run(debug=True)