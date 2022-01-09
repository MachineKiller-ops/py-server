from flask import Flask, send_from_directory
from flask import jsonify

app = Flask(__name__, static_url_path='')


@app.route("/")
def index():

    file_name = 'SM.xlsx'
    return app.send_static_file(filename=file_name)


@app.after_request
def set_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Headers"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "*"
    return response
