from flask import Flask, request
from flask import json
from openpyxl import load_workbook
import os

app = Flask(__name__, static_url_path='')


def print_conteudo(event_data):
    print(event_data.get("body"))


@app.route("/", methods=['POST'])
def index():
    event_data = request.json
    print_conteudo(event_data)
    # load excel file
    workbook = load_workbook(filename=os.path.join('static', "SM.xlsx"))

    # open workbook
    sheet = workbook.active

    # modify the desired cell
    sheet["H3"] = "SE ATALAIA"

    # save the file
    workbook.save(filename=os.path.join('static', "SMout.xlsx"))

    file_name = 'SMout.xlsx'
    return app.send_static_file(filename=file_name)


if os.path.exists(os.path.join('static', "SMout.xlsx")):
    os.remove(os.path.join('static', "SMout.xlsx"))
else:
    print("Extra file doesn't exists")


@app.after_request
def set_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Headers"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "*"
    return response


if __name__ == "__main__":
    app.run(host='127.0.0.1', port=5005)
