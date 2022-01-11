from flask import Flask, request
from flask import json
from openpyxl import load_workbook
import os

app = Flask(__name__, static_url_path='')


def cria_pagina(event_data, page, cont, tipo):
    if(page > 0):
        del event_data[0:32]
    # load excel file
    workbook = load_workbook(filename=os.path.join('static', "SMout.xlsx"))
    page += 1
    # insert at the end (default)
    name = tipo + str(page)
    sheet = workbook[name]
    for i, new_val in enumerate(event_data):
        num = i+8
        cell_ix = "A%d" % num
        cell_time = "B%d" % num
        cell_im = "C%d" % num
        sheet[cell_im] = new_val
        sheet[cell_ix] = cont
        sheet[cell_time] = ':'
        cont += 1
        print(i)
        if (i >= 31):
            workbook.save(filename=os.path.join('static', "SMout.xlsx"))
            cria_pagina(event_data, page, cont, tipo)
            return
        elif(i == len(event_data)-1):
            if((i+1) % 32 != 0):
                num += 1
                cell_ix = "A%d" % num
                cell_time = "B%d" % num
                cell_im = "C%d" % num
                sheet[cell_im] = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
                sheet[cell_ix] = "XXX"
                sheet[cell_time] = "XXX"
            workbook.save(filename=os.path.join('static', "SMout.xlsx"))
            return


""" def grava_conteudo(event_data):
    print(event_data[0])
    page = 1
    # load excel file
    workbook = load_workbook(filename=os.path.join('static', "SM.xlsx"))
    print(workbook.sheetnames)
    # open workbook
    sheet = workbook.active
    cont = 1
    # modify the desired cell
    for i, new_val in enumerate(event_data):
        num = i+8
        cell_ix = "A%d" % num
        cell_time = "B%d" % num
        cell_im = "C%d" % num
        sheet[cell_im] = new_val
        sheet[cell_ix] = cont
        sheet[cell_time] = ':'
        cont += 1
        print(i)
        if (i >= 31):
            workbook.save(filename=os.path.join('static', "SMout.xlsx"))
            cria_pagina(event_data, page, cont)
            print('saiu')
            return
        elif(i == len(event_data)-1):
            print(i)
            if((i+1) % 32 != 0):
                num += 1
                cell_ix = "A%d" % num
                cell_time = "B%d" % num
                cell_im = "C%d" % num
                sheet[cell_im] = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
                sheet[cell_ix] = "XXX"
                sheet[cell_time] = "XXX"
            workbook.save(filename=os.path.join('static', "SMout.xlsx"))
            return
    # save the file
    workbook.save(filename=os.path.join('static', "SMout.xlsx"))
 """


@app.route("/", methods=['POST'])
def index():
    event_data = request.json
    # cria e salva um novo arquivo Excel
    workbook = load_workbook(filename=os.path.join('static', "SM.xlsx"))
    workbook.save(filename=os.path.join('static', "SMout.xlsx"))

    cria_pagina(event_data['seqIsolar'], 0, 1, "ISOLAR-")
    print(event_data['seqNormalizar'])
    cria_pagina(event_data['seqNormalizar'], 0, 1, "NORMALIZAR-")

    file_name = 'SMout.xlsx'
    return app.send_static_file(filename=file_name)


if os.path.exists(os.path.join('static', "SMout.xlsx")):
    os.remove(os.path.join('static', "SMout.xlsx"))
else:
    print("Extra file doesn't exists")


@ app.after_request
def set_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Headers"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "*"
    return response


if __name__ == "__main__":
    app.run(host='127.0.0.1', port=5006)
