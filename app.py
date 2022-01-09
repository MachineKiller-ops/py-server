from flask import Flask, send_from_directory

app = Flask(__name__, static_url_path='')


@app.route("/")
def index():

    file_name = 'SM.xlsx'
    return app.send_static_file(filename=file_name)
