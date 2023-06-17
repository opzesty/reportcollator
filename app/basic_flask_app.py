#!/usr/bin/env python3
from flask import Flask, render_template, request

app = Flask(__name__)

@app.route('/generate_form', methods=["POST"])
def generate_form():
    unit_name = request.form['unit_name']
    am_pm = request.form['am_pm']
    report_type = request.form['report_type']
    filenames = ""
    print(len(request.files))
    for uploaded_file in request.files.getlist("files"):
        filenames += uploaded_file.filename + " "
    return "{} - {} - {} - {}".format(unit_name, am_pm, report_type, filenames)

@app.route('/')
def main():
    return render_template("page_template.html")

if __name__ == '__main__':
    app.run()
