#!/usr/bin/env python3
from flask import Flask, render_template, request
from api.logstat_report_api import api_blueprint
import requests
import os

app = Flask(__name__)

app.register_blueprint(api_blueprint, url_prefix='/api')

API_URL = 'http://localhost:5000/api/generate-logstat'  # Replace with your API URL


@app.route('/generate_form', methods=["POST"])
def generate_form():
    unit_name = request.form['unit_name']
    am_pm = request.form['am_pm']
    day = request.form['day']
    report_type = request.form['report_type']
    file_1 = ""
    file_2 = ""
    output_name = "logstat-{}-day{}-{}.xlsx".format(unit_name,day,am_pm)
    for uploaded_file in request.files.getlist("files"):
        if "EQPT" in uploaded_file.filename:
            file_2 = uploaded_file
        elif "SUPP" in uploaded_file.filename:
            file_1 = uploaded_file
        else:
            return "{} not a valid file".format(uploaded_file.filename)
            
    # Prepare the payload for the API request
    payload = {
        'unit': unit_name,
        'output_filename': output_name
    }
    
    # Add the file data to the payload
    payload['supp_file'] = (file_1.filename, file_1.stream, file_1.content_type)
    payload['equip_file'] = (file_2.filename, file_2.stream, file_2.content_type)

    # Send a POST request to the API
    response = requests.post(API_URL, data=payload, headers={'Content-Type': 'multipart/form-data'})
    
    if response.status_code == 200:
        return response.json()['save_path']
    else:
        return "Failed to generate the report."

@app.route('/')
def main():
    my_obj = { "header" : "Welcome!!!" }
    return render_template("page_template.html", data = my_obj)

if __name__ == '__main__':
    app.run()