#!/usr/bin/env python3
from flask import Flask, render_template, request, Response
from api.logstat_report_api import api_blueprint as logstat_blueprint
from api.perstat_report_api import api_blueprint as perstat_blueprint
import requests
import os
import shutil
import io

app = Flask(__name__)

app.register_blueprint(logstat_blueprint, url_prefix='/api')
app.register_blueprint(perstat_blueprint, url_prefix='/papi')


@app.route('/generate_form', methods=["POST"])
def generate_form():
    unit_name = request.form['unit_name']
    am_pm = request.form['am_pm']
    day = request.form['day']
    report_type = request.form['report_type']
    file_1 = ""
    file_2 = ""
    output_name = ""
 
    files = {} 
    if report_type == "logstat":
        output_name = "logstat-{}-day{}-{}.xlsx".format(unit_name,day,am_pm)
        API_URL = 'http://localhost:5000/api/generate-logstat'  # Replace with your API URL
        for uploaded_file in request.files.getlist("files"):
            if any(keyword in uploaded_file.filename for keyword in ["EQUIP", "EQPT"]):
                print("file_2 is : {}".format(uploaded_file.filename))
                file_2 = uploaded_file
            elif any(keyword in uploaded_file.filename for keyword in ["SUPP", "SUP"]):
                file_1 = uploaded_file
                print("file_1 is : {}".format(uploaded_file.filename))
            else:
                return "{} not a valid file".format(uploaded_file.filename)        
        # Add the file data to the payload
        files['supp_file'] = file_1
        files['equip_file'] = file_2
    else:
        output_name = "perstat-{}-day{}-{}.xlsx".format(unit_name,day,am_pm)
        API_URL = 'http://localhost:5000/papi/generate-perstat'  # Replace with your API URL
        files['perstat_file'] = next(iter(request.files.values()))

    # Prepare the payload for the API request
    payload = {
        'unit': unit_name,
        'output_filename': output_name,
        'am_pm': am_pm
    }

    # Send a POST request to the API
    response = requests.post(API_URL, data=payload, files=files)
    
    if response.status_code == 200:
        # Create an in-memory file buffer
        file_buffer = io.BytesIO(response.content)

        # Set the appropriate headers for the file download
        headers = {
            'Content-Disposition': f'attachment; filename="{output_name}"',
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }

        # Create a Flask Response object with the file buffer and headers
        return Response(file_buffer, headers=headers)
    else:
        return "Failed to generate the report."

@app.route('/')
def main():
    my_obj = { "header" : "87th Training Division Report Generator" }
    return render_template("page_template.html", data = my_obj)

if __name__ == '__main__':
    app.run(host='0.0.0.0')