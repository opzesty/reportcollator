from flask import Blueprint, request, jsonify, Response
from openpyxl import Workbook
from openpyxl.styles import Font
import xlrd
import sys
import os
import io
from datetime import date

api_blueprint = Blueprint('api', __name__)

@api_blueprint.route('/generate-perstat', methods=['POST'])
def generate_perstat():
    # Get the request data
    unit = request.form['unit']
    am_pm = request.form['am_pm']
    output_filename = request.form['output_filename']
    perstat_file = request.files['perstat_file']

    # Perform the report generation logic
    time = set_time(am_pm)
    date = set_date()
    zulutime = combine_datetime(time, date)
    save_path = get_save_path(output_filename)
    workbook = initialize_workbook()
    workbook = format_report(workbook)
    #perstat_info = pull_perstat_info(perstat_file)
    perstat_info = "test"
    workbook = populate_report(workbook, perstat_info, unit)

    # Prepare the response
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    # Prepare the response headers
    headers = {
        'Content-Disposition': f'attachment; filename="{output_filename}"',
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }

    return Response(output, headers=headers)

def get_save_path(filename):
    current_directory = os.getcwd()
    save_path = os.path.join(current_directory, filename)
    return save_path

def check_and_add_unit(sheet, search_unit):
    column_a = sheet["A"]

    # Check if the search string is already present in column A
    for row in range(2, len(column_a) + 1):
        if column_a[row - 1].value == search_unit:
            return row

    # If the search string is not found, add it to the end of column A
    new_row = len(column_a) + 1
    sheet.cell(row=new_row, column=1).value = search_unit

    return new_row
    
def set_time(am_pm):
    if am_pm == "am":
        return "0700"
    elif am_pm == "pm":
        return "1300"
    else:
        raise ValueError("Invalid value for am_pm. Expected 'am' or 'pm'.")
    
def set_date():
    current_date = date.today()
    return current_date
    
def combine_datetime(time, current_date):
    formatted_date = current_date.strftime("%d").upper()
    formatted_time = time[:2] + time[2:]
    formatted_month_year = current_date.strftime("%b%y").upper()
    combined_datetime = f"{formatted_date}{formatted_time}Z{formatted_month_year}"
    return combined_datetime

def initialize_workbook():
    # Create a new workbook
    workbook = Workbook()
    
    return workbook
    
def save_workbook(workbook, save_path):
    workbook.save(save_path)
    
    return(save_path)
    
def pull_perstat_info(perstat_file):
    raise NotImplementedError("Function not yet implemented")

def format_report(workbook):
    # Headers for perstat report
    headers = [
    "Date/Time (DD01300ZJUN23)",
    "LOCATION",
    "UNIT",
    "AUTHORIZED/RANK",
    "ASSIGNED/RANK",
    "BRANCH/MOS",
    "ON HAND",
    "GAINS",
    "REPLACEMENTS",
    "RETURNED TO DUTY",
    "KILLED",
    "WOUNDED",
    "NON-BATTLE LOSS",
    "MISSING",
    "DESERTERS",
    "AWOL",
    "CAPTURED",
    "AUTHENTICATION"
    ]


    # Remove the default first sheet
    default_sheet = workbook.active
    column= 1
    
    for header in headers:
        unit_cell = default_sheet.cell(row=1, column=column)
        unit_cell.value = header
        unit_cell.font = Font(bold=True, underline="single")
        column+= 1

    return workbook

def populate_report(workbook, data, unit):
    # not yet implemented 
    return workbook