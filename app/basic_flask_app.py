#!/usr/bin/env python3
from flask import Flask, render_template

app = Flask(__name__)

@app.route('/generate_form')
def generate_form():
    #todo logic
    return 'Hello, World!'

@app.route('/')
def main():
    return render_template("page_template.html")

if __name__ == '__main__':
    app.run()
