#!/usr/bin/python3
from flask import Flask, request, render_template, send_file

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/static/js/<path:path>")
def dl(path):
    if request.method == "GET":
        return send_file('./jquery.min.js')

app.run(host="0.0.0.0", port=8000)
