import os, zipfile, csv
from flask import Flask, flash, jsonify, request, redirect, render_template, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)

@app.route('/', defaults={'req_path': ''})
@app.route('/<path:req_path>/')

def dir_listing(req_path):
    BASE_DIR = '/Users/nick.hart/desktop/'

    # Joining the base and the requested path
    abs_path = os.path.join(BASE_DIR, req_path)

    # Show directory contents
    files = os.listdir(abs_path)
    return render_template('a.html', files=files)

if __name__ == '__main__':
    app.run(debug=True)
