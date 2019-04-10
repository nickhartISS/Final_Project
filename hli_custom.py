import os, zipfile, csv
from flask import Flask, flash, jsonify, request, redirect, render_template, url_for
from werkzeug.utils import secure_filename

# Directory paths
zip_upload_dir = "/Users/nick.hart/desktop/test/HLI_Zip_OTTs/"
excel_ott_dir = "/Users/nick.hart/desktop/test/"
final_reports_dir = "/Users/nick.hart/desktop/test/Final_Reports/"

# Approved file extensions
zip_extension = ".zip"
xlsx_extension = ".xlsx"
txt_extension = ".txt"

app = Flask(__name__)
app.secret_key = "super secret key"

# Home route - will show the list of files in the test directory
@app.route("/")
@app.route('/home/')
def home():
    # Return error if path doesn't exist
    if not os.path.exists(excel_ott_dir):
        dir_error = 'The directory does not exist.'
        return dir_error
    # Show directory contents
    files = os.listdir(excel_ott_dir)
    return render_template('home.html', files=files, xlsx_extension=xlsx_extension,
        txt_extension=txt_extension)

# Uploading .zip OTTs to a folder, extract OTTs, and delete the .zip files
@app.route("/addfiles/", methods=['GET', 'POST'])
def hli_api_add_files():
    files = os.listdir(zip_upload_dir)
    if request.method == 'POST':
        for file in request.files.getlist('file'):
            filename = secure_filename(file.filename)
            if filename.endswith(zip_extension):
                file.save(os.path.join(zip_upload_dir, filename))
                zip_ref = zipfile.ZipFile(os.path.join(zip_upload_dir, filename), 'r')
                zip_ref.extractall(excel_ott_dir)
                zip_ref.close()
                os.remove(os.path.join(zip_upload_dir, filename))
            else:
                filename = secure_filename(file.filename)
                file.save(os.path.join(excel_ott_dir, filename))
                os.remove(os.path.join(zip_upload_dir, filename))
        return redirect(url_for('home'))
    return render_template('addfile.html', files=files, zip_extension=zip_extension)

@app.route("/deletefiles/", methods=['GET', 'POST'])
def hli_api_delete_files():
    files = os.listdir(excel_ott_dir)
    for file in files:
        if request.method == 'POST' and (file.endswith(xlsx_extension) or file.endswith(txt_extension)):
            os.remove(os.path.join(excel_ott_dir, file))
            return redirect(url_for('hli_api_delete_files'))
    return render_template('deletefile.html', files=files, xlsx_extension=xlsx_extension,
        txt_extension=txt_extension)

if __name__ == '__main__':
    app.run(debug=True)
