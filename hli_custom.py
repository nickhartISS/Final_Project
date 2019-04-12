import os, zipfile, csv, datetime, pandas as pd, numpy as np
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

# Deleting individual files
@app.route("/deletefiles/", methods=['GET', 'POST'])
def hli_api_delete_files():
    files = os.listdir(excel_ott_dir)
    for file in files:
        if request.method == 'POST' and (file.endswith(xlsx_extension) or file.endswith(txt_extension)):
            os.remove(os.path.join(excel_ott_dir, file))
            return redirect(url_for('hli_api_delete_files'))
    return render_template('deletefile.html', files=files, xlsx_extension=xlsx_extension,
        txt_extension=txt_extension)

# Generating final reports
@app.route("/finalfiles/", methods=['GET', 'POST'])
def hli_api_final_files():
    files = os.listdir(excel_ott_dir)
    report_date = datetime.datetime.now()
    df_aggregated_report = pd.DataFrame()

    if request.method == 'POST':

        for file in files:

            if file.endswith(xlsx_extension) and 'HLI-LMS-Q' in file:
                df_quarter = pd.DataFrame()

                # Reads all the HLI ott excel files
                data_xlsx = pd.read_excel(os.path.join(excel_ott_dir, file), na_values=[''], skiprows=3, converters={"Zip Code": str}, skip_blank_lines=False, usecols = "A:I")
                df_quarter = df_quarter.append(data_xlsx)
                df_quarter.insert(loc=0, column='Quarter', value=file[11:15] + ' ' + file[8:10], allow_duplicates=True)

                for file in files:

                    if file.endswith(txt_extension) and 'HLI-LMS-Territory Groupings' in file:
                        df_groupings = pd.DataFrame()

                        # Reads the supplement territory_groupings file
                        data_txt = pd.read_csv(os.path.join(excel_ott_dir, file), delimiter="\t")
                        df_groupings = df_groupings.append(data_txt)
                        df_combine_ott_groupings = df_quarter.merge(df_groupings, left_on='Territory', right_on='TERRITORY_ID', how='left', sort=True)
                        df_combine_ott_groupings.insert(loc=7, column='Division', value=df_combine_ott_groupings['Group: Divisional Managers'], allow_duplicates=True)
                        df_combine_ott_groupings.Division.replace(np.NaN, 'Unassigned', inplace=True)
                        df_aggregated_report = df_aggregated_report.append(df_combine_ott_groupings)

        # Final main dataframe that has aggregated report information
        df_aggregated_report.drop(df_aggregated_report.columns[11:20], axis=1, inplace=True)

        # RIA only report that filters for where Division == Ed Cisowski
        df = df_aggregated_report.query('Division == "Ed Cisowski"')
        df.to_excel(os.path.join(final_reports_dir, 'HLI-LMS-' + report_date.strftime("%m%d%y") + ' 3 Year Trend Analysis (Cisowski - RIA).xlsx'), index=False)

        # Unassigned only report that filters for where Division == Unassigned
        df = df_aggregated_report.query('Division == "Unassigned"')
        df.to_excel(os.path.join(final_reports_dir, 'HLI-LMS-' + report_date.strftime("%m%d%y") + ' 3 Year Trend Analysis (Unassigned).xlsx'), index=False)

        # Retail only report that filters for where Division != Ed Cisowski & Division != Unassigned
        df = df_aggregated_report.query('Division != "Ed Cisowski" & Division != "Unassigned"')
        df.to_excel(os.path.join(final_reports_dir, 'HLI-LMS-' + report_date.strftime("%m%d%y") + ' 3 Year Trend Analysis (Retail).xlsx'), index=False)

    return render_template('finalreport.html')

if __name__ == '__main__':
    app.run(debug=True)
