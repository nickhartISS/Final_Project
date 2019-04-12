import os, zipfile, csv, datetime, pandas as pd, numpy as np

# --GLOBAL VARIABLES--
# Report Directory
report_directory = "/Users/nick.hart/desktop/HLI_Reports"

# Approved file extensions
zip_extension = ".zip"
xlsx_extension = ".xlsx"
txt_extension = ".txt"

# Looking in report_directory
os.chdir(report_directory)

# Reads all the zipped files from dir, unzipes them, then deletes the zipped files
def export_zip_files():

    for file in os.listdir(report_directory): # for all files in the set directory

        if file.endswith(zip_extension): # if the file extension is ".zip"

            file_name = os.path.abspath(file) # get full path of files
            zip_ref = zipfile.ZipFile(file_name) # create zipfile object
            zip_ref.extractall(report_directory) # extract file to dir
            zip_ref.close() # close file
            os.remove(file_name) # delete zipped file

# Read 12 OTT files and appends the needed data to a master excel file
def hli_ott_format():
    formatted_report_date = datetime.datetime.now().strftime("%m%d%y")
    df_aggregated_report = pd.DataFrame()

    for file in os.listdir(report_directory):

        if file.endswith(xlsx_extension) and 'HLI-LMS-Q' in file:
            df_quarter = pd.DataFrame()

            # Reads all the HLI ott excel files
            data_xlsx = pd.read_excel(file, na_values=[''], skiprows=3, converters={"Zip Code": str}, skip_blank_lines=False, usecols = "A:I")
            df_quarter = df_quarter.append(data_xlsx)
            df_quarter.insert(loc=0, column='Quarter', value=file[11:15] + ' ' + file[8:10], allow_duplicates=True)

            for file in os.listdir(report_directory):

                if file.endswith(txt_extension) and 'HLI-LMS-Territory Groupings' in file:
                    df_groupings = pd.DataFrame()

                    # Reads the supplement territory_groupings file
                    data_txt = pd.read_csv(file, delimiter="\t")
                    df_groupings = df_groupings.append(data_txt)
                    df_combine_ott_groupings = df_quarter.merge(df_groupings, left_on='Territory', right_on='TERRITORY_ID', how='left', sort=True)
                    df_combine_ott_groupings.insert(loc=7, column='Division', value=df_combine_ott_groupings['Group: Divisional Managers'], allow_duplicates=True)
                    df_combine_ott_groupings.Division.replace(np.NaN, 'Unassigned', inplace=True)
                    df_aggregated_report = df_aggregated_report.append(df_combine_ott_groupings)

    # Final main dataframe that has aggregated report information
    df_aggregated_report.drop(df_aggregated_report.columns[11:20], axis=1, inplace=True)

    # RIA only report that filters for where Division == Ed Cisowski
    df = df_aggregated_report.query('Division == "Ed Cisowski"')
    df.to_excel('HLI-LMS-' + formatted_report_date + ' 3 Year Trend Analysis (Cisowski - RIA).xlsx', index=False)

    # Unassigned only report that filters for where Division == Unassigned
    df = df_aggregated_report.query('Division == "Unassigned"')
    df.to_excel('HLI-LMS-' + formatted_report_date + ' 3 Year Trend Analysis (Unassigned).xlsx', index=False)

    # Retail only report that filters for where Division != Ed Cisowski & Division != Unassigned
    df = df_aggregated_report.query('Division != "Ed Cisowski" & Division != "Unassigned"')
    df.to_excel('HLI-LMS-' + formatted_report_date + ' 3 Year Trend Analysis (Retail).xlsx', index=False)

# Calling export_zip_files() method
export_zip_files()

# Calling hli_ott_format() method
hli_ott_format()
