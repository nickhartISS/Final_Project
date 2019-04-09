# importing required modules: os for file directories, zipfile for zipped files
import os, zipfile, csv
import pandas as pd, numpy as np

# definining dir path and file extension
directory_name = "/Users/nick.hart/desktop/test"

# Looking in directory_name
os.chdir(directory_name)

# File extension variables
zip_extension = ".zip"
xlsx_extension = ".xlsx"
txt_extension = ".txt"

# Reads all the zipped files from dir, unzipes them, then deletes the zipped files
def export_zip_files():

    for file in os.listdir(directory_name): # for all files in the set directory
        if file.endswith(zip_extension): # if the file extension is ".zip"
            file_name = os.path.abspath(file) # get full path of files
            zip_ref = zipfile.ZipFile(file_name) # create zipfile object
            zip_ref.extractall(directory_name) # extract file to dir
            zip_ref.close() # close file
            os.remove(file_name) # delete zipped file

# Read 12 OTT files and appends the needed data to a master excel file
def hli_ott_format():

    df_aggregated_report = pd.DataFrame()

    for file in os.listdir(directory_name):

        if file.endswith(xlsx_extension) and 'HLI-LMS-Q' in file:
            df_quarter = pd.DataFrame()

            # Reads all the HLI ott excel files
            data_xlsx = pd.read_excel(file, na_values=[''], skiprows=3, converters={"Zip Code": str},
                skip_blank_lines=False, usecols = "A:I")
            df_quarter = df_quarter.append(data_xlsx)
            df_quarter.insert(loc=0, column='Quarter', value=file[11:15] + ' ' + file[8:10], allow_duplicates=True)

            # Reads the supplement territory_gropings file
            for file in os.listdir(directory_name):

                if file.endswith(txt_extension) and 'HLI-LMS-Territory Groupings' in file:

                    df_groupings = pd.DataFrame()

                    data_txt = pd.read_csv(file, delimiter="\t")
                    df_groupings = df_groupings.append(data_txt)
                    df_combine_ott_groupings = df_quarter.merge(df_groupings, left_on='Territory',
                        right_on='TERRITORY_ID', how='left', sort=True)
                    df_combine_ott_groupings.insert(loc=7, column='Division',
                        value=df_combine_ott_groupings['Group: Divisional Managers'], allow_duplicates=True)
                    df_combine_ott_groupings.Division.replace(np.NaN, 'Unassigned', inplace=True)
                    df_aggregated_report = df_aggregated_report.append(df_combine_ott_groupings)

    df_aggregated_report.drop(df_aggregated_report.columns[11:20], axis=1, inplace=True)
    return df_aggregated_report


# Compile the ria report
def ria_report_to_excel(export):
    df = export.query('Division == "Ed Cisowski"')
    return df

# Compile the unassigned report
def unassigned_report_to_excel(export):
    df = export.query('Division == "Unassigned"')
    return df

# Compile the ria report
def retail_report_to_excel(export):
    df = export.query('Division != "Ed Cisowski" & Division != "Unassigned"')
    return df

# Calling the file export function
export_zip_files()
hli_ott_format()

# Calling the 3 year ott format function
export = hli_ott_format()

export_ria = ria_report_to_excel(export)
export_ria.to_excel('HLI-LMS-022619 3 Year Trend Analysis (Cisowski - RIA).xlsx', index=False)

export_unassigned = unassigned_report_to_excel(export)
export_unassigned.to_excel('HLI-LMS-022619 3 Year Trend Analysis (Unassigned).xlsx', index=False)

export_retail = retail_report_to_excel(export)
export_retail.to_excel('HLI-LMS-022619 3 Year Trend Analysis (Retail).xlsx', index=False)
