
# THIS FILE AGGREGATES BPT DATA AND MERGES TOGETHER TO ENABLE FURTHER PROCESSING
# NB it accepts both xlsx and xlsb files as input. But binary files seem to run quicker.

import pandas as pd
import glob
import sqlite3
import os

# Create an empty dataframes
df = pd.DataFrame()
#  headers = set()

path = r"\\fpcl01\public\OFWSHARE\PR24\231002 - Business Plan Submissions\Files as submitted by companies"  # This is the O drive location where we'll save our files
#os.chdir(path)


# Create aggregated f_outputs sheets from companies' files
for file in glob.glob(r".\Inputs\*.xls*"):  # glob is a way of getting all the files of a certain type
    acronym = file[9:12]  # Get the three letter prefix we have added. This assumes that each file has the acronym of the company added at the beginning
    xl = pd.ExcelFile(file)  # Define the excel file
    worksheets = xl.sheet_names  # Get the list of worksheets in the file
    for i in worksheets:  # For each worksheet...
        if "fOut" in i:  # ...if the workseet contains the text "F_Outputs" then...
            df_temp = pd.read_excel(file, sheet_name= i, skiprows=1)  # put the data into a temporary dataframe
            df_temp = df_temp.iloc[1:]  # Chop off the top row (in our file it is unused)
            # headers_temp = set(df_temp.columns)  # In development this and the next row were used to ensure the df would ultimately have all the right headers
            # headers = headers.union(headers_temp)
            df_temp['Acronym'] = acronym  # Add in the acronym name so we we can identify the company the data relates to
            df_temp['Sheet_name'] = i  # Add in the name of the worksheet, so we locate any issues more easily
            print(acronym + " " + i + " completed ok")
            df = pd.concat([df, df_temp])  # Add the temporary df to the master dataframe

df = df[["Acronym", "Reference", "Sheet_name",  "Unit", "Model", "Item description", # Reorder the years columns
         "2011-12", "2012-13", "2013-14", "2014-15", "2015-16", "2016-17", "2017-18", "2018-19", "2019-20", "2020-21", "2021-22", "2022-23", "2023-24", "2024-25", "2025-26", "2026-27", "2027-28", "2028-29", "2029-30", "2030-31", "2031-32", "2032-33", "2033-34", "2034-35",
         "2020-25", "2025-30", "2030-35", "2035-40", "2040-45", "2045-50",
         '2039-40', '2044-45', '2025-55', 'Constant', '2049-50']]

df.reset_index(inplace=True, drop=True)

# Export the results to excel, pickle and sqlite3
df.to_excel(r".\Outputs\All.xlsx", sheet_name="F_Output")
df.to_pickle(r".\Outputs\F_Outputs.pkl")
conn = sqlite3.connect(r'Outputs\All.db')
df.to_sql('F_Outputs', conn, if_exists='replace')




