#!/usr/bin/env python
# coding: utf-8

# In[2]:

def amazon_affiliates_02_processing_job():
    import pandas as pd
    import zipfile
    import glob
    import os
    import time
    import datetime
    import warnings
    import shutil
    import random
    # !pip install xlsxwriter
    # !pip install openpyxl


    # move to download location
    download_path = "/Downloads"
    processing_path = "/amazon_affiliate_daily_dl"
    backup_path = "/daily_backup"
    pickup_path = "/OneDrive/Apps/"
    os.chdir(download_path)

    # Find most recent "report-data" ZIP file
    raw_file_name = glob.glob("Report-Data*.zip")
    latest_file = max(raw_file_name, key=os.path.getctime)


    # Copy most recent report-data zip to the backup location
    file_to_copy = f'./{latest_file}'
    shutil.copy2(file_to_copy, backup_path)

    # Unzip the doc to the processing location
    with zipfile.ZipFile(latest_file,"r") as zip_ref:
        zip_ref.extractall(processing_path)
        print(f"File has been unzipped to {processing_path} and backed up to the daily_backup sub-folder")

    # Change to processing path location, where unzipped file now lives
    os.chdir(processing_path)

    # Suppress default warning about default styles
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

    # Get date  
    current_date = datetime.datetime.now().date() + datetime.timedelta(days=-1) 

    # Set each tab to its own df (different formats each tab)
    unzipped_file = glob.glob('*-XLSX.xlsx')
    # sheet_name = ['sheet1','sheet2' ,'sheet3'] --included for original loop attempt.

    # Open each tab and read it into a df. Done in a for-loop because of data type of unzipped_file. Can't do it directly.
    for file in unzipped_file:
        df_sheet1 = pd.read_excel(file, sheet_name = 'sheet1', engine = "openpyxl")
        df_sheet2 = pd.read_excel(file, sheet_name = 'sheet2', engine = "openpyxl")
        df_sheet3 = pd.read_excel(file, sheet_name = 'sheet3', engine = "openpyxl")

    # remove first row, reset index
    df_sheet1 = df_sheet1.rename(columns=df_sheet1.iloc[0]).drop(df_sheet1.index[0]).reset_index(drop=True).rename(columns={'Date': 'Date2'})
    df_sheet2 = df_sheet2.rename(columns=df_sheet2.iloc[0]).drop(df_sheet2.index[0]).reset_index(drop=True)
    df_sheet3 = df_sheet3.rename(columns=df_sheet3.iloc[0]).drop(df_sheet3.index[0]).reset_index(drop=True)
    
    # Add the date field and update
    df_sheet1["Date"] = current_date
    df_sheet2["Date"] = current_date
    df_sheet3["Date"] = current_date

    # Change to pickup path location, where the new file will be dropped
    os.chdir(pickup_path)
    
    # set file date and name. Including a random ID - why? If run more than 1x a day, file name conflict means OneDrive won't sync.
    random_int = str(random.randint(10000, 100000))
    excel_file_date = current_date.strftime("%Y_%m_%d")
    excel_file_name = "Amazon_Fee_Orders_"+excel_file_date+"_"+random_int+".xlsx"
    
    # write files to excel, deposit in pickup path, where Power Automate/OneDrive will retrieve.
    with pd.ExcelWriter(excel_file_name) as writer:
        df_sheet3.to_excel(writer, sheet_name="Bounty", index=False)
        df_sheet1.to_excel(writer, sheet_name="Fee-Orders", index=False)
        df_sheet2.to_excel(writer, sheet_name="Fee-Earnings", index=False)
    
    
    # In[14]:
    
    
    # Check if new output file has been dropped. If so, delete the processing file and the original ZIP download.
    if os.path.isfile(excel_file_name):
        for name in zip_ref.namelist():
            os.chdir(processing_path)
            os.remove(name)
        os.chdir(download_path)
        os.remove(latest_file)
        print(f"Job complete. Processing file: {file} has been deleted.")
    else:
        print("Delete failed two options: 1) Output file wasn't copied down properly or 2) Raw file was already deleted. Investigate.")

if __name__ == "__main__":
    amazon_affiliates_02_processing_job()
