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
    download_path = "/Users/ben.pollock/Downloads"
    processing_path = "/Users/ben.pollock/Documents/amazon_affiliate_daily_dl"
    backup_path = "/Users/ben.pollock/Documents/amazon_affiliate_daily_dl/daily_backup"
    pickup_path = "/Users/ben.pollock/Library/CloudStorage/OneDrive-Hearst/Apps/"
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
    # sheet_name = ['Fee-Orders','Fee-Earnings' ,'Bounty'] --included for original loop attempt. couldn't get it to work.

    # Open each tab and read it into a df. Done in a for-loop because of data type of unzipped_file. Can't do it directly.
    for file in unzipped_file:
        df_fee_orders = pd.read_excel(file, sheet_name = 'Fee-Orders', engine = "openpyxl")
        df_fee_earnings = pd.read_excel(file, sheet_name = 'Fee-Earnings', engine = "openpyxl")
        df_bounty = pd.read_excel(file, sheet_name = 'Bounty', engine = "openpyxl")

    # remove first row, reset index
    df_fee_orders = df_fee_orders.rename(columns=df_fee_orders.iloc[0]).drop(df_fee_orders.index[0]).reset_index(drop=True).rename(columns={'Date': 'Date2'})
    df_fee_earnings = df_fee_earnings.rename(columns=df_fee_earnings.iloc[0]).drop(df_fee_earnings.index[0]).reset_index(drop=True)
    df_bounty = df_bounty.rename(columns=df_bounty.iloc[0]).drop(df_bounty.index[0]).reset_index(drop=True)
    
    # Add the date field and update
    df_fee_orders["Date"] = current_date
    df_fee_earnings["Date"] = current_date
    df_bounty["Date"] = current_date

    # Change to pickup path location, where the new file will be dropped
    os.chdir(pickup_path)
    
    # set file date and name. Including a random ID - why? If run more than 1x a day, file name conflict means OneDrive won't sync.
    random_int = str(random.randint(10000, 100000))
    excel_file_date = current_date.strftime("%Y_%m_%d")
    excel_file_name = "Amazon_Fee_Orders_"+excel_file_date+"_"+random_int+".xlsx"
    
    # write files to excel, deposit in pickup path, where Power Automate/OneDrive will retrieve.
    with pd.ExcelWriter(excel_file_name) as writer:
        df_bounty.to_excel(writer, sheet_name="Bounty", index=False)
        df_fee_orders.to_excel(writer, sheet_name="Fee-Orders", index=False)
        df_fee_earnings.to_excel(writer, sheet_name="Fee-Earnings", index=False)
    
    
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