from numpy import NaN
import pandas as pd
import warnings
from datetime import date
from data.Division_Projects import project_cleanup
from data.Division_Name import division_names

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

today = date.today()
date_save = today.strftime("%Y-%m-%d")
first_file = {}
second_file = {}

# use this for dev work
# folder_header = r"C:/Users/cmyers/Documents/"

# use this for productions
folder_header = r"S:/National Purchasing/Month over Month Trending/"
division_backup_folder = {
    "0012": "0012 - H&H Fayetteville/Backup Files/",
    "0014": "0014 - H&H Raleigh/Backup Files/",
    "0015": "0015 - H&H Myrtle Beach/Backup Files/",
    "0016": "0016 - H&H Wilmington/Backup Files/",
    "0018": "0018 - H&H Charlotte/Backup Files/",
    "0026": "0026 - Colorado/Backup Files/",
    "0035": "0035 - Jacksonville/Backup Files/",
    "0042": "0042 - Georgia/Backup Files/",
    "0067": "0067 - Orlando/Backup Files/",
    "0070": "0070 - Capitol Division/Backup Files/",
    "0089": "0089 - Texas/Backup Files/",
    "0091": "0091 - COV Houston/Backup Files/",
    "0092": "0092 - COV Dallas/Backup Files/",
    "0093": "0093 - COV Austin/Backup Files/",
    "0094": "0094 - COV San Antonio/Backup Files/",
    "2001": "2001 - Village Park Standard/Backup Files/",
    "2002": "2002 - Village Park Signature/Backup Files/",
    "2003": "2003 - Active Adult/Backup Files/",
}


division_backup_file_end = {
    "0012": " Takeoff Margin HH Fayetteville.xlsx",
    "0014": " Takeoff Margin HH Raleigh.xlsx",
    "0015": " Takeoff Margin HH Myrtle Beach.xlsx",
    "0016": " Takeoff Margin HH Wilmington.xlsx",
    "0018": " Takeoff Margin HH Charlotte.xlsx",
    "0026": " Takeoff Margin Colorado.xlsx",
    "0035": " Takeoff Margin Jacksonville.xlsx",
    "0042": " Takeoff Margin Georgia.xlsx",
    "0067": " Takeoff Margin Orlando.xlsx",
    "0070": " Takeoff Margin CAP.xlsx",
    "0089": " Takeoff Margin Texas.xlsx",
    "0091": " Takeoff Margin COV Houston.xlsx",
    "0092": " Takeoff Margin COV Dallas.xlsx",
    "0093": " Takeoff Margin COV Austin.xlsx",
    "0094": " Takeoff Margin COV San Antonio.xlsx",
    "2001": " Takeoff Margin VPH Standard.xlsx",
    "2002": " Takeoff Margin VPH Signature.xlsx",
    "2003": " Takeoff Margin Active Adult.xlsx",
}


def file_grab(division, current_date, previous_date):
    file_df = pd.DataFrame()
    print(division)
    print(current_date)
    print(previous_date)
    print(date_save)

    first_file = pd.read_excel(
        folder_header
        + division_backup_folder[division]
        + current_date
        + division_backup_file_end[division]
    )

    second_file = pd.read_excel(
        folder_header
        + division_backup_folder[division]
        + previous_date
        + division_backup_file_end[division]
    )

    first_file.rename(columns={"Amount": "Current Amount"}, inplace=True)
    first_file.insert(52, "Previous Amount", NaN)
    second_file.rename(columns={"Amount": "Previous Amount"}, inplace=True)
    second_file.insert(53, "Current Amount", NaN)
    df = project_cleanup(file1=first_file)
    merged2 = project_cleanup(file1=second_file)
    df = df.append(merged2, ignore_index=True)
    file_df = df

    return file_df
