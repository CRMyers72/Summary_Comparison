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
    "0012": "0012 - Fayetteville/Backup Files/",
    "0014": "0014 - Raleigh/Backup Files/",
    "0015": "0015 - Myrtle Beach/Backup Files/",
    "0016": "0016 - Wilmington/Backup Files/",
    "0018": "0018 - Charlotte/Backup Files/",
    "0026": "0026 - Colorado/Backup Files/",
    "0035": "0035 - Jacksonville/Backup Files/",
    "0042": "0042 - Georgia/Backup Files/",
    "0067": "0067 - Orlando/Backup Files/",
    "0070": "0070 - Capitol Division/Backup Files/",
    "0089": "0089 - DFH Austin/Backup Files/",
    "0091": "0091 - COV Houston/Backup Files/",
    "0092": "0092 - COV Dallas/Backup Files/",
    "0093": "0093 - COV Austin/Backup Files/",
    "0094": "0094 - COV San Antonio/Backup Files/",
    "2001": "2001 - Bluffton Standard/Backup Files/",
    "2002": "2002 - Bluffton Signature/Backup Files/",
    "2003": "2003 - Active Adult/Backup Files/",
}


division_backup_file_end = {
    "0012": " Takeoff Margin Fayetteville.xlsx",
    "0014": " Takeoff Margin Raleigh.xlsx",
    "0015": " Takeoff Margin Myrtle Beach.xlsx",
    "0016": " Takeoff Margin Wilmington.xlsx",
    "0018": " Takeoff Margin Charlotte.xlsx",
    "0026": " Takeoff Margin Colorado.xlsx",
    "0035": " Takeoff Margin Jacksonville.xlsx",
    "0042": " Takeoff Margin Georgia.xlsx",
    "0067": " Takeoff Margin Orlando.xlsx",
    "0070": " Takeoff Margin Capitol.xlsx",
    "0089": " Takeoff Margin DFH Austin.xlsx",
    "0091": " Takeoff Margin COV Houston.xlsx",
    "0092": " Takeoff Margin COV Dallas.xlsx",
    "0093": " Takeoff Margin COV Austin.xlsx",
    "0094": " Takeoff Margin COV San Antonio.xlsx",
    "2001": " Takeoff Margin Bluffton Standard.xlsx",
    "2002": " Takeoff Margin Bluffton Signature.xlsx",
    "2003": " Takeoff Margin Active Adult.xlsx",
}


def file_grab(division, current_date, previous_date):
    writer = pd.ExcelWriter(
        "S:/National Purchasing/Clayborn_Working_Folder/Coding_Testing/"
        + division
        + "combined.xlsx"
    )

    file_df = pd.DataFrame()
    print(division)
    print(current_date)
    print(previous_date)
    print(date_save)

    first_file = pd.read_excel(
        folder_header
        + division_backup_folder[division]
        + current_date
        + division_backup_file_end[division],
    )

    second_file = pd.read_excel(
        folder_header
        + division_backup_folder[division]
        + previous_date
        + division_backup_file_end[division],
    )

    first_file.rename(columns={"Amount": "Current Amount"}, inplace=True)
    first_file.insert(52, "Previous Amount", NaN)
    # first_file = first_file.drop(columns=["slprodprice_recid", "Refresh Date"])
    second_file.rename(columns={"Amount": "Previous Amount"}, inplace=True)
    second_file.insert(53, "Current Amount", NaN)
    # second_file = second_file.drop(columns=["slprodprice_recid", "Refresh Date"])
    df = project_cleanup(file1=first_file)
    merged2 = project_cleanup(file1=second_file)
    # df = df.append(
    #     merged2,
    #     ignore_index=True,
    # )
    # print(first_file)
    print(df)
    all_dfs = [df, merged2]
    for df in all_dfs:
        df.columns = [
            "Oper Unit",
            "Project Code",
            "Project Name",
            "Model Code",
            "Model Description",
            "elev",
            "SqFt",
            "Category Code",
            "Product Category",
            "Subcategory Code",
            "Product Subcategory Code",
            "Product Code",
            "Kit Product Code",
            "Kit Qty",
            "Product Description",
            "Product Sales Description",
            "Override Descr",
            "CutOff Task",
            "Product UofM",
            "Sales Office Opt",
            "Design Center Opt",
            "Prod Type",
            "Disallow Global Markup",
            "Select Num of Std",
            "Major Code Rev",
            "Minor Code Rev",
            "Disc",
            "Selling Price Eff Date",
            "Selling Price",
            "Total Cost - Tax Out",
            "Margin $ - Tax Out",
            "Margin % - Tax Out",
            "Total Cost",
            "Margin $",
            "Margin %",
            "Craft Code",
            "Craft Description",
            "Task Code",
            "Task Description",
            "Major Code",
            "Major Description",
            "Minor Code",
            "Supplier Code",
            "Supplier Name",
            "Rec Type",
            "Sub / Part / Bid Rate",
            "Sub / Part / Bid Description",
            "UofM",
            "Qty",
            "Unit Price",
            "Draw",
            "Draw %",
            "Previous Amount",
            "Current Amount",
            "Total Tax",
            "Total Tax Included",
            "Oper Unit 2",
            "Oper Unit Name 2",
            "Project Code 2",
            "No. of Lots",
            "Count of Constr Start Date",
            "Count of C.O. Date",
            "Lots Remaining to Start",
            "Lots Remaining to Close",
            "Status",
            "Include Y/N",
        ]
    # df = first_file.append(second_file, ignore_index=True)
    df = pd.concat(all_dfs, ignore_index=True)
    df.drop(
        columns=[
            # "slprodprice_recid",
            # "Refresh Date",
            "Oper Unit 2",
            "Oper Unit Name 2",
            "Project Code 2",
            "No. of Lots",
            "Count of Constr Start Date",
            "Count of C.O. Date",
            "Lots Remaining to Start",
            "Lots Remaining to Close",
            "Status",
        ],
        axis=1,
    )
    file_df = df
    file_df.to_excel(writer, sheet_name="Sheet1", engine="xlsxwriter", index=False)
    writer.close()
    return file_df
