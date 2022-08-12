from typing import final
from data.File_Pulls import file_grab
import pandas as pd
import numpy as np
from datetime import date, timedelta
from xlsxwriter.utility import xl_col_to_name
import xlsxwriter

today = date.today()
last_month_last_day = today.replace(day=1) - timedelta(days=1)
last_month = last_month_last_day.strftime("%B")
current_month = today.strftime("%B")
year = today.strftime("%Y")
print(year)
date_save = year + " - " + last_month + " - " + current_month

file_save_end = {
    "0012": " - HH Fayetteville Trending File.xlsx",
    "0014": " - HH Raleigh Trending File.xlsx",
    "0015": " - HH Myrtle Beach Trending File.xlsx",
    "0016": " - HH Wilmington Trending File.xlsx",
    "0018": " - HH Charlotte Trending File.xlsx",
    "0026": " - DFH Colorado Trending File.xlsx",
    "0035": " - DFH Jacksonville Trending File.xlsx",
    "0042": " - DFH Georgia Trending File.xlsx",
    "0067": " - DFH Orlando Trending File.xlsx",
    "0070": " - Capitol Division Trending File.xlsx",
    "0089": " - DFH Texas Trending File.xlsx",
    "0091": " - COV Houston Trending File.xlsx",
    "0092": " - COV Dallas Trending File.xlsx",
    "0093": " - COV Austin Trending File.xlsx",
    "0094": " - COV San Antonio Trending File.xlsx",
    "2001": " - VPH Standard Trending File.xlsx",
    "2002": " - VPH Signature Trending File.xlsx",
    "2003": " - Active Adult Trending File.xlsx",
}


def trending_file(current_date, previous_date, division):
    df = file_grab(division, current_date, previous_date)
    print("Starting trending macro")

    df.drop(
        columns=[
            "Oper Unit_x",
            "Oper Unit_y",
            "Project Name_y",
            "elev",
            "SqFt",
            "Category Code",
            "Product Category",
            "Subcategory Code",
            "Product Subcategory",
            "Product Code",
            "Kit Product Code",
            "Product Description",
            "CutOff Task",
            "Product UofM",
            "Sales Office Opt",
            "Prod Type",
            "Disallow Global Markup",
            "Craft Code",
            "Craft Description",
            "Task Code",
            "Task Description",
            "Supplier Code",
            "Supplier Name",
            "UofM",
            "Unit Price",
            "Draw",
            "Draw %",
            "Total Tax",
            "Total Tax Included",
            "slprodprice_recid",
            "ID",
            "Project_Status",
            "Rec Type",
            "Sub / Part / Bid Rate",
            "Sub / Part / Bid Description",
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
        ],
        inplace=True,
        axis=1,
    )

    df.rename(columns={"Project Name_x": "Project Name"})
    df = df[df["Major Code"].between(2165, 3995)]

    df_projects = []
    ##### Dev writer ######
    # writer = pd.ExcelWriter(
    #     "C:/Users/cmyers/Documents/0014 - H&H Raleigh/dummytest.xlsx"
    # )

    ##### Production writer ####
    writer = pd.ExcelWriter(
        "S:/National Purchasing/Clayborn_Working_Folder/Coding_Testing/"
        + date_save
        + file_save_end[division],
        engine="xlsxwriter",
    )

    workbook = xlsxwriter.Workbook(date_save + file_save_end[division])
    format1 = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
    format2 = workbook.add_format({"bg_color": "C6EFCE", "font_color": "#006100"})

    df.set_index(keys=["Project Code"], drop=False, inplace=True)
    project_codes = df["Project Code"].unique().tolist()

    for i, j in enumerate(project_codes, 1):

        locals()[j] = df.loc[df["Project Code"] == str(j)]
        df_projects.insert(i, locals()[j])
    for k, df_final in enumerate(df_projects):
        name = df_final["Project Code"].iat[1]
        pivot_df = pd.pivot_table(
            df_final,
            index=[
                "Major Code",
                "Major Description",
                "Project Name_x",
                "Model Code",
                "Model Description",
            ],
            values=[
                "Previous Amount",
                "Current Amount",
            ],
            aggfunc=np.sum,
        )
        pivot_df["Total Difference"] = (
            pivot_df["Current Amount"] - pivot_df["Previous Amount"]
        )

        ########  BELOW THIS IS WHAT COULD BE USED FOR CONDITIONAL FORMATTING ##################
        ######## MIGHT NEED TO PUSH IT TO A WORKSHEET AND THEN PULL IT BACK TO A DF?????
        
        # target_col = xl_col_to_name(pivot_df.columns.get_loc("Total Difference"))
        # l = [f"{target_col}2:{target_col}1000"]
        # for columns in l:

        #     pivot_df.conditional_format(
        #         columns,
        #         {"type": "cell", "criteria": "<", "value": 0, "format": format1},
        #     )
        #     pivot_df.conditional_format(
        #         columns,
        #         {"type": "cell", "criteria": ">", "value": 0, "format": format2},
        #     )
        #     return pivot_df
        final_pivot = pd.pivot_table(
            pivot_df,
            index=[
                "Major Code",
                "Major Description",
            ],
            columns=["Project Name_x", "Model Code", "Model Description"],
            values=["Previous Amount", "Current Amount", "Total Difference"],
            aggfunc=np.sum,
        )

        final_pivot.columns = final_pivot.columns.swaplevel(0, 1)
        final_pivot.columns = final_pivot.columns.swaplevel(1, 2)
        final_pivot.columns = final_pivot.columns.swaplevel(2, 3)
        final_df = final_pivot.sort_index(axis=1, sort_remaining=False)
        final_df.to_excel(writer, sheet_name=str(name), engine="xlswriter", index=True)
    writer.save()
