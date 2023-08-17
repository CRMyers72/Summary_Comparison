from heapq import merge
import pandas as pd


def project_cleanup(file1):
    project_file = pd.read_excel(
        # Dev
        # r"C:/Users/cmyers/Documents/Division - Project Split Macro.xlsx"
        # Production
        r"S:/National Purchasing/Clayborn_Working_Folder/Useful_Excel_Docs/Project Tracker.xlsx"
    )
    project_df = pd.DataFrame(project_file)
    merged = pd.merge(file1, project_df, how="left", on="Project Code")
    # merged = merged[merged.Project_Status == "Active"]
    return merged


###### Need to update blanks in project status column to unknown ####
