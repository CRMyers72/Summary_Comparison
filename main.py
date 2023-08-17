import pandas as pd
from file_creators.Summary_file import summary_file
from file_creators.Trending_file import trending_file


# current_date = input("What is the first date? Format YYYYMMDD")
# previous_date = input("What is the previous date? Format YYYYMMDD")
# report_choice = input("What report are you running? Trending or Summary")

# In place only to get the trending script running - delete later
current_date = "202308"
previous_date = "202307"
report_choice = "trending"
df = pd.DataFrame({})
division_list = [
    "0012",
    "0014",
    "0015",
    "0016",
    "0018",
    "0026",
    "0035",
    "0042",
    "0067",
    "0070",
    "0089",
    "0091",
    "0092",
    "0093",
    "0094",
    "2001",
    "2002",
    "2003",
]


def get_summaries(current_date, previous_date, report_choice):
    if current_date == previous_date:
        print("Your dates are the same. Try again!")
    else:
        for division in division_list:
            # file_grab(division, current_date, previous_date)
            if report_choice == "Trending" or report_choice == "trending":
                trending_file(current_date, previous_date, division)
            elif report_choice == "Summary" or report_choice == "summary":
                summary_file
            else:
                print("that file choice is not valid")


get_summaries(current_date, previous_date, report_choice)
