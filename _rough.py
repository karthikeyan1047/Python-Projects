
from datetime import datetime, timedelta
from easygui import *

yr_start = 2024
yr_curr = datetime.today().year
months_ranges = [(1,4), (5,8), (9, 12)]
total_files = 0
def get_date_range(year, start_month, end_month):
    date_from = datetime(year, start_month, 1)
    if end_month == 12:
        date_to = datetime(year, 12, 31)
    else:
        date_to = datetime(year, end_month + 1, 1) - timedelta(days=1)
    return date_from, date_to

def process_date_range(date_from: datetime, date_to: datetime):
    global total_files
    alert_text = enterbox("Days : Error/No-Error", "Error")
    if alert_text and alert_text == 'e':
        if (date_to - date_from).days == 0:
            return
        mid_date = date_from + (date_to - date_from) // 2
        next_start = mid_date + timedelta(days=1)

        process_date_range(date_from, mid_date)
        process_date_range(next_start, date_to)
    else:
        total_files += 1
        print(date_from.date(), date_to.date(), f"{date_from.day}-{date_to.day}")


for year in range(yr_start, yr_curr + 1):
    to_process = months_ranges.copy()
    while to_process:
        sm, em = to_process.pop(0)
        alert_text = enterbox("Month : Error/No-Error", "Error")
        if alert_text and alert_text == 'e':
            if sm == em:
                date_from, date_to = get_date_range(year, sm, em)
                process_date_range(date_from, date_to)           
            else:
                mid = (sm + em) // 2
                to_process.insert(0, (mid + 1, em))
                to_process.insert(0, (sm, mid))
        else:
            total_files += 1
            date_from, date_to = get_date_range(year, sm, em)
            print(date_from.date(), date_to.date())


print(total_files)