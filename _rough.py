from datetime import datetime


sdate = datetime(2025, 7, 8).date()
edate = datetime.today().date()

dd = (edate - sdate).days

print(dd)