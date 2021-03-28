#! python3
# weekly_releases.py - generates a list of weekly releases for me to set up.
from products import NewProduct

weekly_list = NewProduct()
weekly_list.date = str(input('Enter Date:'))
weekly_list.open_workbook()
weekly_list.weekly_release_generator()
weekly_list.close_workbook()

