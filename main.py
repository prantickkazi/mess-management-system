import sys
import os


from config import workbook_name,meal_count_sheet_name,info_string
from openpyxl import Workbook,load_workbook
from useful_utils import own_work,create_all_sheets,main_menu

print('Welcome to mess management')
created=False
if not os.path.exists(workbook_name):
    print("Workbook not exist ")
    user_res=input("Create a new workbook: y/n")
    if user_res == 'y':
        wb=Workbook(workbook_name)
        wb.save(workbook_name)
        create_all_sheets()
        
        print(f"Successful: workbook name{workbook_name}")
        created=True
    else:
        print("Need workbook to continue")
else:
    created = True
        
if created:
    while True:
        main_menu()