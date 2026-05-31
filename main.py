import sys
import os
import datetime

from boarder import Boarder
from hostel import Hostel
from config import workbook_name,meal_count_sheet_name
from openpyxl import Workbook,load_workbook
from useful_utils import own_work,create_all_sheets

print('Welcome to mess management')
now=datetime.datetime.now()
date=now.day
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
    my_hostel=Hostel(workbook_name)
    wb=load_workbook(workbook_name)
    ws=wb[meal_count_sheet_name]
    for cell in ws["A"]:
        if cell.value is None:
            break
        else:
            boarder_obb=Boarder(cell.value)
            row_no=cell.row 
            for cell in ws[1]:
                if cell.value == date:
                    start_col=cell.column
                    break
            boarder_obb.row_number=row_no
            boarder_obb.deposit= 0 if ws[f"C{row_no}"].value is None else ws[f"C{row_no}"].value
            boarder_obb.guest_active= True if (ws.cell(row=row_no,column=start_col+1).value == "on" or ws.cell(row=row_no,column=start_col+1).value == "on") else False
            boarder_obb.active=True if (ws.cell(row=row_no,column=start_col).value == "on" or ws.cell(row=row_no,column=start_col+3).value == "on") else False
            my_hostel.boarders.append(boarder_obb)
            loaded=True
    if loaded:
            del my_hostel.boarders[0:2]
    while True:
            user_responce=input("""
                -----------Menu----------\n
                ------enter the corrosponding serial number -------\n
                0.You
                1.Add new Boarder
                2.add deposit
                3.calculate final month bill
                4.meal on 
                5.exit\n""")
            if user_responce == "0":
                own_work(my_hostel)
            if user_responce == "1":
                name=input("enter the boarder name: ")
                # deposit=int(input("enter the deposit amount: "))
                # row_number=int(input("enter the row number"))
                my_hostel.add_boarder(name=name)
                print(f"{name}: boarder added to to hostel mess")
            elif user_responce == "2":
                boarder_name=input("enter the boarder name")
                boarder_amout=int(input("enter the deposit amount"))
                my_hostel.add_deposit_(boarder_name,boarder_amout)
            elif user_responce == "3":
                my_hostel.calculate_meal_charge()
                my_hostel.print_final_balance()
            elif user_responce == "4":
                n_ame=input("enter the boarder name")
                my_hostel._meal_on(n_ame)
            elif user_responce == "5":
                sys.exit()
            else:
                print("Not a valid user input")
            