from openpyxl import load_workbook
from config import workbook_name,meal_count_sheet_name
from meal_count_template import initialize_meal_count_sheet,create_meal_count_sheet
from meal_routin_template import init_meal_routin,create_meal_routin_sheet
from marketing_template import init_marketing_sheet,create_marketing_sheet

import datetime
import calendar
import sys
import os

def create_all_sheets():
     initialize_meal_count_sheet()
     create_meal_count_sheet()
     init_meal_routin()
     create_meal_routin_sheet()
     init_marketing_sheet()
     create_marketing_sheet()
     


def day_month_year():
     now=datetime.datetime.now()
     month=now.month
     year=now.year
     day=now.day
     date_list=[day,month,year]
     return date_list       

def own_work(hostel_object):
     user_input=input("enter the name of user")
     user_name=str(user_input)
     wb=load_workbook(workbook_name)
     ws=wb[meal_count_sheet_name]
     user_present=False
     for cell in ws["A"]:
          if cell.value == user_name:
               user_present=True
               break
     while user_present:
          print("--------------welcome to your profile---------------")
          user_input=input("""
          0.Meal Statue                
          1.Meal On
          2.Meal off
          3.Guest Meal on
          4.Guest Meal off
          5.Deposit amount
          6.Current Balance
          7.Total Meals
          8.Exit to main menu\n
          """)
          if user_input == "0":
               meal_type=input("""
                               1.Own meal
                               2.Guest meal
               """)
               if meal_type =="1":
                    for boarder in hostel_object.boarders:
                         if boarder.name == user_name:
                              status=boarder.active
                              if status:
                                   print(f"{user_name } your meal is active")
                              else:
                                   print(f"{user_name} your meal is not active")
                              pass
               elif meal_type == "2":
                    for boarder in hostel_object.boarders:
                         if boarder.name == user_name:
                              status= boarder.guest_active
                              if status:
                                   print(f"{user_name} your guest meal is active")
                              else:
                                   print(f"{user_name} your guest meal is not active")
                    pass
          elif user_input == "1":
               hostel_object._meal_on(user_name)
               pass
          elif user_input == "2":
               hostel_object._meal_off(user_name)
               pass
          elif user_input == "3":
               hostel_object.guest_meal_on(user_name)
               pass
          elif user_input == "4":
               pass
          elif user_input == "5":
               deposit_amount=int(input("enter the deposit amount only intiger value"))
               hostel_object.add_deposit_(user_name,deposit_amount)
               pass
          elif user_input == "6":
               boarers_list=hostel_object.boarders
               for object in boarers_list:
                    if object.name == user_name:
                         balance=object.deposit
                         print(f"{user_name} your balance in mess is {balance}")
          elif user_input =="7":
               pass
          elif user_input == "8":
               break



        