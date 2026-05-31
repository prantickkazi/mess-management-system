from boarder import Boarder 
from openpyxl import load_workbook
from config import workbook_name,meal_count_sheet_name
from useful_utils import day_month_year
import datetime


class Hostel:
    def __init__(self,filename):
        self.filename= filename
        self.boarders= []
        self.total_marketing_expenses= 0
        self.current_meal_charge= 0
        
    #add boarder and update the excell sheet    
    def add_boarder(self,name):
        boarder=Boarder(name)
        self.boarders.append(boarder)
        wb=load_workbook(workbook_name)
        ws=wb[meal_count_sheet_name]
        last_cell_no=1
        for cl in ws["A"]:
            if cl.value is not None:
                last_cell_no+=1
            else:
                break    
        ws[f"{last_cell_no}"]=name  
        wb.save(workbook_name)     
        msg=f"Boarder {boarder.name} is successfully added to hostel"
        print(msg)
    
     #add the deposit and updatee the excell sheet   
    def  add_deposit_(self,_name,amount):
        success=False
        target_boarder=None
        for boarder in self.boarders:
            if boarder.name == _name:
                boarder.add_deposit(amount)
                target_boarder=boarder
                success=True
                break
        if success:
            wb=load_workbook(workbook_name)
            ws=wb[meal_count_sheet_name]
            for cell in ws["A"]:
                if cell.value == _name:
                    row_number=cell.row
                    break
            ws[f"C{row_number}"]=target_boarder.deposit
            wb.save(workbook_name)  
            print(f"{_name} your deposit of {amount} is successfull.")
        else:
            print(f"{_name}  is not found as a boarder")
     
     #on the meal and update the excell sheet   
    def _meal_on(self,_name):
        success = False
        info = None
        
        #check is user a boarder
        for boarder in self.boarders:
            if boarder.name == _name:
                info=input("""
                "on" for turn on night meal \n
                "od" for turn on day meal \n
                "yes" for both day and night\n""")
                boarder.on_meal(info)
                success=True
                break
        #update the excell sheet
        if success:
            wb=load_workbook(workbook_name)
            ws=wb[meal_count_sheet_name]
            for cell in ws["A"]:
                if cell.value == _name:
                    row_number=cell.row
                    
            #update the meal status
            ws[f"B{row_number}"]="yes"
            now=datetime.datetime.now()
            day=now.day 
                  
            #upadate the meal status of particular day
            for cell in ws[1]:
                if cell.value == day:
                    start_col=cell.column
            if info == "yes":
                ws.cell(row=row_number,column=start_col).value="on"
                ws.cell(row=row_number,column=start_col+2).value="on"
            elif info =="od":
                ws.cell(row=row_number,column=start_col).value="on"
            elif info =="on":
                ws.cell(row=row_number,column=start_col+2).value="on"
            wb.save(workbook_name)
        else:
            print(f"Name {_name} does not exist")
            
    #off the meal and update the excell sheet
    def _meal_off(self,_name):
        info=None
        success=False
        
        #check in memory if user exist or not
        for boarder in self.boarders:
            if boarder.name == _name:
                info=input("""
                "on" for turn off night meal \n
                "od" for turn off day meal \n
                "off" for both day and night\n""")
                boarder.off_meal(info)
                success=True
        
        if success:       
            #update the excel sheet
            wb=load_workbook(workbook_name)
            ws=wb[meal_count_sheet_name]
            date=day_month_year()[0]
            
            for cell in ws["A"]:
                if cell.value == _name:
                    row_number = cell.row
                    for cell in ws[1]:
                        if cell.value == date:
                            start_col=cell.column
                            if info == "off":
                               ws.cell(row=row_number,column=start_col).value = "off"
                               ws.cell(row=row_number,column=start_col+3).value = "off" 
                            elif info == "od":
                                ws.cell(row=row_number,column=start_col).value = "off"
                            elif info == "on":
                                ws.cell(row=row_number,column=start_col+3).value = "off"
                            break
            wb.save(workbook_name)
        else:
            print(f"{_name} you are not a boarder")
    
    #guest meal on 
    def guest_meal_on(self,_name):
        success=False
        for object in self.boarders:
            if object.name == _name:
                info=input("""
                on guest meal in both day and night -yes
                on guest meal in day-od
                on guest meal at night-on\n
                           """)
                object.on_guest_meal(info)
                success=True
                break
        if success:
            wb=load_workbook(workbook_name)
            ws=wb[meal_count_sheet_name]
            now=datetime.datetime.now()
            day=now.day
            for cell in ws["A"]:
                if cell.value == _name:
                    row_number=cell.row
                    break
            for cell in ws[1]:
                if cell.value == day:
                    get_col=cell.column
                    break
            if info == "od":
                ws.cell(row=row_number,column=get_col+1).value="on"
            elif info == "on":
                ws.cell(row=row_number,column=get_col+3).value="on"
            elif info == "yes":
                ws.cell(row=row_number,column=get_col+1).value="on"
                ws.cell(row=row_number,column=get_col+3).value="on"
            wb.save(workbook_name)
        else:
            print(f"no user exist wuth this name") 
               
    def calculate_meal_charge(self):
        all_meals=0
        for i in self.boarders:
            all_meals+=i.total_meal
        if all_meals == 0:
            print("first day of mess meal charge is 0")
        else:
            meal_charge=self.total_marketing_expenses/all_meals
            self.current_meal_charge=meal_charge
            print(f"Current meal charge is {self.current_meal_charge}")
            
    def print_final_balance(self):
        balance_dict={}
        for i in self.boarders:
            final_balance=(i.total_meal * self.current_meal_charge) - i.deposit
            if final_balance == 0:
                balance_dict[i.name]="all clear"
                
            if final_balance < 0:
                balance_dict[i.name]=f"Return from hostel mess {-1*(final_balance)}"
                
            if final_balance > 0:
                balance_dict[i.name]=f"Due from {i.name} : {final_balance}"
        
        for name,status in balance_dict.items():
            print(f"Name {name} : {status}\n")
            

   
            
            
      
        