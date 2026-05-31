from openpyxl import load_workbook
from excel_utils import write_in_excel_file
from config import workbook_name,meal_count_sheet_name

class Boarder:
    def __init__(self,name):
        self.name=name
        self.deposit=0
        self.row_number=None
        self.total_meal=0
        self.total_guest_meal=0
        self.active=None
        self.night_active=None
        self.day_active=None
        self.guest_active=None
        self.guest_day_active=None
        self.guest_night_active=None
        
        
    def on_meal(self,info= "yes"):
        if info == "yes":
            self.night_active=True
            self.day_active=True
            self.active=True
            print(f"your both meal is on")
            self.total_meal+=2
        elif info == "on":
            self.night_active=True
            self.day_active=False
            self.active=True
            print("your night meal is only on")
            self.total_meal+=1
        elif info == "od":
            self.night_active=False
            self.day_active=True
            self.active=True
            print("your day meal is only on")
            self.total_meal+=1
           
    def off_meal(self,info):
        if info == "off":
            self.active = False
            self.day_active = False
            self.night_active = False
        elif info == "od":
            self.day_active == False
            self.night_active == True
            self.active=True
            self.total_meal += 1
        elif info == "on":
            self.day_active == False
            self.night_active == True
            self.active=True
            self.total_meal += 1
            
    def on_guest_meal(self,info):
        if info == "yes":
            self.guest_active = True
            self.guest_day_active = True
            self.guest_night_active = True
            self.guest_meal += 2
        if info == "od":
            self.guest_active = True
            self.guest_day_active = True
            self.guest_night_active = False
            self.guest_meal += 1
            pass
        if info == "on":
            self.guest_active = True
            self.guest_day_active = False
            self.guest_night_active = True
            self.guest_meal += 1
            pass
        
    def off_guest_meal(self,info):
        self.guest_active=False
        
    def add_deposit(self,amount):
        self.deposit+=amount
        # write_in_excel_file(meal_count_sheet_name,f"C{self.row_number}",self.deposit)
        msg=f"{self.name} your deposit of{amount} is successful.New balance {self.deposit}"
        print(msg)
        
    def get_row_number(self):
        
        pass
        
    
        
    
        
    
        
    
        