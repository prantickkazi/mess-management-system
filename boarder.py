from openpyxl import load_workbook
from excel_utils import write_in_excel_file
from config import workbook_name,meal_count_sheet_name

class Boarder:
    def __init__(self,name):
        self.name=name
        self.deposit=0
        self.row_number=None
        self.total_own_meal=0
        self.total_guest_meal=0
        self.active=None
        self.night_active=None
        self.day_active=None
        self.guest_active=None
        self.guest_day_active=None
        self.guest_night_active=None
        self.total_meal=self.total_guest_meal + self.total_own_meal

    def check_active_status(self):
        if self.day_active  or self.night_active :
            self.active = True
        elif self.guest_day_active or self.guest_night_active :
            self.guest_active = True
        else:
            self.active = None
        
    def on_meal(self,info= "yes"):
        if info == "yes":
            if self.night_active  is None and self.day_active is None:
                self.night_active=True
                self.day_active=True
                print(f"your both meal is on")
                self.total_own_meal+=2
            else:
                print("Your meal is active already")
        elif info == "on":
            if self.night_active  is None:
                self.night_active = True
                self.total_own_meal+=1
            else :
                print("Your night meal is already on")
        elif info == "od":
            if self.day_active is None:
                self.day_active=True
                self.total_own_meal+=1
            else :
                print("Your day meal is on already")
        self.check_active_status()
           
    def off_meal(self,info):
        if info == "off":
            if self.day_active  and  self.night_active :
                self.day_active = None
                self.night_active = None
                self.total_own_meal -=2
            else:
                print(f"You do not have any active meal status")
        elif info == "od":
            if self.day_active :
                self.day_active = None
                self.total_own_meal  -= 1
            else :
                print("Your day meal is off already")
        elif info == "on":
            if self.night_active :
                self.night_active = None
                self.total_own_meal -= 1
            else:
                print("Your night meal is off already")
        self.check_active_status()
            
            
    def on_guest_meal(self,info):
        if info == "yes":
            if self.guest_day_active  is None and self.guest_night_active is None:
                self.guest_active = True
                self.guest_day_active = True
                self.guest_night_active = True
                self.total_guest_meal += 2
            else:
                print("your meal is active already")   
        elif info == "od":
            if self.guest_night_active is None:
                self.guest_active = True
                self.guest_day_active = True
                self.total_guest_meal += 1
            else:
                print("active already")
        elif info == "on":
            if self.guest_night_active is None:
                self.guest_active = True
                self.guest_night_active = True
                self.total_guest_meal += 1
            else:
                print("active already")
        self.check_active_status()
        
    def off_guest_meal(self,info):
        if info == "off":
            if self.guest_day_active  and self.guest_night_active :
                self.guest_day_active = None
                self.guest_night_active = None
                self.total_guest_meal -= 2
            else:
                print("Both meal are off already")
        if info == "od":
            if self.guest_day_active :    
                self.guest_day_active = None
                self.total_guest_meal -= 1
            else:
                print("Off already")
        if info == "on":
            if self.guest_night_active :
                self.guest_night_active = None
                self.total_guest_meal -= 1
            else:
                print("off already")
        self.check_active_status()
                
    def add_deposit(self,amount):
        self.deposit+=amount
        msg=f"{self.name} your deposit of{amount} is successful.New balance {self.deposit}"
        print(msg)
        
        
    
        
    
        
    
        
    
        