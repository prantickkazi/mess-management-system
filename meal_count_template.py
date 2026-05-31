import openpyxl
import datetime
import calendar

from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter
from config import workbook_name,meal_count_sheet_name

#initialize the work sheet
def initialize_meal_count_sheet():
    wb=openpyxl.load_workbook(workbook_name)
    ws=wb.create_sheet(title=meal_count_sheet_name)
    wb.save(workbook_name)
  
#create the worksheet completely  
def create_meal_count_sheet():
    wb=openpyxl.load_workbook(workbook_name)
    ws=wb[meal_count_sheet_name]
    #date and month
    now=datetime.datetime.now()
    year=now.year
    month=now.month
    days_in_month=calendar.monthrange(year,month)[1]
    month_name=now.strftime("%B, %Y")
    #insert the current month/year in the first row
    row1=[f"{month_name} date->", "", ""]
    for day in range(1,days_in_month+1):
        row1.extend([day, "","",""])
    ws.append(row1)
    #insert in the second row
    row2 = ["Name", "Status", "Deposit"]
    for day in range(1,days_in_month+1):
        row2.extend(["day", "guest(day)", "night", "guest(night)"])
    ws.append(row2)
    #apply style and colours
    black_bolt=Font(name="Calibri",size=11,bold=True,color="000000")
    center_align=Alignment(horizontal="center",vertical="center")
    fixed_header_fill=PatternFill(start_color="117A65", end_color="117A65", fill_type="solid")
   
    for row in range(1,3):
        for col in range(1,4):
            cell=ws.cell(row=row,column=col)
            cell.fill=fixed_header_fill
            cell.font=black_bolt
            cell.alignment=center_align
    
    #apply alternating colour to days
    colored_day_fill = PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid") 
    uncolored_fill = PatternFill(fill_type=None)
    
    for day in range(1,days_in_month+1):
        start_col=4+(day-1)*4
        end_col=start_col+3
        
        current_fill=colored_day_fill if day % 2 == 0 else uncolored_fill
        
        for col in range(start_col,end_col+1):
            for row in range(1,3):
                cell=ws.cell(row=row,column=col)
                cell.fill=current_fill
                
                if row <=2:
                    cell.font=black_bolt
                    cell.alignment=center_align
    #formatting and layout
    #merge the data cells
    for day in range(1, days_in_month + 1):
        start_col = 4 + (day - 1) * 4
        end_col = start_col + 3
        ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)

    # Adjust base column widths
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 15

    # Calculate exact number of columns used so we don't format empty ones!
    total_cols = 3 + (days_in_month * 4)
    for col in range(4, total_cols + 1):
        col_letter = get_column_letter(col)
        ws.column_dimensions[col_letter].width = 11

    ws.freeze_panes = "D3"
    
    wb.save(workbook_name)
    print(f"Sheet {meal_count_sheet_name} for {month_name} ({days_in_month} days) created in '{workbook_name}'!")

#add boarder in the sheet
def add_boarder_insheet():
    user_name=input("enter your name")
    
    
# initialize_meal_count_sheet()
# create_meal_count_sheet()
                
        
    
    
    

    