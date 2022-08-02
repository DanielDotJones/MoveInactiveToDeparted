import openpyxl
import os
import shutil
import json
from datetime import datetime
from pathlib import Path
import time
from os.path import dirname
import easygui


start_time = time.time()

employee_path = dirname(dirname(os.getcwd())) + '\\'

spreadsheetName = easygui.fileopenbox(msg=None, title="Please select the spreadsheet file", default=os.getcwd()+'\\*.xlsx', filetypes = ["*.xlsx"])

src_path = easygui.diropenbox(msg="Active Employee Folder: ", title=None, default=employee_path)

dst_path = easygui.diropenbox(msg="Inactive Employee Folder: ", title=None, default=employee_path)

print(src_path)
print(dst_path)
print(spreadsheetName)

xlsx_file = Path(spreadsheetName)
wb_obj = openpyxl.load_workbook(xlsx_file)
sheet = wb_obj.active
temp2 = []
temp4 = []

for row in sheet.iter_rows(max_row=566):
    if row[3].value == 'Inactive Employee':
        temp2.append((row[0].value + ', ' + row[1].value + ' ' + row[2].value).strip())
        temp4.append((row[2].value).strip())

    
temp1 = (next(os.walk(src_path))[1])

temp3 = []
index=0
for item4 in temp4:
    for item1 in temp1:
        if item1[-6:-4] == "MC":
            if item1[-6:] == item4:
                temp3.append(temp2[index])
            elif item1 == temp2[index]:
                temp3.append(temp2[index])
    index = index+1
    #if item in temp1: temp3.append(item)

now = datetime.now()
dt_string = now.strftime("%m-%d-%Y_%H-%M-%S")

if not os.path.exists(os.getcwd()+'/logs'):
    os.makedirs(os.getcwd()+'/logs')

temp3.sort()
with open("logs/"+dt_string+'.txt', 'w') as f:
    f.write("Begin Log\n\n")
    for items in temp3:
        shutil.move(src_path+'/'+items, dst_path)
        f.write("   -"+items+"\n")
        print(items)
    f.write("\n")
    f.write("Runtime: --- %s seconds ---\n" % (time.time() - start_time))
    f.write("End of Log")


