import openpyxl
import os
import shutil
import json
from datetime import datetime
from pathlib import Path
import time
from os.path import dirname
import easygui
import logging

start_time = time.time()
now = datetime.now()
dt_string = now.strftime("%m-%d-%Y_%H-%M-%S")
preLogTxt = ''


def preLog(text):
    global preLogTxt
    preLogTxt = preLogTxt + '['+str(time.time())+']:  ' + text + '\n'
    
def checkLogFolder():
    preLog("Checking if log folder exists...")
    if not os.path.exists(os.getcwd()+'/logs'):
        preLog("Folder didnt exist, attempting to create a new log folder")
        folderLocation = os.getcwd()+'/logs'
        os.makedirs(folderLocation)
        preLog("Created a new log folder at "+folderLocation)
    else:
        preLog("Log Folder Exists")

def initLogging():
    logging.basicConfig(filename="logs/"+dt_string+'.txt', level=logging.INFO, format="%(asctime)s %(message)s")

#logFile = open("logs/"+dt_string+'.txt', 'a')

def log(text):
    print('['+str(time.time())+']:  ' + text)
    logging.info('['+str(time.time())+']:  '+text)
    


def main():
    checkLogFolder()
    initLogging()
    
    log("Begin Log")

    log(preLogTxt)

    log("Attempting to create file paths...")

    working_dir = os.getcwd() + '\\'
    log("Retrieved Working Directory Path at " + working_dir)

    spreadsheetName = easygui.fileopenbox(msg=None, title="Please select the spreadsheet file", default=working_dir + '*.xlsx', filetypes = ["*.xlsx"])
    log("Retrieved Spreadsheet File Path at " + spreadsheetName)

    src_path = easygui.diropenbox(msg="Active Employee Folder: ", title=None, default=working_dir)
    log("Retrieved Active Employee Folder Path at " + src_path)

    dst_path = easygui.diropenbox(msg="Inactive Employee Folder: ", title=None, default=working_dir)
    log("Retrieved Inactive Employee Folder Path at " + dst_path)

    log("Attempting to create path object to spreadsheet...")
    xlsx_file = Path(spreadsheetName)
    log("Created path object to spreadsheet")
    log("Attempting to load the spreadsheet workbook...")
    wb_obj = openpyxl.load_workbook(xlsx_file)
    log("Loaded the spreadsheet workbook")
    sheet = wb_obj.active
    log("Loaded the active sheet")

    temp2 = []
    temp4 = []

    log("Attempting to find inactive employees to seperate into two arrays...")
    for row in sheet.iter_rows(max_row=566):
        if row[3].value == 'Inactive Employee':
            temp2.append((row[0].value + ', ' + row[1].value + ' ' + row[2].value).strip())
            temp4.append((row[2].value).strip())
    log("Arrays created")

    log("Attempting to read folder names from the Active (Employee) Folder")
    temp1 = (next(os.walk(src_path))[1])

    log("Attempting to compare the Active Employee Array to the Inactive Employee Array...")
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
    log("Comparison and resulting array created")



    log("Attempting to sort the resulting array...")
    temp3.sort()
    log("Resulting array sorted")

    if not (len(temp3) == 0):
        log("Attempting to move the files below:\n")
        for items in temp3:
            shutil.move(src_path+'/'+items, dst_path)
            log("   -"+items)
            print(items)
        log("Files have been moved")
    else:
        log("No files to move")
        
    log("\n")
    log("Runtime: --- %s seconds ---" % (time.time() - start_time))
    log("End of Log")

if __name__ == "__main__":
   try:
      main()
   except Exception as e:
      logging.exception("main crashed. Error: %s", e)
