# README

This program is a simple compare and move algorithm that compares the 'Inactive Employee' entries in the spreadsheet and compares it to the folder names in the Dropbox 'Active' folder. Before running there are only a couple of things needing to be done.

## Requirements
In order for the program to run, you will need to keep both the config.json and spreadsheet file in the same folder as the application (main.exe). The config.json will need to be edited as seen in the later section describing the whole process. When you would like to run the program you will double click main.exe

All paths should not have any "\\" characters in it. As you should replace all of them with "/".

## Employee Data Spreadsheet

You will need to run a report on Namely, this report will need to include all of the Employees with the columns of data in this order: 
```
Last Name, First Name, Employee Number (MC0000), Active/Inactive Employee
```

## Configure

In order to run this application you need to enter some information into the config.json file. This file can be opened and edited in the default windows Notepad application and is treated as a normal text file in regards to file encoding and decoding.

You will need to change the three variables labeled "ActiveFolderPath", "DepartedFolderPath", and "EmployeeSpreadsheetName". The formatting will appear as such:
```json
"ActiveFolderPath": "EDIT ME"
```

You will want to then put in the path of of the 'Active' folder in dropbox into those quotations and then do the same with the 'Departed' folder.
With the "EmployeeSpreadsheetName", you will want to only include the relative path. Meaning that you will need to include only the NAME of the file as well as the file extension.

## Logs

You can see what employees were moved (if any) in the 'logs' folder. The files in the folder are sorted by date_time so you will be able to find which employees were moved on each date.
