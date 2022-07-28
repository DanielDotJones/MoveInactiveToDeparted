# README

This program is a simple compare and move algorithm that compares the 'Inactive Employee' entries in the spreadsheet and compares it to the folder names in the Dropbox 'Active' folder. Before running there are only a couple of things needing to be done.

## Requirements

You will need to download a report from Namely in the .xls format, to see what columns to include look at the 'Employee Data Spreadsheet' section. After downloading you will need to open the file in excel and export the file as type .xlsx

When starting the program, you will be prompted with a file explorer window where you will need to select the spreadsheet exported from Namely
Next, you will select the folder with all of the active employees (Generally the dropbox folder) and after selecting that it will prompt you to do the same with the departed folder.

## Employee Data Spreadsheet

You will need to run a report on Namely, this report will need to include all of the Employees with the columns of data in this order: 
```
Last Name, First Name, Employee Number (MC0000), Active/Inactive Employee
```

## Logs

You can see what employees were moved (if any) in the 'logs' folder. The files in the folder are sorted by date_time so you will be able to find which employees were moved on each date.
