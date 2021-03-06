# Relevant Info
# 1. Type of Report
# 2. Customer
# 3. Location
# 4. Sample Point
# 5. Sample Date

import os
from tkinter import filedialog
import openpyxl
import datetime
import functions
# import convert

# Open log file
logFile = open('logfile.txt', 'a')
# Promp File Select
fileName = filedialog.askopenfilename()
logFile.write(
    f'<----------Current Date/Time: {datetime.datetime.today()}---------->')
logFile.write(f'\n\nOpening {fileName}')

# Need to retain the file extension for renaming
# I split by '.' because i don't know how long the file extension is: Ex. .pdf or .xlsx
fileSplit = fileName.split('.')

# Add a period back to the last element of the split name
fileExtension = '.' + fileSplit[len(fileSplit) - 1]
logFile.write(f'\nFiletype: {fileExtension}')

data_map = {
    # 'Bacteria Culture Analysis': {'abbreviation': 'BCA', 'multipleFlag': 0, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': '', 'Date Received': '', 'Date Completed': ''},
    'Microbial Incubation Report': {'abbreviation': 'BCA', 'multipleFlag': 0, 'Customer': 'B6', 'Location': 'B7', 'Sample Point': 'B8', 'Sample Date': 'K7', 'Date Received': 'K8', 'Date Completed': 'K9'},
    'Corrosion Coupon Analysis Report': {'abbreviation': 'CCA', 'multipleFlag': 1, 'Customer': 'B6', 'Location': 'B7', 'Sample Point': 'A11', 'Sample Date': 'J5', 'Date Received': 'J6', 'Date Completed': 'J7'},
    'Cold Finger Analysis': {'abbreviation': 'CFA', 'multipleFlag': 0, 'Customer': 'C6', 'Location': 'C7', 'Sample Point': '', 'Sample Date': 'H5', 'Date Received': 'H6', 'Date Completed': 'H7'},
    # 'Chloride Analysis': {'abbreviation': 'CLA', 'multipleFlag': 0, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': '', 'Date Received': '', 'Date Completed': ''},
    'Comprehensive Crude Oil Analysis': {'abbreviation': 'COA', 'multipleFlag': 1, 'Customer': 'C6', 'Location': 'C7', 'Sample Point': 'A12', 'Sample Date': 'H5', 'Date Received': 'H6', 'Date Completed': 'H7'},
    'Sparged Beaker Analysis': {'abbreviation': 'CST', 'multipleFlag': 0, 'Customer': 'C6', 'Location': 'C7', 'Sample Point': '', 'Sample Date': 'H5', 'Date Received': 'H6', 'Date Completed': 'H7'},
    'Comprehensive Oilfield Water Analysis': {'abbreviation': 'CWA', 'multipleFlag': 0, 'Customer': 'C5', 'Location': 'C6', 'Sample Point': 'C7', 'Sample Date': 'H5', 'Date Received': 'H6', 'Date Completed': 'H7'},
    'Iron and Manganese Analysis': {'abbreviation': 'FeMnA', 'multipleFlag': 1, 'Customer': 'C6', 'Location': 'C7', 'Sample Point': 'A10', 'Sample Date': 'G5', 'Date Received': 'G6', 'Date Completed': 'G7'},
    'Iron, Manganese, & Chloride Analysis': {'abbreviation': 'FeMnA', 'multipleFlag': 1, 'Customer': 'C6', 'Location': 'D7', 'Sample Point': 'A10', 'Sample Date': 'J5', 'Date Received': 'J6', 'Date Completed': 'J7'},
    # 'Gas Analysis': {'abbreviation': 'GSA', 'multipleFlag': 0, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': '', 'Date Received': '', 'Date Completed': ''},
    'Membrane Filter Analysis': {'abbreviation': 'MFA', 'multipleFlag': 0, 'Customer': 'C5', 'Location': 'C6', 'Sample Point': 'C7', 'Sample Date': 'I5', 'Date Received': 'I6', 'Date Completed': 'I7'},
    'Oil & Grease in Water by Infrared Analysis': {'abbreviation': 'OGA', 'multipleFlag': 1, 'Customer': 'C6', 'Location': 'C7', 'Sample Point': 'A10', 'Sample Date': 'G5', 'Date Received': 'G6', 'Date Completed': 'G7'},
    # 'Pipe Failure Analysis': {'abbreviation': 'PFA', 'multipleFlag': 0, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': '', 'Date Received': '', 'Date Completed': ''},
    'Scale Coupon Analysis Report': {'abbreviation': 'SCA', 'multipleFlag': 1, 'Customer': 'B6', 'Location': 'B7', 'Sample Point': 'A10', 'Sample Date': 'K5', 'Date Received': 'K6', 'Date Completed': 'K7'},
    'Quantitative Solids Identification': {'abbreviation': 'QSI', 'multipleFlag': 0, 'Customer': 'C5', 'Location': 'C6', 'Sample Point': 'C7', 'Sample Date': 'I5', 'Date Received': 'I6', 'Date Completed': 'I7'},
    # 'Solids ID Analysis': {'abbreviation': 'SIDA', 'multipleFlag': 0, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': '', 'Date Received': '', 'Date Completed': ''},
    # 'Scale Inhibitor Residual': {'abbreviation': 'SIR', 'multipleFlag': 0, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': '', 'Date Received': '', 'Date Completed': ''},
    # 'Total Oil and Grease': {'abbreviation': 'TOG', 'multipleFlag': 0, 'Customer': '', 'Location': '', 'Sample Point': '', 'Sample Date': '', 'Date Received': '', 'Date Completed': ''}
}

# # Check to see if its a .xls spreadsheet
# if fileExtension == '.xls':
# 	wb = convert.open_xls_as_xlsx(fileName)
# 	sheet = wb.active

# else:
# load .xlsx workbook
wb = openpyxl.load_workbook(fileName)
sheet = wb.active

c1 = functions.check_space(sheet['C1'].value)
d1 = functions.check_space(sheet['D1'].value)


# find if the report name is in cell C1 or D1
if c1 in data_map:
    reportName = c1
    logFile.write(f'\nReport Type: {reportName}')

elif d1 in data_map:
    reportName = d1
    logFile.write(f'\nReport Type: {reportName}')

else:
    logFile.write('\n---Error: Cannot find the header---')

# Grab customer name and format for renaming
customer = functions.format_name(sheet[data_map[reportName]['Customer']].value)
logFile.write(f"\nCustomer: {customer}")

# Grab location and format for renaming
location = functions.format_name(sheet[data_map[reportName]['Location']].value)
logFile.write(f'\nLocation: {location}')


# Grab sample point and format for renaming
# ----------------------------------------------------------NEED TO CHECK IF THERE ARE MULTIPLE SAMPLE POINTS
if data_map[reportName]['multipleFlag']:
    # Need to see if there are more than one sample points
    sample1 = data_map[reportName]['Sample Point']
    # Chop up the cell name and increment to 1 cell below
    sample2 = 'A' + str(int(sample1[1:]) + 1)
    # print(f'The cell under {sample1} is {sample2}')
    # print(f'The value in cell {sample2} is {sheet[sample2].value} and is of type {type(sheet[sample2].value)}')

    # if the cell under the first has a value, rename to 'Multiple'
    if sheet[sample2].value != None:
        samplePoint = 'Multiple_'
        logFile.write(
            "\nThere are multiple sample locations, so renaming to 'Multiple'")

    # if the value is empty, it means there is only 1 sample point
    else:
        samplePoint = functions.format_name(
            sheet[data_map[reportName]['Sample Point']].value)
        logFile.write(f'\nSample Point: {samplePoint}')
else:
    if data_map[reportName]['Sample Point'] != '':
        samplePoint = functions.format_name(
            sheet[data_map[reportName]['Sample Point']].value)
        logFile.write(f'\nSample Point: {samplePoint}')

    else:
        samplePoint = ''
        logFile.write(f'\nSample Point Missing')

# Grab sample date and format for renaming
# If sample date is empty, check date received, then date completed

if isinstance(sheet[data_map[reportName]['Sample Date']].value, datetime.date):
    sampleDate = functions.format_date(
        sheet[data_map[reportName]['Sample Date']].value)
    logFile.write(f'\nSample Date: {sampleDate}')

elif isinstance(sheet[data_map[reportName]['Date Received']].value, datetime.date):
    sampleDate = functions.format_date(
        sheet[data_map[reportName]['Date Received']].value)
    logFile.write(f'\nDate Received: {sampleDate}')

elif isinstance(sheet[data_map[reportName]['Date Completed']].value, datetime.date):
    sampleDate = functions.format_date(
        sheet[data_map[reportName]['Date Completed']].value)
    logFile.write(f'\nDate Completed: {sampleDate}')

else:
    sampleDate = ''
    logFile.write(f'\nThe Sample Date is Missing')

# Grab the report abbreviation from the reports dictionary
reportAbr = data_map[reportName]['abbreviation']
logFile.write(f'\nReport Abbreviation: {reportAbr}')

# Assemble the new filename
newName = reportAbr + '_' + customer + location + \
    samplePoint + sampleDate + fileExtension
logFile.write(f'\nNew File Name: {newName}')
# Rename the file to the new name and move the file to the desired path
os.rename(fileName, 'C:\\Users\\tsl\\desktop\\Lab Data\\Renamed\\' + newName)

# 'Log' Message
logFile.write(f'\n\nRenamed:\n{fileName[30:]}\nto:\n{newName}\n\n')
logFile.close()
