# Relevant Info
# 1. Type of Report - Index = 13
# 2. Customer - Index = 23
# 3. Location - Index = 29
# 4. Sample Point - Index = 33
# 5. Sample Date - Index = 26

# report = check_space(text[13])
# customer = check_space(text[23])
# location = check_space(text[29])
# samplePoint = check_space(text[33])
# sampleDate = check_space(text[26])


import os
from tkinter import filedialog
import functions
import PyPDF2

# os.chdir('C:\\Users\\tsl\\Desktop\\Lab Data')
#OneDrive - Pine Island Chemical\\Programming\\Python\\Projects')
fileName = filedialog.askopenfilename()

# Need to retain the file extension for renaming
# I split by '.' because i don't know how long the file extension is: Ex. .pdf or .xlsx
fileSplit = fileName.split('.')

# Add a period back to the last element of the split name
fileExtension = '.' + fileSplit[len(fileSplit) - 1]

pdfFileObj = open(fileName, 'rb')
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

pageObj = pdfReader.getPage(0)
text = pageObj.extractText()
text = text.split('\n')

receivedDate = text[text.index('Received Date:') + 8]
samplePoint = text[text.index('Received Date:') + 13]
customer = text[text.index('Received Date:') + 17]
report = text[text.index('Received Date:') + 28]
location = text[text.index('Received Date:') + 38]

reports = {'Bacteria Culture Analysis': 'BCA',
		'Corrosion Coupon Analysis': 'CCA',
		'Cold Finger Analysis': 'CFA',
		'Chloride Analysis': 'CLA',
		'Crude Oil Analysis': 'COA',
		'Corrosion Selection Test': 'CST',
		'Complete Water Analysis': 'CWA',
		'Iron and Manganese Analysis': 'FeMnA',
		'Gas Analysis': 'GSA',
		'Membrane Filter Analysis': 'MFA',
		'Oil and Grease Analysis': 'OGA',
		'Pipe Failure Analysis': 'PFA',
		'Scale Coupon Analysis': 'SCA',
		'Solids ID Analysis': 'SIDA',
		'Scale Inhibitor Residual': 'SIR',
		'Total Oil and Grease': 'TOG'
		}



print(receivedDate)
print(samplePoint)
print(customer)
print(report)
print(location)
print(fileExtension)

# print(fileName +' renamed to: ' + newName)

# os.rename(fileName, newName)