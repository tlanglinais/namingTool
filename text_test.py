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

print(functions.preceding_zero('2'))


filename = filedialog.askopenfilename()
print(f'Opening {filename}')
text = open(filename, 'r').read()
text = text.split('\n')



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

newName = f'{reports[check_space(text[13])]}_{check_space(text[23])}_{check_space(text[29])}_{check_space(text[33])}_{format_date(text[26])}.txt'

print(filename +' renamed to: ' + newName)

os.rename(filename, newName)