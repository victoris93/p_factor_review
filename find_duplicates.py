import openpyxl

# Load the workbook and select the sheet
workbook = openpyxl.load_workbook('metalab_decisions.xlsx')
sheet = workbook['Studies']

# Create a dictionary to store titles and their row numbers
titles = {}

# Iterate over the rows in the sheet
for row in range(2, sheet.max_row + 1):
    unique_id = sheet[f'A{row}'].value
    title = sheet[f'B{row}'].value 
    decision = sheet[f'I{row}'].value  

    if title in titles:
        sheet[f'A{row}'] = titles[title]['unique_id']
        sheet[f'I{row}'] = 'duplicate'
    else:
        titles[title] = {'unique_id': unique_id, 'row': row}

# Save the workbook
workbook.save('metalab_decisions.xlsx')