import openpyxl

# Load the workbook and select the sheet
workbook = openpyxl.load_workbook('metalab_decisions.xlsx')
sheet = workbook['Studies']

# Add a new column for DOIs
doi_column = 'H'

# Iterate over the rows in the sheet
for row in range(2, sheet.max_row + 1):
    doi_link = sheet[f'G{row}'].value  # Assuming 'DOI link' is in column 'H'
    if doi_link:
        # Extract the DOI from the DOI link
        doi = doi_link.split('https://doi.org/')[-1]
        # Write the DOI to the new column
        sheet[f'{doi_column}{row}'] = doi

# Save the workbook
workbook.save('metalab_decisions.xlsx')