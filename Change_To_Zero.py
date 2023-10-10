

import openpyxl

# Load the Excel workbook
workbook = openpyxl.load_workbook(r"\\AZ99\Accounting\Private\JBusser\00 Automation Production\Close_was_completed.xlsx")

# Select the specific worksheet where you want to change the cells
worksheet = workbook['Sheet1']  # Replace 'Sheet1' with the actual sheet name

# Loop through the cells B1 to B5 and set their values to zero
for row in worksheet.iter_rows(min_row=2, max_row=6, min_col=2, max_col=4):  # B1 to B5
    for cell in row:
        cell.value = 0

# Save the modified workbook
workbook.save(r"\\AZ99\Accounting\Private\JBusser\00 Automation Production\Close_was_completed.xlsx")

# Close the workbook
workbook.close()
