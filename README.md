# Jupyter Notebooks EDA

 Final Project: Black Friday Excel Data
Overview
This project processes sales data from an Excel spreadsheet for a fictional ski shop and performs the following key tasks:

Reading data from the Excel workbook.
Printing column data for easy inspection.
Creating a structured dictionary from the Excel data.
Calculating sales tax and total amounts based on location-specific tax rates.
Writing the calculated data back to the Excel sheet and saving it.
This README provides an overview of the project structure, functionality, and usage.

Prerequisites
Python 3.x
Libraries:
openpyxl (for Excel file manipulation)
A custom module tax_calculator (must be available in the project directory)
pprint (for displaying structured data)
Install openpyxl if it's not already installed:

bash
Copy code
pip install openpyxl
Project Components
1. Reading the Excel File
The project begins by loading the maven_ski_shop_data.xlsx file using the openpyxl library:

python
Copy code
wb = xl.load_workbook(filename='maven_ski_shop_data.xlsx')
orders = wb['Orders_Info']
2. Column Printer Function
The column_printer function facilitates easy viewing of column data without manually opening the Excel file. It prints the cell coordinates and contents of a specified column.

Example Usage:
python
Copy code
column_printer(orders, 'A')  # Prints Order IDs
column_printer(orders, 'C')  # Prints Subtotals
3. Order Data Dictionary
A dictionary is created to store order data from the Orders_Info sheet. The structure is as follows:

Key: Order ID (Column A)
Value: A list containing data from columns B, C, D, G, and H (split into a list).
Implementation:
python
Copy code
order_dict = {
    orders[f'A{order}'].value: [
        orders[f'B{order}'].value,
        orders[f'C{order}'].value,
        orders[f'D{order}'].value,
        orders[f'G{order}'].value,
        str(orders[f'H{order}'].value).split(', ')
    ]
    for order in range(2, orders.max_row + 1)
}
4. Sales Tax Calculation
The sales tax and total amounts owed are calculated based on the location:

Sun Valley: 8%
Mammoth: 7.75%
Stowe: 6%
The tax_calculator function is used to compute:

Sales tax
Total amount owed
Integration Example:
python
Copy code
for order in order_dict.values():
    if order[3] == 'Sun Valley':
        transaction = tax_calculator(order[2], .08)
    elif order[3] == 'Mammoth':
        transaction = tax_calculator(order[2], .0775)
    else:
        transaction = tax_calculator(order[2], .06)
    order.insert(3, transaction[1])
    order.insert(4, transaction[2])
5. Writing Data Back to Excel
The calculated sales tax and total amounts are written into the workbook, which is then saved with the filename maven_ski_shop_data_fixed.xlsx.

Implementation:
python
Copy code
for idx, order in enumerate(order_dict.values(), start=2):
    orders[f'E{idx}'] = order[3]  # Sales Tax
    orders[f'F{idx}'] = order[4]  # Total Amount

wb.save('maven_ski_shop_data_fixed.xlsx')
Usage Instructions
Clone the Repository Download or clone the project files to your local machine.

Ensure Required Files Exist

maven_ski_shop_data.xlsx (input file)
tax_calculator.py (custom module for tax calculations)
Run the Script Execute the Python script to process the data:

bash
Copy code
python main.py
Output

The script outputs dictionary contents and column data for verification.
The modified workbook is saved as maven_ski_shop_data_fixed.xlsx.
Example Output
Dictionary Sample:
python
Copy code
{
    10001: ['John Doe', 200.0, 'Sun Valley', 16.0, 216.0, ['Skis', 'Poles']],
    10002: ['Jane Smith', 150.0, 'Mammoth', 11.625, 161.625, ['Boots']]
}
Excel File:
Sales Tax written in Column E.
Total Amount written in Column F.

