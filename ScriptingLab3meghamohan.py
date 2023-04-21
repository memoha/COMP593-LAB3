import sys
import openpyxl
import os
from openpyxl import Workbook
from datetime import date
import pandas as pd

def main():
    sales_csv = get_sales_csv()
    orders_directory = create_orders_directory(sales_csv)
    process_sales_data(sales_csv, orders_directory)

# Get path of sales data CSV file from the command line
def get_sales_csv():
    # Check whether command line parameter provided
    # Check whether provide parameter is valid path of file
    if len(sys.argv) < 2:
        print("Error: No Command line parameter is given.")
        sys.exit()
    csv_file = sys.argv[1]
    if not os.path.exists(csv_file):
        print("Error: Given File path does not exist.")
        
        
        sys.exit()
    
    return csv_file
    

# Create the directory to hold the individual order Excel sheets

def create_orders_directory(sales_csv):
    csv_directory = os.path.dirname(sales_csv)
    today = date.today().strftime('%Y-%m-%d')
    orders_directory = os.path.join(csv_directory, 'Orders' + today)

    if not os.path.exists(orders_directory):
        os.mkdir(orders_directory)


    return orders_directory

    # Get directory in which sales data CSV file resides
    # Determine the name and path of the directory to hold the order data files
    # Create the order directory if it does not already exist
    

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    # Insert a new "TOTAL PRICE" column into the DataFrame
    # Remove columns from the DataFrame that are not needed
    # Group the rows in the DataFrame by order ID
    # For each order ID:
        # Remove the "ORDER ID" column
        # Sort the items by item number
        # Append a "GRAND TOTAL" row
        # Determine the file name and full path of the Excel sheet
        # Export the data to an Excel sheet
        # TODO: Format the Excel sheet



    sales_data = pd.read_csv(sales_csv)
    sales_data['TOTAL PRICE'] = sales_data['QUANTITY OF ITEM'] * sales_data['PRICE OF ITEM']
    sales_data = sales_data[['DATE OF ORDER', 'ITEM NUMBER', 'PRODUCT LINE', 'PRODUCT CODE', 'QUANTITY OF ITEM', 'PRICE OF ITEM', 'TOTAL PRICE', 'STATUS', 'CUSTOMER NAME']]
    Orders = sales_data.groupby(by='ORDER ID')

    for Order_id, order_data in Orders:

        Order_data = Order_data.sort_values(by='ITEM NUMBER')
        Order_data = Order_data.reset_index(drop=True)
        grand_total = Order_data['TOTAL PRICE'].sum()
        Order_data = Order_data.append({'ITEM NUMBER': 'GRAND TOTAL', 'TOTAL PRICE': grand_total}, ignore_index=True)
       
        Order_file = os.path.join(orders_dir, str(Order_id) + ".xlsx")
        Order_data.to_excel(Order_file, Index=False, Engine='openpyxl')
        # TODO: Format the Excel sheet
        wb = openpyxl.load_workbook(Order_file)
        sheet = wb.worksheets[0]
        # Format the Excel sheet
        sheet.column_dimensions['A'].width = 11
        sheet.column_dimensions['B'].width = 13
        sheet.column_dimensions['C'].width = 15
        sheet.column_dimensions['D'].width = 15
        sheet.column_dimensions['E'].width = 15
        sheet.column_dimensions['F'].width = 13
        sheet.column_dimensions['G'].width = 13
        sheet.column_dimensions['H'].width = 10
        sheet.column_dimensions['I'].width = 30

        
        for row in sheet.item_rows(min_row=2, min_col=4, max_col=5):
            for cell in row:
                cell.number_format = '"$"#,##0.00'
        Workbook.save(Order_file)
        pass

if __name__ == '__main__':
    main()