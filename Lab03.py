from sys import argv, exit
import os
from datetime import date
import pandas as pd
import re 

def main():
    sales_csv = get_sales_csv()    
    orders_dir = create_orders_dir(sales_csv)
    sales_data = process_sales_data(sales_csv, orders_dir)
    return

# Get path of sales data CSV file from the command line
def get_sales_csv():
    # Check whether command line parameter provided
    numb_paras = len(argv) -1
    if numb_paras>= 1:
        sales_csv = argv[1]
        # Check whether provide parameter is valid path of file
        if os.path.isfile(sales_csv):
            return sales_csv
        else:
            print('Error: invalid path')
            exit(1)
    else:
        print('Error: Missing path to sales data CSV file')
        exit(1)

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    # Get directory in which sales data CSV file resides
    sales_dir = os.path.dirname(os.path.abspath(sales_csv))
    # Determine the path of the directory to hold the order data files
    todays_date = date.today().isoformat()
    orders_dir = os.path.join(sales_dir, f'Orders_{todays_date}')
    # Create the order directory if it does not already exist
    if not os.path.isdir(orders_dir):
        os.makedirs(orders_dir)
    return orders_dir

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    sales_df = pd.read_csv(sales_csv)
    # Insert a new "TOTAL PRICE" column into the DataFrame
    sales_df.insert(7, 'TOTAL PRICE', sales_df['ITEM QUANTITY'] * sales_df['ITEM PRICE'])
    # Remove columns from the DataFrame that are not needed
    sales_df.drop(columns = ['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], inplace=True)
    # Group sales data by order ID and save to Excel sheets
    for order_id, order_df in sales_df.groupby('ORDER ID'):
        # Remove the 'ORDER ID' column
        order_df.drop(columns = ['ORDER ID'], inplace=True)
        # Sort the order by item number
        order_df.sort_values(by='ITEM NUMBER', inplace=True)
        # Add the grand total row
        grand_total = order_df['TOTAL PRICE'].sum()
        grand_total_df = pd.DataFrame({'ITEM PRICE': ['GRAND TOTAL'], 'TOTAL PRICE': [grand_total]})
        order_df = pd.concat([order_df, grand_total_df])
        export_order_to_excel(order_id, order_df, orders_dir)
    return 

def export_order_to_excel(order_id, order_df, orders_dir):
    # Determine the file name and path for the order Excel sheet
    customer_name = order_df['CUSTOMER NAME'].values[0]
    customer_name = re.sub(r'\W', '', customer_name)
    order_file = f'ORDER{order_id}_{customer_name}.xlsx'
    order_path = os.path.join(orders_dir, order_file)
    sheet_name = f'Order #{order_id}'
    # Formating the code to put the $ sign and use 2 decimal spots
    money_format = pd.ExcelWriter(order_path, engine='openpyxl')
    order_df.style.format({'TOTAL PRICE': '${:,.2f}', 'ITEM PRICE': '${:,.2f}'}).to_excel(order_path, index=False, sheet_name=sheet_name)
    # Making it so that the column widths match the image given in the lab
    size = money_format.book.get_sheet_by_name(sheet_name)
    size.column_dimensions['ORDER DATE'].width = 11
    size.column_dimensions['ITEM NUMBER'].width = 13
    size.column_dimensions['PRODUCT LINE'].width = 15
    size.column_dimensions['PRODUCT CODE'].width = 15
    size.column_dimensions['ITEM QUANTITY'].width = 15
    size.column_dimensions['ITEM PRICE'].width = 13
    size.column_dimensions['TOTAL PRICE'].width = 13
    size.column_dimensions['STATUS'].width = 10
    size.column_dimensions['CUSTOMER NAME'].width = 30
    money_format.save()
    return

if __name__ == '__main__':
    main()