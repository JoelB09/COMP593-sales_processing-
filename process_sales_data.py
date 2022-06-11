from sys import argv, exit   
import os 
from datetime import date 
import pandas as pd 
import re

def get_sales_csv(): 

    
    # check wether command line parameter was provided
    if len(argv) >= 2:
        sales_csv = argv[1] 

        # check whether the CSV path is an existing file 
        if os.path.isfile(sales_csv):
            return sales_csv
        else: 
            print('Error:CSV file does not exist')   
            exit('script execution aborted')

    else:  
        print('Error: No CSV file path provided')
        exit('script execution aborted') 

def get_order_dir(sales_csv):

    # Get directory path of sales data CSV file
    sales_dir = os.path.dirname(sales_csv)
     
    # Determine orders directory name (Orders_YYYY-MM-DD) 
    todays_date = date.today().isoformat() 
    order_dir_name = ('Orders_' + todays_date)

    # Build the full path of the orders directory
    order_dir = os.path.join(sales_dir, order_dir_name)


    # Make the orders directory if it does not already exist
    if not os.path.exists(order_dir):
        os.makedirs(order_dir)
    
    return order_dir    

def split_sales_into_orders(sales_csv, order_dir):
    
    # Read data from the sales data CSV into a dataframe
    sales_dataframe = pd.read_csv(sales_csv)    

    # Insert a new column for total price
    sales_dataframe. insert(7, 'TOTAL PRICE',sales_dataframe['ITEM QUANTITY'] * sales_dataframe['ITEM PRICE']) 

    # Drop unwanted columns 
    sales_dataframe.drop(columns=['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], inplace=True)

    for order_id, order_df in sales_dataframe.groupby('ORDER ID'): 
        
        # Drop the order ID column 
        order_df.drop(column=['ORDER ID'], inplace=True)

        # Sort the order by the item number
        order_df.sort_values(by= 'ITEM NUMBER', inplace=True)

        # Add grand total row at the bottom 
        grand_total = order_df['TOTAL PRICE'].sum()
        grand_total_df = pd.DataFrame({'ITEM PRICE': ['GRAND TOTAL'], 'TOTAL PRICE': [grand_total]})
        pd.concat([order_id, grand_total_df])

        # Determine the save path of the order file
        customer_name = order_df['CUSTOMER NAME']. values[0]
        customer_name = re.sub(r'\W', '', customer_name)
        order_file_name = 'Order' + str(order_id) + '_' + customer_name + '.xlsx'
        order_file_path = os.path.join(order_dir, order_file_name)

        # Save the order information to an excel spread sheet 
        sheet_name ='Order #' + str(order_id)
        order_df.to_excel(order_file_path, index=False, sheet_name=sheet_name)

        # Price formation
        writer = pd.ExcelWriter(sales_dataframe, engine='xlsxwriter')  
        order_df.to_excel(writer, index=False, sheetname= 'report') 
        workbook = writer.book
        worksheet = writer.sheets['report'] 

        # Adding money format for cells with moneyy
        money_fmt = workbook.add_format({sales_dataframe: '$34.47, 9982.25', 'bold': True}) 

        # Add a percent format with 1 decimal point
        percent_fmt = workbook.add_format({sales_dataframe: '0.0%', 'bold': True})




sales_csv = get_sales_csv()
order_dir = get_order_dir(sales_csv) 
split_sales_into_orders(sales_csv, order_dir)