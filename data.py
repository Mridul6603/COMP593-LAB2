from sys import argv, exit
import os
from datetime import date
import pandas as pd
import re

def get_sales_csv():


    if len(argv) >= 2:
        sales_csv = argv[1]


        if os.path.isfile(sales_csv):
            return sales_csv
        else:
            print("CSV file does not exist")
            exit('Script terminated')  
    else:
        print("Sorry no file was imported as there is no path provided")
        exit('Script terminated')  

def get_order_dir(sales_csv):

    #Get the directory  path of sales csv file
    sales_dir = os.path.dirname(sales_csv)
    #Determine orders directory name (Orders_YYYY_MM_DD)
    todays_date = date.today().isoformat()
    order_dir_name =  'Orders_' + todays_date

    #Build the full path of the orders directory
    order_dir = os.path.join(sales_dir, order_dir_name)

    #make orders directory if it does not already exist
    if not os.path.exists(order_dir):
        os.makedirs(order_dir)

    return order_dir

def split_sales_into_orders(sales_csv,order_dir):


    sales_df = pd.read_csv(sales_csv)

    #Insert a new column as requested in the question
    sales_df.insert(7, 'TOTAL PRICE', sales_df['ITEM QUANTITY'] * sales_df['ITEM PRICE']) 

    sales_df.drop(columns=['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], inplace=True)

    for order_id, order_df in sales_df.groupby('ORDER ID'):

        #DRop the ID column
        order_df.drop(columns=  ['ORDER ID'], inplace=True)

        #Sort by ITem No.
        order_df.sort_values(by='ITEM NUMBER', inplace=True)

        #add grand total at the bottom
        grand_total =  order_df['TOTAL PRICE'].sum()
        grand_total_df = pd.DataFrame({'ITEM PRICE':['GRAND TOTAL:'], 'TOTAL PRICE': [grand_total]})
        order_df = pd.concat([order_df, grand_total_df])

        #Determine the save path of the file
        customer_name = order_df['CUSTOMER NAME'].values[0]
        customer_name = re.sub(r'/W', '', customer_name)
        order_file_name= 'Order' + str(order_id) + '_' + customer_name + '.xlsx'
        order_file_path = os.path.join(order_dir, order_file_name)

        #Save the order of information to Excel
        sheet_name = 'Order #' + str(order_id)
        order_df.to_excel(order_file_path, index=False, sheet_name=sheet_name)

        

        



sales_csv = get_sales_csv()
order_dir = get_order_dir(sales_csv)
split_sales_into_orders(sales_csv,order_dir)


