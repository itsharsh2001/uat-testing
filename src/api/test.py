from flask import Flask, request, jsonify
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
import time
from datetime import datetime

from flask_cors import CORS
import pandas as pd
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import TimeoutException
import os
import win32com.client





def reports_checker(module_number,path,type):
    try:
        print('5')
        if module_number==1:
            print('6')
            sheet_name = 'Debtor_Confirmations'
            new_df = pd.read_excel(path, sheet_name=sheet_name)
            # print(new_df, 7)
            if type=='consolidated':
                download_directory = "C:/Users/harsh.vijaykumar/Downloads"
                file = 'DC_Consolidated.xlsx'
                file_path = os.path.join(download_directory, file)
                original_df = pd.read_excel(file_path)


                print(new_df.columns.to_list())
                print(original_df.columns.to_list())
                new_df = new_df.drop(new_df.columns[0], axis=1)
                original_df = original_df.drop(original_df.columns[0], axis=1)

                new_df.reset_index(drop=True, inplace=True)
                original_df.reset_index(drop=True, inplace=True)
            
            

            # new_df['Created On'] = pd.to_datetime(new_df['Created On'], format='%d-%m-%Y %H:%M:%S')
            # original_df['Created On'] = pd.to_datetime(original_df['Created On'], format='%d-%m-%Y %H:%M:%S')

            # current_date = datetime.now().date()

            
            # filtered_df = new_df[new_df['Created On'].dt.date == current_date]
                
                filtered_df = new_df[new_df['Created On'].isin(original_df['Created On'])]    
                filtered_df.rename(columns={'Total_reminders': 'Total reminders'}, inplace=True)
                filtered_df.reset_index(drop=True, inplace=True)
                print(filtered_df)
                print(original_df)

                lowercase_columns_df1 = [col.strip().lower() for col in filtered_df.columns]
                lowercase_columns_df2 = [col.strip().lower() for col in original_df.columns]

                # lowercase_columns_df1 = [col.lower() for col in original_df.columns]
                # lowercase_columns_df2 = [col.lower() for col in filtered_df.columns]

                print(lowercase_columns_df1,'123')
                print(lowercase_columns_df2,'123')
                print(lowercase_columns_df1 == lowercase_columns_df2)
                # column_names_match = original_df.columns.tolist() == filtered_df.columns.to_list()
                # print(column_names_match, 'columns names match')


                if original_df.shape[0] == filtered_df.shape[0]:
                    print('row count match')
                else:
                    print('row count nnot match')
                # if column_names_match:
                #     print("All column names match the expected list.")
                # else:
                #     print("Column names do not match the expected list.")   


                try:
                    # null_columns = new_df.columns[new_df.isnull().any()]
                    # print(null_columns)


                    print(filtered_df['Debit Balances'])
                    print(original_df['Debit Balances'])

                    mask1 = filtered_df['Debit Balances'].notnull()
                    mask2 = original_df['Debit Balances'].notnull()
                except Exception as e:
                    print('exc', e)    

                try:
                    print(filtered_df.loc[mask1,'Debit Balances']==original_df.loc[mask2,'Debit Balances'])
                    print('1234')
                except Exception as e:
                    print(e)
                    print('4321')
                # pd.set_option('display.max_columns', None)
                # pd.set_option('display.max_rows', None)
                print(filtered_df)
                print(original_df)

                try:
                    # original_df_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                    # filtered_df_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                    # data_equals = original_df_lower.iloc[:, 0:].equals(filtered_df_lower.iloc[:, 0:])
                    # print(data_equals)

                    df1_lower = original_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)
                    df2_lower = filtered_df.applymap(lambda x: x.lower() if isinstance(x, str) else x)

                    # Perform comparison and get the unequal parts
                    # comparison_result = df1_lower.compare(df2_lower)
                    # print(comparison_result)

#                     merged_df = df1_lower.merge(df2_lower, on='Debit Balances', suffixes=('_left', '_right'))

# # Identify rows where values are unequal
#                     unequal_rows = merged_df[merged_df['Debit Balances'] != merged_df['Debit Balances']]
#                     print(unequal_rows)

#                     excel_filename = 'unequal_rows.xlsx'
#                     unequal_rows.to_excel(excel_filename, index=False)

                except Exception as e:
                    print(e, 'total comparision')

                original_df.columns = original_df.columns.str.lower()
                filtered_df.columns = filtered_df.columns.str.lower()

                if list(original_df.columns) != list(filtered_df.columns):
                    print("Column names or order are different.")
                else:
    # Check index values
                    if not original_df.index.equals(filtered_df.index):
                        print("Index values are different.")
                    else:
                        print("DataFrames are identically labeled.")
                try:
                    mask3 = original_df.isnull()
                    mask4 = filtered_df.isnull()
                    placement_match = (mask3 == mask4).all().all()

                    if placement_match:
                        print("Null value placement matches in both DataFrames.")
                    else:
                        print("Null value placement does not match in both DataFrames.")
                    
                except Exception as e:
                    print(e)
                    


    except Exception as e:
        print(e)

if __name__ == "__main__":
    # excel_file = "data.xlsx"  # Change this to your Excel file name
    # check_null_columns(excel_file)
    download_directory = "C:/Users/harsh.vijaykumar/Downloads"
    print('1')
    file_name = "client_Consolidated.xlsx"
    print('2')
    file_path = os.path.join(download_directory, file_name)
    print('3')
    reports_checker(1,file_path,'consolidated')
    print('4')

