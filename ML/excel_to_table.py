# Importing the required libraries
import os
import numpy as np
import pandas as pd
from functions import *
from datetime import datetime

# Main function to convert the excel from the ilovepdf.py to a table
def excel_to_table(excel_path,download_folder):
    try:
        excel_sheets = pd.read_excel(excel_path, sheet_name=None,header=None)

        # Dropping the rows with all NaN values
        for (sheet,sheet_name) in zip(excel_sheets,excel_sheets.keys()):
            excel_sheets[sheet].dropna(axis=0, how='all', inplace=True)
            excel_sheets[sheet].reset_index(drop=True, inplace=True)
        
        # Dropping sheets with less than 5 rows and 4 columns (Not having enough data)
        excel_sheet_list = []
        for (sheet,sheet_name) in zip(excel_sheets,excel_sheets.keys()):
            if len(excel_sheets[sheet]) >= 5 and excel_sheets[sheet].shape[1] >= 4:
                excel_sheet_list.append(excel_sheets[sheet].copy())

        # Parsing the terminology
        terminology = None
        for i in range(len(excel_sheet_list)):
            while True:
                row_1 = excel_sheet_list[i].iloc[0]
                row_1_concatenated = ' '.join(row_1.dropna().astype(str))
                words_count = len(row_1_concatenated.split())
                if words_count > 15 and row_1.dropna().count() < 4:
                    excel_sheet_list[i].drop(labels =0,inplace=True)
                    excel_sheet_list[i].reset_index(drop=True,inplace=True)
                    continue
                
                if "balance" in row_1_concatenated.lower() and (("description" in row_1_concatenated.lower() or "particular" in row_1_concatenated.lower() or "narration" in row_1_concatenated.lower() or "remarks" in row_1_concatenated.lower()) or ("date" in row_1_concatenated.lower())):
                    if "date" not in row_1_concatenated.lower() and "date" in ' '.join(excel_sheet_list[i].iloc[1].dropna().astype(str)).lower():
                        excel_sheet_list[i].iloc[0] = excel_sheet_list[i].iloc[0].astype(str) + ' ' + excel_sheet_list[i].iloc[1].astype(str)
                        excel_sheet_list[i].drop(labels =1,inplace=True)
                        excel_sheet_list[i].reset_index(drop=True, inplace=True)
                    terminology = excel_sheet_list[i].iloc[0]
                    break
                else :
                    excel_sheet_list[i].drop(labels =0,inplace=True)
                    excel_sheet_list[i].reset_index(drop=True,inplace=True)

            if len(excel_sheet_list[i].iloc[0].dropna()) < 3:
                terminology = None
            if terminology is not None:
                print(f"Terminology found in sheet {i+1}")
                break
        
        # Dropping the rows which are not recording transactions
        for i in range(len(excel_sheet_list)):
            diff = excel_sheet_list[i].apply(lambda x : diff_first_last_non_nan(x),axis=1)
            excel_sheet_list[i] = excel_sheet_list[i][diff>2].copy()
            
            for j, row in excel_sheet_list[i].iterrows():
                if  row.dropna().apply(lambda x: is_pure_string(x)).all():
                    excel_sheet_list[i].drop(labels = j,inplace=True)
            excel_sheet_list[i].reset_index(drop=True, inplace=True)

        for i in range(len(excel_sheet_list)):
            date_cell_prev = None
            if excel_sheet_list[i].shape[0] == 0:
                continue
            for j, row in excel_sheet_list[i].iterrows():
                date_cell = None
                for k, cell in enumerate(row):
                    if check_date(cell):
                        date_cell = k
                        break
                if date_cell is None:
                    excel_sheet_list[i].drop(labels = j,inplace=True)
                    continue
                if date_cell_prev is not None:
                    if date_cell != date_cell_prev:
                        excel_sheet_list[i].drop(labels = j,inplace=True)
                        continue
                else :
                    date_cell_prev = date_cell
            
            # drop only those column which are not there in terminology
            excel_sheet_list[i].dropna(axis=1, how='all', inplace=True)
            excel_sheet_list[i].reset_index(drop=True, inplace=True)
            excel_sheet_list[i].columns = range(excel_sheet_list[i].shape[1])   
            
            if excel_sheet_list[i].shape[0] == 0:
                excel_sheet_list.pop(i)         

        # Finding the column names in the terminology
        non_nan_column_indexes = len(terminology_preprocessing(terminology).dropna())

        # Finding the indexes of the columns in the terminology
        dates_idx,dates_num,balance_idx,description_idx,convention = column_index_processed(terminology,excel_sheet_list)

        # Cleaning the individual sheets
        for i in range(len(excel_sheet_list)):
            if excel_sheet_list[i].shape[1] < non_nan_column_indexes:
                # add a nan column jst after description column
                if excel_sheet_list[i][dates_idx[0]].dropna().apply(lambda x: check_date(x)).all():
                    pass
                elif dates_idx[0]>0 and excel_sheet_list[i][dates_idx[0]-1].dropna().apply(lambda x: check_date(x)).all():
                    excel_sheet_list[i].insert(dates_idx[0]-1, 'nan', np.nan)
                    excel_sheet_list[i].columns = range(excel_sheet_list[i].shape[1])
                    continue

                if excel_sheet_list[i][description_idx].apply(lambda x: is_pure_string(x)).all():
                    pass
                elif description_idx>0 and excel_sheet_list[i][description_idx-1].apply(lambda x: is_pure_string(x)).all(): # type: ignore
                    excel_sheet_list[i].insert(description_idx-1, 'nan', np.nan) # type: ignore
                    excel_sheet_list[i].columns = range(excel_sheet_list[i].shape[1])
                    continue


                excel_sheet_list[i].insert(description_idx+1, 'nan', np.nan) # type: ignore
                excel_sheet_list[i].columns = range(excel_sheet_list[i].shape[1])
            
            elif excel_sheet_list[i].shape[1] > non_nan_column_indexes:
                # add a nan column jst after description column
                difference = excel_sheet_list[i].shape[1] - non_nan_column_indexes
                column_to_delete = [description_idx+1+i for i in range(difference)] # type: ignore
                excel_sheet_list[i].drop(columns=column_to_delete,inplace=True) # type: ignore
                excel_sheet_list[i].columns = range(excel_sheet_list[i].shape[1])

        # Extracting the final tables
        final_extracted_sheets = []
        for i in range(len(excel_sheet_list)):
            try:
                extracted_sheet = excel_sheet_list[i].iloc[:,[dates_idx[0],description_idx,convention[1][0],convention[1][1],balance_idx]]

                final_extracted_sheets.append(extracted_sheet)
            except IndexError:
                pass
        
        # Processing the extracted tables
        last_sheets = extracted_sheet_processing(final_extracted_sheets,dates_idx,dates_num,balance_idx,description_idx,convention)
        
        # Concatenating the final tables
        final_df = pd.concat(last_sheets,ignore_index=True)

        print("Final table saved successfully")

        return final_df
    except :
        return None