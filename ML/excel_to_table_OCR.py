import os
import numpy as np
import pandas as pd
from functions import *
from datetime import datetime
import argparse

def excel_to_table_OCR(excel_sheets_folder,download_folder):
    try:
        excel_files = [f for f in os.listdir(excel_sheets_folder) if f.endswith('.csv')]

        excel_files_list = [os.path.join(excel_sheets_folder, f) for f in excel_files]

        # excel_files_list = excel_files_list[2:] + excel_files_list[:2]
        excel_sheet_list = []
        for file in excel_files_list:
            sheet = pd.read_csv(file,header=None)
            sheet.dropna(axis=0, how='all', inplace=True)
            sheet.reset_index(drop=True, inplace=True)
            if len(sheet) >= 5 or sheet.shape[1] >= 4:
                excel_sheet_list.append(sheet)

        terminology = None
        for i in range(len(excel_sheet_list)):
            for j,row in excel_sheet_list[i].iterrows():
                row_1_concatenated = ' '.join(row.dropna().astype(str))

                if "balance" in row_1_concatenated.lower() and (("description" in row_1_concatenated.lower() or "particular" in row_1_concatenated.lower() or "narration" in row_1_concatenated.lower() or "remarks" in row_1_concatenated.lower()) or ("date" in row_1_concatenated.lower())):
                    if "date" not in row_1_concatenated.lower() and "date" in ' '.join(excel_sheet_list[i].iloc[j+1].dropna().astype(str)).lower():
                        excel_sheet_list[i].iloc[j] = excel_sheet_list[i].iloc[j].astype(str) + ' ' + excel_sheet_list[i].iloc[j+1].astype(str)
                        excel_sheet_list[i].drop(labels =j+1,inplace=True)
                        excel_sheet_list[i].reset_index(drop=True, inplace=True)
                    terminology = excel_sheet_list[i].iloc[j]
                    break

            if terminology is not None:
                print(f"Terminology found in sheet {i+1}")
                break

        for i in range(len(excel_sheet_list)):
            j = 0
            first_date_encountered = False
            while j < excel_sheet_list[i].shape[0]:
                is_date = False
                for k, cell in enumerate(excel_sheet_list[i].iloc[j].dropna()):
                    if check_date(cell):
                        if not first_date_encountered:
                            first_date_encountered = True

                        is_date = True
                        break

                if not is_date and first_date_encountered:
                    non_nan_idx = excel_sheet_list[i].iloc[j].dropna().index[0]
                    if excel_sheet_list[i].iloc[j].dropna().count() == 1 and is_pure_string(excel_sheet_list[i].iloc[j].dropna().values[0]) and is_pure_string(excel_sheet_list[i].iloc[j-1][non_nan_idx]):
                        excel_sheet_list[i].iloc[j-1,non_nan_idx] = excel_sheet_list[i].iloc[j-1,non_nan_idx] + ' ' + excel_sheet_list[i].iloc[j,non_nan_idx]
                        excel_sheet_list[i].drop(labels = j,inplace=True)
                        excel_sheet_list[i].reset_index(drop=True, inplace=True)
                    else:
                        excel_sheet_list[i].drop(labels = j,inplace=True)
                        excel_sheet_list[i].reset_index(drop=True, inplace=True)
                else :
                    j += 1

        for i in range(len(excel_sheet_list)):
            diff = excel_sheet_list[i].apply(lambda x : diff_first_last_non_nan(x),axis=1)
            excel_sheet_list[i] = excel_sheet_list[i][diff>2].copy()
            
            for j, row in excel_sheet_list[i].iterrows():
                if  row.dropna().apply(lambda x: is_pure_string(x)).all():
                    excel_sheet_list[i].drop(labels = j,inplace=True)
            excel_sheet_list[i].reset_index(drop=True, inplace=True)

        for i in range(len(excel_sheet_list)):
            date_cell_prev = None
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
            
            excel_sheet_list[i].dropna(axis=1, how='all', inplace=True)
            excel_sheet_list[i].reset_index(drop=True, inplace=True)
            excel_sheet_list[i].columns = range(excel_sheet_list[i].shape[1])     

        non_nan_column_indexes = len(terminology_preprocessing(terminology).dropna())
        dates_idx,dates_num,balance_idx,description_idx,convention = column_index_processed(terminology,excel_sheet_list)
        #print(dates_idx,dates_num,balance_idx,description_idx,convention)
        for i in range(len(excel_sheet_list)):
            for j, row in excel_sheet_list[i].iterrows():
                # drop row if balance is not present
                if np.isnan(amount_parser(row[balance_idx])):
                    excel_sheet_list[i].drop(labels = j,inplace=True)
            excel_sheet_list[i].reset_index(drop=True, inplace=True)

        for i in range(len(excel_sheet_list)):
            if excel_sheet_list[i].shape[0] == 0:
                continue
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
                # if there are zero values in balance, see if there is a column to right and assign it to balance
                difference = excel_sheet_list[i].shape[1] - non_nan_column_indexes
                column_to_delete = [description_idx+1+i for i in range(difference)] # type: ignore
                excel_sheet_list[i].drop(columns=column_to_delete,inplace=True) # type: ignore
                excel_sheet_list[i].columns = range(excel_sheet_list[i].shape[1])

        final_extracted_sheets = []
        for i in range(len(excel_sheet_list)):
            if excel_sheet_list[i].shape[0] == 0:
                continue
            extracted_sheet = excel_sheet_list[i].iloc[:,[dates_idx[0],description_idx,convention[1][0],convention[1][1],balance_idx]]

            final_extracted_sheets.append(extracted_sheet)
        last_sheets = extracted_sheet_processing(final_extracted_sheets,dates_idx,dates_num,balance_idx,description_idx,convention)
        # concatenate all the sheets -
        final_df = pd.concat(last_sheets,ignore_index=True)

        return final_df
    except:
        return None