# Importing required libraries -
from calendar import c
import os
import re
import numpy as np
import pandas as pd
import cv2
import subprocess
from functions import BS_Info, excel_to_text, pdf_to_images,makeDirectories,save_images,read_image,get_ifsc,bank_details,is_scanned
from OcrPdfToCsv import ImgToCsv
import argparse
from excel_to_table import excel_to_table
from excel_to_table_OCR import excel_to_table_OCR
from ilovepdffunc import pdf_downloader
# Function to extract the transactions from the pdf files
def main_converter(files_folder_path):
    print("Starting Transaction Extraction ...")

    # Making a dataframe to store the statement details
    all_processed_csv_path = []
    all_statement_details = pd.DataFrame(columns=["IFSC","Bank","Account Type","Statement Start","Statement End","PAN No","File Path"])
    files = [os.path.join(files_folder_path, f) for f in os.listdir(files_folder_path) if os.path.isfile(os.path.join(files_folder_path, f))]

    # Getting the pdf and excel files
    pdf_files = [f for f in files if f.endswith('.pdf')]
    excel_files = [f for f in files if f.endswith('.csv') or f.endswith('.xlsx')]

    ocr_needed = False

    # List of banks for which OCR is needed
    OCR_Needed_List = ["KOTAK MAHINDRA BANK"]
    total_files = len(pdf_files) + len(excel_files)
    files_done = 0

    # Extracting transactions from the pdf files
    for pdf_file in pdf_files:
        # Extracting details such as IFSC, Bank, Account Type, Statement Period from the first page of the pdf file
        first_image = pdf_to_images(pdf_file,enhancement=False,first_page=1,last_page=1)
        (ifsc,bank, account_type, statement_period,pan_number) = BS_Info(first_image,is_text=False)

        if bank in OCR_Needed_List:
            ocr_needed = True

        # Checking if the pdf file is a scanned document 
        is_scanned_doc = is_scanned(pdf_file)
        if is_scanned_doc:
            ocr_needed = True
        
        # If OCR is not needed
        if not ocr_needed:
            # Making directories to store the excel files
            images_folder_path, processed_images_folder_path, excel_folder_path = makeDirectories(pdf_file,files_folder_path,images_needed=False)

            # Converting the pdf file to excel using the script ilovepdf.py
            print("Here")
            # command = [
            #     "python",
            #     "ilovepdf.py",
            #     "--pdf_path", pdf_file,
            #     "--download_folder", excel_folder_path
            # ]
            # subprocess.run(command)
            pdf_downloader(pdf_file,excel_folder_path)
            print(excel_folder_path)
            excel_file_path = os.path.join(excel_folder_path,os.listdir(excel_folder_path)[0])

            # Extracted dataframe from the excel file
            final_df = excel_to_table(excel_file_path,excel_folder_path)
            if final_df is None:
                continue

            # If the statement period is not present in the first page of the pdf file
            if statement_period['from_date'] == None :
                statement_period['from_date'] = str(final_df.iloc[0,0]) # type: ignore
                statement_period['to_date'] = str(final_df.iloc[-1,0]) # type: ignore

            # Saving the final table to a csv file
            final_excel_file_path = os.path.join(excel_folder_path,"final_table.csv")
            final_df.to_csv(final_excel_file_path,index=False)
            
            # Appending the final table path to the list
            all_processed_csv_path.append(final_excel_file_path)
            all_statement_details.loc[len(all_statement_details.index)] =  [ifsc,bank, account_type, statement_period['from_date'],statement_period['to_date'],pan_number, pdf_file]
            files_done += 1
            print(f"Files Done: {files_done}/{total_files}")
        
        # If OCR is needed
        else:

            # Making directories to store the images and excel files
            images_folder_path, processed_images_folder_path, excel_folder_path = makeDirectories(pdf_file,files_folder_path,images_needed=True)
            images = pdf_to_images(pdf_file,enhancement=False)
            images_path_list = save_images(images, images_folder_path)
            csv_path_list = []

            # OCR code to convert the images to csv files
            for image_path in images_path_list:
                csv_path_list_page = ImgToCsv(image_path, processed_images_folder_path, excel_folder_path)
                csv_path_list.extend(csv_path_list_page)

            # Combining the csv files to a single dataframe and extracting the transactions
            final_df = excel_to_table_OCR(excel_folder_path,excel_folder_path)
            if final_df is None:
                continue
            
            # If the statement period is not present in the first page of the pdf file
            if statement_period['from_date'] == None :
                statement_period['from_date'] = str(final_df.iloc[0,0]) # type: ignore
                statement_period['to_date'] = str(final_df.iloc[-1,0]) # type: ignore

            # Saving the final table to a csv file
            final_excel_file_path = os.path.join(excel_folder_path,"final_table.csv")
            final_df.to_csv(final_excel_file_path,index=False)
            
            # Appending the final table path to the list
            all_processed_csv_path.append(final_excel_file_path)
            all_statement_details.loc[len(all_statement_details.index)] =  [ifsc,bank, account_type, statement_period['from_date'],statement_period['to_date'],pan_number, pdf_file]
            
            files_done += 1
            print(f"Files Done: {files_done}/{total_files}")
    
    # Extracting transactions from the excel files
    for excel_file in excel_files:
        # Making directories to store the images and excel files
        images_folder_path, processed_images_folder_path, excel_folder_path = makeDirectories(excel_file,files_folder_path,images_needed=False)
        excel_sheets = pd.read_excel(excel_file, sheet_name=None,header=None)
        # Take the first sheet -
        first_sheet = list(excel_sheets.keys())[0]
        excel_sheet = excel_sheets[first_sheet] 
        first_page_text = excel_to_text(excel_sheet)
        
        # Extracting details such as IFSC, Bank, Account Type, Statement Period from the first page of the excel file
        (ifsc,bank, account_type, statement_period,pan_number) = BS_Info(first_page_text,is_text=True)

        final_df = excel_to_table(excel_file,excel_folder_path)
        if final_df is None:
                continue
        
        # If the statement period is not present in the first page of the pdf file
        if statement_period['from_date'] == None :
            statement_period['from_date'] = str(final_df.iloc[0,0]) # type: ignore
            statement_period['to_date'] = str(final_df.iloc[-1,0]) # type: ignore

        # Saving the final table to a csv file
        final_excel_file_path = os.path.join(excel_folder_path,"final_table.csv")
        final_df.to_csv(final_excel_file_path,index=False)
        
        # Appending the final table path to the list
        all_processed_csv_path.append(final_excel_file_path)
        all_statement_details.loc[len(all_statement_details.index)] =  [ifsc,bank, account_type, statement_period['from_date'],statement_period['to_date'],pan_number, pdf_file]

        files_done += 1
        print(f"Files Done: {files_done}/{total_files}")

    # Returning the list of csv files and the statement details dataframe
    print("Transaction Extraction done !")
    return all_processed_csv_path,all_statement_details
    

            