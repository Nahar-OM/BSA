import os
import re
import numpy as np
import pandas as pd
import cv2
import subprocess
from functions import pdf_to_images,makeDirectories,save_images,read_image,get_ifsc
from OcrPdfToCsv import ImgToCsv

# Start the code -

if __name__ == "__main__" :
    pdf_path = r"C:\Users\Lenovo\OneDrive\Desktop\Folders\NaharOm\BSA\Main_Project\BankStatement\MVH SOLUTIONS\FY 2020-21 2021-22 2022-23.pdf"

    images_folder_path, processed_images_folder_path, excel_folder_path = makeDirectories(pdf_path)

    images = pdf_to_images(pdf_path,enhancement=False)

    images_path_list = save_images(images, images_folder_path)
    for image_path in images_path_list:
        csv_path_list = ImgToCsv(image_path, processed_images_folder_path, excel_folder_path)
    
    # # Sorted list of paths of excel files in excel folder -
    # excel_files_list = sorted(os.listdir(excel_folder_path))

    # heading = None

    # # Create a dataframe to store the data from all the excel files -
    # df = pd.DataFrame(columns=['Date','Narration','Withdrawal Amount','Deposit Amount','Closing Balance'])

    # # for i,excel_file in enumerate(excel_files_list):
    # #     # Read the excel file -
    #     path = os.path.join(excel_folder_path,excel_file)

    #     tranasaction_df 
    # command = [
    # "python",
    # "PaddleOCR/ppstructure/table/predict_table.py",
    # "--det_model_dir=PaddleOCR/ppstructure/inference/en_PP-OCRv3_det_infer",
    # "--rec_model_dir=PaddleOCR/ppstructure/inference/en_PP-OCRv3_rec_infer",
    # "--table_model_dir=PaddleOCR/ppstructure/inference/en_ppstructure_mobile_v2.0_SLANet_infer",
    # "--rec_char_dict_path=PaddleOCR/ppocr/utils/en_dict.txt",
    # "--table_char_dict_path=PaddleOCR/ppocr/utils/dict/table_structure_dict.txt",
    # f'--image_dir={processed_images_folder_path}',
    # f'--output={excel_folder_path}',
    # "--show_log=False"
    # ]   

    # subprocess.run(command)