# Importing the required libraries
from transaction_extractor import *
from report_generator import *

# Main function to extract the transactions and generate the report
def Main_BSA_Function(data_folder_path,party_name):

    # If the all_processed_csv_path.txt and statement_details.xlsx files are already present in the folder
    if os.path.exists(os.path.join(data_folder_path,"all_processed_csv_path.txt")) and os.path.exists(os.path.join(data_folder_path,"statement_details.xlsx")):
        pass
    else:
        #try:
        # Extracting the transactions from the pdf files
        all_processed_csv_path,all_statement_details = main_converter(data_folder_path)

        all_statement_details.to_excel(os.path.join(data_folder_path,"statement_details.xlsx"),index=False)
        with open(os.path.join(data_folder_path,"all_processed_csv_path.txt"),"w") as f:
            for path in all_processed_csv_path:
                f.write(path+'\n')
        # except:
        #     print("Error in converting the files")
        #     return

    # Generating report directory
    folder_path = os.path.join(data_folder_path,"report_files")
    try:
        os.makedirs(folder_path,exist_ok=True)
    except:
        pass

    # Generating the report
    #try:
        # If the all_processed_csv_path.txt and statement_details.xlsx files are already present in the folder
    if os.path.exists(os.path.join(data_folder_path,"all_processed_csv_path.txt")) and os.path.exists(os.path.join(data_folder_path,"statement_details.xlsx")):
        # read the paths from all_processed_csv_path.txt
        all_processed_csv_path = []
        with open(os.path.join(data_folder_path,"all_processed_csv_path.txt"),"r") as f:
            for line in f:
                all_processed_csv_path.append(line.strip())
        # read the statement_details.xlsx
        all_statement_details = pd.read_excel(os.path.join(data_folder_path,"statement_details.xlsx"))

        # Report generator function
        report_generator(all_processed_csv_path,all_statement_details,party_name,folder_path)
    else:
        report_generator(all_processed_csv_path,all_statement_details,party_name,folder_path)
    # except:
    #     print("Error in generating the report")
    #     return all_processed_csv_path,all_statement_details

    print("Report generated successfully in the folder report_files")
    return

# Correction function to correct the generated report
def BSA_Report_Corrector(data_folder_path,info_extracted_df_path,party_name):
    all_statement_details = pd.read_excel(os.path.join(data_folder_path,"statement_details.xlsx"))

    folder_path = os.path.join(data_folder_path,"report_files")

    try:
        # Report corrector function
        report_corrector(info_extracted_df_path,all_statement_details,party_name,folder_path)
    except:
       print("Error in generating the report")
       return

    print("Report generated successfully in the folder report_files")
    return
