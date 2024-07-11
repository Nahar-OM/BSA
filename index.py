import sys
from BSA_main import Main_BSA_Function, BSA_Report_Corrector

def print_flush(message):
    print(message, flush=True)

data_folder_path = r"/Users/oeuvars/Documents/ekarth-ventures/BSA/bank-statement/LANDCRAFT-RECREATIONS"
party_name = "LANDCRAFT RECREATIONS"

print_flush("Starting BSA process...")
print_flush("Extracting transactions...")
return_val = Main_BSA_Function(data_folder_path=data_folder_path, party_name=party_name)
print_flush(f"BSA process completed. Result: {return_val}")
