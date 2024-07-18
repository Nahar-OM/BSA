import sys
import os
from BSA_main import Main_BSA_Function
import json

def print_flush(message):
    print(message, flush=True)

try:
    data_folder_name = sys.argv[1]
    base_path = "/Users/oeuvars/Documents/ekarth-ventures/BSA/ML/bank-statement"
    data_folder_path = os.path.join(base_path, data_folder_name)
except IndexError:
    print_flush("Error: Folder name not provided")
    sys.exit(1)

party_name = "LANDCRAFT RECREATIONS"

print_flush(f"Using data folder path: {data_folder_path}")
print_flush("Starting BSA process...")
print_flush("Extracting transactions...")
return_val = Main_BSA_Function(data_folder_path=data_folder_path, party_name=party_name)
print_flush(f"BSA process completed. Result: {return_val}")
