# Importing required libraries
import os
from PIL import Image, ImageEnhance
import numpy as np
import cv2
import pytesseract
import pandas as pd
from pdf2image import convert_from_path
import datetime
import time
import re
import json
# from dotenv import load_dotenv
from dateutil.parser import parse
import os
from datetime import datetime
# Need to change according to user's system
# Get the Tesseract executable path from the environment variable
tesseract_path = os.getenv('TESSERACT_PATH', 'tesseract')
pytesseract.pytesseract.tesseract_cmd = tesseract_path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait  # type: ignore
from selenium.webdriver.support import expected_conditions as EC
import time
import fitz

# List of all date formats
date_formats = [
        "%Y-%m-%d",  # "2022-11-28"
        "%d/%m/%Y",  # "11/01/2023"
        "%d-%b-%y",  # "01-FEB-23"
        "%d-%b-%Y",
        "%B %d %Y",
        "%Y-%m-%dT%H:%M:%S",  # ISO format with time
        "%Y-%m-%d",  # ISO format without time
        "%m/%d/%Y",  # Month/Day/Year
        "%m/%d/%y",  # Month/Day/Year (short year)
        "%d-%m-%Y",  # Day-Month-Year
        "%Y%m%d",  # Basic ISO format without separators
        "%d/%m/%Y",  # Day/Month/Year
        "%d/%m/%y",  # Day/Month/Year (short year)
        "%b %d, %Y",  # Month abbreviation Day, Year (e.g., Jan 01, 2023)
        "%B %d, %Y",  # Month full name Day, Year (e.g., January 01, 2023)
        "%d %b %Y",  # Day Month abbreviation Year (e.g., 01 Jan 2023)
        "%d %B %Y",  # Day Month full name Year (e.g., 01 January 2023)
        "%Y-%m-%dT%H:%M:%S.%f",  # ISO format with microseconds
        "%Y-%m-%dT%H:%M:%S.%fZ",  # ISO format with microseconds and Zulu timezone
        "%Y/%m/%d",  # Year/Month/Day
        "%Y.%m.%d",  # Year.Month.Day
        "%d.%m.%Y",  # Day.Month.Year
        "%Y.%m.%d %H:%M:%S",  # Year.Month.Day Hour:Minute:Second
        "%d.%m.%Y %H:%M:%S",  # Day.Month.Year Hour:Minute:Second
        "%d,%m,%Y",
        "%d,%m,%y",
    ]

# Function to parse a date string into a datetime object
def try_multiple_date_formats(date_string):
    try:
        date_string = parse(date_string, fuzzy=False)
    except:
        pass
    if isinstance(date_string, datetime):
        return date_string
    if date_string is None or date_string == '':
        return date_string
    for date_format in date_formats:
        try:
            #print(date_string)
            parsed_date = datetime.strptime(date_string.strip(), date_format)
            return parsed_date
        except ValueError:
            continue
        except AttributeError:
            continue
    # If no format matches, raise an exception or return None as needed.
    return np.nan

# Function to check if a string can be parsed into a date
def check_date_2(date_string):
    try:
        if date_string is None or date_string == '':
            return date_string
        for date_format in date_formats:
            try:
                parsed_date = datetime.strptime(date_string.strip(), date_format)
                return True
            except ValueError:
                continue
        return False
    except:
        return False

# Function to check if a string is a number
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

# Function to fix OCR errors in numbers
def fix_ocr_error(number_str):
    # Replace all periods with commas
    number_str = number_str.replace('.', ',')
    # Replace the last comma with a period
    number_str = number_str[::-1].replace(',', '.', 1)[::-1]
    # Remove all other commas
    number_str = number_str.replace(',', '')
    # Convert to float
    try:
        return float(number_str)
    except:
        return np.nan

# Function to check if a string can be parsed into a date
def is_date(string, fuzzy=False):
    try:
        parse(string, fuzzy=fuzzy)
        return True
    except ValueError:
        return False
    except OverflowError:
        return False

# Check if a string is a pure string (not numeric, not a date, and does not contain any digits)
def is_pure_string(s):
    try :
        return isinstance(s, str) and not is_number(s) and not is_date(s)
    except:
        return False

# Function to parse an amount string into a float
def amount_parser(x):
    if is_number(x):
        return float(x)

    # parse out only the numeric part from the string -
    num = re.sub(r'[^\d\.\-\+]','',x)
    list_str = x.split(" ")
    for str in list_str:
        if is_number(str):
            num = str
            break
    try:
        return float(num)
    except ValueError:
        return fix_ocr_error(num)

# Function to check if CR/DR is present in the string
def is_cr_or_dr(x):
    try:
        if "dr" in x.lower():
            return "dr"
        elif "cr" in x.lower():
            return "cr"
        else:
            return np.nan
    except:
        return np.nan

# Function to find if the pdf is a scanned document
def is_scanned(pdf_path):
    doc = fitz.open(pdf_path)
    coverage_list = []
    for page_num, page in enumerate(doc): # type: ignore
        total_page_area = abs(page.rect)
        text_area = 0.0
        for b in page.get_text("blocks"):
            r = fitz.Rect(b[:4])  # rectangle where block text appears
            text_area = text_area + abs(r)
        coverage_list.append(text_area/total_page_area)
    doc.close()
    coverage_list_sorted = coverage_list.copy()
    coverage_list_sorted.sort()
    if len(coverage_list_sorted) == 1:
        min_coverage = coverage_list_sorted[0]
    else:
        min_coverage = (coverage_list[0]+coverage_list[1])/2

    if min_coverage < 0.1:
        return True
    else:
        return False

# Function to make directories for storing the images and excel files
def makeDirectories(pdf_path,root_directory,images_needed):
    # find the folder name from the root directory by finding the text after the last backslash
    file_name = os.path.basename(pdf_path)
    file_name  = os.path.splitext(file_name)[0]
    dir_name = os.path.dirname(pdf_path)
    dir_name_last = os.path.basename(dir_name)
    folder_path = os.path.join(dir_name_last,file_name)
    #regex = r"BankStatement\\(.+)(?=\.pdf|\.xlsx|\.csv)"
    # here instead of 1
    #folder_path = re.search(regex, pdf_path).group(1) # type: ignore
    #print(folder_path)
    excel_folder_path = os.path.join(root_directory, folder_path, "ExcelData")
    if not os.path.exists(excel_folder_path):
        os.makedirs(excel_folder_path)

    images_folder_path, processed_images_folder_path = None, None
    if images_needed:
        images_folder_path = os.path.join(root_directory, folder_path, "Images")
        if not os.path.exists(images_folder_path):
            os.makedirs(images_folder_path)

        processed_images_folder_path = os.path.join(root_directory, folder_path, "ProcessedImages")
        if not os.path.exists(processed_images_folder_path):
            os.makedirs(processed_images_folder_path)

    return images_folder_path, processed_images_folder_path, excel_folder_path

# Function to convert a pdf to images
def pdf_to_images(pdf_path,enhancement=True,first_page=None,last_page=None):
    images = convert_from_path(pdf_path, 500, first_page=first_page, last_page=last_page)

    final_images = []
    for i, image in enumerate(images):
        if not enhancement:
            grayscaled_image = image.convert('L')
            final_images.append(np.array(grayscaled_image))
        else :
            grayscaled_image = image.convert('L')
            _,binary_image = cv2.threshold(np.array(grayscaled_image), 100, 255, cv2.THRESH_BINARY) # type: ignore
            denoised_image = cv2.GaussianBlur(binary_image, (3,3), 0) # type: ignore
            enhanced_image = Image.fromarray(denoised_image)

            enhancer = ImageEnhance.Contrast(enhanced_image)
            factor = 1.2

            enhanced_image = enhancer.enhance(factor)
            final_images.append(np.array(enhanced_image))

    return final_images

# Function to save images
def save_images(images,images_folder_path):
    path_lists = []
    for i, image in enumerate(images):
        image = Image.fromarray(image)
        image.save(f'{images_folder_path}/page_{(i+1):03d}.png', 'PNG')
        path_lists.append(f'{images_folder_path}\\page_{(i+1):03d}.png')
    return path_lists

# Function to read an image to text
def read_image(image):
    fulltext = pytesseract.image_to_string(image, lang='eng')
    return fulltext

# Function to detect any IFSC code in the text
def get_ifsc(text):
    regex = r"[A-Z]{4,}\d{6,7}"
    matches = re.findall(regex, text)
    return matches

# Function to classify the bank from the text
def classify_bank(text):
    '''
    Takes OCR text from first page and detects IFSC code, to infer the bank by using it
    '''
    highligted_keywords = {"SBI":"STATE BANK OF INDIA"}

    bank_keywords = {"YES": ["YES", "BANK"],
             "HDFC": ["HDFC", "BANK"],
             "ICIC": ["ICICI", "BANK"],
             "KKBK": ["KOTAK", "MAHINDRA", "BANK"],
             "UTIB": ["AXIS", "BANK"],
             "IDIB": ["INDIAN", "BANK"],
             "CNRB": ["CANARA", "BANK"],
             "SBIN": ["STATE BANK OF INDIA"],
             "IOBA": ["INDIAN", "OVERSEAS", "BANK"],
             "CITI": ["CITI", "BANK"],
             "ANDB": ["ANDHRA BANK"],
             "KVBL": ["KARUR", "VYSYA", "BANK"],
             "BARB": ["BANK OF BARODA"],
             "CBIN": ["CENTRAL", "BANK", "OF", "INDIA"],
             "UBIN": ["UNION", "BANK"],
             "PUNB": ["PUNJAB", "NATIONAL", "BANK"],
             "ORBC": ["ORIENTAL", "BANK", "OF", "COMMERCE"],
             "UCBA": ["UCO", "BANK"],
             "BKID": ["BANK OF INDIA"],
             "SBBJ": ["STATE", "BANK", "OF", "BIKANER", "AND", "JAIPUR"],
             "SBHY": ["STATE", "BANK", "OF", "HYDERABAD"],
             "SBM": ["STATE", "BANK", "OF", "MYSORE"],
             "SBP": ["STATE", "BANK", "OF", "PATIALA"],
             "SBT": ["STATE", "BANK", "OF", "TRAVANCORE"],
             "SYNB": ["SYNDICATE", "BANK"],
             "VIJB": ["VIJAYA", "BANK"],
             "IDFB": ["IDFC", "FIRST", "BANK"],
             "BAND": ["BANDHAN", "BANK"],
             "FBL": ["FEDERAL", "BANK"],
             "INDUS": ["INDUSIND", "BANK"],
             "RBL": ["RBL", "BANK"],
             "DBSS": ["DBS", "BANK"],
             "SCBL": ["STANDARD", "CHARTERED", "BANK"],
             "HSBC": ["HSBC", "BANK"],
             "DEUT": ["DEUTSCHE", "BANK"],
             "BARC": ["BANK", "OF", "AMERICA"],
             "SBI": ["STATE", "BANK", "OF", "INDIA"],
             "BOI": ["BANK", "OF", "INDIA"],
             "BOB": ["BANK", "OF", "BARODA"],
             "PNB": ["PUNJAB", "NATIONAL", "BANK"],
             "OBC": ["ORIENTAL", "BANK", "OF", "COMMERCE"]}


    banks = {"YES": "YES BANK", "HDFC": "HDFC BANK", "ICIC": "ICICI BANK", "KKBK": "KOTAK MAHINDRA BANK",
            "SBIN": "STATE BANK OF INDIA", "UTIB": "AXIS BANK", "IDIB": "INDIAN BANK", "CNRB": "CANARA BANK",
            "IOBA": "INDIAN OVERSEAS BANK", "CITI": "CITI BANK", "ANDB": "ANDHRA BANK", "BKID": "BANK OF INDIA",
            "BARB": "BANK OF BARODA", "CBIN": "CENTRAL BANK OF INDIA", "UBIN": "UNION BANK OF INDIA",
            "SBIN": "STATE BANK OF INDIA", "PUNB": "PUNJAB NATIONAL BANK", "ORBC": "ORIENTAL BANK OF COMMERCE",
            "UCBA": "UCO BANK", "SBBJ": "STATE BANK OF BIKANER AND JAIPUR", "SBHY": "STATE BANK OF HYDERABAD",
            "SBIN": "STATE BANK OF INDIA", "SBM": "STATE BANK OF MYSORE", "SBP": "STATE BANK OF PATIALA",
            "SBBJ": "STATE BANK OF BIKANER AND JAIPUR", "SBT": "STATE BANK OF TRAVANCORE", "SYNB": "SYNDICATE BANK",
            "VIJB": "VIJAYA BANK", "IDFB": "IDFC FIRST BANK", "BAND": "BANDHAN BANK", "FBL": "FEDERAL BANK",
            "INDUS": "INDUSIND BANK", "RBL": "RBL BANK", "DBSS": "DBS BANK", "SCBL": "STANDARD CHARTERED BANK",
            "HSBC": "HSBC BANK", "CITI": "CITI BANK", "DEUT": "DEUTSCHE BANK", "BARC": "BANK OF AMERICA",
            "SBI": "STATE BANK OF INDIA", "BOI": "BANK OF INDIA", "BOB": "BANK OF BARODA",
            "PNB": "PUNJAB NATIONAL BANK", "OBC": "ORIENTAL BANK OF COMMERCE","KVBL":"KARUR VYSYA BANK"}
    ifsc = -1
    bank = ""
    ifsc = get_ifsc(text)
    if ifsc != [] :
        ifsc = ifsc[0]

        bank = ""
        for j in banks.keys():
            if j in ifsc:
                bank = banks[j]
                break
    if bank == "" :
        for keyword in highligted_keywords.keys():
            if keyword.lower() in text.lower():
                bank = highligted_keywords[keyword]
                break

    if bank == "" :
        for (bank_id,keywords) in bank_keywords.items():
            for keyword in keywords:
                keys = True
                if keyword.lower() not in text.lower():
                    keys = False
                    break
            if keys:
                bank = banks[bank_id]
                break
    if bank == "":
        bank = "Not Found"
    return (ifsc, bank)

def bank_details(images):
    h, w = images[0].shape
    crop = images[0][:h//3,:]

    info = read_image(crop)
    ifsc, bank = classify_bank(info)
    return (info,ifsc, bank)

# Function to analyse the format/terminology of the transactions in the bank statement
def terminology_analyzer(terminology):
    values = list(terminology.astype(np.str_))
    date_idx = []
    date_num = 0
    balance_idx = None
    description_idx = []
    withdraw_idx, deposit_idx = None, None
    amount_idx = None
    description_synonyms = ["description", "particulars", "narration","remarks"]
    withdraw_synonyms = ["withdraw", "debit"]
    deposit_synonyms = ["deposit", "credit"]
    cr_dr = False
    cr_dr_idx = None
    convention = [None,None]
    for i,value in enumerate(values):
        if value is not np.nan :
            #print(value)
            if "dr" in value.lower() and "cr" in value.lower():
                cr_dr = True
                cr_dr_idx = i
                break

    if not cr_dr:
        for i,value in enumerate(values):
            if value is not np.nan :
                if "withdraw" in value.lower() or "debit" in value.lower():
                    withdraw_idx = i

                elif "deposit" in value.lower() or "credit" in value.lower():
                    deposit_idx = i

        convention = [cr_dr,[withdraw_idx,deposit_idx]]
    else:
        for i,value in enumerate(values):
            if value is not np.nan :
                if "amount" in value.lower():
                    amount_idx = i
        convention = [cr_dr,[cr_dr_idx,amount_idx]]

    if cr_dr==False and withdraw_idx==None and deposit_idx==None:
        for i,value in enumerate(values):
            if value is not np.nan :
                if "amount" in value.lower():
                    amount_idx = i
        convention = [True,[None,amount_idx]]

    for i, value in enumerate(values):
        if value is not np.nan :
            if "date" in value.lower():
                date_idx.append(i)
                date_num += 1
            elif "balance" in value.lower():
                balance_idx = i
            else :
                for synonym in description_synonyms:
                    if synonym in value.lower():
                        description_idx = i
                        break
    return date_idx, date_num, balance_idx, description_idx, convention

# Process the raw information of the transactions extracted from the bank statement
def terminology_preprocessing(terminology):
    terminology = terminology.astype(np.str_)
    terminology_new = terminology.copy()
    # remove spaces from each of the cells in the terminology -
    for i,column_name in enumerate(terminology):
        terminology_new[i] = column_name.replace(" ","")

    for i,column_name in enumerate(terminology):
        # remove all 'nan' from the column names -
        if "nan" in column_name.lower():
            terminology_new[i] = column_name.replace("nan","")

        # if only spaces or empty terminology[i] , drop it -

        if len(terminology_new[i].strip()) == 0:
            terminology_new[i] = np.nan

    return terminology_new.dropna()

# Function to check if a cell is a date
def check_date(cell):
    if isinstance(cell, datetime):
        return True
    else:
        # If the cell is a string, try to parse it as a date
        try:
            return check_date_2(cell)
        except :
            # If parsing fails, return NaN
            return False

# Extracting the column indics from the terminology
def column_index_processed(terminology,excel_sheet_list):
    terminology = terminology_preprocessing(terminology)
    dates_idx,dates_num,balance_idx,description_idx,convention=terminology_analyzer(terminology)
    if convention[0] and convention[1][0]==None:
        for i,row in excel_sheet_list[0].iterrows():
            if row.dropna().apply(lambda x: check_date(x)).any():
                break
        for i,val in enumerate(row):
            if "cr" in str(val).lower() or "dr" in str(val).lower():
                convention[1][0] = row.index[i]
    if convention[0] and not convention[1][0]:
        convention[1][0] = convention[1][1]
    return dates_idx,dates_num,balance_idx,description_idx,convention

# Function to check the span of non-NaN values in a row
def diff_first_last_non_nan(row):
    non_nan_indices = row.dropna().index
    if len(non_nan_indices) >= 2:
        return non_nan_indices[-1] - non_nan_indices[0]
    else:
        return 0

# Function to process the extracted transactions from the bank statement
def extracted_sheet_processing(final_extracted_sheets,dates_idx,dates_num,balance_idx,description_idx,convention):
    extracted_standardized_sheets = []
    if not convention[0]:
        for i in range(len(final_extracted_sheets)):
            extracted_sheet = final_extracted_sheets[i].copy()
            extracted_sheet.columns = ['Date','Description','Debit','Credit','Balance']
            extracted_sheet['Date'] = extracted_sheet['Date'].apply(lambda x : try_multiple_date_formats(x))
            # remove the rows where date values are not date values
            extracted_sheet = extracted_sheet[extracted_sheet['Date'].notna()]
            extracted_sheet['Balance'] = extracted_sheet['Balance'].apply(lambda x : amount_parser(x))
            extracted_sheet['Debit'] = extracted_sheet['Debit'].apply(lambda x : amount_parser(x))
            extracted_sheet['Credit'] = extracted_sheet['Credit'].apply(lambda x : amount_parser(x))
            extracted_sheet = extracted_sheet[extracted_sheet['Date'].notna()]
            extracted_standardized_sheets.append(extracted_sheet)
    else :
        for i in range(len(final_extracted_sheets)):
            extracted_sheet = final_extracted_sheets[i].copy()
            extracted_sheet.columns = ['Date','Description','CR_DR','Amount','Balance']
            extracted_sheet['Date'] = extracted_sheet['Date'].apply(lambda x : try_multiple_date_formats(x))
            # remove the rows where date values are not date values
            extracted_sheet = extracted_sheet[extracted_sheet['Date'].notna()]
            extracted_sheet['Balance'] = extracted_sheet['Balance'].apply(lambda x : amount_parser(x))
            extracted_sheet['CR_DR'] = extracted_sheet['CR_DR'].apply(lambda x : is_cr_or_dr(x))
            extracted_sheet['Amount'] = extracted_sheet['Amount'].apply(lambda x : amount_parser(x))

            # make a new sheet which has debit and credit, and values are in debit if it is dr add in debit column and nan in credit column and vice versa

            extracted_sheet['Debit'] = np.where(extracted_sheet['CR_DR'] == 'dr',extracted_sheet['Amount'],np.nan)
            extracted_sheet['Credit'] = np.where(extracted_sheet['CR_DR'] == 'cr',extracted_sheet['Amount'],np.nan)

            # delete the columns CR_DR and Amount and rearrange the columns
            extracted_sheet.drop(columns=['CR_DR','Amount'],inplace=True)
            extracted_sheet = extracted_sheet[['Date','Description','Debit','Credit','Balance']]
            # remove rows with no date values
            extracted_sheet = extracted_sheet[extracted_sheet['Date'].notna()]
            extracted_standardized_sheets.append(extracted_sheet)

    return extracted_standardized_sheets

# Function to check using ilovepdf if OCR is needed
def OCR_Needed(pdf_path):
    chrome_options = webdriver.ChromeOptions()
    # Set the default download directory
    chrome_options.add_argument("--headless")

    driver = webdriver.Chrome(options=chrome_options)

    url = "https://www.ilovepdf.com/pdf_to_excel"

    driver.get(url)
    driver.implicitly_wait(30)
    upload_button = driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')
    upload_button.send_keys(pdf_path)

    WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="processTask"]/span'))
    )
    time.sleep(10)
    # Check if the ocr is needed -
    ocr_needed = False

    # check if the div with class option__panel__content scanned has style value as display: none -
    scanned_div = driver.find_element(By.CLASS_NAME, 'option__panel__content.scanned')
    style = scanned_div.get_attribute('style')

    if style == 'display: none;':
        ocr_needed = False
    else:
        ocr_needed = True

    driver.quit()

    return ocr_needed

def account_name_processor(account_name):
    account_name_df = pd.DataFrame(account_name)

    account_name_df["next_start"] = account_name_df["start"].shift(-1)

    # check if next start is equal to end of current row

    account_name_df["next_start"] = account_name_df["next_start"].fillna(0)
    i = 0
    while i != len(account_name_df)-1:
        if account_name_df["next_start"].iloc[i] == account_name_df["end"].iloc[i]:
            account_name_df.loc[i, "word"] = account_name_df.loc[i, "word"] + account_name_df.loc[i+1, "word"] # type: ignore
            account_name_df.loc[i, "end"] = account_name_df.loc[i+1, "end"]
            account_name_df.loc[i, "next_start"] = account_name_df.loc[i+1, "next_start"]
            account_name_df = account_name_df.drop(i+1)
            account_name_df = account_name_df.reset_index(drop=True)
        else :
            i += 1
    account_name_df = account_name_df[account_name_df["score"] > 0.85]
    return account_name_df

# Function to find the type of account from the text
def find_account_type(text):
    text = text.lower()
    text = text.replace("\n"," ")
    range_of_words = 20
    if "saving" in text:
        index = text.index("saving")
        context = text[max(0,index-range_of_words):min(len(text),index+range_of_words)]

        if "account" in context or "a/c" in context or "ac" in context or "type" in context:
            return "Saving Account"

    if "current" in text:
        index = text.index("current")
        context = text[max(0,index-range_of_words):min(len(text),index+range_of_words)]
        if "account" in context or "a/c" in context or "ac" in context or "type" in context:
            return "Current Account"

    if "regular" in text:
        index = text.index("regular")
        context = text[max(0,index-range_of_words):min(len(text),index+range_of_words)]

        if "account" in context or "a/c" in context or "ac" in context or "type" in context:
            return "Regular Account"

    return "Unknown"

# Function to extract the statement period from the text
def get_statement_period(text):
        text = text.lower()
        text = text.replace("\n"," ")
        text = text.replace(":","")
        first_pattern = r"(\b\d{2}/\d{2}/\d{4}\b)\s+to\s+(\b\d{2}/\d{2}/\d{4}\b)"
        first_second_pattern = r"(\b\d{2}-\d{2}-\d{4}\b)\s+to\s+(\b\d{2}-\d{2}-\d{4}\b)"
        second_pattern = r"(\b\d{2}-\w{3}-\d{4}\b)\s+to\s+(\b\d{2}-\w{3}-\d{4}\b)"
        third_pattern = r"(\b\d{2}-\w{3,}-\d{4}\b)\s+to\s+(\b\d{2}-\w{3,}-\d{4}\b)"

        patterns = [first_pattern,first_second_pattern, second_pattern, third_pattern]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                from_date = match.group(1)
                to_date = match.group(2)
                try:
                    output_format = "%d-%b-%Y"
                    from_date = try_multiple_date_formats(from_date)
                    from_date = from_date.strftime(output_format) # type: ignore
                    to_date = try_multiple_date_formats(to_date)
                    to_date = to_date.strftime(output_format) # type: ignore
                    return {'from_date': from_date, 'to_date': to_date}
                except Exception as e:
                    print("Statement Period Error", e)
                finally:
                    return {'from_date': from_date, 'to_date': to_date}
            else:
                continue

        return {'from_date': None, 'to_date': None}

def pan_number(text):
    text = text.lower()
    text = text.replace("\n"," ")
    text = text.replace(":","")
    first_pattern = r"([a-z]{5}\d{4}[a-z]{1})"
    match = re.search(first_pattern, text, re.IGNORECASE)
    if match:
        # return in uppercase
        return match.group(1).upper()
    else:
        return None
# Function to find all the above details from the text
def BS_Info(input,is_text=False):
    if is_text:
        text = input
    else:
        text = read_image(input[0])

    text_lines = text.split("\n")

    i = None

    for j in range(len(text_lines)):
        lowered_text = text_lines[j].lower()
        if "balance" in lowered_text and ("particulars" in lowered_text or "description" in lowered_text or "narration" in lowered_text or "remark" in lowered_text):
            i = j
            break

    if i is not None:
        text = "\n".join(text_lines[:i])

    ifsc, bank = classify_bank(text)

    account_type = find_account_type(text)

    statement_period = get_statement_period(text)

    pan_no = pan_number(text)

    return (ifsc, bank, account_type, statement_period,pan_no)

# excel to text conversion
def excel_to_text(dataframe):
    text = ""
    for i in range(dataframe.shape[0]):
        for j in range(dataframe.shape[1]):
            # if not nan
            if dataframe.iloc[i,j] is not np.nan:
                text += str(dataframe.iloc[i,j]) + " "
        text += "\n"
    return text




