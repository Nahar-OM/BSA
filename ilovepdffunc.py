import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
# import pyautogui
import argparse

def wait_for_download(download_folder):
    while True:
        for fname in os.listdir(download_folder):
            if fname.endswith('.crdownload'):
                break
        else:
            return

def pdf_downloader(pdf_path,download_folder):
    chrome_options = webdriver.ChromeOptions()
    prefs = {"download.default_directory" : download_folder}
    chrome_options.add_experimental_option("prefs", prefs)
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

    convert_button = driver.find_element(By.XPATH, '//*[@id="processTask"]/span').click()

    WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="download"]'))
    )
    while True:
        if len(os.listdir(download_folder)) == 0:
            continue
        else:
            if os.listdir(download_folder)[0].endswith('.crdownload'):
                continue
            else:
                break
    driver.quit()

