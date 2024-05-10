from selenium.webdriver.chrome.service import Service
from time import sleep
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import openpyxl
from selenium.webdriver.common.keys import Keys
def open_Site():
    # Proxy details
    username = 'sppqd0zln8'
    password = 'q8ufInygku1q8U_1FX'
    proxy_host = 'gate.smartproxy.com'
    proxy_port = 7000

    
    
    service = Service(ChromeDriverManager().install())
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--no-sandbox")
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36"
    chrome_options.add_argument(f"--user-agent={user_agent}")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-geolocation")
    chrome_options.add_argument("--disable-features=Geolocation")
    proxy = f"{username}:{password}@{proxy_host}:{proxy_port}"
    chrome_options.add_argument(f"--proxy-server=http://{proxy}")
    #chrome_options.add_argument('--headless')
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.maximize_window()
    return driver


def scrapper():
    driver = open_Site()
    driver.get('https://onlinesearch.mns.mu/')
    sleep
    excel_file = 'ListofVATRegPersons.xlsx'
    df = pd.read_excel(excel_file)
    column_data = df.iloc[:, 2]
    column_data.pop(0)
    for index, column in enumerate(column_data):
        try:
            search_box = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@id="company-partnership-text-field"]')))
            search_box.send_keys(Keys.CONTROL + "a")  # Select all text in the input field
            search_box.send_keys(Keys.BACKSPACE)
            search_box.send_keys(column)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//button[@class="btn search-business-partnership-btn shadow"]'))).click()
            sleep(5)
            WebDriverWait(driver,20).until(EC.presence_of_element_located((By.XPATH, '//fa-icon[@title="View"]'))).click()
            sleep(2)
            try:
                financial_date = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, ' //mat-panel-title[text()="PROFIT AND LOSS STATEMENT"]/ancestor::mat-expansion-panel-header/following-sibling::div/div/div/div[1]/label[2]')))
                driver.execute_script("arguments[0].scrollIntoView(true);", financial_date)
                financial_date = financial_date.text
            except:
                financial_date = ''
            try:
                turnover = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//table//th[text()="Turnover"]/following-sibling::td')))
                driver.execute_script("arguments[0].scrollIntoView(true);", turnover)
                turnover = turnover.text
            except:
                turnover = ''
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//button[@class="dialog-close-button"]'))).click()
            row = index+3
            add_data_to_excel(excel_file, 5, row, turnover)
            add_data_to_excel(excel_file, 6, row, financial_date)
        except Exception as e:
            #print(e)
            pass

def add_data_to_excel(file_path, column_index, row_index, new_data):
    try:
        # Load the Excel workbook
        workbook = openpyxl.load_workbook(file_path)

        # Select the active worksheet
        sheet = workbook.active

        # Access the cell at the specified row and column and assign the new data
        sheet.cell(row=row_index, column=column_index).value = new_data

        # Save the changes to the Excel file
        workbook.save(file_path)
        print(f"Data added successfully.{new_data}")

    except Exception as e:
        print(f"An error occurred: {e}")

scrapper()


