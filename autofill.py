import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import random
import logging

# Configure logging
logging.basicConfig(filename="form_submission.log", level=logging.ERROR, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Function to read data from an Excel file starting from a specific row
def read_data_from_excel(excel_path, start_row):
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active
    data = []
    for row in sheet.iter_rows(min_row=start_row, max_col=4, values_only=True):
        if all(row):  # Ensures no empty fields
            data.append(row)
        else:
            logging.warning(f"Skipping incomplete row: {row}")
    return data

# Function to get last processed row
def get_last_row():
    try:
        with open("last_row.txt", "r") as file:
            return int(file.read().strip())
    except FileNotFoundError:
        return 2  # Default to row 2 if no previous record exists

# Function to update last processed row
def update_last_row(row):
    with open("last_row.txt", "w") as file:
        file.write(str(row))

# Function to submit Google Form
def submit_google_form(data, form_url):
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")  # Run in headless mode
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(form_url)
    wait = WebDriverWait(driver, 10)

    for entry in data:
        try:
            name, email, address, student_code = entry

            fields = {
                "name": 'your_xpath_here',
                "email": 'your_xpath_here',
                "address": 'your_xpath_here',
                "student_code": 'your_xpath_here'
            }

            for key, xpath in fields.items():
                field = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                field.send_keys(locals()[key])  # Dynamically get variable value

            submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, 'your_submit_button_xpath')))
            submit_button.click()

            time.sleep(random.randint(45, 80))  # Delay between form submissions
            driver.get(form_url)  # Reload form for next entry

        except Exception as e:
            logging.error(f"Error submitting form: {e}")
            print(f"Skipping entry due to error: {entry}")

    driver.quit()

# Main function
def main():
    excel_path = r'C:\Users\Admin\Downloads\students_data.xlsx'  # Update with your actual path
    form_url = 'your_google_form_url_here'
    
    start_row = get_last_row()
    data = read_data_from_excel(excel_path, start_row)

    if data:
        submit_google_form(data, form_url)
        update_last_row(start_row + len(data))
    else:
        print("No valid data found to process.")

if __name__ == '__main__':
    main()
