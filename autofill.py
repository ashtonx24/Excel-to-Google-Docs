import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import random

# Function to read data from an Excel file starting from a specific row
def read_data_from_excel(excel_path, start_row):
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active
    data = []
    for row in sheet.iter_rows(min_row=start_row, max_col=4, values_only=True):  # Start from the specified row
        data.append(row)
    return data

# Function to submit the form
def submit_google_form(data, form_url):
    driver = webdriver.Chrome()  # Make sure ChromeDriver is in your PATH
    driver.get(form_url)
    
    for entry in data:
        try:
            name, email, address, student_code = entry
            
            name_field = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input')
            email_field = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input')
            address_field = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[]/div[3]/div/div/div[2]/div/div[1]/div[2]/textarea')
            student_code_field = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input')
            #enter your actual xpaths from your google form
            name_field.send_keys(name)
            email_field.send_keys(email)
            address_field.send_keys(address)
            student_code_field.send_keys(student_code)
            
            submit_button = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[]/div[1]/div/span/span')
            submit_button.click()
            
            delay = random.randint(45, 80)
            time.sleep(delay)
            driver.get(form_url)  # Reload the form for the next entry
        
        except Exception as e:
            print(f"An error occurred: {e}")
    
    driver.quit()

# Main function
def main():
    excel_path = r'C:\Users\Admin\Downloads\students_data.xlsx'  # Replace with the path to your Excel file
    form_url = 'https://docs.google.com/forms/d/e/1FAIpQSfVBuhgQ2X6kRuH_5pZexdzU8z-MT8puPCkza9WggjbJ0mhqw/viewform?usp=sf_link'  # Replace with your Google Form URL
    start_row = int(input("Enter row number:"))  # Set the row number from which to start processing. Adjust this as needed.

    data = read_data_from_excel(excel_path, start_row)
    submit_google_form(data, form_url)

if __name__ == '__main__':
    main()
