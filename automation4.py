import logging
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
import os
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from playwright.sync_api import sync_playwright
import pytesseract
from pdf2image import convert_from_path
import random
from dotenv import load_dotenv
import os
from lxml import html
import re
from datetime import datetime
import json
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from utiles import get_informations
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
import sys




logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
load_dotenv()

def currency_to_float(currency_str):
    return float(re.sub(r'[^\d.]', '', currency_str))
def parse_time(time_str):
    return datetime.strptime(time_str, '%I:%M %p')



def human_typing(element, text, delay=0.2):
    for char in text:
        element.send_keys(char)
        time.sleep(delay)
all_names = []
all_links = []

def extract_names_and_links(driver):

    html_doc = driver.page_source
    time.sleep(random.uniform(1, 2))
    soup = BeautifulSoup(html_doc, 'html.parser')
    
    # names = [th.find(class_='fc-resource-custom-name').text for th in soup.find_all('th', role='columnheader') if th.find(class_='fc-resource-custom-name')]
    # print(f'names:{names}')
    # time.sleep(random.uniform(1, 2))
    # print(len(names))
    # if len(names) == 0:
    #     print("No names found. Retrying...")
    #     extract_names_and_links(driver)
    # else:
    #     all_names.extend(names)

    links = [a.get('href') for a in soup.find_all('a', class_='fc-timegrid-event')]
    if len(links) == 0:
        print("No names found. Retrying...")
        extract_names_and_links(driver)
        return
    else:
        all_links.extend(links)
    print(f'links:{links}')
    all_links.extend(links)
    try:
        back_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="calendar-date__back-button"]'))
        )
        back_button.click()
        print("Button clicked successfully")
    except TimeoutException:
        print("Timed out waiting for the button to be clickable")
    
    time.sleep(random.uniform(2, 4)) 

def run(driver):
    try:
        logging.info("Sign-in successful")
        time.sleep(random.uniform(3,5))
        logging.info("Navigating to the archived reviews page")
        driver.get("https://pro.housecallpro.com/pro/calendar_new")
        current_day_of_week = datetime.now().weekday()
        iterations = current_day_of_week + 1

        for _ in range(iterations): 
            extract_names_and_links(driver)
            time.sleep(random.uniform(2,3))
        
        print("All links:", all_links)
        total_links = len(all_links)
        for index, link in enumerate(all_links, start=1):
            percent = (index / total_links) * 100
            bar_length = 50 
            filled_length = int(bar_length * index // total_links)
            bar = '>' * filled_length + '-' * (bar_length - filled_length)

            # Print progress
            sys.stdout.write(f"\rProcessing link {index}/{total_links}: [{bar}] {percent:.2f}%")
            sys.stdout.flush()
            full_url = "https://pro.housecallpro.com" + link  
            driver.get(full_url)
            print(full_url)
            try:
                WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, '//div/div/div/div/div/p[@data-hcp-mask-type="unmasked"]')))

                tree = html.fromstring(driver.page_source)
                customer_name_elements = WebDriverWait(driver, 50).until(EC.presence_of_all_elements_located((By.XPATH, '//div/div/div/div/div/div/div/div/div/div/div/div/div/div/p[@data-hcp-mask-type="unmasked"]')))
                english_letter_elements = [element for element in customer_name_elements if re.match(r'^[A-Za-z]', element.text)]
                info = []
                for row in customer_name_elements:
                    info.append(row.text)

                for data in english_letter_elements:
                    print(data)
                if customer_name_elements:
                    if english_letter_elements[0].text == 'Street view image not available.':

                        customer_name = english_letter_elements[1].text
                        print(f"Customer Name text: {english_letter_elements[1].text}")
                    else:
                        customer_name = english_letter_elements[0].text
                        print(f"Customer text: {english_letter_elements[0].text}")
                else:
                    print("Element not found")
                
                data_element = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//div/div/div/div/div/div/div/div/div/div/div/div/div[@data-hcp-mask-type="unmasked"]')))
                if data_element:
                    date_elements = [element for element in data_element if re.match(r'\w{3}, \w{3} \d{1,2} \'\d{2}', element.text)]
                    print(len(date_elements))
                    if len(date_elements)>1:
                        date = date_elements[1].text
                        print(f"Date: {date}")
                    else:
                        date = date_elements[0].text
                        print(f"Date: {date}")
                else:
                    print("No data elements found")

                job_customer_tags_elements = WebDriverWait(driver, 50).until(EC.presence_of_all_elements_located((By.XPATH, '//textarea[@data-testid="name-tag-chip"]//span')))
                for tag in job_customer_tags_elements:
                    print(tag.text)
                date_obj = datetime.strptime(date, "%a, %b %d '%y")
                formatted_date = date_obj.strftime("%Y-%m-%d")
                getinfo = get_informations(info)
                print(getinfo)
                what_have_done_text = ''
                what_have_done = WebDriverWait(driver, 50).until(EC.presence_of_all_elements_located((By.XPATH, '//textarea[@data-testid="name-input"]')))
                for row in what_have_done:
                    print(row.text)
                if len(what_have_done)>1:
                    what_have_done_text = what_have_done[0].text
                    material = what_have_done[1].text
                elif len(what_have_done)>0:
                    what_have_done_text = what_have_done[0].text
                    material = ''


                parsed_result = json.loads(getinfo)

                customer_name = parsed_result.get("Name")
                adress = parsed_result.get("Address")
                email = parsed_result.get('Email')
                phoneNumber = parsed_result.get('Phone Number')
                current_url = driver.current_url
                print(f"Current URL: {current_url}")
                service = ''
                data = [formatted_date, customer_name, current_url ,  adress, email, phoneNumber, what_have_done_text, material, service]
                print(data)
                # write_to_google_sheet_3(data, '1oK2JkSWbdya1zJdqLzEURQyH7miQyyEHlhWlO5lpvJU', "EmailMarketing")
        


            except Exception as e:
                print('Error:', e)
             
            except Exception as e:
                print('Error:', e)
            title = driver.title
            print(f"Title of {full_url}: {title}")    


        
    except NoSuchElementException:
        logging.error("Sign-in failed or element not found")

    finally:
        logging.info("Closing the browser")
        driver.quit()

EMAIL = os.getenv('EMAIL')
PASSWORD = os.getenv('PASSWORD')

def run_automtion():
    logging.info("Initializing the Chrome driver")

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

    # driver.fullscreen_window()
    driver.maximize_window()

    logging.info("Opening the Covidence sign-in page")
    driver.get("https://pro.housecallpro.com/")
    time.sleep(3)
    logging.info("Locating the email and password fields and sign-in button")
    try:
        email_field = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "email"))
        )
    except TimeoutException:
        print("Email field not found within the given time.")

    try:
        password_field = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "password"))
        )
    except TimeoutException:
        print("Password field not found within the given time.")

    logging.info("Entering credentials")
    human_typing(email_field, EMAIL)
    human_typing(password_field, PASSWORD)

    logging.info("Clicking the sign-in button")
    sign_in_button = driver.find_element(By.XPATH, "//button[contains(@class, 'MuiButtonBase-root') and contains(@class, 'MuiButton-root')]")
    sign_in_button.click()

    logging.info("Waiting for the page to load")
    driver.implicitly_wait(5)
    time.sleep(10)
    run(driver)


while True:
    run_automtion()  # Call your main function
    time.sleep(3600*4) 
