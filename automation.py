            # Print progress
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
from utiles import Warranty_Retail, write_to_google_sheet
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
    time.sleep(random.uniform(2, 4))
    soup = BeautifulSoup(html_doc, 'html.parser')   


    links = [a.get('href') for a in soup.find_all('a', class_='fc-timegrid-event')]
    if len(links) == 0:
        print("No names found. Retrying...")
        extract_names_and_links(driver)
        return
    else:
        all_links.extend(links)
    print(f'links:{links}')
    print(links)
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
        time.sleep(random.uniform(5,10))
        logging.info("Navigating to the archived reviews page")
        driver.get("https://pro.housecallpro.com/pro/calendar_new")
        current_day_of_week = datetime.now().weekday()
        iterations = current_day_of_week + 1

        for _ in range(iterations): 
            extract_names_and_links(driver)
        print("All names:", all_names)
        print("All links:", all_links)
        total_links = len(all_links)
        for index, link in enumerate(all_links, start=1):
            percent = (index / total_links) * 100
            bar_length = 50  # Length of the progress bar
            filled_length = int(bar_length * index // total_links)
            bar = '>' * filled_length + '-' * (bar_length - filled_length)

            # Print progress
            sys.stdout.write(f"\rProcessing link {index}/{total_links}: [{bar}] {percent:.2f}%")
            sys.stdout.flush()
            full_url = "https://pro.housecallpro.com" + link  
            driver.get(full_url)
            print(full_url)
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//div/div/div/div/div/p[@data-hcp-mask-type="unmasked"]')))

                tree = html.fromstring(driver.page_source)
                customer_name_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//div/div/div/div/div/div/div/div/div/div/div/div/div/div/p[@data-hcp-mask-type="unmasked"]')))
                english_letter_elements = [element for element in customer_name_elements if re.match(r'^[A-Za-z]', element.text)]
                if customer_name_elements:
                    if english_letter_elements[0].text == 'Street view image not available.':
                        customer_name = english_letter_elements[1].text
                        print(f"Element text: {english_letter_elements[1].text}")
                    else:
                        customer_name = english_letter_elements[0].text
                        print(f"Element text: {english_letter_elements[0].text}")
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

                omw_time = ''
                clickin_time = ''
                clickout_time = ''
                time_to_clickin = ''
                time_to_clickout = ''
                subtotal2 = 0
                total2 = 0
                discount2 = 0
                total_cost1 = 0
                total_cost2  = 0

                OMW_element = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//div/div/div/div/div/div/div/div/div/div/div/div/div[@data-hcp-mask-type="unmasked"]')))
                if OMW_element:
                    time_elements = [element for element in OMW_element if re.match(r'\d{1,2}:\d{2} (am|pm)', element.text)]
                    if len(time_elements) > 1 and time_elements[1].text:
                        omw = time_elements[1].text
                        omw_time = parse_time(omw)
                    else:
                        omw = None

                    if len(time_elements) > 2 and time_elements[2].text:
                        clickin = time_elements[2].text
                        clickin_time = parse_time(clickin)
                        time_to_clickin = (clickin_time - omw_time).total_seconds() / 60
                    else:
                        clickin = None

                    if len(time_elements) > 3 and time_elements[3].text:
                        clickout = time_elements[3].text
                        clickout_time = parse_time(clickout)
                        time_to_clickout = (clickout_time - clickin_time).total_seconds() / 60    
                    else:
                        clickout = None                   
                else:
                    print("No OMW elements found")                
                
                print(f'OMW : {omw}')
                print(f'clickin : {clickin}')
                print(f'clickout : {clickout}')
                print(f'DRIVE TIME: {time_to_clickin} minutes')
                print(f'TIME SPENT ON JOB: {time_to_clickout} minutes')
                image_elements = []
                try:
                    image_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//div/ul/li/div/a/img')))
                    job_pics = len(image_elements)
                except TimeoutException:
                    job_pics = 0
                
                print(f'JOB PICS : {len(image_elements)}')

                try:
                    name_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//div/div/div/div/div/div/div/div/div/div/div/div/div/div[@data-testid="assignedEmployee"]/span')))
                    for ele in name_elements:
                        if ele.text: 
                            name = ele.text  
                            print(f"First character of element text: {name}")
                            break
                        else:
                            print("Element text is empty")
                except TimeoutException:
                    name = ''
                try:
                    note_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//div/textarea[@data-testid="name-input"]')))
                    for ele in note_elements:
                        if ele.text: 
                            job_note = ele.text  
                            print(f"element text: {job_note}")
                            break
                        else:
                            job_note = ''
                except TimeoutException:
                    job_note = ''
                try:   
                    company_elements = WebDriverWait(driver, 10).until(
                        EC.presence_of_all_elements_located((By.XPATH, '//div[@role="button" and @data-testid="tag-chip"]/span'))
                    )
                    if company_elements:
                        company_info = company_elements[-1].text
                    # for ele in company_elements:
                    #     if ele.text: 
                    #         company_info = ele.text 
                            
                    #         break
                    #     else:
                    #         company_info =''
                except TimeoutException:
                    company_info = ''
                GREEN = '\033[92m'
                # ANSI escape code to reset color
                RESET = '\033[0m'

                print(f"{GREEN}Company text: {company_info}{RESET}")
                warranty_retail_result = Warranty_Retail(company_info)
                print(warranty_retail_result)
                parsed_result = json.loads(warranty_retail_result)

                warranty_and_retail = parsed_result.get("Warranty and retail")
                warranty_company = parsed_result.get("Warranty company")

                print(f"Warranty and Retail: {warranty_and_retail}")
                print(f"Warranty Company: {warranty_company}")

                try:
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//div[@data-testid="AppointmentPaymentStatusStep__Description"]'))
                    )
                    payment_status = element.text
                    print(f"Payment Status: {payment_status}")
                except TimeoutException:
                    payment_status = ''
                    print("Element not found or not visible within the timeout period")
                if payment_status == 'Paid':
                    pay = 10
                elif payment_status:
                    pay = 1
                else: pay =0

                print(pay)
                                            

                date_obj = datetime.strptime(date, "%a, %b %d '%y")
                formatted_date = date_obj.strftime("%Y-%m-%d")
                print(formatted_date)

                if time_to_clickin:
                    time_to_clickin= int(round(time_to_clickin))
                else: time_to_clickin =0
                if time_to_clickout:
                    time_to_clickout= int(round(time_to_clickout))
                else: time_to_clickout =0
                if job_note: 
                    job_note_word_count = len(job_note.split())
                else: 
                    job_note_word_count = 0
                

                try:

                    next_seg_buttons = driver.find_elements(By.XPATH, '//button[@role="tab"]')
                    if len(next_seg_buttons) > 1:
                        next_seg_buttons[1].click()  
                        time.sleep(3) 
                        
                    else:
                        print("No button with the specified attributes found.")
                    
                    try:
                        div_with_subtotal1 = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//div[p[contains(text(), 'Subtotal')]]"))
                        )
                        print("Found div with 'Subtotal 1':", div_with_subtotal1.text)

                        next_div1 = div_with_subtotal1.find_element(By.XPATH, "following-sibling::div")
                        print("Text of the next <div> element:", next_div1.text)
                        subtotal1 = next_div1.text
                        subtotal1 = currency_to_float(subtotal1)

                        div_with_total = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//div[p[contains(text(), 'Total')]]"))
                        )
                        print("Found div with 'Total1':", div_with_total.text)
                        next_total_div1 = div_with_total.find_element(By.XPATH, "following-sibling::div")
                        print("Text of the next <div> element:", next_total_div1.text)
                        total1 = next_total_div1.text
                        total1 = currency_to_float(total1)
                        try:
                            cost_elements = driver.find_elements(By.XPATH, '//input[@data-testid="autosave-unit-cost-input"]')
                            total_cost2 = sum(currency_to_float(element.get_attribute('value')) for element in cost_elements)
                            print(f"Total Cost2: {total_cost2}")
                        except Exception as e:
                            print("Error calculating total cost:", e)

                        discount1 = subtotal1 - total1
                        print(f"Subtotal1: {subtotal1}, Total: {total1}, Discount: {discount1}")

                    except Exception as e:
                        print("Error while finding subtotal or total:", e)

                    if len(next_seg_buttons) > 2:

                        next_seg_buttons[-1].click() 
                        time.sleep(1)
                        if index%2==0:
                            x = 100  
                            y = 100
                        else:
                            x=200
                            y=200

                        try: 
                            actions = ActionChains(driver)
                            actions.move_by_offset(x, y).click().perform()
                        except Exception as e:
                            print("Error while clicking the button:")

                        time.sleep(1)
                        try:
                            div_with_subtotal2 = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//div[p[contains(text(), 'Subtotal')]]"))
                            )
                            print("Found div with 'Subtotal 2':", div_with_subtotal2.text)

                            next_div2 = div_with_subtotal2.find_element(By.XPATH, "following-sibling::div")
                            print("Text of the next <div> element:", next_div2.text)
                            subtotal2 = next_div2.text
                            subtotal2 = currency_to_float(subtotal2)

                            div_with_total2 = WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//div[p[contains(text(), 'Total')]]"))
                            )
                            print("Found div with 'Total':", div_with_total2.text)
                            next_total_div2 = div_with_total2.find_element(By.XPATH, "following-sibling::div")
                            print("Text of the next <div> element:", next_total_div2.text)
                            total2 = next_total_div2.text
                            total2 = currency_to_float(total2)

                            discount2 = subtotal2 - total2
                            print(f"Subtotal2: {subtotal2}, Total: {total2}, Discount: {discount2}")

                        except Exception as e:
                            print("Error while finding subtotal or total for the second segment:", e)
                    else:
                        print("No button with the specified attributes found.")

                    time.sleep(3)                                
                    
                    try:
                        cost_elements = driver.find_elements(By.XPATH, '//input[@data-testid="autosave-unit-cost-input"]')
                        total_cost2 = sum(currency_to_float(element.get_attribute('value')) for element in cost_elements)
                        print(f"Total Cost2: {total_cost2}")
                    except Exception as e:
                        print("Error calculating total cost:", e)

                except NoSuchElementException:
                    print("No segment button found")
                total_price = total1 + total2
                total_unit_costs = total_cost1+ total_cost2
                discount = discount1 + discount2           
            
                data = [formatted_date,customer_name,full_url , ' ' , omw,  clickin, clickout, time_to_clickin, time_to_clickout, ' ', ' ', ' ', ' ',  job_pics, job_note_word_count, ' ', ' ', ' ', ' ',  ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ' ,' ']
                print(data)
                write_to_google_sheet(data, "1Wah4JVOkaiGRvsYO0QOpcTkrg9i59NrGr5LA5YD8CrU", name)
             
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
