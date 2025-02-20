from flask import Blueprint, jsonify, request
import subprocess
import os
import logging
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import requests
import os 
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from PIL import Image
import pytesseract
import google.generativeai as genai
import pandas as pd
import re
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import shutil
api = Blueprint('routes', __name__)

@api.route('/run-seizure-451-bots', methods=['POST'])
def run_451_bots():
    try:
        # Extract username and password from request payload
        payload = request.get_json()
        if not payload or "customData" not in payload:
            return jsonify({
                "status": "error",
                "message": "Missing 'customData' in the request payload"
            }), 400

        custom_data = payload.get("customData", {})
        username = custom_data.get("username")
        password = custom_data.get("password")

        if not username or not password:
            return jsonify({
                "status": "error",
                "message": "Username and password are required in 'customData'"
            }), 400
        
        # chrome_driver_path = 'chromedriver.exe'

        # Initialize Chrome options
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36")
        chrome_options.add_argument("--incognito")

        # # Run in headless mode
        # chrome_options.add_argument("--headless")  # Add this line for headless mode
        # chrome_options.add_argument("--disable-gpu")  # This is necessary for headless mode on Windows
        # chrome_options.add_argument("--no-sandbox")  # For some environments, like CI/CD

        # chrome_options.add_argument(f"executable_path={chrome_driver_path}")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        # driver = webdriver.Chrome(options=chrome_options)

        # Configure logging
        LOG_FILE = "seizure_bot_451_errors.log"
        logging.basicConfig(
            filename=LOG_FILE,
            level=logging.ERROR,
            format="%(asctime)s - %(levelname)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )

        def log_error(error_message: str, exception: Exception):
            """
            Log an error with its message and exception details.
            """
            logging.error(f"{error_message}: {str(exception)}")
        try:
            # Open the website
            driver.get("https://ssl.jpclerkofcourt.us/JeffnetService/MCSearch/public_search/public_search_main.asp?PublicSearch=doc")
            driver.delete_all_cookies()
            driver.execute_script("window.localStorage.clear();")
            driver.execute_script("window.sessionStorage.clear();")

            # Login process
            try:
                wait = WebDriverWait(driver, 10)
                password_field = wait.until(EC.presence_of_element_located((By.ID, "txtPassword")))
                password_field.send_keys(password)

                login_field = wait.until(EC.presence_of_element_located((By.ID, "txtLogin")))
                login_field.send_keys(username, Keys.RETURN)
            except TimeoutException as e:
                log_error("Login fields not found", e)
                raise

            # Navigate to "Mortgage & Conveyance"
            try:
                mortgage_conveyance = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "div[onclick='runAccordion(1);']"))
                )
                mortgage_conveyance.click()

                doc_type = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.LINK_TEXT, "Doc Type"))
                )
                doc_type.click()
                print("Successfully navigated to 'Doc Type'.")
            except TimeoutException as e:
                log_error("Failed to navigate to 'Doc Type'", e)
                raise

            # Select dropdown option
            try:
                dropdown = Select(driver.find_element(By.ID, "cboDocType"))
                dropdown.select_by_value("451")
                print("Option 'Pre-foreclosure/Seizure' has been selected.")
            except NoSuchElementException as e:
                log_error("Dropdown not found or option selection failed", e)
                raise

            # Fill date fields
            try:
                def get_date_parts(date):
                    return date.strftime("%m"), date.strftime("%d"), date.strftime("%y")

                current_date = datetime.now()
                previous_date = current_date - timedelta(days=1)
                from_month, from_day, from_year = get_date_parts(previous_date)
                to_month, to_day, to_year = get_date_parts(current_date)
                # fields = [
                #     ("txtFromMonth", "12"),
                #     ("txtFromDay", "05"),
                #     ("txtFromYear", "24"),
                #     ("txtToMonth", "12"),
                #     ("txtToDay", "05"),
                #     ("txtToYear", "24"),
                # ]
                fields = [
                    ("txtFromMonth", from_month),
                    ("txtFromDay", from_day),
                    ("txtFromYear", from_year),
                    ("txtToMonth", to_month),
                    ("txtToDay", to_day),
                    ("txtToYear", to_year),
                ]
                for field_id, value in fields:
                    field = driver.find_element(By.ID, field_id)
                    field.clear()
                    field.send_keys(value)
            except Exception as e:
                log_error("Error filling date fields", e)
                raise

            # Submit form
            try:
                submit_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.ID, "cmdSubmit"))
                )
                submit_button.click()
                print("Form submitted successfully.")
            except Exception as e:
                log_error("Form submission failed", e)
                raise

            # Image downloading function
            def download_image(folder_path, counter):
                try:
                    image_element = driver.find_element(By.ID, "img_0")
                    image_url = image_element.get_attribute("src")
                    cookies = driver.get_cookies()
                    session = requests.Session()
                    for cookie in cookies:
                        session.cookies.set(cookie['name'], cookie['value'])
                    response = session.get(image_url)
                    if response.status_code == 200 and "image" in response.headers.get("Content-Type", ""):
                        image_name = os.path.join(folder_path, f"downloaded_image_{counter}.jpg")
                        with open(image_name, "wb") as file:
                            file.write(response.content)
                        print(f"Image successfully saved at: {os.path.abspath(image_name)}")
                    else:
                        log_error("Failed to download image", Exception(f"Status code: {response.status_code}"))
                except Exception as e:
                    log_error("Error in download_image function", e)

            # Handle search results and download images
            iframe = WebDriverWait(driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrmPublicSearchResults"))
            )
            ink_links = driver.find_elements(By.XPATH, "//a[contains(@href, 'javascript:fnViewJeffNetImage')]")
            if not ink_links:
                message=f"No data found for the given date: {previous_date} to {current_date} for Pre-foreclosure/Seizure"
                log_error(message, None)
            original_window = driver.current_window_handle
            counter = 1

            for index, ink_link in enumerate(ink_links):
                try:
                    ink_link.click()
                    driver.switch_to.window(driver.window_handles[1])
                    instrument_element = driver.find_element(By.XPATH, "//span[contains(text(), 'Instrument:')]")
                    instrument_text = instrument_element.text
                    valid_folder_name = instrument_text.replace(":", "_")
                    folder_path = f"C:\\Users\\Huraira Akbar\\Desktop\\Docs\\flask_seizure_bot_runner\\Seizure_data\\SEI_1(451)\\{valid_folder_name}"
                    os.makedirs(folder_path, exist_ok=True)
                    # Get total pages from divPageCount
                    wait = WebDriverWait(driver, 10)
                    page_count_element = wait.until(EC.presence_of_element_located((By.ID, "divPageCount")))

                    # Extract the text content of the span
                    page_count_text = page_count_element.text

                    # Print the extracted text
                    print("Page Count Text:", page_count_text)
                
                    total_pages = page_count_text  # Default to 1 if extraction fails
                    for page_num in range(1, int(total_pages)+ 1):
                        print(f"Processing page {page_num} of {total_pages}")
                        download_image(folder_path,counter)
                        counter += 1
                        print(counter)
                        next_page_button = driver.find_element(By.ID, "cmdNextPage")
                        next_page_button.click()
                        time.sleep(5)
                    # time.sleep(5)
                    driver.close()
                    # Wait for the page to load or perform actions
                    time.sleep(5)
                    print(f"Im here, clicked on link {index + 1}")
                    # Go back to the original window
                    driver.switch_to.window(driver.window_handles[0]) 
                    # Wait before continuing (optional)
                    time.sleep(3.5)
                    iframe = WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrmPublicSearchResults")))
                    if iframe:
                        print("TRue")
                except Exception as e:
                    log_error(f"Error handling link {index + 1}", e)

            logout_url = "https://ssl.jpclerkofcourt.us/JeffnetService/logout.asp"
            driver.get(logout_url)
            driver.quit()
            print("Logout action performed.")


            # Suppress GRPC logs
            os.environ["GRPC_VERBOSITY"] = "ERROR"
            os.environ["GRPC_LOG"] = "ERROR"

            # AI configuration
            api_key = "AIzaSyCNYWGVfq7ZYNoJOnf7HgbelmiHLVg10o8"
            if not api_key:
                raise ValueError("API key not provided. Please ensure your API key is set.")
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel("models/gemini-1.5-pro-001")

            prompt = (
                "Please summarize and extract the following key information from the provided text: "
                "Case No, Name, Property Address, Sum Owed, and Auction Date. "
                "The results should be displayed in a clean JSON format with each key-value pair representing the extracted information."
            )

            def ai_request_with_retries(prompt, max_retries=3):
                for attempt in range(max_retries):
                    try:
                        result = model.generate_content(prompt)
                        return result.text
                    except Exception as e:
                        log_error(f"AI request failed on attempt {attempt + 1}: {e}")
                        time.sleep(5)
                log_error("All retries for AI request failed.")
                return None

            def extract_text_from_image(image_path):
                """Extracts text from an image using Tesseract OCR."""
                try:
                    img = Image.open(image_path)
                    return pytesseract.image_to_string(img)
                except Exception as e:
                    log_error(f"Error extracting text from image {image_path}: {e}")
                    return ""

            def extract_key_value_pairs(text, keys):
                """Extracts key-value pairs using regex based on provided keys."""
                data = {}
                for key in keys:
                    match = re.search(rf"{key}:?\s*(.*?)\n", text, re.IGNORECASE)
                    data[key] = match.group(1).strip() if match else "Not mentioned"
                return data

            def clean_data(extracted_data):
                """Cleans up the extracted data to remove extraneous characters."""
                cleaned_data = {}
                for key, value in extracted_data.items():
                    cleaned_value = re.sub(r'[":,]+|null', '', value).strip()
                    cleaned_data[key] = cleaned_value if cleaned_value else "Not mentioned"
                return cleaned_data
            
            def send_to_zapier(file_path):
                try:
                    file_name = os.path.basename(file_path)
                    zapier_base_url = "https://hooks.zapier.com/hooks/catch/13403526/2sc5m31/"
                    zapier_url = f"{zapier_base_url}?file_name_get={file_name}"
                    with open(file_path, 'rb') as file:
                        files = {'file': file}
                        data = {
                            'file_name': file_name
                        }
                        response = requests.post(zapier_url, files=files, data=data)

                        if response.status_code == 200:
                            print(f"File successfully sent to Zapier: {file_path}")
                            folder_path = os.path.dirname(file_path)
                            del file_name
                            shutil.rmtree(folder_path)
                            print(f"Folder and all its contents successfully deleted: {folder_path}")
                        else:
                            log_error(f"Failed to send file to Zapier. Status Code: {response.status_code}, Response: {response.text}")
                except Exception as e:
                    log_error(f"Error sending file to Zapier: {e}")

            def process_folder(folder_path):
                """Processes all images in a folder: first OCR, then apply AI to combine results."""
                folder_name = os.path.basename(folder_path)
                ocr_text_all_images = ""

                print(f"Starting processing for folder: {folder_name}")
                for file in os.listdir(folder_path):
                    if file.lower().endswith(('.png', '.jpg', '.jpeg')):
                        image_path = os.path.join(folder_path, file)
                        print(f"Extracting text from image: {file}")
                        ocr_text = extract_text_from_image(image_path)
                        ocr_text_all_images += "\n" + ocr_text

                if not ocr_text_all_images.strip():
                    logging.warning(f"No text extracted from images in folder {folder_name}. Skipping AI processing.")
                    return None

                print(f"Sending OCR text for AI processing in folder: {folder_name}")
                ai_text = ai_request_with_retries(prompt + ocr_text_all_images)
                if not ai_text:
                    log_error(f"AI processing failed for folder {folder_name}.")
                    return None

                keys = ["Case No", "Name", "Property Address", "Sum Owed", "Auction Date"]
                extracted_data = extract_key_value_pairs(ai_text, keys)
                cleaned_data = clean_data(extracted_data)
                print(f"Processing complete for folder: {folder_name}")
                return cleaned_data

            base_dir = r"C:\Users\Huraira Akbar\Desktop\Docs\flask_seizure_bot_runner\Seizure_data\SEI_1(451)"

            """Processes all folders and subfolders for images."""
            overall_result = []
            original_filename = "SEI_1(Pre_foreclosure_Seizure).xlsx"
            current_date_for_file_name = datetime.now().strftime("%Y-%m-%d")
            new_filename = f"{current_date_for_file_name}_{original_filename}"
            output_file = os.path.join(base_dir, new_filename)
            try:
                for root, dirs, files in os.walk(base_dir):
                    if files:
                        folder_name = os.path.basename(root)
                        print(f"Processing folder: {folder_name}")
                        folder_result = process_folder(root)
                        if folder_result:
                            folder_result["Folder Name"] = folder_name
                            overall_result.append(folder_result)
                        else:
                            logging.warning(f"No valid data found in folder: {folder_name}")

                if overall_result:
                    print(overall_result)
                    df = pd.DataFrame(overall_result)
                    df.to_excel(output_file, index=False)
                    print(f"Results saved to {output_file}")
                        
                    # Send the generated file to Zapier
                    send_to_zapier(output_file)
                else:
                    logging.warning("No valid data found to save.")
            except Exception as e:
                log_error(f"Error processing folders: {e}")

            
        except Exception as e:
            log_error("Logout action failed", e)
        
        return jsonify({
            "status": "success",
            "message": "Seizure-451-bots executed successfully",
        })
    except subprocess.CalledProcessError as e:
        return jsonify({"error": str(e)}), 500


@api.route('/run-seizure-474-bots', methods=['POST'])
def run_474_bots():
    try:
        # Extract username and password from request payload
        payload = request.get_json()
        if not payload or "customData" not in payload:
            return jsonify({
                "status": "error",
                "message": "Missing 'customData' in the request payload"
            }), 400

        custom_data = payload.get("customData", {})
        username = custom_data.get("username")
        password = custom_data.get("password")

        if not username or not password:
            return jsonify({
                "status": "error",
                "message": "Username and password are required in 'customData'"
            }), 400
        
        # chrome_driver_path = 'chromedriver.exe'

        # Initialize Chrome options
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36")
        chrome_options.add_argument("--incognito")

        # # Run in headless mode
        # chrome_options.add_argument("--headless")  # Add this line for headless mode
        # chrome_options.add_argument("--disable-gpu")  # This is necessary for headless mode on Windows
        # chrome_options.add_argument("--no-sandbox")  # For some environments, like CI/CD

        # chrome_options.add_argument(f"executable_path={chrome_driver_path}")
        # driver = webdriver.Chrome(options=chrome_options)
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)


        # Configure logging
        LOG_FILE = "seizure_bot_474_errors.log"
        logging.basicConfig(
            filename=LOG_FILE,
            level=logging.ERROR,
            format="%(asctime)s - %(levelname)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )

        def log_error(error_message: str, exception: Exception):
            """
            Log an error with its message and exception details.
            """
            logging.error(f"{error_message}: {str(exception)}")
        try:
            # Open the website
            driver.get("https://ssl.jpclerkofcourt.us/JeffnetService/MCSearch/public_search/public_search_main.asp?PublicSearch=doc")
            driver.delete_all_cookies()
            driver.execute_script("window.localStorage.clear();")
            driver.execute_script("window.sessionStorage.clear();")

            # Login process
            try:
                wait = WebDriverWait(driver, 10)
                password_field = wait.until(EC.presence_of_element_located((By.ID, "txtPassword")))
                password_field.send_keys(password)

                login_field = wait.until(EC.presence_of_element_located((By.ID, "txtLogin")))
                login_field.send_keys(username, Keys.RETURN)
            except TimeoutException as e:
                log_error("Login fields not found", e)
                raise

            # Navigate to "Mortgage & Conveyance"
            try:
                mortgage_conveyance = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "div[onclick='runAccordion(1);']"))
                )
                mortgage_conveyance.click()

                doc_type = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.LINK_TEXT, "Doc Type"))
                )
                doc_type.click()
                print("Successfully navigated to 'Doc Type'.")
            except TimeoutException as e:
                log_error("Failed to navigate to 'Doc Type'", e)
                raise

            # Select dropdown option
            try:
                dropdown = Select(driver.find_element("id", "cboDocType"))
                dropdown.select_by_value("474")
                print("Option 'Probates/Succession' has been selected.")
            except NoSuchElementException as e:
                log_error("Dropdown not found or option selection failed", e)
                raise

            # Fill date fields
            try:
                def get_date_parts(date):
                    return date.strftime("%m"), date.strftime("%d"), date.strftime("%y")

                current_date = datetime.now()
                previous_date = current_date - timedelta(days=1)
                from_month, from_day, from_year = get_date_parts(previous_date)
                to_month, to_day, to_year = get_date_parts(current_date)
                # fields = [
                #     ("txtFromMonth", "12"),
                #     ("txtFromDay", "05"),
                #     ("txtFromYear", "24"),
                #     ("txtToMonth", "12"),
                #     ("txtToDay", "05"),
                #     ("txtToYear", "24"),
                # ]
                fields = [
                    ("txtFromMonth", from_month),
                    ("txtFromDay", from_day),
                    ("txtFromYear", from_year),
                    ("txtToMonth", to_month),
                    ("txtToDay", to_day),
                    ("txtToYear", to_year),
                ]
                for field_id, value in fields:
                    field = driver.find_element(By.ID, field_id)
                    field.clear()
                    field.send_keys(value)
            except Exception as e:
                log_error("Error filling date fields", e)
                raise

            # Submit form
            try:
                submit_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.ID, "cmdSubmit"))
                )
                submit_button.click()
                print("Form submitted successfully.")
            except Exception as e:
                log_error("Form submission failed", e)
                raise

            # Image downloading function
            def download_image(folder_path, counter):
                try:
                    image_element = driver.find_element(By.ID, "img_0")
                    image_url = image_element.get_attribute("src")
                    cookies = driver.get_cookies()
                    session = requests.Session()
                    for cookie in cookies:
                        session.cookies.set(cookie['name'], cookie['value'])
                    response = session.get(image_url)
                    if response.status_code == 200 and "image" in response.headers.get("Content-Type", ""):
                        image_name = os.path.join(folder_path, f"downloaded_image_{counter}.jpg")
                        with open(image_name, "wb") as file:
                            file.write(response.content)
                        print(f"Image successfully saved at: {os.path.abspath(image_name)}")
                    else:
                        log_error("Failed to download image", Exception(f"Status code: {response.status_code}"))
                except Exception as e:
                    log_error("Error in download_image function", e)

            # Handle search results and download images
            iframe = WebDriverWait(driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrmPublicSearchResults"))
            )
            ink_links = driver.find_elements(By.XPATH, "//a[contains(@href, 'javascript:fnViewJeffNetImage')]")
            if not ink_links:
                message=f"No data found for the given date: {previous_date} to {current_date} for Probates/Succession"
                log_error(message, None)
            original_window = driver.current_window_handle
            counter = 1

            for index, ink_link in enumerate(ink_links):
                try:
                    ink_link.click()
                    driver.switch_to.window(driver.window_handles[1])
                    instrument_element = driver.find_element(By.XPATH, "//span[contains(text(), 'Instrument:')]")
                    instrument_text = instrument_element.text
                    valid_folder_name = instrument_text.replace(":", "_")
                    folder_path = f"C:\\Users\\Huraira Akbar\\Desktop\\Docs\\flask_seizure_bot_runner\\Seizure_data\\SEI_2(474)\\{valid_folder_name}"
                    os.makedirs(folder_path, exist_ok=True)
                    # Get total pages from divPageCount
                    wait = WebDriverWait(driver, 10)
                    page_count_element = wait.until(EC.presence_of_element_located((By.ID, "divPageCount")))

                    # Extract the text content of the span
                    page_count_text = page_count_element.text

                    # Print the extracted text
                    print("Page Count Text:", page_count_text)
                
                    total_pages = page_count_text  # Default to 1 if extraction fails
                    for page_num in range(1, int(total_pages)+ 1):
                        print(f"Processing page {page_num} of {total_pages}")
                        download_image(folder_path,counter)
                        counter += 1
                        print(counter)
                        next_page_button = driver.find_element(By.ID, "cmdNextPage")
                        next_page_button.click()
                        time.sleep(5)
                    # time.sleep(5)
                    driver.close()
                    # Wait for the page to load or perform actions
                    time.sleep(5)
                    print(f"Im here, clicked on link {index + 1}")
                    # Go back to the original window
                    driver.switch_to.window(driver.window_handles[0]) 
                    # Wait before continuing (optional)
                    time.sleep(3.5)
                    iframe = WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "ifrmPublicSearchResults")))
                    if iframe:
                        print("TRue")
                except Exception as e:
                    log_error(f"Error handling link {index + 1}", e)

            logout_url = "https://ssl.jpclerkofcourt.us/JeffnetService/logout.asp"
            driver.get(logout_url)
            driver.quit()
            print("Logout action performed.")


            os.environ["GRPC_VERBOSITY"] = "ERROR"
            os.environ["GRPC_LOG"] = "ERROR"

            # Configure the Generative AI API
            api_key = "AIzaSyCNYWGVfq7ZYNoJOnf7HgbelmiHLVg10o8"
            if not api_key:
                raise ValueError("API key not provided. Please ensure your API key is set.")
            genai.configure(api_key=api_key)

            model = genai.GenerativeModel("models/gemini-1.5-pro-001")
            prompt = (
                "Please summarize and extract the following key information from the provided text: "
                "Decedent Name, Property Address, Representative, Representative's Address, Heir 1, Heir 2, and Heir 3. "
                "The results should be displayed in a clean JSON format with each key-value pair representing the extracted information."
            )

            # Retry mechanism for AI requests
            def ai_request_with_retries(prompt, max_retries=3):
                for attempt in range(max_retries):
                    try:
                        result = model.generate_content(prompt)
                        return result.text
                    except Exception as e:
                        log_error(f"AI request failed on attempt {attempt + 1}: {e}")
                        time.sleep(5)  # Wait before retrying
                log_error("All retries for AI request failed.")
                return None

            def extract_text_from_image(image_path):
                """Extracts text from an image using Tesseract OCR."""
                try:
                    img = Image.open(image_path)
                    return pytesseract.image_to_string(img)
                except Exception as e:
                    log_error(f"Error extracting text from image {image_path}: {e}")
                    return ""

            def extract_key_value_pairs(text, keys):
                """Extracts key-value pairs using regex based on provided keys."""
                data = {}
                for key in keys:
                    match = re.search(rf"{key}:?\s*(.*?)\n", text, re.IGNORECASE)
                    data[key] = match.group(1).strip() if match else "Not mentioned"
                return data

            def clean_data(extracted_data):
                """Cleans up the extracted data to remove extraneous characters."""
                cleaned_data = {}
                for key, value in extracted_data.items():
                    cleaned_value = re.sub(r'[":,]+|null', '', value).strip()
                    cleaned_data[key] = cleaned_value if cleaned_value else "Not mentioned"
                return cleaned_data

            def process_folder(folder_path):
                """Processes all images in a folder: first OCR, then apply AI to combine results."""
                folder_name = os.path.basename(folder_path) 
                ocr_text_all_images = "" 

                for file in os.listdir(folder_path):
                    if file.lower().endswith(('.png', '.jpg', '.jpeg')):
                        image_path = os.path.join(folder_path, file)
                        ocr_text = extract_text_from_image(image_path)
                        ocr_text_all_images += "\n" + ocr_text  

                if not ocr_text_all_images.strip():
                    logging.warning(f"No text extracted from images in folder {folder_path}. Skipping AI processing.")
                    return None

                ai_text = ai_request_with_retries(prompt + ocr_text_all_images)
                if not ai_text:
                    log_error(f"AI processing failed for folder {folder_path}.")
                    return None

                keys = [
                    "Decedent Name", "Property Address", "Representative",
                    "Representative's Address", "Heir 1", "Heir 2", "Heir 3", 
                    "Heir 4", "Heir 5"
                ]
                extracted_data = extract_key_value_pairs(ai_text, keys)
                return clean_data(extracted_data)
            
            def send_to_zapier(file_path):
                try:
                    file_name = os.path.basename(file_path)
                    zapier_base_url = "https://hooks.zapier.com/hooks/catch/13403526/2sc5m31/"
                    zapier_url = f"{zapier_base_url}?file_name_get={file_name}"
                    with open(file_path, 'rb') as file:
                        files = {'file': file}
                        data = {
                            'file_name': file_name
                        }
                        response = requests.post(zapier_url, files=files, data=data)

                        if response.status_code == 200:
                            print(f"File successfully sent to Zapier: {file_path}")
                            folder_path = os.path.dirname(file_path)
                            del file_name
                            shutil.rmtree(folder_path)
                            print(f"Folder and all its contents successfully deleted: {folder_path}")
                        else:
                            log_error(f"Failed to send file to Zapier. Status Code: {response.status_code}, Response: {response.text}")
                except Exception as e:
                    log_error(f"Error sending file to Zapier: {e}")


            base_dir = r"C:\Users\Huraira Akbar\Desktop\Docs\flask_seizure_bot_runner\Seizure_data\SEI_2(474)"

            """Processes all folders and subfolders for images."""
            overall_result = []
            original_filename = "SEI_2(Probates_Succession).xlsx"
            current_date_for_file_name = datetime.now().strftime("%Y-%m-%d")
            new_filename = f"{current_date_for_file_name}_{original_filename}"
            output_file = os.path.join(base_dir, new_filename)
            try:
                for root, dirs, files in os.walk(base_dir):
                    if files:
                        folder_name = os.path.basename(root)
                        print(f"Processing folder: {folder_name}")
                        folder_result = process_folder(root)
                        if folder_result:
                            folder_result["Folder Name"] = folder_name
                            overall_result.append(folder_result)
                        else:
                            logging.warning(f"No valid data found in folder: {folder_name}")

                if overall_result:
                    # print(overall_result)
                    df = pd.DataFrame(overall_result)
                    df.to_excel(output_file, index=False)
                    print(f"Results saved to {output_file}")
                        
                    # Send the generated file to Zapier
                    send_to_zapier(output_file)
                else:
                    logging.warning("No valid data found to save.")
            except Exception as e:
                log_error(f"Error processing folders: {e}")

            

                        
        except Exception as e:
            log_error("Logout action failed", e)
        
        return jsonify({
            "status": "success",
            "message": "Seizure-474-bots executed successfully",
        })
    except subprocess.CalledProcessError as e:
        return jsonify({"error": str(e)}), 500