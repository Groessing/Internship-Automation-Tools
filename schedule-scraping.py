from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
import time
from google.oauth2.service_account import Credentials
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import json
import gspread


driver = login_and_check("username", "password")


# Define the scope of the application — tells Google APIs what level of access is needed.
# In this case, we're requesting permission to read/write Google Sheets
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Load credentials from the service account JSON key file.
# This file contains the private key and other auth info generated in Google Cloud Console
creds = Credentials.from_service_account_file("key.json", scopes=scope)

# Authorize the gspread client with those credentials so we can access Google Sheets.
gsh = gspread.authorize(creds)
dialoger_spreadsheet_id = 'XXXXXX'
dialoger_sheet_name = 'Work schedules'

if driver:
    if navigate_to_arbeitszeitmodelle:
        data = extract_dynamic_tables(driver)

        if data:
            ws = gsh.open_by_key(dialoger_spreadsheet_id) # Opens CSV file by ID
            sh = ws.worksheet(dialoger_sheet_name) # Opens sheet by name
            values = [[row["Name"], row["ID"]] for row in data] # Extracts the "Name" and "ID" values from each row in the data
            sh.update('A3', values) # Starts at A3 and write the values into it (from A3;B3 to A_;B_)
            print("Exported to Google Sheets")
        
        else:
            print("No rows extracted.")
        
        driver.quit()
    
    else:
        print("Could not proceed without login.")

def send_error_message(error_message):
    with open("gmail.json") as g:
        app_password = g.read().strip()
    to_email = "email@domain.com"
    subject = "Error occured - at.xxxxxx.timetac.datascraping"
    body_text = str(error_message)
    from_email = "email@domain.com"

    message = MIMEMultipart()
    message["From"] = from_email
    message["To"] = to_email
    message["Subject"] = subject
    message.attach(MIMEText(body_text,"plain"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(from_email, app_password)
        server.send_message(message)

# Its enters a username and password, clicks on Login and checks if it's successful
def login_and_check(username_str, password_str):
    print("Starting login...")
    login_url = "https://go-sandbox.timetac.com/xxxxxx?auth"
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--window-size=1920,8048")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        driver.get(login_url)
        time.sleep(2)

        username_input = driver.find_element(By.ID, "userName")
        password_input = driver.find_element(By.ID, "userPass")

        username_input.send_keys(username_str)
        password_input.send_keys(password_str)

        login_button = driver.find_element(By.CLASS_NAME, "LoginSubmitForm")
        login_button.click()
        time.sleep(5)


        if "?auth" not in driver.current_url:
            print("Login successful")
            return driver
        
        else:
            print("Login failed")
            driver.quit()
            return None
    
    except Exception as e:
        print("Login error:", e)
        send_error_message(e)
        drive.quit()
        return None

# After Login, it is on dashboard and navigates to Einstellung > Arbeitszeitmodelle (Settings > Working Time Models)
def navigate_to_arbeitszeitmodelle():
    try:
        setting_button = driver.find_element(By.ID, "panel_left_settings-header")
        setting_button.click()
        wait = WebDriverWait(driver, 20)
        arbeitszeitmodell_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='Arbeitszeitmodelle']"))
        )
        arbeitszeitmodell_button.click()


        # It waits until "Hier können Sie verschiedene Arbeitszeitmodelle..." appears which is unique on the Arbeitszeitmodelle page
        wait.until(EC.presence_of_element_located((
            By.XPATH, "//*[contains(text(), 'Hier können Sie verschiedene Arbeitszeitmodelle')]"
        )))

        print("Arbeitszeitmodelle page loaded!")
        return True

    except TimeoutException as te:
        print("Could not load Arbeitszeitmodelle page")
        send_error_message(te)
        return False

def scroll_extjs_virtual(driver, container_selector="#gridview-1295", pause=1, max_attempts=100):
    container = driver.find_element(By.CSS_SELECTOR, container_selector)

    prev_last_name = None
    stable_count = 0

    for i in range(max_attempts):
        tables = container.find_elements(By.CSS_SELECTOR, "table[id^='gridview-']") # Searches for tables containing the id with ="gridview-"

        if not tables:
            print(f"No tables found on attempt {i+1}")
            break
        
        last_table = tables[-1]
        last_name = last_table.find_element(By.CSS_SELECTOR, "tbody > tr > td").text.strip() #Searches in each table and goes to tbody > tr > td

        if last_name = prev_last_name:
            stable_count += 1
            if stable_count >= 3:
                print("Last visible model stable for 3 cycles. Scrolling complete.")
                break
        else:
            stable_cont = 0
            prev_last_name = last_name #prev_last_name gets last_name, just to check they are not same
        
        # Scroll up by 190px (or container height) (it's javascript inside)
        driver.execute_script("""
            const container = arguments[0];
            container.scrollTop += container.clientHeight - 20; // slight overlap
        """, container)
        time.sleep(pause)
    return container

def extract_dynamic_tables(driver):
    wait = WebDriverWait(driver, 30)

    try:
        print("Waiting for table content to start loading...")

        # Wait for initial tables to appear in the container
        wait.until(lambda d: d.execute_script("""
            return document.querySelectorAll('#gridview-1295 table[id^="gridview-"]').length > 0;
        """))
        container = scroll_extjs_virtual(driver, "#gridview-1295", pause=5)

        tables = container.find_elements(By.CSS_SELECTOR, "table[id^='gridview-']")
        print(f"Found {len(tables)} tables.")

        results = []
        for table in tables:
            try:
                table_id = table.get_attribute("data-recordexternalid") 
                first_td = table.find_element(By.CSS_SELECTOR, "tdbody > tr > td") # goes to tbody > tr > td, inside here, there is an id

                text = frist_id.text.strip() # gets the text which is visible for the user (with the id above)
                results.append({
                    "Name": text,
                    "ID": table_id
                })
            except NoSuchElementException:
                continue
        
        return results
    
    except TimeoutException as te:
        send_error_message(te)
        print("Timed out waiting for table rows.")
        return []

                

