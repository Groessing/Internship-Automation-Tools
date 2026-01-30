from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
import time
from time import sleep
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import json


driver = login_and_check("username", "password")
if driver:
    if navigate_to_arbeitszeitmodelle(driver):
        copy_arbeitszeitmodelle(driver)
        rename_arbeitszeitmodelle(driver)
        edit_arbeitszeitmodelle(driver)

        # This recieves country and working hours and searches for it
        # and prints out whether the working schedule exists or not
        country = os.environ.get('COUNTRY')
        working_hours = os.environ.get('WORKINGHOURS')
        print(f"Country: {country}")
        print(f"Working hours: {working_hours}")
        
        working_schedule = convert_to_string(driver, country, working_hours)
        search_work_schedule(driver, working_schedule)
        
        print("Success")
    



# It sends an error message to Florian Bolka
def send_error_message(error_message):
    with open("gmail.json") as g:
        app_password = g.read().strip()
    to_email = "email@domain.com"
    subject = "Error occured - at.xxxxxx.timetac.datascraping (Arbeitszeit Creator)"
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


# It puts country and working_hours together
# For example:
# Country: AT
# Working hours: 40
# Result: AT 40h / week

def convert_to_string(driver, country, working_hours):
    try:
        # If the working hours are < 10, then they need are 0 before (eg. 8 => 08)
        if float(working_hours) < 10:
            return f"{country} 0{working_hours}h / week"
        else:
            return f"{country} {working_hours}h / week"
    
    except Exception as e:
        send_error_message(e)
        return ""


# It searches for the given working schedule and
# prints out whether it was found or not
def search_work_schedule(driver, working_schedule):
    try:
        wait = WebDriverWait(driver, 30)
        wait.until(lambda d: d.execute_script("""
            return document.querySelectorAll('#gridview-1295 table[id^="gridview-"]').length > 0;
        """))

        container = driver.find_element(By.CSS_SELECTOR, "#gridview-1295")
        tables = container.find_elements(By.CSS_SELECTOR, "table[id^='gridview-']")
        found = False

        for table in tables:
            cell = table.find_element(By.CSS_SELECTOR, "tbody > tr > td")
            text = cell.text.strip()
            if text == working_schedule:
                found = True
                print(f"Working Schedule ({working_schedule}) found")
        if found == False:
            print(f"Error: Working Schedule ({working_schedule}) was not found")
    
    except Exception as e:
        send_error_message(e)


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

# It clicks to certains Arbeitszeitmodelle and clicks on "Vorlage kopieren" (copy template)
def copy_arbeitszeitmodelle(driver):
    wait = WebDriverWait(driver, 30)
    print("Waiting for table content to start loading...")

    # Wait for initial tables to appear in the container
    wait.until(lambda d: d.execute_script("""
        return document.querySelectorAll('#gridview-1295 table[id^="gridview-"]').length > 0;
    """))

    # List of Arbeitszeitmodelle names to find and copy
    target_names = [
        "AT 40h / week", "BG 40h / week", "HR 40h / week",
        "HU 40h / week", "PL 40h / week", "RO 40h / week",
        "UA 40h / week", "SK 40h / week", "UA 40h / week"
    ]

    # Loop through the list until all matching templates are found and copied
    while target_names:
        try:
            wait = WebDriverWait(driver, 5)
            name = target_names[0]
            container = driver.find_element(By.CSS_SELECTOR, "#gridview-1295")
            tables = container.find_elements(By.CSS_SELECTOR, "table[id^='gridview-']")
            matched = False

            # Iterate over all found tables to match by model name
            for table in tables:
                try:
                    cell = table.find_element(By.CSS_SELECTOR, "tbody > tr > td")
                    text = cell.text.strip()
                    if text == name:
                        cell.click()
                        copytemplate_button = driver.find_element(By.ID, "button-1302-btnWrap")
                        copytemplate_button.click()
                        print(f"Copied template for: {name}")
                        sleep(2)
                        matched = True
                        break

                # Handle case where DOM updated before access
                except StaleElementReferenceException:
                    print(f"Stale element while processing table for '{name}', skipping...")
                    continue
                    
                # Handle timeout waiting for button
                except TimeoutException:
                    print(f"Timeout waiting for button after clicking '{name}'")
                    continue
                    
            # Remove successfully processed name if it matched
            if matched:
                target_names.pop(0)
            else:
                print(f"{name} not found on page")
                break
        
        except StaleElementReferenceException as sere:
            print("Stale element at top level, retrying...")
            send_error_message(sere)
            sleep(2)
            continue



# Renames Arbeitszeitmodelle entries containing "_copy"
# by replacing "40h" with "5h" and removing the "_copy" suffix.
def rename_arbeitszeitmodelle(driver):

    # List of Arbeitszeitmodelle names to rename (those with "_copy")
    target_names = [
        "AT 40h / week_copy", "BG 40h / week_copy", "HR 40h / week_copy",
        "HU 40h / week_copy", "PL 40h / week_copy", "RO 40h / week_copy",
        "UA 40h / week_copy", "SK 40h / week_copy", "UA 40h / week_copy"
    ]

    # Iterate through each target name
    for name in target_names:
        success = False
        attempts = 3 # Retry attempts for handling dynamic page changes

        # Retry loop to handle possible stale or missing elements
        while attempts > 3 and not success:
            try:
                # Find the container holding the tables
                container = driver.find_elements(By.CSS_SELECTOR, "#gridview-1295")
                tables = container.find_elements(By.CSS_SELECTOR, "table[id^='gridview-']")

                # Loop through all tables to find the matching name
                for table in tables:
                    cell = table.find_element(By.CSS_SELECTOR, "tbody > tr > td")
                    text = cell.text.strip()

                    # If the current cell text matches the target name
                    if text == name:
                        actions = ActionChains(driver)

                        # Double-click the cell to enable editing
                        actions.double_click(cell).perform()
                        sleep(0.5) # Small delay to allow input to appear

                        # Prepare new name by replacing substrings
                        new_value = text.replace("40h", "5h")
                        new_value = new_values.replace("_copy", "")

                        # Clear current text (BACKSPACE), then type new value and submit
                        actions.send_keys(Keys.BACKSPACE)
                        sleep(1) # Allow UI to catch up
                        actions.send_keys(new_value)
                        actions.send_keys(Keys.ENTER)
                        actions.perform()
                        print(f"Renamed: {text} => {new_value}")
                        success = True
                        break # Exit table loop once renamed
            # Page updated causing stale element, retry after delay
            except StaleElementReferenceException as sere:
                print("Stale element encountered. Retrying...")
                send_error_message(sere)
                sleep(1)
                attempts -= 1

            # If element not found in DOM, retry as well
            except NoSuchElementException:
                print(f"Could not find element for {name}")
                attempts -= 1
            # Log if renaming failed after retries
            if not success:
                print(f"{name} not found in current table state")


# Edits all "5h / week" Arbeitszeitmodelle by setting:
# - 1 hour for each weekday (Monday to Friday)
# - 0.5 hours for "Halber Feiertag" (half public holiday)
def edit_arbeitszeitmodelle(driver):
    # List of Arbeitszeitmodelle names to rename (those with "_copy")
    arbeitszeitmodelle_names = [
        "AT 5h / week", "BG 5h / week", "HR 5h / week",
        "HU 5h / week", "PL 5h / week", "RO 5h / week",
        "UA 5h / week", "SK 5h / week", "UA 5h / week"
    ]

    wait = WebDriverWait(driver, 10)

    # Wait until the Arbeitszeitmodelle tables have loaded
    wait.until(lambda d: d.execute_script("""
        return document.querySelectorAll('#gridview-1295 table[id^="gridview-"]').length > 0;
        """))

    # Loop through each model name to edit
    for name in arbeitszeitmodelle_names:
        try:
            # Find the container with the Arbeitszeitmodelle tables
            container = wait.until(EC.presence_of_element_located(By.CSS_SELECTOR, "#gridview-1295"))
            tables = container.find_elements(By.CSS_SELECTOR, "table[id^='gridview-']")
            matched = False


            # Find and select the Arbeitszeitmodell by matching name
            for table in tables:
                try:
                    cell = table.find_element(By.CSS_SELECTOR, "tbody > tr > td")

                    if cell.text.strip() == name:
                        ActionChains(driver).click(cell).perform()
                        matched = True
                        break
                    
                # Handle dynamic DOM updates
                except StaleElementReferenceException as e:
                    print(f"Error: {e}")
                    continue
                
                except Exception as e:
                    print(f"Error: {e}")
                    continue

            if not matched:
                print(f"Arbeitszeitmodell '{name}' not found.")
                continue
            
            # Wait until the Wochenplan (week plan) tables load after selecting the model
            wait.until(lambda d: d.execute_script("""
            return document.querySelectorAll('#gridview-1318 table[id^="gridview-"]').length > 0;
            """))

            # Initialize day counter:
            # Monday = 1, ..., Halber Feiertag (Half Public Holiday) = 9
            # Skip Saturday (6), Sunday (7), Public Holiday (8)
            count = 1

            # Loop through each day in the Wochenplan:
            # Days 1 to 5 correspond to Monday through Friday,
            # Days 6 to 8 (Saturday, Sunday, Public Holiday) are skipped,
            # and day 9 corresponds to "Halber Feiertag" (Half Public Holiday),
            # which requires special handling

            while count <= 9:
                try:
                    container_wochenplan = wait.until(EC.presence_of_element_located(By.CSS_SELECTOR, "#gridview-1318"))
                    tables_wochenplan = container_wochenplan.find_element(By.CSS_SELECTOR, "table[id^='gridview-']")

                    # If the current day (`count`) exceeds the number of available Wochenplan tables,
                    # it means there aren't enough tables loaded to process all days,
                    # so we break out of the loop to prevent errors.

                    if count > len(tables_wochenplan):
                        print(f"Not enough tables to process all days for {name}")
                        break

                    # Select the table corresponding to the current day (`count`) from the list of Wochenplan tables
                    # Since list indices start at 0, we use `count - 1` to get the correct table for day `count`
                    table_wochenplan = tables_wochenplan[count - 1]

                    # Monday's row structure is different: value is in 2nd row, 4th cell
                    if count == 1:
                        cell_wochenplan = table_wochenplan.find_element(By.CSS_SELECTOR, "tbody > tr:nth-of-type(2) > td:nth-of-type(4)")

                    # Other days: value is in first row, 4th cell
                    if count > 1:
                        cell_wochenplan = table_wochenplan.find_element(By.CSS_SELECTOR, "tbody > tr > td:nth-of-type(4)")
                    
                    sleep(2)

                    # Prepare action chain to edit the hours value
                    actions = ActionChains(driver)
                    actions.double_click(cell_wochenplan)
                    sleep(2)
                    actions.send_keys(Keys.BACKSPACE) # Clear current value
                    sleep(2)

                    # Set 1 hour for Monday-Friday
                    if count >= 1 and count <= 5:
                        action.send_keys("1")
                        print("Set hour to 1")

                    # Set 0.5 hour for Halber Feiertag (half public holiday)
                    if count == 9:
                        action.send_keys("0.5")
                        print("Set hour to 0.5")

                    sleep(1)
                    # Submit the change
                    actions.send_keys(Keys.ENTER).perform()

                    count += 1
                    sleep(0.5)

                except Exception as e:
                    print(f"Failed to edit day {count} for '{name}': {e}")
                    count += 1
                    continue
            sleep(0.5)

        except Exception as e:
            print(f"Error processing '{name}': {e}")
    
    print("Edited Arbeitszeitmodelle successfully.")
