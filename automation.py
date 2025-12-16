from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
from utils import Handywrapper
import os
import shutil


username = "OLLATISU001"
password = "Atil@987654321"

def read_excel_to_dicts(file_path, sheet_name=0, header_search_rows=10):
    # Try engines that commonly work with XLS/XLSX
    engines = ["xlrd", "openpyxl", None]

    # Step 1: load a small preview without header to detect header row
    preview = None
    last_err = None
    for eng in engines:
        try:
            preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=header_search_rows, engine=eng)
            break
        except Exception as e:
            last_err = e

    if preview is None:
        raise RuntimeError(f"Unable to read Excel preview: {last_err}")

    header_row = None
    for idx, row in preview.iterrows():
        # Count non-null, non-empty values
        non_null_count = row.count()
        if non_null_count >= 2:
            header_row = idx
            break

    if header_row is None:
        header_row = 0

    # Step 2: read full sheet with detected header
    df = None
    last_err = None
    for eng in engines:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, engine=eng)
            break
        except Exception as e:
            last_err = e

    if df is None:
        raise RuntimeError(f"Unable to read Excel with header row {header_row}: {last_err}")

    # Normalize column names
    df.columns = [str(c).strip() if not pd.isna(c) else "" for c in df.columns]

    # Drop rows that are entirely empty
    df = df.dropna(how="all")

    # Replace NaN with None for JSON-compatibility and convert to list of dicts
    df = df.where(pd.notnull(df), None)
    records = df.to_dict(orient="records")

    return records


def start_automation_process():
    URL = "https://vendorpodium.oocllogistics.com/"

    # If preferred, manually start Chrome, solve Cloudflare, then attach Selenium to it.
    # Set ATTACH_CHROME=1 in the environment before running this script to enable.
    attach = os.getenv("ATTACH_CHROME", "0") == "1"

    service = Service()
    if attach:
        options = Options()
        # connect to a running Chrome that was started with --remote-debugging-port=9222
        options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        driver = webdriver.Chrome(service=service, options=options)
    else:
        options = Options()
        # Copy your existing 'Profile 1' into a non-default user-data-dir so
        # remote debugging can be used while preserving your token stored in that profile.
        orig_user_data_dir = r"C:\\Users\\admin\AppData\\Local\\Google\\Chrome\\User Data"
        profile_name = "Profile 1"
        src_profile = os.path.join(orig_user_data_dir, profile_name)

        temp_parent = r"C:\\temp\\selenium-chrome-profile"
        dst_profile = os.path.join(temp_parent, profile_name)

        os.makedirs(temp_parent, exist_ok=True)
        try:
            if os.path.exists(src_profile) and not os.path.exists(dst_profile):
                shutil.copytree(src_profile, dst_profile)
        except Exception as e:
            print(f"Warning: failed to copy profile (will attempt to continue): {e}")

        # Copy the 'Local State' file (contains encrypted keys) so DPAPI decryption will work
        src_local_state = os.path.join(orig_user_data_dir, "Local State")
        dst_local_state = os.path.join(temp_parent, "Local State")
        try:
            if os.path.exists(src_local_state) and not os.path.exists(dst_local_state):
                shutil.copy2(src_local_state, dst_local_state)
        except Exception as e:
            print(f"Warning: failed to copy Local State (decryption may fail): {e}")

        # Use the copied profile under the non-default user-data-dir
        options.add_argument(f"--user-data-dir={temp_parent}")
        options.add_argument(f"--profile-directory={profile_name}")

        # REQUIRED to avoid crash
        options.add_argument("--remote-allow-origins=*")
        # Enable remote debugging on a fixed port so Chrome can start cleanly for automation
        options.add_argument("--remote-debugging-port=9222")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-background-networking")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option('useAutomationExtension', False)
        options.add_experimental_option('excludeSwitches', ['enable-logging'])


        driver = webdriver.Chrome(service=service, options=options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": """
                        Object.defineProperty(navigator, 'webdriver', {
                            get: () => undefined
                        });
                    """
        })
    try:
        time.sleep(2)  # Wait for browser to be ready
        driver.get(URL)

        def login(username, password):
            try:
                hw = Handywrapper(driver)
                time.sleep(3)  # Wait for page to load
                # Example usage:
                selectors = [
                "div.main-wrapper",    # top-level host CSS selector (change to actual host)
                "div#content",         # element inside that host's shadowRoot (if nested)
                'input[type="checkbox"]'
                ]
                ok = hw.click_in_shadow(selectors)

                hw.wait_explicitly(By.ID, "ext_username-inputEl")
                username_field = hw.find_element(By.ID, "ext_username-inputEl")
                hw.wait_explicitly(By.ID, "ext_password-inputEl")
                password_field = hw.find_element(By.ID, "ext_password-inputEl")

                username_field.send_keys(username)
                password_field.send_keys(password)

                # hw.Click_element(By.XPATH, "//input[@type='checkbox']")
                # Wait for login to complete and click submit
                hw.Click_element(By.ID, "ext_submitLogin-btnInnerEl")
                print("Login successful.")
                time.sleep(10)  # Wait for login to process
            except Exception as e:
                print(f"An error occurred during login: {e}")
                driver.quit()
                exit(1)

        def start_booking():
            try:
                hw = Handywrapper(driver)
                hw.hover(By.ID, "PGMID_1000023")  # Hover over 'Bookings' tab
                hw.Click_element(By.ID, "PGMID_1400045")  # Click on 'Bookings' tab
                print("Navigated to Create Booking page.")
            except Exception as e:
                print(f"An error occurred while navigating to Create Booking: {e}")
                driver.quit()
                exit(1)

        login(username, password)
        start_booking()
        
        records = read_excel_to_dicts("PAK LPO 10-1.xls", sheet_name="Extended")
        automate_process(driver, records)


    except Exception as e:
        print(f"An error occurred while setting up the WebDriver: {e}")
        driver.quit()
        exit(1)

def automate_process(driver, records):
    try:
        hw = Handywrapper(driver)
        for record in records:
            print(f"Processing record: {record}")
            hw.find_element(By.ID, "brCreateForm:poNumTxt").send_keys(record.get("PO Number", ""))

            hw.Click_element(By.ID, "brCreateForm:custComboPo")
            if record.get("Customer Name", "") == "USA":
                hw.Click_element(By.XPATH, f"//select[@id='brCreateForm:custComboPo']//option[.='HADDAD APPAREL GROUP, LTD']")
            elif record.get("Customer Name", "") == "Europe":
                hw.Click_element(By.XPATH, f"//select[@id='brCreateForm:custComboPo']//option[.='Haddad Europe']")

            hw.Click_element(By.ID, "brCreateForm:next")
            time.sleep(2)  # Wait for next page to load

    except Exception as e:
        print(f"An error occurred during automation: {e}")
    finally:
        driver.quit()


start_automation_process()