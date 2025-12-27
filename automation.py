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

import pandas as pd

def read_excel_all_records(file_path, sheet_name=0):
    engines = ["xlrd", "openpyxl", None]

    def read_excel(**kwargs):
        last_error = None
        for engine in engines:
            try:
                return pd.read_excel(
                    file_path,
                    sheet_name=sheet_name,
                    engine=engine,
                    **kwargs
                )
            except Exception as e:
                last_error = e
        raise RuntimeError(f"Failed to read Excel file: {last_error}")

    # Read the sheet with first row as headers
    df = read_excel(header=0)

    # Drop fully empty rows
    df.dropna(how="all", inplace=True)

    # NaN → None (JSON-safe)
    df = df.where(pd.notna(df), None)

    # Convert to list of dictionaries with actual column names
    return df.to_dict(orient="records")


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
        time.sleep(5)  # Wait for page to load
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
                hw.Click_element(By.ID, "accept-cookies")
                hw.Click_element(By.ID, "ext_submitLogin-btnInnerEl")
                print("Login successful.")
                time.sleep(5)  # Wait for login to process
            except Exception as e:
                print(f"An error occurred during login: {e}")
                driver.quit()
                exit(1)

        def start_booking():
            try:
                hw = Handywrapper(driver)
                time.sleep(2)
                hw.wait_explicitly(By.XPATH, "//span[.='No']")
                hw.Click_element(By.XPATH, "//span[.='No']")  # Click 'No' on popup

                hw.wait_explicitly(By.ID, "PGMID_1000023")
                hw.hover(By.ID, "PGMID_1000023")  # Hover over 'Bookings' tab
                hw.Click_element(By.ID, "PGMID_1400045")  # Click on 'Bookings' tab
                print("Navigated to Create Booking page.")
            except Exception as e:
                print(f"An error occurred while navigating to Create Booking: {e}")
                driver.quit()
                exit(1)

        login(username, password)
        start_booking()

        shipment_records = read_excel_all_records(file_path='Shipment Plan_26Dec25 FOR HADDAD.xlsx', sheet_name='Sheet1')
        dmf_records = read_excel_all_records(file_path="DMF MASTER SHEET FOR HADDAD ORDERS_FA'24(AutoRecovered).xlsx", sheet_name='Master')
        automate_process(driver, shipment_records[:2])
             
        edit_details(driver, dmf_records)

    except Exception as e:
        print(f"An error occurred while setting up the WebDriver: {e}")
        driver.quit()
        exit(1)

def automate_process(driver, records):
    try:
        hw = Handywrapper(driver)
        first_record = True
        for record in records:
            if first_record:
                print(f"Processing record: {record}")
                time.sleep(2)  # Wait before processing next record
                po_num = hw.find_element(By.ID, "brCreateForm:poNumTxt")
                po_num.send_keys("0"+record.get("PO#", ""))
                hw.Click_element(By.ID, "brCreateForm:custComboPo")
                if record.get("Customer Name", "") == "USA":
                    hw.Click_element(By.XPATH, f"//select[@id='brCreateForm:custComboPo']//option[.='HADDAD APPAREL GROUP, LTD']")
                elif record.get("Customer Name", "") == "Europe":
                    hw.Click_element(By.XPATH, f"//select[@id='brCreateForm:custComboPo']//option[.='Haddad Europe']")

                hw.Click_element(By.ID, "poItemsRadio:0")

                hw.Click_element(By.ID, "brCreateForm:next")
                time.sleep(2)  # Wait for next page to load

                hw.Click_element(By.ID, "vdrBrAmendForm:movementTypeBRSI001")
                hw.find_element(By.ID, "crdBRSI008_extdate-inputEl").send_keys(record.get("Plan-HOD", ""))
                hw.Click_element(By.XPATH, "//option[@value='CY-CY']")
                hw.Click_element(By.ID, "vdrBrAmendForm:woodPackingBRSI002:1")
                hw.Click_element(By.ID, "vdrBrAmendForm:msdsGoodsBRSI017:1")
                hw.Click_element(By.ID, "vdrBrAmendForm:hasCustomsDeclaration:1")
                first_record = False
            
            time.sleep(1)
            hw.Click_element(By.ID, "vdrBrAmendForm:addItemBtn")

            hw.find_element(By.ID, "vdrBrAmendForm:poNumTxt").send_keys("0" + record.get("PO#", ""))
            time.sleep(1)  # Wait for item to be added
            hw.Click_element(By.ID, "vdrBrAmendForm:j_id1710:0")

            hw.Click_element(By.ID, "gridcolumn-1141-textEl")
            hw.Click_element(By.ID, "vdrBrAmendForm:j_id1738")


    except Exception as e:
        print(f"An error occurred during automation: {e}")
    finally:
        driver.quit()

def edit_details(driver, records):
    try:
        hw = Handywrapper(driver)
        added_pos = hw.find_elements_list_of_text(By.XPATH, "//div[@id='poItemGrid-targetEl']/div[1]//tr[1]/td[3]")
        records = ["0"+str(record.get("PO#", "")) for record in added_pos]
        
        for ind in range(len(records)):
            PO_num = records[ind].get("PO#", "")
            length = int(records[ind].get("L", "")) * 2.54
            width = records[ind].get("W", "") * 2.54
            height = records[ind].get("H", "") * 2.54
            Net_weight = records[ind].get("NT-wt", "")
            Gross_weight = records[ind].get("GR-Wt", "")
            Color_code = records[ind].get("Pantone", "")

            po_elem = hw.find_elements_list(By.XPATH, f"((//div[@id='poItemGrid-targetEl']//table[@class])[2]/tbody/tr)[{ind + 2}]")
            time.sleep(2)  # Wait before processing next record
            hw.Click_element(By.ID, f"((//div[@id='poItemGrid-targetEl']//table[@class])[2]/tbody/tr)[2]/td[1]")
            
            hw.find_element(By.XPATH, "(//table[contains(@id,'combobox')]//input[@value])[2]").clear().send_keys("CARTON")
            
            hw.find_element(By.XPATH, "(//table[contains(@id,'combobox')]//input[@value])[3]").clear().send_keys("PIECE")

            hw.Click_element(By.ID, f"((//div[@id='poItemGrid-targetEl']//table[@class])[2]/tbody/tr)[2]/td[2]")

            #vol gwt len, wid height
            hw.find_element(By.XPATH, "(//div[contains(@id,'form')]//input[contains(@id,'numberfield')])[2]").clear().send_keys(Gross_weight)
            hw.find_element(By.XPATH, "(//div[contains(@id,'form')]//input[contains(@id,'numberfield')])[3]").clear().send_keys(length)
            hw.find_element(By.XPATH, "(//div[contains(@id,'form')]//input[contains(@id,'numberfield')])[4]").clear().send_keys(width)
            hw.find_element(By.XPATH, "(//div[contains(@id,'form')]//input[contains(@id,'numberfield')])[5]").clear().send_keys(height)

            hw.Click_element(By.ID, f"((//div[@id='poItemGrid-targetEl']//table[@class])[2]/tbody/tr)[2]/td[13]")
            hw.find_element(By.XPATH, "//input[contains(@id,'textfield') and contains(@id,'input') and @name='G06']").clear().send_keys(Color_code)
        
        hw.Click_element(By.ID, "vdrBrAmendForm:submitBtn_hm_pre2")
    except Exception as e:
        print(f"An error occurred during editing details: {e}")

start_automation_process()