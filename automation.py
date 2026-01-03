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
from selenium.webdriver.common.action_chains import ActionChains




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


def start_automation_process(shipment_records,dmf_records):
    URL = "https://vendorpodium.oocllogistics.com/"
    
    selectors = [
                "div.main-wrapper",    # top-level host CSS selector (change to actual host)
                "div#content",         # element inside that host's shadowRoot (if nested)
                'input[type="checkbox"]'
                ]

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
        hw = Handywrapper(driver)  # Wait for browser to be ready
        time.sleep(3)
        driver.get(URL)  # Wait for page to load
        time.sleep(2)
        def login(username, password):
            try:
                ok = hw.click_in_shadow(selectors)
                hw.wait_explicitly(By.ID, "ext_username-inputEl")
                username_field = hw.find_element(By.ID, "ext_username-inputEl")
                hw.wait_explicitly(By.ID, "ext_password-inputEl")
                password_field = hw.find_element(By.ID, "ext_password-inputEl")

                username_field.send_keys(username)
                password_field.send_keys(password)

                # hw.Click_element(By.XPATH, "//input[@type='checkbox']")
                # Wait for login to complete and click submit
                # hw.Click_element(By.ID, "accept-cookies")
                hw.Click_element(By.ID, "ext_submitLogin-btnInnerEl")
                ok = hw.click_in_shadow(selectors)
                print("Login successful.")
                time.sleep(2)  # Wait for login to process
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
                # ok = hw.click_in_shadow(selectors)
                time.sleep(2)
                hw.wait_explicitly(By.ID, "PGMID_1000023")
                hw.hover(By.ID, "PGMID_1000023")  # Hover over 'Bookings' tab
                hw.Click_element(By.ID, "PGMID_1400045")
                ok = hw.click_in_shadow(selectors)  # Click on 'Bookings' tab
                print("Navigated to Create Booking page.")
            except Exception as e:
                print(f"An error occurred while navigating to Create Booking: {e}")
                driver.quit()
                exit(1)

        login(username, password)
        start_booking()

        
        automate_process(driver, shipment_records)
             
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
            country = record.get("Country", "")
            if not country.lower() == "united states" and not country.lower() == "europe":
                continue
            
            if first_record:
                print(f"Processing record: {record}")
                time.sleep(1)
                hw.wait_explicitly(By.ID, "brCreateForm:poNumTxt")
                po_num = hw.find_element(By.ID, "brCreateForm:poNumTxt")
                po_num.send_keys("0"+str(record.get("PO#", "")).split("-")[0])
                hw.Click_element(By.ID, "brCreateForm:custComboPo")
                if record.get("Country", "") == "UNITED STATES":
                    hw.Click_element(By.XPATH, f"//select[@id='brCreateForm:custComboPo']//option[.='HADDAD APPAREL GROUP, LTD']")
                elif record.get("Customer Name", "") == "Europe":
                    hw.Click_element(By.XPATH, f"//select[@id='brCreateForm:custComboPo']//option[.='Haddad Europe']")
                else:
                    continue

                hw.Click_element(By.ID, "poItemsRadio:0")

                hw.Click_element(By.ID, "brCreateForm:next")
                time.sleep(2)  # Wait for next page to load
                hw.Click_element(By.ID, "vdrBrAmendForm:movementTypeBRSI001")
                
                plan_hod = record.get("Plan-HOD", "")
                # Convert to mm/dd/yyyy string for send_keys
                plan_hod_str = ""
                if hasattr(plan_hod, 'strftime'):
                    plan_hod_str = plan_hod.strftime("%m/%d/%Y")
                elif isinstance(plan_hod, str) and plan_hod:
                    try:
                        # Try to parse string to datetime
                        plan_hod_dt = pd.to_datetime(plan_hod, errors='coerce')
                        if pd.notnull(plan_hod_dt):
                            plan_hod_str = plan_hod_dt.strftime("%m/%d/%Y")
                        else:
                            plan_hod_str = plan_hod
                    except Exception:
                        plan_hod_str = plan_hod
                elif plan_hod is not None:
                    plan_hod_str = str(plan_hod)

                hw.wait_explicitly(By.ID, "crdBRSI008_extdate-inputEl")
                hw.find_element(By.ID, "crdBRSI008_extdate-inputEl").send_keys(plan_hod_str)
                hw.Click_element(By.XPATH, "//option[@value='CY-CY']")
                hw.Click_element(By.ID, "vdrBrAmendForm:woodPackingBRSI002:1")
                hw.Click_element(By.ID, "vdrBrAmendForm:msdsGoodsBRSI017:1")
                hw.Click_element(By.ID, "vdrBrAmendForm:hasCustomsDeclaration:1")
                first_record = False
            else:
                hw.scroll_to_element(By.ID, "vdrBrAmendForm:addItemBtn")
                hw.Click_element(By.ID, "vdrBrAmendForm:addItemBtn")
                hw.wait_explicitly(By.ID, "vdrBrAmendForm:poNumTxt")
                add_po = hw.find_element(By.ID, "vdrBrAmendForm:poNumTxt")
                add_po.clear()
                add_po.send_keys("0" + str(record.get("PO#", "")))
                hw.Click_element(By.XPATH, "//input[@id='vdrBrAmendForm:j_id1710:0']")
                hw.Click_element(By.XPATH, "//input[contains(@id,'vdrBrAmendForm:j_id') and @value='Retrieve']")
                time.sleep(3)
                no_records_alert = hw.find_element(By.XPATH, "//span[.='No Records Found']")
                if no_records_alert in ["", None]:
                    hw.Click_element(By.XPATH, "(//div[@id='retrievePoItemGridDiv']/div/div/div/div/div[contains(@id,'gridcolumn')]//span)[1]")
                    hw.Click_element(By.ID, "vdrBrAmendForm:j_id1738")
                else:
                    hw.Click_element(By.XPATH, "//input[@id='vdrBrAmendForm:j_id1719' and @value='Cancel']")

    except Exception as e:
        print(f"An error occurred during automation: {e}")


def edit_details(driver, records):
    try:
        hw = Handywrapper(driver)
        actions = ActionChains(driver)
        
        hw.wait_explicitly(By.ID, "poItemGrid-targetEl")
        added_pos = hw.find_elements_list_of_text(By.XPATH, "//div[@id='poItemGrid-targetEl']/div[1]//tr[1]/td[3]")
        # Reorder records to match the sequence in added_pos, matching PO# to po[1:]
        records = [next((record for record in records if str(record["PO#"]) == str(po)[1:]), None) for po in added_pos]
        # Remove any None values if a PO# was not found
        records = [record for record in records if record is not None]
        
        for ind in range(len(records)):
            PO_num = records[ind].get("PO#", "")
            length = int(records[ind].get("L", "")) * 2.54
            width = records[ind].get("W", "") * 2.54
            height = records[ind].get("H", "") * 2.54
            Net_weight = records[ind].get("NT-wt", "")
            Gross_weight = records[ind].get("GR-Wt", "")
            Color_code = records[ind].get("Pantone", "")

            hw.wait_explicitly(By.XPATH, f"((//div[@id='poItemGrid-targetEl']//table[@class])[2]/tbody/tr)[{ind + 2}]")
            # po_elem = hw.find_elements(By.XPATH, f"((//div[@id='poItemGrid-targetEl']//table[@class])[2]/tbody/tr)[{ind + 2}]")
            time.sleep(1)  # Wait before processing next record
            hw.scroll_to_element(By.XPATH, f"((//div[@id='poItemGrid-targetEl']//table[@class])[2]/tbody/tr)[{ind + 2}]/td[1]")
            hw.Click_element(By.XPATH, f"((//div[@id='poItemGrid-targetEl']//table[@class])[2]/tbody/tr)[{ind+ 2}]/td[1]")
            
            time.sleep(1)
            hw.wait_explicitly(By.XPATH, "(//table[contains(@id,'combobox')]//input[@value])[1]")
            package = hw.find_element(By.XPATH, "(//table[contains(@id,'combobox')]//input[@value])[1]")
            package.clear()
            package.send_keys("CARTO")
            package.send_keys("N")
            hw.Click_element(By.XPATH, f"(//li[.='CARTON'])[1]")

            hw.wait_explicitly(By.XPATH, "(//table[contains(@id,'combobox')]//input[@value])[2]")   
            unit = hw.find_element(By.XPATH, "(//table[contains(@id,'combobox')]//input[@value])[2]")
            unit.clear()
            unit.send_keys("PIEC")
            unit.send_keys("E")
            time.sleep(1)
            piece_list = hw.find_elements(By.XPATH, "(//li[.='PIECE'])")
            piece = piece_list[-1]
            hw.Click_element(By.XPATH, element=piece)

            hw.Click_element(By.XPATH, f"(//div[contains(@id,'gridview')])[3]/table")
            time.sleep(0.5)
            hw.Click_element(By.XPATH, f"(((//div[@id='poItemGrid-targetEl']//table[@class])[2]/tbody/tr)/td[2]//table[contains(@id,'ext-gen')])[{ind + 1}]")
            time.sleep(0.5)
            #vol gwt len, wid height
            hw.wait_explicitly(By.XPATH, "(//div[contains(@id,'form')]//input[contains(@id,'numberfield')])[2]")
            # hw.find_element(By.XPATH, "(//div[contains(@id,'form')]//input[contains(@id,'numberfield')])[2]")
            Gross_weight_elem = hw.find_element(By.XPATH, "(//div[contains(@id,'form')]//input[contains(@id,'numberfield')])[2]")
            Gross_weight_elem.clear()
            Gross_weight_elem.send_keys(Gross_weight)

            hw.wait_explicitly(By.XPATH, "(//div[contains(@id,'form')]//input[contains(@id,'numberfield')])[3]")
            length_elem = hw.find_element(By.XPATH, "(//div[contains(@id,'form')]//input[contains(@id,'numberfield')])[3]")
            length_elem.clear()
            length_elem.send_keys(length)
            
            hw.wait_explicitly(By.XPATH, "(//div[contains(@id,'form')]//input[contains(@id,'numberfield')])[4]")
            width_elem = hw.find_element(By.XPATH, "(//div[contains(@id,'form')]//input[contains(@id,'numberfield')])[4]")
            width_elem.clear()
            width_elem.send_keys(width)
            
            hw.wait_explicitly(By.XPATH, "(//div[contains(@id,'form')]//input[contains(@id,'numberfield')])[5]")
            height_elem = hw.find_element(By.XPATH, "(//div[contains(@id,'form')]//input[contains(@id,'numberfield')])[5]")
            height_elem.clear()
            height_elem.send_keys(height)

            hw.Click_element(By.XPATH, f"(//div[contains(@id,'gridview')])[3]/table")
            time.sleep(0.5)
            hw.scroll_to_element(By.XPATH, f"((//div[@id='poItemGrid-targetEl']//table[@class])[2]/tbody/tr)[{ind + 2}]/td[13]")
            hw.Click_element(By.XPATH, f"((//div[@id='poItemGrid-targetEl']//table[@class])[2]/tbody/tr)[{ind + 2}]/td[13]")
            
            time.sleep(0.5)
            actions.send_keys(Color_code).perform()
            # hw.wait_explicitly(By.XPATH, "//input[contains(@id,'textfield') and contains(@name,'G')]")
            # Color_code_elem = hw.find_elements(By.XPATH, "//input[contains(@id,'textfield') and contains(@name,'G')]")
            # Color_code = Color_code_elem[ind]
            # Color_code.clear()
            # time.sleep(0.5)
            # Color_code.send_keys(Color_code)
        time.sleep(1)
        hw.scroll_to_element(By.ID, "vdrBrAmendForm:submitBtn_hm_pre2")
        hw.Click_element(By.ID, "vdrBrAmendForm:submitBtn_hm_pre2")
    except Exception as e:
        print(f"An error occurred during editing details: {e}")
if __name__ == "__main__":
    shipment_records = read_excel_all_records(file_path='Shipment Plan_26Dec25 FOR HADDAD.xlsx', sheet_name='Sheet1')
    dmf_records = read_excel_all_records(file_path="DMF MASTER SHEET FOR HADDAD ORDERS_FA'24(AutoRecovered).xlsx", sheet_name='Master')
    start_automation_process()