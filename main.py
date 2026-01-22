import time
import random
import openpyxl
import os
import pandas as pd
import getpass
from pathlib import Path
import gspread
from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException 
from selenium.webdriver.chrome.service import Service

# ---------------- CONFIGURATION ---------------- #
WEBSITE_URL = input("Website: ")

# --- ROW LIMIT FOR TESTING ---
# Set this to the number of patients you want to process in one run.
ROW_LIMIT = 50

print("=== Login to PDSHIS ===")
USERNAME = input("Email: ")
PASSWORD = getpass.getpass("Password: ")

# ---------------- FLAGS ---------------- #
DRY_RUN_MODE = False 

SHEET_ID = "1IpDKK27QxNXbx4Y-i0BU9-yJXs_bghShjucI6bMnUBY"
PATIENT_SHEET_NAME = "Patient List"
OCRA_SHEET_NAME = "OCRA"

# --- COLUMN SETTINGS ---
PATIENT_STATUS_COL = "I"  
OCRA_STATUS_COL = "A"     

# ---------------- GOOGLE SHEETS ---------------- #

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CREDS = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
client = gspread.authorize(CREDS)

spreadsheet = client.open_by_key(SHEET_ID)
patient_ws = spreadsheet.worksheet(PATIENT_SHEET_NAME)
ocra_ws = spreadsheet.worksheet(OCRA_SHEET_NAME)

def setup_driver_persistent():
    print("--- STEP 1: Launching Browser (Persistent Profile Mode) ---")
    
    profile_path = os.path.join(os.getcwd(), "chrome_profile")
    if not os.path.exists(profile_path):
        os.makedirs(profile_path)
    
    options = webdriver.ChromeOptions()
    options.add_argument(f"user-data-dir={profile_path}")
    
    prefs = {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--start-maximized")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])

    service = Service()

    try:
        driver = webdriver.Chrome(service=service, options=options)
        print(f"   [SUCCESS] Browser Launched.")
        return driver
    except Exception as e:
        print(f"   [ERROR] Launch failed: {e}")
        raise e

def login(driver):
    print("--- STEP 2: Logging In ---")
    try:
        driver.get(WEBSITE_URL)
    except Exception as e:
        print(f"   [CRITICAL] Browser crashed loading page: {e}")
        raise e

    time.sleep(3)
    
    try:
        if driver.find_elements(By.CLASS_NAME, "dashboard-header"):
            print("   [INFO] Already logged in.")
            return
    except: pass
    
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, "email")))
        driver.find_element(By.NAME, "email").send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.NAME, "password").send_keys(Keys.RETURN)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "dashboard-header")))
        print("   [SUCCESS] Logged in.")
    except TimeoutException:
        print("   [WARNING] Login timed out. Are you already logged in?")

def mark_excel_rows(ocra_indices=None, patient_list_index=None):
    try:
        if patient_list_index is not None:
            cell = f"{PATIENT_STATUS_COL}{patient_list_index + 2}"
            print(f"   [UPDATE] Marking Patient Row {patient_list_index + 2} in Col {PATIENT_STATUS_COL}")
            patient_ws.update(range_name=cell, values=[[True]])

        if ocra_indices:
            for idx in ocra_indices:
                cell = f"{OCRA_STATUS_COL}{idx + 3}" 
                print(f"   [UPDATE] Marking OCRA Row {idx + 3} in Col {OCRA_STATUS_COL}")
                ocra_ws.update(range_name=cell, values=[[True]])
    except Exception as e:
        print(f"   [ERROR] Could not save to Google Sheets: {e}")

def create_patient_profile(driver, patient_data):
    print("   [INFO] Attempting to Create Patient Profile...")
    
    button_clicked = False
    opener_strategies = [
        (By.XPATH, "/html/body/div[2]/main/div/div[1]/div[2]/button"),
        (By.XPATH, "//button[@wire:click='createShowModal()']"),
        (By.XPATH, "//button[normalize-space()='Create Patient Profile']"),
        (By.XPATH, "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'create patient profile')]")
    ]

    for attempt in range(2): 
        if attempt == 1: 
            print("   [INFO] Button not found. Resetting to Dashboard...")
            driver.get(WEBSITE_URL)
            time.sleep(3)

        for strategy in opener_strategies:
            try:
                btn = WebDriverWait(driver, 2).until(EC.presence_of_element_located(strategy))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", btn)
                button_clicked = True
                break
            except: continue
        
        if button_clicked: break

    if not button_clicked:
        print("   [ERROR] Could not find 'Create Patient Profile' button.")
        return False

    try:
        WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "firstname")))
        driver.find_element(By.ID, "firstname").send_keys(str(patient_data.get('First Name', '')))
        driver.find_element(By.ID, "lastname").send_keys(str(patient_data.get('Last Name', '')))
        
        try:
            raw_bday = patient_data.get("Birthday (Y-M-D)")
            if raw_bday:
                year_only = pd.to_datetime(raw_bday).year
                driver.find_element(By.ID, "year_of_birth").send_keys(str(year_only))
        except: pass
        
        try: 
            sex = str(patient_data.get('Sex')).capitalize()
            driver.find_element(By.XPATH, f"//label[contains(text(), '{sex}')]").click()
        except: pass

        if DRY_RUN_MODE:
            print("   [DRY RUN] Profile filled. Skipping Save.")
            driver.refresh()
            return True

        print("   [INFO] Saving profile...")
        save_strategies = [
            (By.XPATH, "/html/body/div[2]/main/div/div[2]/div[2]/div[2]/button[2]"),
            (By.XPATH, "//button[@wire:click='create']"),
            (By.XPATH, "//button[@type='submit' and contains(text(), 'Create')]")
        ]
        
        save_clicked = False
        for strategy in save_strategies:
            try:
                save_btn = WebDriverWait(driver, 2).until(EC.element_to_be_clickable(strategy))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", save_btn)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", save_btn)
                save_clicked = True
                break
            except: continue
            
        if not save_clicked:
            print("   [ERROR] Could not click final 'Create' button inside modal.")
            return False

        print("   [INFO] Waiting for creation validation...")
        time.sleep(4)

        current_url = driver.current_url
        if "/patients/" in current_url:
            print(f"   [SUCCESS] Profile created and redirected to: {current_url}")
            return True
            
        try:
            WebDriverWait(driver, 2).until(EC.invisibility_of_element_located((By.ID, "firstname")))
            print("   [SUCCESS] Creation form closed. Assuming success.")
            return True
        except:
            print("   [ERROR] Profile creation failed. Form is still visible (check for validation errors).")
            return False

    except Exception as e:
        print(f"   [ERROR] Form filling failed: {e}")
        return False

def search_for_profile(driver, search_term, target_name_clean, patient_list_index=None, is_already_true=False):
    driver.get(WEBSITE_URL)
    time.sleep(2)
    
    try:
        search_box = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder='Search patients...']")))
        search_box.clear()
        search_box.send_keys(search_term) 
        search_box.send_keys(Keys.RETURN)
        time.sleep(3)
    except:
        driver.refresh()
        return None

    target_name_parts = target_name_clean.split()

    while True:
        try:
            results = driver.find_elements(By.XPATH, "//a[contains(@href, '/patients/')]")
            for res in results:
                res_text_lower = res.text.lower()
                match = True
                for part in target_name_parts:
                    if part not in res_text_lower:
                        match = False
                        break
                if match:
                    found_url = res.get_attribute("href")
                    print(f"   [MATCH] Found profile: {found_url}")
                                    
                    return found_url
        except: pass

        try:
            next_btn = driver.find_element(By.XPATH, "//button[@rel='next']")
            if next_btn.is_enabled():
                driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                driver.execute_script("arguments[0].click();", next_btn)
                time.sleep(3) 
                continue 
            else:
                break 
        except NoSuchElementException:
            break
            
    return None

def process_patients(driver):
    print(f"--- STEP 3: Processing Data (Limited to {ROW_LIMIT} Sets) ---")
    
    try:
        patient_data_raw = patient_ws.get_all_values()
        if not patient_data_raw:
            print("Patient sheet is empty!")
            return
        
        df_patients = pd.DataFrame(patient_data_raw[1:], columns=patient_data_raw[0])
        
        ocra_data_raw = ocra_ws.get_all_values()
        if not ocra_data_raw:
            print("OCRA sheet is empty!")
            return
        df_ocra = pd.DataFrame(ocra_data_raw[2:], columns=ocra_data_raw[1])

        df_patients.columns = df_patients.columns.str.strip()
        df_ocra.columns = df_ocra.columns.str.strip()

        df_patients['Full_Name_Clean'] = (df_patients['First Name'].astype(str).str.strip() + " " + df_patients['Last Name'].astype(str).str.strip()).str.lower().apply(lambda x: " ".join(x.split()))
        
        if 'Date' in df_ocra.columns:
            df_ocra['Date'] = pd.to_datetime(df_ocra['Date']).dt.normalize()
        df_ocra['Full_Name'] = df_ocra['First Name'].astype(str).str.strip() + " " + df_ocra['Last Name'].astype(str).str.strip()
        df_ocra['Full_Name_Clean'] = df_ocra['Full_Name'].str.lower().apply(lambda x: " ".join(x.split()))

    except Exception as e:
        print(f"CRITICAL ERROR loading Sheets: {e}")
        return

    # --- FILTER 1: OCRA Sheet Checkbox (Column A / Index 0) ---
    pending_visits = df_ocra[df_ocra.iloc[:, 0].astype(str).str.strip().str.upper() != 'TRUE'].copy()
    
    if pending_visits.empty:
        print("No pending visits found in OCRA sheet (all Column A ticked)!")
        return
        
    print(f"Found {len(pending_visits)} pending visits based on OCRA Column A.")

    grouped = pending_visits.groupby('Full_Name_Clean', sort=False)
    
    processed_count = 0 

    for patient_name_clean, group in grouped:
        if processed_count >= ROW_LIMIT:
            print(f"--- ROW LIMIT REACHED ({ROW_LIMIT}) ---")
            break

        display_name = group.iloc[0]['Full_Name']
        search_term = str(group.iloc[0]['First Name']).strip()
        
        patient_record = df_patients[df_patients['Full_Name_Clean'] == patient_name_clean]
        
        patient_list_idx = None
        is_already_true = False
        
        # --- FILTER 2: Patient List Checkbox (Column I / Index 8) ---
        if not patient_record.empty:
            patient_list_idx = patient_record.index[0]
            
            try:
                is_patient_done = str(patient_record.iloc[0, 8]).strip().upper() == 'TRUE'
                
                if is_patient_done:
                    is_already_true = True
            except Exception as e:
                print(f"   [WARNING] Could not check Patient List status column: {e}")
        else:
            print(f"   [INFO] Patient '{display_name}' not found in Patient List. Proceeding with OCRA only.")

        processed_count += 1
        print(f"--------------------------------------------------")
        print(f"Processing Set #{processed_count}: {display_name}")

        is_already_true = False 
        target_url = search_for_profile(driver, search_term, patient_name_clean, patient_list_idx, is_already_true)

        if not target_url:
            print(f"   [INFO] Profile NOT found in results. Attempting Creation...")
            
            if patient_record.empty:
                print(f"   [SKIP] Patient not in Master List (Cannot create without details).")
                continue
            
            patient_data = patient_record.iloc[0].to_dict()
            
            if create_patient_profile(driver, patient_data):
                if not DRY_RUN_MODE and patient_list_idx is not None:
                    mark_excel_rows(patient_list_index=patient_list_idx)
                    is_already_true = True
                
                print(f"   [INFO] Profile created. Searching again to find URL...")
                target_url = search_for_profile(driver, search_term, patient_name_clean, patient_list_idx, is_already_true)
            else:
                continue
        else:
            if not DRY_RUN_MODE and patient_list_idx is not None:
                print(f"   [INFO] Patient found. Marking Patient List checkbox as DONE.")
                mark_excel_rows(patient_list_index=patient_list_idx)
                is_already_true = True

        if target_url:
            driver.get(target_url)
            time.sleep(3)
        else:
            print(f"   [ERROR] Could not find profile URL even after creation. Skipping visits.")
            continue

        successful_ocra_indices = []
        visits_by_date = group.groupby('Date')

        for visit_date, visit_rows in visits_by_date:
            try:
                driver.refresh()
                time.sleep(3)

                target_date_str = visit_date.strftime("%b %d, %Y")
                
                is_duplicate = False
                
                while True:
                    try:
                        visit_container_xpath = "/html/body/div[2]/main/div/div[1]/div[1]/div[2]"
                        visit_cards = driver.find_elements(By.XPATH, f"{visit_container_xpath}/div")
                        
                        found_on_page = False
                        for i, card in enumerate(visit_cards, start=1):
                            try:
                                header_date_xpath = f"{visit_container_xpath}/div[{i}]/header/span"
                                try:
                                    card_date_text = driver.find_element(By.XPATH, header_date_xpath).text.strip()
                                except: continue
                                
                                clean_card_date = " ".join(card_date_text.split()) 
                                
                                if target_date_str.lower() in clean_card_date.lower():
                                    diag_div_xpath = f"{visit_container_xpath}/div[{i}]/div[1]"
                                    try:
                                        card_diag_text = driver.find_element(By.XPATH, diag_div_xpath).text.lower()
                                    except: card_diag_text = ""
                                    
                                    all_diags_present = True
                                    for _, r in visit_rows.iterrows():
                                        excel_diag_clean = str(r['Diagnosis']).strip().lower()
                                        diag_parts = excel_diag_clean.split()
                                        for part in diag_parts:
                                            if part not in card_diag_text:
                                                all_diags_present = False
                                                break
                                        if not all_diags_present: break
                                    
                                    if all_diags_present:
                                        is_duplicate = True
                                        found_on_page = True
                                        print(f"   [SKIP] Found duplicate: {target_date_str}")
                                        break
                            except: continue
                        
                        if found_on_page: break

                    except Exception as e: pass

                    try:
                        next_btn = driver.find_element(By.XPATH, "//button[@rel='next']")
                        if next_btn.is_enabled():
                            driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                            driver.execute_script("arguments[0].click();", next_btn)
                            time.sleep(3)
                            continue
                        else: break
                    except NoSuchElementException: break

                if is_duplicate:
                    successful_ocra_indices.extend(visit_rows.index.tolist())
                    continue

                print(f"   Adding visit: {target_date_str}")
                
                nv_clicked = False
                nv_strategies = [
                    (By.XPATH, "/html/body/div[2]/main/div/div[1]/div[2]/button"),
                    (By.XPATH, "//button[@wire:click='showCreateVisitModal()']"),
                    (By.XPATH, "//button[contains(text(), 'New Visit')]")
                ]
                
                for strat in nv_strategies:
                    try:
                        btn = WebDriverWait(driver, 3).until(EC.element_to_be_clickable(strat))
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                        time.sleep(0.5)
                        driver.execute_script("arguments[0].click();", btn)
                        nv_clicked = True
                        break
                    except: continue
                
                if not nv_clicked:
                    print("   [ERROR] Could not click 'New Visit' button.")
                    continue

                time.sleep(2)

                date_input = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "visit_date")))
                date_input.click()
                date_input.clear()
                date_input.send_keys(target_date_str)
                date_input.send_keys(Keys.TAB)

                try:
                    consultant = str(visit_rows.iloc[0].get('Consultant', '')).strip()
                    if consultant and consultant.lower() != 'nan':
                        consultant_xpath = "/html/body/div[2]/main/div/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div[5]/select"
                        select_el = driver.find_element(By.XPATH, consultant_xpath)
                        select = Select(select_el)
                        found = False
                        for opt in select.options:
                            if consultant.split()[0].lower() in opt.text.lower():
                                select.select_by_visible_text(opt.text)
                                found = True
                                break
                        if not found and len(select.options) > 1:
                            select.select_by_index(random.randint(1, len(select.options)-1))
                except: pass

                diag_count = 0
                for _, r in visit_rows.iterrows():
                    diag_text = str(r['Diagnosis']).strip()
                    
                    clicked_add = False
                    try:
                        btn = driver.find_element(By.XPATH, "//button[@wire:click.prevent='addDiagnosis']")
                        driver.execute_script("arguments[0].click();", btn)
                        clicked_add = True
                    except:
                        try:
                            driver.execute_script("""
                                var btns = document.querySelectorAll('button'); 
                                for(var i=0;i<btns.length;i++){
                                    if(btns[i].innerText.includes('Add diagnosis')) { btns[i].click(); break; }
                                }
                            """)
                            clicked_add = True
                        except: pass
                    
                    if not clicked_add:
                        print("      [WARNING] Could not click 'Add diagnosis' button.")

                    time.sleep(2)

                    try:
                        idx_base = 7 + (diag_count * 3)
                        input_xpath = f"/html/body/div[2]/main/div/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div[{idx_base}]/div[2]/input"
                        target_input = driver.find_element(By.XPATH, input_xpath)

                        if target_input:
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_input)
                            time.sleep(0.5)
                            driver.execute_script("arguments[0].focus();", target_input)
                            target_input.clear()
                            target_input.send_keys(diag_text)
                            time.sleep(3) 

                            options = target_input.find_elements(By.XPATH, "following-sibling::div//ul/li//button")
                            clicked = False
                            fallback_btn = None
                            
                            for opt in options:
                                txt = opt.text.strip()
                                if txt.lower() == diag_text.lower():
                                    opt.click()
                                    print(f"      [SUCCESS] Exact match clicked: {txt}")
                                    clicked = True
                                    break
                                if "use diagnosis anyway" in txt.lower():
                                    fallback_btn = opt
                            
                            if not clicked:
                                if fallback_btn:
                                    fallback_btn.click()
                                    print(f"      [INFO] Exact match not found. Used 'Diagnosis not found' option.")
                                else:
                                    target_input.send_keys(Keys.TAB)
                                    target_input.send_keys(Keys.ENTER)
                            
                            try:
                                modal_title = driver.find_element(By.XPATH, "//div[contains(text(), 'Visit Details Form')]")
                                modal_title.click()
                            except: pass

                    except Exception as e: 
                        print(f"      [ERROR] Diagnosis entry failed: {e}")
                    
                    try:
                        visit_type = str(r.get('Diagnosis Visit', '')).strip().lower()
                        target_radio_id = None
                        if "new" in visit_type:
                            target_radio_id = f"diagnoses.{diag_count}.diagnosis_type_new"
                        elif "follow" in visit_type or "ffup" in visit_type:
                            target_radio_id = f"diagnoses.{diag_count}.diagnosis_type_ffup"
                        
                        if target_radio_id:
                            radio_btn = driver.find_element(By.ID, target_radio_id)
                            driver.execute_script("arguments[0].scrollIntoView(true);", radio_btn)
                            radio_btn.click()
                            print(f"      Selected type: {visit_type}")
                    except Exception as e:
                        print(f"      [WARNING] Radio selection failed: {e}")

                    diag_count += 1

                if not DRY_RUN_MODE:
                    save_btn = driver.find_element(By.XPATH, "//button[contains(text(), 'Create')]")
                    driver.execute_script("arguments[0].click();", save_btn)
                    time.sleep(4)
                    print("   [SUCCESS] Visit Saved.")
                    successful_ocra_indices.extend(visit_rows.index.tolist())
                else:
                    print("   [DRY RUN] Would save visit here.")
                    driver.refresh()
                    time.sleep(2)

            except Exception as e:
                print(f"   [ERROR] Visit {visit_date} failed: {e}")
                driver.refresh()
                time.sleep(3)

        if not DRY_RUN_MODE and successful_ocra_indices:
            mark_excel_rows(ocra_indices=successful_ocra_indices)

if __name__ == "__main__":
    driver = setup_driver_persistent()
    try:
        login(driver)
        process_patients(driver)
    except Exception as e:
        print(f"CRITICAL ERROR: {e}")
    finally:
        print("Done.")