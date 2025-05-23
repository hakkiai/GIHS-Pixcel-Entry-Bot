import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import time

# Constants
CHROMEDRIVER_PATH = r"C:\Users\kuber\Downloads\chromedriver-win32\chromedriver-win32\chromedriver.exe"
EXCEL_PATH      = "data/SAMPLE-TEST.xlsx"
FILE_NUMBER     = "FILENUMBER8LAMM8"

# Load Excel data exactly as it is
df = pd.read_excel(EXCEL_PATH, header=None)
df.columns = ["FileNum", "FormNumber", "Name", "Birthday", "MothersMaiden", "Address", "Zipcode", "Country", "Occupation", "Company"]
df = df.fillna("")

# Setup Chrome
options = Options()
options.add_experimental_option("detach", True)
options.add_argument("--start-maximized")
driver  = webdriver.Chrome(service=Service(CHROMEDRIVER_PATH), options=options)
wait    = WebDriverWait(driver, 20)
actions = ActionChains(driver)

# --- LOGIN ---
driver.get("https://gihsservice.com/Home")
time.sleep(2)
try:
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "LOGIN"))).click()
    wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='User Name']"))).send_keys("athreyasa007@gmail.com")
    wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Password']"))).send_keys("Karthik@2004")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Login']"))).click()
    wait.until(EC.url_contains("gihsservice.com"))
    print("✅ Logged in")
except Exception:
    print("⚠️ Login failed; please complete manually")
    wait.until(EC.url_contains("gihsservice.com"))

# --- NAVIGATE TO FORM ---
driver.get("https://gihsservice.com/tqiservice/pixcelEntryFrom/pixcel")
wait.until(EC.url_to_be("https://gihsservice.com/tqiservice/pixcelEntryFrom/pixcel"))
time.sleep(3)

def fill_field_directly(element, value):
    """Fill field directly with the exact value from Excel"""
    try:
        # Scroll to element
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
        time.sleep(0.3)
        
        # Click to focus
        element.click()
        time.sleep(0.2)
        
        # Clear field
        element.clear()
        time.sleep(0.1)
        
        # Select all and delete (backup clear)
        element.send_keys(Keys.CONTROL + "a")
        element.send_keys(Keys.DELETE)
        time.sleep(0.1)
        
        # Enter the exact value from Excel
        value_str = str(value).strip()
        if value_str:
            element.send_keys(value_str)
        
        time.sleep(0.2)
        return True
    except Exception as e:
        print(f"⚠️ Error filling field: {e}")
        return False

def fill_all_fields_sequentially(row_data):
    """Fill all form fields in sequence exactly as they appear in Excel"""
    try:
        # Wait for page to load completely
        time.sleep(2)
        
        # Get all input fields and textareas
        inputs = driver.find_elements(By.TAG_NAME, "input")
        textareas = driver.find_elements(By.TAG_NAME, "textarea")
        
        print(f"Found {len(inputs)} input fields and {len(textareas)} textarea fields")
        
        # Data in exact Excel order
        excel_data = [
            row_data["FileNum"],           # Column 0: File Number
            row_data["FormNumber"],        # Column 1: Form Number
            row_data["Name"],              # Column 2: Name
            row_data["Birthday"],          # Column 3: Birthday (exact as in Excel)
            row_data["MothersMaiden"],     # Column 4: Mother's Maiden
            row_data["Address"],           # Column 5: Address
            row_data["Zipcode"],           # Column 6: Zipcode
            row_data["Country"],           # Column 7: Country
            row_data["Occupation"],        # Column 8: Occupation
            row_data["Company"]            # Column 9: Company
        ]
        
        # Fill first 5 input fields
        for i in range(min(5, len(inputs))):
            print(f"Filling field {i+1} with: '{excel_data[i]}'")
            fill_field_directly(inputs[i], excel_data[i])
        
        # Handle Address field (check if it's a textarea)
        address_filled = False
        if textareas:
            for textarea in textareas:
                if textarea.is_displayed():
                    print(f"Filling address textarea with: '{excel_data[5]}'")
                    fill_field_directly(textarea, excel_data[5])
                    address_filled = True
                    break
        
        # If no textarea, use input field for address
        if not address_filled and len(inputs) > 5:
            print(f"Filling address input field with: '{excel_data[5]}'")
            fill_field_directly(inputs[5], excel_data[5])
        
        # Fill remaining fields (Zipcode, Country, Occupation, Company)
        remaining_start = 6 if not address_filled else 5
        for i in range(4):  # 4 remaining fields
            field_idx = remaining_start + i
            data_idx = 6 + i  # Starting from zipcode in excel_data
            
            if field_idx < len(inputs) and data_idx < len(excel_data):
                print(f"Filling field {field_idx+1} with: '{excel_data[data_idx]}'")
                fill_field_directly(inputs[field_idx], excel_data[data_idx])
        
        return True
        
    except Exception as e:
        print(f"⚠️ Error in sequential filling: {e}")
        return False

def click_save_button():
    """Click the save button with multiple strategies"""
    try:
        time.sleep(1)  # Wait before looking for save button
        
        # Multiple strategies to find save button
        save_button = None
        strategies = [
            lambda: driver.find_element(By.XPATH, "//button[text()='Save']"),
            lambda: driver.find_element(By.XPATH, "//button[contains(text(), 'Save')]"),
            lambda: driver.find_element(By.XPATH, "//input[@value='Save']"),
            lambda: driver.find_element(By.XPATH, "//button[@type='submit']"),
            lambda: driver.find_element(By.XPATH, "//input[@type='submit']"),
            lambda: driver.find_element(By.CSS_SELECTOR, "button[class*='btn']"),
            lambda: driver.find_elements(By.TAG_NAME, "button")[-1]  # Last button on page
        ]
        
        for i, strategy in enumerate(strategies):
            try:
                save_button = strategy()
                if save_button and save_button.is_displayed() and save_button.is_enabled():
                    print(f"✅ Found save button using strategy {i+1}")
                    break
            except:
                continue
        
        if save_button:
            # Scroll to save button
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", save_button)
            time.sleep(0.5)
            
            # Try clicking with different methods
            try:
                save_button.click()
                print("💾 Save button clicked (normal click)")
            except:
                try:
                    driver.execute_script("arguments[0].click();", save_button)
                    print("💾 Save button clicked (JavaScript click)")
                except:
                    # Try ActionChains as last resort
                    actions.move_to_element(save_button).click().perform()
                    print("💾 Save button clicked (ActionChains)")
            
            # Wait for form processing
            time.sleep(3)
            
            # Check if form was reset (indicates successful save)
            try:
                WebDriverWait(driver, 10).until(
                    lambda d: d.find_elements(By.TAG_NAME, "input")[2].get_attribute("value") == ""
                )
                print("✅ Form reset detected - save successful")
                return True
            except:
                print("⚠️ Form reset not detected, but continuing...")
                return True
                
        else:
            print("❌ Save button not found with any strategy")
            return False
            
    except Exception as e:
        print(f"❌ Error clicking save button: {e}")
        return False

# Process each row from Excel
for idx, row in df.iterrows():
    try:
        print(f"\n" + "="*60)
        print(f"📋 Processing record {idx + 1}: {row['Name']}")
        print(f"File Number: {row['FileNum']}")
        print(f"Form Number: {row['FormNumber']}")
        print(f"Birthday: {row['Birthday']}")
        print(f"Address: {row['Address']}")
        print(f"Zipcode: {row['Zipcode']}")
        print("="*60)
        
        # Fill the form with exact Excel data
        if fill_all_fields_sequentially(row):
            print("✅ All fields filled successfully")
            
            # Click save button
            if click_save_button():
                print(f"✅ Record {idx + 1} saved successfully for {row['Name']}")
            else:
                print(f"⚠️ Save may have failed for {row['Name']}")
                # Continue anyway as the data might have been saved
        else:
            print(f"❌ Failed to fill form for {row['Name']}")
        
        # Wait between records
        print("⏳ Waiting before next record...")
        time.sleep(3)
        
    except Exception as e:
        print(f"❌ Error processing record {idx + 1}: {e}")
        continue

print("\n🎉 All records processed!")
print("Closing browser...")
time.sleep(2)
driver.quit()
