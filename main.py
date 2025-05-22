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
EXCEL_PATH      = "data/FILENUMBER8LAMM8.xlsx"
FILE_NUMBER     = "FILENUMBER8LAMM8"

# Load Excel
df = pd.read_excel(EXCEL_PATH, header=None).iloc[:, :9]
df.columns = ["FormNumber","Name","Birthday","MothersMaiden","Address","Zipcode","Country","Occupation","Company"]
df = df.fillna("")

# Setup Chrome
options = Options()
options.add_experimental_option("detach", True)
options.add_argument("--start-maximized")  # Start with browser maximized
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
time.sleep(3)  # Increased wait time for page to load fully

def fill_field_by_label(label_text, value):
    """Fill a form field by finding its associated label and then the corresponding input element"""
    try:
        # Various strategies to find the field
        strategies = [
            # Strategy 1: Find by label text
            lambda: driver.find_element(By.XPATH, f"//label[contains(text(), '{label_text}')]/..//input | //label[contains(text(), '{label_text}')]/..//textarea"),
            
            # Strategy 2: Find by preceding text node
            lambda: driver.find_element(By.XPATH, f"//*[contains(text(), '{label_text}')]/following::input[1] | //*[contains(text(), '{label_text}')]/following::textarea[1]"),
            
            # Strategy 3: Find inputs with matching placeholder
            lambda: driver.find_element(By.XPATH, f"//input[@placeholder='{label_text}'] | //textarea[@placeholder='{label_text}']"),
            
            # Strategy 4: For File Number which might be specially handled
            lambda: driver.find_element(By.XPATH, "//input[1]") if label_text == "File Number" else None
        ]
        
        element = None
        for strategy in strategies:
            try:
                element = strategy()
                if element and element.is_displayed():
                    break
            except:
                continue
                
        if not element:
            print(f"⚠️ Could not find field for '{label_text}'")
            return False
            
        # Scroll to the element
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
        time.sleep(0.7)  # Wait longer for scroll to complete
        
        # Focus the field with a more robust approach
        actions.move_to_element(element).click().perform()
        time.sleep(0.3)
        
        # Clear with both methods to ensure it's empty
        element.clear()
        # Also try to clear with key combos in case the clear() method doesn't work
        element.send_keys(Keys.CONTROL + "a")
        element.send_keys(Keys.DELETE)
        time.sleep(0.2)
        
        # Type more naturally with small delays
        for char in str(value).strip():
            element.send_keys(char)
            time.sleep(0.03)  # Slight delay between characters
        
        # Press Tab to move to next field (helps trigger validation)
        element.send_keys(Keys.TAB)
        time.sleep(0.3)
        
        print(f"✅ Filled '{label_text}' with '{value}'")
        return True
    except Exception as e:
        print(f"⚠️ Error filling '{label_text}': {e}")
        return False

# Alternative method when labels don't work
def fill_all_fields_sequentially(row_data):
    """Fill all form fields by finding them in sequence"""
    try:
        # Format birthday if needed (assuming MM/DD/YYYY format expected on form)
        try:
            if isinstance(row_data["Birthday"], str) and "-" in row_data["Birthday"]:
                bday = datetime.strptime(row_data["Birthday"], "%d-%m-%Y").strftime("%m/%d/%Y")
            else:
                bday = str(row_data["Birthday"])
        except:
            bday = str(row_data["Birthday"])

        # Get all input and textarea elements
        inputs = driver.find_elements(By.TAG_NAME, "input")
        textareas = driver.find_elements(By.TAG_NAME, "textarea")
        
        # Make sure we have enough fields
        if len(inputs) < 9:
            print(f"⚠️ Expected at least 9 input fields, found {len(inputs)}")
            return False
            
        # Map of field positions and values with CORRECT ordering
        # This is the key change - mapping the correct Excel data to the right form fields
        field_map = [
            (inputs[0], FILE_NUMBER),                # File Number 
            (inputs[1], row_data["FormNumber"]),     # Form Number
            (inputs[2], row_data["Name"]),           # Name
            (inputs[3], bday),                       # Birthday
            (inputs[4], row_data["MothersMaiden"]),  # Mother's Maiden
        ]
        
        # Handle Address which might be a textarea
        address_element = next((t for t in textareas if t.is_displayed()), None)
        if address_element:
            field_map.append((address_element, row_data["Address"]))
            next_input_index = 5
        else:
            field_map.append((inputs[5], row_data["Address"]))
            next_input_index = 6
            
        # Continue with remaining fields
        field_map.extend([
            (inputs[next_input_index], row_data["Zipcode"]),      # Zipcode
            (inputs[next_input_index+1], row_data["Country"]),    # Country
            (inputs[next_input_index+2], row_data["Occupation"]), # Occupation
            (inputs[next_input_index+3], row_data["Company"])     # Company
        ])
        
        for element, value in field_map:
            # Scroll to the element
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
            time.sleep(0.5)
            
            # Click with JavaScript for more reliable focus
            driver.execute_script("arguments[0].click();", element)
            time.sleep(0.2)
            
            # Clear with both methods
            element.clear()
            element.send_keys(Keys.CONTROL + "a")
            element.send_keys(Keys.DELETE)
            time.sleep(0.2)
            
            # Type the value
            for char in str(value).strip():
                element.send_keys(char)
                time.sleep(0.03)
            
            # Press Tab to move to next field
            element.send_keys(Keys.TAB)
            time.sleep(0.3)
            
        return True
    except Exception as e:
        print(f"⚠️ Error in sequential field filling: {e}")
        return False

def click_save_button():
    """Click the save button using multiple strategies"""
    save_found = False
    
    # Multiple strategies to find and click the save button
    strategies = [
        # Strategy 1: By button text
        lambda: driver.find_element(By.XPATH, "//button[normalize-space(.)='Save']"),
        
        # Strategy 2: By value attribute
        lambda: driver.find_element(By.XPATH, "//button[@value='Save' or @value='save']"),
        
        # Strategy 3: By class that contains btn and text
        lambda: driver.find_element(By.XPATH, "//button[contains(@class, 'btn') and contains(text(), 'Save')]"),
        
        # Strategy 4: Just any button (last resort)
        lambda: driver.find_elements(By.TAG_NAME, "button")[-1]
    ]
    
    for strategy in strategies:
        try:
            save_btn = strategy()
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", save_btn)
            time.sleep(0.7)
            
            # Try both normal and JavaScript click
            try:
                save_btn.click()
            except:
                driver.execute_script("arguments[0].click();", save_btn)
                
            save_found = True
            print("💾 Save button clicked")
            time.sleep(1)  # Wait after clicking save
            break
        except:
            continue
    
    if not save_found:
        print("⚠️ Could not find Save button")
        return False
        
    # Wait for the form to reset
    try:
        WebDriverWait(driver, 10).until(
            lambda d: len(d.find_elements(By.TAG_NAME, "input")) > 2 and 
                      d.find_elements(By.TAG_NAME, "input")[2].get_attribute("value") == ""
        )
        time.sleep(1.5)  # Extra wait to ensure form is fully reset
        return True
    except:
        print("⚠️ Form might not have reset properly, continuing anyway")
        time.sleep(2)
        return False

for idx, row in df.iterrows():
    try:
        print(f"\n📋 Processing record for {row['Name']}")
        
        # Format birthday properly if needed
        try:
            if isinstance(row["Birthday"], str) and "-" in row["Birthday"]:
                bday = datetime.strptime(row["Birthday"], "%d-%m-%Y").strftime("%m/%d/%Y")
            else:
                bday = str(row["Birthday"])
        except:
            bday = str(row["Birthday"])
        
        # Define the correct field mapping based on the requirements
        field_data = [
            ("File Number", FILE_NUMBER),            # 1. File Number: FILE_NUMBER
            ("Form Number", row["FormNumber"]),      # 2. Form Number: FormNumber from Excel
            ("Name", row["Name"]),                   # 3. Name: Name from Excel
            ("Birthday", bday),                      # 4. Birthday: Birthday from Excel
            ("Mothers Maiden", row["MothersMaiden"]),# 5. Mother's Maiden: MothersMaiden from Excel
            ("Address", row["Address"]),             # 6. Address: Address from Excel
            ("Zipcode", row["Zipcode"]),             # 7. Zipcode: Zipcode from Excel
            ("Country", row["Country"]),             # 8. Country: Country from Excel
            ("Occupation", row["Occupation"]),       # 9. Occupation: Occupation from Excel
            ("Company", row["Company"])              # 10. Company: Company from Excel
        ]
        
        label_method_successful = True
        for label, value in field_data:
            if not fill_field_by_label(label, value):
                label_method_successful = False
                break
                
        if not label_method_successful:
            print("Falling back to sequential field filling...")
            fill_all_fields_sequentially(row)
        
        # Give time for any field validations to complete
        time.sleep(1)
        
        # Click save and wait for form reset
        click_save_button()
        
        print(f"✅ Processed entry for {row['Name']} (Form #{row['FormNumber']})")
        time.sleep(2)  # Pause between submissions

    except Exception as e:
        print(f"❌ Error processing row {idx+1}: {e}")
        continue

print("🎉 All done!")
driver.quit()
