import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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
driver  = webdriver.Chrome(service=Service(CHROMEDRIVER_PATH), options=options)
wait    = WebDriverWait(driver, 20)

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
time.sleep(1)

def fill_input(placeholder, value):
    """Scroll into view, click to focus, clear & send keys."""
    xpath_input = f"//input[@placeholder='{placeholder}']"
    xpath_textarea = f"//textarea[@placeholder='{placeholder}']"
    try:
        try:
            el = wait.until(EC.presence_of_element_located((By.XPATH, xpath_input)))
        except:
            el = wait.until(EC.presence_of_element_located((By.XPATH, xpath_textarea)))
        driver.execute_script("arguments[0].scrollIntoView(true);", el)
        el.click()
        el.clear()
        el.send_keys(str(value).strip())
    except Exception as e:
        print(f"⚠️ Couldn’t fill '{placeholder}': {e}")

for idx, row in df.iterrows():
    try:
        # format birthday to MM/DD/YYYY
        bday = datetime.strptime(str(row["Birthday"]), "%d-%m-%Y").strftime("%m/%d/%Y")

        fill_input("File Number", FILE_NUMBER)
        fill_input("Form Number", row["FormNumber"])
        fill_input("Name", row["Name"])
        fill_input("Birthday", bday)
        fill_input("Monthers Maiden", row["MothersMaiden"])
        fill_input("address", row["Address"])
        fill_input("Zipcode", row["Zipcode"])
        fill_input("Country", row["Country"])
        fill_input("Ocupation", row["Occupation"])
        fill_input("Company", row["Company"])

        # --- CLICK SAVE ---
        try:
            save_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[normalize-space(.)='Save']")))
            try:
                save_btn.click()
            except:
                # fallback to JS click
                driver.execute_script("arguments[0].click();", save_btn)
            print(f"💾 Save clicked for {row['Name']}")

            # wait until Name field is emptied (form reset)
            wait.until(lambda d: d.find_element(
                By.XPATH, "//input[@placeholder='Name']").get_attribute("value") == "")
            time.sleep(0.5)  # cushion
        except Exception as e:
            print(f"⚠️ Save failed for {row['Name']}: {e}")

        print(f"✅ Submitted {row['Name']} (Form #{row['FormNumber']})")

    except Exception as e:
        print(f"❌ Row {idx+1} error: {e}")
        continue

print("🎉 All done!")
driver.quit()
