import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By 
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
from cred import Cred 
# NOTE: Remember to replace 'from cred import Cred' with environment variables (os.environ.get) 
# for better security if possible.

# --- Configuration ---
TARGET_URL = "https://web.spaggiari.eu/home/app/default/login.php?custcode="
WAIT_TIMEOUT = 10 
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads") 
os.makedirs(DOWNLOAD_DIR, exist_ok=True) 

# Selector for the final button inside the export configuration area
# The button is: <button type="button" class="ui-button ui-corner-all ui-widget">Conferma</button>
# Using XPath to target the button by its text 'Conferma' for robustness.
POPUP_CONFIRM_BUTTON_SELECTOR = "//button[text()='Conferma']" 

# Desired dates for the export (DD-MM-YYYY format)
EXPORT_START_DATE = "01-11-2025" 
EXPORT_END_DATE = "01-10-2026" # Set to 1/10/2026 as requested

def open_spaggiari_login_page():
    """
    Automates login, navigation, and export, treating the configuration screen 
    as a modal/static element on the main page. Sets both the start ('dal') 
    and end ('al') dates using send_keys, and clicks the final 'Conferma' button.
    """
    print("Starting Selenium web driver...")
    print(f"Downloads will be saved to: {DOWNLOAD_DIR}")

    # Set up Chrome options for robustness and custom downloads
    chrome_options = Options()
    
    chrome_options.add_argument('--ignore-ssl-errors')
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--start-maximized') 

    # Custom Download Preferences
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True 
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    driver = None
    try:
        driver = webdriver.Chrome(options=chrome_options)
        print(f"Opening URL: {TARGET_URL}")
        driver.get(TARGET_URL)
        
        wait = WebDriverWait(driver, WAIT_TIMEOUT)

        # --- STEP 1 & 2: Login ---
        print("Performing login sequence...")
        username_field = wait.until(EC.presence_of_element_located((By.ID, "login")))
        username_field.send_keys(Cred.usname)
        
        password_field = wait.until(EC.presence_of_element_located((By.ID, "password")))
        password_field.send_keys(Cred.pws)

        # --- STEP 3: Click Login Button ---
        login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.accedi")))
        login_button.click()
        print("Logged in. Waiting for dashboard.")

        # --- STEP 4: Click 'Agenda' Link ---
        agenda_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Agenda")))
        agenda_link.click()
        print("Clicked Agenda. Ready for export.")
        
        # ----------------------------------------------------------------------
        # --- STEP 5: Trigger Export & Interact with Static Configuration Area ---
        # ----------------------------------------------------------------------
        
        # Click the Excel Export Button
        excel_export_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//img[@alt='Esportazione Excel']"))
        )
        excel_export_button.click()
        print("Clicked Excel Export button. Waiting for configuration element to appear...")

        # 6.1 Set the start date ('dal') input field 
        try:
            # Wait for the input field to become visible/interactable
            date_start_field = wait.until(
                EC.visibility_of_element_located((By.ID, "dal"))
            )
            # Clear the default value before sending keys
            date_start_field.clear() 
            date_start_field.send_keys(EXPORT_START_DATE)
            print(f"Set start date ('dal') to: {EXPORT_START_DATE}")
            
        except Exception as date_error:
            print(f"Error setting start date 'dal': {type(date_error).__name__}. Skipping date input.")

        # 6.2 Set the end date ('al') input field 
        try:
            # Wait for the input field to become visible/interactable
            date_end_field = wait.until(
                EC.visibility_of_element_located((By.ID, "al"))
            )
            # Clear the default value before sending keys
            date_end_field.clear() 
            date_end_field.send_keys(EXPORT_END_DATE)
            print(f"Set end date ('al') to: {EXPORT_END_DATE}")
            
        except Exception as date_error:
            print(f"Error setting end date 'al': {type(date_error).__name__}. Skipping end date input.")

        # 6.3 Click the confirmation button
        print(f"Waiting for confirmation button using selector: '{POPUP_CONFIRM_BUTTON_SELECTOR}'")
        confirm_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, POPUP_CONFIRM_BUTTON_SELECTOR))
        )
        confirm_button.click()
        print("Clicked the final confirmation button ('Conferma').")
        
        # Final observation pause
        print("Pausing for 5 seconds to ensure the download completes on the filesystem.")
        time.sleep(5)

    except Exception as e:
        print(f"\n*** AN ERROR OCCURRED ***")
        print(f"Error Type: {type(e).__name__}")
        print(f"Error Message: {e}")
        print("\nDebugging Tip: Check the stability of the 'Conferma' button. If the text changes, the XPath will break.")
    finally:
        if driver:
            print("\nClosing the browser.")
            driver.quit()

if __name__ == "__main__":
    open_spaggiari_login_page()