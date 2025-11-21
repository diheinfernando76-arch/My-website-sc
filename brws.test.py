import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By 
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
from cred import Cred 

# --- Configuration ---
TARGET_URL = "https://web.spaggiari.eu/home/app/default/login.php?custcode="
WAIT_TIMEOUT = 10 # Maximum time to wait for an element in seconds

def open_spaggiari_login_page():
    """
    Initializes a Chrome browser, navigates to the target URL,
    fills in the login and password fields, clicks the login button,
    clicks the 'Agenda' link, and finally clicks the 'Excel Export' button, 
    all using reliable explicit waits.
    """
    print("Starting Selenium web driver...")

    # Set up Chrome options for robustness
    chrome_options = Options()
    
    # 1. Ignore SSL certificate errors (good for troubleshooting CDN/security issues)
    chrome_options.add_argument('--ignore-ssl-errors')
    chrome_options.add_argument('--ignore-certificate-errors')
    
    # 2. Add an argument to maximize the window right away
    chrome_options.add_argument('--start-maximized') 
    
    # Initialize the Chrome driver.
    try:
        driver = webdriver.Chrome(options=chrome_options)
        print(f"Opening URL: {TARGET_URL}")
        driver.get(TARGET_URL)
        
        # Initialize WebDriverWait to handle dynamic loading
        wait = WebDriverWait(driver, WAIT_TIMEOUT)

        print("Page opened. Waiting for login field to be present...")

        # --- STEP 1 & 2: Fill Credentials ---
        
        # 1. Locate the Username/Email field using its ID "login"
        username_field = wait.until(
            EC.presence_of_element_located((By.ID, "login"))
        )
        username_field.send_keys(Cred.usname)
        print(f"Typed '{Cred.usname}' into the Codice Personale / Email field.")
        
        # 2. Locate the Password field using its ID "password"
        password_field = wait.until(
            EC.presence_of_element_located((By.ID, "password"))
        )
        password_field.send_keys(Cred.pws)
        print(f"Typed '{Cred.pws}' into the Password field.")

        # --- STEP 3: Click Login Button ---

        # 3. Locate the Login Button and Click it
        login_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.accedi"))
        )
        login_button.click()
        print("Clicked the 'Entra con le credenziali' button. Waiting for next page to load...")

        # --- STEP 4: Click 'Agenda' Link ---
        
        # Wait for the "Agenda" link to appear on the new dashboard page.
        print("Waiting for 'Agenda' link to appear on the dashboard...")
        agenda_link = wait.until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Agenda"))
        )
        agenda_link.click()
        print("Clicked the 'Agenda' link. Now on the Agenda page.")
        
        # --- STEP 5: Click Excel Export Button ---
        
        # Wait for the Excel image button using its alt text.
        print("Waiting for the 'Esportazione Excel' button...")
        excel_export_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//img[@alt='Esportazione Excel']"))
        )
        excel_export_button.click()
        print("Clicked the 'Esportazione Excel' button to start download.")

        # Final observation pause
        print("Pausing for 5 seconds to confirm the download action.")
        time.sleep(5)

    except Exception as e:
        print(f"An error occurred: {e}")
        print("Debugging notes: Ensure all element selectors are correct and authentication was successful.")
    finally:
        if 'driver' in locals() and driver:
            print("Closing the browser.")
            driver.quit()

if __name__ == "__main__":
    open_spaggiari_login_page()