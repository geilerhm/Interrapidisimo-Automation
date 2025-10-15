import random
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.common.exceptions import UnexpectedAlertPresentException, TimeoutException
from src.utils import create_or_load_excel, handle_alert_and_reopen
from src.config import LOGIN_URL


class AuthenticationError(Exception):
    """Custom exception for authentication errors."""
    pass

# ============================================================
# 🧱 SETUP SELENIUM DRIVER
# ============================================================
def setup_driver(show_browser=True): # Add show_browser parameter
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    if not show_browser: # Add headless argument if show_browser is False
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu") # Recommended for headless
        chrome_options.add_argument("--no-sandbox") # Recommended for headless
    driver = webdriver.Chrome(service=Service(), options=chrome_options)
    print("✅ [setup_driver] Driver started successfully.")
    return driver

# ============================================================
# 🧱 LOGIN
# ============================================================
def login(driver, username, password):
    print("🌐 [STEP 1] Logging in...")
    driver.get(LOGIN_URL)
    wait = WebDriverWait(driver, 25) # A single wait object with a 25s timeout

    user_input = wait.until(EC.presence_of_element_located((By.ID, "usernameLogin")))
    pass_input = wait.until(EC.presence_of_element_located((By.ID, "passwordLogin")))
    user_input.send_keys(username)
    pass_input.send_keys(password)

    driver.find_element(By.ID, "botonLogin").click()
    print("🔐 [login] Waiting for validation result...")

    try:
        # This inner function will be our custom wait condition.
        # It returns a truthy value ("success" or "error") when the condition is met,
        # and False otherwise, telling the wait to continue.
        def check_login_status(d):
            try:
                # Find the alert element
                alert_element = d.find_element(By.ID, "swal2-title")
                
                # If the alert is not visible, login is successful.
                if not alert_element.is_displayed():
                    return "success"
                
                # If the alert shows the error text, it's a failure.
                if "Error de autenticación" in alert_element.text:
                    return "error"
                
                # If none of the above, the "Validando..." message is still up.
                # Return False to keep waiting.
                return False

            except:
                # If the element can't be found in the DOM, it means the alert is gone.
                # This is a successful login.
                return "success"

        # We wait until our custom check_login_status function returns something other than False.
        final_status = wait.until(check_login_status)
        
        print(f"  [login] Wait finished with status: {final_status}")

        if final_status == "error":
            raise AuthenticationError("Credenciales incorrectas. Por favor, verifique su usuario y contraseña.")

    except TimeoutException:
        # This will be raised by wait.until() if check_login_status always returns False for 25 seconds.
        print("  [login] Timeout: 'Validando...' message was stuck on screen.")
        raise Exception("Login failed: validation timed out.")
    except AuthenticationError:
        raise # Re-raise for the UI to catch
    except Exception as e:
        print(f"❌ [login] An unexpected error occurred during validation: {e}")
        raise

    print("✅ [login] Logged in successfully.")
    return wait

# ============================================================
# 🧱 OPEN 'SHIPMENT EXPLORER' MODULE
# ============================================================
def open_shipment_explorer(driver, timeout=40):
    wait = WebDriverWait(driver, timeout)
    print("📦 [STEP 2] Opening 'Shipment Explorer' module...")

    try:
        initial_tabs = driver.window_handles
        card = wait.until(EC.element_to_be_clickable((By.XPATH, "//p[contains(.,'Explorador Envios')]")))
        card.click()
        print("🖱️ [STEP 2] Click performed, waiting for load...")

        start_time = time.time()
        while time.time() - start_time < timeout:
            current_tabs = driver.window_handles
            if len(current_tabs) > len(initial_tabs):
                new_tab = list(set(current_tabs) - set(initial_tabs))[0]
                driver.switch_to.window(new_tab)
                print("🆕 [STEP 2] New tab activated.")
                break
            if "ExploradorEnvios.aspx" in driver.current_url:
                print("🌐 [STEP 2] Page loading in the same tab.")
                break
            time.sleep(1)

        wait.until(EC.presence_of_element_located((By.ID, "tbxNumeroGuia")))
        print("✅ [STEP 2] Shipment Explorer loaded successfully.")
        return wait

    except Exception as e:
        print(f"❌ [STEP 2] Error opening Shipment Explorer: {e}")
        raise

# ============================================================
# 🧱 PROCESS SINGLE SHIPMENT
# ============================================================
def process_single_shipment(driver, shipment, progress_callback_for_one_item, wb, ws, file_name):
    wait = WebDriverWait(driver, 30)

    # Check for alerts and handle redirection
    if handle_alert_and_reopen(driver):
        print("⚠️ [process_single_shipment] Alert handled, redirection occurred. Explorer needs re-opening.")
        return False, True # False for success, True for needs_reopen

    print(f"\n🔎 [process_single_shipment] Processing shipment {shipment}...")
    try:
        input_field = wait.until(EC.visibility_of_element_located((By.ID, "tbxNumeroGuia")))
        input_field.clear()
        time.sleep(0.8)
        sanitized_shipment = shipment.encode('ascii', 'ignore').decode('ascii') # Sanitize input
        input_field.send_keys(sanitized_shipment)

        search_button = wait.until(EC.element_to_be_clickable((By.ID, "btnShow")))
        driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
        time.sleep(0.5)
        search_button.click()
        time.sleep(1.2)

        # Extract data
        tracking_number = driver.find_element(By.ID, "tbxNumeroGuia1").get_attribute("value").strip()
        name = driver.find_element(By.ID, "tbxNombreDes").get_attribute("value").strip()
        phone = driver.find_element(By.ID, "tbxTelefonoDes").get_attribute("value").strip()
        value = driver.find_element(By.ID, "tbxValorComercial").get_attribute("value").strip()

        if not tracking_number:
            print("⚠️ [process_single_shipment] No valid data found for this shipment.")
            return False, False # Failed for this shipment, no re-open needed

        ws.append([tracking_number, name, phone, value])
        wb.save(file_name)
        print(f"✅ [process_single_shipment] Shipment {shipment} saved ({name} | {phone} | {value})")

        time.sleep(random.uniform(1, 4))
        
        if progress_callback_for_one_item:
            progress_callback_for_one_item() # Just signal that one item is done

        return True, False # Success for this shipment, no re-open needed

    except UnexpectedAlertPresentException as e:
        print(f"❌ [process_single_shipment] Unexpected alert during processing of {shipment}: {e}")
        return False, True # False for success, True for needs_reopen

    except Exception as e:
        print(f"❌ [process_single_shipment] Error processing shipment {shipment}: {e}")
        return False, False
