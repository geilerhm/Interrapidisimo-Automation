import os
import datetime
from openpyxl import Workbook, load_workbook
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC



# ============================================================
# 🧱 CREATE OR LOAD EXCEL FILE
# ============================================================
def create_or_load_excel():
    today = datetime.date.today().strftime("%Y-%m-%d")
    file_name = f"shipments_{today}.xlsx"
    if os.path.exists(file_name):
        wb = load_workbook(file_name)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["TrackingNumber", "RecipientName", "Phone", "CommercialValue"])
    return wb, ws, file_name

# ============================================================
# 🧱 HANDLE ALERTS IN BROWSER
# ============================================================
def handle_alert_and_reopen(driver):
    try:
        alert = driver.switch_to.alert
        msg = alert.text
        print(f"⚠️ [alert] Detected: {msg}")
        if "register your user" in msg.lower():
            alert.accept()
            print("🔁 Session expired, redirecting to home...")
            WebDriverWait(driver, 30).until(EC.url_contains("home/applications"))
            return True
        alert.dismiss()
    except NoAlertPresentException:
        pass
    return False



