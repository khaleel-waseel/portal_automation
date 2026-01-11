from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains


from time import sleep
import time
from datetime import datetime
from dateutil.relativedelta import relativedelta

import EmailReceiver as EmailReceiver
import json
import os
from pathlib import Path
import shutil


CLIENT_ID = "e5727438-df7b-4945-93b2-3cfb6dace85c"
TENANT_ID = "7c01a8ce-61c6-4a98-a1c3-0062f98a7cc9"
receiver = EmailReceiver.EmailReceiver(CLIENT_ID=CLIENT_ID, TENANT_ID=TENANT_ID)

MONTHS = json.loads(open('config.json').read())['months']
USERS_INFO = json.loads(open('config.json').read())['users_info']

home = Path.home()

UPLOAD_DEST = (home / "OneDrive - Waseel ASP (1)" / "automations" / "TAWUNIYA REJECTION" / "OUTPUT")
UPLOAD_DEST.mkdir(parents=True, exist_ok=True)

options = webdriver.ChromeOptions()


def get_username(output_dir=UPLOAD_DEST):
    output_files = os.listdir(output_dir)
    names = [f.split('_')[0].lower() for f in output_files]
    usernames = []
    for name in names:
        if name not in usernames:
            usernames.append(name)

    return usernames



def check_batch_ref(download_dir=UPLOAD_DEST): #gets all the batch ref from downloaded files
    downloaded_files = os.listdir(download_dir)
    refs = [f.split('_')[1] for f in downloaded_files if '_' in f]
    return refs


def get_otp_code(retries=3, delay=20): #get otp code from email with retry
    for attempt in range(1, retries + 1):
        sleep(delay)
        messages = receiver.top_messages(number_of_messages=1)

        for message in messages:
            subject = message.get("subject")
            sender = message.get("from", {}).get("emailAddress", {}).get("address")

            if sender == "otp@waseel.net" and subject == "Waseel: One Time Passcode":
                body_preview = message.get("bodyPreview", "")[90:125]
                otp_digits = "".join(filter(str.isdigit, body_preview))

                if len(otp_digits) >= 6:
                    print(f"OTP received on attempt {attempt}")
                    return otp_digits

        print(f"OTP not received yet. Retry {attempt}/{retries}")

    return None

def check_req_table(wait):
    try:
        rows = wait.until(
            EC.presence_of_all_elements_located((By.XPATH, "//table/tbody/tr")))
        if not rows:
            print("No table or data found. Exiting.")
            return False
        return True
    except TimeoutException:
        print("No table or data found. Exiting.")
        return False


def uploading_file(row, wait):
    pass


def upload_automation(username, password):
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)

    driver.get("https://sso.waseel.com/")
    wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
    driver.find_element(By.ID, "password").send_keys(password + Keys.RETURN)

    otp_input = wait.until(EC.presence_of_element_located((By.ID, "code")))

    otp_code = get_otp_code(retries=5, delay=10)
    if not otp_code:
        print("OTP not received after retries. Exiting.")
        driver.quit()
        return

    otp_input.send_keys(otp_code + Keys.RETURN)

    wait.until(
        EC.visibility_of_element_located(
            (By.XPATH, "/html/body/app-root/app-home-page/div/div[2]/div[2]/div/div/div[3]/app-registered-app-card/button")))

    driver.get("https://jisr.waseel.com/payers/102")

    try:
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "MuiBackdrop-root")))
    except TimeoutException:
        pass

    try:
        popup = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@role='dialog']")))
        lets_go_btn = popup.find_element(By.XPATH, ".//button")
        driver.execute_script("arguments[0].click();", lets_go_btn)
        wait.until(EC.invisibility_of_element_located((By.XPATH, "//div[@role='dialog']")))
    except TimeoutException:
        pass

    #Statement of Accounts page
    driver.get("https://jisr.waseel.com/payers/102/reports/statement-of-accounts")

    #Wait for table to load
    wait.until(EC.presence_of_element_located((By.XPATH, "//table/tbody/tr")))

    if not check_req_table(wait):
        driver.quit()
        return

    installed_files = check_batch_ref()
    processed_files = set()

    while True:
        #Always re-fetch rows to avoid stale elements
        rows = wait.until(
            EC.presence_of_all_elements_located((By.XPATH, "//table/tbody/tr")))

        entered_any = False

        for row in rows:
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                batch_ref = cells[1].text.strip()

                #Skip already processed batch refs
                if batch_ref in processed_files:
                    continue

                #Process only required batch refs
                if batch_ref in installed_files:
                    batch_cell = row.find_element(By.XPATH, "./td[2]")

                    driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
                        batch_cell)

                    #
                    wait.until(EC.visibility_of_element_located((By.XPATH,"/html/body/div/section/section/div/div[2]/div/div[2]/div[1]/ul/li[1]/div")))
                    print(f"Batch ref: {batch_ref}")
                    processed_files.add(batch_ref)
                    print(f"add to processed files: {batch_ref}")
                    entered_any = True
                    print("Done!")

                    #Go back and wait for table to reload
                    driver.back()
                    wait.until(EC.presence_of_element_located((By.XPATH, "//table/tbody/tr")))
                    break

            except Exception as e:
                print("Error:", e)
                raise

        #If no batch was entered on this page, go to next page
        if not entered_any:
            try:
                next_button = wait.until(EC.element_to_be_clickable((
                    By.XPATH,
                    "/html/body/div/section/section/div/div[2]/div/div[2]/div[2]/div[2]/div/div[3]/button[2]")))
                next_button.click()
                wait.until(EC.staleness_of(rows[0]))
            except TimeoutException:
                break

    driver.quit()


if __name__ == "__main__":
    file_names = get_username()

    for user in USERS_INFO:
        for file in file_names:
            if file == user['file_name']:
                sleep(2)
                upload_automation(username=user['username'],password=user['password'],)

