import sys
import traceback

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException,NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains


from time import sleep
import time
from datetime import datetime
from dateutil.relativedelta import relativedelta

from tools.EmailReceiver import EmailReceiver
import json
import os
from pathlib import Path
import shutil
import logging
from tools.email import SendEmail
from tools.logging_setup import init as init_logging



def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)  # When running as EXE
    else:
        return os.path.dirname(os.path.abspath(__file__))  # When running as script

CONFIG_PATH = Path(get_base_path()) / 'config.json'
CONFIG = json.load(open(CONFIG_PATH))
MONTHS = CONFIG['months']
USERS_INFO = CONFIG['users_info']
BUSINESS_RECIPIENTS = CONFIG['business_recipients']


options = webdriver.ChromeOptions()

def get_username(output_dir):

    output_files = os.listdir(output_dir)
    names = [f.split('_')[0].lower() for f in output_files]
    usernames = []
    for name in names:
        if name not in usernames:
            usernames.append(name)

    return usernames

def check_batch_ref(download_dir): #gets all the batch ref from downloaded files
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
                    logging.info(f"OTP received on attempt {attempt}")
                    return otp_digits

        print(f"OTP not received yet. Retry {attempt}/{retries}")
        logging.warn(f"OTP not received yet. Retry {attempt}/{retries}")

    return None

def check_req_table(wait):

    try:

        rows = wait.until(
            EC.presence_of_all_elements_located((By.XPATH, "//table/tbody/tr")))

        if not rows:

            print("No table or data found. Exiting.")
            logging.info("No table or data found.")

            return False

        return True

    except TimeoutException as e:

        print("No table or data found. Exiting.")
        logging.error(traceback.print_exc())
        logging.error(e)

        return False



def uploading_file():
    pass


def upload_automation(username, password):

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)

    driver.get("https://sso.waseel.com/") #portal
    logging.info("Enter portal.")

    wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username) #username
    logging.info("Send username.")

    driver.find_element(By.ID, "password").send_keys(password + Keys.RETURN) #password
    logging.info("Send password.")

    MAX_OTP_SUBMIT_ATTEMPTS = 3

    otp_verified = False

    for attempt in range(1, MAX_OTP_SUBMIT_ATTEMPTS + 1):

        logging.info(f"OTP submit attempt {attempt}/{MAX_OTP_SUBMIT_ATTEMPTS}")

        otp_input = wait.until(EC.presence_of_element_located((By.ID, "code")))
        otp_input.clear()

        #call get_otp_code
        otp_code = get_otp_code(retries=3, delay=10)

        if not otp_code:

            logging.warning("OTP not received after retries.")

            continue

        otp_input.send_keys(otp_code + Keys.RETURN)
        logging.info("OTP submitted.")

        try:
            #success condition
            wait.until(
                EC.visibility_of_element_located(
                    (By.XPATH, "/html/body/app-root/app-home-page")))

            otp_verified = True
            logging.info("OTP verified successfully.")

            break

        except TimeoutException:

            logging.warning("OTP invalid or expired, retrying...")

        #resend button
        try:

            resend_btn = driver.find_element(
                By.XPATH, "/html/body/div/div[2]/div[3]/div/div[2]/div/form/div[2]/div[2]/a")
            resend_btn.click()
            logging.info("Resend OTP clicked.")

        except NoSuchElementException as e:

            logging.error(traceback.print_exc())
            logging.error(e)

            pass


    if not otp_verified:

        logging.error("OTP verification failed after all attempts.")
        driver.quit()

        return

    driver.get("https://jisr.waseel.com/payers/102")


    driver.get("https://jisr.waseel.com/payers/102")

    try:

        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "MuiBackdrop-root")))

    except TimeoutException:

        logging.error("Timed out while waiting for page to load.")

        pass

    try:

        popup = wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@role='dialog']")))
        lets_go_btn = popup.find_element(By.XPATH, ".//button")
        driver.execute_script("arguments[0].click();", lets_go_btn)
        wait.until(EC.invisibility_of_element_located((By.XPATH, "//div[@role='dialog']")))

    except TimeoutException:

        logging.error("Timed out while waiting for page to load.")

        pass

    #Statement of Accounts page
    driver.get("https://jisr.waseel.com/payers/102/reports/statement-of-accounts")
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    logging.info("Enter Statement of accounts.")

    if not check_req_table(wait):

        driver.quit()
        return
    logging.info("Statement of accounts Table checked")

    installed_files = check_batch_ref(UPLOAD_DEST)
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

                    logging.info(f"Batch ref {batch_ref} already processed.")
                    continue

                #Process only required batch refs
                if batch_ref in installed_files:

                    batch_cell = row.find_element(By.XPATH, "./td[2]")

                    driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
                        batch_cell)

                    wait.until(EC.visibility_of_element_located((By.XPATH,"/html/body/div/section/section/div/div[2]/div/div[2]/div[1]/ul/li[1]/div")))
                    logging.info(f"Batch ref: {batch_ref}")

                    processed_files.add(batch_ref)
                    logging.info(f"Add to processed files: {batch_ref}")

                    entered_any = True
                    logging.info("Done")

                    #Go back and wait for table to reload
                    driver.back()
                    logging.info("Back to main page.")
                    wait.until(EC.presence_of_element_located((By.XPATH, "//table/tbody/tr")))
                    break

            except Exception as e:

                print("Error:", e)
                logging.error(traceback.print_exc())
                logging.error(e)

                raise

        #If no batch was entered on this page, go to next page
        if not entered_any:

            try:

                next_button = wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH,
                         "/html/body/div/section/section/div/div[2]/div/div[2]/div[2]/div[2]/div/div[3]/button[2]")))
                next_button.click()
                wait.until(EC.staleness_of(rows[0]))
                logging.info("[Table] Next page")

            except TimeoutException:

                logging.info("[Table] No next page found.")

                break

            except StaleElementReferenceException:

                logging.warning("[Table] Page refreshed while paginating. Stopping.")

                break

    driver.quit()



def init():

    UPLOAD_DEST = Path(os.getenv(CONFIG['local_onedrive_path'])) / CONFIG['sp_input_folder']
    UPLOAD_DEST.mkdir(parents=True, exist_ok=True)
    init_logging(prefix='Log_Upload')

    CLIENT_ID = CONFIG['CLIENT_ID']
    TENANT_ID = CONFIG['TENANT_ID']
    receiver = EmailReceiver(CLIENT_ID=CLIENT_ID, TENANT_ID=TENANT_ID, token_file_path=get_base_path())

    return UPLOAD_DEST, receiver



if __name__ == "__main__":

    try:

        UPLOAD_DEST, receiver = init()

        file_names = get_username(UPLOAD_DEST)

        for user in USERS_INFO:
            for file in file_names:
                if file == user['file_name']:
                    logging.info(f'Processing user : {user['username']}')
                    sleep(2)
                    upload_automation(username=user['username'],password=os.getenv(user['password']),)

    except Exception as e:

        logging.error(traceback.print_exc())
        logging.error(e)