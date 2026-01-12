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


CONFIG = json.load(open("config.json"))
MONTHS = CONFIG['months']
USERS_INFO = CONFIG['users_info']
BUSINESS_RECIPIENTS = CONFIG['business_recipients']


FINAL_DEST = Path(os.getenv(CONFIG['local_onedrive_path'])) / CONFIG['sp_input_folder']
FINAL_DEST.mkdir(parents=True, exist_ok=True)

options = webdriver.ChromeOptions()


def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)  # When running as EXE
    else:
        return os.path.dirname(os.path.abspath(__file__))  # When running as script

def get_batch_ref(row, driver, wait): #gets the batch ref and file name before downloading the file

    downloaded_files = []

    file_name_cell = row.find_element(By.XPATH, "./td[1]")
    file_name = file_name_cell.text.strip()
    downloaded_files.append(file_name)

    batch_ref_btn = row.find_element(By.XPATH, "./td[2]//button")
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
        batch_ref_btn)

    batch_ref_cell = wait.until(
        EC.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[3]/div/div[1]")))
    batch_ref = batch_ref_cell.text.strip()
    downloaded_files.append(batch_ref[11:])


    ActionChains(driver).send_keys(Keys.ESCAPE).perform()

    return downloaded_files

def get_file_status(row): #failed: retry button, preparing: no button wait, ready: download button

    cells = row.find_element(By.XPATH, "./td[6]")
    file_status = cells.text.strip()

    return file_status

def rename_file(filename, batch_ref, download_dir): #rename per file

    if "StatementOfAccount" not in filename:

        print(f"Unexpected filename: {filename}")
        logging.warn(f"Unexpected filename: {filename}")

        return False

    new_name = filename.replace("StatementOfAccount", batch_ref)

    old_path = os.path.join(download_dir, filename)
    new_path = os.path.join(download_dir, new_name)

    if os.path.exists(new_path):

        print(f"File already exists: {new_name}")
        logging.info(f"File already exists: {new_name}")

        return True

    os.rename(old_path, new_path)
    print(f"Renamed: {filename} -> {new_name}")
    logging.info(f"Renamed: {filename} -> {new_name}")

    return True

def wait_for_single_download(before_files, download_dir, timeout=300):

    end_time = time.time() + timeout
    try:
        while time.time() < end_time:

            current_files = set(os.listdir(download_dir))
            new_files = current_files - before_files

            #ignore temp chrome files
            completed_files = [f for f in new_files if not f.endswith(".crdownload")]

            if completed_files:

                return completed_files[0]

            sleep(1)

    except (TimeoutError, TimeoutException) as e:
        print("Single file download timeout")
        logging.error(traceback.print_exc())
        logging.error(e)
        logging.error("Single file download timeout")

def check_before_req(download_dir=FINAL_DEST): #check if batch ref already installed

    downloaded_files = os.listdir(download_dir)
    refs = [f.split('_')[1] for f in downloaded_files if '_' in f]

    return refs


def month_in_range(month_text):

    row_month = datetime.strptime(month_text, "%m-%Y")
    now = datetime.now()
    start_month = now - relativedelta(months=MONTHS)

    return start_month.replace(day=1) <= row_month.replace(day=1) <= now.replace(day=1)


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

def download_center(driver, row, retries=5, delay=10):

    status_list = ["Ready to Download", "Preparing Download"]
    existing_files = set(os.listdir(DOWNLOAD_DIR))

    for attempt in range(1, retries + 1):

        try:

            file_status = get_file_status(row)

            if file_status == status_list[0]:

                print(file_status)
                download_btn = row.find_element(By.XPATH, "./td[7]//button")
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
                    download_btn)
                sleep(delay)

            elif file_status == status_list[1]: #preparing to downalod: wait

                print(file_status)
                logging.info(f"{file_status} waiting..")
                sleep(50)

                if file_status == status_list[0]:

                    sleep(delay)
                    wait_for_download_btn = row.find_element(By.XPATH, "./td[7]//button")
                    driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
                        wait_for_download_btn)
                    sleep(delay)

            elif file_status not in status_list: #Failed to download: Retry Button

                print(file_status)
                retry_btn = row.find_element(By.XPATH, "./td[7]//button")

                if retry_btn:

                    sleep(delay)
                    driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
                        retry_btn)
                    sleep(delay)

            new_files = set(os.listdir(DOWNLOAD_DIR))

            if new_files != existing_files:

                return True

        except Exception as e:

            print(f"Download center retry {attempt} failed:", e)
            logging.error(f"Download center retry {attempt} failed.")
            logging.error(traceback.print_exc())
            logging.error(e)

            sleep(delay)

    return False


def download_automation(username, password):

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

    #request files to dowanload

    driver.get("https://jisr.waseel.com/payers/102/reports/statement-of-accounts")
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    logging.info("Enter Statement of accounts.")

    if not check_req_table(wait):

        driver.quit()
        return
    logging.info("Statement of accounts Table checked")

    files_reqs = 0

    installed_files = check_before_req()

    while True:

        rows = driver.find_elements(By.XPATH, "//table/tbody/tr")

        for row in rows:

            try:

                cells = row.find_elements(By.TAG_NAME, "td")
                month_text = cells[2].text.strip()
                batch_ref = cells[1].text.strip()

                if month_in_range(month_text):  # check month & batch ref before click request

                    if batch_ref not in installed_files:

                        download_btn = row.find_element(By.XPATH, "./td[8]//button")
                        driver.execute_script(
                            "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
                            download_btn)
                        download_req = wait.until(
                            EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[3]/div[2]/button")))
                        driver.execute_script("arguments[0].click();", download_req)
                        files_reqs += 1
                        print(f"Requested download for {batch_ref} {month_text}")
                        logging.info(f"Requested download for {batch_ref} {month_text}")

                    elif batch_ref in installed_files:

                        print(f"{batch_ref} {month_text} already downloaded.")
                        logging.info(f"Already downloaded {batch_ref} {month_text}")

                        pass

            except Exception as e:

                print("Error: ", e)
                logging.error(traceback.print_exc())
                logging.error(e)

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

        except Exception as e:

            logging.exception("[Table] Unexpected error while clicking Next")
            logging.error(traceback.print_exc())
            logging.error(e)

            break

    if files_reqs == 0:

        print("No files requested. Exiting.")
        logging.info("No files requested. Exiting.")
        driver.quit()

        return

    #Download Center

    driver.get("https://jisr.waseel.com/payers/102/reports/download-center")
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    logging.info("Enter Download Center")

    if not check_req_table(wait):

        driver.quit()
        return

    logging.info("Downalod center Table checked")

    rows = driver.find_elements(By.XPATH, "//table/tbody/tr")
    latest_rows = rows[:files_reqs]

    for row in latest_rows:

        try:

            file_info = get_batch_ref(row, driver, wait)
            original_name, batch_ref = file_info

            before_files = set(os.listdir(DOWNLOAD_DIR))

            if not download_center(driver, row):

                print("Download failed after retries")
                logging.warning(f"{original_name}| {batch_ref} failed to download after retries")
                continue

            downloaded_file = wait_for_single_download(before_files, DOWNLOAD_DIR)

            if not rename_file(downloaded_file, batch_ref, DOWNLOAD_DIR):

                print("Failed to rename file")
                logging.error("Failed to rename file.")

        except Exception as e:

            print("Error processing row:", e)
            logging.error(traceback.print_exc())
            logging.error(e)

    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    driver.quit()
    logging.info(f'Done. Exiting..')


def move_files(source: Path = None, destination: Path = FINAL_DEST):

    #destination.mkdir(parents=True, exist_ok=True)

    excel_extensions = {".xls", ".xlsx", ".xlsm", ".xlsb"}

    for file in source.iterdir():

        if not file.is_file():

            continue

        if file.suffix.lower() not in excel_extensions:

            logging.warn(f"File {file} not an Excel file.")
            file.unlink()
            logging.warn(f"File {file} removed.")

            continue  #skip&remove non-Excel files

        target = destination / file.name

        if target.exists():

            print(f"{file} already exists, skipping")
            logging.info(f"{file} already exists, skipping")
            continue

        shutil.copy(file, target)
        logging.info(f"{file} copied to {target}")


def get_all_files_downloaded(download_dir):

    #Current time
    now = time.time()
    one_hour_ago = now - 3600  #3600 seconds = 1 hour

    #List all files
    files = os.listdir(download_dir)
    files = [f for f in files if os.path.isfile(os.path.join(download_dir, f))]

    #Filter files modified within the last hour
    recent_files = [f for f in files if os.path.getmtime(os.path.join(download_dir, f)) >= one_hour_ago]

    #Sort by most recently modified
    recent_files.sort(key=lambda f: os.path.getmtime(os.path.join(download_dir, f)), reverse=True)

    return recent_files

def send_report(business_recipients=BUSINESS_RECIPIENTS): #sends report

    email_sender = SendEmail(
    smtp_host=CONFIG["smtp_host"],
    smtp_port=CONFIG["smtp_port"],
    smtp_username=CONFIG["smtp_username"],
    smtp_password=os.getenv(CONFIG["smtp_password"]),
    smtp_auth=True,
    is_auth=True)

    recent_files = get_all_files_downloaded(DOWNLOAD_DIR)
    subject = "Tawuniya rejections"
    bodypreview = f"Hi,<br><br>New Tawuniya rejections are available. Following are the new files:<br><br>{recent_files}<br><br>Thanks,<br><br>Waseel RPA"

    if len(recent_files) == 0:
        logging.info(f"[Report]No new files found, Email not sent.")
    else:
        email_sender.send_email(subject=subject,
                                body=bodypreview,
                                business_recipients=business_recipients,
                                attachment_list=[])
        logging.info(f"[Report]Email sent to {business_recipients}")



def init():

    DOWNLOAD_DIR = Path(get_base_path()) / 'downloads'
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    init_logging(prefix='Log_Download')

    CLIENT_ID = CONFIG['CLIENT_ID'] # replace with config
    TENANT_ID = CONFIG['TENANT_ID']
    receiver = EmailReceiver(CLIENT_ID=CLIENT_ID, TENANT_ID=TENANT_ID, token_file_path=get_base_path())

    options.add_experimental_option("prefs", {
        "download.default_directory": str(DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })

    return receiver, DOWNLOAD_DIR



if __name__ == '__main__':

    try:

        receiver, DOWNLOAD_DIR = init()

        download_automation(username="rpa1.dkmc", password="Rpa@waseel123")
        move_files(source=DOWNLOAD_DIR)
        send_report()

    except Exception as e:

        logging.error(traceback.print_exc())
        logging.error(e)