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
from tools.email import SendEmail



CLIENT_ID = "e5727438-df7b-4945-93b2-3cfb6dace85c"
TENANT_ID = "7c01a8ce-61c6-4a98-a1c3-0062f98a7cc9"
receiver = EmailReceiver.EmailReceiver(CLIENT_ID=CLIENT_ID, TENANT_ID=TENANT_ID)

CONFIG = json.load(open("config.json"))
MONTHS = json.loads(open('config.json').read())['months']
USERS_INFO = json.loads(open('config.json').read())['users_info']
BUSINESS_RECIPIENTS = json.loads(open('config.json').read())['business_recipients']

home = Path.home()
DOWNLOAD_DIR = Path("downloads").resolve()
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

FINAL_DEST = (home / "OneDrive - Waseel ASP (1)" / "automations" / "TAWUNIYA REJECTION" / "INPUT")
FINAL_DEST.mkdir(parents=True, exist_ok=True)

options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
    "download.default_directory": str(DOWNLOAD_DIR),
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

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

def rename_file(filename, batch_ref, download_dir=DOWNLOAD_DIR): #rename per file

    if "StatementOfAccount" not in filename:

        print(f"Unexpected filename: {filename}")

        return False

    new_name = filename.replace("StatementOfAccount", batch_ref)

    old_path = os.path.join(download_dir, filename)
    new_path = os.path.join(download_dir, new_name)

    if os.path.exists(new_path):

        print(f"File already exists: {new_name}")

        return True

    os.rename(old_path, new_path)
    print(f"Renamed: {filename} -> {new_name}")

    return True

def wait_for_single_download(before_files, download_dir, timeout=300):

    end_time = time.time() + timeout

    while time.time() < end_time:

        current_files = set(os.listdir(download_dir))
        new_files = current_files - before_files

        #ignore temp chrome files
        completed_files = [f for f in new_files if not f.endswith(".crdownload")]

        if completed_files:

            return completed_files[0]

        sleep(1)

    raise TimeoutError("Single file download timeout")

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

def download_center(driver, row, wait, retries=5, delay=10):

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
            sleep(delay)

    return False


def download_automation(username, password):
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)

    driver.get("https://sso.waseel.com/") #portal

    wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username) #username
    driver.find_element(By.ID, "password").send_keys(password + Keys.RETURN) #password

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

    #request files to dowanload

    driver.get("https://jisr.waseel.com/payers/102/reports/statement-of-accounts")
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

    if not check_req_table(wait):

        driver.quit()
        return

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

                    elif batch_ref in installed_files:

                        print(f"{batch_ref} {month_text} already downloaded.")
                        pass

            except Exception as e:
                print("Error: ", e)

        try:

            next_button = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH,
                     "/html/body/div/section/section/div/div[2]/div/div[2]/div[2]/div[2]/div/div[3]/button[2]")))
            next_button.click()
            wait.until(EC.staleness_of(rows[0]))

        except TimeoutException:
            break

    if files_reqs == 0:

        print("No files requested. Exiting.")
        driver.quit()
        return

    #Download Center

    driver.get("https://jisr.waseel.com/payers/102/reports/download-center")
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

    if not check_req_table(wait):

        driver.quit()
        return

    rows = driver.find_elements(By.XPATH, "//table/tbody/tr")
    latest_rows = rows[:files_reqs]

    for row in latest_rows:

        try:

            file_info = get_batch_ref(row, driver, wait)
            original_name, batch_ref = file_info

            before_files = set(os.listdir(DOWNLOAD_DIR))

            if not download_center(driver, row, wait):

                print("Download failed after retries")
                continue

            downloaded_file = wait_for_single_download(before_files, DOWNLOAD_DIR)

            if not rename_file(downloaded_file, batch_ref):

                print("Failed to rename file")

        except Exception as e:

            print("Error processing row:", e)

    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    driver.quit()


def move_files(source: Path = DOWNLOAD_DIR, destination: Path = FINAL_DEST):

    destination.mkdir(parents=True, exist_ok=True)

    excel_extensions = {".xls", ".xlsx", ".xlsm", ".xlsb"}

    for file in source.iterdir():
        if not file.is_file():
            continue

        if file.suffix.lower() not in excel_extensions:
            file.unlink()
            continue  #skip&remove non-Excel files

        target = destination / file.name

        if target.exists():

            print(f"{file} already exists, skipping")
            continue

        shutil.copy(file, target)


def get_all_files_downloaded(download_dir=DOWNLOAD_DIR):

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
    smtp_password=CONFIG["smtp_password"],
    smtp_auth=True,
    is_auth=True)

    recent_files = get_all_files_downloaded()
    subject = "Tawuniya rejections"
    bodypreview = f"Hi,<br><br>New Tawuniya rejections are available. Following are the new files:<br><br>{recent_files}<br><br>Thanks,<br><br>Waseel RPA"
    email_sender.send_email(subject=subject,
                            body=bodypreview,
                            business_recipients=business_recipients,
                            attachment_list=[])


if __name__ == '__main__':

    for user in USERS_INFO:
        sleep(2)
        download_automation(username=user['username'],password=user['password'],)

    move_files()
    send_report()