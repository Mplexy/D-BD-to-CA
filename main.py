import pandas as pd
import time
from seleniumbase import Driver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import requests
import re
import multiprocessing
from datetime import datetime

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/91.0.4472.124 Safari/537.36",
    "Other-Header": "Value"
}
session = requests.Session()

def get_emails(email_address):
    url = f"https://www.1secmail.com/api/v1/?action=getMessages&login={email_address.split('@')[0]}&domain={email_address.split('@')[1]}"
    response = requests.get(url)
    emails = response.json()
    return emails


def fetch_email_content(email_address, email_id):
    url = f"https://www.1secmail.com/api/v1/?action=readMessage&login={email_address.split('@')[0]}&domain={email_address.split('@')[1]}&id={email_id}"
    response = requests.get(url)
    email_content = response.json()
    return email_content


def extract_otp(email_body):
    # Regular expression pattern to match OTP (6 digits)
    otp_pattern = r'(\b\d{6}\b)'

    # Search for OTP pattern in email body
    otp_match = re.search(otp_pattern, email_body)
    if otp_match:
        return otp_match.group(1)
    else:
        return None


def process_row(row):
    print("Opening driver for row:", row)
    driver = Driver(uc=True, headless=False)
    driver.maximize_window()
    url = "https://www.vfsglobal.ca/IRCC-AppointmentWave1/Account/RegisteredLogin?q=shSA0YnE4pLF9Xzwon/x/CQ1P0LBKn66dLdNUfueK+wmzUtYpC3tDHc7KgSJKZ7FG9C7PVGEduPHD32HU8RF1yxd7I8Ev5jXEiVQL8EwcvA="

    email = str(row['Account Email'])
    password = str(row['Account Password'])
    selection_centre = str(row['Selection Centre'])
    no_of_applicants = str(row['No Of Applicants'])
    appointment_category = str(row['Appointment Category'])
    irccNo = str(row['IRCC NO'])
    bdate = str(row['Birth Date'])
    fname = str(row['First Name'])
    lname = str(row['Last Name'])
    mobile_number = str(row['Mobile Number'])
    applicant_email = str(row['Applicant Email'])
    otpWaitTime = row['OTP Wait Time']
    startTime = row['After OTP Start Time']

    if len(str(bdate))>10:
        bdate = bdate[:10]
    else:
        date_object = datetime.strptime(bdate, "%m/%d/%Y")
        bdate = date_object.strftime("%Y-%m-%d")

    try:
        driver.get(url)

        try:
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="EmailId"]'))).send_keys(email)
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Password"]'))).send_keys(
                password)
        except Exception as e:
            print('Not Found')

        while True:
            try:
                captcha_input = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="CaptchaInputText"]')))
                captcha_input.send_keys()
                captcha_value = captcha_input.get_attribute("value")
                if captcha_value.strip() == "":
                    print("Captcha input is empty. Waiting...")
                    time.sleep(15)
                else:
                    print("Captcha input is not empty.")
                    break
            except Exception as e:
                print('Error occurred:')

        try:
            WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ApplicantListForm"]/div[4]/input'))).click()
        except Exception as e:
            print('Not Found')

        try:
            WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="Accordion1"]/div/div[2]/div/ul/li[1]'))).click()
        except Exception as e:
            print('Not Found')

        try:
            dropdown = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="LocationId"]')))
            select = Select(dropdown)
            select.select_by_visible_text(selection_centre)
        except Exception as e:
            print('Not Found')

        try:
            dropdown = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="NoOfApplicantId"]')))
            select = Select(dropdown)
            select.select_by_visible_text(no_of_applicants)
        except Exception as e:
            print('Not Found')

        try:
            dropdown = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="VisaCategoryId"]')))
            select = Select(dropdown)
            select.select_by_visible_text(appointment_category)
        except Exception as e:
            print('Not Found')

        try:
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="IAgree"]'))).click()
        except Exception as e:
            print('Not Found')

        try:
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnContinue"]'))).click()
        except Exception as e:
            print('Not Found')

        try:
            WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[1]/div[3]/div[2]/a'))).click()
        except Exception as e:
            print('Not Found')

        if len(str(irccNo)) > 5:
            try:
                WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="BILNumber"]'))).send_keys(irccNo)
            except Exception as e:
                print('Not Found')

        try:
            WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="DateOfBirth"]'))).send_keys(bdate)
        except Exception as e:
            print('Not Found')

        try:
            element_first_name = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="FirstName"]')))
            element_first_name.clear()
            element_first_name.send_keys(fname)
        except Exception as e:
            print('Not Found')

        try:
            element_last_name = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="LastName"]')))
            element_last_name.clear()
            element_last_name.send_keys(lname)
        except Exception as e:
            print('Not Found')

        try:
            element_mobile = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Mobile"]')))
            element_mobile.clear()
            element_mobile.send_keys(mobile_number)
        except Exception as e:
            print('Not Found')

        try:
            element_email = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="validateEmailId"]')))
            element_email.clear()
            element_email.send_keys(applicant_email)
        except Exception as e:
            print('Not Found')

        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="submitbuttonId"]'))).click()
        except Exception as e:
            print('Not Found')

        try:
            alert = WebDriverWait(driver, 30).until(EC.alert_is_present())
            alert.accept()
        except Exception as e:
            print('No alert present:')

        try:
            WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="ApplicantListForm"]/div[2]/input'))).click()
        except Exception as e:
            print('Not Found')

        email_address = email
        time.sleep(otpWaitTime)

        otp = ''

        while True:
            emails = get_emails(email_address)
            if emails:
                email_content = fetch_email_content(email_address, emails[0]['id'])
                otp = extract_otp(email_content['body'])
                if otp:
                    print(otp)
                    try:
                        WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="OTPe"]'))).send_keys(otp)
                        break
                    except Exception as e:
                        print('Error occurred:')
                        time.sleep(10)
                else:
                    print("No OTP found.")
            else:
                print("No new emails found.")

        current_date = time.strftime("%Y-%m-%d", time.localtime())

        # Concatenate current date and start time
        start_datetime_str = f"{current_date} {startTime}"

        # Convert the start datetime string to a struct_time
        start_time_struct = time.strptime(start_datetime_str, "%Y-%m-%d %H:%M:%S")

        # Convert the start time struct to seconds since epoch
        start_time_seconds = time.mktime(start_time_struct)

        # Get the current time in seconds since epoch
        current_time_seconds = time.time()

        # Calculate the difference in seconds between current time and start time
        time_diff_seconds = start_time_seconds - current_time_seconds

        # If the difference is positive, wait until it's time to start
        if time_diff_seconds > 0:
            print("Waiting for {} seconds...".format(time_diff_seconds))
            time.sleep(time_diff_seconds)

        # Start your main script logic here
        print("It's time to start the script!")

        try:
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="txtsub"]'))).click()
        except Exception as e:
            print('Not Found')

        try:
            WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="calendar"]/table/tr/td[3]/span'))).click()
        except Exception as e:
            print('Not Found')

        try:
            # Waiting for the element to be clickable
            target_td = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, 'td[style="background-color: rgb(188, 237, 145); cursor: pointer;"]'))
            )
            target_td.click()
            print("Clicked successfully!")
        except Exception as e:
            print("Error:", e)

        try:
            first_checkbox = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, 'input[type="radio"][name="selectedTimeBand"]:first-child'))
            )
            # first_checkbox = driver.find_element_by_css_selector('input[type="radio"][name="selectedTimeBand"]:first-child')
            first_checkbox.click()
            print("Clicked first checkbox successfully!")
        except Exception as e:
            print("Error:", e)

        try:
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnConfirm"]'))).click()
        except Exception as e:
            print('Not Found')

        try:
            alert = WebDriverWait(driver, 30).until(EC.alert_is_present())
            alert.accept()
        except Exception as e:
            print('No alert present:')

        try:
            WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="inline-popups"]/table/tbody/tr[2]/td[4]/span/a'))).click()
            print('File Downloaded')
            send_url = 'https://visa.mplexy.com/receive_data.php?country=CANADA&email=' +email + '&pass=' + password + '&fname='+fname+'&lname='+lname+'&applicants='+no_of_applicants+'&category='+appointment_category+'&bdate='+bdate
            response = session.get(send_url, headers=headers)
            print(response.text)
        except Exception as e:
            print('Not Found')

        print('*** END ***')
        driver.quit()

    finally:
        driver.quit()

if __name__ == '__main__':
    file_path = 'input/data.xlsx'
    xl_data = pd.read_excel(file_path)

    processes = []

    for index, row in xl_data.iterrows():
        process = multiprocessing.Process(target=process_row, args=(row,))
        processes.append(process)
        process.start()
        time.sleep(15)

    for process in processes:
        process.join()
