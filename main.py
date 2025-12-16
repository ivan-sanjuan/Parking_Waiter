import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from keys import key_details
import cv2
import pyautogui
from PIL import Image
from datetime import date, timedelta, datetime
import time

# REMINDERS:
# 1. Check that email grouping is disabled.
# 2. Check if Reading pane is set to Fill screen.

def answer_form():
    keys = key_details()
    start = WebDriverWait(driver,100).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="form-main-content1"]/div/button')))
    driver.execute_script("arguments[0].click();",start)
    test = WebDriverWait(driver,100).until(EC.visibility_of_element_located((By.CSS_SELECTOR,'[data-automation-id="questionItem"]')))
    if test.is_displayed():
        questions = driver.find_elements(By.CSS_SELECTOR,'[data-automation-id="questionItem"]')
        for q in questions:
            question = q.find_element(By.CLASS_NAME,'text-format-content').text.strip()
            if question == 'Employee Number':
                driver.execute_script("arguments[0].scrollIntoView();",q)
                time.sleep(1)
                emp_num = q.find_element(By.CSS_SELECTOR,'[data-automation-id="textInput"]')
                emp_num.send_keys(keys[0]['employee_num'])
            elif question == 'Last Name':
                driver.execute_script("arguments[0].scrollIntoView();",q)
                time.sleep(1)
                emp_num = q.find_element(By.CSS_SELECTOR,'[data-automation-id="textInput"]')
                emp_num.send_keys(keys[0]['last_name'])
            elif question == 'First Name':
                driver.execute_script("arguments[0].scrollIntoView();",q)
                time.sleep(1)
                emp_num = q.find_element(By.CSS_SELECTOR,'[data-automation-id="textInput"]')
                emp_num.send_keys(keys[0]['first_name'])
            elif question == 'Vehicle Make and Model':
                driver.execute_script("arguments[0].scrollIntoView();",q)
                time.sleep(1)
                emp_num = q.find_element(By.CSS_SELECTOR,'[data-automation-id="textInput"]')
                emp_num.send_keys(keys[0]['vehicle'])
            elif question == 'Plate Number':
                driver.execute_script("arguments[0].scrollIntoView();",q)
                time.sleep(1)
                emp_num = q.find_element(By.CSS_SELECTOR,'[data-automation-id="textInput"]')
                emp_num.send_keys(keys[0]['plate_num'])
            elif question.startswith('Shift'):
                driver.execute_script("arguments[0].scrollIntoView();",q)
                time.sleep(1)
                emp_num = q.find_element(By.CSS_SELECTOR,'[data-automation-id="textInput"]')
                emp_num.send_keys(keys[0]['shift'])
            elif question.startswith('Workdays'):
                driver.execute_script("arguments[0].scrollIntoView();",q)
                time.sleep(1)
                q.find_element(By.CSS_SELECTOR,'input[value="Monday"]').click()
                q.find_element(By.CSS_SELECTOR,'input[value="Tuesday"]').click()
                q.find_element(By.CSS_SELECTOR,'input[value="Wednesday"]').click()
                q.find_element(By.CSS_SELECTOR,'input[value="Thursday"]').click()
                q.find_element(By.CSS_SELECTOR,'input[value="Friday"]').click()
        time.sleep(10)
        submit = driver.find_element(By.CSS_SELECTOR,'[data-automation-id="submitButton"]')
        driver.execute_script("arguments[0].scrollIntoView();",submit)
        time.sleep(5)
        driver.execute_script("arguments[0].click();",submit)

driver = webdriver.Chrome()
options = Options()
driver.set_window_size(1920,1080)
url = 'https://outlook.office.com/mail/'
driver.get(url)
parking_folder = WebDriverWait(driver,200).until(EC.presence_of_element_located((By.CSS_SELECTOR,'[data-folder-name="parking"]')))
if parking_folder:
    driver.execute_script("arguments[0].click();",parking_folder)
    time.sleep(10) #return to 60
    unread = WebDriverWait(driver,39600).until(EC.presence_of_element_located((By.CSS_SELECTOR,'[aria-label*="Unread"]')))
    driver.execute_script("arguments[0].scrollIntoView();",unread)
    # WebDriverWait(driver,250).until(EC.visibility_of_element_located((By.CSS_SELECTOR,'[aria-label*="Unread"]')))
    if unread:
        try:
            ActionChains(driver).click(unread).perform()
            time.sleep(30) #return to 90
            message_body = driver.find_element(By.CSS_SELECTOR,'[role="document"]')
            link = message_body.find_element(By.XPATH, "//*[starts-with(@href, 'https://forms.office.com/')]").get_attribute('href')
            driver.switch_to.new_window('tab')
            driver.get(link)
            answer_form()
            print('RESERVATION via LINK, SUCCESSFUL.')
        except:
            ActionChains(driver).click(unread).perform()
            time.sleep(30)
            screenshot = pyautogui.screenshot()
            screenshot.save('screen.png')
            img = cv2.imread("screen.png")
            detector = cv2.QRCodeDetector()
            data, bbox, _ = detector.detectAndDecode(img)
            if data:
                print("QR Code Data:", data)
                driver.switch_to.new_window('tab')
                driver.get(data)
                answer_form()
                # time.sleep(30) #return to 30 // can be removed
                result = pyautogui.screenshot()
                result.save('result.png')
                print('RESERVATION via QR, SUCCESSFUL.')
    else:
        print("No Link or QR code found.")

driver.quit()
time.sleep(20)