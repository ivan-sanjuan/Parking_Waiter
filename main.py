import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import cv2
import pyautogui
from PIL import Image
from datetime import date, timedelta, datetime
from urllib.parse import urljoin
import time

driver = webdriver.Chrome()
options = Options()
driver.set_window_size(1920,1080)
url = 'https://outlook.office.com/mail/'
driver.get(url)
parking_folder = WebDriverWait(driver,250).until(EC.presence_of_element_located((By.CSS_SELECTOR,'[data-folder-name="parking"]')))
if parking_folder:
    driver.execute_script("arguments[0].click();",parking_folder)
    time.sleep(60)
    unread = WebDriverWait(driver,250).until(EC.presence_of_element_located((By.CSS_SELECTOR,'[aria-label*="Unread"]')))
    driver.execute_script("arguments[0].scrollIntoView();",unread)
    WebDriverWait(driver,250).until(EC.visibility_of_element_located((By.CSS_SELECTOR,'[aria-label*="Unread"]')))
    if unread:
        ActionChains(driver).click(unread).perform()
        time.sleep(90)
        screenshot = pyautogui.screenshot()
        screenshot.save('screen.png')
        img = cv2.imread("screen.png")
        detector = cv2.QRCodeDetector()
        data, bbox, _ = detector.detectAndDecode(img)
        if data:
            print("QR Code Data:", data)
            driver.switch_to.new_window('tab')
            driver.get(data)
            time.sleep(30)
            result = pyautogui.screenshot()
            result.save('result.png')
        else:
            print("No QR code found.")


time.sleep(20)