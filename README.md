from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
import pyautogui
import time

options = Options()
# DO NOT run headless
options.add_argument("--start-maximized")

service = Service("C:\\Users\\gnandgow@12725\\bin\\edgedriver-136.0.3240.50\\msedgedriver.exe")
driver = webdriver.Edge(service=service, options=options)

driver.get("https://your-angular-site.com")
time.sleep(10)  # Wait for full render

# Simulate Ctrl+S
pyautogui.hotkey('ctrl', 's')
time.sleep(2)

# Enter file name and hit Enter
pyautogui.typewrite('saved_page')
pyautogui.press('enter')

time.sleep(3)
driver.quit()