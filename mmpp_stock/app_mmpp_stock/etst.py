from selenium import webdriver
from selenium .webdriver.chrome.options import Options
import time
from selenium.webdriver.chrome.service import Service

def gsshot(url):
    service = Service(executable_path=r'C:\Users\snortiz\Documents\projects\sap_extraction\chromedriver.exe')
    c_options = Options()
    c_options.add_argument("--headless")
    c_options.add_argument("--start-maximized")

    driver = webdriver.Chrome(service=service,options=c_options)

    driver.get(url)
    time.sleep(2)

    element = driver.find_element_by_xpath("//body")
    width = 1928
    height = element.size['height'] +1000

    driver.set_window_size(width,height)
    time.sleep(2)
    driver.save_screenshot("test.png")
    driver.quit()

gsshot("https://app.powerbi.com/groups/7c1add41-a7d1-4e15-a1f5-a9309203b8ee/reports/05210caa-895c-4d54-ae1d-da9ee7f02eb1/ReportSectiona176c07c7fb31cc6bdb0?experience=power-bi")