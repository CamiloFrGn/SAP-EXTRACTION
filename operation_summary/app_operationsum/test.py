from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
# Set the path to the web driver executable
driver_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
os.environ["PATH"] += os.pathsep + os.path.dirname(driver_path)
# Configure Chrome options
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run Chrome in headless mode (without GUI)

# Create a Chrome driver instance
driver = webdriver.Chrome( options=chrome_options)

# Load the Power BI web report URL
report_url = "https://app.powerbi.com/groups/7c1add41-a7d1-4e15-a1f5-a9309203b8ee/reports/f2ab0139-543a-458f-bd50-bedcb081b6cd/ReportSection53d3333863acde0c3e82?experience=power-bi"
driver.get(report_url)

# Perform any necessary sign-in steps programmatically
# For example, fill in the username and password fields and submit the form
# You may need to inspect the web elements and provide appropriate selectors
username = "santiagoandres.ortiz@cemex.com"
password = "Asigbi_2023*"

username_field = driver.find_element_by_id("email")
password_field = driver.find_element_by_id("password")

username_field.send_keys(username)
password_field.send_keys(password)

# Submit the login form
login_button = driver.find_element_by_id("submitBtn")
login_button.click()

# Find the refresh button and click it
refresh_button = driver.find_element_by_xpath("//button[contains(text(),'Refresh')]")
refresh_button.click()

# Wait for the refresh to finish
wait = WebDriverWait(driver, 60)  # Adjust the timeout as needed
wait.until(EC.invisibility_of_element(refresh_button))

# Quit the browser
driver.quit()
