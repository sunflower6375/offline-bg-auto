import gspread
import gspread.utils
from selenium import webdriver 
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
import time

# source google sheet
url = "https://docs.google.com/spreadsheets/d/1rMfCfMiKN3WtYXlTFaNZ8ji-exFpVzgz-O90ReDsedI/edit#gid=761486896"

# access the gsheet
gc = gspread.service_account(filename="credentials.json")
workbook = gc.open_by_url(url)
sheet = workbook.worksheet("0. Detail Overall Dashboard")

# define range to screenshot
start_cell = "B18"
end_cell = "AC121"
start_row, start_col = gspread.utils.a1_to_rowcol(start_cell)
end_row, end_col = gspread.utils.a1_to_rowcol(end_cell)

# Setup headless Chrome
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1920x1080")
driver = webdriver.Chrome(options=chrome_options, executable_path="C:\Users\linh.mynguyen\chromedriver_win32\chromedriver.exe")

# navigate to sheet
driver.get(url)

# wait for sheet to load
time.sleep(10)

# scroll to the start cell
start_cell_element = driver.find_element(By.XPATH, f'//div[@data-row="{start_row-1}"][@data-col="{start_col-1}"]')
actions = ActionChains(driver)
actions.move_to_element(start_cell_element).perform()

# wait for scroll to complete
time.sleep(6)

# take screenshot of defined range
screenshot_path = "screenshot.png"
driver.save_screenshot(screenshot_path)

driver.quit()

print(f"Screenshot saved to {screenshot_path}")

