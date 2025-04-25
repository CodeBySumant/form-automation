from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import random
from selenium.webdriver.common.action_chains import ActionChains

# Function to simulate human-like typing
def slow_typing(element, text):
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(0.1, 0.3))  # Random delay between keystrokes

# Function to scroll into view before clicking
def scroll_and_click(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
    time.sleep(random.uniform(1.0, 1.1))
    element.click()

# Load data from Excel
df = pd.read_excel(r"D:\google\class7.xlsx")  # Update with your correct path

# Setup WebDriver
options = Options()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Open the Google Form
form_url = "https://forms.gle/YehRZMniAMF7Bxex6"
driver.get(form_url)
driver.maximize_window()  # Maximize the browser window

# Wait for the form to load
time.sleep(random.uniform(2.5, 4.0))

# Loop through each student entry
for index, row in df.iterrows():
    # Fill Email
    email_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/div/div[2]/div/div/div/div[2]/div/div[1]/div/div[1]/input")
    slow_typing(email_field, "test@example.com")  # Modify if email is in Excel
    time.sleep(random.uniform(1.0, 1.1))

    # Click the "Next" button
    wait = WebDriverWait(driver, 10)
    next_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Next']")))
    scroll_and_click(driver, next_button)

    # Wait for the next page to load
    time.sleep(random.uniform(2.5, 4.0))

    # Fill Name
    name_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input")
    slow_typing(name_field, row["Name"])
    time.sleep(random.uniform(1.0, 1.1))

    # Select Gender
    gender_xpath = f"//span[text()='{row['Gender']}']"
    gender_button = driver.find_element(By.XPATH, gender_xpath)
    scroll_and_click(driver, gender_button)

    # Fill Father's Name
    father_name_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input")
    slow_typing(father_name_field, row["Father's Name"])
    time.sleep(random.uniform(1.0, 1.1))

    actions = ActionChains(driver)

    # Move focus to the dropdown
    father_name_field.send_keys(Keys.TAB)
    time.sleep(random.uniform(1.0, 1.1))

    # Press "B" key using ActionChains
    actions.send_keys("B").perform()
    time.sleep(random.uniform(1.0, 1.1))

    # Fill UDISE Code
    udise_code_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[1]/div/div[1]/input")
    slow_typing(udise_code_field, "10140807501")
    time.sleep(random.uniform(1.0, 1.1))

    # Fill School Name
    school_name_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/div/div[2]/div[7]/div/div/div[2]/div/div[1]/div/div[1]/input")
    slow_typing(school_name_field, "UMS BHIKHANPUR KOTHI")
    time.sleep(random.uniform(1.0, 1.1))

    # Select Class
    class_radio_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//div[@role="radio" and @aria-label="7"]'))
    )
    scroll_and_click(driver, class_radio_button)

    # Fill Class Roll Number
    class_roll_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/div/div[2]/div[9]/div/div/div[2]/div/div[1]/div/div[1]/input")
    slow_typing(class_roll_field, str(row["Class Roll Number"]))  # Convert to string
    time.sleep(random.uniform(1.0, 1.1))

    # Move focus to the dropdown
    class_roll_field.send_keys(Keys.TAB)
    time.sleep(random.uniform(1.1, 2.5))

    # Press "MU" key using ActionChains
    actions.send_keys("MU").perform()
    time.sleep(random.uniform(1.0, 1.1))

    # Click "Next" button
    next_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@role='button']//span[contains(text(), 'Next')]"))
    )
    scroll_and_click(driver, next_button)

    # Wait for the next page to load
    time.sleep(random.uniform(2.5, 4.0))

    answers = [
        "D) Calculator",                                     # Q1
        "A) Uniform Resource Locator",                       # Q2
        "B) To perform calculations",                        # Q3
        "d) Photoshop",                                      # Q4
        "C) Hard Disk",                                      # Q5
        "B) To manage hardware and software resources",      # Q6
        "B) Output device",                                  # Q7
        "A) Dropbox",                                        # Q8
        "B) It displays websites",                           # Q9
        "A) Ctrl + C"                                        # Q10
    ]

    for i, answer in enumerate(answers):
        option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//span[contains(text(), '{answer}')]"))
        )
        
        # Scroll into view before clicking
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", option)
        time.sleep(random.uniform(1.0, 1.1))  # Random delay to mimic human behavior
        
        option.click()
        
        # Scroll down after answering Question 4
        if i == 3:  # Index 3 = Question 4
            driver.execute_script("window.scrollBy(0, 500);")  # Scroll down by 500 pixels
            time.sleep(random.uniform(1.0, 1.1))  # Random delay

    # Submit form
    final_submit_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='Submit']"))
    )
    scroll_and_click(driver, final_submit_button)

    time.sleep(random.uniform(2.5, 4.0))

    df.at[index, "Status"] = "Done"
    df.to_excel(r"D:\google\class7_progress.xlsx", index=False)
    print(f"S.No {index + 1} - {row['Name']} marked as Done")

    driver.get(form_url)  # Reload the form for the next student
    time.sleep(random.uniform(2.5, 4.0))
