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
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

# Load data from Excel
df = pd.read_excel(r"D:\google\students.xlsx")  # Update with your correct path

# Setup WebDriver
options = Options()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Open the Google Form
form_url = "https://forms.gle/SRz32a2hyo4bLfGy7"
driver.get(form_url)
driver.maximize_window()  # Maximize the browser window


# Wait for the form to load
time.sleep(3)

# Loop through each student entry
for index, row in df.iterrows():
    # Find the email input field and enter email
    email_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/div/div[2]/div/div/div/div[2]/div/div[1]/div/div[1]/input")
    email_field.send_keys("test@example.com")  # You can modify to take email from Excel
    time.sleep(1)

    # Click the "Next" button
    wait = WebDriverWait(driver, 10)
    next_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Next']")))
    next_button.click()
    time.sleep(1)

    # Wait for the next page to load
    time.sleep(3)

    # Fill Name
    name_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input")
    name_field.send_keys(row["Name"])
    time.sleep(1)

    gender_xpath = f"//span[text()='{row['Gender']}']"
    gender_button = driver.find_element(By.XPATH, gender_xpath)
    gender_button.click()
    time.sleep(1)

    father_name_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input")
    father_name_field.send_keys(row["Father's Name"])
    time.sleep(1)

    actions = ActionChains(driver)

    # Move focus to the dropdown
    father_name_field.send_keys(Keys.TAB)
    time.sleep(1)  # Wait to ensure focus shift

    # Press "B" key using ActionChains
    actions.send_keys("B").perform()
    time.sleep(1)  # Allow time for selection

    udise_code_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[1]/div/div[1]/input")
    udise_code_field.send_keys("10140807501")  # Replace with actual UDISE code
    time.sleep(1)

    school_name_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/div/div[2]/div[7]/div/div/div[2]/div/div[1]/div/div[1]/input")
    school_name_field.send_keys("UMS BHIKHANPUR KOTHI")  # Replace with actual school name
    time.sleep(1)

        # Wait for the element to be clickable
    class_radio_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//div[@role="radio" and @aria-label="06"]'))
    )

    # Click the element
    class_radio_button.click()

    class_roll_field = driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/div/div[2]/div[9]/div/div/div[2]/div/div[1]/div/div[1]/input")
    class_roll_field.send_keys(row["Class Roll Number"])  # Get value from Excel
    time.sleep(1)

    actions = ActionChains(driver)

    # Move focus to the dropdown
    class_roll_field.send_keys(Keys.TAB)
    time.sleep(1)  # Wait to ensure focus shift

    # Press "B" key using ActionChains
    actions.send_keys("MU").perform()
    time.sleep(1)  # Allow time for selection

    # Wait for the Next button to be clickable
    next_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@role='button']//span[contains(text(), 'Next')]"))
    )

    # Click the Next button
    next_button.click()
    #page 3
    time.sleep(3)

    answers = [
        "c) Hardware",               # Q1
        "a) Speaker",                # Q2
        "b) Google Chrome",          # Q3
        "b) Uniform Resource Locator",  # Q4
        "d) CPU",                    # Q5
        "a) Hyper Text Markup Language",  # Q6
        "c) Hard Disk",              # Q7
        "c) Flash Drive",            # Q8
        "d) YouTube",                # Q9
        "c) Personal Computer"       # Q10
    ]

    for i, answer in enumerate(answers):
        option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//span[contains(text(), '{answer}')]"))
        )
        option.click()

        # Scroll down after answering Question 4
        if i == 3:  # Index 3 = Question 4
            driver.execute_script("window.scrollBy(0, 500);")  # Scroll down by 500 pixels

    final_submit_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[text()='Submit']"))
    )
    final_submit_button.click()

    time.sleep(3)  # Allow time for submission to process

    df.at[index, "Status"] = "Done"  # Mark student as 'Done'
    df.to_excel(r"D:\google\students_progress.xlsx", index=False)  # Save progress after each submission
    print(f"S.No {index + 1} - {row['Name']} marked as Done")


    driver.get(form_url)  # Reload the form
    time.sleep(3)  # Wait for the form to fully load
