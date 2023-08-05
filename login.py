import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl

workbook = openpyxl.load_workbook('Files_and_Tags.xlsx')
sheet = workbook.active
filesNtags = [filename for filename in sheet.iter_rows(values_only=True, min_row=5)]


try:
    driver = webdriver.Chrome()  # You can replace 'Chrome' with 'Firefox' if using Firefox

    # Website URL
    website_url = "https://stg.tagsmarter.com/auth/login"
    eventName = input("Enter the name of the Event \nMake sure it is precise and concise:   ")  # Replace with the actual option text

    # Navigate to the login page
    driver.get(website_url)
    driver.maximize_window()

    # Login credentials

    # Find and fill in the email and password fields
    email_field = driver.find_element(By.CSS_SELECTOR, 'app-login input[formcontrolname="email"]')
    password_field = driver.find_element(By.CSS_SELECTOR, 'app-login input[formcontrolname="password"]')

    login_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Log in')]")

    login_button.click()

    # Wait for Assets tab to load and click on it
    Asset = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH , '//*[@id="primarynavigation"]/ul/li[4]/div/a')))
    Asset.click()

    # Folder name from Excel
    folder_name = 'Anthony Albanese Australia'

    # Click on the specified folder
    folder_element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, f"//span[@class='item-folder-name' and text()='{folder_name}']")))
    folder_element.click()

    # Wait for assets to load
    time.sleep(1)

    # imageName = "Devendra_Fadnavis_sarkarnama_2F2023_08_2F0055e5bf_c63e_4120_9382_0f3bf8fd090f_2Fdevendra_Fadnvis_sangram_Thopate.jpg"
    for iterfiles in filesNtags:
        imageName = iterfiles[0]

        # Find the assetCard (image) element
        assetCard = driver.find_element(By.XPATH, f"//img[contains(@src,'{imageName}')]")

        # Scroll the assetCard (image) element into view
        driver.execute_script("arguments[0].scrollIntoView();", assetCard)

        # Find the corresponding checkbox element based on the assetCard (image) element
        checkbox_element = assetCard.find_element(By.XPATH, "../following-sibling::div//input[@type='checkbox']")

        driver.execute_script("arguments[0].click();", checkbox_element)
        time.sleep(2)
        rightClickOnAsset = ActionChains(driver)
        rightClickOnAsset.context_click(assetCard).perform()
        clickOnTag = driver.find_element(By.XPATH,'//*[@id="context-menu1"]/ul/li[3]/a/span')
        clickOnTag.click()

        eventDropdown = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, "//div[@id='nav-home']//app-assets-master//p-dialog//p-dropdown")))
        eventDropdown.click()


        options = WebDriverWait(driver,10).until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, "ul.p-dropdown-items li")))

        for option in options:
            if option.text == eventName:
                option.click()
                break
        
        tag_input = driver.find_element(By.XPATH, '//div[@id="nav-home"]//app-assets-master//p-dialog[2]//form//div[4]//p-chips//input')

        tags = iterfiles[1].split(',')
        for tag in tags:
            driver.execute_script(f'arguments[0].value = "{tag}";', tag_input)
            tag_input.send_keys(Keys.RETURN)

        completedButton = driver.find_element(By.NAME, "assetStatus")
        completedButton.click()

        approveButton =driver.find_element(By.XPATH, "//button[contains(text(), 'Approve')]")
        approveButton.click()
        time.sleep(1)

except : 
    pass
finally:
    time.sleep(20)
    driver.quit()
