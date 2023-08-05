import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException, StaleElementReferenceException


import time
import openpyxl

workbook = openpyxl.load_workbook('Files_and_Tags.xlsx')
sheet = workbook.active
filesNtags = [filename for filename in sheet.iter_rows(values_only=True, min_row=2)]

try:
    driver = webdriver.Chrome()

    # Website URL and login credentials
    website_url = "https://stg.tagsmarter.com/auth/login"
    user_email = input("Enter the EMAIL of the user:    ")
    user_password = input("Enter the PASSWORD of the user: ")


    folder_name = input("Enter the name of the FOLDER:  ")
    eventName = input("Enter the name of the EVENT:   ")


    # Navigate to the login page
    driver.get(website_url)
    driver.maximize_window()

    # Find and fill in the email and password fields
    email_field = driver.find_element(By.CSS_SELECTOR, 'app-login input[formcontrolname="email"]')
    password_field = driver.find_element(By.CSS_SELECTOR, 'app-login input[formcontrolname="password"]')
    login_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Log in')]")

    email_field.send_keys(user_email)
    password_field.send_keys(user_password)
    login_button.click()

    # Wait for Assets tab to load and click on it
    Asset = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH , '//*[@id="primarynavigation"]/ul/li[4]/div/a')))
    Asset.click()

    # Click on the specified folder
    folder_element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, f"//span[@class='item-folder-name' and text()='{folder_name}']")))
    folder_element.click()
    # Wait for assets to load
    time.sleep(1)

    for totalTagged, iterfiles in enumerate(filesNtags):
        try:
            imageName = iterfiles[0]

            #Change the size of the page to 1000 assets 
            if totalTagged > 49 : 
                pageSize = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="nav-home"]//app-assets-master//p-dropdown')))
                driver.execute_script("window.scrollTo({top: 0, behavior: 'instant'});")
                pageSize.click()
                pageOptions = WebDriverWait(driver,10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul.p-dropdown-items li")))
                for page in pageOptions:
                    if page.text == '1000':
                        page.click()
                        break

            # Find the assetCard (image) element
            assetCard = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, f"//img[contains(@src,'{imageName}')]")))
            # Scroll the assetCard (image) element into view
            driver.execute_script("arguments[0].scrollIntoView();", assetCard)

            # Find the corresponding checkbox element based on the assetCard (image) element
            checkbox_element = assetCard.find_element(By.XPATH, "../following-sibling::div//input[@type='checkbox']")
            driver.execute_script("arguments[0].click();", checkbox_element)
#             time.sleep(1)

            try:
                # Trigger the context menu using JavaScript
                driver.execute_script("var evt = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, view: window }); arguments[0].dispatchEvent(evt);", assetCard)

#                 clickOnTag = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="context-menu1"]/ul/li[3]/a/span')))
                js_click_on_tag = """var contextMenu = document.getElementById('context-menu1');
                                    var tagOption = contextMenu.querySelector('ul li:nth-child(3) a span');
                                    tagOption.click();
                                    """
                driver.execute_script(js_click_on_tag)

            except StaleElementReferenceException:
                # Re-locate the assetCard element
                assetCard = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, f"//img[contains(@src,'{imageName}')]")))
                # Trigger the context menu using JavaScript again
                driver.execute_script("var evt = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, view: window }); arguments[0].dispatchEvent(evt);", assetCard)

#                 clickOnTag = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="context-menu1"]/ul/li[3]/a/span')))
                js_click_on_tag = """var contextMenu = document.getElementById('context-menu1');
                                    var tagOption = contextMenu.querySelector('ul li:nth-child(3) a span');
                                    tagOption.click();
                                    """
                driver.execute_script(js_click_on_tag)
            # time.sleep(1)
            eventDropdown = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@id='nav-home']//app-assets-master//p-dialog//p-dropdown")))
            
            attempts = 0
            while attempts < 3:
                try:
                    eventDropdown.click()
                    options = WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, "ul.p-dropdown-items li")))
                    for option in options:
                        if option.text == eventName:
                            option.click()
                            break
                    break
                except StaleElementReferenceException:
                    attempts += 1
                    time.sleep(1)

            
            tag_input = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.XPATH, '//div[@id="nav-home"]//app-assets-master//p-dialog[2]//form//div[4]//p-chips//input')))
            tags = iterfiles[1].split(',')
            
            for tag in tags:
                driver.execute_script(f'arguments[0].value = "{tag.strip()}";', tag_input)
                tag_input.send_keys(Keys.RETURN)

            completedButton = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.NAME, "assetStatus")))
            completedButton.click()

            approveButton = WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.XPATH, "//button[contains(text(), 'Approve')]")))
            approveButton.click()
            print(f'Finished Tagging {totalTagged} Assets !')
            time.sleep(3)

        
        except ElementClickInterceptedException as e:
            print(f"Element click intercepted: {e}")
            continue
        except TimeoutException as e:
            print(f"Timeout exception: {e}")
            continue
        except Exception as e:
            print(f"An error occurred inside the loop: {e}")
            continue

except Exception as e:
    print(f"An error occurred outside the loop: {e}")
finally:
    # time.sleep(25)
    driver.quit()
