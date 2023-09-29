import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException, StaleElementReferenceException
from dotenv import load_dotenv
import time
import openpyxl
import os
import logging
from concurrent.futures import ThreadPoolExecutor
from retryrightclick import rightclick

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', filename='damnit.log')
logger = logging.getLogger(__name__)
workbook = openpyxl.load_workbook('Files_and_Tags.xlsx')
sheet = workbook.active
filesNtags = [filename for filename in sheet.iter_rows(values_only=True, min_row=2)]
load_dotenv()
try:
    website_url = "https://stg.tagsmarter.com/auth/login"
    # user_email = input("Enter the EMAIL of the user:    ")
    # user_password = input("Enter the PASSWORD of the user: ")


    # folder_name = input("Enter the name of the FOLDER:  ")
    # eventName = input("Enter the name of the EVENT:   ")


    # Website URL and login credentials
    user_email = os.getenv('EMAIL')
    user_password= os.getenv('PASSWORD')
    folder_name = os.getenv('FOLDER_NAME')
    eventName = os.getenv('EVENT_NAME')

    driveroptions = webdriver.ChromeOptions()
    # driveroptions.add_argument('--headless')
    driveroptions.add_argument("--disable-extensions")
    driveroptions.add_argument("--disable-gpu")
    driver = webdriver.Chrome(options=driveroptions)

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
    driver.execute_script("arguments[0].click();", Asset)
    logger.info(f'Successfully logged in to the DAM PORTAL..')
    # Click on the specified folder
    folder_element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, f"//span[@class='item-folder-name' and text()='{folder_name}']")))
    driver.execute_script("arguments[0].click();", folder_element)
    logger.info(f'Navigated to the folder {folder_name} and starting the tagging process...')
    # Wait for assets to load
    # time.sleep(1)
    finished_tagging = 0
    for totalTagged, iterfiles in enumerate(filesNtags,start=1):
        try:
            imageName = iterfiles[0]

            #Change the size of the page to 1000 assets 
            if totalTagged > -1 : 
                size = '1000'
                pageSize = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="dropdownMenuLink"]')))
                ## driver.execute_script("window.scrollTo({top: 0, behavior: 'instant'});")
                ## pageSize.click()
                driver.execute_script("arguments[0].click();", pageSize)
                # pageOptions = WebDriverWait(driver  ,10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul.p-dropdown-items li")))
                js_click_option = """
                            var selectedSize = arguments[0];
                            var options = document.querySelectorAll('ul.dropdown-menu.show li a');
                            for (var i = 0; i < options.length; i++) {
                                if (options[i].textContent.trim() === selectedSize.toString()) {
                                    options[i].click();
                                    break;
                                }
                            }
                            """
                driver.execute_script(js_click_option,size)

                # for page in pageOptions:
                #     if page.text == '1000':
                #         # page.click()
                #         driver.execute_script("arguments[0].click();", page)
                #         break

            # Find the assetCard (image) element
            assetCard = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, f"//img[contains(@src,'{imageName}')]")))
            # Scroll the assetCard (image) element into view
            driver.execute_script("arguments[0].scrollIntoView();", assetCard)
            logger.info(f'Successfully located and found the Asset {imageName}')

            # Uncheck any previously selected checkboxes 
            uncheck_previous_checkboxes_js = """
                var checkboxes = document.querySelectorAll('input[type="checkbox"]');
                for (var i = 0; i < checkboxes.length; i++) {
                    checkboxes[i].checked = false;
                }
            """
            driver.execute_script(uncheck_previous_checkboxes_js)  
            logger.info(f'Cleared all the checkboxes if present for the iteration number: {totalTagged}')          

            # Find the corresponding checkbox element based on the assetCard (image) element
            checkbox_element = assetCard.find_element(By.XPATH, "../following-sibling::div//input[@type='checkbox']")
            driver.execute_script("arguments[0].click();", checkbox_element)
#             time.sleep(1)
            logger.info(f'Clicked on the check box for asset : {imageName}')

            driver.execute_script("var evt = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, view: window }); arguments[0].dispatchEvent(evt);", assetCard)

            js_click_on_tag = """var contextMenu = document.getElementById('context-menu1');
                                var tagOptions = contextMenu.querySelectorAll('ul li a span');
                                var found = false;
                                for (var i = 0; i < tagOptions.length; i++) {
                                    if (tagOptions[i].textContent.trim() === 'Tag Asset') {
                                        tagOptions[i].click();
                                        found = true;
                                        break;
                                    }
                                }
                                return found;
                            """
            tag_asset_clicked = driver.execute_script(js_click_on_tag)

            if tag_asset_clicked:
                logger.info(f'Successfully Clicked on the "Tag Asset" for Asset: {imageName}')
            else:
                logger.warning('Unable to find and click "Tag Asset" option on the context menu, Retrying !!!')
                rightclick(folder_name,imageName,totalTagged,driver,logger)
    

            # time.sleep(1)
            eventDropdown = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@id='nav-home']//app-assets-master//p-dialog//p-dropdown/div/div[2]")))
            attempts = 0
            while attempts < 3:
                try:
                    # eventDropdown.click()
                    driver.execute_script("arguments[0].click();", eventDropdown)

                    # Use JavaScript to find and click the desired option based on the variable
                    js_click_option = """
                                    var options = document.querySelectorAll('ul.p-dropdown-items li');
                                    for (var i = 0; i < options.length; i++) {
                                        if (options[i].textContent.trim() === arguments[0]) {
                                            options[i].click();
                                            break;
                                        }
                                    }
                                """
                    driver.execute_script(js_click_option, eventName)
                    logger.info(f'Selected the {eventName} from the dropdown in {attempts} attemps.')
                    break
                except StaleElementReferenceException:
                    attempts += 1
                    # time.sleep(1)

            
            tag_input = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="nav-home"]/div/div[2]/div/app-assets-master/p-dialog[3]/div/div/div[2]/div/div/div/div/div[2]/div/div[2]/div/div/div/form/div[1]/div[1]/div[4]/div/div/div/p-chips/div/ul/li/input')))
            tags = iterfiles[1].split(',')
            
            for tag in tags:
                driver.execute_script(f'arguments[0].value = "{tag.strip()}";', tag_input)
                tag_input.send_keys(Keys.RETURN)
            logger.info(f'Done with the tagging process of the Asset {totalTagged} Image: {imageName}')

            completedButton = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.NAME, "assetStatus")))
            # completedButton.click()
            driver.execute_script("arguments[0].click();", completedButton)

            approveButton = WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.XPATH, "//button[contains(text(), 'Approve')]")))
            # approveButton.click()
            driver.execute_script("arguments[0].click();", approveButton)
            logger.info(f'Finished Tagging {totalTagged} Assets !')
            WebDriverWait(driver, 20).until(EC.invisibility_of_element_located((By.XPATH, "//div[@class='overlay']")))
            finished_tagging+= 1
            time.sleep(2)

        
        except ElementClickInterceptedException as e:
            logger.error(f"Element click intercepted: {e}")
            continue
        except TimeoutException as e:
            logger.error(f"Timeout exception !!! \n\n\n\n ")
            continue
        except StaleElementReferenceException:
            # time.sleep(2)
            # rightclick(folder_name,imageName,totalTagged,driver,logger)
            continue
        except Exception as e:
            logger.error(f"Unable to Tag Asset No- {totalTagged}.An error occurred inside the loop: {e}")
            continue

except Exception as e:
    logger.critical(f"An error occurred outside the loop: {e}")
except KeyboardInterrupt:
    logger.error(f'Unable to complete the tagging process : KeyboardInterrupt')
finally:
    # time.sleep(25)
    logger.info(f'Done with the tagging process of {finished_tagging} Assets !!!\n\n\n\n')
    driver.quit()
