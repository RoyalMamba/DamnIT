import logging
import os
import time
from concurrent.futures import ThreadPoolExecutor
from dotenv import load_dotenv
from selenium import webdriver
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException, StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl

class DAMTagger:
    def __init__(self):
        self.logger = self.setup_logger()
        self.load_environment_variables()
        self.initialize_driver()
        self.login()

    def setup_logger(self):
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', filename='damnit.log')
        return logging.getLogger(__name__)

    def load_environment_variables(self):
        load_dotenv()
        self.website_url = "https://stg.tagsmarter.com/auth/login"
        self.user_email = os.getenv('EMAIL')
        self.user_password = os.getenv('PASSWORD')
        self.folder_name = os.getenv('FOLDER_NAME')
        self.event_name = os.getenv('EVENT_NAME')

    def initialize_driver(self):
        driver_options = webdriver.ChromeOptions()
        driver_options.add_argument("--disable-extensions")
        driver_options.add_argument("--disable-gpu")
        self.driver = webdriver.Chrome(options=driver_options)
        self.driver.maximize_window()

    def login(self):
        self.driver.get(self.website_url)
        email_field = self.driver.find_element(By.CSS_SELECTOR, 'app-login input[formcontrolname="email"]')
        password_field = self.driver.find_element(By.CSS_SELECTOR, 'app-login input[formcontrolname="password"]')
        login_button = self.driver.find_element(By.XPATH, "//button[contains(text(), 'Log in')]")
        email_field.send_keys(self.user_email)
        password_field.send_keys(self.user_password)
        login_button.click()
        WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="primarynavigation"]/ul/li[4]/div/a')))
        self.logger.info('Successfully logged in to the DAM PORTAL..')

    def navigate_to_folder(self):
        Asset = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH , '//*[@id="primarynavigation"]/ul/li[4]/div/a')))
        Asset.click()
        folder_element = WebDriverWait(self.driver, 20).until(EC.visibility_of_element_located((By.XPATH, f"//span[@class='item-folder-name' and text()='{self.folder_name}']")))
        folder_element.click()
        self.logger.info(f'Navigated to the folder {self.folder_name} and starting the tagging process...')

    def change_page_size(self, size):
        page_size_dropdown = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="nav-home"]/div/div[2]/div/app-assets-master/div[1]/div/div/div[1]/div/p-dropdown')))
        page_size_dropdown.click()
        option_xpath = f'//ul[@class="p-dropdown-items"]/li//span[text()="{size}"]'
        option_element = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, option_xpath)))
        option_element.click()

    def locate_asset_card(self, image_name):
        asset_xpath = f"//img[contains(@src, '{image_name}')]"
        assetCard = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, asset_xpath)))
        self.driver.execute_script("arguments[0].scrollIntoView();", assetCard)
        return assetCard

    def tag_asset(self, image_name, tags):
        assetCard = self.locate_asset_card(image_name)

        checkbox_xpath = f"{assetCard}/../following-sibling::div//input[@type='checkbox']"
        checkbox_element = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, checkbox_xpath)))
        self.driver.execute_script("arguments[0].click();", checkbox_element)

        self.trigger_context_menu(assetCard)
        self.click_on_tag_asset()

        self.select_event_from_dropdown()
        self.enter_tags(tags)
        self.complete_tagging()

    def trigger_context_menu(self, assetCard):
        js_script = "var evt = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, view: window }); arguments[0].dispatchEvent(evt);"
        self.driver.execute_script(js_script, assetCard)

    def click_on_tag_asset(self):
        js_script = """var contextMenu = document.getElementById('context-menu1');
                       var tagOption = contextMenu.querySelector('ul li:nth-child(3) a span');
                       tagOption.click();"""
        self.driver.execute_script(js_script)

    def select_event_from_dropdown(self):
        event_dropdown_xpath = "//div[@id='nav-home']//app-assets-master//p-dialog//p-dropdown/div/div[2]"
        eventDropdown = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, event_dropdown_xpath)))
        
        attempts = 0
        while attempts < 3:
            try:
                self.driver.execute_script("arguments[0].click();", eventDropdown)
                self.driver.execute_script(js_click_option, self.event_name)
                self.logger.info(f'Selected the {self.event_name} from the dropdown in {attempts} attempts.')
                break
            except StaleElementReferenceException:
                attempts += 1

    def enter_tags(self, tags):
        tag_input_xpath = '//div[@id="nav-home"]//app-assets-master//p-dialog[2]//form//div[4]//p-chips//input'
        tag_input = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, tag_input_xpath)))
        
        for tag in tags:
            js_script = f'arguments[0].value = "{tag.strip()}";'
            self.driver.execute_script(js_script, tag_input)
            tag_input.send_keys(Keys.RETURN)

    def complete_tagging(self):
        completed_button = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.NAME, "assetStatus")))
        self.driver.execute_script("arguments[0].click();", completed_button)

        approve_button_xpath = "//button[contains(text(), 'Approve')]"
        approveButton = WebDriverWait(self.driver, 20).until(EC.visibility_of_element_located((By.XPATH, approve_button_xpath)))
        self.driver.execute_script("arguments[0].click();", approveButton)

        WebDriverWait(self.driver, 20).until(EC.invisibility_of_element_located((By.XPATH, "//div[@class='overlay']")))
        time.sleep(2)
        self.logger.info('Finished Tagging Assets!')

    def run(self):
        try:
            self.navigate_to_folder()
            finished_tagging = 0
            for totalTagged, iterfiles in enumerate(filesNtags, start=1):
                try:
                    imageName = iterfiles[0]
                    self.change_page_size(1000)
                    self.tag_asset(imageName, iterfiles[1])
                    finished_tagging += 1
                except ElementClickInterceptedException as e:
                    self.logger.error(f"Element click intercepted: {e}")
                    continue
                except TimeoutException as e:
                    self.logger.error(f"Timeout exception !!! \n\n\n\n ")
                    continue
                except Exception as e:
                    self.logger.error(f"Unable to Tag Asset No- {totalTagged}. An error occurred inside the loop: {e}")
                    continue

        except Exception as e:
            self.logger.critical(f"An error occurred outside the loop: {e}")
        except KeyboardInterrupt:
            self.logger.error(f'Unable to complete the tagging process : KeyboardInterrupt')
        finally:
            self.logger.info(f'Done with the tagging process of {finished_tagging} Assets !!!\n\n\n\n')
            self.driver.quit()

if __name__ == "__main__":
    workbook = openpyxl.load_workbook('Files_and_Tags.xlsx')
    sheet = workbook.active
    filesNtags = [filename for filename in sheet.iter_rows(values_only=True, min_row=2)]
    dam_tagger = DAMTagger()
    dam_tagger.run()
