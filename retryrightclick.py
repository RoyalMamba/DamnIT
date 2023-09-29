# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def rightclick(folder_name,imageName,totalTagged,driver,logger):
    folder_element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, f"//span[@class='item-folder-name' and text()='{folder_name}']")))
    driver.execute_script("arguments[0].click();", folder_element)
    driver.execute_script("arguments[0].click();", folder_element)
    logger.info(f'Navigated to the folder {folder_name} and starting the tagging process...')
    if totalTagged > -1 : 
        size = '1000'
        pageSize = WebDriverWait(driver,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="nav-home"]/div/div[2]/div/app-assets-master/div[1]/div/div/div[1]/div/p-dropdown/div/div[2]')))
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
        rightclick(folder_element,imageName,totalTagged,driver,logger)