#!/usr/bin/env python
# coding: utf-8
def amazon_affiliates_01_download_report():
    from selenium import webdriver
    from webdriver_manager.firefox import GeckoDriverManager
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support.expected_conditions import presence_of_element_located
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.firefox.service import Service
    from selenium.webdriver.firefox.options import Options
    from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
    from selenium.webdriver.common.action_chains import ActionChains
    import time
    
    # Set File path and creds
    download_dir = r"/.../amazon_affiliate_daily_dl"
    user = "[]"
    password = "[]"

    # Set driver options
    options = Options()
    options.add_argument=("--headless")
    options.add_argument=("browser.preferences.instantApply",True)
    options.add_argument=("browser.preferences.instantApply",True)
    options.add_argument=("browser.helperApps.neverAsk.saveToDisk", "text/plain, application/octet-stream, application/binary, text/csv, application/csv, application/excel, text/comma-separated-values, text/xml, application/xml")
    options.add_argument=("browser.helperApps.alwaysAsk.force",False)
    options.add_argument=("browser.download.manager.showWhenStarting",False)
    options.add_argument=("browser.download.folderList",0)
    options.add_argument=("browser.download.dir", download_dir)
    
    #Launch the website, sign in, navigate to yesterday's dataset and download
    with webdriver.Firefox() as driver:
        driver.get("https://affiliate-program.amazon.com/")
        driver.find_element(by="id", value="a-autoid-0").click()
        driver.find_element(by="id", value="ap_email").send_keys(user)
        driver.find_element(by="id", value="continue").click()
        driver.find_element(by="id", value="ap_password").send_keys(password)
        driver.find_element(by="id", value="signInSubmit").click()
        driver.get("https://affiliate-program.amazon.com/home/reports")
        driver.find_element(by="id", value="ac-report-download-launcher").click()
        actions = ActionChains(driver)
        wait = WebDriverWait(driver, 5)
        dropdown = wait.until(EC.presence_of_element_located((By.ID, "ac-daterange-popover-report-download-timeInterval")))
        actions.click(dropdown)
        actions.perform()

        # Hover over the Date Drop Down, define actions to perform
        actions = ActionChains(driver)
        wait = WebDriverWait(driver, 180)
        # dropdown = wait.until(EC.presence_of_element_located(by="id", value="ac-daterange-popover-report-download-timeInterval"))
        radio_button = driver.find_element(by='xpath', value = "//div[@id='ac-daterange-radio-report-download-timeInterval-yesterday' and contains(@class, 'a-radio')]/label/i[contains(@class, 'a-icon')]")    
        apply_button = driver.find_element(by="id", value="ac-daterange-ok-button-report-download-timeInterval-announce")
        gen_reports_button = driver.find_element(by="id", value="ac-reports-download-generate-announce")
        download_button = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='ac-report-download-container' and contains(@class, 'ac-modal-container')]/div[contains(@class, 'ac-modal-tbl')]/div[contains(@class, 'ac-report-download-tbl-content')]/table[contains(@class, 'a-dtt-table')]/tbody[contains(@class, 'a-dtt-tbody')]/tr/td[4]/a")))
    
        # Perform the actions
        # actions.pause(5)
        # actions.move_to_element(dropdown)
        actions.pause(2)
        actions.click(radio_button)
        actions.click(apply_button)
        actions.click(gen_reports_button)
        actions.perform()
        actions.reset_actions()
        actions.click(download_button)
        actions.pause(5)
        actions.perform()
    
        # Give the site a few seconds to begin the download
        WebDriverWait(driver, 10)
        driver.close()
    
if __name__ == "__main__":
    amazon_affiliates_01_download_report()
