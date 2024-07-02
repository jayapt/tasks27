"""
TEST MAIN
"""
from Data import data
from Locators import locators

import pytest
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC



class TestOrangeHRM():


    @pytest.fixture
    def boot(self):
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        self.driver.get(data.WebData().url)
        self.driver.maximize_window()
        self.wait = WebDriverWait(self.driver, 10)
        yield
        self.driver.quit()

    @pytest.mark.html
    def testLogin(self, boot):

        try:
            # Username - 2
            # Password - 3
            # Test Results - 7
            locator = locators.WebLocators()
            sheet = data.WebData()
            for row in range(2, data.WebData().rowCount() + 1):
                username = sheet.readData(row, 2)
                password = sheet.readData(row, 3)
                # To get current date , time & username in the excel.
                sheet.writeData(row, 4, datetime.now().strftime("%Y-%m-%d"))
                sheet.writeData(row, 5, datetime.now().strftime("%H:%M:%S"))
                sheet.writeData(row, 6, "Jaya")
                self.wait.until(EC.presence_of_element_located((By.NAME, locator.usernameLocator))).send_keys(username)
                self.wait.until(EC.presence_of_element_located((By.NAME, locator.passwordLocator))).send_keys(password)
                self.wait.until(EC.presence_of_element_located((By.XPATH, locator.buttonLocator))).click()
                # Check whether we have loged in to the webpage successfully or not

                currenturl = self.driver.current_url
                if (currenturl == data.WebData().dashboardURL):

                    print("login successfully")
                    sheet.writeData(row, 7, "Passed")


                    self.wait.until(EC.presence_of_element_located((By.XPATH, locator.dropdownMenuLocator))).click()
                    self.wait.until(EC.presence_of_element_located((By.XPATH, locator.logoutLocator))).click()
                else:
                    print("Login failed")
                    sheet.writeData(row, 7, "Failed")


        except Exception as e:
            print(e)