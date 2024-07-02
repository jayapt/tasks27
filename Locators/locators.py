"""
Locator.py
"""
from selenium.webdriver.common.by import By

class WebLocators:


   def __init__(self):
       self.usernameLocator = "username"
       self.passwordLocator = "password"
       self.buttonLocator = "//button[@type='submit']"
       self.dropdownMenuLocator ='//*[@id="app"]/div[1]/div[1]/header/div[1]/div[2]/ul/li/span'
       self.logoutLocator ='//*[@id="app"]/div[1]/div[1]/header/div[1]/div[2]/ul/li/ul/li[4]/a'

   def enterText(self, driver, locator, textValue):
       element = driver.find_element(by=By.ID, value=locator)
       element.clear()
       element.send_keys(textValue)

   def clickButton(self, driver, locator):
       driver.find_element(by=By.ID, value=locator).click()
