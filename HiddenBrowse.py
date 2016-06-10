#!/usr/bin/env python
import os
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
#from pyvirtualdisplay import Display

#display = Display(visible=False)
#display.start()

#chromeOptions = webdriver.ChromeOptions()
#chromeOptions.add_experimental_option("prefs", {'safebrowsing.enabled':1})


driver = webdriver.Remote(desired_capabilities=webdriver.DesiredCapabilities.HTMLUNITWITHJS)
driver.get('http://www.google.com')
driver.maximize_window()

print driver.title

