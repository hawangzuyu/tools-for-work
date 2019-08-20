from selenium import webdriver
from selenium.webdriver.firefox.options import Options

options = Options()
options.headless = True
driver = webdriver.Firefox(options=options,
                           executable_path=r'C:\Users\Administrator\AppData\Local\Programs\Python\Python37\geckodriver.exe')
driver.get("https://www.baidu.com/")
driver.save_screenshot('baidu.png')
print("Headless Firefox Initialized")
driver.quit()
