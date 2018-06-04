from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains



option = webdriver.ChromeOptions()
option.add_argument('disable-infobars')
driver = webdriver.Chrome(chrome_options=option)
input('dd')
driver.get('http://www.baidu.com')
driver.maximize_window()
a=driver.find_element_by_css_selector('#su')
ActionChains(driver).send_keys('敬江'+Keys.ENTER).perform()

input(...)
driver.quit()