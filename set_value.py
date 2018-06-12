from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains



option = webdriver.ChromeOptions()
option.add_argument('disable-infobars')
driver = webdriver.Chrome(chrome_options=option)
input('dd')
driver.get('http://www.baidu.com')
driver.maximize_window()
a=driver.find_element_by_css_selector('#su').location_once_scrolled_into_view
ActionChains(driver).send_keys('敬江'+Keys.ENTER).perform()

input(...)
driver.quit()

a=driver.find_elements_by_css_selector('table.datagrid-btable')[-1]

driver.find_element_by_css_selector('button[name="batchEdit"]').click()

driver.find_element_by_css_selector('''select[messages='{required:"请选择销售商！"}']''')

driver.find_element_by_css_selector('''select[messages='{required:"请选择业务名称！"}']''')

driver.find_element_by_css_selector('''input[messages='{floatIntervalCheck:"持有天数区间输入不规范"}']''')
