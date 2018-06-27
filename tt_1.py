from selenium import webdriver
option = webdriver.ChromeOptions()
option.add_argument('disable-infobars')
driver = webdriver.Chrome(chrome_options=option)
driver.get('http://10.2.130.78:8080/bomp/login.html')
driver.maximize_window()

aa.find_eles(By.CSS_SELECTOR,'#usernameInput0')[0].send_keys('10816')
aa.find_eles(By.CSS_SELECTOR,'#passwordInput0')[0].send_keys('123456789')
aa.find_eles(By.CSS_SELECTOR,'.login_btn')[0].click()
# =========================
driver.switch_to.default_content()
driver.switch_to.frame('frame-tab-132')
driver.switch_to.frame('26-frame')
driver.switch_to.frame('sysinfo_fundInfoBase-frame')


dts=driver.find_elements(By.CSS_SELECTOR,'dt')
dds=driver.find_elements(By.CSS_SELECTOR,'dd')
dict(zip([x.text for x in dts],[x.text for x in dds]))
driver.find_elements(By.CSS_SELECTOR,'label')

dts[1].location_once_scrolled_into_view

d['*基金名称'].find_element(By.CSS_SELECTOR,'input').get_attribute('value')

aa.top_window()
d['*基金名称'].find_element(By.CSS_SELECTOR,'input').send_keys()


js='document.getElementById("frame-tab-132").contentWindow.document.getElementById("26-frame").\
contentWindow.document.getElementById("sysinfo_fundInfoBase-frame").contentWindow.document.getElementById("infoForm")'




if driver.execute_script('return arguments[0].querySelector("select")',dds[4]):
    jse.executeScript("arguments[0].setAttribute('style', arguments[1])", div, "height: 1000px")


    print('is a select ele')
else:
    print('is not a select ele')











