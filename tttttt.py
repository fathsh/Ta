# import win32com.client, win32api, win32con, win32gui
# parent=win32gui.FindWindow(0,'文件上传')
# hwndChildList=[]
# win32gui.EnumChildWindows(parent, lambda hwnd, param: param.append(hwnd),  hwndChildList)
# btn=[x for x in hwndChildList if win32gui.GetWindowText(x)=='打开(&O)'][0]
# edit=sorted([(win32gui.GetWindowRect(x)[1],x) for x in hwndChildList if win32gui.GetClassName(x)=='Edit'])[1][1]
# win32api.SendMessage(edit, 0x000C, 0, '文件路径')
# win32api.PostMessage(btn, win32con.WM_LBUTTONDOWN, 0, 0)
# win32api.PostMessage(btn, win32con.WM_LBUTTONUP, 0, 0)


# print(btn)
# print(edit)
# for x in hwndChildList:
#     # print(win32gui.GetClassName(x))
#     print(win32gui.GetWindowText(x))
def f(a):
    c(lambda :a)
def c(f1):
    print(f1())

print(type(f))
exit()



def compaer(data_excel, data_sys):
    if data_excel != '':
        data_excel = data_excel.replace('：', ':').split(':')[0]
    try:
        return True if float(data_excel) - float(data_sys.replace(',', '')) == 0 else False
    except:
        return True if data_excel == data_sys else False


#
# [key for key in data_excel if not compaer(data_excel[key],data_sys.get(key)) and key in data_sys]
# [print(key,data_excel[key],data_sys[key]) for key in [key for key in data_excel if not compaer(data_excel[key],data_sys.get(key)) and key in data_sys]]


driver.switch_to.frame('frame-tab-sysinfo_fundInfo-add-fund')
driver.switch_to.frame('layui-layer-iframe3')
dds=driver.find_elements_by_css_selector('dd')
s=dds[2].find_element_by_css_selector('select').send_keys('ddd')
driver.execute_script('arguments[0].setAttribute("style", arguments[1])', s, "display:block")


driver.switch_to.frame('frame-tab-sysinfo_fundInfo-add-fund')
driver.switch_to.frame('layui-layer-iframe1')
dd=driver.find_elements_by_css_selector('dd')
inp=dd[2].find_element_by_css_selector('input')

dm.SetWindowState(hwnd,1)
dm.SetWindowState(hwnd,4)
driver.execute_script('arguments[0].focus()', inp)
uia.Win32API.SendKeys('011')



print('\033[1;33m{}'.format('ddsasa') )


