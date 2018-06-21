import win32com.client, win32api, win32con, win32gui
from datetime import datetime
import pickle
def getExcelSheetName(sheet):
    return LazyExcel.SheetGetName(sheet)[0]

def getExcelrows(sheet):
    return LazyExcel.SheetRowsCount(sheet)[0]

def getExceltColumns(sheet):
    return LazyExcel.SheetColumnsCount(sheet)[0]

def loadDbase_pickle(filename):
    f = open(filename, 'rb')
    obj=pickle.load(f)
    f.close()
    return obj
# format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")
a={'a':1,'b':2}
print(a)
a.pop('a')
print(a)





#
# LazyExcel = win32com.client.Dispatch('Lazy.LxjExcel')
# LazyExcel.ExcelOpen(r'e:\ta\云TA基金信息模板 v2.5（金湖&百川&混沌）.xlsx', 1)
# sheet_count = LazyExcel.SheetCount()
# for sheet in range(1, sheet_count + 1):
#     sheet_name = getExcelSheetName(sheet)
#     if sheet_name not in ['清算天数设置']:
#         continue
#     print(sheet_name)
#     print(getExceltColumns(sheet))
#
# def f(a):
#     c(lambda :a)
# def c(f1):
#     print(f1())


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


