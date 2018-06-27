import win32com.client, win32api, win32con, win32gui
from datetime import datetime
from sys import _getframe
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

def aaa1():
    print(type(_getframe().f_code.co_name))

# a=['序号','基金代码','基金名称','基金英文名称','单位净值长度','分红方式','基金状态','管理人名称','托管行名称','所属部门','清算批次号','投资方向','风险等级','发行价格(元)','最低募集金额','最高募集金额','募集起始日期','募集截止日期','验资日期','成立日期','封闭结束日期','巨额赎回比例(%)','最低资产限额','最低账户数量','最大账户数量','认购利息处理方式','基金申请确认天数','强制赎回方式','级差控制方式','净值公布频率按月计算方式','净值公布月','净值公布频率','净值公布日','巨额赎回顺延方式','最低账面处理方式','赎回后资产低于最低账面金额处理方式','赎回最少持有天数','认申购费用计算方式','首次投资是否含费','追加投资是否包含费用','利息计算是否含费','赎回费计算模式','合同类型','是否OTC产品','OTC产品编码	','OTC产品类别','销售商','母代码']
# b=['序号','基金代码','基金名称','基金英文名称','单位净值长度','分红方式','基金状态','管理人名称','托管行名称','所属部门','清算批次号','投资方向','风险等级','发行价格(元)','最低募集金额','最高募集金额','募集起始日期','募集截止日期','验资日期','成立日期','封闭结束日期','巨额赎回比例(%)','最低资产限额','最低账户数量','最大账户数量','认购利息处理方式','基金申请确认天数','强制赎回方式','级差控制方式','净值公布频率按月计算方式','净值公布月','净值公布频率','净值公布日','巨额赎回顺延方式','最低账面处理方式','赎回后资产低于最低账面金额处理方式','赎回最少持有天数','认申购费用计算方式','首次投资是否含费','追加投资是否包含费用','利息计算是否含费','赎回费计算模式','合同类型','是否OTC产品','OTC产品编码','OTC产品类别','销售商','母代码']
a='ghg1233'

print(a.startswith('ghg'))

# format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")
# aaa1()





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


