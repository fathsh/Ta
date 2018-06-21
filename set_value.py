import win32com.client, win32api, win32con
from datetime import datetime
def ExcelRead(sheet, row, cell):
    t=LazyExcel.ExcelRead(sheet, row, cell)[0]
    if isinstance(t,datetime):
        t=str(t).split(' ')[0]
    t=str(t)
    if t=='None':
        return ''
    return t[:-2] if t[-2:]=='.0' else t


LazyExcel = win32com.client.Dispatch('Lazy.LxjExcel')
LazyExcel.ExcelOpen(r'e:\ta\云TA基金信息模板 v2.5（金湖&百川&混沌）.xlsx', 1)
a=ExcelRead(1,2,2)
print(a)
