from Ta_task import *
from SubTask_insert_code import *

def click(x, y=None, method=0, waittime=0.02):
    if isinstance(x, tuple):
        x, y = x[0], x[1]
    dm.MoveTo(x, y)
    time.sleep(0.02)
    if method == 0:
        dm.LeftClick()
    elif method == 1:
        dm.LeftDoubleClick()
    elif method == 2:
        dm.RightClick()
    time.sleep(waittime)

def f():
    click(119,147,waittime=1)
    click(285,183)
    uia.SendKeys('500114{Enter}',waitTime=1)
    click(593,448,waittime=2)
    click(575,181,waittime=1)
    click(899,506,waittime=2)





if __name__=="__main__":
    aa=SubTask_insert_code()

    # aa.get_excel_data()
    # exit()
    aa.login_ta()
    # aa.del_code('500001')
    # exit()
    co=0
    while True:
        co+=1
        print('='*100)
        print(co)
        # aa.into_frame_tab_132()
        aa.get_new_code()
        aa.add_code()
        aa.copy_code()
        aa.compare_after_copy_code()
        aa.set_value_after_compare()
        aa.after_add_code0()
        aa.after_add_code1()
#
        # break
        aa.del_code('500114')
        # f()
#
#     # aa.after_add_code1()
    msgbox(aa.log)
    aa.driver.quit()