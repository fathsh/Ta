from Ta_task import *
from SubTask_insert_code import *


if __name__=="__main__":
    ta=SubTask_insert_code()
    ta.get_excel_data()
    ta.login_ta()
    co=0
    while True:
        t = time.time()
        co+=1
        print('='*100)
        print(co)
        ta.get_new_code()
        ta.add_code()
        ta.copy_code()
        ta.compare_after_copy_code()
        ta.set_value_after_compare()
        ta.set_value_after_add_code()
        print('\033[1;33m{} {}\033[0m'.format('total time:', time.time() - t))
        exit()
        ta.driver.refresh()
#
        # break
        # ta.del_code(ta.new_code)

    msgbox(ta.log)
    ta.driver.quit()