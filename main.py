from Ta_task import *
from SubTask_insert_code import *


def error(ta,e):
    if e.find('Message: Cannot locate option with value:') >= 0:
        msgbox('单选下拉框数据非法，{}|code={}'.format(e,ta.new_code))
    if e.find('multiple invalid data') >= 0:
        msgbox('多选下拉框数据非法，{}|code={}'.format(e,ta.new_code))
    if e.find('invalid data') >= 0:
        msgbox('数据非法，{}|code={}'.format(e,ta.new_code))
    # input('press a key')

# def main():
#     ta = SubTask_insert_code()
#     ta.run()

if __name__=="__main__":
    try:
        ta = SubTask_insert_code()
        ta.run()
        # print(ta.log)
    except TaError as e:
        error(ta,str(e))



