from Ta_task import *
from SubTask_insert_code import *
import pickle

def saveDbase_pickle(filename, object):
    f=open(filename,'wb')
    pickle.dump(object,f)
    f.close()

def loadDbase_pickle(filename):
    f = open(filename, 'rb')
    obj=pickle.load(f)
    f.close()
    return obj



if __name__=="__main__":
    ta=SubTask_insert_code()
    # ta.get_excel_data()
    # saveDbase_pickle('data_excel,',ta.excel_datas)
    ta.excel_datas=loadDbase_pickle('data_excel,')
    a=ta.excel_datas
    # print(a['基金信息'][0]+a['集中备份信息-第一次填写'][0])
    [ta.excel_datas['基金信息'][0].update({k:v}) for k,v in a['集中备份信息-第一次填写'][0].items() if k not in a['基金信息'][0]]
    print(ta.excel_datas['产品个户交易限制信息'])

    # exit()


    # print(a['集中备份信息-第一次填写'][0].keys())
    # print([x for x in a['集中备份信息-第一次填写'][0] if x in a['基金信息'][0]])


    # exit()


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
        ta.set_value_after_compare0()
        ta.set_value_after_compare1()
        # exit()
        ta.set_value_after_compare2()
        ta.set_value_after_compare3()
        # ta.set_value_after_add_code()
        print('\033[1;33m{} {}\033[0m'.format('total time:', time.time() - t))
        exit()
        ta.driver.refresh()
#
        # break
        # ta.del_code(ta.new_code)

    msgbox(ta.log)
    ta.driver.quit()