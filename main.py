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

def error(e):
    if e.find('Message: Cannot locate option with value:') >= 0:
        print('单选下拉框数据非法，{}'.format(e))
    if e.find('multiple invalid data') >= 0:
        print('多选下拉框数据非法，{}'.format(e))
    if e.find('invalid data') >= 0:
        print('数据非法，{}'.format(e))
    input('press a key')

def main():
    ta = SubTask_insert_code()
    datas=ta.get_excel_data()
    # print(datas)
    ta.login_ta()
    co = 0
    # while True:
    for data in datas:
        print(data.keys())
        data_fundInfoBase_fundParameterEdit=data['基金信息']
        [data_fundInfoBase_fundParameterEdit.update({key: value})
                                             for key, value in data['集中备份信息-第一次填写'].items() if key not in data['基金信息']]
        print(data_fundInfoBase_fundParameterEdit)
        # t = time.time()
        # ta.get_new_code()
        # ta.add_code()
        # ta.copy_code()
        # ta.set_value_fundInfoBase(data_fundInfoBase_fundParameterEdit)
        # ta.set_value_arLimitList(data[ '产品个户交易限制信息'])
        # ta.set_value_fundParameterEdit(data_fundInfoBase_fundParameterEdit)
        # ta.set_value_fundBelongAssetList(data['归基金资产比例'])
        # ta.set_value_fundSetupInfoList(data['基金成立信息'])
        # ta.set_value_after_add_code()
        ta.f(data)
        # print('\033[1;33m{}_total time: {}\033[0m'.format(ta.new_code, round(time.time() - t,2)))
        # exit()
        # continue
        input(...)
        exit()
        ta.driver.refresh()
        #
        # break
        # ta.del_code(ta.new_code)

    msgbox(ta.log)
    ta.driver.quit()

if __name__=="__main__":
    try:
        main()
    except TaError as e:
        error(str(e))



