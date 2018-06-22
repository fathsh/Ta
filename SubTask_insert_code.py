from Ta_task import *
class SubTask_insert_code(TaTask):
    '''SubTask_insert_code用于在Ta系统新增代码'''

    def data_pre_treated(self,data):
        self.log_write('data_pre_treated')
        for x in data:
            if self.otc:
                pass
            else:
                x['基金信息']['管理费计提比例(%)']=str(float(x['基金信息']['管理费计提比例(%)'])*100)
                x['基金信息']['巨额赎回比例(%)']=str(float(x['基金信息']['巨额赎回比例(%)'])*100)
                x['基金信息']['募集起始日期']=xlrd.xldate.xldate_as_datetime(float(x['基金信息']['募集起始日期']), 0).strftime("%Y-%m-%d")
                x['归基金资产比例']['持有天数区间']='' if (x['归基金资产比例']['最低持有天数']=='0' and x['归基金资产比例']['最高持有天数'])\
                            =='999999999' else '{},{}'.format(x['归基金资产比例']['最低持有天数'],x['归基金资产比例']['最高持有天数'])

        return data

    def get_excel_data_raw(self):
        def is_valid(data):
            '''判断是否参数数据'''
            for x in data[1:min(10, len(data))]:
                if x:
                    return True
            return False
# ================================================================
        self.log_write('get_excel_data_raw')
        # workbook = xlrd.open_workbook(r'e:\ta\云TA基金信息模板 v2.5（金湖&百川&混沌）.xlsx')
        workbook = xlrd.open_workbook(r'e:\ta\【2018-6-14 20 605 GF0401-GF0405】云TA基金信息模板 - OTC - v2.3.xlsx')
        excel_data_raw = {}
        for sheet in workbook.sheets():
            sheet_data = []
            for row in range(1, sheet.nrows):
                row_data = sheet.row_values(row)
                if is_valid(row_data):
                    sheet_data.append(row_data)
                else:
                    excel_data_raw[sheet.name] = sheet_data
                    break
        self.otc=False if excel_data_raw['基金信息'][1][1]=='' else True
        self.excel_data_raw = excel_data_raw
        print(self.otc)
        # exit()

    def clean_datas(self):
        self.log_write('clean_datas')
        self.read_config()
        print(self.data_show)
        # exit()
        file_datas={}
        datas={key:value for key,value in self.excel_data_raw.items() if key in self.sheet_show}
        for sheet_name,values in datas.items():
            if sheet_name=='清算天数设置':
                file_datas['清算天数设置_托管行清算天数设置'], file_datas['清算天数设置_销售商清算天数设置'] = [], []
            else:
                file_datas[sheet_name]=[]
            for i,value in enumerate(values):
                data=[self.key_map[x] if x in self.key_map else x for x in value]
                data=[str(x)[:-2] if str(x)[-2:]=='.0' else str(x) for x in data]
                if sheet_name=='清算天数设置':
                    if i==0:
                        header1,header2=data[:9],data[:3]+data[9:-4]
                    else:
                        data1,data2=data[:9],data[:3]+data[9:-4]
                        data_tmp1,data_tmp2=dict(zip(header1,data1)),dict(zip(header2, data2))
                        data_tmp1={key:value for key,value in data_tmp1.items() if (key in self.data_show[sheet_name]
                                                                                  if sheet_name in self.data_show else True)}
                        data_tmp2={key:value for key,value in data_tmp2.items() if (key in self.data_show[sheet_name]
                                                                                  if sheet_name in self.data_show else True)}
                        file_datas['清算天数设置_托管行清算天数设置'].append(data_tmp1)
                        file_datas['清算天数设置_销售商清算天数设置'].append(data_tmp2)
                else:
                    if i==0:
                        header=data
                    else:
                        data_tmp=dict(zip(header,data))
                        data_tmp={key:value for key,value in data_tmp.items() if (key in self.data_show[sheet_name]
                                                                                  if sheet_name in self.data_show else True)}
                        file_datas[sheet_name].append(data_tmp)
        return file_datas

    def stack_datas(self,datas):
        self.log_write('stack_datas')
        codes=[x['基金代码' if x['基金代码']!='' else '基金名称'] for x in datas['基金信息']]
        file_datas=[]
        print(codes)
        for code in codes:
            data_tmp={}
            for key,values in datas.items():
                data_tmp[key]=[x for x in values if x.get('基金代码')==code or x.get('基金名称')==code] if key=='产品个户交易限制信息' else \
                    [x for x in values if x.get('基金代码')==code or x.get('基金名称')==code][0]
            file_datas.append(data_tmp)
        return file_datas

    def login_ta(self):
        '''登录ta系统'''
        self.log_write('login_ta')
        if not self.driver:
            option = webdriver.ChromeOptions()
            option.add_argument('disable-infobars')
            self.driver = webdriver.Chrome(chrome_options=option)
        self.driver.get(self.url)
        self.driver.maximize_window()
        self.super_find_eles('#usernameInput0').send_keys('10816')
        self.super_find_eles('#passwordInput0').send_keys('123456789')
        self.super_find_eles('form').submit()

    def get_new_code(self,data):
        '''获取新代码函数'''
        code=data['基金代码']
        self.name=data['基金名称']
        self.log_write('{}\n{} get_new_code'.format('='*100,datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        self.driver.switch_to.default_content()
        self.super_click(self.super_find_eles('a[data-text="信息维护"]',remark='信息维护'))
        self.super_click(self.super_find_eles('ul.up-sub-list>li',frames='frame-tab-24',remark='基金信息设置'),waittime=1)
        if code=='':
            if self.new_code is None:
                lis_gupiao_gm=self.super_find_eles('#id_0_0>li',frames='frame-tab-132',return_all=True)
                lis_gupiao_sm=self.super_find_eles('#id_1_0>li',return_all=True)                        #获取私募-股票下的所有li元素
                lis_qiquan_sm=self.super_find_eles('#id_1_7>li',return_all=True)                               #获取私募-期权下的所有li元素
                lis=lis_gupiao_gm+lis_gupiao_sm+lis_qiquan_sm
                self.code_be_copy=self.skip_Exception(lambda :sorted([x.get_attribute('innerText') for x in lis])[-1].split('：')[0])     #获取最后一个代码
                self.new_code = '{}{}'.format(self.code_be_copy[:3],
                                              int(self.code_be_copy[-3:]) + (2 if self.code_be_copy[-1] == '3' else 1))
                print(self.code_be_copy)
                print(self.new_code)
            else:
                self.code_be_copy=self.new_code
                self.new_code='{}{}'.format(self.code_be_copy[:3],int(self.code_be_copy[-3:])+(2 if self.code_be_copy[-1]=='3' else 1))
            self.log_write('【new_code={},name={},code_be_copy={}】'.format(self.new_code, self.name,self.code_be_copy))
        else:
            pass

    def add_code(self):
        '''新增代码函数'''
        self.log_write('add_code')
        self.skip_Exception(lambda :self.super_find_eles('#new-fund', frames=['frame-tab-132', -1]).click())
        datas={'基金模板':'210000201:一对多专户净值型产品子模板1',
               '基金代码':self.new_code,
               '基金名称':self.name,        #'TA名称':'87:广发证券股份有限公司',
                '管理人名称':'gf0002:广发证券柜台交易市场部' if self.otc else 'gf0001:广发托管外包'}
        # datas['基金代码']='870022'
        eles = self.get_eles_to_set(datas.keys(),frames=['frame-tab-132',-1,-1])
        self.set_values(eles,datas,'add code')
        self.super_find_eles('#dialog-btn-save').click()

    def copy_code(self):
        '''复制代码'''
        self.log_write('copy_code')
        ele=self.super_find_eles('#fundcode-copy',frames='frame-tab-sysinfo_fundInfo-add-fund')
        # self.code_be_copy='870022'
        self.skip_Exception(lambda :self.set_value(ele,self.code_be_copy,sel='#fundcode-copy'),remark='copy_code')
        self.super_click(self.super_find_eles('#copy-fundinfo'))

    def set_value_fundInfoBase(self,data_fundInfoBase):
        '''比对之后设置0'''
        self.log_write('set_value_fundInfoBase')
        for i in range(2):
            result_compare=self.compare_values(data_fundInfoBase,
                                        frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_fundInfoBase-frame'],length=59)
            # print(result_compare)
            self.log_write(str(result_compare) if result_compare else '【all data compare correctly!】')
            eles=self.get_eles_to_set(result_compare.keys())
            data={key:value[0] for key,value in result_compare.items()}
            self.set_values(eles,data,'set_value_fundInfoBase' if i==0 else 'set_value_fundInfoBase_repeat',mode=1)
            result_compare=self.compare_values(data_fundInfoBase,
                                            frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_fundInfoBase-frame'],length=59)
            if not result_compare:
                return
        raise

    def set_value_arLimitList(self,data_arLimitList):
        '''比对之后设置1'''
        self.log_write('set_value_arLimitList')
        for excel_data in data_arLimitList:
            client_type=excel_data['客户类型'].split(':')[-1]
            if client_type=='产品':
                continue
            self.super_find_eles('div.fund-tab.clear > ul > li:nth-child(2) > a',
                                 frames='frame-tab-sysinfo_fundInfo-add-fund').click()
            result_compare,tds,data_sys=self.compare_values(excel_data,frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_arLimitList-frame'],
                                               data_mode='table',selcet_data_by_value=client_type)
            # result_compare={key:value for key,value in result_compare.items() if key not in ['客户类型','销售商'] or value[0].replace('：',':').split(':')[-1]!=value[1]}
            # self.log_write(str(result_compare) if result_compare else '【all data compare correctly!】')
            if data_sys is None:
                print('data_sys is None')
                self.skip_Exception(lambda: self.super_find_eles('#addButton').click())
            elif result_compare['销售商'][1] == result_compare['销售商'][0].replace('：',':').split(':')[-1]:
                print('销售商 correctly!')
                result_compare={key:value for key,value in result_compare.items() if key not in ['客户类型', '销售商','key_not_find']}
                self.log_write('销售商 correctly!'+client_type+':'+(str(result_compare) if result_compare
                                                                 else '【all data compare correctly!】'))
                if not result_compare:
                    continue
                self.super_find_eles('input', ele_parent=tds[0]).click()
                self.super_find_eles('button[name="batchEdit"]').click()
            elif result_compare['销售商'][1] != result_compare['销售商'][0].replace('：',':').split(':')[-1]:
                print('销售商 wrong!')
                data_sys=None
                self.super_find_eles('input',ele_parent=tds[0]).click()
                self.super_click(self.super_find_eles('button[name="batchDelete"]'),waittime=1)
                self.super_click(self.super_find_eles('#batchdelete-sure'),mode=0)
                self.skip_Exception(lambda :self.super_find_eles('#addButton').click())
            eles = self.get_eles_to_set(result_compare.keys() if result_compare else excel_data.keys(),
                                        frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_arLimitList-frame',-1])
            data = {key: value[0] for key, value in result_compare.items() if key != '基金代码'} if data_sys else excel_data
            self.set_values(eles, data, 'set_value_arLimitList_{}'.format(excel_data['客户类型'].split(':')[-1]))
            self.super_find_eles('#dialog-btn-save').click()
            self.super_find_eles('td',frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_arLimitList-frame'],waittime=2)

    def set_value_fundParameterEdit(self,data_fundParameterEdit):
        self.log_write('set_value_fundParameterEdit')
        self.super_find_eles('div.fund-tab.clear > ul > li:nth-child(3) > a',
                             frames='frame-tab-sysinfo_fundInfo-add-fund').click()
        for i in range(2):
            result_compare = self.compare_values(data_fundParameterEdit,
                                    frames=['frame-tab-sysinfo_fundInfo-add-fund','sysInfo_fundParameterEdit-frame'], length=85)
            # key_not_find=result_compare['key_not_find']
            # print([x for x in key_not_find if x in self.tmp])
            result_compare={key:value for key,value in result_compare.items() if key not in ['基金代码', 'key_not_find']}
            # print(result_compare)
            self.log_write(str(result_compare) if result_compare else '【all data compare correctly!】')
            eles = self.get_eles_to_set(result_compare.keys())
            # data = {key: value[0] for key, value in result_compare.items() if key != '基金代码'}
            data = {key: value[0] for key, value in result_compare.items()}
            self.set_values(eles,data,'set_value_fundParameterEdit' if i==0 else 'set_value_fundParameterEdit_repeat')
            result_compare = self.compare_values(data_fundParameterEdit)
            result_compare = {key:value for key, value in result_compare.items() if key not in ['基金代码', 'key_not_find']}
            # print(result_compare)
            if not result_compare:
                self.super_find_eles('div.fund-tab.clear > ul > li:nth-child(5) > a',
                                     frames='frame-tab-sysinfo_fundInfo-add-fund').click()
                return
        raise

    def set_value_fundBelongAssetList(self,data_fundBelongAssetList):
        self.log_write('set_value_fundBelongAssetList')
        self.super_find_eles('div.fund-tab.clear > ul > li:nth-child(5) > a',
                             frames='frame-tab-sysinfo_fundInfo-add-fund').click()
        # print(data_fundBelongAssetList)
        self.super_find_eles('button[name="trading-new"]',
                             frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_fundBelongAssetList-frame']).click()
        eles={'销售商': self.super_find_eles('''select[messages='{required:"请选择销售商！"}']''',frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_fundBelongAssetList-frame',-1]),
              '业务名称': self.super_find_eles('''select[messages='{required:"请选择业务名称！"}']'''),
              '持有天数区间':self.super_find_eles('''input[messages='{floatIntervalCheck:"持有天数区间输入不规范"}']''')}
        self.set_values(eles,data_fundBelongAssetList,'set_value_fundBelongAssetList')
        eles['持有天数区间'].send_keys(Keys.ENTER)
        self.super_find_eles('#save').click()

    def set_value_fundSetupInfoList(self,data_fundSetupInfoList):
        '''比对之后设置'''
        self.log_write('set_value_fundSetupInfoList')
        self.super_find_eles('div.fund-tab.clear > ul > li:nth-child(6) > a',
                             frames='frame-tab-sysinfo_fundInfo-add-fund').click()
        excel_data = data_fundSetupInfoList
        # print(excel_data)
        result_compare, tds, data_sys = self.compare_values(excel_data, frames=['frame-tab-sysinfo_fundInfo-add-fund',
                                            'sysinfo_fundSetupInfoList-frame'],data_mode='table')
        self.log_write(str(result_compare) if result_compare else '【all data compare correctly!】')
        # print(result_compare)
        if data_sys is None:
            print('data_sys is None')
            self.skip_Exception(lambda: self.super_find_eles('#addButton').click())
        elif result_compare['销售商'][1] == result_compare['销售商'][0].replace('：', ':').split(':')[-1]\
                and '利息类别' not in result_compare:
            print('销售商 and 利息类别 correctly!')
            result_compare = {key: value for key, value in result_compare.items() if key not in ['销售商','利息类别']}
            if not result_compare:
                return
            self.super_find_eles('input', ele_parent=tds[0]).click()
            self.super_find_eles('button[name="batchEdit"]').click()
        elif result_compare['销售商'][1] != result_compare['销售商'][0].replace('：', ':').split(':')[-1]\
                or '利息类别' in result_compare:
            print('销售商 and 利息类别 wrong!')
            excel_data['销售商']='GF8:私募直销'
            excel_data['计息结束日期']='2099-12-31'
            data_sys = None
            self.super_find_eles('input', ele_parent=tds[0]).click()
            self.super_click(self.super_find_eles('button[name="batchDelete"]'), waittime=1)
            self.super_click(self.super_find_eles('#batchdelete-sure'), mode=0)
            self.skip_Exception(lambda: self.super_find_eles('#addButton').click())
        eles = self.get_eles_to_set(result_compare.keys() if data_sys else excel_data.keys(),
                                    frames=['frame-tab-sysinfo_fundInfo-add-fund', 'sysinfo_fundSetupInfoList-frame', -1])

        data = {key: value[0] for key, value in result_compare.items() if
                key != '基金代码'} if data_sys else excel_data
        self.set_values(eles, data, 'set_value_fundSetupInfoList')
        self.super_find_eles('#dialog-btn-save').click()
        self.super_find_eles('td', frames=['frame-tab-sysinfo_fundInfo-add-fund', 'sysinfo_fundSetupInfoList-frame'],
                             waittime=2)

    def set_value_fundInfo_add_fund(self,seller):
        '''基金信息新增'''
        self.log_write('set_value_fundInfo_add_fund')
        self.super_find_eles('#finish', frames='frame-tab-sysinfo_fundInfo-add-fund').click()
        data=[{'销售商':seller,'代销标志':'1:代销'},
              {'销售商':seller},]
        for i in range(2):
            self.super_click(self.super_find_eles('div.fund-result-list>ul>li:nth-child({}) a'.format(i+1),
                                                  frames='frame-tab-sysinfo_fundInfo-add-fund'), waittime=1)
            eles = self.get_eles_to_set(data[i].keys(), frames=['frame-tab-sysinfo_fundInfo-add-fund', -1])
            self.set_values(eles, data[i], 'set_value_fundInfo_add_fund{}'.format(i))
            self.super_find_eles('#finish', frames=['frame-tab-sysinfo_fundInfo-add-fund', -1]).click()\
                if i==1 else self.super_find_eles('#dialog-btn-save').click()
            f = lambda: self.super_find_eles('div.fund-result-list>ul >li:nth-child({})'.format(i+1),
                                             frames='frame-tab-sysinfo_fundInfo-add-fund').get_attribute('innerText')
            self.skip_Exception(f, lambda: f().find(' 完成') > -0)

    def set_value_frame_tab_119_161(self,data):
        self.log_write('set_value_frame_tab_119_161')
        self.driver.refresh()
        for i in range(2):
            key_name='清算天数设置_托管行清算天数设置' if i==0 else '清算天数设置_销售商清算天数设置'
            self.driver.switch_to.default_content()
            self.super_click(self.super_find_eles('a[data-text="信息维护"]',remark='信息维护'))
            ele=self.super_find_eles('ul.menu-tab-list>li:nth-child(4)',frames='frame-tab-24',remark='清算天数设置')
            self.super_click(ele,waittime=1)
            li=self.super_find_eles('li',ele_parent=ele,return_all=True)[i]
            excel_data=data[key_name]
            excel_data['基金名称']=self.new_code
            keys=list(excel_data.keys())
            keys.insert(0,keys[-1])
            excel_data={key:excel_data[key] for key in keys[:-1]}
            self.super_find_eles('a',ele_parent=li).click()
            self.super_click(self.super_find_eles('button[name="trading-new"]',frames='frame-tab-119' if i==0 else 'frame-tab-161',
                                                  waittime=1),mode=0)
            eles=self.get_eles_to_set(frames=['frame-tab-119' if i==0 else 'frame-tab-161',-1])
            self.log_write('{}:{}'.format(key_name,excel_data))
            self.set_values(eles,excel_data)
            self.super_find_eles('#dialog-btn-save').click()
            if i==1:
                return

    def run(self):
        self.get_excel_data_raw()
        datas=self.clean_datas()
        datas = self.stack_datas(datas)
        datas = self.data_pre_treated(datas)
        # exit()
        print(datas[0]['基金信息'])
        # exit()
        self.login_ta()
        for data in datas:
            # print(data)
            data_fundInfoBase_fundParameterEdit = data['基金信息']
            [data_fundInfoBase_fundParameterEdit.update({key: value})
             for key, value in data['集中备份信息-第一次填写'].items() if key not in data['基金信息']]
            t = time.time()
            self.get_new_code(data_fundInfoBase_fundParameterEdit)
            self.add_code()
            self.copy_code()
            self.set_value_fundInfoBase(data_fundInfoBase_fundParameterEdit)
            self.set_value_arLimitList(data['产品个户交易限制信息'])
            self.set_value_fundParameterEdit(data_fundInfoBase_fundParameterEdit)
            self.set_value_fundBelongAssetList(data['归基金资产比例'])
            self.set_value_fundSetupInfoList(data['基金成立信息'])
            self.set_value_fundInfo_add_fund(data_fundInfoBase_fundParameterEdit['销售商'])
            self.set_value_frame_tab_119_161(data)
            print('\033[1;33m{}_total time: {}\033[0m'.format(self.new_code, round(time.time() - t, 2)))
            self.log_write('{} total time: {}'.format(self.new_code, round(time.time() - t, 2)))
            print(self.log)
            break
            self.driver.refresh()
        self.driver.quit()





