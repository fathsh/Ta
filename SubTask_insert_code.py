from Ta_task import *



class SubTask_insert_code(TaTask):
    '''SubTask_insert_code用于在Ta系统新增代码'''
    def __init__(self):
        TaTask.__init__(self)
        self.data_show={'基金信息':['序号', '基金代码', '基金名称', '单位净值长度', '分红方式', '托管行名称', '清算批次号', '管理费计提比例(%)',
                        '投资方向', '份额类别', '最低募集金额', '最高募集金额', '募集起始日期', '巨额赎回比例(%)', '最低资产限额', '最低账户数量',
                        '最大账户数量', '认购利息处理方式', '基金申请确认天数', '强制赎回方式', '级差控制方式', '级差的处理方式', '存续期数(月)',
                        '户数限制处理方式', '净值公布频率按月计算方式', '净值公布月', '净值公布频率', '净值公布日', '巨额赎回顺延方式',
                        '赎回后资产低于最低账面金额处理方式', '赎回最少持有天数', '利息计算是否含费', '按收益率设置费率模式',
                        '是否只针对有限合伙人计提', '股权收益率类型', '股权收益率(%)', '股权计提比例', '安全垫产品持有天数', '销售商']}
        self.excel_datas={}
        self.onchange_key = ['基金代码','销售商最大折扣','基金管理人','基金申请确认天数']

    def get_excel_data(self):
        '''读取excel数据'''
        LazyExcel.ExcelOpen(r'e:\ta\云TA基金信息模板 v2.5（金湖&百川&混沌）.xlsx', 1)
        sheet_count = int(LazyExcel.SheetCount())
        for sheet in range(1, sheet_count + 1):
            sheet_name = getExcelSheetName(sheet)
            columns_count = getExceltColumns(sheet)
            datas = []
            for i in range(2, 999):
                data_tmp = get_row_data(sheet, i, columns_count)
                if data_is_valid(data_tmp):
                    datas.append(get_row_data(sheet, i, columns_count))
                else:
                    break
            self.excel_datas[sheet_name] = filter([dict(zip(datas[0], x)) for x in datas[1:]],sheet_name,self.data_show)  # 生成字典，去隐藏
        print(self.excel_datas)

    def login_ta(self):
        '''登录ta系统'''
        if not self.driver:
            option = webdriver.ChromeOptions()
            option.add_argument('disable-infobars')
            self.driver = webdriver.Chrome(chrome_options=option)
        self.driver.get('http://10.2.130.78:8080/bomp/login.html')
        self.driver.maximize_window()
        self.super_find_eles('#usernameInput0').send_keys('10816')
        self.super_find_eles('#passwordInput0').send_keys('123456789')
        self.super_find_eles('form').submit()

    def get_new_code(self):
        '''获取新代码函数'''
        self.driver.switch_to.default_content()
        self.super_click(self.super_find_eles('a[data-text="信息维护"]',log='login_ta',remark='信息维护'))
        self.super_click(self.super_find_eles('ul.up-sub-list>li',frames='frame-tab-24',remark='基金信息设置'),waittime=0.8)
        lis_gupiao_gm=self.super_find_eles('#id_0_0>li',frames='frame-tab-132',return_all=True)
        lis_gupiao_sm=self.super_find_eles('#id_1_0>li',return_all=True)       #获取私募-股票下的所有li元素
        lis_qiquan_sm=self.super_find_eles('#id_1_7>li',return_all=True)                               #获取私募-期权下的所有li元素
        lis=lis_gupiao_gm+lis_gupiao_sm+lis_qiquan_sm
        t=self.skip_Exception(lambda :sorted([x.get_attribute('innerText') for x in lis])[-1].split('：')[0])     #获取最后一个代码
        self.new_code='{}{}'.format(t[:3],int(t[-3:])+1)

    def add_code(self):
        '''新增代码函数'''
        # self.skip_Exception(lambda :self.super_find_eles('#new-fund', frames=['frame-tab-132', -1], waittime=1).click())
        self.skip_Exception(lambda :self.super_find_eles('#new-fund', frames=['frame-tab-132', -1]).click())
        datas={'基金模板':'210000201:一对多专户净值型产品子模板1','基金代码':self.new_code,
               '基金名称':'test1','TA名称':'87:广发证券股份有限公司','管理人名称':'gf0002:广发证券柜台交易市场部'}
        # datas['基金代码']='870022'
        eles = self.get_eles_to_set(datas.keys(),frames=['frame-tab-132',-1,-1])
        self.set_values(eles,datas,'add code')
        self.check_invalid_data()
        self.form_submit(self.super_find_eles('#dialog-btn-save'))

    def copy_code(self):
        '''复制代码'''
        ele=self.super_find_eles('#fundcode-copy',frames='frame-tab-sysinfo_fundInfo-add-fund',log='add_code')
        self.code_be_copy='{}{}'.format(self.new_code[:3],int(self.new_code[-3:])-1)      #self.code_be_copy=self.new_code-1
        self.code_be_copy='870022'
        self.skip_Exception(lambda :self.set_value(ele,self.code_be_copy,sel='#fundcode-copy'),remark='copy_code')
        # print('\033[1;33m{} {}\033[0m'.format('set_value select',time.time()-t))
        self.super_click(self.super_find_eles('#copy-fundinfo'))

    def set_value_after_compare0(self):
        '''比对之后设置0'''
        # self.excel_datas['基金信息'][0]['募集起始日期']='2018-05-16'
        for i in range(2):
            result_compare=self.compare_values(self.excel_datas['基金信息'][0],
                                                        frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_fundInfoBase-frame'],length=59)
            print(result_compare)
            eles=self.get_eles_to_set(result_compare.keys())
            data={key:value[0] for key,value in result_compare.items() if key!='基金代码'}
            self.set_values(eles,data,'set_value_after_compare0' if i==0 else 'set_value_after_compare0_repeat',mode=1)
            result_compare=self.compare_values(self.excel_datas['基金信息'][0],
                                                        frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_fundInfoBase-frame'],length=59)
            self.check_invalid_data()
            if not result_compare or (len(result_compare)==1 and result_compare['基金代码']):
                return
        raise
        #     self.form_submit(self.super_find_eles('#finish',frames='frame-tab-sysinfo_fundInfo-add-fund'),
        #                      frames_label=['frame-tab-sysinfo_fundInfo-add-fund',-1])

    def set_value_after_compare1(self):
        '''比对之后设置1'''
        ti=time.time()
        for excel_data in self.excel_datas['产品个户交易限制信息']:
            show_data=['客户类型','销售商','首次投资最低金额','最少追加金额','级差金额','最低账面金额']
            excel_data={key:excel_data[key] for key in excel_data if key in show_data}
            if excel_data['客户类型'].split(':')[-1]=='产品':
                continue
            self.super_find_eles('div.fund-tab.clear > ul > li:nth-child(2) > a',
                                 frames='frame-tab-sysinfo_fundInfo-add-fund').click()
            table=self.super_find_eles('table.datagrid-htable',frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_arLimitList-frame'],return_all=True)[-1]
            tds=self.super_find_eles('td',ele_parent=table,return_all=True)
            header=[self.driver.execute_script("return arguments[0].innerText", x).replace('\xa0','').strip() for x in tds[2:]]
            print(header)
            table_trs=self.super_find_eles('table.datagrid-btable',return_all=True)[-1]
            trs=self.super_find_eles('tr',ele_parent=table_trs,return_all=True)[::-1]
            tdss=[self.super_find_eles('td',ele_parent=tr,return_all=True) for tr in trs]
            tmp=[x for x in tdss if x[2].text==excel_data[ '客户类型'].split(':')[-1]]
            tds=tmp[0] if tmp else None
            data_sys=dict(zip(header,[self.driver.execute_script("return arguments[0].innerText", x).replace('\xa0','').strip()
                                      for x in tds[2:]])) if tds else None
            excel_data['销售商']='040：xx'
            result_compare = {}
            if data_sys and data_sys['销售商'] in excel_data['销售商'].split(':')[-1]:
                print('销售商 correctly!')
                for k,v in data_sys.items():
                    if k in ['客户类型', '销售商'] or k not in excel_data:
                        continue
                    if self.compaer_value(excel_data[k],v):
                        print(k,excel_data[k],v,'correctly!')
                    else:
                        result_compare[k]=(excel_data[k],v)
                print(result_compare)
                if not result_compare:
                    continue
                self.super_find_eles('input', ele_parent=tds[0]).click()
                self.super_find_eles('button[name="batchEdit"]').click()
            elif data_sys and data_sys['销售商'] not in excel_data['销售商'].split(':')[-1]:
                print('销售商 wrong!')
                self.super_find_eles('input',ele_parent=tds[0]).click()
                self.super_click(self.super_find_eles('button[name="batchDelete"]'),waittime=1)
                self.super_click(self.super_find_eles('#batchdelete-sure'),mode=0)
                self.skip_Exception(lambda :self.super_find_eles('#addButton').click())
            elif data_sys is None:
                self.skip_Exception(lambda: self.super_find_eles('#addButton').click())
            eles = self.get_eles_to_set(result_compare.keys() if result_compare else excel_data.keys(),
                                        frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_arLimitList-frame',-1])
            data = {key: value[0] for key, value in result_compare.items() if key != '基金代码'} if result_compare else excel_data
            self.set_values(eles, data, 'set_value_after_compare1_{}'.format(excel_data['客户类型'].split(':')[-1]))
            self.super_find_eles('#dialog-btn-save').click()
            self.super_find_eles('td',frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_arLimitList-frame'],waittime=2)
            # msgbox(1)
            # exit()
            # self.super_find_eles('table.datagrid-htable',frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_arLimitList-frame'],return_all=True)[-1]
        print((time.time()-ti))

    def set_value_after_compare2(self):
        self.super_find_eles('div.fund-tab.clear > ul > li:nth-child(3) > a',
                             frames='frame-tab-sysinfo_fundInfo-add-fund').click()
        for i in range(2):
            result_compare = self.compare_values(self.excel_datas['基金信息'][0],
                                            frames=['frame-tab-sysinfo_fundInfo-add-fund','sysInfo_fundParameterEdit-frame'], length=85)
            print(result_compare)
            eles = self.get_eles_to_set(result_compare.keys())
            data = {key: value[0] for key, value in result_compare.items() if key != '基金代码'}
            self.set_values(eles,data,'set_value_after_compare2' if i==0 else 'set_value_after_compare2_repeat')
            result_compare = self.compare_values(self.excel_datas['基金信息'][0])
            if not result_compare or (len(result_compare) == 1 and '基金代码' in result_compare):
                # print('set_value_after_compare0 right')
                self.super_find_eles('div.fund-tab.clear > ul > li:nth-child(5) > a',
                                     frames='frame-tab-sysinfo_fundInfo-add-fund').click()
                return
        raise

    def set_value_after_compare3(self):
        self.super_find_eles('div.fund-tab.clear > ul > li:nth-child(5) > a',
                             frames='frame-tab-sysinfo_fundInfo-add-fund').click()
        data={'销售商': '040：全部销售商', '业务名称': '03:赎回', '持有天数区间':'9999999'}
        print(self.excel_datas['归基金资产比例'][0])
        self.super_find_eles('button[name="trading-new"]',
                             frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_fundBelongAssetList-frame']).click()
        eles={'销售商': self.super_find_eles('''select[messages='{required:"请选择销售商！"}']''',frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_fundBelongAssetList-frame',-1]),
              '业务名称': self.super_find_eles('''select[messages='{required:"请选择业务名称！"}']'''),
              '持有天数区间':self.super_find_eles('''input[messages='{floatIntervalCheck:"持有天数区间输入不规范"}']''')}
        self.set_values(eles,data)













    def set_value_after_add_code(self):
        '''基金信息新增'''
        data=[{'销售商':'011:xinye,040:lldj,376:陆金所,005:hhg','代销标志':'1:代销'},
              {'销售商': '011:xinye,040:lldj,376:陆金所,005:hhg'},
              {'转换入基金': 'SM0483:xinye', '转换出份额类别': '*:ssd,A:ttt', '销售商': 'GF9:yyrrg',
               '转换入份额类别': 'b:fff,c:rrr'},
              {'业务类型': '01:xinye,02:gg', '费用类型': '0:ssd,1:ttt', '销售商': 'GF9:yyrrg',
               '归销售商比例(%)': '30', '归注册机构比例(%)': '50', '费用分成模式': '2', '对方基金名称': 'tttrrdsd'},
              {  # '费用类型':'0:ssd',
                  '销售商': 'GF9:yyrrg', '份额类别': '*:ssd', '业务代码': '02','销售商最大折扣': '0.4'}]
        data1={'核对电子合同':'0:否','销售服务费起始日':'2017-08-01','销售服务费截止日':'2098-12-31'}
        for i in range(2):
            self.super_click(self.super_find_eles('div.fund-result-list>ul>li:nth-child({}) a'.format(i+1),
                                                  frames='frame-tab-sysinfo_fundInfo-add-fund'), mode=1, waittime=1)
            eles = self.get_eles_to_set(data[i].keys(), frames=['frame-tab-sysinfo_fundInfo-add-fund', -1])
            self.set_values(eles, data[i], 'set_value_after_add_code{}'.format(i))
            if i==1:
                eles = self.get_eles_to_set(data1.keys(), frames=['frame-tab-sysinfo_fundInfo-add-fund', -1,-1])
                self.set_values(eles, data1, 'set_value_after_add_code{}_2'.format(i))
            self.form_submit(self.super_find_eles('#finish', frames=['frame-tab-sysinfo_fundInfo-add-fund', -1]))\
                if i==1 else self.form_submit(self.super_find_eles('#dialog-btn-save'))
            f = lambda: self.super_find_eles('div.fund-result-list>ul >li:nth-child({})'.format(i+1),
                                             frames='frame-tab-sysinfo_fundInfo-add-fund').get_attribute('innerText')
            self.skip_Exception(f, lambda: f().find(' 完成') > -0)
    #
    # def del_code(self,code):
    #     self.driver.switch_to.default_content()
    #     self.super_click(self.super_find_eles('a[data-text="信息维护"]',log='login_ta',remark='信息维护'))
    #     self.super_click(self.super_find_eles('ul.up-sub-list>li',frames='frame-tab-24',remark='基金信息设置'))
    #     self.super_find_eles('input[placeholder="请输入代码或名称"]',frames='frame-tab-132').send_keys(code+Keys.ENTER)
    #     time.sleep(2)
    #     self.super_find_eles('#search-fund-list > li').click()
    #     time.sleep(2)
    #     # self.driver.switch_to.frame(self.super_find_eles('iframe',return_all=True,waittime=0.5)[-1])
    #     self.skip_Exception(lambda :self.super_find_eles('#delete-fund',frames=['frame-tab-132',-1]).click())
    #     self.driver.switch_to.default_content()
    #     self.driver.switch_to.frame('frame-tab-132')
    #     self.driver.switch_to.frame(self.super_find_eles('iframe', return_all=True, waittime=0.5)[-1])
    #     self.super_click(self.super_find_eles('#delete-dailog > div.dialog-btn > button.hs-ui-btn.hs-blue-btn.delete-btn-sure'),
    #                      mode=2,waittime=1)

def ExcelRead(sheet, row, cell):
    t=LazyExcel.ExcelRead(sheet, row, cell)[0]
    if isinstance(t,datetime):
        t=str(t).split(' ')[0]
    t=str(t)
    if t=='None':
        return ''
    return t[:-2] if t[-2:]=='.0' else t

def ExcelWrite(sheet, row, cell, contents):
    LazyExcel.ExcelWrite(sheet, row, cell, str(contents))

def getExcelSheetName(sheet):
    return LazyExcel.SheetGetName(sheet)[0]

def getExcelrows(sheet):
    return LazyExcel.SheetRowsCount(sheet)[0]

def getExceltColumns(sheet):
    return LazyExcel.SheetColumnsCount(sheet)[0]

def get_row_data(sheet,row,columns_count):
    datas=[]
    for column in range(1,columns_count):
        datas.append(ExcelRead(sheet,row,column))
    return datas

def data_is_valid(data):
    '''判断是否参数数据'''
    for x in data[1:min(10,len(data))]:
        if x:
            return True
    return False

def filter(datas,sheet_name,data_show):
    '''去隐藏'''
    ret=[]
    for data in datas:
        ret.append({key:str(float(value)*100) if key[-3:]=='(%)' else value for key,value in data.items()
                   if (key in data_show[sheet_name] if sheet_name in data_show else True)})
    return ret

