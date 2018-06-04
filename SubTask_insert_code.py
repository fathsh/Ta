from Ta_task import *



class SubTask_insert_code(TaTask):
    def __init__(self):
        TaTask.__init__(self)
        self.data_show={'基金信息':['序号', '基金代码', '基金名称', '单位净值长度', '分红方式', '托管行名称', '清算批次号', '管理费计提比例(%)',
                        '投资方向', '份额类别', '最低募集金额', '最高募集金额', '募集起始日期', '巨额赎回比例(%)', '最低资产限额', '最低账户数量',
                        '最大账户数量', '认购利息处理方式', '基金申请确认天数', '强制赎回方式', '级差控制方式', '级差的处理方式', '存续期数(月)',
                        '户数限制处理方式', '净值公布频率按月计算方式', '净值公布月', '净值公布频率', '净值公布日', '巨额赎回顺延方式',
                        '赎回后资产低于最低账面金额处理方式', '赎回最少持有天数', '利息计算是否含费', '按收益率设置费率模式',
                        '是否只针对有限合伙人计提', '股权收益率类型', '股权收益率(%)', '股权计提比例', '安全垫产品持有天数', '销售商']}
        self.excel_datas={}
        self.onchange_key = ['基金代码']

    def get_excel_data(self):
        LazyExcel.ExcelOpen(r'e:\ta\云TA基金信息模板 v2.5（金湖&百川&混沌）.xlsx', 1)
        sheet_count = int(LazyExcel.SheetCount())
        for sheet in range(1, sheet_count + 1):
            sheet_name = getExcelSheetName(sheet)
            print(sheet_name)
            # if sheet_name not in self.data_show.keys():
            #     continue
            columns_count = getExceltColumns(sheet)
            datas = []
            for i in range(2, 999):
                data_tmp = get_row_data(sheet, i, columns_count)
                if data_is_valid(data_tmp):
                    datas.append(get_row_data(sheet, i, columns_count))
                else:
                    break
            self.excel_datas[sheet_name] = filter([dict(zip(datas[0], x)) for x in datas[1:]],sheet_name,self.data_show)  # 生成字典，去隐藏
        print(self.excel_datas['集中备份信息-第一次填写'][0])
        print(self.excel_datas['集中备份信息-第一次填写'][0].keys())

    def login_ta(self):
        self.driver.get('http://10.2.130.78:8080/bomp/login.html')
        self.driver.maximize_window()
        self.super_find_eles('#usernameInput0').send_keys('10816')
        self.super_find_eles('#passwordInput0').send_keys('123456789')
        self.super_find_eles('form').submit()

    def get_new_code(self):
        self.super_click(self.super_find_eles('a[data-text="信息维护"]',log='login_ta',remark='信息维护'))
        self.super_click(self.super_find_eles('ul.up-sub-list>li',frames='frame-tab-24',remark='基金信息设置'))
        lis_gupiao=self.super_find_eles('#id_1_0>li',frames='frame-tab-132',return_all=True)
        lis_qiquan=self.super_find_eles('#id_1_7>li',return_all=True)
        t=self.wait_Exception(lambda  :[x.get_attribute('innerText') for x in lis_qiquan][-1].split('：')[0])     #获取最后一个代码
        self.new_code=t[:3]+str(int(t[-3:])+1)

    def add_code(self):
        # self.wait_Exception(lambda :self.super_click(self.super_find_eles('#new-fund',frames=['frame-tab-132',-1],waittime=1),mode=1),
        #                     lambda :self.super_find_eles('dt', frames=['frame-tab-132', -1, -1]).get_attribute('innerText')=='*基金模板:')
        self.wait_Exception(lambda :self.super_find_eles('#new-fund', frames=['frame-tab-132', -1], waittime=1).click())
        datas={'基金模板':'210000201:一对多专户净值型产品子模板1','基金代码':'500114',
               '基金名称':'test1','TA名称':'87:广发证券股份有限公司','管理人名称':'gf0002:广发证券柜台交易市场部'}
        eles = self.get_eles_to_set(datas.keys(),frames=['frame-tab-132',-1,-1])
        self.set_values(eles,datas,'add code')
        self.super_find_eles('#dialog-btn-save').click()

    def copy_code(self):
        ele=self.super_find_eles('#fundcode-copy',frames='frame-tab-sysinfo_fundInfo-add-fund',log='add_code')
        t=time.time()
        self.set_value(ele,'870022:广发理财-招商-2',sel='#fundcode-copy')
        print('\033[1;33m{} {}\033[0m'.format('set_value select',time.time()-t))
        self.super_click(self.super_find_eles('#copy-fundinfo'))


    def compare_after_copy_code(self):
        data_sys =self.get_data_sys(frames=['frame-tab-sysinfo_fundInfo-add-fund','sysinfo_fundInfoBase-frame'],length=59)
        data_excel={'基金代码': '500114', '基金名称': '和聚(玉融)量化空盈9号私募基金', '单位净值长度': '3',
                    '分红方式': '0:再投资', '托管行名称': '605:广发证券', '清算批次号': '1：正常批次', '管理费计提比例(%)': '1.5',
                    '投资方向': '0:股票', '份额类别': 'A:前收费', '最低募集金额': '1000000', '最高募集金额': '10000000000',
                    '募集起始日期': '2015-05-16', '巨额赎回比例(%)': '20.0', '最低资产限额': '1000000.0', '最低账户数量': '1',
                    '最大账户数量': '200', '认购利息处理方式': '1:利息转份额', '基金申请确认天数': '2.0',
                    '强制赎回方式': '2:发生赎回|转换出|转托管出确认', '级差控制方式': '2:追加投资级差控制', '级差的处理方式': '2:确认整数倍金额',
                    '存续期数(月)': '1200.0', '户数限制处理方式': '1:按申请时间/申请金额/基金账号', '净值公布频率按月计算方式': '',
                    '净值公布月': '', '净值公布频率': '1:每周', '净值公布日': '5:周五', '巨额赎回顺延方式': '1:顺延到下个开放日',
                    '赎回后资产低于最低账面金额处理方式': '1:全部确认', '赎回最少持有天数': '0', '利息计算是否含费': '0：包含认购费',
                    '按收益率设置费率模式': '0:不按收益率设置费率', '是否只针对有限合伙人计提': '2:否', '股权收益率类型': '1:年化收益率',
                    '股权收益率(%)': '0', '股权计提比例': '0', '安全垫产品持有天数': '0','投资交易是否导出基金成立数据':'1:是',
                    '赎回份额明细处理方式':'0:后进先出'}
        result_compare=self.compaer_excel_sys_data(data_excel,data_sys)
        self.log_write('compare_after_copy_code')
        self.result_compare=result_compare   #such as {'基金名称': ('和聚(玉融)量化空盈9号私募基金', '广发理财-招商-2')}

    def set_value_after_compare(self):
        if not self.result_compare:                                #self.result_compare={}，比对正确
            self.super_click(self.super_find_eles('#finish',frames='frame-tab-sysinfo_fundInfo-add-fund'))
            return
        print(self.result_compare)
        eles=self.get_eles_to_set(self.result_compare.keys())
        data={key:value[0] for key,value in self.result_compare.items()}
        # print(self.result_compare)
        # print(data)
        # exit()
        self.set_values(eles,data,'set_value_after_compare')
        self.compare_after_copy_code()
        if not self.result_compare:
            self.log_write('set_value_after_compare')
            self.super_click(self.super_find_eles('#finish', frames='frame-tab-sysinfo_fundInfo-add-fund'),mode=1)
            return
        else:
            msgbox('value is wrong\n{}'.format(self.result_compare))
            time.sleep(1)

    def after_add_code0(self):
        data={'销售商':'011:xinye,040:lldj,376:陆金所,005:hhg','代销标志':'1:代销'}
        self.super_find_eles('div.fund-result-list>ul>li:nth-child(1) a',frames='frame-tab-sysinfo_fundInfo-add-fund').click()
        eles = self.get_eles_to_set(data.keys(),frames=['frame-tab-sysinfo_fundInfo-add-fund',-1])
        self.set_values(eles, data)
        self.super_find_eles('form').submit()
        f=lambda :self.super_find_eles('div.fund-result-list>ul >li:nth-child(1)',
                    frames='frame-tab-sysinfo_fundInfo-add-fund').get_attribute('innerText')
        self.wait_Exception(f,lambda :f().find(' 完成') > 0)

    def after_add_code1(self):
        data1={'销售商':'011:xinye,040:lldj,376:陆金所,005:hhg'}
        data={'核对电子合同':'0:否','销售服务费起始日':'2017-08-01','销售服务费截止日':'2098-12-31'}
        # self.super_find_eles('div.fund-result-list>ul >li:nth-child(2) a',frames='frame-tab-sysinfo_fundInfo-add-fund').click()
        self.super_click(self.super_find_eles('div.fund-result-list>ul>li:nth-child(2) a',
                                              frames='frame-tab-sysinfo_fundInfo-add-fund'),mode=1,waittime=1)
        # self.driver.switch_to.frame('layui-layer-iframe1')
        eles=self.get_eles_to_set(data1.keys(),frames=['frame-tab-sysinfo_fundInfo-add-fund',-1])
        self.set_values(eles, data1,'after_add_code1_1')
        # self.driver.switch_to.frame('sysinfo_fundinfo_fundagencyparameterAdd-frame')
        eles = self.get_eles_to_set(data.keys(),frames=['frame-tab-sysinfo_fundInfo-add-fund',-1,'sysinfo_fundinfo_fundagencyparameterAdd-frame'])
        self.set_values(eles, data,'after_add_code1_2')
        self.super_click(self.super_find_eles('#finish',frames=['frame-tab-sysinfo_fundInfo-add-fund','layui-layer-iframe1']))
        f=lambda :self.super_find_eles('div.fund-result-list>ul >li:nth-child(2)',
                    frames='frame-tab-sysinfo_fundInfo-add-fund').get_attribute('innerText')
        self.wait_Exception(f,lambda :f().find(' 完成') > -2)

    def del_code(self,code):
        self.driver.switch_to.default_content()
        self.super_click(self.super_find_eles('a[data-text="信息维护"]',log='login_ta',remark='信息维护'))
        self.super_click(self.super_find_eles('ul.up-sub-list>li',frames='frame-tab-24',remark='基金信息设置'))
        self.super_find_eles('input[placeholder="请输入代码或名称"]',frames='frame-tab-132').send_keys(code+Keys.ENTER)
        time.sleep(2)
        self.super_find_eles('#search-fund-list > li').click()
        time.sleep(2)
        # self.driver.switch_to.frame(self.super_find_eles('iframe',return_all=True,waittime=0.5)[-1])
        self.wait_Exception(lambda :self.super_find_eles('#delete-fund',frames=['frame-tab-132',-1]).click())
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame('frame-tab-132')
        self.driver.switch_to.frame(self.super_find_eles('iframe', return_all=True, waittime=0.5)[-1])
        self.super_click(self.super_find_eles('#delete-dailog > div.dialog-btn > button.hs-ui-btn.hs-blue-btn.delete-btn-sure'),
                         mode=2,waittime=1)

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
    for x in data[1:min(10,len(data))]:
        if x:
            return True
    return False

def filter(datas,sheet_name,data_show):
    ret=[]
    for data in datas:
        ret.append({key:str(float(value)*100) if key[-3:]=='(%)' else value for key,value in data.items()
                   if (key in data_show[sheet_name] if sheet_name in data_show else True)})
    return ret

