import pandas

库区表 = pandas.read_excel('data\\Config.xlsx',sheet_name='库区表')
库位表 = pandas.read_excel('data\\库位表.xlsx',sheet_name='库位表')
料名班组表 = pandas.read_excel('data\\Data.xlsx',sheet_name='料名班组表',index_col=0)

def UpdateInfo():
    global 库区表
    global 库位表
    global 料名班组表
    库区表 = pandas.read_excel('data\\Config.xlsx',sheet_name='库区表')
    库位表 = pandas.read_excel('data\\库位表.xlsx',sheet_name='库位表')
    料名班组表 = pandas.read_excel('data\\Data.xlsx',sheet_name='料名班组表',index_col=0)

def WriteToData(table:pandas.DataFrame,sheetName):
    with pandas.ExcelWriter('data\\Data.xlsx',mode='a',if_sheet_exists='replace') as writer:
        table.to_excel(writer,sheet_name=sheetName)
    

def FormatStockTable(stockTable:pandas.DataFrame):
    #移动列
    tmp = stockTable.pop('库位调整时间')
    stockTable.insert(stockTable.columns.get_loc('物料子图号'),'库位调整时间',tmp)
    tmp = stockTable.pop('库位编码')
    stockTable.insert(stockTable.columns.get_loc('供应商编码'),'库位编码',tmp)
    tmp = stockTable.pop('未锁账数量')
    stockTable.insert(stockTable.columns.get_loc('供应商编码'),'未锁账数量',tmp)
    tmp = stockTable.pop('锁帐数量')
    stockTable.insert(stockTable.columns.get_loc('供应商编码'),'锁帐数量',tmp)
    tmp = stockTable.pop('总库存数量')
    stockTable.insert(stockTable.columns.get_loc('供应商编码'),'总库存数量',tmp)

    #生成列
    stockTable.insert(stockTable.columns.get_loc('库位编码'),'上架库位',[None] * stockTable.index.size)
    stockTable.insert(stockTable.columns.get_loc('库位编码'),'差异日期',[None] * stockTable.index.size)
    
def DeleteGarbageRow(stockTable:pandas.DataFrame):
    #删除c01和c02库区以外的行
    matchFlags=stockTable['库区编码'].str.match('^3123\d\d\.S0[^12]$',na=True)
    stockTable.drop(stockTable[matchFlags].index,inplace = True)
    return 

def PartsSubCodeToStockCode(row):
    global 料名班组表
    if 库位表[库位表['物料子图号'] == row['物料子图号']]['库位编码'].empty:
        if 料名班组表[料名班组表['物料名称'] == row['物料名称']]['班组名称'].empty:
            print('有未知物料！\n物料编码：'+row['物料编码']+'\n物料名称：'+row['物料名称'])
#            teamName = input('请输入班组：')
#            if teamName != '':
#                tempdf=pandas.DataFrame({'物料名称':[row['物料名称']],'班组名称':[teamName]})
#                料名班组表 = pandas.concat([料名班组表, tempdf])
#                WriteToData(料名班组表,'料名班组表')
#                return teamName
            return None
        else:
            return 料名班组表[料名班组表['物料名称'] == row['物料名称']]['班组名称'].iloc[0]
    else:
        return 库位表[库位表['物料子图号'] == row['物料子图号']]['库位编码'].iloc[0]

def InsertData(stockTable:pandas.DataFrame,diffTable:pandas.DataFrame):
    #根据库位表填写对应库位
    stockTable['上架库位'] = stockTable.apply(PartsSubCodeToStockCode,axis = 'columns')

    #根据缺件表标记差异
    stockTable['差异日期'] = stockTable.apply(lambda r : None if diffTable[diffTable['物料编码'] == r['物料编码']]['物料编码'].empty else diffTable[diffTable['物料编码'] == r['物料编码']]['生产日期'].iloc[0],axis = 'columns')
    return

def Process(stockTablePath, diffTablePath):
    UpdateInfo()
    stockTable = pandas.read_excel(stockTablePath,sheet_name='exportFile',index_col='序号')
    diffTable = pandas.read_excel( diffTablePath,sheet_name='exportFile')
    FormatStockTable(stockTable)
    DeleteGarbageRow(stockTable)
    InsertData(stockTable,diffTable)
    stockTable.sort_values(by=['差异日期','库位调整时间'],inplace=True)
    stockTable.style.apply(lambda row : ['background-color: yellow'] * len(row),axis = 'columns')
    stockTable.to_excel('target\差异上架表.xlsx')



if __name__ == '__main__':
    while True:
        input('请更新S库与缺件表，而后回车开始执行')
        stockTablePath='tmp\\物料库存查询-S库.xlsx'
        diffTablePath='tmp\\仓储配送计划缺件执行.xlsx'
        Process(stockTablePath, diffTablePath)