import pandas
import re

from . import Base

不出库供应商表 = pandas.read_excel('data\\Config.xlsx',sheet_name='不出库供应商表',converters={'供应商编码':str})
不出库库区表 = pandas.read_excel('data\\Config.xlsx',sheet_name='不出库库区表')
临时不出库物料表 = pandas.read_excel('data\\Config.xlsx',sheet_name='临时不出库物料表')
排产完成状况表 = pandas.read_excel('data\\Config.xlsx',sheet_name='排产完成状况表')
缺件不出库表 = pandas.read_excel('data\\Data.xlsx',sheet_name='缺件不出库表',index_col=0)
撤单件出库表 = pandas.read_excel('data\\Data.xlsx',sheet_name='撤单件出库表',index_col=0)

def UpdateInfo():
    global 不出库供应商表
    global 不出库库区表
    global 临时不出库物料表
    global 排产完成状况表
    global 缺件不出库表
    global 撤单件出库表

    不出库供应商表 = pandas.read_excel('data\\Config.xlsx',sheet_name='不出库供应商表',converters={'供应商编码':str})
    不出库库区表 = pandas.read_excel('data\\Config.xlsx',sheet_name='不出库库区表')
    临时不出库物料表 = pandas.read_excel('data\\Config.xlsx',sheet_name='临时不出库物料表')
    排产完成状况表 = pandas.read_excel('data\\Config.xlsx',sheet_name='排产完成状况表')
    缺件不出库表 = pandas.read_excel('data\\Data.xlsx',sheet_name='缺件不出库表',index_col=0)
    撤单件出库表 = pandas.read_excel('data\\Data.xlsx',sheet_name='撤单件出库表',index_col=0)

def WriteToData(table:pandas.DataFrame,sheetName):
    with pandas.ExcelWriter('data\\Data.xlsx',mode='a',if_sheet_exists='replace') as writer:
        table.to_excel(writer,sheet_name=sheetName)

def DeleteNonOutbondSupplyerRow(diffTable:pandas.DataFrame):
    #删除表中指定供应商
    matchFlags = diffTable['供应商编码'].isin(不出库供应商表['供应商编码'])
    diffTable.drop(diffTable[matchFlags].index,inplace = True)

    #删除N开头供应商
    matchFlags=diffTable['供应商编码'].str.match('^N',na=False)
    diffTable.drop(diffTable[matchFlags].index,inplace = True)

def DeleteProductionCompletedDateRow(diffTable:pandas.DataFrame):
    #删除一线已完成日期
    matchFlags=diffTable.apply(lambda row :  (not 排产完成状况表.isnull().loc[排产完成状况表['生产日期']==row['生产日期']].iloc[0]['总装一线'])&(row['生产线']=='L1'),axis='columns')
    diffTable.drop(diffTable[matchFlags].index,inplace = True)

    #删除二线已完成日期
    matchFlags=diffTable.apply(lambda row :  (not 排产完成状况表.isnull().loc[排产完成状况表['生产日期']==row['生产日期']].iloc[0]['总装二线'])&(row['生产线']=='L2'),axis='columns')
    diffTable.drop(diffTable[matchFlags].index,inplace = True)

def DeleteNonOutbondRow(diffTable:pandas.DataFrame):
    #删除不出库条目
    matchFlags=diffTable.apply(lambda row : False if 缺件不出库表[(缺件不出库表['生产日期']==row['生产日期']) & (缺件不出库表['生产线'] == row['生产线']) & (缺件不出库表['物料编码'] == row['物料编码'])].empty else True,axis='columns')
    diffTable.drop(diffTable[matchFlags].index,inplace = True)

    #删除临时不出库条目
    matchFlags = diffTable.apply(lambda row :not 临时不出库物料表[临时不出库物料表['物料编码'] == row['物料编码']].empty,axis='columns')
    diffTable.drop(diffTable[matchFlags].index,inplace = True)

def CheckRevoke(diffTable:pandas.DataFrame):
    global 缺件不出库表
    global 撤单件出库表
    #筛选新增撤单
    matchFlags = diffTable.apply(lambda row : False if 撤单件出库表[(撤单件出库表['生产日期']==row['生产日期']) & (撤单件出库表['生产线'] == row['生产线']) & (撤单件出库表['物料编码'] == row['物料编码']) & (撤单件出库表['缺件原因'] == row['缺件原因'])].empty else True,axis='columns')
    tempTable = diffTable.drop(diffTable[matchFlags].index)
    matchFlags = tempTable['缺件原因'].str.contains('撤',na=False)
    NewRevokeTable = tempTable[matchFlags]
    if NewRevokeTable.empty:
        return
    dropIndexs = list()
    tempDropCode = list()
    for index,row in NewRevokeTable.iterrows():
        if row['物料编码'] in tempDropCode:
            continue
        Base.Print('有新增撤单需求请处理：\n'+'生产日期：'+row['生产日期']+'\n生产线：'+row['生产线']+'\n物料编码：'+row['物料编码']+'\n物料名称：'+row['物料名称']+'\n缺件原因：'+row['缺件原因'])
        while 1:
            isOutStock = input('是否出库？（y/tn/n）')
            if isOutStock == 'y':
                tempDf=row.to_frame()
                撤单件出库表 = pandas.concat([撤单件出库表, tempDf.T])
                WriteToData(撤单件出库表,'撤单件出库表')
                break
            elif isOutStock == 'n':
                tempDf=row.to_frame()
                缺件不出库表 = pandas.concat([缺件不出库表, tempDf.T])
                WriteToData(缺件不出库表,'缺件不出库表')
                dropIndexs.append(index)
                break
            elif isOutStock == 'tn':
                tempDropCode.append(row['物料编码'])
                break
    diffTable.drop(dropIndexs,inplace = True)
    #移除本次不出物料
    matchFlags = diffTable.apply(lambda row : row['物料编码'] in tempDropCode,axis='columns')
    diffTable.drop(diffTable[matchFlags].index,inplace = True)

def DeleteNonOutbondStock(stockTable:pandas.DataFrame):
    #找出不出库库区库存条目
    matchFlags = stockTable.apply(lambda row :not 不出库库区表[不出库库区表['库区编码'] == (row['库位编码'][:10] if type(row['库位编码']) == str else '')].empty,axis='columns')
    dropTable = stockTable[matchFlags].copy()

    #处理C15库区特殊库存
    searchStrBool=lambda pattern,string:not (re.search(pattern,string) is None)
    matchFlags = dropTable.apply(lambda row :((row['库位编码'] if type(row['库位编码']) == str else '') == '312302.C15.030101') | searchStrBool('取力器',row['物料名称']),axis='columns')
    dropTable.drop(dropTable[matchFlags].index,inplace = True)

    stockTable.drop(dropTable.index,inplace = True)

    #删除未锁帐为0条目
    stockTable.drop(stockTable[stockTable['未锁账数量'] == 0].index,inplace = True)

def DeleteNoStockRow(diffTable:pandas.DataFrame,stockTable:pandas.DataFrame):
    #删除15库物料
    stockTable.drop(stockTable[stockTable['仓库编码'] == '312315'].index,inplace = True)

    matchFlags = diffTable.apply(lambda row : stockTable[stockTable['物料编码'] == row['物料编码']].empty,axis='columns')
    diffTable.drop(diffTable[matchFlags].index,inplace = True)

def GenOutbondDiffTable(diffTablePath):
    UpdateInfo()
    diffTable = pandas.read_excel(diffTablePath,sheet_name='exportFile',index_col='序号',converters={'供应商编码':str})
    DeleteNonOutbondSupplyerRow(diffTable)
    DeleteProductionCompletedDateRow(diffTable)
    DeleteNonOutbondRow(diffTable)
    CheckRevoke(diffTable)
    diffTable.to_excel('target\\可出库差异表.xlsx')
    return diffTable

def GenInStockDiffTable(diffTable:pandas.DataFrame,stockTablePath):
    stockTable = pandas.read_excel(stockTablePath,sheet_name='exportFile',index_col='序号')
    DeleteNonOutbondStock(stockTable)
    stockTable.to_excel('target\\差异在库库存表.xlsx')
    DeleteNoStockRow(diffTable,stockTable)
    diffTable.to_excel('target\\在库差异表.xlsx')
    return diffTable

if __name__ == '__main__':
    while True:
        input('请更新缺件表，而后回车开始执行')
        diffTablePath='tmp\\仓储配送计划缺件执行.xlsx'
        stockTablePath = 'tmp\\物料库存查询.xlsx'
        outbondDiffTable = GenOutbondDiffTable(diffTablePath)
        outbondDiffTable['物料编码'].to_clipboard(sep=' ',index=False, header=None)
        input('已将差异物料编码复制致剪切板，请更新物料库存查询。按任意键继续')
        GenInStockDiffTable(outbondDiffTable,stockTablePath)