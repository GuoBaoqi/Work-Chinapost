import pandas

from . import Base
from . import WMSControl

def UpdateStockLocationTable():
    #读取现有库位表
    oldStockLocationTable = pandas.read_excel("data\\库位表.xlsx",sheet_name='库位表',index_col=0)
    
    #下载"15库,平库,立库"库存查询并合并
    WMSControl.DownLoadStockTable("download\\库存表.xlsx",repositories=["制造部平面仓储5库"],downloadTimeout=0)
    stockTable = pandas.read_excel("download\\库存表.xlsx")
    WMSControl.DownLoadStockTable("download\\库存表.xlsx",repositories=["制造部智能立体库"],downloadTimeout=0)
    stockTable = pandas.concat([stockTable, pandas.read_excel("download\\库存表.xlsx")])
    WMSControl.DownLoadStockTable("download\\库存表.xlsx",repositories=["制造部平面仓储库"],stockAreaCode="C0",downloadTimeout=0)
    stockTable = pandas.concat([stockTable, pandas.read_excel("download\\库存表.xlsx")])
    WMSControl.DownLoadStockTable("download\\库存表.xlsx",repositories=["制造部平面仓储库"],stockAreaCode="C1",downloadTimeout=0)
    stockTable = pandas.concat([stockTable, pandas.read_excel("download\\库存表.xlsx")])
    WMSControl.DownLoadStockTable("download\\库存表.xlsx",repositories=["制造部平面仓储库"],stockAreaCode="C2",downloadTimeout=0)
    stockTable = pandas.concat([stockTable, pandas.read_excel("download\\库存表.xlsx")])
    WMSControl.DownLoadStockTable("download\\库存表.xlsx",repositories=["制造部平面仓储库"],stockAreaCode="C3",downloadTimeout=0)
    stockTable = pandas.concat([stockTable, pandas.read_excel("download\\库存表.xlsx")])
    WMSControl.DownLoadStockTable("download\\库存表.xlsx",repositories=["制造部平面仓储库"],stockAreaCode="C4",downloadTimeout=0)
    stockTable = pandas.concat([stockTable, pandas.read_excel("download\\库存表.xlsx")])
    WMSControl.DownLoadStockTable("download\\库存表.xlsx",repositories=["制造部平面仓储库"],stockAreaCode="C5",downloadTimeout=0)
    stockTable = pandas.concat([stockTable, pandas.read_excel("download\\库存表.xlsx")])
    stockTable.reset_index(drop=True)
    stockTable.to_excel("target\\库存表.xlsx")

    #生成新库位表
    newStockLocationTable = pandas.DataFrame()
    newStockLocationTable.index.name = "序号"
    tmp = stockTable.pop("库位编码")
    newStockLocationTable.insert(0,"库位编码",tmp)
    tmp = stockTable.pop("物料名称")
    newStockLocationTable.insert(0,"物料名称",tmp)
    tmp = stockTable.pop("物料子图号")
    newStockLocationTable.insert(0,"物料子图号",tmp)

    #新库位表去重
    newStockLocationTable.drop_duplicates(subset=["物料子图号","库位编码"],inplace=True)
    newStockLocationTable.reset_index(drop=True)

    #删除临时库位物料
    newStockLocationTable.apply(lambda row : row["库位编码"]=="312302.C02.084206",axis='columns')
    newStockLocationTable.drop(newStockLocationTable[matchFlags].index,inplace = True)

    newStockLocationTable.apply(lambda row : row["库位编码"][:10]=="312302.C13.88",axis='columns')
    newStockLocationTable.drop(newStockLocationTable[matchFlags].index,inplace = True)

    #删除旧库位表中新库位表有数据的物料
    
    matchFlags=oldStockLocationTable.apply(lambda row : False if newStockLocationTable[(newStockLocationTable["物料子图号"]==row["物料子图号"])].empty else True,axis='columns')
    oldStockLocationTable.drop(oldStockLocationTable[matchFlags].index,inplace = True)
    oldStockLocationTable.reset_index(drop=True)

    #将新库位表中没有的历史库位加入库位表
    newStockLocationTable = pandas.concat([newStockLocationTable, oldStockLocationTable])
    newStockLocationTable.reset_index(drop=True)

    #保存新库位表
    newStockLocationTable.to_excel("data\\库位表.xlsx",sheet_name='库位表')

def Process():
    while 1:
        try:
            UpdateStockLocationTable()
            break
        except:
            chose = input("库位表更新失败，输入y重试：")
            if chose != "y":
                break
            Base.Print("库位表更新失败重试中。。。")
    Base.Print("库位表更新成功。")