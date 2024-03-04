from playwright.sync_api import sync_playwright, Playwright
import pandas

from . import GenInStockDiffTable
from . import WMSControl
from . import Base

def PreprocessData():
    diffTable = pandas.read_excel(r".\target\在库差异表.xlsx")
    dateConfig = pandas.read_excel(r".\data\Config.xlsx",sheet_name = "日期配置")
    分装工位表 = pandas.read_excel(r".\data\Config.xlsx",sheet_name = "分装工位表")

    diffStockItems = pandas.DataFrame()
    indexCount = 0
    for date in diffTable['生产日期'].drop_duplicates():
        #跨分拣
        if pandas.to_datetime(date)  <= dateConfig["跨分拣截止日期"].iloc[0]:
            itemCodes = " ".join(diffTable[(diffTable['生产日期'] == date) & (diffTable['生产线'] == "L1")]["物料编码"].drop_duplicates().to_list())
            if itemCodes:
                tempDf = pandas.DataFrame({"生产日期":date,"生产线":"L1","跨分拣":"是" ,"物料编码":itemCodes},index=[indexCount])
                diffStockItems = pandas.concat([diffStockItems, tempDf])
                indexCount += 1

            itemCodes = " ".join(diffTable[(diffTable['生产日期'] == date) & (diffTable['生产线'] == "L2")]["物料编码"].drop_duplicates().to_list())
            if itemCodes:
                tempDf = pandas.DataFrame({'生产日期':date,"生产线":"L2","跨分拣":"是" ,"物料编码":itemCodes},index=[indexCount])
                diffStockItems = pandas.concat([diffStockItems, tempDf])
                indexCount += 1

        #跨分装
        elif pandas.to_datetime(date) <= dateConfig["跨分装截止日期"].iloc[0]:
            #分装工位
            itemCodes = " ".join(diffTable[diffTable.apply(lambda row : (row["生产日期"] == date) and (row["生产线"] == "L1") and (not 分装工位表[分装工位表["工位"] == row["工位"]].empty),axis='columns')]["物料编码"].drop_duplicates().to_list())
            if itemCodes:
                tempDf = pandas.DataFrame({"生产日期":date,"生产线":"L1","跨分拣":"是" ,"物料编码":itemCodes},index=[indexCount])
                diffStockItems = pandas.concat([diffStockItems, tempDf])
                indexCount += 1

            itemCodes = " ".join(diffTable[diffTable.apply(lambda row : (row["生产日期"] == date) and (row["生产线"] == "L2") and (not 分装工位表[分装工位表["工位"] == row["工位"]].empty),axis='columns')]["物料编码"].drop_duplicates().to_list())
            if itemCodes:
                tempDf = pandas.DataFrame({'生产日期':date,"生产线":"L2","跨分拣":"否","物料编码":itemCodes},index=[indexCount])
                diffStockItems = pandas.concat([diffStockItems, tempDf])
                indexCount += 1
            #非分装工位
            itemCodes = " ".join(diffTable[diffTable.apply(lambda row : (row["生产日期"] == date) and (row["生产线"] == "L1") and (分装工位表[分装工位表["工位"] == row["工位"]].empty),axis='columns')]["物料编码"].drop_duplicates().to_list())
            if itemCodes:
                tempDf = pandas.DataFrame({"生产日期":date,"生产线":"L1","跨分拣":"否" ,"物料编码":itemCodes},index=[indexCount])
                diffStockItems = pandas.concat([diffStockItems, tempDf])
                indexCount += 1

            itemCodes = " ".join(diffTable[diffTable.apply(lambda row : (row["生产日期"] == date) and (row["生产线"] == "L2") and (分装工位表[分装工位表["工位"] == row["工位"]].empty),axis='columns')]["物料编码"].drop_duplicates().to_list())
            if itemCodes:
                tempDf = pandas.DataFrame({'生产日期':date,"生产线":"L2","跨分拣":"是","物料编码":itemCodes},index=[indexCount])
                diffStockItems = pandas.concat([diffStockItems, tempDf])
                indexCount += 1

        #执行出库
        else:
            itemCodes = " ".join(diffTable[(diffTable['生产日期'] == date) & (diffTable['生产线'] == "L1")]["物料编码"].drop_duplicates().to_list())
            if itemCodes:
                tempDf = pandas.DataFrame({"生产日期":date,"生产线":"L1","跨分拣":"否" ,"物料编码":itemCodes},index=[indexCount])
                diffStockItems = pandas.concat([diffStockItems, tempDf])
                indexCount += 1

            itemCodes = " ".join(diffTable[(diffTable['生产日期'] == date) & (diffTable['生产线'] == "L2")]["物料编码"].drop_duplicates().to_list())
            if itemCodes:
                tempDf = pandas.DataFrame({'生产日期':date,"生产线":"L2","跨分拣":"否","物料编码":itemCodes},index=[indexCount])
                diffStockItems = pandas.concat([diffStockItems, tempDf])
                indexCount += 1

    diffStockItems.to_excel(r".\target\diffStockItems.xlsx")
    return diffStockItems



def Process():
    while 1:
        try:
            WMSControl.DownLoadDiffTable()
            outbondDiffTable = GenInStockDiffTable.GenOutbondDiffTable(r".\download\仓储配送计划缺件执行.xlsx")
            WMSControl.DownLoadStockTable(r".\download\差异物料库存查询.xlsx",itemCodes = outbondDiffTable["物料编码"],repositories=["制造部平面仓储5库","制造部平面仓储库","制造部智能立体库"],itemStatus = "合格",minimumStock = "1")
            GenInStockDiffTable.GenInStockDiffTable(outbondDiffTable,r".\download\差异物料库存查询.xlsx")
            diffStockItems = PreprocessData()
#            input("更改后按回车继续")
            diffStockItems = pandas.read_excel(r".\target\diffStockItems.xlsx")
            WMSControl.OutboundDiffItem(diffStockItems)
            break
        except:
            chose = input("出差异失败，输入y重试：")
            if chose != "y":
                break
            Base.Print("出差异失败重试中。。。")
    Base.Print("出差异成功。")

if __name__ == "__main__":
    with sync_playwright() as playwright:
        WMSControl.InitBrowser(playwright)
        while 1:
            WMSControl.DownLoadDiffTable()
            outbondDiffTable = GenInStockDiffTable.GenOutbondDiffTable(r".\download\仓储配送计划缺件执行.xlsx")
            WMSControl.DownLoadStockTable(r".\download\差异物料库存查询.xlsx",itemCodes = outbondDiffTable["物料编码"],repositories=["制造部平面仓储5库","制造部平面仓储库","制造部智能立体库"],itemStatus = "合格",minimumStock = "1")
            GenInStockDiffTable.GenInStockDiffTable(outbondDiffTable,r".\download\差异物料库存查询.xlsx")
            diffStockItems = PreprocessData()
            WMSControl.OutboundDiffItem(diffStockItems)
            input("已完成请按回车继续")
