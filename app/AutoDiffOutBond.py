from playwright.sync_api import sync_playwright, Playwright
import pandas

import app.GenInStockDiffTable as GenDiffInStockTable
import app.WMSControl as WMSControl

def PreprocessData():
    diffTable = pandas.read_excel(r".\target\在库差异表.xlsx")

    diffStockItems = pandas.DataFrame()
    indexCount = 0
    for date in diffTable['生产日期'].drop_duplicates():
        itemCodes = " ".join(diffTable[(diffTable['生产日期'] == date) & (diffTable['生产线'] == "L1")]["物料编码"].to_list())
        if itemCodes:
            tempDf = pandas.DataFrame({'生产日期':date,"生产线":"L1","物料编码":itemCodes},index=[indexCount])
            diffStockItems = pandas.concat([diffStockItems, tempDf])
            indexCount += 1

        itemCodes = " ".join(diffTable[(diffTable['生产日期'] == date) & (diffTable['生产线'] == "L2")]["物料编码"].to_list())
        if itemCodes:
            tempDf = pandas.DataFrame({'生产日期':date,"生产线":"L2","物料编码":itemCodes},index=[indexCount])
            diffStockItems = pandas.concat([diffStockItems, tempDf])
            indexCount += 1

    diffStockItems.to_excel(r".\target\diffStockItems.xlsx")
    return diffStockItems



def Process():
    while 1:
        try:
            WMSControl.DownLoadDiffTable()
            GenDiffInStockTable.Process(r".\download\仓储配送计划缺件执行.xlsx",r".\download\差异物料库存查询.xlsx",WMSControl.DownLoadDiffInStockTable)
            diffStockItems = PreprocessData()
            WMSControl.OutboundDiffItem(diffStockItems)
            break
        except:
            print("出差异失败自动重试中。。。")
    print("出差异成功。")

if __name__ == "__main__":
    with sync_playwright() as playwright:
        WMSControl.InitBrowser(playwright)
        while 1:
            WMSControl.DownLoadDiffTable()
            GenDiffInStockTable.Process(r".\download\仓储配送计划缺件执行.xlsx",r".\download\差异物料库存查询.xlsx",WMSControl.DownLoadDiffInStockTable)
            diffStockItems = PreprocessData()
            WMSControl.OutboundDiffItem(diffStockItems)
            input("已完成请按回车继续")
