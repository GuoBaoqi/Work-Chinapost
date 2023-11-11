from playwright.sync_api import sync_playwright, Playwright


import app.WMSControl as WMSControl
import app.GenUnshelfStockTable as GenUnshelfStockTable

def Process():
    while 1:
        try:
            WMSControl.DownLoadDiffTable()
            WMSControl.DownLoadUnshelvedStockTable()
            GenUnshelfStockTable.Process(r".\download\S库物料库存查询.xlsx",r".\download\仓储配送计划缺件执行.xlsx")
            break
        except:
            print("催上架失败自动重试中。。。")
    print("催上架成功。")

if __name__ == "__main__":
    with sync_playwright() as playwright:
        WMSControl.InitBrowser(playwright)
        while 1:
            WMSControl.DownLoadUnshelvedStockTable()
            GenUnshelfStockTable.Process(r".\download\S库物料库存查询.xlsx",r".\download\仓储配送计划缺件执行.xlsx")
            input("已完成请按回车继续")
