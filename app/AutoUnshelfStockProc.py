from playwright.sync_api import sync_playwright, Playwright


from . import WMSControl
from . import GenUnshelfStockTable
from . import Base

def Process():
    while 1:
        try:
            WMSControl.UpdateInfo()
            WMSControl.DownLoadDiffTable()
            WMSControl.DownLoadStockTable(r".\download\S库物料库存查询.xlsx",repositories=["制造部平面仓储5库","制造部平面仓储库","制造部智能立体库"],stockAreaCode="S")
            GenUnshelfStockTable.Process(r".\download\S库物料库存查询.xlsx",r".\download\仓储配送计划缺件执行.xlsx")
            break
        except:
            chose = input("催上架失败，输入y重试：")
            if chose != "y":
                break
            Base.Print("催上架失败重试中。。。")
    Base.Print("催上架成功。")

if __name__ == "__main__":
    with sync_playwright() as playwright:
        WMSControl.InitBrowser(playwright)
        while 1:
            WMSControl.DownLoadStockTable(r".\download\S库物料库存查询.xlsx",stockAreaCode="S")
            GenUnshelfStockTable.Process(r".\download\S库物料库存查询.xlsx",r".\download\仓储配送计划缺件执行.xlsx")
            input("已完成请按回车继续")
