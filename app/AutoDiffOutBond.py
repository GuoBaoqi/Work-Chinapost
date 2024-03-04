from playwright.sync_api import sync_playwright, Playwright
import pandas

from . import GenInStockDiffTable
from . import WMSControl
from . import Base


def Process():
    while 1:
        try:
            WMSControl.DownLoadDiffTable()
            outbondDiffTable = GenInStockDiffTable.GenOutbondDiffTable(r".\download\仓储配送计划缺件执行.xlsx")
            WMSControl.DownLoadStockTable(r".\download\差异物料库存查询.xlsx",itemCodes = outbondDiffTable["物料编码"],repositories=["制造部平面仓储5库","制造部平面仓储库","制造部智能立体库"],itemStatus = "合格",minimumStock = "1")
            GenInStockDiffTable.GenInStockDiffTable(outbondDiffTable,r".\download\差异物料库存查询.xlsx")
            diffTable = pandas.read_excel(r".\target\在库差异表.xlsx")
            WMSControl.OutboundDiffItem(diffTable)
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
            diffTable = pandas.read_excel(r".\target\在库差异表.xlsx")
            WMSControl.OutboundDiffItem(diffTable)
            input("已完成请按回车继续")
