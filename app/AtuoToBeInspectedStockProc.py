from playwright.sync_api import sync_playwright, Playwright
import pandas

from . import Base
from . import WMSControl



def DownloadToBeInspectedStockTable():
    diffTable = pandas.read_excel("target\\可出库差异表.xlsx")
    WMSControl.DownLoadStockTable("download\\待检差异库存查询.xlsx",repositories=["制造部平面仓储5库","制造部平面仓储库","制造部智能立体库"],itemCodes=diffTable["物料编码"],itemStatus="待检")