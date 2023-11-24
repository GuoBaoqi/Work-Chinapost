from playwright.sync_api import sync_playwright, Playwright
import pandas

import app.WMSControl as WMSControl
import app.Base as Base



def DownloadToBeInspectedStockTable():
    diffTable = pandas.read_excel("target\\可出库差异表.xlsx")
    WMSControl.DownLoadStockTable("download\\待检差异库存查询.xlsx",itemCodes=diffTable["物料编码"],itemStatus="待检")