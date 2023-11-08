from playwright.sync_api import sync_playwright, Playwright
import pandas

import app.AutoDiffOutBond as AutoDiffOutBond
import app.AutoUnshelfStockProc as AutoUnshelfStockProc
import app.WMSControl as WMSControl

if __name__ == "__main__":
    with sync_playwright() as playwright:
        WMSControl.InitBrowser(playwright)
        while 1:
            print("出差异开始")
            AutoDiffOutBond.Process()
            print("催上架表格生成开始")
            AutoUnshelfStockProc.Process()
