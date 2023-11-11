from playwright.sync_api import sync_playwright, Playwright
import pandas

import app.AutoDiffOutBond as AutoDiffOutBond
import app.AutoUnshelfStockProc as AutoUnshelfStockProc
import app.WMSControl as WMSControl
import app.Base as Base

if __name__ == "__main__":
    with sync_playwright() as playwright:
        WMSControl.InitBrowser(playwright)
        while 1:
            Base.Print("出差异开始")
            AutoDiffOutBond.Process()
            input("出差异成功，按任意键开始新一轮：")