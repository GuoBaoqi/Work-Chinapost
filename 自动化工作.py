from playwright.sync_api import sync_playwright, Playwright

import app.AutoDiffOutBond as AutoDiffOutBond
import app.AutoUnshelfStockProc as AutoUnshelfStockProc
import app.WMSControl as WMSControl
import app.Base as Base

if __name__ == "__main__":
    with sync_playwright() as playwright:
        WMSControl.InitBrowser(playwright)
        while 1:
            Base.Print("准备完毕：\n1.催上架表格生成\n2.自动筛差异\n3.待检差异表生成\n请输入要做执行的任务序号：")
            chose = input()

            if chose == "1":
                Base.Print("生成催上架表格开始")
                AutoUnshelfStockProc.Process()
                Base.Print("生成催上架表格结束")
            elif chose == "2":
                Base.Print("出差异开始")
                AutoDiffOutBond.Process()
                Base.Print("出差异结束")
            elif chose == "3":
                Base.Print("待检差异表生成功能未完成，操作无效。")
