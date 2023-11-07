from playwright.sync_api import sync_playwright, Playwright
import pandas

import app.GenInStockDiffTable as GenDiffInStockTable

context = None
dateConfig = pandas.read_excel(r".\data\Config.xlsx",sheet_name = "日期配置")

def init(playwright: Playwright):

    global context
    context = playwright.chromium.launch_persistent_context(user_data_dir="C:\\Users\\Administrator\\AppData\\Local\\Google\\Chrome\\User Data\\Default",headless=False)

#    browser = playwright.chromium.launch(headless=False)
#    context = browser.new_context()

    page = context.new_page()
    page.goto("http://wms.sinotruk.com/")
    page.wait_for_timeout(2000)
    loc=page.get_by_role("textbox", name="请输入用户名")
    if loc.count() == 0:
        page.close()
        return context

    #登录
    page.goto("http://wms.sinotruk.com/Login")
    page.locator("i").first.click()
    page.get_by_role("list").get_by_text("莱芜工厂").click()
    page.get_by_role("textbox", name="请输入用户名").click()
    page.get_by_role("textbox", name="请输入用户名").fill("LW3017")
    page.get_by_role("textbox", name="请输入密码").click()
    page.get_by_role("textbox", name="请输入密码").fill("Zz13579*")
    page.get_by_role("button", name="登录").click()
    page.wait_for_timeout(3000)
    page.close()
    return

def DownLoadDiffTable():
    page = context.new_page()
    page.goto("http://wms.sinotruk.com/") 
    #打开缺件执行
    page.get_by_role("menuitem", name="物流配送 ").hover()
    page.get_by_title("下架补货").hover()
    page.get_by_title("仓储配送计划缺件执行").click()
    page.wait_for_timeout(1000)
    #点击下箭头
    page.get_by_label("chevrondown").click()
    #输入起始日期
    page.locator(".dx-texteditor-input").first.click()
    page.locator(".dx-texteditor-input").first.fill(dateConfig["起始日期"].iloc[0].strftime('%Y-%m-%d'))
    page.locator(".dx-texteditor-input").first.press("Enter")
    #输入结束日期
    page.locator("div:nth-child(2) > .search-content > .search-content-textBox > .dx-datebox > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").click()
    page.locator("div:nth-child(2) > .search-content > .search-content-textBox > .dx-datebox > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").fill(dateConfig["截止日期"].iloc[0].strftime('%Y-%m-%d'))
    page.locator("div:nth-child(2) > .search-content > .search-content-textBox > .dx-datebox > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").press("Enter")
    #点击搜索
    page.get_by_label("find").click()
    page.wait_for_timeout(1000)
    #下载表格
    page.get_by_label("导出").click()
    with page.expect_download() as download_info:  
        page.get_by_text("导出所有数据").click()
    download = download_info.value
    download.save_as(r".\download\仓储配送计划缺件执行.xlsx")
    page.close()
    return


def DownLoadStockTable(diffTable:pandas.DataFrame):

    diffItem = " ".join(diffTable["物料编码"])

    page = context.new_page()
    page.goto("http://wms.sinotruk.com/") 
    
    #打开库存查询
    page.get_by_role("menuitem", name="报表管理 ").hover()
    page.wait_for_timeout(200)
    page.get_by_title("库存查询", exact=True).hover()
    page.wait_for_timeout(200)
    page.get_by_title("物料库存查询").click()
    page.wait_for_timeout(1000)
    #点击下箭头
    page.get_by_label("chevrondown").click()
    #输入物料编码
    page.locator(".dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container").first.click()
    page.locator(".dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").first.fill(diffItem)
    page.locator(".dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").first.press("Enter")
    #选择仓库
    page.locator("div:nth-child(14) > .search-content > .search-content-textBox > div > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-buttons-container > .dx-widget").click()
    page.get_by_text("仓库编码（批量）").click()
    page.get_by_role("row", name="选择行 制造部平面仓储5库").get_by_label("选择行").click()
    page.get_by_role("row", name="选择行 制造部平面仓储库").get_by_label("选择行").click()
    page.get_by_role("row", name="选择行 制造部智能立体库").get_by_label("选择行").click()
    page.locator("div:nth-child(14) > .search-content > .search-content-textBox > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-buttons-container > .dx-widget").click()
    #选择合格
    page.locator("div:nth-child(19) > .search-content > .search-content-textBox > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-buttons-container > .dx-widget").click()
    page.get_by_role("option", name="_ 合格").locator("div").click()
    #最小库存1
    page.locator("div:nth-child(31) > .search-content > .search-content-textBox > #kuCunQty > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").click()
    page.locator("div:nth-child(31) > .search-content > .search-content-textBox > #kuCunQty > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").fill("1")
    page.locator("div:nth-child(31) > .search-content > .search-content-textBox > #kuCunQty > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").press("Enter")
    #点击上箭头
    page.get_by_label("chevronup").click()
    #点击搜索
    page.get_by_label("find").click()
    page.wait_for_timeout(1000)

    #下载表格
    page.get_by_label("导出").click()
    with page.expect_download() as download_info:  
        page.get_by_text("导出所有数据").click()
    download = download_info.value
    download.save_as(r".\download\差异物料库存查询.xlsx")
    page.close()

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

def OutboundDiffItem(diffStockItems:pandas.DataFrame):
    page = context.new_page()
    page.goto("http://wms.sinotruk.com/") 
    #打开缺件执行
    page.get_by_role("menuitem", name="物流配送 ").hover()
    page.get_by_title("下架补货", exact=True).hover()
    page.get_by_title("仓储配送计划缺件执行").click()
    page.wait_for_timeout(1000)
    #选择最大5000条
    page.get_by_label("Display 5000 items on page").click()

    for index,row in diffStockItems.iterrows():
        #点击下箭头
        page.get_by_label("chevrondown").click()
        page.wait_for_timeout(1000)
        #输入起始日期
        page.locator(".dx-texteditor-input").first.click()
        page.locator(".dx-texteditor-input").first.fill(row["生产日期"])
        page.locator(".dx-texteditor-input").first.press("Enter")
        #输入结束日期
        page.locator("div:nth-child(2) > .search-content > .search-content-textBox > .dx-datebox > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").click()
        page.locator("div:nth-child(2) > .search-content > .search-content-textBox > .dx-datebox > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").fill(row["生产日期"])
        page.locator("div:nth-child(2) > .search-content > .search-content-textBox > .dx-datebox > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").press("Enter")
        #选择生产线
        page.locator(".dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-buttons-container > .dx-widget").first.click()
        page.get_by_text("总装一线" if row["生产线"] == "L1" else "总装二线").click()
        #输入物料编码（批量）
        page.locator(".search-content > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").press("Backspace")
        page.locator(".search-content > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container").click()
        page.locator(".search-content > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").fill(row["物料编码"])
        page.locator(".search-content > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").press("Enter")
        #点击搜索
        page.get_by_label("find").click()
        page.wait_for_timeout(1000)
        #点击全选
        page.get_by_role("columnheader", name="全选").click()
        page.wait_for_timeout(1000)
        #出库
        if pandas.to_datetime(row["生产日期"]) <= dateConfig["跨分拣截止日期"].iloc[0]:
            page.get_by_label("跨分拣执行").click()
        else:
            page.get_by_label("执行出库").click()

        #前序继续
        try:
            page.get_by_role("button", name="继续").click()
        except:
            pass
        #点击确定
        page.get_by_role("button", name="确 定").click()
    page.close()

def Process(argContext)
    global context
    context = argContext
    while 1:
        try:
            DownLoadDiffTable()
            GenDiffInStockTable.Process(r".\download\仓储配送计划缺件执行.xlsx",r".\download\差异物料库存查询.xlsx",DownLoadStockTable)
            diffStockItems = PreprocessData()
            OutboundDiffItem(diffStockItems)
            break
        except:
            print("出差异失败自动重试中。。。")

if __name__ == "__main__":
    with sync_playwright() as playwright:
        init(playwright)
        while 1:
            DownLoadDiffTable()
            GenDiffInStockTable.Process(r".\download\仓储配送计划缺件执行.xlsx",r".\download\差异物料库存查询.xlsx",DownLoadStockTable)
            diffStockItems = PreprocessData()
            OutboundDiffItem(diffStockItems)
            input("已完成请按回车继续")
