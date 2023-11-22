from playwright.sync_api import sync_playwright, Playwright
import pandas

context = None
dateConfig = pandas.read_excel(r".\data\Config.xlsx",sheet_name = "日期配置")

def UpdateInfo():
    global dateConfig
    dateConfig = pandas.read_excel(r".\data\Config.xlsx",sheet_name = "日期配置")

def InitBrowser(playwright: Playwright):

    global context
    context = playwright.chromium.launch_persistent_context(chromium_sandbox = True,user_data_dir="C:\\Users\\Administrator\\AppData\\Local\\Google\\Chrome\\User Data",no_viewport=True,headless=False)

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
    page.get_by_role("textbox", name="请输入密码").fill("Zz13579**")
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

def DownLoadDiffInStockTable(diffTable:pandas.DataFrame):

    diffItem = " ".join(diffTable["物料编码"].drop_duplicates())

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
        page.locator("div:nth-child(16) > .search-content > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container").click(click_count=3,delay=100)
        page.wait_for_timeout(500)
        try:
            page.get_by_label("没有要显示的数据").get_by_text("清空").click(timeout=1500)
        except:
            pass
        page.wait_for_timeout(500)
        page.locator("div:nth-child(16) > .search-content > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").fill(row["物料编码"])
        page.locator("div:nth-child(16) > .search-content > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").press("Enter")
        #点击搜索
        page.get_by_label("find").click()
        page.wait_for_timeout(1000)
        #点击全选
        page.get_by_role("columnheader", name="全选").click()
        page.wait_for_timeout(1000)
        #出库
        if row["跨分拣"] == "是":
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

def DownLoadUnshelvedStockTable():
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
    #输入库区编码
    page.locator("#kuQuId > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").click()
    page.locator("#kuQuId > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").fill("S")
    page.locator("#kuQuId > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").press("Enter")
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
    download.save_as(".\download\S库物料库存查询.xlsx")
    page.close()
