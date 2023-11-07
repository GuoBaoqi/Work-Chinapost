from playwright.sync_api import sync_playwright, Playwright
import app.GenShelfTable as GenShelfTable

def init(playwright: Playwright):

    context = playwright.chromium.launch_persistent_context(user_data_dir="C:\\Users\\Administrator\\AppData\\Local\\Google\\Chrome\\User Data\\Default",headless=False)

#    browser = playwright.chromium.launch(headless=False)
#    context = browser.new_context()

    page = context.new_page()
    page.goto("http://wms.sinotruk.com/")
    if page.get_by_title("仓储配送计划缺件执行").is_visible():
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
    return context

def DownLoadDiffTable(context):
    page = context.new_page()
    page.goto("http://wms.sinotruk.com/") 
    #打开缺件执行
    page.get_by_role("menuitem", name="物流配送 ").hover()
    page.get_by_title("下架补货").hover()
    page.get_by_title("仓储配送计划缺件执行").click()
    #点击下箭头
    page.get_by_label("chevrondown").click()
    #输入起始日期
    page.locator(".dx-texteditor-input").first.click()
    page.locator(".dx-texteditor-input").first.fill("2023-10-30")
    page.locator(".dx-texteditor-input").first.press("Enter")
    #输入结束日期
    page.locator("div:nth-child(2) > .search-content > .search-content-textBox > .dx-datebox > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").click()
    page.locator("div:nth-child(2) > .search-content > .search-content-textBox > .dx-datebox > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").fill("2023-11-09")
    page.locator("div:nth-child(2) > .search-content > .search-content-textBox > .dx-datebox > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").press("Enter")
    #点击搜索
    page.get_by_label("find").click()
    #下载表格
    page.get_by_label("导出").click()
    with page.expect_download() as download_info:  
        page.get_by_text("导出所有数据").click()
    download = download_info.value
    download.save_as(".\download\仓储配送计划缺件执行.xlsx")
    page.close()


def DownLoadStockTable(context):
    page = context.new_page()
    page.goto("http://wms.sinotruk.com/") 
    
    #打开库存查询


    #下载表格
    page.get_by_label("导出").click()
    with page.expect_download() as download_info:  
        page.get_by_text("导出所有数据").click()
    download = download_info.value
    download.save_as(".\download\查询.xlsx")
    page.close()

if __name__ == "__main__":

    with sync_playwright() as playwright:
        context = init(playwright)
        DownLoadDiffTable(context)
        GenShelfTable.Process(".\download\仓储配送计划缺件执行.xlsx",".\download\物料库存查询.xlsx",DownLoadStockTable)
