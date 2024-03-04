from playwright.sync_api import sync_playwright, Playwright
import pandas
import time

from . import GetConfig

context = None

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
    page.locator(".dx-texteditor-input").first.fill(GetConfig.GetDiffStartDate().strftime('%Y-%m-%d'))
    page.locator(".dx-texteditor-input").first.press("Enter")
    #输入结束日期
    page.locator("div:nth-child(2) > .search-content > .search-content-textBox > .dx-datebox > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").click()
    page.locator("div:nth-child(2) > .search-content > .search-content-textBox > .dx-datebox > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").fill(GetConfig.GetDiffEndDate().strftime('%Y-%m-%d'))
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

def OutboundDiffItem(diffTable:pandas.DataFrame):
    subassemblyStationTable = GetConfig.GetSubassemblyStationTable()
    outboundTime = pandas.Timestamp.now()
    #缺件页面
    pageDiff = context.new_page()
    pageDiff.goto("http://wms.sinotruk.com/") 
    #打开缺件执行
    pageDiff.get_by_role("menuitem", name="物流配送 ").hover()
    pageDiff.get_by_title("下架补货", exact=True).hover()
    pageDiff.get_by_title("仓储配送计划缺件执行").click()
    pageDiff.wait_for_timeout(1000)
    #选择最大5000条
    pageDiff.get_by_label("Display 5000 items on page").click()

    for index,dateTableRow in GetConfig.GetDiffOutboundDateTable().iterrows():
        while True:

            itemCodes = diffTable[diffTable.apply(lambda row : (row["生产日期"] ==dateTableRow["生产日期"]) and (row["生产线"] == dateTableRow["生产线"]) and ((not subassemblyStationTable[subassemblyStationTable["工位"] == row["工位"]].empty) if dateTableRow["跨分拣"]=="分装" else True),axis='columns')]["物料编码"].drop_duplicates()

            #点击下箭头
            pageDiff.get_by_label("chevrondown").click()
            pageDiff.wait_for_timeout(1000)
            #输入起始日期
            pageDiff.locator(".dx-texteditor-input").first.click()
            pageDiff.locator(".dx-texteditor-input").first.fill(dateTableRow["生产日期"].strftime('%Y-%m-%d'))
            pageDiff.locator(".dx-texteditor-input").first.press("Enter")
            #输入结束日期
            pageDiff.locator("div:nth-child(2) > .search-content > .search-content-textBox > .dx-datebox > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").click()
            pageDiff.locator("div:nth-child(2) > .search-content > .search-content-textBox > .dx-datebox > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").fill(dateTableRow["生产日期"].strftime('%Y-%m-%d'))
            pageDiff.locator("div:nth-child(2) > .search-content > .search-content-textBox > .dx-datebox > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").press("Enter")
            #选择生产线
            pageDiff.locator(".dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-buttons-container > .dx-widget").first.click()
            pageDiff.get_by_text("总装一线" if dateTableRow["生产线"] == "L1" else "总装二线").click()
            #输入物料编码（批量）
            pageDiff.locator("div:nth-child(16) > .search-content > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container").click(position={"x": 5, "y": 5})
            pageDiff.wait_for_timeout(500)
            pageDiff.locator("div:nth-child(16) > .search-content > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").fill("未清空!")
            pageDiff.locator("div:nth-child(16) > .search-content > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").press("Enter")
            pageDiff.get_by_label("没有要显示的数据").get_by_text("清空",exact=True).click()
            pageDiff.locator("div:nth-child(16) > .search-content > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").fill(" ".join(itemCodes.to_list()))
            pageDiff.locator("div:nth-child(16) > .search-content > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").press("Enter")
            #点击搜索
            pageDiff.get_by_label("find").click()
            pageDiff.wait_for_timeout(1000)
            #点击全选
            pageDiff.get_by_role("columnheader", name="全选").click()
            pageDiff.wait_for_timeout(1000)

            #查询库存变动
            DownLoadStockTable("download\\差异物料库存查询.xlsx",itemCodes=itemCodes,repositories=["制造部平面仓储库","制造部智能立体库"],itemStatus="合格")
            diffStockTable = pandas.read_excel("download\\差异物料库存查询.xlsx")
            #如果库存有所变动
            changedDiffStockTable = diffStockTable[diffStockTable["库位调整时间"] > outboundTime]
            if not changedDiffStockTable.empty:
                matchFlags = diffTable.apply(lambda row :not changedDiffStockTable[changedDiffStockTable['物料编码'] == row['物料编码']].empty,axis='columns')
                diffTable.drop(diffTable[matchFlags].index,inplace = True)
                continue
            break

        #出库
        if dateTableRow["跨分拣"] != "否":
            pageDiff.get_by_label("跨分拣执行").click()
        else:
            pageDiff.get_by_label("执行出库").click()

        #前序继续
        try:
            pageDiff.get_by_role("button", name="继续").click(timeout=2000)
        except:
            pass

        #点击确定
        pageDiff.get_by_role("button", name="确 定").click(timeout=60000)

        #记录时间
        outboundTime = pandas.Timestamp.now()
    #收尾
    pageDiff.close()

def DownLoadStockTable(filePath:str,itemCodes:pandas.Series = None,repositories:"list[str]"=None,itemStatus:str = None,stockAreaCode:str=None,minimumStock:str=None,downloadTimeout=30000):

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
    #输入物料编码(批量)
    if not (itemCodes is None):
        itemCodes = " ".join(itemCodes.drop_duplicates())
        page.locator(".dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container").first.click()
        page.locator(".dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").first.fill(itemCodes)
        page.locator(".dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").first.press("Enter")
    #选择仓库
    if not (repositories is None):
        page.locator("div:nth-child(14) > .search-content > .search-content-textBox > div > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-buttons-container > .dx-widget").click()
        page.get_by_text("仓库编码（批量）").click()
        for repositorie in repositories:
            page.get_by_role("row", name="选择行 " + repositorie).get_by_label("选择行").click()
        page.locator("div:nth-child(14) > .search-content > .search-content-textBox > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-buttons-container > .dx-widget").click()
    #选择状态
    if not (itemStatus is None):
        page.locator("div:nth-child(19) > .search-content > .search-content-textBox > .dx-show-invalid-badge > .dx-dropdowneditor-input-wrapper > .dx-texteditor-container > .dx-texteditor-buttons-container > .dx-widget").click()
        page.get_by_role("option", name="_ " + itemStatus).locator("div").click()
    #输入库区编码
    if not (stockAreaCode is None):
        page.locator("#kuQuId > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").click()
        page.locator("#kuQuId > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").fill(stockAreaCode)
        page.locator("#kuQuId > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").press("Enter")
    #最小库存
    if not (minimumStock is None):
        page.locator("div:nth-child(31) > .search-content > .search-content-textBox > #kuCunQty > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").click()
        page.locator("div:nth-child(31) > .search-content > .search-content-textBox > #kuCunQty > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").fill(minimumStock)
        page.locator("div:nth-child(31) > .search-content > .search-content-textBox > #kuCunQty > .dx-texteditor-container > .dx-texteditor-input-container > .dx-texteditor-input").press("Enter")
    #点击上箭头
    page.get_by_label("chevronup").click()
    #点击搜索
    page.get_by_label("find").click()
    page.wait_for_timeout(1000)

    #下载表格
    page.get_by_label("导出").click()
    with page.expect_download(timeout=downloadTimeout) as download_info:  
        page.get_by_text("导出所有数据").click()
    download = download_info.value
    download.save_as(filePath)
    page.close()