from playwright.sync_api import sync_playwright, Playwright
import app.UpdateStockLocationTable as UpdateStockLocationTable
import app.WMSControl as WMSControl

if __name__ == "__main__":
    with sync_playwright() as playwright:
        WMSControl.InitBrowser(playwright)
        UpdateStockLocationTable.UpdateStockLocationTable()