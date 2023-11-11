import pandas
import re

import app.Base as Base
import app.GenInStockDiffTable as GenDiffInStockTable

if __name__ == '__main__':
    while True:
        input('请更新缺件表，而后回车开始执行')
        diffTablePath='download\\仓储配送计划缺件执行.xlsx'
        stockTablePath = 'download\\差异物料库存查询.xlsx',GenDiffInStockTable.Process(r".\download\仓储配送计划缺件执行.xlsx",r".\download\差异物料库存查询.xlsx",GenDiffInStockTable.downloadStockTableCallBack)
