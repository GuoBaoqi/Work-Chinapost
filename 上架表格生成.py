import app.GenUnshelfStockTable as GenUnshelfStockTable
import app.Base as Base

if __name__ == '__main__':
    while True:
        Base.Print('请更新S库与缺件表，而后回车开始执行')
        input()
        stockTablePath='download\\S库物料库存查询.xlsx'
        diffTablePath='download\\仓储配送计划缺件执行.xlsx'
        GenUnshelfStockTable.Process(stockTablePath,diffTablePath)