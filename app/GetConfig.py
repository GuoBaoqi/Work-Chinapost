import pandas

def GetSubassemblyStationTable():
    subassemblyStationTable = pandas.read_excel(r".\data\Config.xlsx",sheet_name = "分装工位表")
    return subassemblyStationTable

def GetDiffStartDate():
    planStateTable = pandas.read_excel(r".\data\Config.xlsx",sheet_name = "排产状况表")
    matchFlags = planStateTable.apply(lambda row : not (row.isnull().loc["一线完成情况"] | row.isnull().loc["二线完成情况"]),axis='columns')
    planStateTable.drop(planStateTable[matchFlags].index,inplace = True)
    return planStateTable["生产日期"].min()

def GetDiffEndDate():
    planStateTable = pandas.read_excel(r".\data\Config.xlsx",sheet_name = "排产状况表")
    matchFlags = planStateTable.apply(lambda row : not (row.isnull().loc["一线完成情况"] | row.isnull().loc["二线完成情况"]),axis='columns')
    planStateTable.drop(planStateTable[matchFlags].index,inplace = True)
    return planStateTable["生产日期"].max()

def GetDiffOutboundDateTable():
    planStateTable = pandas.read_excel(r".\data\Config.xlsx",sheet_name = "排产状况表")
    matchFlags = planStateTable.apply(lambda row : not (row.isnull().loc["一线完成情况"] | row.isnull().loc["二线完成情况"]),axis='columns')
    planStateTable.drop(planStateTable[matchFlags].index,inplace = True)

    DiffOutboundDateTable = pandas.DataFrame()
    indexCount = 0
    for index,row in planStateTable.iterrows():
        if row.isnull().loc["一线完成情况"]:
            if row.isnull().loc["一线跨"]:
                tempDf = pandas.DataFrame({"生产日期":row["生产日期"],"生产线":"L1","跨分拣":"否" },index=[indexCount])
                DiffOutboundDateTable = pandas.concat([DiffOutboundDateTable, tempDf])
                indexCount += 1
            elif row["一线跨"]== "全部":
                tempDf = pandas.DataFrame({"生产日期":row["生产日期"],"生产线":"L1","跨分拣":"是" },index=[indexCount])
                DiffOutboundDateTable = pandas.concat([DiffOutboundDateTable, tempDf])
                indexCount += 1
            elif row["一线跨"]== "分装":
                tempDf = pandas.DataFrame({"生产日期":row["生产日期"],"生产线":"L1","跨分拣":"分装" },index=[indexCount])
                DiffOutboundDateTable = pandas.concat([DiffOutboundDateTable, tempDf])
                indexCount += 1
                tempDf = pandas.DataFrame({"生产日期":row["生产日期"],"生产线":"L1","跨分拣":"否" },index=[indexCount])
                DiffOutboundDateTable = pandas.concat([DiffOutboundDateTable, tempDf])
                indexCount += 1
        if row.isnull().loc["二线完成情况"]:
            if row.isnull().loc["二线跨"]:
                tempDf = pandas.DataFrame({"生产日期":row["生产日期"],"生产线":"L2","跨分拣":"否" },index=[indexCount])
                DiffOutboundDateTable = pandas.concat([DiffOutboundDateTable, tempDf])
                indexCount += 1
            elif row["二线跨"]== "全部":
                tempDf = pandas.DataFrame({"生产日期":row["生产日期"],"生产线":"L2","跨分拣":"是" },index=[indexCount])
                DiffOutboundDateTable = pandas.concat([DiffOutboundDateTable, tempDf])
                indexCount += 1
            elif row["二线跨"]== "分装":
                tempDf = pandas.DataFrame({"生产日期":row["生产日期"],"生产线":"L2","跨分拣":"分装" },index=[indexCount])
                DiffOutboundDateTable = pandas.concat([DiffOutboundDateTable, tempDf])
                indexCount += 1
                tempDf = pandas.DataFrame({"生产日期":row["生产日期"],"生产线":"L2","跨分拣":"否" },index=[indexCount])
                DiffOutboundDateTable = pandas.concat([DiffOutboundDateTable, tempDf])
                indexCount += 1
    DiffOutboundDateTable.to_excel(".\\target\\DiffOutboundDateTable.xlsx")
    return DiffOutboundDateTable