import pandas

def GetDiffStartDate():
    planStateTable = pandas.read_excel(r".\data\Config.xlsx",sheet_name = "排产状况表")
    matchFlags = planStateTable.apply(lambda row : not (row["总装一线"].isnull() | row["总装二线"].isnull()),axis='columns')
    planStateTable.drop(planStateTable[matchFlags].index,inplace = True)
    return planStateTable["生产日期"].min()

def GetDiffEndDate():
    planStateTable = pandas.read_excel(r".\data\Config.xlsx",sheet_name = "排产状况表")
    matchFlags = planStateTable.apply(lambda row : not (row["总装一线"].isnull() | row["总装二线"].isnull()),axis='columns')
    planStateTable.drop(planStateTable[matchFlags].index,inplace = True)
    return planStateTable["生产日期"].max()

def GetDiffOutboundDateTable():
    planStateTable = pandas.read_excel(r".\data\Config.xlsx",sheet_name = "排产状况表")
    matchFlags = planStateTable.apply(lambda row : not (row["总装一线"].isnull() | row["总装二线"].isnull()),axis='columns')
    planStateTable.drop(planStateTable[matchFlags].index,inplace = True)
    return planStateTable