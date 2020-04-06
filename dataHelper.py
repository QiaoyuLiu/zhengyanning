import pandas as pd
import copy
import xlwt

def getDataFrame(fileName, header = 5, sheetName = 0):
    df = pd.read_excel(fileName, sheet_name = sheetName, header = header-1)
    return df


def filterFrame(dataFrame, rule, selects):
    data = copy.deepcopy(dataFrame)
    filters = []
    for prop in rule:
        if prop != "name":
            if prop == "excludes":
                excludes = rule[prop]
                for exclude in excludes:
                    excludeFilterConditions = True
                    for item in excludes[exclude]:
                        excludeFilterConditions = excludeFilterConditions & (data[exclude] !=item )
                    data = data.loc[excludeFilterConditions]
            else:
                filters.append(prop)
                filterConditions = False
                if rule[prop] != []:
                    for item in rule[prop]:
                        filterConditions = filterConditions | (item == data[prop])
                    data = data.loc[filterConditions]
    mixedFilters = filters + selects
    data = data[mixedFilters]
    #data = data.sort_values(filters,ascending = (True,True))
    data.loc[rule["name"]] = data[selects].sum()
    return data.iloc[-1: ][selects],rule["name"]


def appendDataFrames(dataFramesToAppend):
    frameList = []
    for dataFrame in dataFramesToAppend:
        frameList.append(dataFrame)
    res = pd.concat(frameList)
    return res


    
def writeDataToExcel(fileName, name, dataFrame):
    if name == None:
        dataFrame.to_excel(fileName)
    else: 
        dataFrame.to_excel(fileName,sheet_name = name)
