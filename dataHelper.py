import pandas as pd
import copy
from openpyxl import load_workbook

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
                    # Set initial condition as "True" to make sure excluding nothing if "exludes" array is empty or match nothing.
                    excludeFilterConditions = True
                    for item in excludes[exclude]:
                        # Add the target which includes item
                        if isinstance(item,dict):
                            excludeFilterConditions = excludeFilterConditions | (data[exclude].str.contains(item["like"]))
                        excludeFilterConditions = excludeFilterConditions & (data[exclude] !=item )
                    data = data.loc[excludeFilterConditions]
            else:
                filters.append(prop)
                # Similar as excludes
                filterConditions = False
                if rule[prop] != []:
                    for item in rule[prop]:
                        if isinstance(item, dict):
                            filterConditions = filterConditions | (data[prop].str.contains(item["like"]))
                            print(data[prop].str.contains(item["like"]))
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
        '''
        excelApp = xlw.App(False,False)
        exFile = excelApp.books.open(fileName)
        sheets = exFile.sheets
        sheetsNamesList = [sheet.name for sheet in sheets]
        if not name in sheetsNamesList:
            exFile.sheets.add(name)
        exFile.save(fileName)
        exFile.close()
        '''
        print(name)
        writer = pd.ExcelWriter(fileName)
        dataFrame.to_excel(writer,sheet_name = name)
        writer.save()
        writer.close()
