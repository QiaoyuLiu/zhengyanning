import dataHelper
import jsonHelper

if __name__ == '__main__':
    properties = jsonHelper.getProperties()
    fileName = properties["fileName"]
    header = properties["indexRow"]
    rules = properties["rules"]
    selects = properties["selects"]
    output = properties["outPutFileName"]
    sheetName = properties["sheetName"]
    df = dataHelper.getDataFrame(fileName, header,sheetName = sheetName)
    dataFramesToAppend = []
    for rule in rules:
        res,name = dataHelper.filterFrame(rule = rule, dataFrame= df, selects = selects)
        dataFramesToAppend.append(res)
    dataFrameToWrite = dataHelper.appendDataFrames(dataFramesToAppend)
    dataHelper.writeDataToExcel(output,None,dataFrameToWrite)