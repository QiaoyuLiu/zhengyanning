import dataHelper
import jsonHelper
import os
import time

def startProg(rulesFileName,config):
    properties = jsonHelper.getProperties(rulesFileName)
    fileName =config["dataDir"] + properties["fileName"]
    header = properties["indexRow"]
    rules = properties["rules"]
    selects = properties["selects"]
    outputDir = time.strftime("%Y-%m-%d",time.localtime())+"-result/"
    if not os.path.isdir(outputDir):
        os.mkdir(outputDir)
    output = outputDir+properties["outputFileName"]
    outputSheetName = properties["outputSheetName"]
    sheetName = properties["sheetName"]
    df = dataHelper.getDataFrame(fileName, header,sheetName = sheetName)
    dataFramesToAppend = []
    for rule in rules:
        res,name = dataHelper.filterFrame(rule = rule, dataFrame= df, selects = selects)
        dataFramesToAppend.append(res)
    dataFrameToWrite = dataHelper.appendDataFrames(dataFramesToAppend)
    dataHelper.writeDataToExcel(output,outputSheetName,dataFrameToWrite)

if __name__ == '__main__':
    os.chdir("c:/Users/zhengG01/Desktop/programing/Jupyterlab")
    config = jsonHelper.getConfigs()
    indexs = config["rulesIndexs"]
    for index in indexs:
        rulesFileName = config["rulesDir"]+config["rulesFiles"][index-1]
        startProg(rulesFileName,config)

    