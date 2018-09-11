#!/usr/bin/python3
# _*_ coding: utf-8 _*_
import datetime
from mailOperate import downloadReports,sendResults
from mergeWordReport import mergeWordReport
from mergeExcelReport import mergeExcelReport
from fileUtil import scanDir,clearTempDirs,clearFiles,createTempDir,renameFile,loadConfig

def generateHistoryReport():
    tempHistorytDirName='history_'+str(datetime.datetime.now().month)+str(datetime.datetime.now().day)    
    createTempDir(tempHistorytDirName)
    # 下载历史团队周报
    downloadReports(emailaddress,password,pop3_server,teamNumber,14,-14,'~汇总.xls',tempHistorytDirName)
    historyFiles=scanDir(tempHistorytDirName)
    histroryFileName=tempHistorytDirName+'/history.xlsx'
    renameFile(historyFiles[0],histroryFileName)
    return histroryFileName, tempHistorytDirName

if __name__ == '__main__':
    emailaddress,password,pop3_server,smtp_server,teamNumber,orgName,toAddress=loadConfig()
    timeStampe=str(datetime.datetime.now().month)+str(datetime.datetime.now().day)
    tempReportDirName='reports_'+timeStampe  
    tempResultDirName='result_'+timeStampe    
    createTempDir(tempReportDirName)
    createTempDir(tempResultDirName)
    # 下载团队成员周报
    downloadReports(emailaddress,password,pop3_server,teamNumber,14,-14,'周报',tempReportDirName)
    # 扫描文件夹
    reportInfoList=scanDir(tempReportDirName)
    # 合并到word
    wordResult = mergeWordReport(tempResultDirName,teamNumber,orgName,reportInfoList)

    histroryFileName, tempHistorytDirName = generateHistoryReport()
    # 合并到excel
    excelResult = mergeExcelReport(tempResultDirName,orgName,histroryFileName)
    # 将成果作为邮件附件发送到管理邮箱中
    sendResults([wordResult,excelResult],emailaddress,password,smtp_server)
    clearTempDirs(tempReportDirName)
    clearTempDirs(tempHistorytDirName)
    clearTempDirs(tempResultDirName)