#!/usr/bin/python3
# _*_ coding: utf-8 _*_
import datetime
from mailOperate import downloadReports,sendResults
from mergeWordReport import mergeWordReport
from mergeExcelReport import mergeExcelReport
from fileUtil import scanDir,clearTempDirs,clearFiles,createTempDir,renameFile,loadConfig

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
    mergeWordReport(tempResultDirName,teamNumber,orgName,reportInfoList)
    clearTempDirs(tempReportDirName)