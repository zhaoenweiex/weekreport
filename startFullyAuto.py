#!/usr/bin/python3
# _*_ coding: utf-8 _*_
import datetime
from mailOperate import downloadReports,sendResults
from mergeWordReport import mergeWordReport
from mergeExcelReport import mergeExcelReport
from fileUtil import scanDir,clearTempDirs,clearFiles,createTempDir,renameFile

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
    # 输入邮件地址, 口令和POP3服务器地址:
    emailaddress = '18622939753@163.com'
    # 注意使用开通POP，SMTP等的授权码
    password = '860124Ww'
    pop3_server = 'pop.163.com'
    smtp_server='smtp.163.com'
    teamNumber=7
    orgName="软件二组"
    timeStampe=str(datetime.datetime.now().month)+str(datetime.datetime.now().day)
    tempReportDirName='reports_'+timeStampe  
    tempResultDirName='result_'+timeStampe    
    createTempDir(tempReportDirName)
    createTempDir(tempResultDirName)
    # 下载团队成员周报
    downloadReports(emailaddress,password,pop3_server,teamNumber,7,-7,'周报',tempReportDirName)
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