#!/usr/bin/python3
# _*_ coding: utf-8 _*_
import datetime
from mailOperate import downloadReports,sendResults
from mergeWordReport import mergeWordReport
from mergeExcelReport import mergeExcelReport
from fileUtil import scanDir,clearTempDirs,clearFiles,createTempDir,renameFile,getRelativeName

def generateHistoryReport():
    tempHistorytDirName='history_'+str(datetime.datetime.now().month)+str(datetime.datetime.now().day)    
    createTempDir(tempHistorytDirName)
    # 下载历史团队周报
    downloadReports(emailaddress,password,pop3_server,teamNumber,14,-14,'~汇总.xls',tempHistorytDirName)
    historyFiles=scanDir(tempHistorytDirName)
    histroryFileName=tempHistorytDirName+'/history.xlsx'
    renameFile(historyFiles[0],histroryFileName)
    return histroryFileName, tempHistorytDirName

def getRelativeNameList(dirName,fullNameList):
    out=[]
    for fullName in fullNameList:
        name=getRelativeName(fullName)
        out.append(dirName+'/'+name)
    return out
if __name__ == '__main__':
    emailaddress,password,pop3_server,smtp_server,teamNumber,orgName,toAddress=loadConfig()
    # 输入邮件地址, 口令和POP3服务器地址:
    # emailaddress = '18622939753@163.com'
    # # 注意使用开通POP，SMTP等的授权码
    # password = '860124Ww'
    # pop3_server = 'pop.163.com'
    # smtp_server='smtp.163.com'
    # teamNumber=7
    # orgName="软件二组"
    timeStampe=str(datetime.datetime.now().month)+str(datetime.datetime.now().day)
    tempResultDirName='result_'+timeStampe    
    createTempDir(tempResultDirName)
    histroryFileName, tempHistorytDirName = generateHistoryReport()
    # 合并到excel
    mergeExcelReport(tempResultDirName,orgName,histroryFileName)
    # 扫描成果文件夹
    resultList=scanDir(tempResultDirName)
    # 将成果作为邮件附件发送到管理邮箱中
    sendResults(getRelativeNameList(tempResultDirName,resultList),emailaddress,password,smtp_server)    
    clearTempDirs(tempHistorytDirName)
    clearTempDirs(tempResultDirName)