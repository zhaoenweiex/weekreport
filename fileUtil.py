#!/usr/bin/python3
# _*_ coding: utf-8 _*_
import os
def scanDir(dirName):
    reportsInfo=[]
    currentPath=os.getcwd()
    for fileName in os.listdir(dirName):
        print(fileName)        
        reportsInfo.append(currentPath+'\\'+dirName+'\\'+fileName)
    return reportsInfo
def clearTempDirs(dirName):
    print("开始清理临时文件")
    currentPath=os.getcwd()
    for fileName in os.listdir(dirName):      
        os.remove(currentPath+'\\'+dirName+'\\'+fileName)
    os.removedirs(dirName)
    print("完成清理临时文件")
    return
def clearFiles(array):
    for fileName in array:      
        os.remove(fileName)
    print("完成清理临时成果")
    return
def createTempDir(name):
    if not os.path.exists(name):
        os.makedirs(name)
def renameFile(o,t):
    os.rename(o,t)

def getRelativeName(filename):
    (filepath,tempFileName) = os.path.split(filename)
    return tempFileName
def getModifyTime(name):
    return os.path.getmtime(name)