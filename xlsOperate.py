#!/usr/bin/python3
# _*_ coding: utf-8 _*_
import xlrd
import docx
import xlwt
import time
def merge2HistoryXlsx(resultDir,orgName,doneDict,histroryFileName):
    historyFilePath=histroryFileName
    workbook = xlrd.open_workbook(historyFilePath)
    #用索引取第一个sheet
    table = workbook.sheet_by_index(0)
    # 获取全部数据
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    row_list =[]
    for rownum in range(nrows):
         row = table.row_values(rownum)
         if row:
             row_list.append(row)
    
    #读一行数据
    memberNames = table.row_values(0)
    newRow=[]    
    timeStr= time.strftime("%Y%m%d", time.localtime())     
    newRow.append(timeStr)
    for name in memberNames:
        value=doneDict.get(name)
        newRow.append(value)
    row_list.append(newRow)
    mergeName=resultDir+'/'+orgName+timeStr+'~汇总.xls'
    write_to_excel(mergeName,row_list)

    return mergeName
def write_to_excel(filename,datas):
    wb = xlwt.Workbook()
    sheet = wb.add_sheet('test')  

    for row in range(len(datas)):
        for col in range(len(datas[row])):
            sheet.write(row, col, datas[row][col])
    wb.save(filename)
    
if __name__ == '__main__':
    merge2HistoryXlsx('软件二组',{'路研研':'测试'})