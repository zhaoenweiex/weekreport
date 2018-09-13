#!/usr/bin/python3
# _*_ coding: utf-8 _*_
import xlrd
import docx
import xlwt
import time


def mergeAllinfo2HistoryXlsx(resultDir, orgName, weekDict, historyFilePath):
    timeStr = time.strftime("%Y%m%d", time.localtime())
    workbook = xlrd.open_workbook(historyFilePath)
    sheets = workbook.sheets()
    array = []
    for table in sheets:
        nrows = table.nrows  # 行数
        ncols = table.ncols  # 列数
        row_list = []
        for rownum in range(nrows):
            row = table.row_values(rownum)
            if row:
                row_list.append(row)
        if not 'Sheet1' in table.name:
            array.append({'name':table.name,'datas': row_list})
        if weekDict.__contains__(table.name):
            row_list.append(weekDict.get(table.name))
            weekDict.pop(table.name)
    for member_name in weekDict:
        info=weekDict.get(member_name)
        array.append({'name':member_name,'datas': [info]})
    mergeName = resultDir+'/'+orgName+timeStr+'_all.xls'
    write_to_excel(mergeName, array)
    return mergeName


def merge2HistoryXlsx(resultDir, orgName, doneDict, histroryFileName):
    output_sheet_array=[]
    historyFilePath = histroryFileName
    workbook = xlrd.open_workbook(historyFilePath)
    # 用索引取第一个sheet
    table = workbook.sheet_by_index(0)
   
    # 获取全部数据
    nrows = table.nrows  # 行数
    ncols = table.ncols  # 列数
    row_list = []
    for rownum in range(nrows):
        row = table.row_values(rownum)
        if row:
            row_list.append(row)

    # 读一行数据
    memberNames = table.row_values(0)
    newRow = []
    timeStr = time.strftime("%Y%m%d", time.localtime())
    newRow.append(timeStr)
    for name in memberNames:
        value = doneDict.get(name)
        newRow.append(value)
    row_list.append(newRow)
    output_sheet_array.append({'name':'汇总','datas': row_list})
    sheets=workbook.sheets()
    if len(sheets)>1:
        count=0
        for sheetOfBook in sheets:
            if count>0:
                nrows_sheet = sheetOfBook.nrows  # 行数
                ncols_sheet = sheetOfBook.ncols  # 列数
                row_list_sheet = []
                for rownum in range(nrows):
                    row = sheetOfBook.row_values(rownum)
                    if row:
                        row_list_sheet.append(row)
                sheetData={'name':sheetOfBook.name,'datas':row_list_sheet}
                output_sheet_array.append(sheetData)
            count=count+1
    mergeName = resultDir+'/'+orgName+timeStr+'_汇总.xls'
    write_to_excel(mergeName,output_sheet_array)
    return mergeName



def write_to_excel(filename, arrayInfo):
    wb = xlwt.Workbook()
    for sheetData in arrayInfo:
        sheet = wb.add_sheet(sheetData.get('name'))
        datas=sheetData.get('datas')
        for row in range(len(datas)):
            for col in range(len(datas[row])):
                sheet.write(row, col, datas[row][col])
    wb.save(filename)


if __name__ == '__main__':
    merge2HistoryXlsx('软件二组', {'路研研': '测试'})
