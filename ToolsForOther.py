__author__ = 'swallow'
__language__= 'python 3.0'

from email.errors import StartBoundaryNotFoundDefect
from fileinput import filename
import os
import sys
import time
from datetime import datetime

import json
import random
from weakref import ref
from cv2 import FileNode_NAMED
import requests
import re

import pythoncom
import win32com
from win32com.client import Dispatch,DispatchEx
from PIL import ImageGrab, Image

import CommonFunc
from CommonFunc import isEmptyValue
from CommonFunc import copyspecFile
from CommonFunc import getCellValueinString
from CommonFunc import rgb_to_hex
from CommonFunc import findFile
from shutil import copyfile
import filecmp
#import pymysql as MySqldb

#import myDB 

AMEND_COLOR = (255,102,255)
ERROR_COLOR = (255,0,0)
REMOVE_COLOR = (127,127,127)
RED_COLOR = (255, 0, 0)
BLACK_COLOR = (0,0,0)
WHITE_COLOR = (255,255,255)
YELLOW_COLOR = (255,255,0)

DIFF_COLOR = (122,201,250)
EXPORT_COLOR = (250,161,232)
CHECK_COLOR = (215,251,157)
OBJ_COLOR = (157,175,249)

PURPLE_COLOR = (204,153,255)
PINK_COLOR = (255,153,204)
GRAY_COLOR = (192,192,192)

START_ROWNO = 1
COLUMN_MAX = 11

header = {

    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',

    'Accept-Encoding': 'gzip, deflate, sdch',

    'Accept-Language': 'zh-CN,zh;q=0.8',

    'Connection': 'keep-alive',

    'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/43.0.235'

}

timeout = random.choice(range(80, 180))



class  ToolsFixer(object):
    result_list = []
    mulity_state_list = []
    shapes_list0 = {}
    shapes_list1 = {}

    excelApp = None
    resultBook = None

    def __init__(self,path='.'):
        self._path=path
        self.abspath=os.path.abspath(self._path) # 默认当前目录
        self.excelApp = DispatchEx("Excel.Application")
        self.excelApp.visible = True
        self.excelApp.DisplayAlerts = False

        self.excelApp.Workbooks.Add()
        self.resultBook = self.excelApp.ActiveWorkBook

    def __exit_(self, *args):
        self.resultBook.Close(SaveChanges = 1)
        self.excelApp.Application.Quit()


    def getFilename_fromCd(self, cd):
        """
        Get filename from content-disposition
        """
        if not cd:
            return None
        
        fname = re.findall('filename=(.+)', cd)
        
        if len(fname) == 0:
            return None
        
        return fname[0]

    def SaveResultFile(self, resultName):
        resultFileName  = os.getcwd() + "\\" + resultName
        if os.path.exists(resultFileName): 
            os.remove(resultFileName)

        self.resultBook.SaveAs(resultFileName)        

    def testDB(self, root):
        for root, dirs,files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root,name)

                print('Search in this ->',fileName)    # 绝对路径
                file = os.getcwd() + "\\" + fileName
                costWb  = self.excelApp.Workbooks.Open(file)
                specSheet = costWb.Sheets('DATA') 

                info = specSheet.UsedRange
                nrows = info.Rows.Count
                row = 2

                while row <= nrows:
                    self.insertData(specSheet, row)
                    row += 1

    def insertData(self, sheet, rowNo):

        mydb = myDB.MysqlDb()
        mydb.initDB()

        localtime = time.strftime("%Y%m%d%H%M%S", time.localtime())

        insertSql = "INSERT INTO specid_table  VALUES ('" \
                    + getCellValueinString(sheet.Cells(rowNo, 8)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 1)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 2)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 3)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 4)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 5)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 6)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 7)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 9)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 10)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 11)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 12)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 13)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 14)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 15)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 16)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 17)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 18)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 19)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 20)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 21)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 22)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 23)) + "','" \
                    + getCellValueinString(sheet.Cells(rowNo, 24)) + "','" \
                    + localtime + "');"


        print(insertSql)
        count = mydb.insert_data(insertSql, None)

        mydb.close_connect()
        result = {'inserted_item_count': count}

        return result


    def selectDB(self):
        mydb = myDB()
        mydb.initDB()

        selectSql = "SELECT DISTINCT * FROM specid_table;"

        # print (select_sql)

        result = mydb.select_data(selectSql)
        mydb.close_connect()

        return result['result']


    def testCellValue(self, root):

        for root, dirs,files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root,name)

                print('Search in this ->',fileName)    # 绝对路径
                file = os.getcwd() + "\\" + fileName
                costWb  = self.excelApp.Workbooks.Open(file)
                specSheet = costWb.Sheets('test') 

                info = specSheet.UsedRange
                nrows = info.Rows.Count
                row = 2

                while row <= nrows:
                    mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                    spec_id0 =''
                    spec_id0 = spec_id0.join([getCellValueinString(specSheet.Cells(row, 2)), "_", getCellValueinString(specSheet.Cells(row, 3)) ,"_", getCellValueinString(specSheet.Cells(row, 4)) ,"_", getCellValueinString(specSheet.Cells(row, 5))])
                    print(spec_id0)
                    row += mergeCount
    
    def printSheetName(self, root):
        for root, dirs,files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root,name)

                print('Search in this ->',fileName)    # 绝对路径
                file = os.getcwd() + "\\" + fileName
                costWb  = self.excelApp.Workbooks.Open(file)
                sheetCount = (costWb.Worksheets.Count)

                for i in range(1, sheetCount + 1):
                    sheet_name0 = costWb.Worksheets(i).Name
                    print(sheet_name0)
                    i = i + 1

                costWb.Close(SaveChanges = 0)

    def MergeCostFile(self, root):

        CostTotalfile = os.getcwd() + "\\res\\Cost\\【21mm中国】式样变更管理表-合并.xlsx"
        CostTotalWb  = self.excelApp.Workbooks.Open(CostTotalfile)
        CostTotalSheet = CostTotalWb.Sheets('工数预估') 

        for root, dirs,files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root,name)

                try:
                    print('Search in this ->',fileName)    # 绝对路径
                    file = os.getcwd() + "\\" + fileName
                    costWb  = self.excelApp.Workbooks.Open(file)
                    costSheet = costWb.Sheets('test') 

                    rowNo  = 11

                    info = costSheet.UsedRange
                    nrows = info.Rows.Count

                    while rowNo < nrows:
                        if costSheet.Cells(rowNo, 103).Value is not None:
                            self.FindInTotalfile(CostTotalSheet, costSheet, rowNo)
                        rowNo += 1

                    costWb.Close(SaveChanges = 0)
                except Exception as e:
                    print(e)


        CostTotalWb.Close(SaveChanges = 1)

    def FindInTotalfile(self, CostTotalSheet, costSheet, keyRowNo):

        keycol = 0
        keyValue = ""

        if isEmptyValue(costSheet.Cells(keyRowNo, 6)) is not None:
            keyValue = getCellValueinString(costSheet.Cells(keyRowNo, 6))
            keycol =6
        elif isEmptyValue(costSheet.Cells(keyRowNo, 5)) is not None:
            keyValue = getCellValueinString(costSheet.Cells(keyRowNo, 5))
            keycol =5
        elif isEmptyValue(costSheet.Cells(keyRowNo, 4)) is not None:
            keyValue = getCellValueinString(costSheet.Cells(keyRowNo, 4))
            keycol =4
        elif isEmptyValue(costSheet.Cells(keyRowNo, 3)) is not None:
            keycol =3
            keyValue = getCellValueinString(costSheet.Cells(keyRowNo, 3))
        elif isEmptyValue(costSheet.Cells(keyRowNo, 2)) is not None:
            keycol =2
            keyValue = getCellValueinString(costSheet.Cells(keyRowNo, 2))
        else:
            keycol =1
            keyValue = getCellValueinString(costSheet.Cells(keyRowNo, 1))


        info = CostTotalSheet.UsedRange
        nrows = info.Rows.Count
        group_str =  getCellValueinString(costSheet.Cells(keyRowNo, 103))

        rowNo  = 11
        while rowNo < nrows:
            KeyInTotal = getCellValueinString(CostTotalSheet.Cells(rowNo,keycol))
            if keyValue == KeyInTotal:
                mergeCount = CostTotalSheet.Cells(rowNo, 4).MergeArea.Rows.Count
                for i in range(0, mergeCount):
                    group_total_str = CostTotalSheet.Cells(rowNo + i, 103).Value
                    if  group_total_str is not None:
                        if group_total_str == group_str:
                            range_s= "AA" + str(keyRowNo) + ":" + "DA" + str(keyRowNo)
                            range_t= "AA" + str(rowNo) + ":" + "DA" + str(rowNo)
                            
                            costSheet.Range(range_s).Copy(CostTotalSheet.Cells.Range(range_t))
                            CostTotalSheet.Cells(rowNo, 10).Value = "UPD" + group_str
                            CostTotalSheet.Cells(rowNo, 10).interior.color = rgb_to_hex((147,205,221))
                            break
                    rowNo += 1
                range_a = "A" + str(rowNo) + ":DA" + str(rowNo) 
                CostTotalSheet.Range(range_a).EntireRow.Insert()
                range_s= "AA" + str(keyRowNo) + ":" + "DA" + str(keyRowNo)
                range_t= "AA" + str(rowNo) + ":" + "DA" + str(rowNo)
                costSheet.Range(range_s).Copy(CostTotalSheet.Cells.Range(range_t))
                CostTotalSheet.Cells(rowNo, 10).interior.color = rgb_to_hex((205,205,221))
                CostTotalSheet.Cells(rowNo, 10).Value = "INS" + group_str

                itemMergeCnt = CostTotalSheet.Cells(rowNo , keycol).MergeArea.Rows.Count
                if itemMergeCnt == 1:
                    merge_range = chr(ord("A") + keycol - 1) + str(rowNo - itemMergeCnt -1 ) + ":" + chr(ord("A") + keycol - 1)  + str(rowNo)
                    #CostTotalSheet.Range(merge_range).Merge()

                nrows += 1
                break

            rowNo += 1

    def FindSpecPara(self, root):

        wRow = 1
        for root, dirs, files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root,name)
                try:
                    print('Search in this ->',fileName)    # 绝对路径
                    file = os.getcwd() + "\\" + fileName
                    specWb  = self.excelApp.Workbooks.Open(file)
                    sheetCount = (specWb.Worksheets.Count)
                    wRow += 1                   
                    catalogSheet = specWb.Sheets('Catalog')

                    copytoRow  = 100
                    for i in range(1, sheetCount):
                        specSheet = specWb.Worksheets(i)
                        if specSheet.Name not in ("Catalog", "変更履歴","Cover","DropDownDataList","propertyDownDataListCCtype"):
                            info = specSheet.UsedRange
                            nrows = info.Rows.Count
                            ncolumns = info.Columns.Count
                            row = 1
                            print("Starting to process " + specSheet.Name)
                            while row < nrows:
                                #cellValue = getCellValueinString(specSheet.Cells(row, 3))
                                if specSheet.Cells(row, 3).Value is not None:
                                    if  specSheet.Cells(row, 3).Value == "View of Screen":
                                        while (specSheet.Cells(row, 5).Value != "Parts Name") :
                                            row += 1

                                        startRow = row
                                        while not isEmptyValue(specSheet.Cells(row,2)):
                                            mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                                            row += mergeCount
                                        endRow = row - 1

                                        if specSheet.Cells(startRow, 19).Value != "Display Condition" :
                                            copyspecFile(fileName, "wrong format")
                                            break
                                        range_s= "S" + str(startRow) + ":" + "Y" + str(endRow)
                                        range_t= "B" + str(copytoRow) + ":" + "H" + str(copytoRow + endRow - startRow )
                                        specSheet.Range(range_s).Copy(catalogSheet.Cells.Range(range_t))
                                        range_s= "M" + str(startRow) + ":" + "R" + str(endRow)
                                        range_t= "I" + str(copytoRow) + ":" + "N" + str(copytoRow + endRow - startRow)                                    
                                        specSheet.Range(range_s).Copy(catalogSheet.Cells.Range(range_t))

                                        range_s= "B" + str(copytoRow) + ":" + "N" + str(copytoRow + endRow - startRow)
                                        range_t= "M" + str(startRow) + ":" + "Y" + str(endRow)       
                                        catalogSheet.Range(range_s).Copy(specSheet.Cells.Range(range_t))

                                        catalogSheet.Range(range_s).EntireRow.Delete()

                                    if  specSheet.Cells(row, 3).Value == "View of Soft Button":
                                        while (specSheet.Cells(row, 5).Value != "Button Name") :
                                            row += 1

                                        startRow = row
                                        while not isEmptyValue(specSheet.Cells(row,2)):
                                            mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                                            row += mergeCount
                                        endRow = row - 1
                            
                                        if specSheet.Cells(startRow, 19).Value != "Condition" :
                                            copyspecFile(fileName, "wrong format")
                                            break
                                        
                                        range_s= "S" + str(startRow) + ":" + "W" + str(endRow)
                                        range_t= "B" + str(copytoRow) + ":" + "F" + str(copytoRow + endRow - startRow )
                                        specSheet.Range(range_s).Copy(catalogSheet.Cells.Range(range_t))
                                        range_s= "I" + str(startRow) + ":" + "R" + str(endRow)
                                        range_t= "G" + str(copytoRow) + ":" + "P" + str(copytoRow + endRow - startRow)                                    
                                        specSheet.Range(range_s).Copy(catalogSheet.Cells.Range(range_t))

                                        range_s= "B" + str(copytoRow) + ":" + "P" + str(copytoRow + endRow - startRow)
                                        range_t= "I" + str(startRow) + ":" + "W" + str(endRow)       
                                        catalogSheet.Range(range_s).Copy(specSheet.Cells.Range(range_t))

                                        catalogSheet.Range(range_s).EntireRow.Delete()

                                    if  specSheet.Cells(row, 3).Value == "Soft Button Action":
                                        while (specSheet.Cells(row, 5).Value != "Button Name") :
                                            row += 1

                                        startRow = row
                                        while not isEmptyValue(specSheet.Cells(row,2)):
                                            mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                                            row += mergeCount
                                        endRow = row - 1
                                        
                                        if specSheet.Cells(startRow, 19).Value != "Condition of Action" :
                                            copyspecFile(fileName, "wrong format")
                                            break

                                        range_s= "S" + str(startRow) + ":" + "W" + str(endRow)
                                        range_t= "B" + str(copytoRow) + ":" + "F" + str(copytoRow + endRow - startRow )
                                        specSheet.Range(range_s).Copy(catalogSheet.Cells.Range(range_t))
                                        range_s= "L" + str(startRow) + ":" + "R" + str(endRow)
                                        range_t= "G" + str(copytoRow) + ":" + "M" + str(copytoRow + endRow - startRow)                                    
                                        specSheet.Range(range_s).Copy(catalogSheet.Cells.Range(range_t))

                                        range_s= "B" + str(copytoRow) + ":" + "M" + str(copytoRow + endRow - startRow)
                                        range_t= "L" + str(startRow) + ":" + "W" + str(endRow)       
                                        catalogSheet.Range(range_s).Copy(specSheet.Cells.Range(range_t))

                                        catalogSheet.Range(range_s).EntireRow.Delete()
                                        specSheet.Columns("L:L").ColumnWidth = 8
                                
                                row += 1

                    specWb.Close(SaveChanges = 1)


                except Exception as e:
                    print(e)
                    print('!!! Error in ->', fileName)    # 绝对路径
                    copyspecFile(fileName, "error_files")

        print('...........................................')

    def FindStringID(self, root):
        self.resultBook.ActiveSheet.Name = "string id"
        resultSheet = self.resultBook.ActiveSheet
        resultSheet.Range("A1, F1000").Font.name = "Microsoft YaHei UI"

        wRow = 1
        for root, dirs, files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root,name)
                try:
                    print('Search in this ->',fileName)    # 绝对路径
                    file = os.getcwd() + "\\" + fileName
                    specWb  = self.excelApp.Workbooks.Open(file)
                    sheetCount = (specWb.Worksheets.Count)
                    wRow += 1                   
                    catalogSheet = specWb.Sheets('Catalog')

                    copytoRow  = 100
                    for i in range(1, sheetCount):
                        specSheet = specWb.Worksheets(i)
                        if specSheet.Name not in ("Catalog", "変更履歴","Cover","DropDownDataList","propertyDownDataListCCtype","History"):
                            info = specSheet.UsedRange
                            nrows = info.Rows.Count
                            ncolumns = info.Columns.Count
                            row = 1
                            print("Starting to find stringID " + specSheet.Name)

                            prefixString = "SP_" + specSheet.Name + "_"
                            while row < nrows:
                                #cellValue = getCellValueinString(specSheet.Cells(row, 3))
                                if specSheet.Cells(row, 3).Value is not None:
                                    if  specSheet.Cells(row, 9).Value == "Display Content":
                                        row += 1

                                        while ((specSheet.Cells(row, 3).Value != "View of Soft Button")):
                                            mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                                            content_value = getCellValueinString(specSheet.Cells(row, 32))
                                            if ((content_value.find("文言") >= 0) or  (content_value.find("ボタン") >= 0)):
                                                resultSheet.Cells(wRow, 1).Value = prefixString + getCellValueinString(specSheet.Cells(row,2)) + "_"\
                                                    + getCellValueinString(specSheet.Cells(row,3)) + "_" + getCellValueinString(specSheet.Cells(row,4)) + "_" + getCellValueinString(specSheet.Cells(row,5))
                                                resultSheet.Cells(wRow, 2).Value = specSheet.Cells(row,26).Value

                                                wRow += 1
                                            row += mergeCount

                                        break
                                row += 1

                    specWb.Close(SaveChanges = 0)


                except Exception as e:
                    print(e)
                    print('!!! Error in ->', fileName)    # 绝对路径
                    copyspecFile(fileName, "error_files")

        print('...........................................')

    def SearchScreenID(self, root, ressultFileName):
        self.resultBook.ActiveSheet.Name = "screenlist"
        resultSheet = self.resultBook.ActiveSheet
        resultSheet.Range("A1, F1000").Font.name = "Microsoft YaHei UI"
        wRow = 1
        for root, dirs, files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root,name)
                try:
                    print('Search in this ->',fileName)    # 绝对路径
                    file = os.getcwd() + "\\" + fileName
                    specWb  = self.excelApp.Workbooks.Open(file)
                    sheetCount = (specWb.Worksheets.Count)
                    resultSheet.Cells(wRow, 1).Value = fileName
                    wRow += 1                    
                    for i in range(1, sheetCount):
                        specSheet = specWb.Worksheets(i)
                        if specSheet.Name not in ("Catalog", "変更履歴","History"):
                            info = specSheet.UsedRange
                            nrows = info.Rows.Count
                            ncolumns = info.Columns.Count


                    specWb.Close(SaveChanges = 0)

                except Exception as e:
                    print(e)
                    print('!!! Error in ->', fileName)    # 绝对路径
                    copyspecFile(fileName, "error_files")

    def CheckScreenUUID(self, root):

        pythoncom.CoInitialize()
        print('...........................................')

        for root0, dirs,files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root0, name)
                try:
                    print('Research this ->',fileName)    # 绝对路径
                    file = os.getcwd() + "\\" + fileName
                    specWb  = self.excelApp.Workbooks.Open(file)
                    sheetCount = (specWb.Worksheets.Count)
                    specCatalog = specWb.Sheets("Catalog")

                    bSave = False
                    wRow = 50
                    for i in range(1, sheetCount):
                        bChange = False
                        sheet_name = specWb.Worksheets(i).Name
                        
                        if sheet_name not in ("Cover", "Catalog","History","DropDownDataList","propertyDownDataListCCtype"):
                            specSheet = specWb.Worksheets(i)
                            info = specSheet.UsedRange
                            nrows = info.Rows.Count
                            ncolumns = info.Columns.Count
                            row = 1
                            print("Starting to process " + sheet_name)

                            screen_uuid = getCellValueinString(specSheet.Cells(2,47))

                            row = 3
                            while row < nrows:
                                if not isEmptyValue(specSheet.Cells(row, 47)):
                                    item_uuid = getCellValueinString(specSheet.cells(row, 47))

                                    if item_uuid not in ("UUID","Invalud_1","-"):
                                        if item_uuid.startswith(screen_uuid) is not True:
                                            para_no = getCellValueinString(specSheet.Cells(row,2))
                                            key_value = getCellValueinString(specSheet.Cells(row,6))
                                            if para_no == "4":
                                                if key_value == "PWR":
                                                    specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_HK_Panel_PWR_1"
                                                elif key_value == "VOL UP":
                                                    specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_HK_Panel_VOL_UP_1"
                                                elif key_value == "VOL DOWN":
                                                    specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_HK_Panel_VOL_DOWN_1"
                                                elif key_value == "VOL +":
                                                    specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_HK_Steering_Switch_VOL_+_1"
                                                elif key_value == "VOL -":
                                                    specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_HK_Steering_Switch_VOL_-_1"
                                                elif key_value == "Track/Seek/Ch +":
                                                    specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_HK_Steering_Switch_Track/Seek/Ch_+_1"
                                                elif key_value == "Track/Seek/Ch -":
                                                    specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_HK_Steering_Switch_Track/Seek/Ch_-_1"
                                                elif key_value == "MODE":
                                                    specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_HK_Steering_Switch_MODE_1"
                                                elif key_value == "PTT":
                                                    specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_HK_Steering_Switch_PTT_1"
                                                elif key_value == "OnHook":
                                                    specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_HK_Steering_Switch_OnHook_1"
                                                elif key_value == "OffHook":
                                                    specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_HK_Steering_Switch_OffHook_1"
                                                elif key_value == "TEL SW":
                                                    specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_HK_Steering_Switch_TEL_SW_1"
                                            elif para_no == "5":
                                                specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_Init_"+ getCellValueinString(specSheet.Cells(row,2)) \
                                                                            + getCellValueinString(specSheet.Cells(row,3)) + getCellValueinString(specSheet.Cells(row,4))
                                            elif para_no == "6":
                                                specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_Status_"+ getCellValueinString(specSheet.Cells(row,2)) \
                                                                            + getCellValueinString(specSheet.Cells(row,3)) + getCellValueinString(specSheet.Cells(row,4)) \
                                                                            + getCellValueinString(specSheet.Cells(row,5))
                                            elif para_no == "7":
                                                specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_Trans_"+ getCellValueinString(specSheet.Cells(row,2)) \
                                                                            + getCellValueinString(specSheet.Cells(row,3)) + getCellValueinString(specSheet.Cells(row,4)) \
                                                                            + getCellValueinString(specSheet.Cells(row,5))
                                            elif para_no == "8":
                                                specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_Trig_"+str(row)
                                            else:
                                                specSheet.cells(row, 47).Value = screen_uuid + "->" + screen_uuid + "_XXXX_"+str(row)

                                            specSheet.cells(row, 47).interior.color = rgb_to_hex(AMEND_COLOR)
                                            bSave = True
                                            bChange = True
                                
                                row += 1
                        if bChange:
                            #specCatalog.Cells(wRow, 2).Value = sheet_name
                            wRow += 1

                    if bSave:
                        specWb.Close(SaveChanges = 1)
                        copyspecFile(fileName, "uuid_files")
                    else:
                        specWb.Close(SaveChanges = 0)

                except  Exception as e:                     
                    print(e)
                    print('!!! Error in ->', fileName)    # 绝对路径
                    copyspecFile(fileName, "error_files")

    def CheckScreenSpecID(self, root):

        pythoncom.CoInitialize()
        print('...........................................')

        for root0, dirs,files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root0, name)
                try:
                    print('Research this ->',fileName)    # 绝对路径
                    file = os.getcwd() + "\\" + fileName
                    specWb  = self.excelApp.Workbooks.Open(file)
                    sheetCount = (specWb.Worksheets.Count)
                    specCatalog = specWb.Sheets("Catalog")
                    
                    bSave = False
                    for i in range(1, sheetCount):
                        sheet_name = specWb.Worksheets(i).Name

                        if sheet_name not in ("Cover", "Catalog","History","DropDownDataList","propertyDownDataListCCtype"):
                            specSheet = specWb.Worksheets(i)
                            info = specSheet.UsedRange
                            nrows = info.Rows.Count
                            ncolumns = info.Columns.Count
                            row = 1
                            print("Starting to process " + sheet_name)

                            specId_list = {}

                            row = 3
                            while row < nrows:
                                mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                                if not isEmptyValue(specSheet.Cells(row, 47)):
                                    item_uuid = getCellValueinString(specSheet.cells(row, 47))

                                    if item_uuid not in ("UUID","Invalud_1","-"):
                                        para_no = getCellValueinString(specSheet.Cells(row,2))
                                        key_value = getCellValueinString(specSheet.Cells(row,6))
                                        spec_id =''
                                        spec_id = spec_id.join([getCellValueinString(specSheet.Cells(row, 2)), "_", getCellValueinString(specSheet.Cells(row, 3)), \
                                                    "_", getCellValueinString(specSheet.Cells(row, 4)) ,"_", getCellValueinString(specSheet.Cells(row, 5))])
                                        
                                        if spec_id in specId_list:
                                            specSheet.Cells(row, 1).Value = "Duplicated spec id"
                                            specSheet.Cells(row, 1).interior.color = rgb_to_hex(RED_COLOR)
                                            bSave = True
                                        else:
                                            specId_list[spec_id] = getCellValueinString(specSheet.Cells(row, 26))
                                
                                row += mergeCount

                    if bSave:
                        specWb.Close(SaveChanges = 1)
                        copyspecFile(fileName, "specid_duplicated")
                    else:
                        specWb.Close(SaveChanges = 0)

                except  Exception as e:                     
                    print(e)
                    print('!!! Error in ->', fileName)    # 绝对路径
                    copyspecFile(fileName, "error_files")


    def DeleteExportedOnly(self, root, opeFlag):

        rptfile = os.getcwd() + "\\doc\\report.xlsx"
        rptWb  = self.excelApp.Workbooks.Open(rptfile)
        rptSheet = rptWb.Sheets("issues")

        rptinfo = rptSheet.UsedRange
        nRptRrows = rptinfo.Rows.Count
        ncolumns = rptinfo.Columns.Count
        rptRow = 2
        
        screen_list = {}

        while rptRow <= nRptRrows:
            screen_list[getCellValueinString(rptSheet.Cells(rptRow, 4))] = str(rptRow)
            rptRow += 1

        pythoncom.CoInitialize()
        print('...........................................')

        for root0, dirs,files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root0, name)
                try:
                    print('Research this ->',fileName)    # 绝对路径
                    file = os.getcwd() + "\\" + fileName
                    specWb  = self.excelApp.Workbooks.Open(file)
                    sheetCount = (specWb.Worksheets.Count)
                    specCatalog = specWb.Sheets("Catalog")

                    bSave = False
                    wRow = 50
                    for i in range(1, sheetCount):
                        bChange = False
                        sheet_name = specWb.Worksheets(i).Name
                        
                        if sheet_name not in ("Cover", "Catalog","History","DropDownDataList","propertyDownDataListCCtype"):
                            specSheet = specWb.Worksheets(i)
                            info = specSheet.UsedRange
                            nrows = info.Rows.Count
                            ncolumns = info.Columns.Count
                            row = 1
                            print("Starting to process " + sheet_name)

                            screen_uuid = getCellValueinString(specSheet.Cells(2,47))

                            row = 3
                            point_cnt = 0
                            empty_cnt = 0
                            while row < nrows:
                                mergeCnt = specSheet.Cells(row,2).MergeArea.Rows.count
                                item_tag = getCellValueinString(specSheet.cells(row, 1))

                                if item_tag.upper() == "EXPORTED ONLY":
                                    if opeFlag == "DEL_MARK":
                                        spec_id = getCellValueinString(specSheet.Cells(row, 2))\
                                                + "_" + getCellValueinString(specSheet.Cells(row, 3))\
                                                + "_" + getCellValueinString(specSheet.Cells(row, 4))\
                                                + "_" + getCellValueinString(specSheet.Cells(row, 5))\
                                                + "_" + getCellValueinString(specSheet.Cells(row, 6))
                                        print("delete in ", sheet_name + "：" +spec_id + ":" + item_tag)
                                        mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                                        specSheet.Range("A" + str(row) + ":BM" + str(row + mergeCount -1)).EntireRow.Delete()
    
                                        bSave = True
                                        bChange = True
                                        row -= mergeCnt
                                    else:
                                        point_cnt += 1

                                elif item_tag.upper().find("SAME") >= 0:
                                    if opeFlag == "DEL_MARK":
                                        specSheet.cells(row, 1).Value = item_tag.replace("same", "")
                                        bSave = True
                                elif item_tag.upper().find("DIFF") >= 0 :
                                    if opeFlag == "DEL_MARK":
                                        specSheet.cells(row, 1).Value = item_tag.replace("diff", "")
                                        bSave = True
                                    else:
                                        point_cnt += 1
                                elif item_tag.upper().find("THIS ONLY") >= 0:
                                    if opeFlag == "DEL_MARK":
                                        specSheet.cells(row, 1).Value = item_tag.replace("This only", "")
                                        bSave = True
                                    else:
                                        point_cnt += 1
                                elif specSheet.Cells(row, 2).interior.color == rgb_to_hex(EXPORT_COLOR):
                                    if opeFlag == "DEL_MARK":
                                        print("delete pink line ", sheet_name)
                                        mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                                        specSheet.Range("A" + str(row) + ":BM" + str(row + mergeCount -1)).EntireRow.Delete()
    
                                        bSave = True
                                        bChange = True
                                        row -= mergeCnt
                                elif specSheet.Cells(row, 2).Font.Strikethrough == True:
                                    if opeFlag == "DEL_MARK":
                                        specSheet.Cells(row, 47).Value = getCellValueinString(specSheet.Cells(row, 47)) +"_UUID_Deleted_Item_" + str(row)

                                i = 1
                                end_flag = False
                                while(isEmptyValue(specSheet.Cells(row + i, 3))):
                                    i += 1
                                    if i > 20:
                                        end_flag = True
                                        break

                                if end_flag:
                                    break
                                row += mergeCnt
                            
                            try:
                                rptSheet.Cells(int(screen_list[sheet_name]), 9).Value = point_cnt
                            except Exception as e:
                                print("no this screen name", sheet_name)

                            if opeFlag == "DEL_MARK":
                                specSheet.Range("A1:A"+str(row)).interior.color = rgb_to_hex(WHITE_COLOR)

                        if bChange:
                            #specCatalog.Cells(wRow, 2).Value = sheet_name
                            wRow += 1

                    specWb.Close(SaveChanges = 1)

                except  Exception as e:                     
                    print(e)
                    print('!!! Error in ->', fileName)    # 绝对路径
                    copyspecFile(fileName, "error_files")
        
        if opeFlag != "DEL_MARK":
            rptWb.Close(SaveChanges = 1)
        else:
            rptWb.Close(SaveChanges = 0)

    def GetCompareFileList(self, root0, root1, compare_list):
        old_root = root0
        new_root = root1

        for root0,dirs,files in os.walk(root0):
            for name in files: 
                old_file = os.path.join(root0, name)
                if name in self.result_list:
                    file_ext = os.path.splitext(name)[1].upper()                
                    if file_ext in ('.XL', '.XLSX'):
                        name0 = name.split(".")[0]
                        new_name = name0 + "_newexported.xlsx"
                        new_name = new_name.replace("[", "")
                        new_name = new_name.replace("]", "")
                        new_name = new_name.replace(" ", "")
                        root1 = root0.replace(old_root, new_root)
                        new_orig_file = os.path.join(root1, name)
                        if os.path.exists(new_orig_file):
                            file1 = os.path.join(root1, name)
                            new_shortname_file = os.path.join(root1, new_name)
                            if not os.path.exists(new_shortname_file):
                                os.rename(file1, new_shortname_file)
                            compare_list[old_file] = new_shortname_file
                else:
                    start_fileName = name.split(".")[0]
                    curfile = []
                    foundfile = findFile(start_fileName, root1, curfile)
                    if foundfile is not None:
                        compare_list[old_file]= foundfile

    def WanderInFiles(self, root0, root1, cmp_key):
        filelist=[]
        compare_list={}
        self.GetCompareFileList(root0, root1, compare_list)

        print('...........................................')
        
        rptfile = os.getcwd() + "\\doc\\report_round_workshop.xlsx"
        rptWb  = self.excelApp.Workbooks.Open(rptfile)
        rptSheet = rptWb.Sheets("issues")

        rptinfo = rptSheet.UsedRange
        nRptRrows = rptinfo.Rows.Count
        ncolumns = rptinfo.Columns.Count
        rptRow = 2
        
        screen_list = {}

        while rptRow <= nRptRrows:
            screen_list[getCellValueinString(rptSheet.Cells(rptRow, 6))] = str(rptRow)
            rptRow += 1
        
        pythoncom.CoInitialize()

        resultSheet = self.resultBook.Sheets.Add(Before = None, After = self.resultBook.Sheets(self.resultBook.Sheets.count))
        resultSheet.Name = "PartsID"

        for fileName, compilefile in compare_list.items():            
            if os.path.isfile(fileName):
                try:
                    print('Research this ->',fileName)    # 绝对路径
                    file1 = os.getcwd() + "\\" + fileName
                    specWb0  = self.excelApp.Workbooks.Open(file1)
                    sheetCount0 = (specWb0.Worksheets.Count)


                    file2 =  os.getcwd() + "\\" + compare_list[fileName]
                    specWb1  = self.excelApp.Workbooks.Open(file2)       
                    sheetCount1 = (specWb1.Worksheets.Count)


                    for i in range(1, sheetCount0 + 1):
                        sheet_name0 = specWb0.Worksheets(i).Name
                        
                        if sheet_name0 not in ("Cover", "Catalog","History","DropDownDataList","propertyDownDataListCCtype"):
                            specSheet0 = specWb0.Worksheets(i)
                            info = specSheet0.UsedRange
                            nrows = info.Rows.Count
                            ncolumns = info.Columns.Count
                            row = 1
                            print("Starting to process " + sheet_name0)

                            specSheet1 = None
                            try:
                                specSheet1 = specWb1.Sheets(sheet_name0)
                            except  Exception as e:                     
                                print("Not find this sheet in exported file:", sheet_name0)

                            if specSheet1 is not None:
                                partlist0 = {}
                                partlist1 = {}

                                para_row_list = {}
                                
                                partlist1["problem_cnt"] = "0"

                                wRowAfterVersion0 = self.FillDifferenceContent(specSheet1, partlist1, "GETLIST", None, para_row_list, cmp_key)
                                lastRow = self.FillDifferenceContent(specSheet0, partlist1, "CMP", specSheet1, para_row_list, cmp_key)
                                self.PrintRemainItem(specSheet0, specSheet1, lastRow, partlist1, para_row_list)
                                print("Issue in this screen ", sheet_name0 + ":" + partlist1["problem_cnt"])
                                try:
                                    rptSheet.Cells(int(screen_list[sheet_name0]), 10).Value = partlist1["problem_cnt"]
                                except Exception as e:
                                    print("no this screen name", sheet_name0)

                    specWb0.Close(SaveChanges = 1)
                    specWb1.Close(SaveChanges = 0)
                    
                except  Exception as e:                     
                    print(e)
                    print('!!! Error in ->', fileName)    # 绝对路径
                    copyspecFile(fileName, "error_files")
    

        #self.SaveResultFile(resultName)
        pythoncom.CoUninitialize()

    def PrintRemainItem(self, specSheet, diffSheet, lastRow, partlist, para_row_list):
        problem_cnt = int(partlist["problem_cnt"])
        startRow = lastRow + 1
        for key in partlist.keys():
            try:            
                if key != "problem_cnt":
                    uuid_id = key
                    value = partlist[uuid_id]
                    row_in_other = value.split("<row>")[1]
                    mergeCnt = diffSheet.Cells(row_in_other,2).MergeArea.Rows.count

                    if uuid_id.find("_HK_") >= 0:
                        print("Abnormal HarkKey found " + uuid_id)
                        startRow = int(para_row_list["para_4_row"])
                    elif uuid_id.find("_Init_") >= 0:
                        startRow = int(para_row_list["para_5_row"])
                    elif uuid_id.find("_Trig_") >= 0:
                        startRow = int(para_row_list["para_8_row"])
                    elif uuid_id.find("_Status_") >= 0:
                        startRow = int(para_row_list["para_6_row"])
                    elif uuid_id.find("_Trans_") >= 0:
                        startRow = int(para_row_list["para_7_row"])
                    elif uuid_id.startswith('a'):
                        startRow = int(para_row_list["para_1_row"])
                    elif uuid_id.startswith('b'):
                        startRow = int(para_row_list["para_2_row"])
                    elif uuid_id.startswith('c'):
                        startRow = int(para_row_list["para_3_row"])
                    else:
                        startRow = lastRow + 1

                    range_a = "A" + str(startRow) + ":BM" + str(startRow + mergeCnt -1)
                    range_b = "A" + str(row_in_other) + ":BM" + str(int(row_in_other) + mergeCnt -1)

                    specSheet.Range(range_a).EntireRow.Insert()
                    diffSheet.Range(range_b).Copy(specSheet.Cells.Range(range_a))
                    specSheet.Cells(startRow,1).Value = "Exported Only"
                    specSheet.Range(range_a).interior.color = rgb_to_hex(OBJ_COLOR)
                    problem_cnt += 1

                    if int(para_row_list["para_1_row"]) >= startRow: 
                        para_row_list["para_1_row"] = str(int(para_row_list["para_1_row"]) + mergeCnt)
                    if int(para_row_list["para_2_row"]) >= startRow: 
                        para_row_list["para_2_row"] = str(int(para_row_list["para_2_row"]) + mergeCnt)
                    if int(para_row_list["para_3_row"]) >= startRow: 
                        para_row_list["para_3_row"] = str(int(para_row_list["para_3_row"]) + mergeCnt)
                    if int(para_row_list["para_4_row"]) >= startRow: 
                        para_row_list["para_4_row"] = str(int(para_row_list["para_4_row"]) + mergeCnt)
                    if int(para_row_list["para_5_row"]) >= startRow: 
                        para_row_list["para_5_row"] = str(int(para_row_list["para_5_row"]) + mergeCnt)
                    if int(para_row_list["para_6_row"]) >= startRow: 
                        para_row_list["para_6_row"] = str(int(para_row_list["para_6_row"]) + mergeCnt)
                    if int(para_row_list["para_7_row"]) >= startRow: 
                        para_row_list["para_7_row"] = str(int(para_row_list["para_7_row"]) + mergeCnt)
                    if int(para_row_list["para_8_row"]) >= startRow: 
                        para_row_list["para_8_row"] = str(int(para_row_list["para_8_row"]) + mergeCnt)

                    lastRow += mergeCnt                
            except Exception as e:
                print(e, key)
        partlist["problem_cnt"] = str(problem_cnt)

    def FillDifferenceContent(self, specSheet, partlist,ope_flag, diffSheet, para_row_list, cmp_key):
        row = START_ROWNO                        

        # Find in first version
        info = specSheet.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count

        while row < nrows:
            
            # 0 - Outline
            if (getCellValueinString(specSheet.Cells(row, 3)) == "Outline"):
                if ope_flag == "CMP":
                    spec_str0 = getCellValueinString(specSheet.Cells(row+1, 3))
                    self.DiffWithExportSheetForComment(spec_str0, specSheet, "outline", row, partlist)
                else:
                    partlist["outline"] = getCellValueinString(specSheet.Cells(row+1, 3))
                
                row += 1

            # 1 View of Screen
            if getCellValueinString(specSheet.Cells(row, 3)) is not None:
                if  specSheet.Cells(row, 3).Value == "View of Screen":
                    while (specSheet.Cells(row, 5).Value != "Parts Name") :
                        row += 1

                    comment_para_1 = ""
                    while not isEmptyValue(specSheet.Cells(row,2)):
                        mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                        uuid_id = getCellValueinString(specSheet.Cells(row, 47))
                        if uuid_id is not None: # is part
                            if uuid_id != "UUID":
                                uuid_id = ''
                                uuid_id = uuid_id .join(["a", getCellValueinString(specSheet.Cells(row, 47))])
                                spec_id0 =''
                                spec_id0 = spec_id0.join([getCellValueinString(specSheet.Cells(row, 2)), "_", getCellValueinString(specSheet.Cells(row, 3)) ,"_", getCellValueinString(specSheet.Cells(row, 4)) ,"_", getCellValueinString(specSheet.Cells(row, 5))])
                                part_name0 = getCellValueinString(specSheet.Cells(row, 6))
                                Display_Content0 = getCellValueinString(specSheet.Cells(row, 9))
                                Conditional_combination0 = getCellValueinString(specSheet.Cells(row, 13))
                                Conditional_description0 = getCellValueinString(specSheet.Cells(row, 16))
                                Display_Condition0 = ""
                                for i in range(0, mergeCount):
                                    Display_Condition0 += getCellValueinString(specSheet.Cells(row + 1, 20))

                                Display_In_Such_Condition0 = ""
                                for i in range(0, mergeCount):
                                    Display_In_Such_Condition0 = Display_In_Such_Condition0.join(getCellValueinString(specSheet.Cells(row + 1, 26)))
                                Property_Type0 = getCellValueinString(specSheet.Cells(row, 32))
                                Data_Range0 = getCellValueinString(specSheet.Cells(row, 38))
                                Remark0 = getCellValueinString(specSheet.Cells(row, 41))

                                spec_str0 = ''
                                spec_str0 = spec_str0.join([spec_id0, str(part_name0), str(Display_Content0), Conditional_combination0, Conditional_description0, Display_Condition0, Display_In_Such_Condition0 ,Property_Type0, Data_Range0, Remark0])
                                
                                if cmp_key == "UUID":
                                    key_str = uuid_id
                                else:
                                    key_str = spec_id0
                                                            
                                if ope_flag == "CMP":
                                    mergeCntDiff = self.DiffWithExportSheet(spec_str0, specSheet, diffSheet, key_str, row, partlist)
                                    row += mergeCntDiff
                                    nrows += mergeCntDiff

                                else:
                                    spec_str0 += "<row>" + str(row)
                                    partlist[key_str] = spec_str0

                                            
                        row += mergeCount
                    para_row_list["para_1_row"] =  str(row)
                    while not isEmptyValue(specSheet.Cells(row,3)):
                        if getCellValueinString(specSheet.Cells(row, 3)) == "View of Soft Button":
                            break
                        comment_para_1 += getCellValueinString(specSheet.Cells(row, 3))
                        row += 1

                    if ope_flag == "CMP":
                        spec_str0 = comment_para_1
                        self.DiffWithExportSheetForComment(spec_str0, specSheet, "comment_para_1", row, partlist)
                    else:
                        partlist["comment_para_1"]  = comment_para_1 + "<row>" + str(row)

            # 2 View of Soft Button
            if  specSheet.Cells(row, 3).Value == "View of Soft Button":
                    while (specSheet.Cells(row, 5).Value != "Button Name") :
                        row += 1

                    comment_para_2 = ""
                    while not isEmptyValue(specSheet.Cells(row,2)):
                        mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                        uuid_id = getCellValueinString(specSheet.Cells(row, 47))
                        if uuid_id is not None  : # is button
                            if uuid_id != "UUID":
                                uuid_id = "b"+ getCellValueinString(specSheet.Cells(row, 47))
                                spec_id0 = getCellValueinString(specSheet.Cells(row, 2)) \
                                        + "_" + getCellValueinString(specSheet.Cells(row, 3)) \
                                        + "_" + getCellValueinString(specSheet.Cells(row, 4)) \
                                        + "_" + getCellValueinString(specSheet.Cells(row, 5))

                                button_name0 = getCellValueinString(specSheet.Cells(row, 6))
                                Conditional_combination0 = getCellValueinString(specSheet.Cells(row, 9))
                                Conditional_description0 = getCellValueinString(specSheet.Cells(row, 13))
                                Condition0 = ""
                                for i in range(0, mergeCount):
                                    Condition0 += getCellValueinString(specSheet.Cells(row + 1, 20))
                                
                                Display_In_Such_Condition0 = ""
                                for i in range(0, mergeCount):
                                    Display_In_Such_Condition0 += getCellValueinString(specSheet.Cells(row + 1, 24))

                                Property_Type0 = getCellValueinString(specSheet.Cells(row, 32))
                                DuringDriving0 = getCellValueinString(specSheet.Cells(row, 38))
                                Remark0 = getCellValueinString(specSheet.Cells(row, 41))

                                if cmp_key == "UUID":
                                    key_str = uuid_id
                                else:
                                    key_str = spec_id0

                                if ope_flag == "CMP":
                                    spec_str0 = spec_id0 + button_name0+ Conditional_combination0 + Conditional_description0 + Condition0 \
                                            + Display_In_Such_Condition0 + Property_Type0 + DuringDriving0 + Remark0
                                    mergeCntDiff = self.DiffWithExportSheet(spec_str0, specSheet, diffSheet, key_str, row, partlist)
                                    row += mergeCntDiff
                                    nrows += mergeCntDiff
                                else:
                                    partlist[key_str] = spec_id0 + button_name0+ Conditional_combination0 \
                                                    + Conditional_description0 + Condition0 + Display_In_Such_Condition0 \
                                                    + Property_Type0 + DuringDriving0 + Remark0 + "<row>" + str(row)
                                
                        row += mergeCount

                    para_row_list["para_2_row"] =  str(row)

                    while not isEmptyValue(specSheet.Cells(row,3)):
                        if getCellValueinString(specSheet.Cells(row, 3)) == "Soft Button Action":
                            break
                        comment_para_2 += getCellValueinString(specSheet.Cells(row, 3))
                        row += 1

                    if ope_flag == "CMP":
                        spec_str0 = comment_para_2
                        self.DiffWithExportSheetForComment(spec_str0, specSheet, "comment_para_2", row, partlist)
                    else:
                        partlist["comment_para_2"]  = comment_para_2 + "<row>" + str(row)

            # 3 Soft Button Action
            if  specSheet.Cells(row, 3).Value == "Soft Button Action":
                while (specSheet.Cells(row, 5).Value != "Button Name") :
                    row += 1

                comment_para_3 = ""
                while not isEmptyValue(specSheet.Cells(row,2)):
                    mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                    uuid_id = getCellValueinString(specSheet.Cells(row, 47))
                    if uuid_id is not None: # is button
                        if uuid_id != "UUID" :
                            uuid_id = "c"+ getCellValueinString(specSheet.Cells(row, 47))
                            spec_id0 = getCellValueinString(specSheet.Cells(row, 2))  \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 3)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 4)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 5))

                            button_name0 = getCellValueinString(specSheet.Cells(row, 6))
                            Operation_type0 = getCellValueinString(specSheet.Cells(row, 9))
                            Conditional_combination0 = getCellValueinString(specSheet.Cells(row, 12))
                            Conditional_description0 = getCellValueinString(specSheet.Cells(row, 14))
                            Condition_Action0 = ""
                            for i in range(0, mergeCount):
                                Condition_Action0 += getCellValueinString(specSheet.Cells(row + 1, 20))

                            Action_In_Such_Condition0 = ""    
                            for i in range(0, mergeCount):
                                Action_In_Such_Condition0 += getCellValueinString(specSheet.Cells(row + 1, 24))

                            transition0 = getCellValueinString(specSheet.Cells(row, 32))
                            Sound0 = getCellValueinString(specSheet.Cells(row, 35))       
                            Remark0 = getCellValueinString(specSheet.Cells(row, 41))

                            if cmp_key == "UUID":
                                key_str = uuid_id
                            else:
                                key_str = spec_id0

                            if ope_flag == "CMP":
                                spec_str0 = spec_id0 + button_name0+ Operation_type0 + Conditional_combination0 \
                                        + Conditional_description0 + Condition_Action0 + Action_In_Such_Condition0 + transition0 + Sound0 + Remark0
                                mergeCntDiff = self.DiffWithExportSheet(spec_str0, specSheet, diffSheet, key_str,row, partlist)
                                row += mergeCntDiff
                                nrows += mergeCntDiff
                            else:
                                partlist[key_str] = spec_id0 + button_name0+ Operation_type0 + Conditional_combination0 \
                                                + Conditional_description0 + Condition_Action0 + Action_In_Such_Condition0 \
                                                + transition0 + Sound0 + Remark0 + "<row>" + str(row)


                    row += mergeCount
                para_row_list["para_3_row"] =  str(row)

                while not isEmptyValue(specSheet.Cells(row,3)):
                    if getCellValueinString(specSheet.Cells(row, 3)) == "Hard Key Action":
                        break
                    comment_para_3 += getCellValueinString(specSheet.Cells(row, 3))
                    row += 1

                if ope_flag == "CMP":
                    spec_str0 = comment_para_3
                    self.DiffWithExportSheetForComment(spec_str0, specSheet, "comment_para_3", row, partlist)
                else:

                    partlist["comment_para_3"]  = comment_para_3 + "<row>" + str(row)
                    
            # 4 - Hardkey 
            if  specSheet.Cells(row, 3).Value == "Hard Key Action":
                comment_para_4 = ""
                row += 2
                if  getCellValueinString(specSheet.Cells(row, 4)) == "＜Panel＞":
                    row += 1
                
                while not isEmptyValue(specSheet.Cells(row,2)):
                    mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                    if  getCellValueinString(specSheet.Cells(row, 4)) ==  "＜Steering Switch＞" :
                        row += mergeCount

                    mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count

                    uuid_id = getCellValueinString(specSheet.Cells(row, 47))
                    if uuid_id is not None: # is button
                        if uuid_id != "UUID" :
                            spec_id0 = getCellValueinString(specSheet.Cells(row, 2)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 3)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 4)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 5))

                            key_name0 = getCellValueinString(specSheet.Cells(row, 6))
                            Key_Spec_str0 = getCellValueinString(specSheet.Cells(row, 9)) \
                                        + getCellValueinString(specSheet.Cells(row, 12)) \
                                        + getCellValueinString(specSheet.Cells(row, 14)) \
                                        + getCellValueinString(specSheet.Cells(row, 35)) \
                                        + getCellValueinString(specSheet.Cells(row, 38)) \
                                        + getCellValueinString(specSheet.Cells(row, 41))

                            Key_Spec_Condition_Str = ""
                            for i in range(0, mergeCount):
                                Key_Spec_Condition_Str += getCellValueinString(specSheet.Cells(row + 1, 20))  \
                                                        + getCellValueinString(specSheet.Cells(row + 1, 24)) \
                                                        + getCellValueinString(specSheet.Cells(row + 1, 32))


                            if cmp_key == "UUID":
                                key_str = uuid_id
                            else:
                                key_str = spec_id0

                            if ope_flag == "CMP":
                                spec_str0 = spec_id0 + key_name0+ Key_Spec_str0 + Key_Spec_Condition_Str
                                mergeCntDiff = self.DiffWithExportSheet(spec_str0, specSheet, diffSheet, key_str, row, partlist)
                                row += mergeCntDiff
                                nrows += mergeCntDiff                            
                            else:
                                partlist[key_str] = spec_id0 + key_name0+ Key_Spec_str0 + Key_Spec_Condition_Str + "<row>" + str(row)

                    row += mergeCount

                para_row_list["para_4_row"] =  str(row)

                while not isEmptyValue(specSheet.Cells(row,3)):
                    if getCellValueinString(specSheet.Cells(row, 3)) == "Initialized Status":
                        break
                    comment_para_4 += getCellValueinString(specSheet.Cells(row, 3))
                    row += 1

                if ope_flag == "CMP":
                    spec_str0 = comment_para_4
                    self.DiffWithExportSheetForComment(spec_str0, specSheet, "comment_para_4", row, partlist)
                else:
                    partlist["comment_para_4"]  = comment_para_4 + "<row>" + str(row)

            # 5 - Initialized Status 
            if  specSheet.Cells(row, 3).Value == "Initialized Status":
                row += 1
                comment_para_5 = ""

                while not isEmptyValue(specSheet.Cells(row,2)):
                    mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                    uuid_id = getCellValueinString(specSheet.Cells(row, 47))
                    if uuid_id is not None: # is button
                        if uuid_id != "UUID" :
                            spec_id0 = getCellValueinString(specSheet.Cells(row, 2)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 3)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 4)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 5))

                            Spec_str0 = getCellValueinString(specSheet.Cells(row, 9)) \
                                    + getCellValueinString(specSheet.Cells(row, 12)) \
                                    + getCellValueinString(specSheet.Cells(row, 38)) \
                                    + getCellValueinString(specSheet.Cells(row, 41))

                            Spec_Condition_Str0 = ""
                            for i in range(0, mergeCount):
                                Spec_Condition_Str0 += getCellValueinString(specSheet.Cells(row + 1, 18)) \
                                                    + getCellValueinString(specSheet.Cells(row + 1, 24)) \
                                                    + getCellValueinString(specSheet.Cells(row + 1, 32))

                            if cmp_key == "UUID":
                                key_str = uuid_id
                            else:
                                key_str = spec_id0

                            if ope_flag == "CMP":
                                spec_str0 = spec_id0 + Spec_str0  + Spec_Condition_Str0
                                mergeCntDiff = self.DiffWithExportSheet(spec_str0, specSheet, diffSheet, key_str,row, partlist)
                                row += mergeCntDiff
                                nrows += mergeCntDiff                            
                            else:
                                partlist[key_str] = spec_id0 + Spec_str0  + Spec_Condition_Str0 + "<row>" + str(row)

                    row += mergeCount

                para_row_list["para_5_row"] =  str(row)

                while not isEmptyValue(specSheet.Cells(row,3)):
                    if getCellValueinString(specSheet.Cells(row, 3)) == "Action on Status change":
                        break
                    comment_para_5 += getCellValueinString(specSheet.Cells(row, 3))
                    row += 1

                if ope_flag == "CMP":
                    spec_str0 = comment_para_5
                    self.DiffWithExportSheetForComment(spec_str0, specSheet, "comment_para_5", row, partlist)
                else:
                    partlist["comment_para_5"]  = comment_para_5 + "<row>" + str(row)

            # 6 - Action on Status change
            if  specSheet.Cells(row, 3).Value == "Action on Status change":
                row += 3
                comment_para_6 = ""

                while not isEmptyValue(specSheet.Cells(row,2)):
                    mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                    uuid_id = getCellValueinString(specSheet.Cells(row, 47))
                    if uuid_id is not None: # is button
                        if uuid_id != "UUID" :
                            spec_id0 = getCellValueinString(specSheet.Cells(row, 2)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 3)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 4)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 5))

                            Spec_str0 = getCellValueinString(specSheet.Cells(row, 6)) \
                                    + getCellValueinString(specSheet.Cells(row, 8)) \
                                    + getCellValueinString(specSheet.Cells(row, 9)) \
                                    + getCellValueinString(specSheet.Cells(row, 20)) \
                                    + getCellValueinString(specSheet.Cells(row, 21)) \
                                    + getCellValueinString(specSheet.Cells(row, 32)) \
                                    + getCellValueinString(specSheet.Cells(row, 33)) 

                            Spec_Condition_Str0 = ""
                            for i in range(0, mergeCount):
                                Spec_Condition_Str0 += getCellValueinString(specSheet.Cells(row + 1, 13)) \
                                                    + getCellValueinString(specSheet.Cells(row + 1, 16)) \
                                                    + getCellValueinString(specSheet.Cells(row + 1, 25)) \
                                                    + getCellValueinString(specSheet.Cells(row + 1, 28)) \
                                                    + getCellValueinString(specSheet.Cells(row + 1, 38)) \
                                                    + getCellValueinString(specSheet.Cells(row + 1, 41)) \

                            if cmp_key == "UUID":
                                key_str = uuid_id
                            else:
                                key_str = spec_id0

                            if ope_flag == "CMP":
                                spec_str0 = spec_id0 + Spec_str0 + Spec_Condition_Str0
                                mergeCntDiff = self.DiffWithExportSheet(spec_str0, specSheet,diffSheet,  key_str, row, partlist)
                                row += mergeCntDiff
                                nrows += mergeCntDiff                            
                            else:

                                partlist[key_str] = spec_id0 + Spec_str0 + Spec_Condition_Str0 + "<row>" + str(row)

                    row += mergeCount

                para_row_list["para_6_row"] =  str(row)

                while not isEmptyValue(specSheet.Cells(row,3)):
                    if getCellValueinString(specSheet.Cells(row, 3)) == "Action on Transition":
                        break
                    comment_para_6 += getCellValueinString(specSheet.Cells(row, 3))
                    row += 1

                if ope_flag == "CMP":
                    spec_str0 = comment_para_6
                    self.DiffWithExportSheetForComment(spec_str0, specSheet, "comment_para_6", row, partlist)
                else:
                    partlist["comment_para_6"]  = comment_para_6 + "<row>" + str(row)

            # 7 - Action on Transition 
            if  specSheet.Cells(row, 3).Value == "Action on Transition":
                row += 3
                comment_para_7 = ""

                while not isEmptyValue(specSheet.Cells(row,2)):
                    mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                    uuid_id = getCellValueinString(specSheet.Cells(row, 47))
                    if uuid_id is not None: # is button
                        if uuid_id != "UUID" :
                            spec_id0 = getCellValueinString(specSheet.Cells(row, 2)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 3)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 4)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 5))

                            Spec_str0 = getCellValueinString(specSheet.Cells(row, 6)) \
                                    + getCellValueinString(specSheet.Cells(row, 9)) \
                                    + getCellValueinString(specSheet.Cells(row, 20)) \
                                    + getCellValueinString(specSheet.Cells(row, 21)) \
                                    + getCellValueinString(specSheet.Cells(row, 32)) \
                                    + getCellValueinString(specSheet.Cells(row, 33)) \

                            Spec_Condition_Str0 = ""
                            for i in range(0, mergeCount):
                                Spec_Condition_Str0 += getCellValueinString(specSheet.Cells(row + 1, 13)) \
                                                    + getCellValueinString(specSheet.Cells(row + 1, 16)) \
                                                    + getCellValueinString(specSheet.Cells(row + 1, 25)) \
                                                    + getCellValueinString(specSheet.Cells(row + 1, 28)) \
                                                    + getCellValueinString(specSheet.Cells(row + 1, 38)) \
                                                    + getCellValueinString(specSheet.Cells(row + 1, 41)) \


                            if cmp_key == "UUID":
                                key_str = uuid_id
                            else:
                                key_str = spec_id0

                            if ope_flag == "CMP":
                                spec_str0 = spec_id0 + Spec_str0 +  Spec_Condition_Str0
                                mergeCntDiff = self.DiffWithExportSheet(spec_str0, specSheet,diffSheet, key_str, row, partlist)
                                row += mergeCntDiff
                                nrows += mergeCntDiff                            
                            else:
                                partlist[key_str] = spec_id0 + Spec_str0 +  Spec_Condition_Str0 + "<row>" + str(row)

                    row += mergeCount
                para_row_list["para_7_row"] =  str(row)

                while not isEmptyValue(specSheet.Cells(row,3)):
                    if getCellValueinString(specSheet.Cells(row, 3)) == "Trigger Action":
                        break
                    comment_para_7 += getCellValueinString(specSheet.Cells(row, 3))
                    row += 1

                if ope_flag == "CMP":
                    spec_str0 = comment_para_7
                    self.DiffWithExportSheetForComment(spec_str0, specSheet, "comment_para_7", row, partlist)
                else:
                    partlist["comment_para_7"]  = comment_para_7 + "<row>" + str(row)

            # 8 - Trigger Action 
            if  specSheet.Cells(row, 3).Value == "Trigger Action":
                row += 2
                comment_para_8 = ""

                while not isEmptyValue(specSheet.Cells(row,2)):
                    mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                    uuid_id = getCellValueinString(specSheet.Cells(row, 47))
                    if uuid_id is not None: # is button
                        if uuid_id != "UUID" :
                            spec_id0 = getCellValueinString(specSheet.Cells(row, 2)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 3)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 4)) \
                                    + "_" + getCellValueinString(specSheet.Cells(row, 5))

                            Spec_str0 = getCellValueinString(specSheet.Cells(row, 6)) \
                                    + getCellValueinString(specSheet.Cells(row, 9)) \
                                    + getCellValueinString(specSheet.Cells(row, 12)) \
                                    + getCellValueinString(specSheet.Cells(row, 35)) \
                                    + getCellValueinString(specSheet.Cells(row, 38)) \
                                    + getCellValueinString(specSheet.Cells(row, 41)) 

                            Spec_Condition_Str0 = ""
                            for i in range(0, mergeCount):
                                Spec_Condition_Str0 += getCellValueinString(specSheet.Cells(row + 1, 17)) \
                                                    + getCellValueinString(specSheet.Cells(row + 1, 22)) \
                                                    + getCellValueinString(specSheet.Cells(row + 1, 32)) 

                            if cmp_key == "UUID":
                                key_str = uuid_id
                            else:
                                key_str = spec_id0

                            if ope_flag == "CMP":
                                spec_str0 = spec_id0 + Spec_str0 + Spec_Condition_Str0
                                mergeCntDiff = self.DiffWithExportSheet(spec_str0, specSheet, diffSheet, key_str, row, partlist)
                                row += mergeCntDiff
                                nrows += mergeCntDiff                            
                            else:
                                    
                                partlist[key_str] = spec_id0 + Spec_str0 + Spec_Condition_Str0 + "<row>" + str(row)

                    row += mergeCount
                para_row_list["para_8_row"] =  str(row)
                while not isEmptyValue(specSheet.Cells(row,3)):
                    if getCellValueinString(specSheet.Cells(row, 3)) == "":
                        break
                    comment_para_8 += getCellValueinString(specSheet.Cells(row, 3))
                    row += 1

                if ope_flag == "CMP":
                    spec_str0 = comment_para_8
                    self.DiffWithExportSheetForComment(spec_str0, specSheet, "comment_para_8", row, partlist)
                else:
                    partlist["comment_para_8"]  = comment_para_8 + "<row>" + str(row)

                break

            i = 1
            end_flag = False
            while(isEmptyValue(specSheet.Cells(row + i, 3))):
                i += 1
                if i > 20:
                    end_flag = True
                    break

            if end_flag:
                break

            row += 1

        return row

    def DiffWithExportSheetForComment(self, spec_str0, specSheet, comment_id, row, partlist):
        problem_cnt = int(partlist["problem_cnt"])
        try:
            spec_str1 = partlist[comment_id].split("<row>")[0]
            if spec_str0 != spec_str1: #diff
                #print("find diff in this spec")
                #print(spec_str0)
                #print(spec_str1)
                specSheet.Cells(row-1, 1).Value = "diff"+ getCellValueinString(specSheet.Cells(row-1, 1))
                specSheet.Cells(row-1, 1).interior.color =  rgb_to_hex(DIFF_COLOR)
                range_a = "C" + str(row) + ":BM" + str(row)
                range_b = "A" + str(row) + ":BM" + str(row)

                specSheet.Range(range_a).EntireRow.Insert()
                specSheet.Cells(row, 3).Value = spec_str1
                specSheet.Range(range_a).Merge()
                specSheet.Range(range_b).interior.color = rgb_to_hex(EXPORT_COLOR)

                problem_cnt += 1

            else: #same
                specSheet.Cells(row-1, 1).Value = "same"+ getCellValueinString(specSheet.Cells(row-1, 1))
            
            del partlist[comment_id]

        except Exception as e:
            specSheet.Cells(row-1, 1).Value = "This only" + + getCellValueinString(specSheet.Cells(row-1, 1))
            specSheet.Cells(row, 1).interior.color = rgb_to_hex(CHECK_COLOR)
            problem_cnt += 1

        partlist["problem_cnt"] = str(problem_cnt)

        return 1

    def DiffWithExportSheet(self, spec_str0, specSheet, diffSheet, uuid_id,row, partlist):

        mergeCntDiff = 0
        problem_cnt = int(partlist["problem_cnt"])
        try:
            spec_str1 = partlist[uuid_id].split("<row>")[0]
            row_in_other = partlist[uuid_id].split("<row>")[1]
            mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
            if spec_str0 != spec_str1: #diff
                #print("find diff in this spec")
                #print(spec_str0)
                #print(spec_str1)
                specSheet.Cells(row, 1).Value = "diff" + getCellValueinString(specSheet.Cells(row, 1))
                mergeCntDiff = diffSheet.Cells(row_in_other, 2).MergeArea.Rows.Count
                range_a = "A" + str(row) + ":A" + str(row + mergeCount -1)
                specSheet.Range(range_a).interior.color = rgb_to_hex(DIFF_COLOR)

                for merC in range(0,mergeCntDiff):
                    specSheet.Range("A" + str(row + mergeCount) + ":BM" + str(row + mergeCount)).EntireRow.Insert()
                range_a = "A" + str(row + mergeCount) + ":BM" + str(row + mergeCount + mergeCntDiff -1)
                range_b = "A" + str(row_in_other) + ":BM" + str(int(row_in_other) + mergeCntDiff -1)
                
                specSheet.Range(range_a).UnMerge()
                diffSheet.Range(range_b).Copy(specSheet.Cells.Range(range_a))
                specSheet.Range(range_a).interior.color = rgb_to_hex(EXPORT_COLOR)
    
                problem_cnt += 1

            else: #same
                specSheet.Cells(row, 1).Value = "same" + getCellValueinString(specSheet.Cells(row, 1))
            
            del partlist[uuid_id]


        except Exception as e:
            #print(e)
            specSheet.Cells(row, 1).Value = "This only"+ getCellValueinString(specSheet.Cells(row, 1))
            specSheet.Cells(row, 1).interior.color = rgb_to_hex(CHECK_COLOR)
            problem_cnt += 1
    
        partlist["problem_cnt"] = str(problem_cnt)

        return mergeCntDiff

    def CompareFolder(self, fold0, fold1):
        dirobj = filecmp.dircmp( os.getcwd() + "/" +fold0,  os.getcwd() + "/" +fold1, ['.DS_Store','___exportspecs___'])
        self.resultBook.ActiveSheet.Name = "FileFolder"
        resultSheet = self.resultBook.ActiveSheet
        resultSheet.Range("A1, F1000").Font.name = "Microsoft YaHei UI"

        resultSheet.Cells(2,1).Value = "Only in " + fold0
        resultSheet.Cells(2,2).Value = "Only in " + fold1
        resultSheet.Cells(2, 1).interior.color = rgb_to_hex((231,230,255))
        resultSheet.Cells(2, 2).interior.color = rgb_to_hex((231,230,255))

        resultSheet.Columns("A:A").ColumnWidth = 50
        resultSheet.Columns("A:A").WrapText = True

        resultSheet.Columns("B:B").ColumnWidth = 50
        resultSheet.Columns("B:B").WrapText = True

        wRow=[3]
        self.GetResultFromDirObj(dirobj, '', resultSheet, wRow)

    def GetResultFromDirObj(self, dirobj, rootName, cursheet, wRow):
        if dirobj is None:
            return

        for samefile in dirobj.common_files:
            self.result_list.append(samefile)

        for dir0 in dirobj.left_only:
            cursheet.Cells(wRow[0], 1).Value = dir0
            wRow[0] += 1

        for dir1 in dirobj.right_only:
            cursheet.Cells(wRow[0], 2).Value = dir1
            wRow[0] += 1

        '''
        for file0 in dirobj.diff_files:
            cursheet.Cells(wRow[0], 1).Value = file0
            wRow[0] += 1
        '''
        
        for dir2 in dirobj.common_dirs:
            cursheet.Cells(wRow[0], 1).Value = dir2
            cursheet.Cells(wRow[0], 1).interior.color = rgb_to_hex((147,205,221))
            cursheet.Cells(wRow[0], 2).interior.color = rgb_to_hex((147,205,221))
            wRow[0] += 1
            self.GetResultFromDirObj(dirobj.subdirs[dir2], rootName + "/" + dir2, cursheet, wRow)

    def SaveResultFile(self, resultName):
        resultFileName  = os.getcwd() + "\\" + resultName
        if os.path.exists(resultFileName): 
            os.remove(resultFileName)

        self.resultBook.SaveAs(resultFileName)
        #self.resultBook.Close(Savechanges = 1)

    def AbstractHistoryUpdate(self, root0, root1):
        compare_list={}
        self.GetCompareFileList(root0, root1, compare_list)

        for fileName, compilefile in compare_list.items():            
            if os.path.isfile(fileName):
                try:
                    print('Research this ->',fileName)    # 绝对路径
                    file1 = os.getcwd() + "\\" + fileName
                    specWb0  = self.excelApp.Workbooks.Open(file1)
                    specSheet0 = specWb0.Sheets("History")

                    file2 =  os.getcwd() + "\\" + compare_list[fileName]
                    specWb1  = self.excelApp.Workbooks.Open(file2)       
                    specSheet1 = specWb1.Sheets("History")

                    info = specSheet1.UsedRange
                    nrows = info.Rows.Count
                    ncolumns = info.Columns.Count
                    row = 1
                    bSave = False
                    while (row <= nrows):
                        history_item0 = getCellValueinString(specSheet0.Cells(row, 2))
                        col = 3
                        while (col <= 5):
                            history_item0 += getCellValueinString(specSheet0.Cells(row, col))
                            col += 1

                        history_item1 = getCellValueinString(specSheet1.Cells(row, 2))
                        col = 3
                        while (col <= 5):
                            history_item1 += getCellValueinString(specSheet1.Cells(row, col))
                            col += 1

                        if (history_item0 != history_item1):
                            range_a = "A" + str(row) + ":Z" + str(row)
                            range_b = "A" + str(row) + ":Z" + str(row)
                            specSheet0.Range(range_a).EntireRow.Insert()
                            specSheet1.Range(range_a).Copy(specSheet0.Cells.Range(range_b))
                            specSheet0.Cells(row, 2).interior.color = rgb_to_hex(AMEND_COLOR)
                            bSave = True
                            print(history_item0) 
                            print(history_item1)                         
                        row += 1

                    if bSave:        
                        specWb0.Close(SaveChanges = 1)
                    else:
                        specWb0.Close(SaveChanges = 0)

                    specWb1.Close(SaveChanges = 0)
                    
                except  Exception as e:                     
                    print(e)
                    print('!!! Error in ->', fileName)    # 绝对路径
                    copyspecFile(fileName, "error_files")

    def AddHistoryComment(self, root):

        for root, dirs, files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root,name)

                if os.path.isfile(fileName):
                    try:
                        print('Research this ->',fileName)    # 绝对路径
                        file1 = os.getcwd() + "\\" + fileName
                        specWb  = self.excelApp.Workbooks.Open(file1)
                        specSheet = specWb.Sheets("History")

                        info = specSheet.UsedRange
                        nrows = info.Rows.Count
                        ncolumns = info.Columns.Count

                        row = 5
                        while (not isEmptyValue(specSheet.Cells(row, 2))):
                            if getCellValueinString(specSheet.Cells(row, 3)).upper().find("V") >= 0 :
                                specSheet.Cells(row, 3).Value = specSheet.Cells(row, 3).replace("V","")
                            row += 1

                        
                        lastVersionNo = int(specSheet.Cells(row-1, 2).Value)

                        range_a = "A" + str(row) + ":Z" + str(row)
                        range_b = "A" + str(row-1) + ":Z" + str(row-1)
                        specSheet.Range(range_a).EntireRow.Insert()
                        specSheet.Range(range_b).Copy(specSheet.Cells.Range(range_a))
                        specSheet.Cells(row, 2).Value = lastVersionNo + 1
                        specSheet.Cells(row, 3).Value = "1.14"
                        specSheet.Cells(row, 4).Value = "ALL"
                        specSheet.Cells(row, 5).Value = "iAuto"
                        specSheet.Cells(row, 6).Value = "2020/12/11"
                        specSheet.Cells(row, 7).Value = "-"
                        specSheet.Cells(row, 8).Value = "-"
                        specSheet.Cells(row, 9).Value = "iAuto 内部作业"
                        specSheet.Cells(row, 10).Value = "Descrip. Change"
                        specSheet.Cells(row, 11).Value = "Ope import to Sketch"
                        specSheet.Cells(row, 12).Value = "ALL"
                        specSheet.Cells(row, 13).Value = "Not synchronized to sketch"
                        specSheet.Cells(row, 14).Value = "Update Sketch exported info"
                        specSheet.Cells(row, 15).Value = "2020/12/11"
                        specSheet.Cells(row, 16).Value = "SL Group"
                        specSheet.Cells(row, 17).Value = "-"
                    
                        
                        specWb.Close(SaveChanges = 1)
                        
                    except  Exception as e:                     
                        print(e)
                        print('!!! Error in ->', fileName)    # 绝对路径
                        copyspecFile(fileName, "error_files")

    def AddComment (self, root):

        rptfile = os.getcwd() + "\\doc\\report_round_final.xlsx"
        rptWb  = self.excelApp.Workbooks.Open(rptfile)
        rptSheet = rptWb.Sheets("issues")

        rptinfo = rptSheet.UsedRange
        nRptRrows = rptinfo.Rows.Count
        ncolumns = rptinfo.Columns.Count
        rptRow = 2
        
        screen_list = {}

        while rptRow <= nRptRrows:
            screen_list[getCellValueinString(rptSheet.Cells(rptRow, 6))] = str(rptRow)
            rptRow += 1


        wRow = 1
        for root, dirs, files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root,name)
                try:
                    print('Search in this ->',fileName)    # 绝对路径
                    file = os.getcwd() + "\\" + fileName
                    specWb  = self.excelApp.Workbooks.Open(file)
                    sheetCount = (specWb.Worksheets.Count)
                    wRow += 1                   
                    catalogSheet = specWb.Sheets('Catalog')

                    copytoRow  = 100
                    for i in range(1, sheetCount):
                        specSheet = specWb.Worksheets(i)
                        if specSheet.Name not in ("Catalog", "History","Cover","DropDownDataList","propertyDownDataListCCtype"):
                            info = specSheet.UsedRange
                            nrows = info.Rows.Count
                            ncolumns = info.Columns.Count
                            row = 1
                            print("Starting to process " + specSheet.Name)


                            # Find in first version
                            while row < nrows:
            
                                # 0 - Outline
                                if (getCellValueinString(specSheet.Cells(row, 3)) == "Outline"):                                   
                                    row += 1

                                # 1 View of Screen
                                bHasScroll = False
                                if getCellValueinString(specSheet.Cells(row, 3)) is not None:
                                    if  specSheet.Cells(row, 3).Value == "View of Screen":
                                        while (specSheet.Cells(row, 5).Value != "Parts Name") :
                                            row += 1

                                        while not isEmptyValue(specSheet.Cells(row,2)):
                                            mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                                            partName =  getCellValueinString(specSheet.Cells(row, 6)) 
                                            if partName.upper().find("SCROLLBAR") >= 0:
                                                bHasScroll = True
                                            row += mergeCount

                                        bNoListcomment = True
                                        comment_cnt = 0
                                        while not isEmptyValue(specSheet.Cells(row,3)):
                                            if getCellValueinString(specSheet.Cells(row, 3)) == "View of Soft Button":
                                                if bHasScroll and bNoListcomment:
                                                    range_a = "A" + str(row) + ":AR" + str(row)
                                                    range_b = "A" + str(row-1) + ":AR" + str(row-1)
                                                    specSheet.Range(range_a).EntireRow.Insert()
                                                    specSheet.Range(range_b).Copy(specSheet.Cells.Range(range_a))
                                                    specSheet.Cells(row, 3).Value = "(#)1-" + str(comment_cnt + 1) + " :Refer to 21MM_Material_05_ListCommonProcess.xlsx MM_05_00_05 1-3-P1"
                                                    specSheet.Cells(row, 1).Value = "V1.14"
                                                    specSheet.Range(range_a).font.color = rgb_to_hex((43,15,249))
                                                    specSheet.Range(range_a).font.color = rgb_to_hex((43,15,249))

                                                    nrows += 1
                                                    rptSheet.Cells(int(screen_list[specSheet.Name]), 3).Value = specSheet.Cells(row, 3).Value
                                                    print("Add comment ", specSheet.Name)
                                                break
                                            comment_str = getCellValueinString(specSheet.Cells(row, 3))
                                            if comment_str.upper().find("MATERIAL_05_LISTCOMMONPROCESS") >= 0:
                                                bNoListcomment = False
                                    
                                            row += 1
                                            comment_cnt += 1

                                # 2 View of Soft Button
                                if  specSheet.Cells(row, 3).Value == "View of Soft Button":
                                    while (specSheet.Cells(row, 5).Value != "Button Name") :
                                        row += 1

                                    while not isEmptyValue(specSheet.Cells(row,2)):
                                        mergeCount = specSheet.Cells(row, 2).MergeArea.Rows.Count
                                        row += mergeCount

                                    comment_cnt = 0
                                    bNoListcomment = True
                                    while not isEmptyValue(specSheet.Cells(row,3)):
                                        if getCellValueinString(specSheet.Cells(row, 3)) == "Soft Button Action":
                                            if bHasScroll and bNoListcomment:
                                                range_a = "A" + str(row) + ":AR" + str(row)
                                                range_b = "A" + str(row-1) + ":AR" + str(row-1)
                                                specSheet.Range(range_a).EntireRow.Insert()
                                                specSheet.Range(range_b).Copy(specSheet.Cells.Range(range_a))
                                                specSheet.Cells(row, 3).Value = "(#)2-" + str(comment_cnt + 1) + " :Refer to 21MM_Material_05_ListCommonProcess.xlsx"
                                                specSheet.Cells(row, 1).Value = "V1.14"
                                                specSheet.Range(range_a).font.color = rgb_to_hex((43,15,249))
                                                specSheet.Range(range_a).font.color = rgb_to_hex((43,15,249))

                                                nrows += 1
                                                if isEmptyValue(rptSheet.Cells(int(screen_list[specSheet.Name]), 3)):
                                                    rptSheet.Cells(int(screen_list[specSheet.Name]), 3).Value = specSheet.Cells(row, 3).Value
                                                else:
                                                    rptSheet.Cells(int(screen_list[specSheet.Name]), 3).Value += specSheet.Cells(row, 3).Value

                                                print("Add comment ", specSheet.Name)
                                            break
                                        comment_str = getCellValueinString(specSheet.Cells(row, 3))
                                        if comment_str.upper().find("MATERIAL_05_LISTCOMMONPROCESS") >= 0:
                                            bNoListcomment = False
                                        row += 1
                                        comment_cnt += 1

                                    break

                                row += 1

                    specWb.Close(SaveChanges = 1)

                except Exception as e:
                    print(e)
                    print('!!! Error in ->', fileName)    # 绝对路径
                    copyspecFile(fileName, "error_files")

        print('...........................................')

    def MergeRFQSheet_uranus(self, root):
        CostTotalfile = os.getcwd() + "\\doc\\CostFiles_uranus.xlsx"
        CostTotalWb  = self.excelApp.Workbooks.Open(CostTotalfile)
        CostSheet =  CostTotalWb.Sheets('list') 


        st_screen_list= {}
        st_redmine_list = {}
        row = 2
        while(not isEmptyValue(CostSheet.Cells(row, 4))):
            issue_id = getCellValueinString(CostSheet.Cells(row, 4))
            redmine_id = getCellValueinString(CostSheet.Cells(row, 1))
            st_screen_list[issue_id] = str(row)
            st_redmine_list[redmine_id] = str(row)
            
            row += 1        

        '''
        for root, dirs,files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root, name)
                try:
                    print('Search in this ->',fileName)    # 绝对路径
                    file = os.getcwd() + "\\" + fileName
                    costWb  = self.excelApp.Workbooks.Open(file)
                    costsingleSheet = costWb.Sheets('工数预估') 
                    info = costsingleSheet.UsedRange
                    nrows = info.Rows.Count

                    folder_name = root[5:20]
                    rowNo = 11
                    while rowNo <= nrows:
                        if isEmptyValue(costsingleSheet.Cells(rowNo,2)):
                            break
                        
                        if getCellValueinString(costsingleSheet.Cells(rowNo, 90)) == "SL":
                            costRow = int(st_screen_list[folder_name])
                            CostSheet.Cells(costRow, 11).Value = costsingleSheet.Cells(rowNo, 20).Value/20
                            print(folder_name)
                        rowNo += 1
                    

                    costWb.Close(SaveChanges = 0)

                except Exception as e:
                    print(e)
                    #costWb.Close(SaveChanges = 0)
                    copyspecFile(fileName, "error_files")
        
        '''
        redminecostfile = os.getcwd() + "\\doc\\CostFiles01.xlsx"
        redmineCostWb  = self.excelApp.Workbooks.Open(redminecostfile)
        redmineSheet =  redmineCostWb.Sheets('Summary') 
        info = redmineSheet.UsedRange
        nrows = info.Rows.Count
        rowNo = 3
        costRow = 0
        while rowNo <= nrows:

            if isEmptyValue(redmineSheet.Cells(rowNo,3)):
                break

            redmine_id = getCellValueinString(redmineSheet.Cells(rowNo, 1))
            
            try:
                costRow = int(st_redmine_list[redmine_id])
            except Exception as e:
                costRow = 80 + rowNo
                CostSheet.Cells(costRow, 1).Value = redmineSheet.Cells(rowNo, 1).Value

            value = 0
            if not isEmptyValue(CostSheet.Cells(costRow, 11)):
                value = CostSheet.Cells(costRow, 11).Value
            
            CostSheet.Cells(costRow, 11).Value = value + float(redmineSheet.Cells(rowNo, 15).Value)

            rowNo += 1

        
        CostTotalWb.Close(SaveChanges = 1)
        

    def MergeRFQSheet(self, root):
        CostTotalfile = os.getcwd() + "\\doc\\CostFiles.xlsx"
        CostTotalWb  = self.excelApp.Workbooks.Open(CostTotalfile)
        CostSampleSheet =  CostTotalWb.Sheets('sample') 
        CostSummarySheet = CostTotalWb.Sheets('Summary') 

        group_taglist = []
        summaryCnt = 3
        for root, dirs,files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root, name)
                try:
                    print('Search in this ->',fileName)    # 绝对路径
                    file = os.getcwd() + "\\" + fileName
                    costWb  = self.excelApp.Workbooks.Open(file)
                    costSheet = costWb.Sheets('DATA') 
                    info = costSheet.UsedRange
                    nrows = info.Rows.Count


                    costTotalSheet = CostTotalWb.Worksheets.Add()

                    filter_filename = name.upper().replace(".XLSX", "")[36:].replace("21MM MID AVN","").replace("中国専用仕様書","").replace("VER","")
                    filter_filename = filter_filename.replace("【","").replace("】","").replace("-","_").replace("中国","")

                    #print(filter_filename)
                    costTotalSheet.Name = filter_filename

                    CostSampleSheet.Range("A1:AT1").Copy(costTotalSheet.Cells.Range("A1:AT1"))
                    rowNo  = 3
                    summary_SL = 0
                    summary_Assess = 0
                    summary_HMI = 0
                    summary_Vehicle = 0
                    summary_Connectivity = 0
                    summary_BCCC = 0
                    summary_Media = 0
                    summary_FWSY = 0
                    summary_Voice = 0
                    summary_Telema = 0
                    summary_PF = 0
                    summary_FW = 0
                    summary_UIFW = 0
                    summary_Update= 0

                    sm_row = 2
                    while rowNo <= nrows:
                        if isEmptyValue(costSheet.Cells(rowNo, 23)):
                            break

                        isGoodValue = (not isEmptyValue(costSheet.Cells(rowNo, 59))) and getCellValueinString(costSheet.Cells(rowNo, 59)) not in ("0", "0.0") 
                        hasRefValue = not isEmptyValue(costSheet.Cells(rowNo, 91)) 
                        isNotGoodValue = (getCellValueinString(costSheet.Cells(rowNo, 59)) in ("0","0.0")) or isEmptyValue(costSheet.Cells(rowNo, 59))
                        isEmptyComment = isEmptyValue(costSheet.Cells(rowNo, 64)) and isEmptyValue(costSheet.Cells(rowNo, 65))
                        isEmptySign = isEmptyValue(costSheet.Cells(rowNo, 89))
                        group_tag = getCellValueinString(costSheet.Cells(rowNo, 24))
                        if group_tag in ("SL"):
                            if  isGoodValue or hasRefValue or (isNotGoodValue and isEmptySign and isEmptyComment) :

                                range_a = "H" + str(rowNo) + ":H" + str(rowNo)
                                range_b = "A" + str(sm_row) + ":A" + str(sm_row)
                                costSheet.Range(range_a).Copy(costTotalSheet.Cells.Range(range_b))
                    
                                range_a = "K" + str(rowNo) + ":AB" + str(rowNo)
                                range_b = "B" + str(sm_row) + ":S" + str(sm_row)
                                costSheet.Range(range_a).Copy(costTotalSheet.Cells.Range(range_b))

                                range_a = "AM" + str(rowNo) + ":AV" + str(rowNo)
                                range_b = "T" + str(sm_row) + ":AC" + str(sm_row)
                                costSheet.Range(range_a).Copy(costTotalSheet.Cells.Range(range_b))

                                range_a = "BF" + str(rowNo) + ":BG" + str(rowNo)
                                range_b = "AD" + str(sm_row) + ":AE" + str(sm_row)
                                costSheet.Range(range_a).Copy(costTotalSheet.Cells.Range(range_b))

                                range_a = "BJ" + str(rowNo) + ":BQ" + str(rowNo)
                                range_b = "AF" + str(sm_row) + ":AM" + str(sm_row)
                                costSheet.Range(range_a).Copy(costTotalSheet.Cells.Range(range_b))

                                range_a = "CN" + str(rowNo) + ":CO" + str(rowNo)
                                range_b = "AN" + str(sm_row) + ":AO" + str(sm_row)
                                costSheet.Range(range_a).Copy(costTotalSheet.Cells.Range(range_b))

                                range_a = "CP" + str(rowNo) + ":CP" + str(rowNo)
                                range_b = "AP" + str(sm_row) + ":AP" + str(sm_row)
                                costSheet.Range(range_a).Copy(costTotalSheet.Cells.Range(range_b))

                                range_a = "CT" + str(rowNo) + ":CU" + str(rowNo)
                                range_b = "AQ" + str(sm_row) + ":AR" + str(sm_row)
                                costSheet.Range(range_a).Copy(costTotalSheet.Cells.Range(range_b))

                                range_a = "AC" + str(rowNo) + ":AE" + str(rowNo)
                                range_b = "AS" + str(sm_row) + ":AT" + str(sm_row)
                                costSheet.Range(range_a).Copy(costTotalSheet.Cells.Range(range_b))

                                
                                #print(group_tag)
                                if group_tag not in group_taglist:  
                                    group_taglist.append(group_tag)
                                if CommonFunc.isContentValid(costSheet.Cells(rowNo, 59)):
                                    cost_value = 0
                                    try:
                                        cost_value = float(costSheet.Cells(rowNo, 59).Value)
                                    except Exception as e:
                                        cost_value = 0

                                    if group_tag in ("评价"):
                                        summary_Assess += cost_value                        
                                    elif group_tag in ("SL"):
                                        summary_SL += cost_value
                                    elif group_tag in ("HMI"):
                                        summary_HMI += cost_value
                                    elif group_tag in ("Vehicle"):
                                        summary_Vehicle += cost_value
                                    elif group_tag in ("Connectivity"):
                                        summary_Connectivity += cost_value
                                    elif group_tag in ("BCCC"):
                                        summary_BCCC += cost_value
                                    elif group_tag in ("Media"):
                                        summary_Media += cost_value
                                    elif group_tag in ("FW-SY"):
                                        summary_FWSY += cost_value
                                    elif group_tag in ("Voice"):
                                        summary_Voice += cost_value
                                    elif group_tag in ("Telema"):
                                        summary_Telema += cost_value
                                    elif group_tag in ("PF"):
                                        summary_PF += cost_value
                                    elif group_tag in ("FW"):
                                        summary_FW += cost_value
                                    elif group_tag in ("UIFW"):
                                        summary_UIFW += cost_value
                                    elif group_tag in ("update"):
                                        summary_Update += cost_value
                                
                                sm_row += 1

                        rowNo += 1
                                
                    costTotalSheet.Range("AM2:AN" + str(nrows)).RowHeight = 40
                    
                    summary_col = 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = getCellValueinString(costSheet.Cells(3, 6))
                    summary_col += 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = getCellValueinString(costSheet.Cells(3, 17))
                    summary_col += 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = summary_HMI
                    summary_col += 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = summary_Vehicle
                    summary_col += 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = summary_Connectivity
                    summary_col += 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = summary_BCCC
                    summary_col += 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = summary_Media
                    summary_col += 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = summary_FWSY
                    summary_col += 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = summary_Voice
                    summary_col += 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = summary_Telema
                    summary_col += 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = summary_PF
                    summary_col += 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = summary_FW
                    summary_col += 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = summary_UIFW
                    summary_col += 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = summary_Update
                    summary_col += 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = summary_SL
                    summary_col += 1
                    CostSummarySheet.Cells(summaryCnt, summary_col).Value = summary_Assess

                    summaryCnt +=1

                    costWb.Close(SaveChanges = 0)
                except Exception as e:
                    print(e)
                    copyspecFile(fileName, "error_files")
        
        print(group_taglist)

        CostTotalWb.Close(SaveChanges = 1)

    def MergeNoCostSheet(self, root):
        CostTotalfile = os.getcwd() + "\\doc\\未报价内容整理.xlsx"
        CostTotalWb  = self.excelApp.Workbooks.Open(CostTotalfile)
        costTotalSheet = CostTotalWb.Sheets('Summary') 

        sm_row = 3
        for root, dirs,files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root, name)
                try:
                    print('Search in this ->',fileName)    # 绝对路径
                    file = os.getcwd() + "\\" + fileName
                    costWb  = self.excelApp.Workbooks.Open(file)
                    costSheet = costWb.Sheets('DATA') 
                    info = costSheet.UsedRange
                    nrows = info.Rows.Count

                    rowNo  = 3

                    while rowNo <= nrows:
                        if isEmptyValue(costSheet.Cells(rowNo, 23)):
                            break
                        
                        noCost = getCellValueinString(costSheet.Cells(rowNo, 5))
                        groupTag = getCellValueinString(costSheet.Cells(rowNo, 24))
                        if (noCost.find("无外部") >= 0) and (groupTag.upper() == "SL"):
                            range_a = "F" + str(rowNo) + ":H" + str(rowNo)
                            range_b = "A" + str(sm_row) + ":C" + str(sm_row)
                            costSheet.Range(range_a).Copy(costTotalSheet.Cells.Range(range_b))
                
                            range_a = "K" + str(rowNo) + ":V" + str(rowNo)
                            range_b = "D" + str(sm_row) + ":O" + str(sm_row)
                            costSheet.Range(range_a).Copy(costTotalSheet.Cells.Range(range_b))

                            range_a = "AA" + str(rowNo) + ":AB" + str(rowNo)
                            range_b = "P" + str(sm_row) + ":Q" + str(sm_row)
                            costSheet.Range(range_a).Copy(costTotalSheet.Cells.Range(range_b))

                            range_a = "BJ" + str(rowNo) + ":BK" + str(rowNo)
                            range_b = "R" + str(sm_row) + ":S" + str(sm_row)
                            costSheet.Range(range_a).Copy(costTotalSheet.Cells.Range(range_b))
                            sm_row += 1

                        rowNo += 1
                                

                    costWb.Close(SaveChanges = 0)
                except Exception as e:
                    print(e)
        
        CostTotalWb.Close(SaveChanges = 1)

    def MergeSpecID(self, root):
        stringTablefile = os.getcwd() + "\\doc\\21MM_Material_09_StringTable_research.xlsx"
        stringTableWb  = self.excelApp.Workbooks.Open(stringTablefile)
        stringTableSheet = stringTableWb.Sheets('All Words') 
        unknownStringTableSheet = stringTableWb.Sheets('UnKnown') 

        st_screen_list= {}
        row = 3
        unknowRow =  3
        while(not isEmptyValue(stringTableSheet.Cells(row, 1))):
            screen_id = getCellValueinString(stringTableSheet.Cells(row, 5))
            original_str = getCellValueinString(stringTableSheet.Cells(row, 8))
            status_str = getCellValueinString(stringTableSheet.Cells(row, 25))
            combine_str = screen_id + "|" + original_str
            if status_str.upper() != "NOUSE":
                if combine_str in st_screen_list:
                    st_screen_list[combine_str] = st_screen_list[combine_str] + "|"+ str(row)
                else:
                    st_screen_list[combine_str] = str(row)
            
            row += 1
        

        for root, dirs,files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root, name)
                try:

                    file = os.getcwd() + "\\" + fileName
                    specWb  = self.excelApp.Workbooks.Open(file)
                    sheetCount = (specWb.Worksheets.Count)
                    for i in range(1, sheetCount + 1):
                        specSheet = specWb.Worksheets(i)
                        sheet_name = specWb.Worksheets(i).Name
                        
                        if sheet_name not in ("Cover", "Catalog","History","DropDownDataList","propertyDownDataListCCtype"):
                            print('Search in this ->',sheet_name)    # 绝对路径                            
    
                            info = specSheet.UsedRange
                            nrows = info.Rows.Count

                            rowNo  = 3
                            
                            screen_id = getCellValueinString(specSheet.Cells(2, 2))
                            specID_string_list = {}
                            while rowNo <= nrows:
                                mergeCount = specSheet.Cells(rowNo, 2).MergeArea.Rows.Count

                                disp_content = getCellValueinString(specSheet.Cells(rowNo, 26))
                                if disp_content.find("\"") >= 0:
                                    disp_content = disp_content[1:]
                                    disp_content = disp_content[0:len(disp_content)-1]
                                    spec_id = getCellValueinString(specSheet.Cells(rowNo, 2)) \
                                            + "_" + getCellValueinString(specSheet.Cells(rowNo, 3)) \
                                            + "_" + getCellValueinString(specSheet.Cells(rowNo, 4)) \
                                            + "_" + getCellValueinString(specSheet.Cells(rowNo, 5))

                                    combine_str1 = screen_id + "|" + disp_content
                                    if combine_str1 not in specID_string_list:
                                        specID_string_list[combine_str1] = spec_id
                                    else:
                                        specID_string_list[combine_str1] = specID_string_list[combine_str1] + "|" + spec_id

                                i = 1
                                end_flag = False
                                while(isEmptyValue(specSheet.Cells(rowNo + i, 3))):
                                    i += 1
                                    if i > 20:
                                        end_flag = True
                                        break

                                if end_flag:
                                    rowNo = nrows + 1

                                rowNo += mergeCount
                            
                            for key in specID_string_list:
                                try:
                                    spec_value = specID_string_list[key]
                                    row_info = st_screen_list[key]
                                    if row_info.find("|") >= 0 :
                                        row_list = row_info.split("|")
                                        for row_no in row_list:
                                            stringTableSheet.Cells(int(row_no), 6).Value = spec_value
                                            stringTableSheet.Cells(int(row_no), 6).interior.color = rgb_to_hex(AMEND_COLOR)
                                        print(row_info + " ->>>" + spec_value)
                                        #stringTableSheet.Cells(row, 6).Value = spec_value
                                    else:
                                        stringTableSheet.Cells(int(row_info), 6).Value = spec_value
                                        stringTableSheet.Cells(int(row_info), 6).font.color = rgb_to_hex((43,15,249))
                                        stringTableSheet.Cells(int(row_info), 6).font.color = rgb_to_hex((43,15,249))

                                except Exception as e:
                                    unknownStringTableSheet.Cells(unknowRow, 1).Value = key.split("|")[0]
                                    unknownStringTableSheet.Cells(unknowRow, 2).Value = specID_string_list[key]
                                    unknownStringTableSheet.Cells(unknowRow, 4).Value = key.split("|")[1]
                                    unknowRow += 1
                                    print(e)
                                    print(specID_string_list[key])

                        i += 1

                    specWb.Close(SaveChanges = 0)
                except Exception as e:
                    print(e)
        
        stringTableWb.Close(SaveChanges = 1)

    def CompareUIResult(self):
        UIStringTablefile = os.getcwd() + "\\doc\\All_Words_210524_check.xlsx"
        UIStringTableWb  = self.excelApp.Workbooks.Open(UIStringTablefile)
        UIStringTableSheet = UIStringTableWb.Sheets('All Words') 

        stringTablefile = os.getcwd() + "\\doc\\21MM_Material_09_StringTable.xlsx"
        stringTableWb  = self.excelApp.Workbooks.Open(stringTablefile)
        stringTableSheet = stringTableWb.Sheets('All Words') 

        info = stringTableSheet.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count

        string_table = {}
        for i in range(3, nrows + 1):
            chk_mark = getCellValueinString(stringTableSheet.Cells(i, 25))
            if chk_mark.upper() not in ("REPEAT","NOUSE"):
                string_id = getCellValueinString(stringTableSheet.Cells(i, 7))
                string_table[string_id] = str(i)
            i += 1

        uiInfo = UIStringTableSheet.UsedRange
        nUIrows = uiInfo.Rows.Count

        for i in range(3, nUIrows + 1):
            ui_stringid = getCellValueinString(UIStringTableSheet.Cells(i, 7))
            ui_st_en = getCellValueinString(UIStringTableSheet.Cells(i, 11))
            ui_st_zh = getCellValueinString(UIStringTableSheet.Cells(i, 14))
            st_status = getCellValueinString(UIStringTableSheet.Cells(i, 25))
            if st_status == "变更":
                try:
                    row_no = int(string_table[ui_stringid])
                    sl_en = getCellValueinString(stringTableSheet.Cells(row_no, 11))
                    sl_zh = getCellValueinString(stringTableSheet.Cells(row_no, 23))
                    UIStringTableSheet.Cells(i,13).Value = sl_en 
                    UIStringTableSheet.Cells(i,17).Value = sl_zh 
                    if ui_st_en != sl_en :
                        UIStringTableSheet.Cells(i, 18).Value = "en_diff"
                    else:
                        UIStringTableSheet.Cells(i, 18).Value = "en_same"

                    if ui_st_zh != sl_zh :
                        UIStringTableSheet.Cells(i, 18).Value = UIStringTableSheet.Cells(i, 18).Value + "|zh_diff"
                    else:
                        UIStringTableSheet.Cells(i, 18).Value = UIStringTableSheet.Cells(i, 18).Value + "|zh_same"

                except Exception as e:
                    print(e)

            i += 1

        UIStringTableWb.Close(SaveChanges = 1)
        stringTableWb.Close(SaveChanges = 0)

    def CheckUIResult(self):
        UIStringTablefile = os.getcwd() + "\\doc\\All_Words0601.xlsx"
        UIStringTableWb  = self.excelApp.Workbooks.Open(UIStringTablefile)
        UIStringTableSheet = UIStringTableWb.Sheets('All Words') 

        stringTablefile = os.getcwd() + "\\doc\\21MM_Material_09_StringTable.xlsx"
        stringTableWb  = self.excelApp.Workbooks.Open(stringTablefile)
        stringTableSheet = stringTableWb.Sheets('All Words') 
        historySheet = stringTableWb.Sheets('History') 

        info = UIStringTableSheet.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count

        ui_string_table = {}
        for i in range(3, nrows + 1):
            ui_screenid = getCellValueinString(UIStringTableSheet.Cells(i, 5))
            ui_originid = getCellValueinString(UIStringTableSheet.Cells(i, 8))
            ui_en = getCellValueinString(UIStringTableSheet.Cells(i, 11))
            ui_key = ui_screenid + "_" + ui_originid + "_" + ui_en
            ui_string_table[ui_key] = str(i)
            i += 1

        i = 5

        while(not isEmptyValue(historySheet.Cells(i, 3))):
            i += 1

        history_start_row = i

        rtInfo = stringTableSheet.UsedRange
        nRTrows = rtInfo.Rows.Count
        nRTcols = rtInfo.Columns.Count

        for i in range(3, nRTrows + 1):
            st_screenid = getCellValueinString(stringTableSheet.Cells(i, 5))
            st_originid = getCellValueinString(stringTableSheet.Cells(i, 9))
            st_en = getCellValueinString(stringTableSheet.Cells(i, 12))
            st_key = st_screenid + "_" + st_originid + "_" + st_en
            st_status = getCellValueinString(stringTableSheet.Cells(i, 11))
            if st_status != "noUse":
                try:
                    row_no = int(ui_string_table[st_key])
                    UI_Stringid = getCellValueinString(UIStringTableSheet.Cells(row_no, 7))
                    stringTableSheet.Cells(i, 8).Value = UI_Stringid
                    string_id = getCellValueinString(stringTableSheet.Cells(i, 7))
                    if UI_Stringid != string_id:
                        stringTableSheet.Cells(i, 8).interior.color = rgb_to_hex((0,0,255))
                    else:
                        UIStringTableSheet.Cells(row_no, 7).interior.color = rgb_to_hex((0,0,255)) 
                    #del UI_Stringid[st_key]

                except Exception as e:
                    print(e)
                    print(st_key)

            i += 1

        UIStringTableWb.Close(SaveChanges = 1)
        stringTableWb.Close(SaveChanges = 1)

    def UpdateAllwordsToStringTable(self):
        rtStringTablefile = os.getcwd() + "\\doc\\rt_stringtable_collection.xlsx"
        rtStringTableWb  = self.excelApp.Workbooks.Open(rtStringTablefile)
        rtStringTableSheet = rtStringTableWb.Sheets('All Words') 

        stringTablefile = os.getcwd() + "\\doc\\21MM_Material_09_StringTable.xlsx"
        stringTableWb  = self.excelApp.Workbooks.Open(stringTablefile)
        stringTableSheet = stringTableWb.Sheets('All Words') 
        historySheet = stringTableWb.Sheets('History') 

        info = stringTableSheet.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count

        st_string_table = {}
        for i in range(3, nrows + 1):
            st_states = getCellValueinString(stringTableSheet.Cells(i, 25))
            if st_states != "noUse" and st_states != "wrong" and st_states != "repeat":
                st_id = getCellValueinString(stringTableSheet.Cells(i, 7))              
                st_string_table[st_id] = str(i)
            i += 1
        

        i = 5

        while(not isEmptyValue(historySheet.Cells(i, 3))):
            i += 1

        history_start_row = i

        rtInfo = rtStringTableSheet.UsedRange
        nRTrows = rtInfo.Rows.Count
        nRTcols = rtInfo.Columns.Count

        for i in range(2, nRTrows + 1):
            if isEmptyValue(rtStringTableSheet.Cells(i, 3)):
                break
            rt_status = getCellValueinString(rtStringTableSheet.Cells(i, 14))
            if rt_status == "Need update":
                rt_ids = getCellValueinString(rtStringTableSheet.Cells(i, 10))
                rt_en = getCellValueinString(rtStringTableSheet.Cells(i, 6))
                rt_cn = getCellValueinString(rtStringTableSheet.Cells(i, 7))

                rt_id_list = rt_ids.split("\n")
                for rt_id in rt_id_list:
                    try:
                        st_row =  int(st_string_table[rt_id])
                        stringTableSheet.Cells(st_row, 11).Value = rt_en
                        stringTableSheet.Cells(st_row, 23).Value = rt_cn

                        stringTableSheet.Cells(st_row, 11).interior.color = rgb_to_hex((255,223,249))                
                        stringTableSheet.Cells(st_row, 23).interior.color = rgb_to_hex((255,223,249))                
                        
                        rtStringTableSheet.Cells(i, 14).Value = "Need update - updated"
                                
                        stringTableSheet.Cells(st_row, 25).Value = "Using"


                        range_a = "A" + str(history_start_row) + ":R" + str(history_start_row) 
                        historySheet.Range(range_a).EntireRow.Insert()
                        range_s= "A" + str(history_start_row-1) + ":" + "R" + str(history_start_row-1)
                        range_t= "A" + str(history_start_row) + ":" + "R" + str(history_start_row)

                        historySheet.Range(range_s).Copy(historySheet.Cells.Range(range_t))
                        historySheet.Cells(history_start_row, 8).Value = getCellValueinString(rtStringTableSheet.Cells(i, 1))

                        historySheet.Cells(history_start_row, 13).Value = getCellValueinString(stringTableSheet.Cells(st_row, 1))
                        historySheet.Cells(history_start_row, 14).Value = getCellValueinString(rtStringTableSheet.Cells(i, 12)) + "\n"+ getCellValueinString(rtStringTableSheet.Cells(i, 13)) 
                        historySheet.Cells(history_start_row, 15).Value = getCellValueinString(stringTableSheet.Cells(st_row, 11)) + "\n"+ getCellValueinString(stringTableSheet.Cells(st_row, 23)) 

                        history_start_row += 1
                    except Exception as e:
                        print("not found")
                        rtStringTableSheet.Cells(i, 14).Value = "Need update - failed"
            i += 1

        rtStringTableWb.Close(SaveChanges = 1)
        stringTableWb.Close(SaveChanges = 1)

    def CheckNTString(self):
        rtStringTablefile = os.getcwd() + "\\doc\\rt_stringtable_collection.xlsx"
        rtStringTableWb  = self.excelApp.Workbooks.Open(rtStringTablefile)
        rtStringTableSheet = rtStringTableWb.Sheets('All Words') 

        stringTablefile = os.getcwd() + "\\doc\\21MM_Material_09_StringTable.xlsx"
        stringTableWb  = self.excelApp.Workbooks.Open(stringTablefile)
        stringTableSheet = stringTableWb.Sheets('All Words') 

        info = stringTableSheet.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count

        st_string_table = {}
        for i in range(3, nrows + 1):
            st_states = getCellValueinString(stringTableSheet.Cells(i, 25))
            if st_states.upper() not in ("NOUSE","WRONG"):

                st_id = getCellValueinString(stringTableSheet.Cells(i, 7))
                st_en = getCellValueinString(stringTableSheet.Cells(i, 11))
                st_cn = getCellValueinString(stringTableSheet.Cells(i, 23))
                
                st_string_table[st_id] = st_en + "|" + st_cn
            i += 1
        
        rtInfo = rtStringTableSheet.UsedRange
        nRTrows = rtInfo.Rows.Count
        nRTcols = rtInfo.Columns.Count

        for i in range(2, nRTrows + 1):
            if isEmptyValue(rtStringTableSheet.Cells(i, 3)):
                break
            if not isEmptyValue(rtStringTableSheet.Cells(i, 10)):
                rt_ids = getCellValueinString(rtStringTableSheet.Cells(i, 10))
                rtStringTableSheet.Cells(i, 15).Value = ">"
                rtStringTableSheet.Cells(i, 16).Value = ">"
                if rt_ids not in ("Mixed string", "Deleted","Variable string","ignore","repeat","Impl_picture"):
                    rt_en = getCellValueinString(rtStringTableSheet.Cells(i, 6))
                    rt_cn = getCellValueinString(rtStringTableSheet.Cells(i, 7))

                    rt_id_list = rt_ids.split("\n")
                    for rt_id in rt_id_list:
                        try:
                            st_en_list =  st_string_table[rt_id].split("|")[0]
                            st_cn_list =  st_string_table[rt_id].split("|")[1]

                            rtStringTableSheet.Cells(i, 15).Value  = st_en_list + "|" + rtStringTableSheet.Cells(i, 15).Value
                            rtStringTableSheet.Cells(i, 16).Value  = st_cn_list + "|" + rtStringTableSheet.Cells(i, 16).Value

                            rtStringTableSheet.Cells(i, 12).Value = st_en_list
                            rtStringTableSheet.Cells(i, 13).Value = st_cn_list

                            if rt_en != st_en_list or rt_cn != st_cn_list:
                                rtStringTableSheet.Cells(i, 14).Value = "Need update"
                            

                            stringTableSheet.Cells(i, 25).Value = "Using"
                        except Exception as e:
                            print("not found")
                            rtStringTableSheet.Cells(i, 14).Value = "wrong string id"
            i += 1

        rtStringTableWb.Close(SaveChanges = 1)
        stringTableWb.Close(SaveChanges = 1)

    def RetrieveNTStringIDforRT(self):
        rtStringTablefile = os.getcwd() + "\\doc\\rt_ntstringtable.xlsx"
        rtStringTableWb  = self.excelApp.Workbooks.Open(rtStringTablefile)
        rtStringTableSheet = rtStringTableWb.Sheets('All Words') 

        stringTablefile = os.getcwd() + "\\doc\\21MM_Material_09_StringTable.xlsx"
        stringTableWb  = self.excelApp.Workbooks.Open(stringTablefile)
        stringTableSheet = stringTableWb.Sheets('All Words') 

        rtInfo = rtStringTableSheet.UsedRange
        nRTrows = rtInfo.Rows.Count
        nRTcols = rtInfo.Columns.Count

        info = stringTableSheet.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count

        systemView_list = []
        for i in range(2,nRTrows +1):
            for j in range(3, nrows + 1):
                st_id = getCellValueinString(stringTableSheet.Cells(j,1))
                if (getCellValueinString(stringTableSheet.Cells(j, 2)) == "systemview") and (st_id != "matched"):
                    rt_en_string = getCellValueinString(rtStringTableSheet.Cells(i, 6))
                    st_en_string = getCellValueinString(stringTableSheet.Cells(j, 11))
                    st_res_id = getCellValueinString(stringTableSheet.Cells(j, 7))
                    if rt_en_string == st_en_string:
                        rtStringTableSheet.Cells(i, 8).Value = getCellValueinString(stringTableSheet.Cells(j, 5))
                        rtStringTableSheet.Cells(i, 9).Value = st_res_id
                        stringTableSheet.Cells(j, 1).Value = "matched"
                        break
                j += 1
            i += 1

        rtStringTableWb.Close(SaveChanges = 1)
        stringTableWb.Close(SaveChanges = 1)

    def RetrieveStringIDforRT(self):
        rtStringTablefile = os.getcwd() + "\\doc\\rt_stringtable_collection.xlsx"
        rtStringTableWb  = self.excelApp.Workbooks.Open(rtStringTablefile)
        rtStringTableSheet = rtStringTableWb.Sheets('All Words') 

        stringTablefile = os.getcwd() + "\\doc\\21MM_Material_09_StringTable.xlsx"
        stringTableWb  = self.excelApp.Workbooks.Open(stringTablefile)
        stringTableSheet = stringTableWb.Sheets('All Words') 

        rtInfo = rtStringTableSheet.UsedRange
        nRTrows = rtInfo.Rows.Count
        nRTcols = rtInfo.Columns.Count

        info = stringTableSheet.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count

        for i in range(2,nRTrows +1):
            if isEmptyValue(rtStringTableSheet.Cells(i, 10)) :
                rt_function = getCellValueinString(rtStringTableSheet.Cells(i, 2))
                rt_screenids = getCellValueinString(rtStringTableSheet.Cells(i, 8))
                rt_screen_list = rt_screenids.split("\n")
                rt_en_string = getCellValueinString(rtStringTableSheet.Cells(i, 6))
                for j in range(3, nrows +1):
                    st_id = getCellValueinString(stringTableSheet.Cells(j,1))
                    if st_id != "matched" and rt_function != "diag" and rt_function != "debug":
                        st_function = getCellValueinString(stringTableSheet.Cells(j, 4))
                        st_screenid = getCellValueinString(stringTableSheet.Cells(j, 5))
                        st_en_string = getCellValueinString(stringTableSheet.Cells(j, 11))
                        if st_screenid in rt_screen_list:
                            if (st_screenid + rt_en_string) == (st_screenid + st_en_string):
                                st_res_id = getCellValueinString(stringTableSheet.Cells(j, 7)) + "|"+ getCellValueinString(rtStringTableSheet.Cells(i, 10))
                                rtStringTableSheet.Cells(i, 10).Value = st_res_id
                                stringTableSheet.Cells(j, 1).Value = "matched"
                                rt_screen_list.remove(st_screenid)
                                if len(rt_screen_list) == 0:
                                    break
                    j += 1
            i += 1

        rtStringTableWb.Close(SaveChanges = 1)
        stringTableWb.Close(SaveChanges = 1)

    def AbstractRTScreenAllwords(self, root):
        RTStringTablefile = os.getcwd() + "\\doc\\rt_stringtable_collection.xlsx"
        rtStringTableWb  = self.excelApp.Workbooks.Open(RTStringTablefile)
        stringTableSheet = rtStringTableWb.Sheets('All Words new') 

        wRow = 2
        for root, dirs,files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root, name)
                try:

                    file = os.getcwd() + "\\" + fileName
                    specWb  = self.excelApp.Workbooks.Open(file)
                    sheetCount = (specWb.Worksheets.Count)

                    for i in range(1, sheetCount + 1):
                        specSheet = specWb.Worksheets(i)
                        sheet_name = specWb.Worksheets(i).Name
                        
                        if (sheet_name.upper() != "RELEASE NOTE") and (sheet_name.find("削除") < 0):
                            print('Search in this ->',sheet_name)    # 绝对路径                            
    
                            info = specSheet.UsedRange
                            nrows = info.Rows.Count
                            ncols = info.Columns.Count
                            
                            screen_id = getCellValueinString(specSheet.Cells(2, 1))

                            rowNo  = 3
                            partID = ""
                            No_colNo = 0
                            English_colNo = 0
                            Chinese_colNo = 0
                            partName_colNo = 0

                            while rowNo <= nrows:
                                for i in range(2, ncols):
                                    if getCellValueinString(specSheet.Cells(rowNo, i)).find("No.") >= 0 \
                                        or getCellValueinString(specSheet.Cells(rowNo, i)).find("画面ID") >= 0:
                                        No_colNo = i
                                    if getCellValueinString(specSheet.Cells(rowNo, i)).find("英语") >= 0 \
                                        or getCellValueinString(specSheet.Cells(rowNo, i)).find("英語") >= 0\
                                        or getCellValueinString(specSheet.Cells(rowNo, i)).find("英文") >= 0: 
                                        English_colNo = i
                                    if getCellValueinString(specSheet.Cells(rowNo, i)).find("中文") >= 0 \
                                    or getCellValueinString(specSheet.Cells(rowNo, i)).find("中国語") >= 0:
                                        Chinese_colNo = i
                                    if getCellValueinString(specSheet.Cells(rowNo, i)).find("部品") >= 0 :
                                        partName_colNo = i
                                
                                if No_colNo > 0:
                                    rowNo += 1
                                    break

                                rowNo += 1

                            while rowNo <= nrows:
                                if getCellValueinString(specSheet.Cells(rowNo, No_colNo)).find("备注") >= 0: 
                                    break


                                if not isEmptyValue(specSheet.Cells(rowNo, 2)):
                                    partID = getCellValueinString(specSheet.Cells(rowNo, 2))
                                    
                                if (specSheet.Cells(rowNo, partName_colNo).Font.Strikethrough == False) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, English_colNo)).find("对象外") < 0) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, English_colNo)).upper().find("同SLIDE IN MENU") < 0) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, English_colNo)).upper().find("同SLIDE-IN MENU") < 0) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, Chinese_colNo)).find("对象外") < 0) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, Chinese_colNo)).find("对向外") < 0) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, English_colNo)).find("参见") < 0) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, No_colNo)).find("No.") < 0) \
                                    and not isEmptyValue(specSheet.Cells(rowNo, English_colNo)):
                                    if isEmptyValue(specSheet.Cells(rowNo, No_colNo)):
                                        stringTableSheet.Cells(wRow, 4).Value = partID
                                    else:
                                        stringTableSheet.Cells(wRow, 4).Value =  getCellValueinString(specSheet.Cells(rowNo, No_colNo))

                                    stringTableSheet.Cells(wRow, 5).Value =  getCellValueinString(specSheet.Cells(rowNo, partName_colNo))
                                    stringTableSheet.Cells(wRow, 6).Value =  getCellValueinString(specSheet.Cells(rowNo, English_colNo))
                                    stringTableSheet.Cells(wRow, 7).Value =  getCellValueinString(specSheet.Cells(rowNo, Chinese_colNo))
                                    stringTableSheet.Cells(wRow, 3).Value = screen_id.replace("画面ID：","")

                                    stringTableSheet.Cells(wRow, 1).Value = fileName
                                    wRow += 1

                                rowNo += 1

                    specWb.Close(SaveChanges = 0)
                except Exception as e:
                    print(e)

        rtStringTableWb.Close(SaveChanges = 1)

    def AbstractRTNotificationAllwords(self, root):
        RTStringTablefile = os.getcwd() + "\\doc\\rt_ntstringtable_collection.xlsx"
        rtStringTableWb  = self.excelApp.Workbooks.Open(RTStringTablefile)
        stringTableSheet = rtStringTableWb.Sheets('All Words new') 

        wRow = 2
        for root, dirs,files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root, name)
                try:

                    file = os.getcwd() + "\\" + fileName
                    specWb  = self.excelApp.Workbooks.Open(file)
                    sheetCount = (specWb.Worksheets.Count)

                    stringTableSheet.Cells(wRow, 1).Value = fileName

                    for i in range(1, sheetCount + 1):
                        specSheet = specWb.Worksheets(i)
                        sheet_name = specWb.Worksheets(i).Name
                        
                        if (sheet_name.upper() != "RELEASE NOTE") and (sheet_name.find("削除") < 0):
                            print('Search in this ->',sheet_name)    # 绝对路径                            
    
                            info = specSheet.UsedRange
                            nrows = info.Rows.Count
                            ncols = info.Columns.Count
                            
                            screen_id = getCellValueinString(specSheet.Cells(2, 1))

                            stringTableSheet.Cells(wRow, 3).Value = sheet_name
                            rowNo  = 1
                            partID = ""
                            No_colNo = 0
                            title_colNo = 0
                            title_en_colNo = 0
                            title_cn_colNo = 0
                            content_colNo = 0
                            content_en_colNo = 0
                            content_cn_colNo = 0
                            switch_colNo = 0
                            switch_en_colNo = 0
                            switch_cn_colNo = 0

                            while rowNo <= nrows:
                                for i in range(1, ncols):
                                    if getCellValueinString(specSheet.Cells(rowNo, i)).find("No.") >= 0 :
                                        No_colNo = i
                                    if (getCellValueinString(specSheet.Cells(rowNo, i)).find("英文") >= 0) and (title_colNo > 0) and (title_en_colNo == 0):
                                        title_en_colNo = i
                                    if (getCellValueinString(specSheet.Cells(rowNo, i)).find("中文") >= 0) and (title_colNo > 0) and (title_cn_colNo == 0) :
                                        title_cn_colNo = i
                                    if getCellValueinString(specSheet.Cells(rowNo, i)).find("ONS标题") >= 0 :
                                        title_colNo = i
                                    if (getCellValueinString(specSheet.Cells(rowNo, i)).find("英文") >= 0 ) and (content_colNo > 0) and (content_en_colNo == 0) :
                                        content_en_colNo = i
                                    if (getCellValueinString(specSheet.Cells(rowNo, i)).find("中文") >= 0 ) and (content_colNo > 0) and (content_cn_colNo == 0) :
                                        content_cn_colNo = i
                                    if getCellValueinString(specSheet.Cells(rowNo, i)).find("ONS内容") >= 0 :
                                        content_colNo = i
                                    if (getCellValueinString(specSheet.Cells(rowNo, i)).find("英文") >= 0 ) and (switch_colNo > 0) and (switch_en_colNo == 0) :
                                        switch_en_colNo = i
                                    if (getCellValueinString(specSheet.Cells(rowNo, i)).find("中文") >= 0 ) and (switch_colNo > 0) and (switch_cn_colNo == 0) :
                                        switch_cn_colNo = i
                                    if getCellValueinString(specSheet.Cells(rowNo, i)).find("Text") >= 0 :
                                        switch_colNo = i

                                if No_colNo > 0:
                                    rowNo += 1
                                    break

                                rowNo += 1

                            while rowNo <= nrows:
                                stringTableSheet.Cells(wRow, 3).Value = sheet_name

                                if getCellValueinString(specSheet.Cells(rowNo, No_colNo)).find("备注") >= 0: 
                                    break

                                if (specSheet.Cells(rowNo, title_colNo).Font.Strikethrough == False) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, title_en_colNo)).find("对象外") < 0) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, title_cn_colNo)).find("对象外") < 0) \
                                    and getCellValueinString(specSheet.Cells(rowNo, title_cn_colNo)) != "-" :
                                    
                                    stringTableSheet.Cells(wRow, 4).Value =  getCellValueinString(specSheet.Cells(rowNo, No_colNo))
                                    stringTableSheet.Cells(wRow, 5).Value =  getCellValueinString(specSheet.Cells(rowNo, title_colNo))
                                    stringTableSheet.Cells(wRow, 6).Value =  getCellValueinString(specSheet.Cells(rowNo, title_en_colNo))
                                    stringTableSheet.Cells(wRow, 7).Value =  getCellValueinString(specSheet.Cells(rowNo, title_cn_colNo))
                                    wRow += 1

                                if (specSheet.Cells(rowNo, title_colNo).Font.Strikethrough == False) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, content_en_colNo)).find("对象外") < 0) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, content_cn_colNo)).find("对象外") < 0) \
                                    and getCellValueinString(specSheet.Cells(rowNo, content_cn_colNo)) != "-" :
                                    
                                    stringTableSheet.Cells(wRow, 4).Value =  getCellValueinString(specSheet.Cells(rowNo, No_colNo))
                                    stringTableSheet.Cells(wRow, 5).Value =  getCellValueinString(specSheet.Cells(rowNo, content_colNo))
                                    stringTableSheet.Cells(wRow, 6).Value =  getCellValueinString(specSheet.Cells(rowNo, content_en_colNo))
                                    stringTableSheet.Cells(wRow, 7).Value =  getCellValueinString(specSheet.Cells(rowNo, content_cn_colNo))

                                    wRow += 1

                                if (specSheet.Cells(rowNo, title_colNo).Font.Strikethrough == False) and switch_colNo >0 :

                                    if  (getCellValueinString(specSheet.Cells(rowNo, switch_en_colNo)).find("对象外") < 0) \
                                        and (getCellValueinString(specSheet.Cells(rowNo, switch_cn_colNo)).find("对象外") < 0) \
                                        and getCellValueinString(specSheet.Cells(rowNo, switch_colNo)) != "-" :
                                        
                                        stringTableSheet.Cells(wRow, 4).Value =  getCellValueinString(specSheet.Cells(rowNo, No_colNo))
                                        stringTableSheet.Cells(wRow, 5).Value =  getCellValueinString(specSheet.Cells(rowNo, switch_colNo))
                                        stringTableSheet.Cells(wRow, 6).Value =  getCellValueinString(specSheet.Cells(rowNo, switch_en_colNo))
                                        stringTableSheet.Cells(wRow, 7).Value =  getCellValueinString(specSheet.Cells(rowNo, switch_cn_colNo))

                                        wRow += 1

                                rowNo += 1

                    specWb.Close(SaveChanges = 0)
                except Exception as e:
                    print(e)

        rtStringTableWb.Close(SaveChanges = 1)

    def UpdateRTScreenAllwords(self, root):
        RTStringTablefile = os.getcwd() + "\\doc\\rt_stringtable_collection.xlsx"
        rtStringTableWb  = self.excelApp.Workbooks.Open(RTStringTablefile)
        stringTableSheet = rtStringTableWb.Sheets('All Words') 

        wRow = 2
        for root, dirs,files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root, name)
                try:

                    file = os.getcwd() + "\\" + fileName
                    specWb  = self.excelApp.Workbooks.Open(file)
                    sheetCount = (specWb.Worksheets.Count)

                    stringTableSheet.Cells(wRow, 1).Value = fileName
                    for i in range(1, sheetCount + 1):
                        specSheet = specWb.Worksheets(i)
                        sheet_name = specWb.Worksheets(i).Name
                        
                        if (sheet_name.upper() != "RELEASE NOTE") and (sheet_name.find("削除") < 0):
                            print('Search in this ->',sheet_name)    # 绝对路径                            
    
                            info = specSheet.UsedRange
                            nrows = info.Rows.Count
                            ncols = info.Columns.Count
                            
                            screen_id = getCellValueinString(specSheet.Cells(2, 1))

                            rowNo  = 3
                            partID = ""
                            No_colNo = 0
                            English_colNo = 0
                            Chinese_colNo = 0
                            partName_colNo = 0

                            while rowNo <= nrows:
                                for i in range(2, ncols):
                                    if getCellValueinString(specSheet.Cells(rowNo, i)).find("No.") >= 0 \
                                        or getCellValueinString(specSheet.Cells(rowNo, i)).find("画面ID") >= 0:
                                        No_colNo = i
                                    if getCellValueinString(specSheet.Cells(rowNo, i)).find("英语") >= 0 \
                                        or getCellValueinString(specSheet.Cells(rowNo, i)).find("英語") >= 0\
                                        or getCellValueinString(specSheet.Cells(rowNo, i)).find("英文") >= 0: 
                                        English_colNo = i
                                    if getCellValueinString(specSheet.Cells(rowNo, i)).find("中文") >= 0 \
                                    or getCellValueinString(specSheet.Cells(rowNo, i)).find("中国語") >= 0:
                                        Chinese_colNo = i
                                    if getCellValueinString(specSheet.Cells(rowNo, i)).find("部品") >= 0 :
                                        partName_colNo = i
                                
                                if No_colNo > 0:
                                    rowNo += 1
                                    break

                                rowNo += 1

                            while rowNo <= nrows:
                                if getCellValueinString(specSheet.Cells(rowNo, No_colNo)).find("备注") >= 0: 
                                    break


                                if not isEmptyValue(specSheet.Cells(rowNo, 2)):
                                    partID = getCellValueinString(specSheet.Cells(rowNo, 2))
                                    
                                if (specSheet.Cells(rowNo, partName_colNo).Font.Strikethrough == False) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, English_colNo)).find("对象外") < 0) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, English_colNo)).upper().find("同SLIDE IN MENU") < 0) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, English_colNo)).upper().find("同SLIDE-IN MENU") < 0) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, Chinese_colNo)).find("对象外") < 0) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, Chinese_colNo)).find("对向外") < 0) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, English_colNo)).find("参见") < 0) \
                                    and (getCellValueinString(specSheet.Cells(rowNo, No_colNo)).find("No.") < 0) \
                                    and not isEmptyValue(specSheet.Cells(rowNo, English_colNo)):
                                    if isEmptyValue(specSheet.Cells(rowNo, No_colNo)):
                                        stringTableSheet.Cells(wRow, 4).Value = partID
                                    else:
                                        stringTableSheet.Cells(wRow, 4).Value =  getCellValueinString(specSheet.Cells(rowNo, No_colNo))

                                    stringTableSheet.Cells(wRow, 5).Value =  getCellValueinString(specSheet.Cells(rowNo, partName_colNo))
                                    stringTableSheet.Cells(wRow, 6).Value =  getCellValueinString(specSheet.Cells(rowNo, English_colNo))
                                    stringTableSheet.Cells(wRow, 7).Value =  getCellValueinString(specSheet.Cells(rowNo, Chinese_colNo))
                                    stringTableSheet.Cells(wRow, 3).Value = screen_id.replace("画面ID：","")
                                    wRow += 1

                                rowNo += 1

                    specWb.Close(SaveChanges = 0)
                except Exception as e:
                    print(e)

        rtStringTableWb.Close(SaveChanges = 1)

    def UpdateTMCAllwords(self):
        tmna_allwords_file = os.getcwd() + "\\res\TMNA\\21MM_HMI_NA Allwords_V2.7_Translation.xlsx"
        cf_file = os.getcwd() + "\\res\TMNA\\21MM_HMI_NA Allwords_V2.541R_CA-FR v4.xlsx"
        ms_file = os.getcwd() + "\\res\TMNA\\21MM_HMI_NA Allwords_V2.541R_MX-SP v4.xlsx"

        tmna_stringwb  = self.excelApp.Workbooks.Open(tmna_allwords_file)
        tmna_allwords_sheet = tmna_stringwb.Sheets('Allwords') 
        change_sheet = tmna_stringwb.Sheets('word_list') 

        cf_file_wb = self.excelApp.Workbooks.Open(cf_file)
        cf_sheet = cf_file_wb.Sheets('Allwords') 

        ms_file_wb = self.excelApp.Workbooks.Open(ms_file)
        ms_sheet = ms_file_wb.Sheets('Allwords') 


        info = change_sheet.UsedRange
        nrows = info.Rows.Count

        change_list = {}
        for i in range(1, nrows + 1):
            change_list[getCellValueinString(change_sheet.Cells(i, 4))] = getCellValueinString(change_sheet.Cells(i, 7))
            i += 1
        
        info = cf_sheet.UsedRange
        nrows = info.Rows.Count

        for i in range(5, nrows + 1):
            keystr_cf =  getCellValueinString(cf_sheet.Cells(i, 5)) + getCellValueinString(cf_sheet.Cells(i, 6))+ getCellValueinString(cf_sheet.Cells(i, 7)) + getCellValueinString(cf_sheet.Cells(i, 4))
            if keystr_cf in change_list:
                row_no = int(change_list[keystr_cf])
                tmna_allwords_sheet.Cells(row_no, 9).Value = getCellValueinString(cf_sheet.Cells(i, 9))
                tmna_allwords_sheet.Cells(row_no, 10).Value = getCellValueinString(cf_sheet.Cells(i, 10))
                tmna_allwords_sheet.Cells(row_no, 9).interior.color = rgb_to_hex(PINK_COLOR)
                tmna_allwords_sheet.Cells(row_no, 10).interior.color = rgb_to_hex(PINK_COLOR)

            i += 1

        info = ms_sheet.UsedRange
        nrows = info.Rows.Count
        for i in range(5, nrows + 1):
            keystr_ms =  getCellValueinString(ms_sheet.Cells(i, 5)) + getCellValueinString(ms_sheet.Cells(i, 6)) +  getCellValueinString(ms_sheet.Cells(i, 7)) + getCellValueinString(ms_sheet.Cells(i, 4))
            if keystr_ms in change_list:
                row_no = int(change_list[keystr_ms])
                tmna_allwords_sheet.Cells(row_no, 11).Value = getCellValueinString(ms_sheet.Cells(i, 9))
                tmna_allwords_sheet.Cells(row_no, 12).Value = getCellValueinString(ms_sheet.Cells(i, 10))
                tmna_allwords_sheet.Cells(row_no, 11).interior.color = rgb_to_hex(PINK_COLOR)
                tmna_allwords_sheet.Cells(row_no, 12).interior.color = rgb_to_hex(PINK_COLOR)

                del change_list[keystr_ms]
            i += 1

        i = 1
        for keystr in change_list.keys():
            rowno = int(change_list[keystr])
            tmna_allwords_sheet.Cells(rowno, 7).interior.color = rgb_to_hex(GRAY_COLOR) 
            i += 1

        tmna_stringwb.Close(SaveChanges = 1)
        cf_file_wb.Close(SaveChanges = 0)
        ms_file_wb.Close(SaveChanges = 0)


    def CheckTBDItem(self):
        specSchedulefile = os.getcwd() + "\\doc\\机能式样书TBD一览.xlsx"
        specScheduleTableWb  = self.excelApp.Workbooks.Open(specSchedulefile)
        specScheduleSheet = specScheduleTableWb.Sheets('Sheet1') 
        
        wRow = 2
        for root, dirs, files in os.walk("res\\Tarim"):
            for name in files: 
                fileName = os.path.join(root, name)
                try:
                    file = os.getcwd() + "\\" + fileName
                    specWb  = self.excelApp.Workbooks.Open(file)
                    specScheduleSheet.Cells(wRow, 1).Value = name
                    specScheduleSheet.Cells(wRow, 2).Value = "无"
                    specSheet = specWb.Sheets("TBD情况")
                    
                    startRow = 4

                    info = specSheet.UsedRange
                    nrows = info.Rows.Count

                    while startRow <= nrows:
                        specScheduleSheet.Cells(wRow, 1).Value = name
                        specScheduleSheet.Cells(wRow, 2).Value = getCellValueinString(specSheet.Cells(startRow, 2))
                        specScheduleSheet.Cells(wRow, 3).Value = getCellValueinString(specSheet.Cells(startRow, 3))
                        specScheduleSheet.Cells(wRow, 4).Value = getCellValueinString(specSheet.Cells(startRow, 4))
                        specScheduleSheet.Cells(wRow, 5).Value = getCellValueinString(specSheet.Cells(startRow, 5))
                        specScheduleSheet.Cells(wRow, 6).Value = getCellValueinString(specSheet.Cells(startRow, 6))
                        specScheduleSheet.Cells(wRow, 7).Value = getCellValueinString(specSheet.Cells(startRow, 7))

                        wRow += 1
                        startRow += 1

                    specWb.Close(SaveChanges = 0)

                except Exception as e:
                    print(e)
                    print("No TBD info in this file", fileName)
                    

        specScheduleTableWb.Close(SaveChanges = 1)

    def trace_feature_func(self):
        specFuncFeaturefile = os.getcwd() + "\\featurelist\\CCS5.0_IVI_Feature_Func_list.xlsx"
        specFuncFeatureWb  = self.excelApp.Workbooks.Open(specFuncFeaturefile)
        specFuncFeatureSheet = specFuncFeatureWb.Sheets('Func list')
        specFeatureListSheet = specFuncFeatureWb.Sheets('Feature List')

        startRow = 4
        info = specFeatureListSheet.UsedRange
        nrows = info.Rows.Count
        
        feature_id_list = {}
        while startRow <= nrows:
            feature_id_list[getCellValueinString(specFeatureListSheet.Cells(startRow, 8))] = str(startRow)
            startRow += 1
      
        startRow = 2
        info = specFuncFeatureSheet.UsedRange
        nrows = info.Rows.Count   

        while startRow <= nrows:
            feature_in_func = getCellValueinString(specFuncFeatureSheet.Cells(startRow, 4))
            if feature_in_func in feature_id_list:
                specFuncFeatureSheet.Cells(startRow, 7).Value = feature_in_func
                found_row = feature_id_list[feature_in_func]
                specFuncFeatureSheet.Cells(startRow, 8).Value = getCellValueinString(specFeatureListSheet.Cells(int(found_row),9))

                func_info = getCellValueinString(specFuncFeatureSheet.Cells(startRow, 1)) + ": " + \
                            getCellValueinString(specFuncFeatureSheet.Cells(startRow, 2)) + " "  + \
                            getCellValueinString(specFuncFeatureSheet.Cells(startRow, 3))

                if isEmptyValue(specFeatureListSheet.Cells(int(found_row),10)) :
                    specFeatureListSheet.Cells(int(found_row),10).Value = func_info
                else:
                    last_info = getCellValueinString(specFeatureListSheet.Cells(int(found_row),10))
                    specFeatureListSheet.Cells(int(found_row),10).Value = last_info  + "\n" + func_info
            else:
                specFuncFeatureSheet.Cells(startRow, 7).Value = "not found in featurelist file"
            
            startRow += 1

    def extract_funcid(self):
        self.resultBook.ActiveSheet.Name = "Func list"
        resultSheet = self.resultBook.ActiveSheet
        resultSheet.Cells(1,1).Value = "File Name"
        resultSheet.Cells(1,2).Value = "Func ID"
        resultSheet.Cells(1,3).Value = "Func Name"
        resultSheet.Cells(1,4).Value = "Feature ID"
        resultSheet.Cells(1,5).Value = "Cell"
        
        wRow = 2
        for root, dirs, files in os.walk("func_119"):
            for name in files: 
                fileName = os.path.join(root, name)
                try:
                    file = os.getcwd() + "\\" + fileName
                    specWb  = self.excelApp.Workbooks.Open(file)
                    
                    sheetCount = specWb.Worksheets.Count
                    bHasData = False
                    for i in range(1, sheetCount + 1):
                        sheet_name = specWb.Worksheets(i).Name
                        if sheet_name.startswith("03.") :
                            print("start to extract this file:", name)
                            specSheet = specWb.Worksheets(i)

                            startRow = 2
                            info = specSheet.UsedRange
                            nrows = info.Rows.Count
                            ncols = info.Columns.Count
                            while startRow <= nrows:
                                if not isEmptyValue(specSheet.Cells(startRow,2)):
                                    bHasData = True
                                    func_id = getCellValueinString(specSheet.Cells(startRow,1))
                                    feature_id = getCellValueinString(specSheet.Cells(startRow,2))
                                    feature_id_list = feature_id.split('/')
                                    for j in range(0, len(feature_id_list)):
                                        resultSheet.Cells(wRow,1).Value = name
                                        resultSheet.Cells(wRow,2).Value = func_id

                                        title_str = ""
                                        colNo = 3
                                        while colNo < 27:
                                            if getCellValueinString(specSheet.Cells(startRow, colNo)) != "功能":
                                                title_str = title_str + getCellValueinString(specSheet.Cells(startRow, colNo))
                                            colNo += 1

                                        resultSheet.Cells(wRow,3).Value = title_str

                                        resultSheet.Cells(wRow,4).Value = feature_id_list[j]
                                        if specSheet.Cells(startRow,2).Font.Strikethrough == True:
                                            resultSheet.Cells(wRow,4).Font.Strikethrough = True
                                        
                                        resultSheet.Cells(wRow,5).Value = sheet_name + ":B" + str(startRow)

                                        wRow += 1
                                startRow += 1

                            if not bHasData:
                                resultSheet.Cells(wRow,1).Value = name
                                resultSheet.Cells(wRow,2).Value = "No Feature Id in this file"
                                wRow += 1

                    specWb.Close(SaveChanges = 0)
                except Exception as e:
                    print(e)
                    print("Open file error", fileName)
                    copyspecFile(fileName, "error_files")

        self.SaveResultFile("CCS5.0_Func_list1.xlsx")


    def fillFormForJira(self):
        featurefile = os.getcwd() + "\\CCS5.0_IVI_silverbox_Featurelist_Release_Plan (合并) V1.0_JIRA导入用_20211118.xlsx"
        featureTableWb  = self.excelApp.Workbooks.Open(featurefile)
        featureSheet = featureTableWb.Sheets('Feature List') 

        jiraFile = os.getcwd() + "\\jira批量导入模板-1.xlsx"
        jiraTableWb  = self.excelApp.Workbooks.Open(jiraFile)
        jiraSheet = jiraTableWb.Sheets('FeatureTask') 

        info = featureSheet.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count
        
        startRow = 11
        wRow = 2
        component_str = ""
        while startRow <= nrows:
            component_id = getCellValueinString(featureSheet.Cells(startRow, 9))

            if component_id.startswith("1."):

                summary_str = getCellValueinString(featureSheet.Cells(startRow, 7)) + " " + getCellValueinString(featureSheet.Cells(startRow, 8))

                if not isEmptyValue(featureSheet.Cells(startRow, 8)):
                    component_str= getCellValueinString(featureSheet.Cells(startRow, 4)).replace(" ","")

                if component_str == "GUI市场2.0（DA+Meter+VPA助手）":
                    component_str = "GUI商城"
                elif component_str.startswith("OTA"):
                    component_str = "OTA"
                elif component_str.startswith("VR&VPA"):
                    component_str = "VR"
                elif component_str.startswith("懒人听书"):
                    component_str = "听书"
                elif component_str.startswith("本地多媒体"):
                    component_str = "多媒体"
                elif component_str.startswith("电子用户手册"):
                    component_str = "电子用户手册"

                mid_fid = getCellValueinString(featureSheet.Cells(startRow, 7))
                mid_funcname = getCellValueinString(featureSheet.Cells(startRow, 8))

                description_str = getCellValueinString(featureSheet.Cells(startRow, 9)) + " " + getCellValueinString(featureSheet.Cells(startRow, 10))

                while (isEmptyValue(featureSheet.Cells(startRow + 1, 7))):
                    startRow += 1                    
                    f_id_name = getCellValueinString(featureSheet.Cells(startRow, 9)) + " " + getCellValueinString(featureSheet.Cells(startRow, 10))
                    description_str += "\n" +  f_id_name

            
                jiraSheet.Cells(wRow, 4).Value = summary_str
                jiraSheet.Cells(wRow, 6).Value = description_str
                jiraSheet.Cells(wRow, 8).Value = component_str
                jiraSheet.Cells(wRow, 9).Value = mid_fid
                jiraSheet.Cells(wRow, 10).Value = mid_funcname
                wRow += 1

            startRow += 1


        jiraTableWb.Close(SaveChanges = 1)

    def fillSusukiUpdateDate(self):
        
        for root, dirs, files in os.walk("spec"):
            for name in files: 
                if name not in (".DS_Store","Thumbs.db"):
                    print("start extract this file:", name)
                    fileName = os.path.join(root, name)
                    try:
                        file = os.getcwd() + "\\" + fileName
                        specWb  = self.excelApp.Workbooks.Open(file)
                        sheetCount = (specWb.Worksheets.Count)

                        for i in range(1, sheetCount + 1):
                            specSheet = specWb.Worksheets(i)
                            sheet_name = specWb.Worksheets(i).Name
                            info = specSheet.UsedRange
                            nrows = info.Rows.Count
                            row = 1
                            if sheet_name.upper() not in ("COVER","HISTORY"):
                                while (row < nrows):
                                    if getCellValueinString(specSheet.Cells(row,1)).upper() == "ID":
                                        specSheet.Cells(row, 37).Value = "2021/12/31"
                                    row += 1
                        
                        specWb.Close(SaveChanges = 1)

                    except Exception as e:
                        print(e)
                        copyspecFile(fileName, "error_files")

    def getGlobalInfo(self):
        rfqfile = os.getcwd() + "\\doc\\20211208 入手24MM LEXUS正式RFQ仕様書一览表_21cy对比.xlsx"
        rfqWb  = self.excelApp.Workbooks.Open(rfqfile)
        rfqSheet = rfqWb.Sheets('sheet1') 
        rfqFolderSheet = rfqWb.Sheets('sheet5')

        startRow = 2
        info = rfqFolderSheet.UsedRange
        nrows = info.Rows.Count
        rfq_folder_list = {}
        while startRow <= 1634:
            rfq_no = getCellValueinString(rfqFolderSheet.Cells(startRow, 3))
            rfq_folder_list[rfq_no] = getCellValueinString(rfqFolderSheet.Cells(startRow, 4))
            startRow += 1

        startRow = 2
        info = rfqSheet.UsedRange
        nrows = info.Rows.Count

        while startRow <= nrows:
            try:
                if not isEmptyValue(rfqSheet.Cells(startRow, 6)) and isEmptyValue(rfqSheet.Cells(startRow, 7)):
                    rfq_name = getCellValueinString(rfqSheet.Cells(startRow, 6))

                    pure_name = rfq_name[rfq_name.rfind("\\") + 1:]
                    if pure_name not in ("No this folder in 21CY","Not found in 21cy global"):
                        pure_name = pure_name[:24]
                        for old_file in rfq_folder_list:
                            if old_file.find(pure_name) >= 0 :
                                rfqSheet.Cells(startRow, 7).Value = rfq_folder_list[old_file]
                                rfqSheet.Cells(startRow, 8).Value = old_file

            except Exception as e:
                print(e)

            startRow += 1

        '''
        startRow = 1
        info = rfqFolderSheet.UsedRange
        nrows = info.Rows.Count
        rfq_folder_list = {}
        while startRow <= nrows:
            rfq_no = getCellValueinString(rfqFolderSheet.Cells(startRow, 1))[0:3]
            rfq_folder_list[rfq_no] = getCellValueinString(rfqFolderSheet.Cells(startRow, 1))
            startRow += 1

        startRow = 2
        info = rfqSheet.UsedRange
        nrows = info.Rows.Count

        while startRow <= nrows:
            try:
                folder_name = getCellValueinString(rfqSheet.Cells(startRow, 1))
                if folder_name.find("Global Spec\\") >= 0 and isEmptyValue(rfqSheet.Cells(startRow, 6)):
                    rfq_no = getCellValueinString(rfqSheet.Cells(startRow, 1)).split("Global Spec\\")[1][0:3]
                    rfq_name = getCellValueinString(rfqSheet.Cells(startRow, 2))
                    rfq_extra_name = rfq_name[rfq_name.rfind("."):]
                    if rfq_extra_name.upper() not in (".TXT",".DB",".ZIP"):
                        rfq_name = getCellValueinString(rfqSheet.Cells(startRow, 2))
                        rfqSheet.Cells(startRow, 6).Value = "No this folder in 21CY"

                        if rfq_no in rfq_folder_list:
                            rfq_folder = rfq_folder_list[rfq_no]
                            bFound = False
                            for root, dirs,files in os.walk("C:\\workspace\\shiryu\\RFQ&TOYOTA仕様書\\toyota_spec\\")  + rfq_folder):
                                if bFound:
                                    break
                                for fileName  in  files: 
                                    extra_name = fileName[fileName.rfind("."):]
                                    if rfq_extra_name.upper() == extra_name.upper():
                                        file_len = len(fileName) - 8
                                        if file_len <= 11:
                                            file_len += 8
                                        if rfq_name.find(fileName[0:file_len]) >= 0:
                                            rfqSheet.Cells(startRow, 6).Value = os.path.join(root,fileName)
                                            bFound = True
                                            break;                                            
                                        else:
                                            rfqSheet.Cells(startRow, 6).Value = "Not found in 21cy global"

            except Exception as e:
                print(e)

            startRow += 1
        '''

        rfqWb.Close(SaveChanges = 1)


    def findGlobalInfo(self, root):
        rfqfile = os.getcwd() + "\\doc\\20211208 入手24MM LEXUS正式RFQ仕様書一览表_21cy对比.xlsx"
        rfqWb  = self.excelApp.Workbooks.Open(rfqfile)
        rfqSheet = rfqWb.Sheets('sheet4') 

        wRow = 2
        for root, dirs,files in os.walk(root):
            for name in files: 
                fileName = os.path.join(root, name)
                if fileName.find(".xls") >= 0: 
                    try:
                        file = os.getcwd() + "\\" + fileName
                        specWb  = self.excelApp.Workbooks.Open(file)
                        specSheet = specWb.Sheets("00_Source List")
                        print('Search in this ->', fileName)    # 绝对路径

                        startRow = 2
                        info = specSheet.UsedRange
                        nrows = info.Rows.Count
                        ncols = info.Columns.Count
                        startRecord = False
                        while startRow <= nrows:
                            if startRecord and not isEmptyValue(specSheet.Cells(startRow,1)):
                                rfqSheet.Cells(wRow, 1).Value =  fileName 

                                range_t = "B" + str(wRow) + ":" + "G" + str(wRow) 
                                range_s = "B" + str(startRow) + ":" + "G" + str(startRow) 
                            
                                specSheet.Cells.Range(range_s).Copy(rfqSheet.Cells.Range(range_t))

                                wRow += 1                          

                            if getCellValueinString(specSheet.Cells(startRow,1)) == "Upstream Documents":
                                startRecord = True
                                startRow += 1

                            if getCellValueinString(specSheet.Cells(startRow,1)) == "Reference Documents":
                                break
                            
                            startRow += 1

                        specWb.Close(SaveChanges = 0)

                    except Exception as e:
                        print(e)
                        copyspecFile(fileName, "no source files")

        rfqWb.Close(SaveChanges = 1)


    def CheckVersionInfo_24MM(self):
        rfqfile = os.getcwd() + "\\doc\\24MM LEXUS正式RFQ仕様書一览表_21cy对比_0209.xlsx"
        rfqWb  = self.excelApp.Workbooks.Open(rfqfile)
        rfqSheet = rfqWb.Sheets('对比结果（软件）') 

        startRow = 2
        info = rfqSheet.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count
        startRecord = False
        rfq_file_list = {}
        rfq_2124_file_list = {}

        while startRow <= nrows:
            rfq_filename = getCellValueinString(rfqSheet.Cells(startRow, 2))
            versionMark = getCellValueinString(rfqSheet.Cells(startRow, 8))
            version21 = getCellValueinString(rfqSheet.Cells(startRow, 7))
            version24 = getCellValueinString(rfqSheet.Cells(startRow, 3))

            version21 = version21.replace("V","")
            version24 = version24.replace("V","")
            
            result = ""
            result_file = ""

            if isEmptyValue(rfqSheet.Cells(startRow, 6)):
                result = "24新规"
            else:
                fileName_21cy = getCellValueinString(rfqSheet.Cells(startRow, 6))
                pure_name = fileName_21cy[fileName_21cy.rfind("\\") + 1:]
                if pure_name in ("No this folder in 21CY","Not found in 21cy global"):
                    result = "24新规"
                elif isEmptyValue(rfqSheet.Cells(startRow, 5)):
                    if pure_name.find(version21) >= 0 and rfq_filename.find(version24) >= 0:
                        if pure_name == rfq_filename:
                            result_file = "文件名一致"
                        else:
                            result_file = "文件名近似"
                       
                        if version21 > version24:
                            result = "21版本更高"
                        elif version21 < version24:
                            result = "24版本更高"
                        else:
                            result = "21/24版本一致"
                    else:
                        if pure_name.find(version21) < 0 :
                            result = "21版本信息不明"
                        if rfq_filename.find(version24) < 0 :
                            result += " 24版本信息不明"
                    
                    rfq_2124_file_list[rfq_filename] = pure_name

                else:
                    result = "24MM新规(23MM开头）"
            
            rfq_file_list[rfq_filename] = result_file + result
            rfqSheet.Cells(startRow, 9).Value = result_file + result
            startRow += 1

        rfqWb.Close(SaveChanges = 1)

        rfq24file = os.getcwd() + "\\doc\\24MM_要件分析表_是否只是版本号不一致_220211.xlsx"
        rfq24Wb  = self.excelApp.Workbooks.Open(rfq24file)
        rfq24Sheet = rfq24Wb.Sheets('软件') 
        rfqfileSheet = rfq24Wb.Sheets('文件对应') 

        i= 1
        for file in rfq_2124_file_list:
            rfqfileSheet.Cells(i, 1).Value = rfq_2124_file_list[file]
            rfqfileSheet.Cells(i, 2).Value = file
            i += 1

        '''
        startRow = 3
        info = rfq24Sheet.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count

        rfq21Sheet = rfq24Wb.Sheets('Sheet1') 
        startRow = 2
        info = rfq21Sheet.UsedRange
        nrows = info.Rows.Count

        rfq21_list = {}
        while startRow <= nrows:
            pure_file = getCellValueinString(rfq21Sheet.Cells(startRow, 12)) 
            charpter_info = getCellValueinString(rfq21Sheet.Cells(startRow, 13)).replace(" ","").replace(".","").replace("．","")
            key_str_21 = pure_file + charpter_info
            if key_str_21 in rfq21_list:
                key_value_21 = rfq21_list[key_str_21]
               
                obj_info_new =  getCellValueinString(rfq21Sheet.Cells(startRow, 15))
                imp_info_new =  getCellValueinString(rfq21Sheet.Cells(startRow, 16))

                if obj_info_new in ("该当","該当"):
                    key_value_21 = key_value_21.replace("非該当", obj_info_new)
                if imp_info_new in ("是"):
                    key_value_21 = key_value_21.replace("否", imp_info_new)
                
                rfq21_list[key_str_21] = key_value_21

            else:
                key_value = getCellValueinString(rfq21Sheet.Cells(startRow, 15)) + "|" + getCellValueinString(rfq21Sheet.Cells(startRow, 16))
                rfq21_list[pure_file + charpter_info] = key_value
            
            startRow += 1
        

        startRow = 3
        info = rfq24Sheet.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count

        while startRow <= nrows:
            rfq_filename = getCellValueinString(rfq24Sheet.Cells(startRow, 8))
            charpter_info = getCellValueinString(rfq24Sheet.Cells(startRow, 10)).replace(" ","").replace(".","").replace("．","")
            try:
                result = rfq_file_list[rfq_filename]
                rfq24Sheet.Cells(startRow, 16).Value =  result
            except Exception as e:
                rfq24Sheet.Cells(startRow, 16).Value =  "对比文件中没有此文档"
            
            if rfq_filename.upper().find("HISTORY") >= 0 or rfq_filename.upper().find("COVER") >= 0 or rfq_filename.upper().find("表紙") >= 0 or \
              rfq_filename.upper().find("CHECKSHEET") >= 0 or rfq_filename.find("チェックシート") >= 0 or rfq_filename.find("変更履歴") >= 0:
                rfq24Sheet.Cells(startRow, 24).Value = "Cover/History类，可以忽略"

            rfq_filename = getCellValueinString(rfq24Sheet.Cells(startRow, 8))
            if rfq_filename in rfq_2124_file_list:
                filename_21 = rfq_2124_file_list[rfq_filename]
                key_str = filename_21 + charpter_info
                if key_str in rfq21_list:
                    value_str = rfq21_list[key_str]
                    info_obj = value_str.split("|")[0]
                    imp_obj = value_str.split("|")[1]

                    rfq24Sheet.Cells(startRow, 26).Value = info_obj
                    rfq24Sheet.Cells(startRow, 27).Value = imp_obj
                else:
                    rfq24Sheet.Cells(startRow, 26).Value = "21没有此次文件章节信息"
            else:
                rfq24Sheet.Cells(startRow, 26).Value = "21没有此信息"
            startRow += 1
        '''
        rfq24Wb.Close(SaveChanges = 1)

    def CheckImpInfo_24MM(self):
        rfq24file = os.getcwd() + "\\doc\\24MM_要件分析表_是否只是版本号不一致_220211.xlsx"
        rfq24Wb  = self.excelApp.Workbooks.Open(rfq24file)
        rfq24Sheet = rfq24Wb.Sheets('软件') 
        rfqfileSheet = rfq24Wb.Sheets('文件对应') 
        rfq21Sheet = rfq24Wb.Sheets('21要件') 
        rfq21_filter_Sheet = rfq24Wb.Sheets('21要件_FILTER') 
        rfq24_filter_Sheet = rfq24Wb.Sheets('24要件_FILTER')

        i= 1
        startRow = 1
        info = rfqfileSheet.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count
        rfq_2124_file_list = {}
        while startRow <= nrows:
            rfq_2124_file_list[getCellValueinString(rfqfileSheet.Cells(startRow, 1))] = getCellValueinString(rfqfileSheet.Cells(startRow, 2))
            startRow += 1

        startRow = 3
        info = rfq24Sheet.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count

        startRow = 2
        info = rfq21Sheet.UsedRange
        nrows = info.Rows.Count

        rfq21_list = {}
        while startRow <= nrows:
            pure_file = getCellValueinString(rfq21Sheet.Cells(startRow, 12)) 
            charpter_info = getCellValueinString(rfq21Sheet.Cells(startRow, 13)).replace(" ","").replace(".","").replace("．","").upper()
            key_str_21 = pure_file + charpter_info
            if key_str_21 in rfq21_list:
                key_value_21 = rfq21_list[key_str_21]
               
                obj_info_new =  getCellValueinString(rfq21Sheet.Cells(startRow, 15))
                imp_info_new =  getCellValueinString(rfq21Sheet.Cells(startRow, 16))

                if obj_info_new in ("该当","該当"):
                    key_value_21 = key_value_21.replace("非該当", obj_info_new)
                if imp_info_new in ("是"):
                    key_value_21 = key_value_21.replace("否", imp_info_new)
                
                rfq21_list[key_str_21] = key_value_21

            else:
                key_value = getCellValueinString(rfq21Sheet.Cells(startRow, 15)) + "|" + getCellValueinString(rfq21Sheet.Cells(startRow, 16)) + "|" + getCellValueinString(rfq21Sheet.Cells(startRow, 1))
                rfq21_list[pure_file + charpter_info] = key_value
            
            startRow += 1      

        i = 1
        for key in rfq21_list:
            rfq21_filter_Sheet.Cells(i, 1).Value = key
            rfq21_filter_Sheet.Cells(i, 2).Value = rfq21_list[key]
            i +=1

        startRow = 3
        info = rfq24Sheet.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count
        while startRow <= nrows:
            rfq_filename = getCellValueinString(rfq24Sheet.Cells(startRow, 8))
            rfq_charpter = getCellValueinString(rfq24Sheet.Cells(startRow, 10)).replace(" ","").replace(".","").replace("．","").upper()

            rfq24_filter_Sheet.Cells(startRow, 1).Value = rfq_filename
            rfq24_filter_Sheet.Cells(startRow, 2).Value = rfq_charpter
            '''
            if rfq_filename in rfq_2124_file_list:
                filename_21 = rfq_2124_file_list[rfq_filename]
                key_str = filename_21 + rfq_charpter
                if key_str in rfq21_list:
                    value_str = rfq21_list[key_str]
                    info_obj = value_str.split("|")[0]
                    imp_obj = value_str.split("|")[1]
                    rfq_id = value_str.split("|")[2]
                     

                    rfq24Sheet.Cells(startRow, 26).Value = info_obj
                    rfq24Sheet.Cells(startRow, 27).Value = imp_obj
                    rfq24Sheet.Cells(startRow, 25).Value = rfq_id
                else:
                    rfq24Sheet.Cells(startRow, 26).Value = "21没有此次文件章节信息"
            else:
                rfq24Sheet.Cells(startRow, 26).Value = "21没有此信息"
            '''
            startRow += 1

        rfq24Wb.Close(SaveChanges = 1)

    def ExtractHistoryVersion(self, root):
        versionfile = os.getcwd() + "\\doc\\Amend_list.xlsx"
        versionWb  = self.excelApp.Workbooks.Open(versionfile)
        verSheet = versionWb.Sheets('sheet1') 
        wRow = 2

        for root, dirs,files in os.walk(root):
            for name in files: 
                if name not in (".DS_Store","Thumbs.db"):
                    print("start extract this file:", name)
                    fileName = os.path.join(root, name)
                    try:
                        file = os.getcwd() + "\\" + fileName
                        specWb  = self.excelApp.Workbooks.Open(file)
                        specSheet = specWb.Sheets("History")

                        info = specSheet.UsedRange
                        nrows = info.Rows.Count

                        startrow = 7

                        for i in range(startrow, nrows):
                            if getCellValueinString(specSheet.Cells(i,3)) == "1.01":
                                verSheet.Cells(wRow, 1).Value = getCellValueinString(specSheet.Cells(i, 3))
                                verSheet.Cells(wRow, 2).Value = getCellValueinString(specSheet.Cells(i, 10))
                                verSheet.Cells(wRow, 3).Value = getCellValueinString(specSheet.Cells(i, 16))
                                verSheet.Cells(wRow, 4).Value = getCellValueinString(specSheet.Cells(i, 17))
                                verSheet.Cells(wRow, 5).Value = getCellValueinString(specSheet.Cells(i, 25))
                                verSheet.Cells(wRow, 6).Value = getCellValueinString(specSheet.Cells(i, 33))
                                verSheet.Cells(wRow, 7).Value = getCellValueinString(specSheet.Cells(i, 40))
                                verSheet.Cells(wRow, 8).Value = getCellValueinString(specSheet.Cells(i, 44))
                                verSheet.Cells(wRow, 9).Value = fileName
                                wRow += 1
                        specWb.Close(SaveChanges = 0)

                    except Exception as e:
                        copyspecFile("errorfile", filename)

        versionWb.Close(SaveChanges = 1)

    def resetDocument(self, root):
        RTStringTablefile = os.getcwd() + "\\doc\\reset_gaeaspec.xlsx"
        rtStringTableWb  = self.excelApp.Workbooks.Open(RTStringTablefile)
        stringTableSheet = rtStringTableWb.Sheets('All Words new') 

        wRow = 2
        for root, dirs,files in os.walk(root):
            for name in files: 
                if name not in (".DS_Store","Thumbs.db"):
                    print("start extract this file:", name)
                    fileName = os.path.join(root, name)
                    try:
                        file = os.getcwd() + "\\" + fileName
                        specWb  = self.excelApp.Workbooks.Open(file)
                        sheetCount = (specWb.Worksheets.Count)

                        b_delete_func = False
                        delete_row = 0
                        for i in range(1, sheetCount + 1):
                            specSheet = specWb.Worksheets(i)
                            sheet_name = specWb.Worksheets(i).Name
                            info = specSheet.UsedRange
                            nrows = info.Rows.Count
                            row = 1
                            empty_row = 0
                            if sheet_name.upper() == "HISTORY":
                                row = 6
                                specSheet.Range("A" + str(row) + ":R" + str(row + nrows -1)).EntireRow.Delete()
                                specSheet.Cells(5,2).Value = "1"
                                specSheet.Cells(5,3).Value = "0.10"
                                specSheet.Cells(5,16).Value = "2022/6/2"
                                specSheet.Cells(5,17).Value = "SL Reset"
                            elif sheet_name.upper() != "COVER":
                                while (row <= nrows):
                                    for j in range (1, 50):
                                        specSheet.Cells(row, j).font.color = rgb_to_hex((0,0,0))
                                        specSheet.Cells(row, j).font.color = rgb_to_hex((0,0,0))

                                        if specSheet.Cells(row, j).Font.Strikethrough == True:
                                            stringTableSheet.Cells(wRow, 1).Value = fileName
                                            stringTableSheet.Cells(wRow, 2).Value = sheet_name
                                            stringTableSheet.Cells(wRow, 3).Value = "Delete in Line" + str(row)
                                            wRow += 1

                                        if not isEmptyValue(specSheet.Cells(row, j)):

                                            if getCellValueinString(specSheet.Cells(row, j)).upper().find("VER") >=0 and j >= 30:
                                                #specSheet.Cells(row, j).Value = ""
                                                specSheet.Cells(row, j).interior.color = rgb_to_hex((122,122,122))

                                            if getCellValueinString(specSheet.Cells(row, j)).upper().find("PANA") >=0 or getCellValueinString(specSheet.Cells(row, j)).upper().find("CYCLE") >=0 :
                                                stringTableSheet.Cells(wRow, 1).Value = fileName
                                                stringTableSheet.Cells(wRow, 2).Value = sheet_name
                                                stringTableSheet.Cells(wRow, 3).Value = "find Gaea in " + str(row)
                                                wRow += 1

                                            if getCellValueinString(specSheet.Cells(row, j)).upper() == "IDX":
                                                b_delete_func = True
                                                delete_row = row

                                            if getCellValueinString(specSheet.Cells(row, j)).upper() == "ID" and b_delete_func == True :
                                                specSheet.Range("A" + str(delete_row) + ":R" + str(row -1)).EntireRow.Delete()                                       
                                                b_delete_func = False
                                                delete_row = 0

                                                stringTableSheet.Cells(wRow, 1).Value = fileName
                                                stringTableSheet.Cells(wRow, 2).Value = sheet_name
                                                stringTableSheet.Cells(wRow, 3).Value = "Delete func" + str(row)
                                                wRow += 1
                                        
                                            
                                    row += 1

                        specWb.Close(SaveChanges = 1)

                    except Exception as e:
                        print(e)
                        copyspecFile(fileName, "no source files")

    def CreateBaseline(self, root):

        for root, dirs,files in os.walk(root):
            for name in files: 
                if name not in (".DS_Store","Thumbs.db"):
                    print("start extract this file:", name)
                    fileName = os.path.join(root, name)
                    try:
                        file = os.getcwd() + "\\" + fileName
                        specWb  = self.excelApp.Workbooks.Open(file)
                        specSheet = specWb.Sheets("History")

                        info = specSheet.UsedRange
                        nrows = info.Rows.Count

                        startrow = 7
                        for i in range(startrow, nrows):
                            if isEmptyValue(specSheet.Cells(i, 3)):
                                specSheet.Cells(i, 3).Value = "2.00"
                                specSheet.Cells(i, 4).Value = "ALL"
                                specSheet.Cells(i, 5).Value = "-"
                                specSheet.Cells(i, 6).Value = "-"
                                specSheet.Cells(i, 7).Value = "-"
                                specSheet.Cells(i, 8).Value = "-"
                                specSheet.Cells(i, 9).Value = "-"
                                specSheet.Cells(i, 10).Value = "2.0 Baseline"
                                specSheet.Cells(i, 11).Value = "Spec Confirm"
                                specSheet.Cells(i, 16).Value = "Spec Fix"
                                specSheet.Cells(i, 17).Value = "ALL"
                                specSheet.Cells(i, 25).Value = "-"
                                specSheet.Cells(i, 33).Value = "2.0 Baseline"
                                specSheet.Cells(i, 40).Value = "2023/1/17"
                                specSheet.Cells(i, 44).Value = "SL Hxy"

                                break

                        specWb.Close(SaveChanges = 1)

                    except Exception as e:
                        copyspecFile("errorfile", filename)

    def UpdateSpecFunc_ModuleAnalysis(self, root):
        rfqfile = os.getcwd() + "\\doc\\22TDEM_ModuleAnalysis.xlsx"
        rfqWb  = self.excelApp.Workbooks.Open(rfqfile)
        rfqSheet = rfqWb.Sheets('Modules')  # Modules PF Connectivity Diag FW MediaDL MediaFW Sensor Vehicle   UIFW
        #suzukiSheet = rfqWb.Sheets('suzuki')  # Modules PF Connectivity Diag FW MediaDL MediaFW Sensor Vehicle   UIFW

        '''
        startRow = 2
        info = suzukiSheet.UsedRange
        nrows = info.Rows.Count
        rfq_func_list = {}
        while startRow <= 445:
            rfq_prev = getCellValueinString(suzukiSheet.Cells(startRow, 2)) + getCellValueinString(suzukiSheet.Cells(startRow, 3)) 
            
            func_id_list = getCellValueinString(suzukiSheet.Cells(startRow, 4)).replace("ID","").split(",")
            for func_id in func_id_list:
                rfq_no = rfq_prev + "ID" + str(func_id)
                if rfq_no in func_id_list:
                    rfq_func_list[rfq_no] = func_id_list[rfq_no] + "\n"+ getCellValueinString(suzukiSheet.Cells(startRow, 1))
                else:
                    rfq_func_list[rfq_no] = getCellValueinString(suzukiSheet.Cells(startRow, 1))

            startRow += 1

        startRow = 20
        info = rfqSheet.UsedRange
        nrows = info.Rows.Count
        while startRow <= 1634:
            rfq_no = getCellValueinString(rfqSheet.Cells(startRow, 6)) + getCellValueinString(rfqSheet.Cells(startRow, 7)) + getCellValueinString(rfqSheet.Cells(startRow, 9))
            if rfq_no in rfq_func_list:
                rfqSheet.Cells(startRow, 14).Value = rfq_func_list[rfq_no]
            startRow += 1


        '''
        wRow = 20
        for root, dirs,files in os.walk(root):
            for name in files: 
                if name not in (".DS_Store","Thumbs.db"):
                    print("start extract this file:", name)
                    fileName = os.path.join(root, name)
                    try:
                        file = os.getcwd() + "\\" + fileName
                        specWb  = self.excelApp.Workbooks.Open(file)
                        sheetCount = (specWb.Worksheets.Count)

                        for i in range(1, sheetCount + 1):
                            specSheet = specWb.Worksheets(i)
                            sheet_name = specWb.Worksheets(i).Name
                            info = specSheet.UsedRange
                            nrows = info.Rows.Count
                            row = 1
                            if sheet_name.upper() not in ("COVER","HISTORY") and specSheet.visible == -1:
                                print(sheet_name + ":" )
                                rfqSheet.Cells(wRow , 3).Value = getCellValueinString(specSheet.Cells(2,2))                                        
                                rfqSheet.Cells(wRow, 6).Value = name
                                rfqSheet.Cells(wRow, 7).Value = sheet_name
                                wRow += 1
                                while (row < nrows + 1):
                                    id_tag = getCellValueinString(specSheet.Cells(row, 1))
                                    if id_tag.upper().replace(" ","") == "ID":
                                        spec_key = name.replace(".xlsx","") + sheet_name + id_tag + " " + getCellValueinString(specSheet.Cells(row, 2))
                                        rfqSheet.Cells(wRow , 3).Value = getCellValueinString(specSheet.Cells(2,2))                                        
                                        rfqSheet.Cells(wRow, 6).Value = name
                                        rfqSheet.Cells(wRow, 7).Value = sheet_name
                                        rfqSheet.Cells(wRow , 9).Value = id_tag + getCellValueinString(specSheet.Cells(row, 2))
                                        title_content = ""
                                        for j in range(4, 26):
                                            title_content += getCellValueinString(specSheet.Cells(row, j))
                                        rfqSheet.Cells(wRow , 10).Value = title_content.replace("Func.", "")
                                        '''
                                        if spec_key in rfq_func_list:
                                            rfqSheet.Cells(wRow , 13).Value = getCellValueinString(rfqSheet.Cells(int(rfq_func_list[spec_key]),13))
                                            rfqSheet.Cells(wRow , 14).Value = getCellValueinString(rfqSheet.Cells(int(rfq_func_list[spec_key]),14))
                                        '''
                                        wRow += 1

                                    row += 1
                        specWb.Close(SaveChanges = 0)
                    except Exception as e:
                        copyspecFile("errorfile", filename)
             
        
        rfqWb.Close(SaveChanges = 1)

    def diffFordFIP(self):
        fipfile_old = os.getcwd() + "\\res\\Ford\\Ford China 【SYNC+4.0】CX771 & CX821 IVI FIP- Release Version V2.8_20221020.xlsx"
        fipWb_old  = self.excelApp.Workbooks.Open(fipfile_old)
        
        fipfile = os.getcwd() + "\\res\\Ford\\Ford China 【SYNC+4.0】CX771 & CX821 IVI FIP- Release Version V2.9_20221026.xlsx"
        fipWb  = self.excelApp.Workbooks.Open(fipfile)

        fip_list_01 = {}
        specSheet_01 = fipWb.Sheets("01_CustomerFacing") 
        info = specSheet_01.UsedRange
        nrows = info.Rows.Count
        row = 2
        while (row <= nrows):
            if getCellValueinString(specSheet_01.Cells(row, 14)).upper().replace(" ","") == "X": 
                fip_key = getCellValueinString(specSheet_01.Cells(row, 1)) + getCellValueinString(specSheet_01.Cells(row, 5))
                fip_key = fip_key.upper().replace(" ","")
                fip_list_01[fip_key] = row
            
            row += 1

        fip_list_02 = {}
        specSheet_02 = fipWb.Sheets("02_Fundamental Function")
        info = specSheet_02.UsedRange
        nrows = info.Rows.Count
        row = 2
        while (row <= nrows):
            if getCellValueinString(specSheet_02.Cells(row, 11)).upper().replace(" ","") == "X": 
                fip_key = getCellValueinString(specSheet_02.Cells(row, 1)) + getCellValueinString(specSheet_02.Cells(row,4))
                fip_key = fip_key.upper().replace(" ","")
                fip_list_02[fip_key] = row
            
            row += 1

        fip_list_03 = {}
        specSheet_03 = fipWb.Sheets("03 BOF (CX821&&CX771  06062022)")   
        info = specSheet_03.UsedRange
        nrows = info.Rows.Count
        row = 3
        while (row <= nrows):
            if getCellValueinString(specSheet_03.Cells(row, 13)).upper().replace(" ","") == "X": 
                fip_key = getCellValueinString(specSheet_03.Cells(row, 1)) + getCellValueinString(specSheet_03.Cells(row, 3))
                fip_key = fip_key.upper().replace(" ","")
                fip_list_03[fip_key] = row
            
            row += 1

        # open old version
        fip_old_list_01 = {}
        specSheet_old_01 = fipWb_old.Sheets("01_CustomerFacing") 
        info = specSheet_old_01.UsedRange
        nrows = info.Rows.Count
        row = 2
        while (row <= nrows):
            if getCellValueinString(specSheet_old_01.Cells(row, 14)).upper().replace(" ","") == "X": 
                fip_old_key = getCellValueinString(specSheet_old_01.Cells(row, 1)) + getCellValueinString(specSheet_old_01.Cells(row, 5))
                fip_old_key = fip_old_key.upper().replace(" ","")
                if fip_old_key in fip_list_01:
                    fip_old_long_str = ""
                    fip_long_str = ""
                    for i in range(1,27):
                        fip_old_long_str += getCellValueinString(specSheet_old_01.Cells(row, i)).upper().replace(" ","")
                        fip_long_str += getCellValueinString(specSheet_01.Cells(fip_list_01[fip_old_key], i)).upper().replace(" ","")
                    
                    if fip_old_long_str == fip_long_str:
                        del fip_list_01[fip_old_key] 
                else:
                    fip_old_list_01[fip_old_key] = row
            row += 1

        fip_old_list_02 = {}
        specSheet_old_02 = fipWb_old.Sheets("02_Fundamental Function") 
        info = specSheet_old_02.UsedRange
        nrows = info.Rows.Count
        row = 2
        while (row <= nrows):
            if getCellValueinString(specSheet_old_02.Cells(row, 11)).upper().replace(" ","") == "X": 
                fip_old_key = getCellValueinString(specSheet_old_02.Cells(row, 1)) + getCellValueinString(specSheet_old_02.Cells(row, 4))
                fip_old_key = fip_old_key.upper().replace(" ","")
                if fip_old_key in fip_list_02:
                    fip_old_long_str = ""
                    fip_long_str = ""
                    for i in range(1,20):
                        fip_old_long_str += getCellValueinString(specSheet_old_02.Cells(row, i)).upper().replace(" ","")
                        fip_long_str += getCellValueinString(specSheet_02.Cells(fip_list_02[fip_old_key], i)).upper().replace(" ","")
                    
                    if fip_old_long_str == fip_long_str:
                        del fip_list_02[fip_old_key] 
                else:
                    fip_old_list_02[fip_old_key] = row
            row += 1

        fip_old_list_03 = {}
        specSheet_old_03 = fipWb_old.Sheets("03 BOF (CX821&&CX771  06062022)") 
        info = specSheet_old_03.UsedRange
        nrows = info.Rows.Count
        row = 3
        while (row <= nrows):
            if getCellValueinString(specSheet_old_03.Cells(row, 13)).upper().replace(" ","") == "X": 
                fip_old_key = getCellValueinString(specSheet_old_03.Cells(row, 1)) + getCellValueinString(specSheet_old_03.Cells(row, 3))
                fip_old_key = fip_old_key.upper().replace(" ","")
                if fip_old_key in fip_list_03:
                    fip_old_long_str = ""
                    fip_long_str = ""
                    for i in range(1,32):
                        fip_old_long_str += getCellValueinString(specSheet_old_03.Cells(row, i)).upper().replace(" ","")
                        fip_long_str += getCellValueinString(specSheet_03.Cells(fip_list_03[fip_old_key], i)).upper().replace(" ","")
                    
                    if fip_old_long_str == fip_long_str:
                        del fip_list_03[fip_old_key] 
                else:
                    fip_old_list_03[fip_old_key] = row
            row += 1


        featureListfile = os.getcwd() + "\\res\\Ford\\Ford China 【SYNC+4.0】CX771 & CX821 IVI FeatureList.xlsx"
        featureListWb  = self.excelApp.Workbooks.Open(featureListfile)

        featureListSheet = featureListWb.Sheets('FeatureList') 

        startRow = 3
        info = featureListSheet.UsedRange
        nrows = info.Rows.Count
        feature_list = {}
        while startRow <= nrows:
            feature_key = getCellValueinString(featureListSheet.Cells(startRow, 4)) + getCellValueinString(featureListSheet.Cells(startRow, 5))
            feature_key = feature_key.upper().replace(" ","")
            if feature_key in fip_list_01: 
                featureListSheet.Cells(startRow,1).interior.color = rgb_to_hex(AMEND_COLOR)
                if not isEmptyValue(featureListSheet.Cells(startRow,1)):
                    featureListSheet.Cells(startRow,1).Value = "V0.42 " + getCellValueinString(featureListSheet.Cells(startRow,1))
                else:
                    featureListSheet.Cells(startRow,1).Value = "V0.42" 
                
                #Ford_Feature_Group
                if getCellValueinString(featureListSheet.Cells(startRow, 3)) != getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 3)):
                    featureListSheet.Cells(startRow,3).Value = getCellValueinString(featureListSheet.Cells(startRow,3)) + "{" + getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 3)) + "}" 
                
                #Ford_Feature_ID
                if getCellValueinString(featureListSheet.Cells(startRow, 4)) != getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 1)):
                    featureListSheet.Cells(startRow,4).Value = getCellValueinString(featureListSheet.Cells(startRow,4)) + "{" + getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 1)) + "}" 
                
                #Feature_Name
                if getCellValueinString(featureListSheet.Cells(startRow, 5)) != getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 5)):
                    featureListSheet.Cells(startRow, 5).Value = getCellValueinString(featureListSheet.Cells(startRow, 5)) + "{" + getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 5)) + "}" 
                
                #Feature_Description
                if getCellValueinString(featureListSheet.Cells(startRow, 6)) != getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 10)):
                    featureListSheet.Cells(startRow, 6).Value = getCellValueinString(featureListSheet.Cells(startRow, 6)) + "{" + getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 10)) + "}" 
                
                #Dev_Spec
                if getCellValueinString(featureListSheet.Cells(startRow, 10)) != getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 11)):
                    featureListSheet.Cells(startRow, 10).Value = getCellValueinString(featureListSheet.Cells(startRow, 10)) + "{" + getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 11)) + "}" 

                #CX771
                if getCellValueinString(featureListSheet.Cells(startRow, 12)) != getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 24)):
                    featureListSheet.Cells(startRow, 12).Value = getCellValueinString(featureListSheet.Cells(startRow, 12)) + "{" + getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 24)) + "}" 

                #CX821
                if getCellValueinString(featureListSheet.Cells(startRow, 13)) != getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 25)):
                    featureListSheet.Cells(startRow, 13).Value = getCellValueinString(featureListSheet.Cells(startRow, 13)) + "{" + getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 25)) + "}" 

                #Feature Development Leader
                if getCellValueinString(featureListSheet.Cells(startRow, 14)) != getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 12)):
                    featureListSheet.Cells(startRow, 14).Value = getCellValueinString(featureListSheet.Cells(startRow, 14)) + "{" + getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 12)) + "}" 

                #Package1
                if getCellValueinString(featureListSheet.Cells(startRow, 15)) != getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 13)):
                    featureListSheet.Cells(startRow, 15).Value = getCellValueinString(featureListSheet.Cells(startRow, 15)) + "{" + getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 13)) + "}" 

                #Package2
                if getCellValueinString(featureListSheet.Cells(startRow, 16)) != getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 14)):
                    featureListSheet.Cells(startRow, 16).Value = getCellValueinString(featureListSheet.Cells(startRow, 16)) + "{" + getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 14)) + "}" 

                #Package3
                if getCellValueinString(featureListSheet.Cells(startRow, 17)) != getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 15)):
                    featureListSheet.Cells(startRow, 17).Value = getCellValueinString(featureListSheet.Cells(startRow, 17)) + "{" + getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 15)) + "}" 

                #Inhouse APP
                if getCellValueinString(featureListSheet.Cells(startRow, 18)) != getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 16)):
                    featureListSheet.Cells(startRow, 18).Value = getCellValueinString(featureListSheet.Cells(startRow, 18)) + "{" + getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 16)) + "}" 

                #Baidu
                if getCellValueinString(featureListSheet.Cells(startRow, 19)) != getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 17)):
                    featureListSheet.Cells(startRow, 19).Value = getCellValueinString(featureListSheet.Cells(startRow, 19)) + "{" + getCellValueinString(specSheet_01.Cells(fip_list_01[feature_key], 17)) + "}" 

                del fip_list_01[feature_key]
            
            if feature_key in fip_list_02: 
                featureListSheet.Cells(startRow,1).interior.color = rgb_to_hex(AMEND_COLOR)
                if not isEmptyValue(featureListSheet.Cells(startRow,1)):
                    featureListSheet.Cells(startRow,1).Value = "V0.42 " + getCellValueinString(featureListSheet.Cells(startRow,1))
                else:
                    featureListSheet.Cells(startRow,1).Value = "V0.42" 

                #Ford_Feature_Group
                if getCellValueinString(featureListSheet.Cells(startRow, 3)) != getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 3)):
                    featureListSheet.Cells(startRow,3).Value = getCellValueinString(featureListSheet.Cells(startRow,3)) + "{" + getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 3)) + "}" 

                #Ford_Feature_ID
                if getCellValueinString(featureListSheet.Cells(startRow, 4)) != getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 1)):
                    featureListSheet.Cells(startRow,4).Value = getCellValueinString(featureListSheet.Cells(startRow,4)) + "{" + getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 1)) + "}" 

                #Feature_Name
                if getCellValueinString(featureListSheet.Cells(startRow, 5)) != getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 4)):
                    featureListSheet.Cells(startRow, 5).Value = getCellValueinString(featureListSheet.Cells(startRow,5)) + "{" + getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 4)) + "}" 

                #Feature_Description
                if getCellValueinString(featureListSheet.Cells(startRow, 6)) != getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 5)):
                    featureListSheet.Cells(startRow, 6).Value = getCellValueinString(featureListSheet.Cells(startRow,6)) + "{" + getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 5)) + "}" 

                #Dev_Spec
                if getCellValueinString(featureListSheet.Cells(startRow, 10)) != getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 8)):
                    featureListSheet.Cells(startRow, 10).Value = getCellValueinString(featureListSheet.Cells(startRow, 10)) + "{" + getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 8)) + "}" 

                #CX771
                if getCellValueinString(featureListSheet.Cells(startRow, 12)) != getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 15)):
                    featureListSheet.Cells(startRow, 12).Value = getCellValueinString(featureListSheet.Cells(startRow, 12)) + "{" + getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 15)) + "}" 

                #CX821
                if getCellValueinString(featureListSheet.Cells(startRow, 13)) != getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 16)):
                    featureListSheet.Cells(startRow, 13).Value = getCellValueinString(featureListSheet.Cells(startRow, 13)) + "{" + getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 16)) + "}" 

                #Feature Development Leader
                if getCellValueinString(featureListSheet.Cells(startRow, 14)) != getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 9)):
                    featureListSheet.Cells(startRow, 14).Value = getCellValueinString(featureListSheet.Cells(startRow, 14)) + "{" + getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 9)) + "}" 

                #Package1
                if getCellValueinString(featureListSheet.Cells(startRow, 15)) != getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 10)):
                    featureListSheet.Cells(startRow, 15).Value = getCellValueinString(featureListSheet.Cells(startRow, 15)) + "{" + getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 10)) + "}" 

                #Package2
                if getCellValueinString(featureListSheet.Cells(startRow, 16)) != getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 11)):
                    featureListSheet.Cells(startRow, 16).Value = getCellValueinString(featureListSheet.Cells(startRow, 16)) + "{" + getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 11)) + "}" 

                #Package3
                if getCellValueinString(featureListSheet.Cells(startRow, 17)) != getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 12)):
                    featureListSheet.Cells(startRow, 17).Value = getCellValueinString(featureListSheet.Cells(startRow, 17)) + "{" + getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 12)) + "}" 

                #Baidu                
                if getCellValueinString(featureListSheet.Cells(startRow, 19)) != getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 13)):
                    featureListSheet.Cells(startRow, 19).Value = getCellValueinString(featureListSheet.Cells(startRow, 19)) + "{" + getCellValueinString(specSheet_02.Cells(fip_list_02[feature_key], 13)) + "}" 

               
                del fip_list_02[feature_key]
            
            if feature_key in fip_list_03: 
                featureListSheet.Cells(startRow,1).interior.color = rgb_to_hex(AMEND_COLOR)
                if not isEmptyValue(featureListSheet.Cells(startRow,1)):
                    featureListSheet.Cells(startRow,1).Value = "V0.42 " + getCellValueinString(featureListSheet.Cells(startRow,1))
                else:
                    featureListSheet.Cells(startRow,1).Value = "V0.42" 
                #Ford_Feature_ID
                if getCellValueinString(featureListSheet.Cells(startRow, 4)) != getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 1)):
                    featureListSheet.Cells(startRow,4).Value = getCellValueinString(featureListSheet.Cells(startRow,4)) + "{" + getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 1)) + "}" 

                #Feature_Name
                if getCellValueinString(featureListSheet.Cells(startRow, 5)) != getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 3)):
                    featureListSheet.Cells(startRow, 5).Value = getCellValueinString(featureListSheet.Cells(startRow,5)) + "{" + getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 3)) + "}" 

                #Feature_Description
                if getCellValueinString(featureListSheet.Cells(startRow, 6)) != getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 7)):
                    featureListSheet.Cells(startRow, 6).Value = getCellValueinString(featureListSheet.Cells(startRow, 6)) + "{" + getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 7)) + "}" 

                #Dev_Spec
                if getCellValueinString(featureListSheet.Cells(startRow, 10)) != getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 9)):
                    featureListSheet.Cells(startRow, 10).Value = getCellValueinString(featureListSheet.Cells(startRow, 10)) + "{" + getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 9)) + "}" 

                #CX771
                if getCellValueinString(featureListSheet.Cells(startRow, 12)) != getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 28)):
                    featureListSheet.Cells(startRow, 12).Value = getCellValueinString(featureListSheet.Cells(startRow, 12)) + "{" + getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 28)) + "}" 

                #CX821
                if getCellValueinString(featureListSheet.Cells(startRow, 13)) != getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 29)):
                    featureListSheet.Cells(startRow, 13).Value = getCellValueinString(featureListSheet.Cells(startRow, 13)) + "{" + getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 29)) + "}" 

                #Feature Development Leader
                if getCellValueinString(featureListSheet.Cells(startRow, 14)) != getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 11)):
                    featureListSheet.Cells(startRow, 14).Value = getCellValueinString(featureListSheet.Cells(startRow, 14)) + "{" + getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 11)) + "}" 

                #Package1
                if getCellValueinString(featureListSheet.Cells(startRow, 15)) != getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 12)):
                    featureListSheet.Cells(startRow, 15).Value = getCellValueinString(featureListSheet.Cells(startRow, 15)) + "{" + getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 12)) + "}" 

                #Package2
                if getCellValueinString(featureListSheet.Cells(startRow, 16)) != getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 13)):
                    featureListSheet.Cells(startRow, 16).Value = getCellValueinString(featureListSheet.Cells(startRow, 16)) + "{" + getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 13)) + "}" 

                #Package3
                if getCellValueinString(featureListSheet.Cells(startRow, 17)) != getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 14)):
                    featureListSheet.Cells(startRow, 17).Value = getCellValueinString(featureListSheet.Cells(startRow, 17)) + "{" + getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 14)) + "}" 

                #Inhouse APP
                if getCellValueinString(featureListSheet.Cells(startRow, 18)) != getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 15)):
                    featureListSheet.Cells(startRow, 18).Value = getCellValueinString(featureListSheet.Cells(startRow, 18)) + "{" + getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 15)) + "}" 

                #Baidu
                if getCellValueinString(featureListSheet.Cells(startRow, 19)) != getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 16)):
                    featureListSheet.Cells(startRow, 19).Value = getCellValueinString(featureListSheet.Cells(startRow, 19)) + "{" + getCellValueinString(specSheet_03.Cells(fip_list_03[feature_key], 15)) + "}" 

                del fip_list_03[feature_key]

            if feature_key in fip_old_list_01:
                if not isEmptyValue(featureListSheet.Cells(startRow,1)):
                    featureListSheet.Cells(startRow,1).Value = "V0.42 " + getCellValueinString(featureListSheet.Cells(startRow,1))
                else:
                    featureListSheet.Cells(startRow,1).Value = "V0.42" 

                featureListSheet.Cells(startRow,1).interior.color = rgb_to_hex(REMOVE_COLOR)
                del fip_old_list_01[feature_key]

            if feature_key in fip_old_list_02:
                if not isEmptyValue(featureListSheet.Cells(startRow,1)):
                    featureListSheet.Cells(startRow,1).Value = "V0.42 " + getCellValueinString(featureListSheet.Cells(startRow,1))
                else:
                    featureListSheet.Cells(startRow,1).Value = "V0.42" 

                featureListSheet.Cells(startRow,1).interior.color = rgb_to_hex(REMOVE_COLOR)    
                del fip_old_list_02[feature_key]

            if feature_key in fip_old_list_03:
                if not isEmptyValue(featureListSheet.Cells(startRow,1)):
                    featureListSheet.Cells(startRow,1).Value = "V0.42 " + getCellValueinString(featureListSheet.Cells(startRow,1))
                else:
                    featureListSheet.Cells(startRow,1).Value = "V0.42" 

                featureListSheet.Cells(startRow,1).interior.color = rgb_to_hex(REMOVE_COLOR)
                del fip_old_list_03[feature_key]
            
            startRow +=1
            
        wRow = 150
        for key in fip_list_01:
            featureListSheet.Cells(wRow,1).interior.color = rgb_to_hex(AMEND_COLOR)
            featureListSheet.Cells(wRow,1).Value = "V0.42" 

            #Ford_Feature_Group
            featureListSheet.Cells(wRow,3).Value = getCellValueinString(specSheet_01.Cells(fip_list_01[key],3)) #Feature_Group
            #Ford_Feature_ID
            featureListSheet.Cells(wRow,4).Value = getCellValueinString(specSheet_01.Cells(fip_list_01[key],1))  #Feature_ID
            #Feature_Name
            featureListSheet.Cells(wRow,5).Value = getCellValueinString(specSheet_01.Cells(fip_list_01[key],5))  #Feature_Name
            #Feature_Description
            featureListSheet.Cells(wRow,6).Value = getCellValueinString(specSheet_01.Cells(fip_list_01[key],10))  #Feature_Description 
            #Dev_Spec
            featureListSheet.Cells(wRow,10).Value = getCellValueinString(specSheet_01.Cells(fip_list_01[key],11)) #Dev Spec infor
            #CX771
            featureListSheet.Cells(wRow,12).Value = getCellValueinString(specSheet_01.Cells(fip_list_01[key],24)) #CX771 FNV2.1 Mar-2024
            #CX821
            featureListSheet.Cells(wRow,13).Value = getCellValueinString(specSheet_01.Cells(fip_list_01[key],25)) #CX821 FNV2.1 Mar-2024
            #Feature Development Leader
            featureListSheet.Cells(wRow,14).Value = getCellValueinString(specSheet_01.Cells(fip_list_01[key],12)) #Feature Development Leader
            #Package1
            featureListSheet.Cells(wRow,15).Value = getCellValueinString(specSheet_01.Cells(fip_list_01[key],13)) #Package1
            #Package2
            featureListSheet.Cells(wRow,16).Value = getCellValueinString(specSheet_01.Cells(fip_list_01[key],14)) #Package2
            #Package3
            featureListSheet.Cells(wRow,17).Value = getCellValueinString(specSheet_01.Cells(fip_list_01[key],15)) #Package3
            #Inhouse APP
            featureListSheet.Cells(wRow,18).Value = getCellValueinString(specSheet_01.Cells(fip_list_01[key],16)) #Inhouse APP
            #Baidu
            featureListSheet.Cells(wRow,19).Value = getCellValueinString(specSheet_01.Cells(fip_list_01[key],17)) #Potential Baidu
            wRow += 1

        for key in fip_list_02:
            featureListSheet.Cells(wRow,1).interior.color = rgb_to_hex(AMEND_COLOR)
            featureListSheet.Cells(wRow,1).Value = "V0.42" 

            #Ford_Feature_Group
            featureListSheet.Cells(wRow,3).Value = getCellValueinString(specSheet_02.Cells(fip_list_02[key],3)) #Feature Group
            #Ford_Feature_ID
            featureListSheet.Cells(wRow,4).Value = getCellValueinString(specSheet_02.Cells(fip_list_02[key],1)) #Feature ID
            #Feature_Name
            featureListSheet.Cells(wRow,5).Value = getCellValueinString(specSheet_02.Cells(fip_list_02[key],4)) #Feature Name
            #Feature_Description 
            featureListSheet.Cells(wRow,6).Value = getCellValueinString(specSheet_02.Cells(fip_list_02[key],5)) #Feature_Description 
            #Dev_Spec
            featureListSheet.Cells(wRow,10).Value = getCellValueinString(specSheet_02.Cells(fip_list_02[key],8)) #Dev_Spec
            #CX771
            featureListSheet.Cells(wRow,12).Value = getCellValueinString(specSheet_02.Cells(fip_list_02[key],15)) #CX771
            #CX821
            featureListSheet.Cells(wRow,13).Value = getCellValueinString(specSheet_02.Cells(fip_list_02[key],16)) #CX821            
            #Feature Development Leader
            featureListSheet.Cells(wRow,14).Value = getCellValueinString(specSheet_02.Cells(fip_list_02[key],9)) #Feature Development Leader
            #Package1
            featureListSheet.Cells(wRow,15).Value = getCellValueinString(specSheet_02.Cells(fip_list_02[key],10)) #Package1
            #Package2
            featureListSheet.Cells(wRow,16).Value = getCellValueinString(specSheet_02.Cells(fip_list_02[key],11)) #Package2
            #Package3
            featureListSheet.Cells(wRow,17).Value = getCellValueinString(specSheet_02.Cells(fip_list_02[key],12)) #Package3
            #Inhouse APP 
            featureListSheet.Cells(wRow,18).Value = ""
            #Baidu
            featureListSheet.Cells(wRow,19).Value = getCellValueinString(specSheet_02.Cells(fip_list_02[key],13)) #Potential Baidu

            wRow += 1              

        for key in fip_list_03:
            featureListSheet.Cells(wRow,1).interior.color = rgb_to_hex(AMEND_COLOR)
            featureListSheet.Cells(wRow,1).Value = "V0.42"             
            #Ford_Feature_Group
            featureListSheet.Cells(wRow,3).Value = ""
            #Ford_Feature_ID
            featureListSheet.Cells(wRow,4).Value = getCellValueinString(specSheet_03.Cells(fip_list_03[key],1)) #Key
            #Feature_Name
            featureListSheet.Cells(wRow,5).Value = getCellValueinString(specSheet_03.Cells(fip_list_03[key],3)) #Summary
            #Feature_Description
            featureListSheet.Cells(wRow,6).Value = getCellValueinString(specSheet_03.Cells(fip_list_03[key],7)) #Description
            #Dev_Spec
            featureListSheet.Cells(wRow,10).Value = getCellValueinString(specSheet_03.Cells(fip_list_03[key],9)) #Feature Spec Name
            #CX771
            featureListSheet.Cells(wRow,12).Value = getCellValueinString(specSheet_03.Cells(fip_list_03[key],28)) #CX771
            #CX821
            featureListSheet.Cells(wRow,13).Value = getCellValueinString(specSheet_03.Cells(fip_list_03[key],29)) #CX821
            #Feature Development Leader
            featureListSheet.Cells(wRow,14).Value = getCellValueinString(specSheet_03.Cells(fip_list_03[key],11)) #Feature Development Leader
            #Package1
            featureListSheet.Cells(wRow,15).Value = getCellValueinString(specSheet_03.Cells(fip_list_03[key],12)) #Package1
            #Package2
            featureListSheet.Cells(wRow,16).Value = getCellValueinString(specSheet_03.Cells(fip_list_03[key],13)) #Package2
            #Package3
            featureListSheet.Cells(wRow,17).Value = getCellValueinString(specSheet_03.Cells(fip_list_03[key],14)) #Package3
            #Inhouse APP
            featureListSheet.Cells(wRow,18).Value = getCellValueinString(specSheet_03.Cells(fip_list_03[key],15)) #Inhouse APP
            #Baidu
            featureListSheet.Cells(wRow,19).Value = getCellValueinString(specSheet_03.Cells(fip_list_03[key],16)) #Potential Baidu
            wRow += 1
        
        featureListWb.Close(SaveChanges = 1)

    def UpdateFordInfo(self):
        fipfile_old = os.getcwd() + "\\res\\Ford\\发布计划日程_FordSync+4.0_20220801.xlsx"
        fipWb_old  = self.excelApp.Workbooks.Open(fipfile_old)
        fip_milestone_sheet = fipWb_old.Sheets('sheet1') 
        
        featureListfile = os.getcwd() + "\\res\\Ford\\Ford China 【SYNC+4.0】CX771 & CX821 IVI FeatureList.xlsx"
        featureListWb  = self.excelApp.Workbooks.Open(featureListfile)

        featureListSheet = featureListWb.Sheets('FeatureList') 

        startRow = 2
        info = fip_milestone_sheet.UsedRange
        nrows = info.Rows.Count
        milestone_list = {}
        while startRow <= nrows:
            milestone_list[getCellValueinString(fip_milestone_sheet.Cells(startRow, 3)).replace(" ","").upper()] = getCellValueinString(fip_milestone_sheet.Cells(startRow, 4))
            startRow += 1


        startRow = 3
        info = featureListSheet.UsedRange
        nrows = info.Rows.Count
        while startRow <= nrows:
            if getCellValueinString(featureListSheet.Cells(startRow, 5)).replace(" ","").upper() in milestone_list:
                featureListSheet.Cells(startRow,21).Value = milestone_list[getCellValueinString(featureListSheet.Cells(startRow, 5)).replace(" ","").upper()] 
            startRow += 1

        featureListWb.Close(SaveChanges = 1)

    def DiffDeveloperInfo(self):
        fipfile_old = os.getcwd() + "\\res\\Ford\\CX821 IVI Software Maturity Plan_v1.0_20220913-PKG2-0930.xlsx"
        fipWb_old  = self.excelApp.Workbooks.Open(fipfile_old)
        fip_milestone_sheet = fipWb_old.Sheets('(UPV0) IVI & Cluster') 
        
        featureListfile = os.getcwd() + "\\res\\Ford\\Ford China 【SYNC+4.0】CX771 & CX821 IVI FeatureList.xlsx"
        featureListWb  = self.excelApp.Workbooks.Open(featureListfile)
        featureListSheet = featureListWb.Sheets('FeatureList') 

        startRow = 21
        info = fip_milestone_sheet.UsedRange
        nrows = info.Rows.Count
        devInfo_list = {}
        while startRow <= 140:
            if not isEmptyValue(fip_milestone_sheet.Cells(startRow,9)):
                devInfo_list[getCellValueinString(fip_milestone_sheet.Cells(startRow, 6)).replace(" ","").upper()] = getCellValueinString(fip_milestone_sheet.Cells(startRow, 9))
            startRow += 1

        startRow = 16
        while startRow <= 140:
            if not isEmptyValue(fip_milestone_sheet.Cells(startRow,19)):
                devInfo_list[getCellValueinString(fip_milestone_sheet.Cells(startRow, 16)).replace(" ","").upper()] = getCellValueinString(fip_milestone_sheet.Cells(startRow, 19))
            startRow += 1


        startRow = 3
        info = featureListSheet.UsedRange
        nrows = info.Rows.Count
        while startRow <= nrows:
            dev_key = getCellValueinString(featureListSheet.Cells(startRow, 5)).replace(" ","").upper()
            if dev_key in devInfo_list:
                if getCellValueinString(featureListSheet.Cells(startRow, 21)) != devInfo_list[dev_key] :
                    featureListSheet.Cells(startRow, 21).interior.color = rgb_to_hex((84,129,53))
                    if not isEmptyValue(featureListSheet.Cells(startRow, 21)) : 
                        featureListSheet.Cells(startRow, 21).Value = getCellValueinString(featureListSheet.Cells(startRow, 21)) + "{" + devInfo_list[dev_key] + "}"
                    else:
                        featureListSheet.Cells(startRow, 21).Value = "{" + devInfo_list[dev_key] + "}"
                

                del devInfo_list[dev_key]
            startRow += 1

        featureListWb.Close(SaveChanges = 1)

    def UpdateESCForFraser(self):
        screen_list = [
            "MM_02_01_11",
            "MM_02_01_12",
            "MM_03_11_13",
            "MM_03_11_14",
            "MM_03_11_15",
            "MM_08_01_01",
            "MM_08_01_02",
            "MM_08_01_03",
            "MM_08_01_05",
            "MM_08_01_10",
            "MM_08_01_11",
            "MM_08_01_12",
            "MM_08_01_13",
            "MM_08_01_14",
            "MM_08_01_15",
            "MM_08_01_27",
            "MM_08_01_28",
            "MM_08_01_29",
            "MM_08_01_31",
            "MM_08_01_32",
            "MM_08_01_33",
            "MM_08_01_34",
            "MM_08_01_35",
            "MM_08_01_37",
            "MM_08_01_38",
            "MM_08_01_39",
            "MM_08_01_40",
            "MM_08_01_41",
            "MM_08_01_42",
            "MM_08_01_43",
            "MM_08_01_44",
            "MM_08_02_02",
            "MM_08_02_03",
            "MM_08_02_04",
            "MM_08_02_05",
            "MM_08_02_06",
            "MM_08_02_07",
            "MM_08_02_08",
            "MM_08_02_09",
            "MM_08_02_11",
            "MM_08_02_14",
            "MM_08_04_02",
            "MM_08_04_03",
            "MM_08_04_04",
            "MM_08_04_05",
            "MM_08_04_06",
            "MM_08_05_01",
            "MM_08_05_04",
            "MM_08_05_05",
            "MM_08_08_01",
            "MM_08_11_01",
            "MM_08_12_01",
            "MM_08_12_02",
            "MM_08_13_01",
            "MM_08_13_02",
            "MM_08_16_01",
            "MM_08_16_02",
            "MM_08_16_03",
            "MM_08_16_04",
            "MM_08_16_05",
            "MM_08_16_06",
            "MM_08_16_07",
            "MM_08_16_09",
            "MM_08_16_10",
            "MM_08_18_01",
            "MM_08_18_02",
            "MM_08_18_03",
            "MM_08_18_04",
            "MM_08_20_01",
            "MM_08_20_02",
            "MM_08_20_03",
            "MM_08_22_01",
            "MM_08_22_02",
            "MM_08_22_03",
            "MM_08_22_04",
            "MM_08_22_05",
            "MM_08_23_01",
            "MM_08_23_02",
            "MM_08_23_03",
            "MM_08_23_04",
            "MM_08_23_05",
            "MM_08_23_06",
            "MM_08_24_01",
            "MM_08_24_02",
            "MM_08_24_03",
            "MM_08_24_04",
            "MM_08_24_05",
            "MM_10_01_29" 
        ]
        
        for root, dirs, files in os.walk("res\\remin"):
            for name in files: 
                if name not in (".DS_Store","Thumbs.db"):
                    print("start extract this file:", name)
                    fileName = os.path.join(root, name)
                    try:
                        opeTableWb  = self.excelApp.Workbooks.Open(os.getcwd() + "\\" +  fileName)
                        historySheet = None
                        try:
                            historySheet = opeTableWb.Sheets("History")

                            info = historySheet.UsedRange
                            nrows = info.Rows.Count
                            row = 5
                            last_history_row = 0
                            while (row <= nrows):
                                if isEmptyValue(historySheet.Cells(row, 3)):
                                    last_history_row = row
                                    break
                                row += 1
                        
                        except Exception as e: 
                            print("No history in this document")

                        sheetCount = opeTableWb.Worksheets.Count
                        bChanged = 0
                        for i in range(1, sheetCount + 1):
                            opeSheet = opeTableWb.Worksheets(i)
                            sheet_name = opeSheet.Name
            
                            if sheet_name in screen_list:
                                info = opeSheet.UsedRange
                                nrows = info.Rows.Count
                                row = 2
                                esc_remark_row = 0
                                esc_id = "\n"
                                max_number = 0
                                while (row <= nrows):
                                    if getCellValueinString(opeSheet.Cells(row, 2)).upper() == "3" and \
                                        (getCellValueinString(opeSheet.Cells(row, 5)).upper().replace(" ","") in ("ESC","閉じる") or getCellValueinString(opeSheet.Cells(row, 6)).upper().replace(" ","") in ("ESC","閉じる")):
                                        esc_remark_row = row
                                        esc_id += getCellValueinString(opeSheet.Cells(esc_remark_row,2))
                                        for j in range(3,5):
                                            esc_id += "-" + getCellValueinString(opeSheet.Cells(esc_remark_row,j))
                                    if getCellValueinString(opeSheet.Cells(row, 2)).upper() == "4":
                                        if getCellValueinString(opeSheet.Cells(row-1, 3)).find("(#)3-") >= 0:
                                            max_number = int(getCellValueinString(opeSheet.Cells(row-1, 3))[5:6])
                                        else:
                                            max_number = 0
                                        opeSheet.Cells(esc_remark_row, 46).Value = "(#)3-" + str(max_number + 1)
                                        opeSheet.Cells(esc_remark_row, 46).font.color = rgb_to_hex((255,25,255))
                                        opeSheet.Cells(esc_remark_row, 46).font.color = rgb_to_hex((255,25,255))

                                        opeSheet.Cells(esc_remark_row, 46).Font.Strikethrough = False
                                        range_a = "A" + str(row) + ":DA" + str(row) 
                                        opeSheet.Range(range_a).EntireRow.Insert()
                                        opeSheet.Cells(row, 3).Value = "(#)3-" + str(max_number + 1) + ":<操作仕様書:MM_01_02_08  [ESC]ボタン>を参照。"
                                        opeSheet.Cells(row, 3).font.color = rgb_to_hex((255,25,255))
                                        opeSheet.Cells(row, 3).font.color = rgb_to_hex((255,25,255))

                                        opeSheet.Cells(row, 3).Font.Strikethrough = False
                                        opeSheet.Range(range_a).interior.color = rgb_to_hex((255,255,255))
                                        row += 1
                                        bChanged = 1
                                        break
                                    row += 1

                                historySheet.Cells(last_history_row, 3).Value = "1.09"
                                historySheet.Cells(last_history_row, 4).Value = "-"
                                historySheet.Cells(last_history_row, 5).Value = "-"
                                historySheet.Cells(last_history_row, 6).Value = "-"
                                historySheet.Cells(last_history_row, 7).Value = "仕様変更(仕様変更管理)"
                                historySheet.Cells(last_history_row, 8).Value = "AV-NAVIの閉じるボタン仕様統一"
                                historySheet.Cells(last_history_row, 9).Value = sheet_name + esc_id
                                historySheet.Cells(last_history_row, 10).Value = "-"
                                historySheet.Cells(last_history_row, 11).Value = "Remark 追加：\n" +  "(#)3-" + str(max_number + 1)
                                historySheet.Cells(last_history_row, 12).Value = "2022/11/1"
                                historySheet.Cells(last_history_row, 13).Value = "PSET Huang"
                                historySheet.Cells(last_history_row, 15).Value = "-"
                                last_history_row += 1
                                
                                historySheet.Cells(last_history_row, 3).Value = "1.09"
                                historySheet.Cells(last_history_row, 4).Value = "-"
                                historySheet.Cells(last_history_row, 5).Value = "-"
                                historySheet.Cells(last_history_row, 6).Value = "-"
                                historySheet.Cells(last_history_row, 7).Value = "仕様変更(仕様変更管理)"
                                historySheet.Cells(last_history_row, 8).Value = "AV-NAVIの閉じるボタン仕様統一"
                                historySheet.Cells(last_history_row, 9).Value = sheet_name + "\n(#)3-" + str(max_number + 1)
                                historySheet.Cells(last_history_row, 10).Value = "-"
                                historySheet.Cells(last_history_row, 11).Value = "追加：\n" + getCellValueinString(opeSheet.Cells(row-1, 3))
                                historySheet.Cells(last_history_row, 12).Value = "2022/11/1"
                                historySheet.Cells(last_history_row, 13).Value = "PSET Huang"
                                historySheet.Cells(last_history_row, 15).Value = "-"
                                last_history_row += 1

                                screen_list.remove(sheet_name)

                        opeTableWb.Close(SaveChanges = bChanged)
                    except Exception as e: 
                        print("File error")

    def extractIFSpec_history(self):
        wRow = 1
        self.resultBook.ActiveSheet.Name = "Screen list"
        resultSheet = self.resultBook.ActiveSheet
        for root, dirs, files in os.walk("res\\lomami"):
            for name in files: 
                if name not in (".DS_Store","Thumbs.db"):
                    print("start extract this file:", name)
                    try:
                        fileName = os.path.join(root, name)
                        IFSpecWb  = self.excelApp.Workbooks.Open(os.getcwd() + "\\" +  fileName)
                        sheetCount = IFSpecWb.Worksheets.Count
                        for i in range(1, sheetCount + 1):
                            opeSheet = IFSpecWb.Worksheets(i)
                            sheet_name = opeSheet.Name
            
                            if sheet_name.upper() not in ("HISTORY","COVER","FLOW","CATALOG"):
                                resultSheet.Cells(wRow,1).Value = name
                                resultSheet.Cells(wRow,2).Value = sheet_name
                                wRow += 1
                    
                        IFSpecWb.Close(SaveChanges = 0)
                    except Exception as e: 
                        print(e)
                        #print("No history in this document")


        self.resultBook.Close(SaveChanges = 1)

    def deleteSameFile(self):
         same_files = {}
         for root,dirs,files in os.walk("res\\DMS0"):
            for name in files: 
                foundfile = os.path.join(root, name)
                file_label = name[0:name.find("-")]
                ver_no = name.split("-")[1]
                if file_label in same_files:
                    ver_no = name.split("-")[1]
                    list_ver_no = same_files[file_label].split("|")[0]
                    list_file = same_files[file_label].split("|")[1]
                    if int(ver_no) >= int(list_ver_no):
                        listfile =os.getcwd() + "\\"+ os.path.join(root, list_file)
                        os.remove(listfile)
                    else:
                        foundfile = os.getcwd() + "\\"+ foundfile
                        os.remove(name)
                else:
                    
                    same_files[file_label] = ver_no + "|" + name
    
    def FindDataLabel(self):
        ba_file = os.getcwd() + "\\res\\test\\gncanvehcs-PFV3-9000-a-BITASSIGN.xls"
        ba_wb  = self.excelApp.Workbooks.Open(ba_file)
        ba_sheet = ba_wb.Sheets("Transmit")
        info = ba_sheet.UsedRange
        nrows = info.Rows.Count
        row = 14
        Msg_label = ""
        BA_List = {}
        while (row <= nrows):
            BA_List[getCellValueinString(ba_sheet.Cells(row, 26))] = getCellValueinString(ba_sheet.Cells(row, 24)) + "|" + getCellValueinString(ba_sheet.Cells(row, 1))
            row += 1

        spec_file = os.getcwd() + "\\res\\test\\【参考】多感覚_EhternetCAN変換表_100.xlsx"
        specWb  = self.excelApp.Workbooks.Open(spec_file) 
        specSheet = specWb.Sheets("Sheet2")
        info = specSheet.UsedRange
        nrows = info.Rows.Count
        row = 6
        while (row <= nrows+1):
            comment_value = getCellValueinString(specSheet.Cells(row, 28))
                                                
            if comment_value in BA_List:
                Data_label = BA_List[comment_value].split("|")[0]
                Msg_label = BA_List[comment_value].split("|")[1]
                specSheet.Cells(row, 27).Value = Data_label
                specSheet.Cells(row, 23).Value = Msg_label
            
            row += 1
        
        specWb.Close(SaveChanges = 1)



    def FillBAMessage(self):
        ba_file = os.getcwd() + "\\res\\test\\24CY_Multi-Media_system_CAN_Com_spec_v3.30_19PFv3_BA.xlsx"
        ba_wb  = self.excelApp.Workbooks.Open(ba_file)

        ba_sub_file = os.getcwd() + "\\res\\test\\24CY_MM-SUB-BUS_Com_spec_Ver.3.20_BA.xlsm"
        ba_sub_wb  = self.excelApp.Workbooks.Open(ba_sub_file)   

        ba_sheet = ba_wb.Sheets("ビットアサイン表")
        info = ba_sheet.UsedRange
        nrows = info.Rows.Count
        row = 14
        Msg_label = ""
        BA_List = {}
        while (row <= nrows):
            '''
            if not isEmptyValue(ba_sheet.Cells(row, 1)):
                Msg_label = getCellValueinString(ba_sheet.Cells(row, 1))
            else:
                ba_sheet.Cells(row, 1).Value = Msg_label
            '''
            BA_List[getCellValueinString(ba_sheet.Cells(row, 24))] = getCellValueinString(ba_sheet.Cells(row, 1)) + "|" + getCellValueinString(ba_sheet.Cells(row, 38)) + "|"  \
                            + getCellValueinString(ba_sheet.Cells(row, 21)) + "|"+ getCellValueinString(ba_sheet.Cells(row, 22)) + "|"+ getCellValueinString(ba_sheet.Cells(row, 23)) +"|"+ getCellValueinString(ba_sheet.Cells(row, 25))

            
            row += 1
        
        ba_sub_sheet = ba_sub_wb.Sheets("Transmit")
        info = ba_sub_sheet.UsedRange
        nrows = info.Rows.Count
        row = 14
        Msg_label = ""
        BA_Sub_List = {}
        while (row <= nrows):
            '''
            if not isEmptyValue(ba_sheet.Cells(row, 1)):
                Msg_label = getCellValueinString(ba_sheet.Cells(row, 1))
            else:
                ba_sheet.Cells(row, 1).Value = Msg_label初始值	F/S值	"Signal Description
                信号描述"	"Diagram
                示意图"	"Comments
                备注"	"Source: Reference Chapter / Page / QA No.
                溯源：参考文档 章节/页码/QA号"		"Revision
                修订版本"	Applicable Projects 适用项目				"Comments
                备注"		
                                    "24MM中国Lexus
                L1"	"24MM中国Lexus
                L2"		23MM Mid AVN	23MM Hi AVN	23MM Low DA	24MM中国Lexus			
                                                All	All	All	All			
                
            '''
            BA_Sub_List[getCellValueinString(ba_sub_sheet.Cells(row, 34))] = getCellValueinString(ba_sub_sheet.Cells(row, 1)) + "|" + getCellValueinString(ba_sub_sheet.Cells(row, 53)) + "|"  \
                            + getCellValueinString(ba_sub_sheet.Cells(row, 26)) + "|"+ getCellValueinString(ba_sub_sheet.Cells(row, 32)) + "|"+ getCellValueinString(ba_sub_sheet.Cells(row, 33)) + "|"+ getCellValueinString(ba_sub_sheet.Cells(row, 35))
            row += 1
        
        
        spec_file = os.getcwd() + "\\res\\test\\test.xlsx"
        specWb  = self.excelApp.Workbooks.Open(spec_file) 
        specSheet = specWb.Sheets("Sheet1")
        info = specSheet.UsedRange
        nrows = info.Rows.Count
        row = 2
        while (row <= nrows+1):
            Data_label = getCellValueinString(specSheet.Cells(row, 11)).replace(" ","").replace("※","").upper()
            if Data_label in BA_List and Data_label != "":
                bit_info = BA_List[Data_label]
                Msg_label = bit_info.split("|")[0]
                initial_value = bit_info.split("|")[1]
                signal_flag = bit_info.split("|")[2]
                data_pos = bit_info.split("|")[3]
                data_len = bit_info.split("|")[4]
                data_name = bit_info.split("|")[5]
                specSheet.Cells(row, 8).Value = Msg_label
                specSheet.Cells(row, 13).Value = initial_value
                specSheet.Cells(row, 9).Value = data_pos
                specSheet.Cells(row, 10).Value = data_len
                specSheet.Cells(row, 3).Value = data_name

                '''
                Data_label = getCellValueinString(specSheet.Cells(row, 25)).replace(" ","").replace("※","")
                if Data_label in BA_List and Data_label != "":
                    bit_info = BA_List[Data_label]
                    Msg_label = bit_info.split("|")[0]
                    initial_value = bit_info.split("|")[1]
                    signal_flag = bit_info.split("|")[2]
                    data_pos = bit_info.split("|")[3]
                    data_len = bit_info.split("|")[4]
                    specSheet.Cells(row, 36).Value = initial_value
                    specSheet.Cells(row, 23).Value = data_pos
                    specSheet.Cells(row, 24).Value = data_len

                '''

                for root,dirs,files in os.walk("res\\DMS\\"+Msg_label):
                    for name in files: 
                        file_label = name[0:name.find("-")]
                        if Data_label==file_label:
                            foundfile = os.path.join(root, name)
                            dms_file = os.getcwd() + "\\"+foundfile
                            dmsWb  = self.excelApp.Workbooks.Open(dms_file)
                            dmsSheet = dmsWb.Sheets("Data Master Sheet")

                            dms_info = dmsSheet.UsedRange
                            dms_nrows = dms_info.Rows.Count
                            dms_row = 2
                            t_row_no = 1
                            r_row_no = 1
                            while dms_row <= dms_nrows:
                                if getCellValueinString(dmsSheet.Cells(dms_row, 9)).find("Transmit Data Inf.") >= 0:
                                    t_row_no = dms_row
                                if getCellValueinString(dmsSheet.Cells(dms_row, 9)).find("Receive Data Inf.") >= 0:
                                    if  getCellValueinString(dmsSheet.Cells(dms_row+1, 5)).find("AVN") >= 0:
                                        r_row_no = dms_row
                                        break
                                dms_row += 1
                            
                            print(Data_label, signal_flag)
                            if signal_flag == "T":
                                specSheet.Cells(row, 14).Value  = getCellValueinString(dmsSheet.Cells(t_row_no + 10,13))
                                specSheet.Cells(row, 4).Value = "Operate Signal\n操作信号"
                            if signal_flag == "R":
                                specSheet.Cells(row, 14).Value  = getCellValueinString(dmsSheet.Cells(r_row_no + 4,13))
                                specSheet.Cells(row, 4).Value = "Display Signal\n显示信号"
                            
                            specSheet.Cells(row, 12).Value  = getCellValueinString(dmsSheet.Cells(36,3))
                            specSheet.Cells(row, 1).Value  = dms_file
                            
                            dmsWb.Close(SaveChanges = 0)
                            
                
            if Data_label in BA_Sub_List and Data_label != "":
                bit_info = BA_Sub_List[Data_label]
                Msg_label = bit_info.split("|")[0]
                initial_value = bit_info.split("|")[1]
                signal_flag = bit_info.split("|")[2]
                data_pos = bit_info.split("|")[3]
                data_len = bit_info.split("|")[4]
                data_name = bit_info.split("|")[5]
                specSheet.Cells(row, 8).Value = Msg_label
                specSheet.Cells(row, 13).Value = initial_value
                specSheet.Cells(row, 9).Value = data_pos
                specSheet.Cells(row, 10).Value = data_len
                specSheet.Cells(row, 3).Value = data_name
                for root,dirs,files in os.walk("res\\DMS_MM-SUB-BUS\\"+Msg_label):
                    for name in files: 
                        file_label = name[0:name.find("-")]
                        if Data_label==file_label:
                            foundfile = os.path.join(root, name)
                            dms_file = os.getcwd() + "\\"+foundfile
                            dmsWb  = self.excelApp.Workbooks.Open(dms_file)
                            dmsSheet = dmsWb.Sheets("Data Master Sheet")
                            print(Data_label, signal_flag)
                            if signal_flag == "T":
                                specSheet.Cells(row, 14).Value  = getCellValueinString(dmsSheet.Cells(22,13))
                                specSheet.Cells(row, 4).Value = "Operate Signal\n操作信号"
                            if signal_flag == "R":
                                specSheet.Cells(row, 14).Value  = getCellValueinString(dmsSheet.Cells(82,13))
                                specSheet.Cells(row, 4).Value = "Display Signal\n显示信号"
                            
                            specSheet.Cells(row, 1).Value  = dms_file
                            specSheet.Cells(row, 12).Value  = getCellValueinString(dmsSheet.Cells(36,3))
                            dmsWb.Close(SaveChanges = 0)
                            
            
            row += 1

        ba_wb.Close(SaveChanges = 0)
        ba_sub_wb.Close(SaveChanges = 0)
        specWb.Close(SaveChanges = 1)

    def findSignalInfo(self):
        dev_file = os.getcwd() + "\\doc\\丰田_24MM_VehicleHAL_propertyid定义.xlsx"
        devWb  = self.excelApp.Workbooks.Open(dev_file)
        devSheet = devWb.Sheets("VehicleSetting")
        info = devSheet.UsedRange
        nrows = info.Rows.Count
        row = 3
        vd_info = {}
        while (row <= nrows):
            keyID =  getCellValueinString(devSheet.Cells(row, 19)).upper()
            vd_info[keyID] = getCellValueinString(devSheet.Cells(row, 23)).upper()
            row += 1
        
        spec_file = os.getcwd() + "\\doc\\24MM 09_01_01_Spec_Vehicle exterior.xlsx"
        specWb  = self.excelApp.Workbooks.Open(spec_file)
        specSheet = specWb.Sheets("02_Signal List_vs 19PFv3")
        info = devSheet.UsedRange
        nrows = info.Rows.Count
        row = 5

        while (row <= nrows):
            keyID =  getCellValueinString(specSheet.Cells(row, 12)).upper()
            if keyID in vd_info:
                specSheet.Cells(row, 6).value = vd_info[keyID]
            row += 1

        specWb.Close(SaveChanges =1)

    def fillVCinfo(self):
        spec_file = os.getcwd() + "\\res\\test\\test.xlsx"
        specWb  = self.excelApp.Workbooks.Open(spec_file)
        devSheet = specWb.Sheets("Sheet")
        info = devSheet.UsedRange
        nrows = info.Rows.Count
        row = 4
        vd_info = {}
        while (row <= nrows):
            keyID =  getCellValueinString(devSheet.Cells(row, 11)).upper()
            vd_info[keyID] = str(row)
            row += 1
        
        specSheet = specWb.Sheets("02_Signal List 19PFv3")
        info = specSheet.UsedRange
        nrows = info.Rows.Count
        row = 5
        while (row <= nrows):
            keyID =  getCellValueinString(specSheet.Cells(row, 11))
            if keyID in vd_info:
                specSheet.Cells(row, 2).Value = getCellValueinString(devSheet.Cells(int(vd_info[keyID]), 4))
                specSheet.Cells(row, 3).Value = getCellValueinString(devSheet.Cells(int(vd_info[keyID]), 5))
                specSheet.Cells(row, 6).Value = getCellValueinString(devSheet.Cells(int(vd_info[keyID]), 6))

                devSheet.Cells(int(vd_info[keyID]), 11).font.color = rgb_to_hex((0,0,255)) 
                devSheet.Cells(int(vd_info[keyID]), 11).font.color = rgb_to_hex((0,0,255)) 

            row += 1

        specWb.Close(SaveChanges = 1)

    def CopyFileToReleaseFolder(self):
        for root, dirs0, files in os.walk("res\\checkhistory"):
            for name in files: 
                if name not in (".DS_Store","Thumbs.db") and name.find(".xls", 0) >= 0:
                    path_no = name[5:13]
                    fileName = os.path.join(root, name)
                    src_file = os.getcwd() + "\\" + fileName
                    for root_sub, dirs, files0 in os.walk("C:\\workspace\\Rhine\\01_开发库\\11_需求管理\\05_需求规格说明书\\03_SpecRelease\\01_Func Release"):
                        for dir in dirs:
                            if dir.find(path_no) >= 0:
                                target_path =  "C:\\workspace\\Rhine\\01_开发库\\11_需求管理\\05_需求规格说明书\\03_SpecRelease\\01_Func Release\\" + dir + "\\v1.50\\" + name
                                copyfile(src_file, target_path)
                                new_name = name.replace(".xlsx","_v1.50.xlsx")
                                target_path = root_sub + "\\"+ dir + "\\v1.50\\"
                                old_fileName = os.path.join(target_path, name)
                                new_fileName = os.path.join(target_path, new_name)
                                if os.path.exists(new_fileName):
                                    os.remove(new_fileName)
                                os.rename(old_fileName, new_fileName)

                                print("copy", new_fileName)
                                break




    def diff22DTEM(self):
        fipfile_old = os.getcwd() + "\\res\\22dtem\\22TDEM_destinationcar matrix_R07.xlsx"
        fipWb_old  = self.excelApp.Workbooks.Open(fipfile_old)
        
        fipfile = os.getcwd() + "\\res\\22dtem\\22TDEM_destinationcar matrix_R08.xlsx"
        fipWb  = self.excelApp.Workbooks.Open(fipfile)


        opeSheet_old = fipWb_old.Sheets("Destination & Car (22TDEM)")
        opeSheet_new = fipWb.Sheets("Destination & Car (22TDEM)")

        info = opeSheet_old.UsedRange
        nrows = info.Rows.Count
        row = 1
        while (row <= nrows):
            rowText_old = ""
            rowText_new = ""

            for j in range (1,100):
                rowText_old = getCellValueinString(opeSheet_old.Cells(row, j)).replace(" ","")
                rowText_new = getCellValueinString(opeSheet_new.Cells(row, j)).replace(" ","")

                if rowText_old != rowText_new:
                    if opeSheet_new.Cells(row,j).interior.color != rgb_to_hex((255,255,0)):
                        opeSheet_new.Cells(row,j).interior.color = rgb_to_hex((255,204,255))
                        opeSheet_new.Cells(row,j).Value = getCellValueinString(opeSheet_new.Cells(row,j)) + " { " + rowText_old + " } "
            
            row +=1
        
        fipWb_old.Close(SaveChanges = 0)
        fipWb.Close(SaveChanges = 1)

    def deleteCellsStyle(self):
        spec_file = os.getcwd() + "\\res\\test\\Vehicle setting HMI specification.xlsx"
        specWb  = self.excelApp.Workbooks.Open(spec_file) 
        
        style_count = specWb.Styles.count

        for i in range(0, style_count):
            print(specWb.Styles[0])
            specWb.Styles[0].delete

            

        specWb.Close(SaveChanges = 1)

    def check_hidden_sheet_for_24MM(self):
        self.resultBook.ActiveSheet.Name = "document list"
        resultSheet = self.resultBook.ActiveSheet
        wRow = 2

        for root, dirs, files in os.walk("res\\tesline"):
            for name in files: 
                if name not in (".DS_Store","Thumbs.db") and name.find(".xls", 0) >= 0:
                    fileName = os.path.join(root, name)
                    file = os.getcwd() + "\\" + fileName
                    print("start analyz:", fileName)
                    try:
                        specWb  = self.excelApp.Workbooks.Open(file)
                        sheetCount = specWb.Worksheets.Count
                        for i in range(1, sheetCount + 1):
                            sheet_name = specWb.Worksheets(i).Name
                            if specWb.Worksheets(i).visible != -1:
                                resultSheet.Cells(wRow, 1).Value = fileName
                                resultSheet.Cells(wRow, 2).Value = sheet_name
                                wRow += 1
                        specWb.Close(SaveChanges = 0)
                    
                    except Exception as e:
                        print(e)
                        copyspecFile(fileName, "errorfile")
        
        self.resultBook.Close(SaveChanges = 1)

    def extract_sourcelist_for_24MM(self):
        self.resultBook.ActiveSheet.Name = "document list"
        resultSheet = self.resultBook.ActiveSheet
        wRow = 2

        for root, dirs, files in os.walk("res\\tesline"):
            for name in files: 
                if name not in (".DS_Store","Thumbs.db") and name.find(".xls", 0) >= 0:
                    fileName = os.path.join(root, name)
                    file = os.getcwd() + "\\" + fileName
                    print("start analyz:", fileName)
                    try:
                        startRow = 4
                        specWb  = self.excelApp.Workbooks.Open(file)
                        resultSheet.Cells(wRow, 1).Value = fileName
                        spec_sheet = specWb.Sheets("00_Source List")
                        info = spec_sheet.UsedRange
                        nrows = info.Rows.Count
                        ncols = info.Columns.Count
                        while(startRow <= nrows):
                            resultSheet.Cells(wRow, 1).Value = fileName
                            if not isEmptyValue(spec_sheet.Cells(startRow, 1)) and getCellValueinString(spec_sheet.Cells(startRow, 1)) != "No.":
                                range_s= "A" + str(startRow) + ":G" + str(startRow)
                                range_t= "B" + str(wRow) + ":H" + str(wRow)
                                spec_sheet.Cells.Range(range_s).Copy(resultSheet.Range(range_t))
                                wRow += 1
                            startRow += 1
                       
                        specWb.Close(SaveChanges = 0)

                    except Exception as e:
                        print(e)
                        copyspecFile(fileName, "errorfile")
            
        self.SaveResultFile("func_sourcelist.xlsx")     


    def extract_speccontent_for_24MM(self):
        self.resultBook.ActiveSheet.Name = "document list"
        resultSheet = self.resultBook.ActiveSheet
        wRow = 2

        for root, dirs, files in os.walk("res\\test"):
            for name in files: 
                if name not in (".DS_Store","Thumbs.db") and name.find(".xls", 0) >= 0:
                    fileName = os.path.join(root, name)
                    file = os.getcwd() + "\\" + fileName
                    print("start analyz:", fileName)
                    try:
                        specWb  = self.excelApp.Workbooks.Open(file)
                        sheetCount = specWb.Worksheets.Count
                        resultSheet.Cells(wRow, 1).Value = fileName

                        for i in range(1, sheetCount + 1):
                            sheet_name = specWb.Worksheets(i).Name
                            if sheet_name.upper().replace(" ","") not in ("00_SOURCELIST","CHANGEHISTORY") \
                                and sheet_name.upper().find("ABBREVIATION&TERMINOLOGY")<0 and sheet_name.upper().find("APPENDIX",0) < 0 and sheet_name.upper().find("COVER",0) < 0:
                                spec_sheet = specWb.Sheets(sheet_name)
                                if spec_sheet.visible == -1:
                                    resultSheet.Cells(wRow, 1).Value = fileName
                                    resultSheet.Cells(wRow, 2).Value = sheet_name
                                    print("start analyz sheet:", sheet_name)
                                    startRow = 1
                                    info = spec_sheet.UsedRange
                                    nrows = info.Rows.Count
                                    ncols = info.Columns.Count
                                    if ncols > 200:
                                        ncols = 200
                                    version_col = 0
                                    require_id_col = 0
                                    desc_id_col = 0
                                    comment_id_col = 0
                                    shimuke_col = 0
                                    bFound_title = False
                                    while(startRow <= nrows):
                                        if not bFound_title:
                                            for  startCol in range(1, ncols+5):
                                                if getCellValueinString(spec_sheet.Cells(startRow, startCol)).upper().replace("\n","").replace(" ","").find("REQUIREMENT", 0) >= 0:
                                                    require_id_col = startCol
                                                if getCellValueinString(spec_sheet.Cells(startRow, startCol)).upper().replace("\n","").replace(" ","").find("REVISION", 0) >= 0:
                                                    version_col = startCol
                                                if getCellValueinString(spec_sheet.Cells(startRow, startCol)).upper().replace("\n","").replace(" ","").find("DESCRIPTION", 0) >= 0:
                                                    desc_id_col = startCol
                                                if getCellValueinString(spec_sheet.Cells(startRow, startCol)).upper().replace("\n","").replace(" ","").find("COMMENTS", 0) >= 0:
                                                    comment_id_col = startCol
                                                if getCellValueinString(spec_sheet.Cells(startRow, startCol)).upper().replace("\n","").replace(" ","").find("LEXUS&TOYOTA差分", 0) >= 0:
                                                    shimuke_col = startCol
                                                if startRow > 30:
                                                    break
                                            if require_id_col > 0 and version_col > 0 :
                                                bFound_title = True
                                                startRow += spec_sheet.Cells(startRow, require_id_col).MergeArea.Rows.Count - 1
                                        else:
                                            require_id =  getCellValueinString(spec_sheet.Cells(startRow, require_id_col)).replace(" ","") 
                                            resultSheet.Cells(wRow, 1).Value = file
                                            resultSheet.Cells(wRow, 2).Value = sheet_name
                                            resultSheet.Cells(wRow, 3).Value = str(startRow)
                                            resultSheet.Cells(wRow, 4).Value = require_id
                                            if desc_id_col > 0:
                                                range_s= chr(desc_id_col + 64) + str(startRow)
                                                range_t= "E" + str(wRow)
                                                spec_sheet.Cells.Range(range_s).Copy(resultSheet.Range(range_t))
                                            # surpose column won't be larger than 52
                                            if comment_id_col > 0:
                                                if comment_id_col > 26:
                                                    range_s= "A" + chr(comment_id_col-26 + 64) + str(startRow)
                                                else:
                                                    range_s= chr(comment_id_col + 64) + str(startRow)
                                            
                                                range_t= "F" + str(wRow)
                                                spec_sheet.Cells.Range(range_s).Copy(resultSheet.Range(range_t))
                                            if shimuke_col > 0:
                                                if shimuke_col > 26:
                                                    range_s= "A" + chr(shimuke_col-26 + 64) + str(startRow)
                                                else:
                                                    range_s= chr(shimuke_col + 64) + str(startRow)
                                                range_t= "G" + str(wRow)
                                                spec_sheet.Cells.Range(range_s).Copy(resultSheet.Range(range_t))
                                            wRow += 1

                                        startRow += 1

                        
                        specWb.Close(SaveChanges = 0)

                    except Exception as e:
                        print(e)
                        copyspecFile(fileName, "errorfile")
            
        self.SaveResultFile("func_all.xlsx")     

    def checkhistory_for_24MM(self):
        self.resultBook.ActiveSheet.Name = "document list"
        resultSheet = self.resultBook.ActiveSheet
        wRow = 2

        for root, dirs, files in os.walk("res\\tesline"):
            for name in files: 
                if name not in (".DS_Store","Thumbs.db") and name.find(".xls", 0) >= 0:
                    fileName = os.path.join(root, name)
                    file = os.getcwd() + "\\" + fileName
                    print("start analyz:", fileName)
                    try:
                        specWb  = self.excelApp.Workbooks.Open(file)
                        history_list = {}
                        sheetCount = specWb.Worksheets.Count
                        resultSheet.Cells(wRow, 1).Value = fileName

                        for i in range(1, sheetCount + 1):
                            sheet_name = specWb.Worksheets(i).Name
                            if sheet_name.upper().replace(" ","").find("CHANGEHISTORY", 0) >= 0:
                                history_sheet = specWb.Sheets(sheet_name)
                                startRow = 4
                                info = history_sheet.UsedRange
                                nrows = info.Rows.Count
                                if nrows > 200:
                                    nrows = 200
                                while(startRow <= nrows):
                                    version_no = getCellValueinString(history_sheet.Cells(startRow, 2)).upper().replace(" ","")
                                    requirement_value = getCellValueinString(history_sheet.Cells(startRow, 12))
                                    history_comment = getCellValueinString(history_sheet.Cells(startRow, 13))
                                    if version_no.find("2.00", 0) >= 0:
                                        requirements_list = requirement_value.split("\n")
                                        for require_id in requirements_list:
                                            if require_id != "-":
                                                if require_id not in history_list:
                                                    history_list[require_id] = history_comment
                                          
                                                                                       
                                    startRow += 1
                            elif sheet_name.upper().replace(" ","") not in ("00_SOURCELIST","04_ABBREVIATION&TERMINOLOGY","COVER"):
                                spec_sheet = specWb.Sheets(sheet_name)
                                if spec_sheet.visible == -1:
                                    resultSheet.Cells(wRow, 1).Value = name
                                    resultSheet.Cells(wRow, 2).Value = sheet_name
                                    print("start analyz sheet:", sheet_name)
                                    startRow = 1
                                    info = spec_sheet.UsedRange
                                    nrows = info.Rows.Count
                                    ncols = info.Columns.Count
                                    if ncols > 200:
                                        ncols = 200
                                    version_col = 0
                                    require_id_col = 0
                                    desc_id_col = 0
                                    bFound_title = False
                                    while(startRow <= nrows +1):
                                        if not bFound_title:
                                            for  startCol in range(1, ncols):
                                                if getCellValueinString(spec_sheet.Cells(startRow, startCol)).upper().replace("\n","").replace(" ","").find("REQUIREMENTSID", 0) >= 0:
                                                    require_id_col = startCol
                                                if getCellValueinString(spec_sheet.Cells(startRow, startCol)).upper().replace("\n","").replace(" ","").find("REVISION", 0) >= 0:
                                                    version_col = startCol
                                                if getCellValueinString(spec_sheet.Cells(startRow, startCol)).upper().replace("\n","").replace(" ","").find("DESCRIPTION", 0) >= 0:
                                                    desc_id_col = startCol
                                                if require_id_col > 0 and version_col > 0:
                                                    bFound_title = True
                                                    # startRow +=  (spec_sheet.Cells(startRow, require_id_col).MergeArea.Rows.Count -1)
                                                    break
                                            if startRow > 30 :
                                                break
                                        else:
                                            #if isEmptyValue(spec_sheet.Cells(startRow, require_id_col)):
                                            if startRow > 2000:
                                                break

                                            version_inFunc = getCellValueinString(spec_sheet.Cells(startRow, version_col)).upper().replace(" ","") 
                                            if version_inFunc.find("2.00") >= 0:
                                                require_id =  getCellValueinString(spec_sheet.Cells(startRow, require_id_col))
                                                resultSheet.Cells(wRow, 1).Value = name
                                                resultSheet.Cells(wRow, 2).Value = sheet_name
                                                resultSheet.Cells(wRow, 3).Value = str(startRow)
                                                resultSheet.Cells(wRow, 4).Value = require_id
                                                '''
                                                if desc_id_col > 0:
                                                    resultSheet.Cells(wRow, 5).Value = getCellValueinString(spec_sheet.Cells(startRow, desc_id_col)).replace(" ","") 
                                                wRow += 1
                                                '''
                                                if require_id not in history_list:
                                                    resultSheet.Cells(wRow, 1).Value = name
                                                    resultSheet.Cells(wRow, 2).Value = sheet_name
                                                    resultSheet.Cells(wRow, 3).Value = str(startRow)
                                                    resultSheet.Cells(wRow, 4).Value = require_id
                                                    resultSheet.Cells(wRow, 5).Value = version_inFunc

                                                    if spec_sheet.Cells(startRow, require_id_col).Font.Strikethrough == True:
                                                        resultSheet.Cells(wRow, 6).Value = "strikethrough true"
                                                    else:
                                                        resultSheet.Cells(wRow, 6).Value = "History漏记"                                                        
                                                    if len(history_list) == 0:
                                                        resultSheet.Cells(wRow, 6).Value = "可能初版新规"
                                                    wRow += 1
                                                else:
                                                    resultSheet.Cells(wRow, 3).Value = ""
                                                    resultSheet.Cells(wRow, 4).Value = ""

                                                    del history_list[require_id]                                

                                        if bFound_title:
                                            if spec_sheet.Cells(startRow,require_id_col).MergeArea.Rows.count > 1:
                                                startRow += spec_sheet.Cells(startRow,require_id_col).MergeArea.Rows.count
                                            else:
                                                startRow += 1
                                        else:
                                            startRow += 1
                                    if require_id_col == 0 or version_col == 0:
                                        resultSheet.Cells(wRow, 6).Value = "本Sheet无Requirement/revision列"
                                        wRow += 1
                                    
                                        
                                else:
                                    wRow += 1
                        
                        
                        for history in history_list:
                            resultSheet.Cells(wRow, 1).Value = name
                            resultSheet.Cells(wRow, 4).Value = history
                            resultSheet.Cells(wRow, 6).Value = "正文版本号没有更新"
                            resultSheet.Cells(wRow, 7).Value = history_list[history]
                            wRow += 1


                        specWb.Close(SaveChanges = 0)

                    except Exception as e:
                        print(e)
                        copyspecFile(fileName, "errorfile")

        self.SaveResultFile("func_history_missing.xlsx")
    
    def GuidelinesApply(self):
        guidefile = os.getcwd() + "\\res\\guidelines\\24MM机能需求规格说明书Guideline.xlsx"
        guideWb  = self.excelApp.Workbooks.Open(guidefile)
        guideSheet = guideWb.Sheets("模板-00_Source List")

        for root, dirs, files in os.walk("res\\checkhistory"):
            for name in files: 
                if name not in (".DS_Store","Thumbs.db") and name.find(".xls", 0) >= 0:
                    fileName = os.path.join(root, name)
                    file = os.getcwd() + "\\" + fileName
                    print("start analyz:", fileName)
                    try:
                        specWb  = self.excelApp.Workbooks.Open(file)
                        specSheet = specWb.Sheets("00_Source List")

                        #specSheet.Range("A1:G1").unmerge()
                        guideSheet.Cells.Range("A1:G1").Copy(specSheet.Range("A1:G1"))
                        #specSheet.Range("A1:A7").merge()
                        specSheet.Cells(1,1).font.size = 10
                        specSheet.Cells(1,1).font.name = "等线"

                        specWb.Close(SaveChanges  =1)

                    except Exception as e:
                        print(e)
                        copyspecFile(fileName, "errorfile")

        guideWb.Close(SaveChanges = 0)


    def testCellTextColor(self):
        fipfile = os.getcwd() + "\\res\\24MM 09_13_01_Spec_ETC.xlsx"
        fipWb  = self.excelApp.Workbooks.Open(fipfile)
        fipSheet = fipWb.Sheets("03_Appendix")

        info = fipSheet.UsedRange
        nrows = info.Rows.Count
        ncols = info.Columns.Count
        
        startRow = 1

        while startRow <= nrows:
            range_s= "A" + str(startRow) + ":" + "Z" + str(startRow)
            textChar = fipSheet.Range(range_s).Characters
            print(textChar)
            startRow += 1
        
        fipWb(SaveChanges = 0)
        
    def ReplaceInShape(self): 
        specfile = os.getcwd() + "\\res\\24MM 09_04_01_Spec_Charging Function.xlsx"
        specWb  = self.excelApp.Workbooks.Open(specfile)
        specSheet = specWb.Sheets("03 Sequence Timer設定変更処理")
        
        for shape in specSheet.shapes:
            try :
                print(shape.TextFrame.Characters().text)
                cell_text = shape.TextFrame.Characters().text
                shape.TextFrame.Characters().text = cell_text.replace("で","  中 ")
            
            except Exception as e:
                print(e) 

        specWb.Close(SaveChanges = 1)

        
    def ReleaseCheckFor24MM(self):
        for root, dirs, files in os.walk("C:\\workspace\Rhine\\01_开发库\\11_需求管理\\05_需求规格说明书\\03_SpecRelease\\01_Func Release"):
            for name in files: 
                try:
                    if name not in (".DS_Store","Thumbs.db") and name.find(".xls", 0) >= 0:
                        if root.find("1.50") >=0 :
                            spec_path = root.replace("C:\\workspace\Rhine\\01_开发库\\11_需求管理\\05_需求规格说明书\\03_SpecRelease\\01_Func Release\\","")
                            spec_path = spec_path.replace("\\v1.50","")
                            spec_number = spec_path[0:8]
                            spec_name = spec_path[9:]
                            if name.find("Cover_Sheet") >= 0:
                                fileName = os.path.join(root, name)
                                print("start analyz:", fileName)
                                specWb  = self.excelApp.Workbooks.Open(fileName)
                                specSheet =  specWb.Sheets("Func. Spec Cover Temp._24中国")
                                specSheet.Cells(6,7).Value = "Ver.1.50"
                                specSheet.Cells(6,5).Value = spec_number
                                specWb.Close(SaveChanges = 1)
                            if name.find("Check_Sheet") >= 0:
                                fileName = os.path.join(root, name)
                                print("start analyz:", fileName)
                                specWb  = self.excelApp.Workbooks.Open(fileName)
                                specSheet =  specWb.Sheets("機能仕様チェックシート")
                                specSheet.Cells(7,8).Value = "Ver.1.50"
                                specSheet.Cells(3,10).Value = spec_number
                                specSheet.Cells(5,8).Value = spec_name
                                specWb.Close(SaveChanges = 1)
                except Exception as e:
                    print(e)


    def CheckModelType_for23raku(self):
        model_info_list = ["&NX853AA","&NX853BA","&NX853CA","&NX854C","&NX854D","&NX854E","&NX855A","&NX855B","&NX855C","&NX855D","&NX855E","&NX856C",\
                            "&NX856D","&NX856DA","&NX856E","&NX856EA","&NX857D","&NX857E"]

        self.excelApp.Workbooks.Add()
        resultBook = self.excelApp.ActiveWorkBook
        resultBook.ActiveSheet.Name = "Model info"
        resultSheet =resultBook.ActiveSheet
        wRow = 1

        for root, dirs, files in os.walk("C:\\workspace\\NX816_Spec\\ope\\AV\\fix"):
            for name in files: 
                try:
                    wRow += 1
                    if name not in (".DS_Store","Thumbs.db") and name.find(".xls", 0) >= 0:
                        fileName = os.path.join(root, name)
                        specWb  = self.excelApp.Workbooks.Open(fileName)
                        resultSheet.Cells(wRow, 1).Value = fileName
                        wRow += 1
                        specSheet =  specWb.Sheets("Catalog")
                        if getCellValueinString(specSheet.Cells(6,3)) == "Screen Name":
                            columnNo = 27
                            while columnNo <= 44:
                                if specSheet.Cells(6, columnNo).Font.Strikethrough != True:
                                    model_info = getCellValueinString(specSheet.Cells(6, columnNo))
                                    if model_info not in model_info_list:
                                        resultSheet.Cells(wRow, columnNo).Value = model_info
                                columnNo += 1
                        
                        specWb.Close(SaveChanges= 0)
                                    
                except Exception as e:
                    print(e)
                    
        resultBook.Close(SaveChanges = 1)

    def CheckMessageTimer_for23raku(self):
        self.excelApp.Workbooks.Add()
        resultBook = self.excelApp.ActiveWorkBook
        resultBook.ActiveSheet.Name = "Msg info"
        resultSheet =resultBook.ActiveSheet
        wRow = 1

        for root, dirs, files in os.walk("C:\\workspace\\NX816_Spec\\ope\\AV\\fix"):
            for name in files: 
                try:
                    wRow += 1
                    if name not in (".DS_Store","Thumbs.db") and name.find(".xls", 0) >= 0:
                        fileName = os.path.join(root, name)
                        specWb  = self.excelApp.Workbooks.Open(fileName)
                        sheetCount = specWb.Worksheets.Count
        
                        for i in range(1, sheetCount + 1):
                            sheet_name = specWb.Worksheets(i).Name
                            message_All_list = {}
                            message_list = {}

                            if sheet_name not in ("表紙","Catalog","History"):
                                specSheet = specWb.Worksheets(i)
                                if specSheet.Cells(2, 2).Font.Strikethrough != True:

                                    info = specSheet.UsedRange
                                    nrows = info.Rows.Count
                                    ncols = info.Columns.Count
                                    row = 66
                                    message_state = 1
                                    while(row <= nrows):
                                        #if specSheet.Cells(row,4).Font.Strikethrough != True:
                                        if getCellValueinString(specSheet.Cells(row,4)).find("メッセージ") > 0 and getCellValueinString(specSheet.Cells(row,2)) == "1":
                                            for j in range(2, 42):
                                                if not isEmptyValue(specSheet.Cells(row+1, j)):
                                                    message_list["1-"+ getCellValueinString(specSheet.Cells(row,3))] = getCellValueinString(specSheet.Cells(row+1, j))
                                                    break

                                        if getCellValueinString(specSheet.Cells(row,3)) == "Trigger Action":
                                            message_state = getCellValueinString(specSheet.Cells(row, 3))
                                            row += 1
                                            break

                                        row += 1


                                    while(row <= nrows):
                                        #if specSheet.Cells(row,42).Font.Strikethrough != True:
                                        if not isEmptyValue(specSheet.Cells(row,42)):
                                            if getCellValueinString(specSheet.Cells(row,42)) not in ("-","'-","Timer"):
                                                msg_tag = "1-"+ getCellValueinString(specSheet.Cells(row,3))
                                                if msg_tag in message_list:
                                                    msg_id = message_list[msg_tag]
                                                    message_All_list[msg_id] = getCellValueinString(specSheet.Cells(row,42)) + "|" + str(row)
                                        row += 1

                            for msgid in message_All_list:
                                resultSheet.Cells(wRow, 1).Value = fileName
                                resultSheet.Cells(wRow, 2).Value = sheet_name
                                resultSheet.Cells(wRow, 3).Value = msgid

                                messge_info = message_All_list[msgid].split("|")[0]
                                
                                specSheet.Range("AP" + message_All_list[msgid].split("|")[1]).Copy(resultSheet.Cells.Range("D"+str(wRow)))
                                #resultSheet.Cells(wRow, 4).Value = message_All_list[msgid]
                                wRow += 1

                        specWb.Close(SaveChanges= 0)
                                    
                except Exception as e:
                    print(e)
                    

        resultBook.Close(SaveChanges = 1)

    def CheckTeling_for23raku(self):
        self.excelApp.Workbooks.Add()
        resultBook = self.excelApp.ActiveWorkBook
        resultBook.ActiveSheet.Name = "Msg info"
        resultSheet =resultBook.ActiveSheet
        wRow = 1

        for root, dirs, files in os.walk("C:\\workspace\\NX816_Spec\\ope\\AV\\fix"):
            for name in files: 
                try:
                    wRow += 1
                    if name not in (".DS_Store","Thumbs.db") and name.find(".xls", 0) >= 0:
                        fileName = os.path.join(root, name)
                        specWb  = self.excelApp.Workbooks.Open(fileName)
                        sheetCount = specWb.Worksheets.Count
        
                        for i in range(1, sheetCount + 1):
                            sheet_name = specWb.Worksheets(i).Name
                            message_All_list = {}
                            message_list = {}

                            if sheet_name not in ("表紙","Catalog","History"):
                                specSheet = specWb.Worksheets(i)
                                if specSheet.Cells(2, 2).Font.Strikethrough != True:

                                    info = specSheet.UsedRange
                                    nrows = info.Rows.Count
                                    ncols = info.Columns.Count
                                    row = 50

                                    while(row <= nrows):
                                        for j in range(2, 42):
                                            if not isEmptyValue(specSheet.Cells(row, j)):
                                                if getCellValueinString(specSheet.Cells(row, j)).find("電話中") >= 0:
                                                    resultSheet.Cells(wRow, 1).Value = fileName
                                                    resultSheet.Cells(wRow, 2).Value = sheet_name
                                                    resultSheet.Cells(wRow, 3).Value = str(row)
                                                    resultSheet.Cells(wRow, 4).Value = getCellValueinString(specSheet.Cells(row, j))
                                                    wRow += 1

                                        row += 1

                        specWb.Close(SaveChanges= 0)
                                    
                except Exception as e:
                    print(e)
                    
        resultBook.Close(SaveChanges = 1)

if __name__ == '__main__':
    print("Start running ", time.localtime(time.time()))
    metaF = ToolsFixer()
    if len(sys.argv) == 2:
        if sys.argv[1] == "RETRIEVE_STR":
            metaF.RetrieveStringIDforRT()
        elif sys.argv[1] == "RETRIEVE_NT_STR":
            metaF.RetrieveNTStringIDforRT()
        elif sys.argv[1] == "CHECK_NT_STR":
            metaF.CheckNTString()
        elif sys.argv[1] == "UPDATE_ALLWORDS":
            metaF.UpdateAllwordsToStringTable()
        elif sys.argv[1] == "CHECK_UI_ALLWORDS":
            metaF.CompareUIResult()
        elif sys.argv[1] == "GET_ZILIAO":
            # metaF.getDataFromZiliaoZhan()
            metaF.testRedmineApi()
        elif sys.argv[1] == "TMC_ALLWORDS":
            metaF.UpdateTMCAllwords()
        elif sys.argv[1] == "CHK_TBD":
            metaF.CheckTBDItem()
        elif sys.argv[1] == "EXTRACT_FUNC":
            metaF.extract_funcid()
        elif sys.argv[1] == "TRACE_FUNC":
            metaF.trace_feature_func()
        elif sys.argv[1] == "FILL_UPDATEDATE":
            metaF.fillSusukiUpdateDate()
        elif sys.argv[1] == "FOR_JIRA":
            metaF.fillFormForJira()
        elif sys.argv[1] == "GET_GLOBAL":
            metaF.getGlobalInfo()
        elif sys.argv[1] == "CHECK_GLOBAL":
            metaF.CheckImpInfo_24MM()
        elif sys.argv[1] == "UPD_FORD_FIP":
            metaF.diffFordFIP()
        elif sys.argv[1] == "UPD_INFO_FORD":
            metaF.DiffDeveloperInfo()
        elif sys.argv[1] == "UPD_ESC":
            metaF.UpdateESCForFraser()
        elif sys.argv[1] == "DIFF_22":
            metaF.diff22DTEM()
        elif sys.argv[1] == "EXTRACT_IF_SPEC":
            metaF.extractIFSpec_history()
        elif sys.argv[1] == "CHECK_HISTORY_MISSING":
            metaF.checkhistory_for_24MM()

        elif sys.argv[1] == "TEST_COLOR":
            metaF.testCellTextColor()

        elif sys.argv[1] == "EXTTRACT_SPEC":
            metaF.extract_speccontent_for_24MM()

        elif sys.argv[1] == "CHK_HIDDEN_SPEC":
            metaF.check_hidden_sheet_for_24MM()

        elif sys.argv[1] == "CHK_SOURCELIST_SPEC":
            metaF.extract_sourcelist_for_24MM()
        elif sys.argv[1] == "FILL_BA":
            metaF.FillBAMessage()
        elif sys.argv[1] == "FILL_VC":
            metaF.fillVCinfo()
        elif sys.argv[1] == "DEL_FILE":
            metaF.deleteSameFile()
        elif sys.argv[1] == "DEL_STYLE":
            metaF.deleteCellsStyle()
        elif sys.argv[1] == "REP_SHAPE_TEXT":
            metaF.ReplaceInShape()
        elif sys.argv[1] == "GUIDELINE_APP":
            metaF.GuidelinesApply()
        elif sys.argv[1] == "COPY_FILE_TO_RELEASE_FOLDER":
            metaF.CopyFileToReleaseFolder()

        elif sys.argv[1] == "RELEASE_CHK":
            metaF.ReleaseCheckFor24MM()
        
        elif sys.argv[1] == "FILL_DATALABLE":
            metaF.FindDataLabel()

        elif sys.argv[1] == "FILL_SIG":
            metaF.findSignalInfo()
        elif sys.argv[1] == "CHK_MODLE_23RAKU":
            metaF.CheckMessageTimer_for23raku()
        elif sys.argv[1] == "CHK_MODLE_23RAKU_TEL":
            metaF.CheckTeling_for23raku()

    elif len(sys.argv) == 3:
    # 生成所有多状态对象列表
        if sys.argv[2] == "EDIT_FORMART":
            metaF.FindSpecPara(sys.argv[1])
        elif sys.argv[2] == "FIND_STRING":
            metaF.FindStringID(sys.argv[1])
        elif sys.argv[2] == "MERGE_COST":
            metaF.MergeCostFile(sys.argv[1])
        elif sys.argv[2] == "CHK_UUID":
            metaF.CheckScreenUUID(sys.argv[1])
        elif sys.argv[2] == "CHK_SPECID":
            metaF.CheckScreenSpecID(sys.argv[1])
        elif sys.argv[2] == "MERGE_RFQ":
            metaF.MergeRFQSheet(sys.argv[1])
        elif sys.argv[2] == "MERGE_NOCOST":
            metaF.MergeNoCostSheet(sys.argv[1])
        elif sys.argv[2] == "ADD_COMMENT":
            metaF.AddComment(sys.argv[1])
        elif sys.argv[2] == "ADD_HISTORY":
            metaF.AddHistoryComment(sys.argv[1])
        elif sys.argv[2] == "TEST_CELL":
            metaF.testCellValue(sys.argv[1])
        elif sys.argv[2] == "TEST_MYSQL":
            metaF.testDB(sys.argv[1])
        elif  sys.argv[2] == "TEST_SHEETNAME":
            metaF.printSheetName(sys.argv[1])
        elif  sys.argv[2] == "MERGE_SPECID":
            metaF.MergeSpecID(sys.argv[1])
        elif sys.argv[2] == "MERGE_RFQ_URANUS":
            metaF.MergeRFQSheet_uranus(sys.argv[1])
        elif sys.argv[2] == "RT_ALLWORDS":
            metaF.AbstractRTScreenAllwords(sys.argv[1])
        elif sys.argv[2] == "NF_ALLWORDS":
            metaF.AbstractRTNotificationAllwords(sys.argv[1])
        elif sys.argv[1] == "GLOBAL_INFO":
            metaF.findGlobalInfo(sys.argv[2])
        elif sys.argv[1] == "RESET_DOC":
            metaF.resetDocument(sys.argv[2])
        elif sys.argv[1] == "BASELINE":
            metaF.CreateBaseline(sys.argv[2])
        elif sys.argv[1] == "EXTRACT_HISTORY":
            metaF.ExtractHistoryVersion(sys.argv[2])
        elif sys.argv[1] == "UPDATE_MODULE_ANALYSIS":
            metaF.UpdateSpecFunc_ModuleAnalysis(sys.argv[2])
            
    elif len(sys.argv) == 4:
        # 生成所有多状态对象列表
        if sys.argv[3] == "SEARCH_SCREEN_ID":
            metaF.SearchScreenID(sys.argv[1],sys.argv[2])
        # 对比两个版本的part id差异
        print( 'Start searching...')
        if sys.argv[3] == "SYNC_HISTORY":
            metaF.CompareFolder(sys.argv[1], sys.argv[2])            
            metaF.AbstractHistoryUpdate(sys.argv[1], sys.argv[2])
        elif sys.argv[2] == "DEL_EXPORTED":
            metaF.DeleteExportedOnly(sys.argv[1], sys.argv[3])

    elif len(sys.argv) == 5:
        if sys.argv[3] in ("CMP_PART_ID") :
            metaF.CompareFolder(sys.argv[1], sys.argv[2])            
            metaF.WanderInFiles(sys.argv[1], sys.argv[2],sys.argv[4])

                     