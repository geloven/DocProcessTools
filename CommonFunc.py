__author__ = 'swallow'
__language__= 'python 3.0'

import os
import sys
import time
import re
import datetime

from shutil import copyfile
START_ROWNO = 49

PART_ID_COLNO = 1
DISPLAY_NAME_COLNO = 2
PART_TYPE_COLNO = 3
PART_NAME_COLNO = 4

DISPLAY_FORMULAR_NO = 5
DISPLAY_CONDITION_TAG_NO = 6
DISPLAY_CONDITION_NO = 7

DISP_SAME_CONDITION_PARTS_ID = 8
DISP_SAME_CONDITION_SCREEN_ID = 9
DISP_ELSE_NO = 10

JAPANESE_NO = 11
JAPAN_FIXED_WORDS_NO = 12
US_English_NO = 13
US_ENGLISH_FIXED_WORDS_NO = 14
UK_ENGLISH_NO = 15
UK_ENGLISH_FIXED_WORD_NO = 16

FIXED_IMAGE_NO = 17

DELETE_CONDITION_MOTION_FORMULA_NO = 18
DELETE_CONDITION_MOTION_CONDITION_TAG_NO = 19
DELETE_CONDITION_MOTION_NO = 20

TEXT_IMAGE_SAME_CONDITION_PARTS_ID = 21
TEXT_IMAGE_SAME_CONDITION_SCREEN_ID = 22
TEXT_IMAGE_DISPLAY_CONTENT = 23

OUTSIDE_INPUT_FORMAT_NO = 24
OUTSIDE_INPUT_RANGE_NO = 25
OUTSIDE_VALIDATION_NO = 26

TONEDOWN_FORMULA_NO = 27
TONEDOWN_CONDITION_TAG_NO = 28
TONEDOWN_CONDITION_ID = 29

TONEDOWN_SAME_CONDITION_PARTS_ID = 30
TONEDOWN_SAME_CONDITION_SCREEN_ID = 31


TONEDOWN_EXCEPT_FORMULAR_NO = 32
TONEDOWN_EXCEPT_CONDITION_TAG_NO = 33
TONEDOWN_EXCEPT_CONDITION_NO = 34

TONEDOWN_EXCEPTSAME_CONDITION_PARTS_ID = 35
TONEDOWN_EXCEPTSAME_CONDITION_SCREEN_ID = 36

SELECTED_CONDITION_FORMULAR_NO = 37
SELECTED_CONDITION_TAG_NO = 38
SELECTED_CONDITION_NO = 39

SELECTED_EXCEPTSAME_CONDITION_PARTS_ID = 40
SELECTED_EXCEPTSAME_CONDITION_SCREEN_ID = 41

SW_OPERATION_PATTERN_NO = 42
OPERATION_RESULT_SCNTRAN_NO = 43
OPERATION_START_FUNCTION_NO = 44
OPERATION_RESULT_SETTING_VALUE_CHANGE_NO = 45
OPERATION_RESULT_OTHER = 46
BEEP_COL_NO = 47
UUID_COLNO = 48

def copyspecFile(fileName, pathName):
    try:
        folder = os.path.exists(os.getcwd() + "\\" +pathName)
        if not folder:
            os.makedirs(pathName)
        src_file = os.getcwd() + "\\"+ fileName
        target_file = fileName.split("\\")[len(fileName.split("\\"))-1]
        target_fileName = os.getcwd() + "\\" + pathName + "\\" + target_file
        #if not os.path.exists(target_fileName):
        copyfile(src_file, target_fileName)
    except Exception as e:
        print(e)

def convertFloatToStr(str_number):
    try:
        return str(float(str_number))
    except ValueError:
        return str(str_number)

def convertIntToStr(str_number):
    try:
        return str(int(str_number))
    except ValueError:
        return str(str_number)

def getCellValueinString(cell):
    value = cell.value
    if value is not None:
        if value != "":
            strval = ""
            if isinstance(value, int):
                strval = convertIntToStr(value)
            elif isinstance(value, float):
                if value % 1 == 0:
                    strval = convertIntToStr(value)
                else:
                    strval = convertFloatToStr(value)
            elif isinstance(value, datetime.date):
                strval = str(value)
            else:
                strval = value
            
            return strval

    return ""
    

def getValueinString(value):
    if value is not None:
        if value != "":
            if str(value).find("_") < 0:
                return convertFloatToStr(value)
            else:
                return value

    return ""

def findFile(keyword,root, curfile):
    filelist=[]
    for root,dirs,files in os.walk(root):
        for name in files: 
            if name.startwith(keyword) >= 0:
                foundfile = os.path.join(root, name)
                curfile.append(name)
                return foundfile

    return None

def rgb_to_hex(rgb):
    bgr = (rgb[2],rgb[1],rgb[0])
    strValue = '%02x%02x%02x' % bgr
    iValue = int(strValue, 16)
    return iValue

def isEmptyValue(cell):
    if cell.Value is None:
        return True

    if str(cell.Value).strip() == "" :
        return True
    
    return False

def isContentValid(cell):
    if not isEmptyValue(cell):
        if str(cell.Value).strip() != "-" :
            return True
        else:
            return False
    else:
        return False

def CheckCellHasChild(find_uuid, specSheet):
    try:
        info = specSheet.UsedRange
        totalRow = info.Rows.Count
        rowNo = START_ROWNO
        while rowNo < totalRow:
            cellValue = specSheet.Cells(rowNo, UUID_COLNO).Value
            if cellValue is not None :
                if cellValue.find(find_uuid + "_state") >= 0:
                    return rowNo
            rowNo += specSheet.Cells(rowNo, UUID_COLNO).MergeArea.Rows.Count
        return -1
    except Exception as e:
        print(e)

def getFormulaStirng(cell1, cell2):
    formula_str = cell1.Value
    try:
        if isContentValid(cell1):
            part_row = cell1.MergeArea.Rows.Count
            row_asc = ord("A")
            replace_asc = 145
            replace_time = 0
            condition_dict = {}
            for i in range(part_row):
                if not isEmptyValue(cell2.MergeArea.Cells(i+1, 1)):                   
                    original_str= cell2.MergeArea.Cells(i+1, 1).Value
                    condition_dict[chr(row_asc + i)] = original_str
                else:
                    original_str = " "
                    condition_dict[chr(row_asc + i)] = original_str
    
                replace_time = replace_time + 1

            for i in range(replace_time):
                for j in range(len(condition_dict)):
                    condition_dict[chr(row_asc + j)] = condition_dict[chr(row_asc + j)].replace(chr(row_asc+i), chr(replace_asc+i))

            for i in range(replace_time):
                formula_str = formula_str.replace(chr(row_asc + i), condition_dict[chr(row_asc + i)])

            for i in range(replace_time):
                formula_str = formula_str.replace(chr(replace_asc + i), chr(row_asc + i))
    except Exception as e:
        print(e)

    return formula_str

def FindPartInAnotherSpec(keyValue, specSheet, keyFlag):
    try:
        if keyValue not in ("-","ー"):
            info = specSheet.UsedRange
            totalRow = info.Rows.Count
            rowNo = START_ROWNO
            while rowNo < totalRow:
                if keyFlag == "PART_ID":
                    if getCellValueinString(specSheet.Cells(rowNo, PART_ID_COLNO)) == keyValue :
                        return rowNo
                if keyFlag == "UUID_ID":
                    if getCellValueinString(specSheet.Cells(rowNo, UUID_COLNO)) == keyValue :
                        return rowNo
                if keyFlag == "PART_NAME":
                    if getCellValueinString(specSheet.Cells(rowNo, PART_NAME_COLNO)) == keyValue :
                        return rowNo
                if keyFlag == "DISP_NAME":
                    if getCellValueinString(specSheet.Cells(rowNo, DISPLAY_NAME_COLNO)) == keyValue :
                        return rowNo                        
                if keyFlag == "DISP_STR":
                    dsp_str = getFormulaStirng(specSheet.Cells(rowNo, DISPLAY_FORMULAR_NO), specSheet.Cells(rowNo, DISPLAY_CONDITION_NO))
                    if dsp_str == keyValue:
                        return rowNo
                if keyFlag == "UNION_STR":
                    dsp_str = getCellValueinString(specSheet.Cells(rowNo, PART_ID_COLNO)) + getCellValueinString(specSheet.Cells(rowNo, DISPLAY_NAME_COLNO))+getCellValueinString(specSheet.Cells(rowNo, PART_NAME_COLNO))
                    if dsp_str == keyValue:
                        return rowNo
                if keyFlag == "UNION_STR_MIN":
                    dsp_str = getCellValueinString(specSheet.Cells(rowNo, DISPLAY_NAME_COLNO))+getCellValueinString(specSheet.Cells(rowNo, PART_NAME_COLNO))
                    if dsp_str == keyValue:
                        return rowNo
                rowNo += specSheet.Cells(rowNo, PART_ID_COLNO).MergeArea.Rows.Count

        return -1
    except Exception as e:
        print(e)

def CollectPartInAnotherSpec(keyValue, specSheet, keyFlag, partlist, startRowNo):
    try:
        keystr = str(keyValue)
        if partlist.get(keystr) is not None:
           return partlist[keystr]

        if keystr not in ("-","ー"):
            info = specSheet.UsedRange
            totalRow = info.Rows.Count
            rowNo = START_ROWNO
            while rowNo < totalRow:
                part_id = getCellValueinString(specSheet.Cells(rowNo, PART_ID_COLNO))
                if partlist.get(part_id) is None:
                    partlist[part_id] = rowNo
                if keyFlag == "PART_ID":
                    if part_id == keystr :
                        return rowNo
                if keyFlag == "UUID_ID":
                    if specSheet.Cells(rowNo, UUID_COLNO).Value == keystr :
                        return rowNo

                if keyFlag == "UNION_STR":
                    union_str =  getCellValueinString(specSheet.Cells(rowNo, DISPLAY_NAME_COLNO)) + getCellValueinString(specSheet.Cells(rowNo, PART_TYPE_COLNO)) + getCellValueinString(specSheet.Cells(rowNo, PART_NAME_COLNO))
                    if union_str == keystr :
                        return rowNo

                rowNo += specSheet.Cells(rowNo, PART_ID_COLNO).MergeArea.Rows.Count

        return -1
    except Exception as e:
        print(e)

def getShapeInSheet(cellPosition, shapes):
    for shape in shapes:
        shape.Copy()
        print(str(shape.Top) + " " + str(shape.Left) + " " + str(shape.Width) + " " + str(shape.Height))
        if (cellPosition[0] <= shape.Left and ((shape.Top - cellPosition[1]) < cellPosition[3] and (shape.Top - cellPosition[1]) >=0)\
        and cellPosition[2] >= shape.Width and cellPosition[3] >= shape.Height ):
            return shape
            #shape.Copy()
    return None


class nodeObj(object):
    parent_node = None
    sibling_node = None
    child_node = None
    node_value = None
    node_level = -1
    node_content = ""
    node_extraVal = ""

    def get_parent(self):
        return self.parent_node
    
    def get_sibling(self):
        return self.sibling_node
    
    def get_child(self):
        return self.child_node

    def set_parent(self, node):
        self.parent_node = node
    
    def set_sibling(self, node):
        self.sibling_node = node
    
    def set_child(self, node):
        self.child_node = node

    def set_value(self, value):
        self.node_value = value

    def set_level(self, level):
        self.node_level = level
    
    def set_content(self, content):
        self.node_content = content
    
    def set_extraVal(self, extraVal):
        self.node_extraVal = extraVal

class nodeTree(object):
    root_node = None
    current_node = None

    def __init__(self, root_node):
        self.root_node = root_node

    def traversal(self):

        return 




if __name__ == '__main__':
    print("This is a common function package")