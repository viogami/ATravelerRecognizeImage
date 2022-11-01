import os
import xlrd
import pyautogui

#将鼠标移到屏幕的左上角，来抛出failSafeException异常
pyautogui.FAILSAFE = True

#定义判断图像是否存在 参数：img:图片位置或名称  similar:相似度
def imgexist(img,similar):
    location=pyautogui.locateCenterOnScreen(img,confidence=similar)
    if location is not None:
        return True
    elif location is None:
        return False

#定义匹配图片返回坐标的函数 参数：img:图片位置或名称  similar:相似度 region:指定区域
def imgloc(img,similar,region):
    location=pyautogui.locateCenterOnScreen(img,confidence=similar)
    if location is not None:
        return location


#定义读取excel文件的函数，返回sheet1的表格内容
def readexcel(filepath):
    #获得路径下的文件名称(带后缀)
    f=os.path.basename(filepath)
    #用xlrd库打开excel文件
    wb = xlrd.open_workbook(f)
    #通过索引获取表格sheet页，只需要sheet1即可
    sheet1 = wb.sheet_by_index(0)
    return sheet1