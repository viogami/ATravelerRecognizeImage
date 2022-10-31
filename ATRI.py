import os
import pyautogui
import time
import xlrd
from gooey import Gooey, GooeyParser
import mainwork

# 定义数据检查
# cmdType.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮 
# ctype     空：0
#           字符串：1
#           数字：2
#           日期：3
#           布尔：4
#           error：5
def dataCheck(sheet1):
    checkCmd = True
    #行数检查
    if sheet1.nrows<2:
        print("没有输入操作内容")
        checkCmd = False
    #每行数据检查
    i = 1
    while i < sheet1.nrows:
        # 第1列 操作类型检查
        cmdType = sheet1.row(i)[0]
        if (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0 
            and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0
            and cmdType.value !=7.0):
            print('第1列，第',i+1,"行,请输入正确的操作类型（数字）")
            checkCmd = False
        # 第2列 内容检查
        cmdValue = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value ==1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
                print('第2列，第',i+1,"行,数据输入错误")
                checkCmd = False
        # 输入类型，内容不能为空
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('第2列，第',i+1,"行,数据输入错误")
                checkCmd = False
        # 等待类型，内容必须为数字
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print('第2列，第',i+1,"行,数据输入错误")
                checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('第2列，第',i+1,"行,数据输入错误")
                checkCmd = False
        i += 1
    return checkCmd

     


#定义判断图像是否存在 参数：img:图片位置或名称  confidence:相似度
def imgexist(img,similar):
    location=pyautogui.locateCenterOnScreen(img,confidence=similar)
    if location is not None:
        return True
    elif location is None:
        return False

#定义读取excel文件的函数，返回sheet1的表格内容
def readexcel(filepath):
    #获得路径下的文件名称(带后缀)
    f=os.path.basename(filepath)
    #用xlrd库打开excel文件
    wb = xlrd.open_workbook(f)
    #通过索引获取表格sheet页，只需要sheet1即可
    sheet1 = wb.sheet_by_index(0)
    return sheet1

@Gooey(
    richtext_controls=True,  # 打开终端对颜色支持
    program_name="ATravelerRecognizeImage",  # 程序名称
    encoding="utf-8",  # 设置编码格式，打包的时候遇到问题
    #设置菜单
    menu=[{
           'name': '菜单',
           'items': [{
               'type': 'AboutDialog',
               'menuTitle': '关于',
               'name': 'ATravelerRecongnizeImage',
               'description': '作者:kagami',
               'version': 'v1.1',
               'website': 'https://github.com/Violetmail/ATravelerRecognizeImage'}]
       }
       ])
def main():
    #设置图形化界面
    # ,default="D:\GitHub\ATravelerRecognizeImage\\test1.xlsx"
    # ,default="D:\GitHub\ATravelerRecognizeImage\\image"
    parser = GooeyParser(description="基于图像匹配的自动操作脚本") 
    parser.add_argument('Excelpath', metavar="Excel路径", widget="FileChooser")
    parser.add_argument('Imagepath', metavar="图片文件夹", widget="DirChooser")
    parser.add_argument('-loop', help="是否无限循环?",metavar="循环",widget="CheckBox",action="store_true")
    parser.add_argument('-skip', help="十次未匹配不到图片跳过该图",metavar="跳过",widget="CheckBox",action="store_true")
    args = parser.parse_args()

    ##start##
    print('欢迎使用,萝卜子开始运行...')
    sheet1=readexcel(args.Excelpath)
    #数据检查
    checkCmd = dataCheck(sheet1)
    if checkCmd:
        if args.loop:
             #循环，无限运行
            while True:
                mainwork.mainWork(sheet1,args.Imagepath,args.skip)
                time.sleep(0.1)
                print("等待0.1秒")
        else:
            #不循环，运行一次
            mainwork.mainWork(sheet1,args.Imagepath,args.skip)
            print("成功运行流程一次！")  
    else:
        print('Excel表格数据输入有误!')
    

if __name__ == '__main__':
    main()
##