import os
import time
from gooey import Gooey, GooeyParser
import mainwork,Datecheck,tool

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
       },
       {
        'name': '帮助',
           'items': [{
               'type': 'AboutDialog',
               'menuTitle': '使用方法',
               'name': 'how to use'},
               { 
                'type': 'AboutDialog',
               'menuTitle': '注意事项',
               'name': 'you should be remind'}
               ]
       }

       ])
def Atri():
    #制作图形化界面
    nowpath=os.getcwd()
    parser = GooeyParser(description="一个基于图像匹配的自动操作脚本") 
    parser.add_argument('Excelpath', metavar="Excel路径", widget="FileChooser",default=nowpath+"\\test1.xlsx")
    parser.add_argument('Imagepath', metavar="图片文件夹", widget="DirChooser",default=nowpath+"\\worktest")
    parser.add_argument('-loop', help="是否无限循环?",metavar="循环",widget="CheckBox",action="store_true")
    parser.add_argument('-skip', help="十次匹配不到图片跳过该指令",metavar="跳过",widget="CheckBox",action="store_true")
    parser.add_argument('-saveimage', help="截图保存文件夹位置",metavar="截图",widget="DirChooser",default=nowpath+"\\ScreenShot")

    mouseattribute=parser.add_argument_group("鼠标参数")
    mouseattribute.add_argument('-mouseinterval', help="鼠标点击时间间隔/s",metavar="intervaltime",type=float,default="0.2")
    mouseattribute.add_argument('-mouseduration', help="执行指令鼠标移动持续时间/s",metavar="durationtime",type=float,default="0.2")

 #   verbosity = parser.add_mutually_exclusive_group()
 #   verbosity.add_argument('-t', dest='顺次式判断',
 #                         action="store_true", help="匹配到指定图片顺次运行指令")
 #   verbosity.add_argument('-q', dest='跳过式判断',
 #                          action="store_true", help="匹配到图片则跳过指令")
    args = parser.parse_args()

    ##start##
    print('欢迎使用,萝卜子开始运行...')
    sheet1=tool.readexcel(args.Excelpath)
    #数据检查
    checkCmd = Datecheck.dataCheck(sheet1)
    if checkCmd:
        loopcount='a'
        if args.loop:
             #循环，无限运行
            while True:
                mainwork.mainWork(sheet1,args.Imagepath,args.skip,args.mouseinterval,args.mouseduration,args.saveimage,loopcount)
                loopcount=chr(ord(loopcount)+1)
                time.sleep(0.1)
                print("指令流程结束一次，休息0.1秒~")
        else:
            #不循环，运行一次
            mainwork.mainWork(sheet1,args.Imagepath,args.skip,args.mouseinterval,args.mouseduration,args.saveimage,loopcount)
            print("成功运行流程一次~")  
    else:
        print('Excel表格数据输入有误!')
    

if __name__ == '__main__':
    Atri()
##