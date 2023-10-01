import os
import time
from gooey import Gooey, GooeyParser
import mainwork_first,mainwork,Datecheck,tool

@Gooey(
        advanced=True,
        tabbed_groups=True,
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
               'version': 'v1.5',
               'website': 'https://github.com/Violetmail/ATravelerRecognizeImage'}]
       },
       {
        'name': '帮助',
           'items': [{
               'type': 'MessageDialog',
               'menuTitle': '使用方法',
               'caption': '使用方法',
               'message': '在commands.xlsx文件中配置每一步的指令，根据表格提示填入对应的内容，指令123456789，分别具备单击，双击，右键单击，输入，等待，滚动，移动，截图，键盘按键的功能。',},
               { 
                'type': 'MessageDialog',
               'menuTitle': '注意事项',
               'caption': '注意事项',
               'message': '- 文件夹必须放在英文路径下\n- python版本不要超过3.8，因为在3.9版本中改变了xlrd包的部分函数名，导致读取.xlsx文件会出现异常\n- 鼠标点击和移动默认为图片的中心位置',
               }
               ]
       }

       ])
def Atri():
    #制作图形化界面
    nowpath=os.getcwd()
    parser = GooeyParser(description="一个基于图像匹配的自动操作脚本") 
    parser.add_argument('Excelpath', metavar="指令文件路径", widget="FileChooser",default=nowpath+"\\commands.xlsx")
    parser.add_argument('Imagepath', metavar="图片文件夹", widget="DirChooser",default=nowpath+"\\TargetImage")
    
    parser.add_argument('-loop', help="是否无限循环?",metavar="循环",widget="CheckBox",action="store_true",default=True)
    parser.add_argument('-skip', help="十次匹配不到目标图片则跳过该指令",metavar="跳过",widget="CheckBox",action="store_true")
    parser.add_argument('-retrytime', help="单一指令重复间隔",metavar="重复间隔",widget="DecimalField",type=float,default="0.1")
    parser.add_argument('-saveimage', help="截图保存文件夹位置",metavar="截图",widget="DirChooser",default=nowpath+"\\ScreenShot")

    mouseattribute=parser.add_argument_group("鼠标参数")
    mouseattribute.add_argument('-mouseinterval', help="鼠标点击时间间隔/s",metavar="intervaltime",widget="DecimalField",type=float,default="0.2")
    mouseattribute.add_argument('-mouseduration', help="执行指令鼠标移动持续时间/s",metavar="durationtime",widget="DecimalField",type=float,default="0.2")

 #   browser= parser.add_argument_group("JavaScript运行设置")
 #   browser.add_argument('-chrome', dest='Chrome浏览器',action="store_true", help="js脚本的运行只能实现在Chrome浏览器中",default=True)
 #   browser.add_argument('-edge', dest='Edge浏览器',action="store_true", help="js脚本的运行只能实现在Edge浏览器中")
 #   browser.add_argument('jspath', metavar="JS脚本路径", widget="FileChooser",default=nowpath+"\\web-JS\\user.js")

 #   commands=parser.add_argument_group("新建指令")
 #   commands.add_argument('-commands_num', help="选择要添加的指令",metavar="操作",widget="Dropdown")
 #   commands.add_argument('-commands_image', help="选择图片或者附加值",metavar="操作参数",default=" ")
 #   commands.add_argument('-commands_repeat', help="指令是否重复",metavar=" ",widget="CheckBox",action="store_true")

    args = parser.parse_args()

    ##start##
    print('欢迎使用,萝卜子开始运行...')
    sheet1=tool.readexcel(args.Excelpath)
    #数据检查
    checkCmd = Datecheck.dataCheck(sheet1)

    if checkCmd:
        loopcount='A'
        if args.loop:
            #所有指令预先运行一次
            mainwork_first.mainWork(sheet1,args.Imagepath,args.skip,args.mouseinterval,args.mouseduration,args.saveimage,loopcount,args.retrytime)
            print("指令流程结束一次，休息0.1秒~")
            #循环，无限运行
            while True:
                mainwork.mainWork(sheet1,args.Imagepath,args.skip,args.mouseinterval,args.mouseduration,args.saveimage,loopcount,args.retrytime)
                loopcount=chr(ord(loopcount)+1)
                time.sleep(0.1)
                print("指令流程结束一次，休息0.1秒~")
        else:
            #不循环，运行一次
            mainwork_first.mainWork(sheet1,args.Imagepath,args.skip,args.mouseinterval,args.mouseduration,args.saveimage,loopcount,args.retrytime)
            print("成功运行流程一次~")  
    else:
        print('Excel表格数据输入有误,请检查指令表格')
    
if __name__ == '__main__':
    Atri()
##