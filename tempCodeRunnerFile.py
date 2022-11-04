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
    # 
    # 
    parser = GooeyParser(description="一个基于图像匹配的自动操作脚本") 
    parser.add_argument('Excelpath', metavar="Excel路径", widget="FileChooser",default="D:\GitHub\ATravelerRecognizeImage\\test1.xlsx")
    parser.add_argument('Imagepath', metavar="图片文件夹", widget="DirChooser",default="D:\GitHub\ATravelerRecognizeImage\\worktest")
    parser.add_argument('-loop', help="是否无限循环?",metavar="循环",widget="CheckBox",action="store_true")
    parser.add_argument('-skip', help="十次匹配不到图片跳过该图",metavar="跳过",widget="CheckBox",action="store_true")
    parser.add_argument('-mouseclicktime','--mouse', help="点击间隔",metavar="设置鼠标参数")
    
    parser.add_argument('-mouseduration', help="持续时间",metavar="设置鼠标参数")

    verbosity = parser.add_mutually_exclusive_group()
    verbosity.add_argument('-t', '--verbozze', dest='verbose',
                           action="store_true", help="Show more details")
    verbosity.add_argument('-q', '--quiet', dest='quiet',
                           action="store_true", help="Only output on error")
    args = parser.parse_args()

    ##start##
    print('欢迎使用,萝卜子开始运行...')
    sheet1=tool.readexcel(args.Excelpath)
    #数据检查
    checkCmd = Datecheck.dataCheck(sheet1)
    if checkCmd:
        if args.loop:
             #循环，无限运行
            while True:
                mainwork.mainWork(sheet1,args.Imagepath,args.skip)
                time.sleep(0.1)
                print("完整流程结束一次，休息0.1秒~")
        else:
            #不循环，运行一次
            mainwork.mainWork(sheet1,args.Imagepath,args.skip)
            print("成功运行流程一次~")  
    else:
        print('Excel表格数据输入有误!')
    

if __name__ == '__main__':
    Atri()
##