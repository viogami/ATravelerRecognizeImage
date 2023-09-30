from tool import readregion
# 定义数据检查
# cmdType.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮 7.0 移动 8.0 截屏
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
        cmdType_str=["鼠标拖拽","单击","双击","右键","输入","等待","滚轮","移动","截图","键盘按键"]
        # 完成文字和数字代码的映射
        if cmdType.value in cmdType_str:
                cmdType.value=cmdType_str.index(cmdType.value)

        if (cmdType.value not in [0.0,1.0,2.0,3.0,4.0,5.0,6.0,7.0,8.0,9.0,10.0]):
            print('第1列，第',i+1,"行,请输入正确的操作类型")
            checkCmd = False

        # 第2列 内容检查
        cmdValue = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value in [0.0,1.0,2.0,3.0]:
            if cmdValue.ctype != 1:
                print('第2列，第',i+1,"行,数据输入错误")
                checkCmd = False

        # 输入类型，内容不能为空
        if cmdType.value == 4.0 or cmdType.value == 9.0:
            if cmdValue.ctype == 0:
                print('第2列，第',i+1,"行,数据输入错误")
                checkCmd = False

        # 等待/滚轮事件，内容必须为数字
        if cmdType.value == 5.0 or cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('第2列，第',i+1,"行,数据输入错误")
                checkCmd = False

        # 截屏事件，内容必须为"[0,0,200,300]"的格式
        if cmdType.value == 8.0:
            if cmdValue.ctype != 1:
                print('第2列，第',i+1,"行,数据输入错误")
                checkCmd = False
            else:
                if cmdValue.value[0]=="[" and cmdValue.value[-1]=="]":
                    # 使用逗号分割字符串，得到一个包含各部分的列表
                    parts = readregion(cmdValue.value)
                    if len(parts) == 4 and all(part.isdigit() for part in parts):
                        pass
                    else:
                        print('第2列，第',i+1,"行,内容必须类似为'[0,0,200,300]'的坐标格式")
                        checkCmd = False
                else:
                    print('第2列，第',i+1,"行,内容必须携带中括号")
                    checkCmd = False
        
        #第四列，可选参数验证
        cmdValue = sheet1.row(i)[3]
        if cmdType.value==0.0:
            if cmdValue.ctype == 0:
                print('第4列，第',i+1,"行,鼠标拖拽操作必须含有两个参数")
                checkCmd = False


        i += 1
    return checkCmd