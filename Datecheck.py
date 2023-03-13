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
        if (cmdType.value not in [1.0,2.0,3.0,4.0,5.0,6.0,7.0,8.0,9.0]):
            print('第1列，第',i+1,"行,请输入正确的操作类型（数字）")
            checkCmd = False

        # 第2列 内容检查
        cmdValue = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value ==1.0 or cmdType.value == 2.0 or cmdType.value == 3.0 or cmdType.value == 7.0:
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

        # 截屏事件，内容必须为"0,0,200,300"的格式
        if cmdType.value == 8.0:
            if cmdValue.ctype != 1:
                print('第2列，第',i+1,"行,数据输入错误")
                checkCmd = False

        i += 1
    return checkCmd