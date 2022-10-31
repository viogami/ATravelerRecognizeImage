import pyautogui
import time
import pyperclip

#定义鼠标事件
def mouseClick(clickTimes,lOrR,img,reTry):
    if reTry == 1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                break
            print("未匹配到图片"+img+",0.1秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
            print("未匹配到图片"+img+",0.1秒后重试")
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                print("0.1秒后重复执行一次！")
                i += 1
            time.sleep(0.1)
    #retry=0,意味着多次匹配不到可以跳过
    elif reTry==0:
        location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
        i=0
        while i<10:
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                break
            i=i+1
            print("未匹配到图片"+img+",0.2秒后重试")
            time.sleep(0.2)



#定义任务 对于参数有： 
# sheet1:字符型，Excel表格文件位置
# imgpath：字符型，图片文件位置
# skip：布尔型，是否retry=0

def mainWork(sheet1,imgpath,skip):
    i = 1
    while i < sheet1.nrows:
        #取本行指令的操作类型
        cmdType = sheet1.row(i)[0]
        if cmdType.value == 1.0:
            #取图片名称
            img = imgpath+"\\"+sheet1.row(i)[1].value
            reTry = 1
            if skip:
                reTry=0
            elif sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1,"left",img,reTry)
            print("单击左键",img)
        #2代表双击左键
        elif cmdType.value == 2.0:
            #取图片名称
            img = imgpath+"\\"+sheet1.row(i)[1].value
            #取重试次数
            reTry = 1
            if skip:
                reTry=0
            elif sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2,"left",img,reTry)
            print("左键双击",img)
        #3代表右键
        elif cmdType.value == 3.0:
            #取图片名称
            img = imgpath+"\\"+ sheet1.row(i)[1].value
            #取重试次数
            reTry = 1
            if skip:
                reTry=0
            elif sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1,"right",img,reTry)
            print("右键单击",img) 
        #4代表输入
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl','v')
            time.sleep(0.5)
            print("输入:",inputValue)                                        
        #5代表等待
        elif cmdType.value == 5.0:
            #取等待时间
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("等待",waitTime,"秒")
        #6代表滚轮
        elif cmdType.value == 6.0:
            #取滚动距离
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("滚轮滑动，距离：",int(scroll))  
        elif cmdType.value == 7.0:
            continue                  
        i += 1