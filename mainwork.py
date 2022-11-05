import pyautogui
import time
import pyperclip
from tool import readregion

#将鼠标移到屏幕的左上角，来抛出failSafeException异常
pyautogui.FAILSAFE = True

#定义鼠标点击事件
def mouseClick(clickTimes,lOrR,img,reTry,intervaltime,durationtime):
    if reTry == 1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=intervaltime,duration=durationtime,button=lOrR)
                break
            print("未匹配到图片"+img+",0.1秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=intervaltime,duration=durationtime,button=lOrR)
            print("未匹配到图片"+img+",0.1秒后重试")
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=intervaltime,duration=durationtime,button=lOrR)
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

#定义鼠标移动事件
def mousemove(img,reTry):
    if reTry == 1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.moveTo(location.x,location.y,duration=0.2)
                break
            print("未匹配到图片"+img+",0.1秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.moveTo(location.x,location.y,duration=0.2)
            print("未匹配到图片"+img+",0.1秒后重试")
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.moveTo(location.x,location.y,duration=0.2)
                print("0.1秒后重复执行一次！")
                i += 1
            time.sleep(0.1)
    #retry=0,意味着多次匹配不到可以跳过
    elif reTry==0:
        location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
        i=0
        while i<10:
            if location is not None:
                pyautogui.moveTo(location.x,location.y,duration=0.2)
                break
            i=i+1
            print("未匹配到图片"+img+",0.2秒后重试")
            time.sleep(0.2)



#定义键盘事件
def keyboardClick(key):
    pyautogui.press(key)



#定义任务 对于参数有： 
# sheet1:字符型，Excel表格文件位置
# imgpath：字符型，图片文件位置
# skip：布尔型，是否使得retry=0

def mainWork(sheet1,imgpath,skip,time1,time2,saveimage,loopcount):
    imagecount=1
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
            mouseClick(1,"left",img,reTry,time1,time2)
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
            mouseClick(2,"left",img,reTry,time1,time2)
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
            mouseClick(1,"right",img,reTry,time1,time2)
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
            #取图片名称
            img = imgpath+"\\"+ sheet1.row(i)[1].value   
            #取重试次数
            reTry = 1
            if skip:
                reTry=0
            elif sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value 
            mousemove(img,reTry)
            print("鼠标移动到图片\n"+img+" 的位置") 
        elif cmdType.value == 8.0 :
           region=sheet1.row(i)[1].value
           if region=="":
            pyautogui.screenshot(saveimage+"\\screenshot_"+loopcount+str(imagecount)+".png")
            imagecount=imagecount+1
           else:
            regionlist=readregion(region)
            pyautogui.screenshot(saveimage+"\\screenshot_"+loopcount+str(imagecount)+".png",region=regionlist)
            imagecount=imagecount+1
           print("截图已经保存到"+saveimage) 

        i += 1