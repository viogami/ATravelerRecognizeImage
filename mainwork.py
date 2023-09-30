import pyautogui
import time
import pyperclip
from tool import readregion

#将鼠标移到屏幕的左上角，来抛出failSafeException异常
pyautogui.FAILSAFE = True

#定义鼠标点击事件
def mouseClick(clickTimes,lOrR,img,reTry,intervaltime,durationtime,retrytime):
    if reTry == 1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=intervaltime,duration=durationtime,button=lOrR)
                break
            print("未匹配到图片"+img+","+str(retrytime)+"秒后重试点击")
            time.sleep(retrytime)
    
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=intervaltime,duration=durationtime,button=lOrR)
                print(str(retrytime)+"秒后重复执行一次！")
            else:
                reTry=0
                mouseClick(clickTimes,lOrR,img,reTry,intervaltime,durationtime,retrytime)
                break

            i += 1
            time.sleep(retrytime)

    #retry=0,意味着多次匹配不到可以跳过
    elif reTry==0:
        location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
        i=0
        while i<10:
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                break
            i=i+1
            time.sleep(retrytime)
        print("10次未匹配到图片"+img+"，跳过该指令:")

    #-1表示只运行一次，运行后跳出循环
    elif reTry == -1:
        print("跳过该指令:")

#定义鼠标移动事件
def mousemove(img,reTry,retrytime,coor_x=0,coor_y=0):
    #如果有指定坐标
    if coor_x!=0 and coor_y!=0:
        location=pyautogui.position(coor_x,coor_y)
    else:
        location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
    
    if reTry ==1:
        while True:
            if location is not None:
                pyautogui.moveTo(location.x,location.y,duration=0.2)
                break
            print("未匹配到图片"+img+","+str(retrytime)+"秒后重试移动")
            time.sleep(retrytime)

    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            if location is not None:
                pyautogui.moveTo(location.x,location.y,duration=0.2)
                print(str(retrytime)+"秒后重复执行一次！")
            i += 1
            time.sleep(retrytime)
            
    #retry=0,意味着多次匹配不到可以跳过
    elif reTry==0:
        i=0
        while i<10:
            if location is not None:
                pyautogui.moveTo(location.x,location.y,duration=0.2)
                break
            i=i+1
            time.sleep(retrytime)
        print("10次未匹配到图片"+img+",跳过该指令:")
    
    #-1表示只运行一次，运行后跳出循环
    elif reTry == -1:
        print("跳过该指令:")


#定义鼠标拖拽事件
def drag_or_hold(img1,img2,retry,retrytime):
    # 查找 img1的位置
    img1_center= pyautogui.locateCenterOnScreen(img1,confidence=0.9)
    if retry==1:
        while True:
            if img1_center is not None:
                if isinstance(img2, float):
                    # 如果 img2 是数字，进行长按操作
                    pyautogui.moveTo(img1_center.x,img1_center.y)
                    pyautogui.mouseDown()
                    time.sleep(img2)  # 长按指定的秒数
                    pyautogui.mouseUp()
                    print("长按图片"+img1+","+str(img2)+"秒")
                    break
                else:
                    # 否则进行拖拽操作
                    img2_center = pyautogui.locateCenterOnScreen(img2,confidence=0.9)
                    pyautogui.moveTo(img1_center.x, img1_center.y)
                    pyautogui.mouseDown()
                    pyautogui.moveTo(img2_center.x,img2_center.y)
                    pyautogui.mouseUp()
                    print("从"+img1+"位置 拖拽鼠标至"+img2+"处")
                    break
            else:
                print("未匹配到图片"+img1+","+str(retrytime)+"秒后重试拖拽")
                time.sleep(retrytime)

    elif retry>1:
        i = 1
        while i < retry + 1:
            if img1_center is not None:
                if isinstance(img2, int):
                    # 如果 img2 是数字，进行长按操作
                    pyautogui.mouseDown(x=img1_center.x, y=img1_center.y)
                    time.sleep(img2)  # 长按指定的秒数
                    pyautogui.mouseUp(x=img1_center.x, y=img1_center.y)
                    print("长按图片"+img1+","+str(img2)+"秒")
                    break
                else:
                    # 否则进行拖拽操作
                    img2_center = pyautogui.locateCenterOnScreen(img2)
                    pyautogui.moveTo(img1_center.x, img1_center.y)
                    pyautogui.mouseDown()
                    pyautogui.moveTo(img2_center.x,img2_center.y)
                    pyautogui.mouseUp()
                    print("从"+img1+"位置 拖拽鼠标至"+img2+"处")
                    break
            else:
                print("重复一次，在"+str(retrytime)+"秒后重试")
                time.sleep(retrytime)

            i += 1
            time.sleep(retrytime)
    
    #retry=0,意味着多次匹配不到可以跳过
    elif retry==0:
        i=0
        while i<10:
            if img1_center is not None:
                if isinstance(img2, int):
                    # 如果 img2 是数字，进行长按操作
                    pyautogui.mouseDown(x=img1_center.x, y=img1_center.y)
                    time.sleep(img2)  # 长按指定的秒数
                    pyautogui.mouseUp(x=img1_center.x, y=img1_center.y)
                    print("长按图片"+img1+","+str(img2)+"秒")
                    break
                else:
                    # 否则进行拖拽操作
                    img2_center = pyautogui.locateCenterOnScreen(img2)
                    pyautogui.moveTo(img1_center.x, img1_center.y)
                    pyautogui.mouseDown()
                    pyautogui.moveTo(img2_center.x,img2_center.y)
                    pyautogui.mouseUp()
                    print("从"+img1+"位置 拖拽鼠标至"+img2+"处")
                    break
            else:
                print("未匹配到图片"+img1+","+str(retrytime)+"秒后重试")
                
            i=i+1
            time.sleep(retrytime)
        print("10次未匹配到图片"+img1+",跳过该指令:")
    
    #-1表示只运行一次，运行后跳出循环
    elif retry == -1:
        print("跳过该指令:")

#定义键盘事件
def keyboardClick(key,reTry,retrytime,delay=0):

    #-1表示只运行一次，运行后跳出循环
    if reTry == -1:
        print("跳过该指令:")
    else:
        for i in range(int(reTry)):
            if delay!=0:
                pyautogui.keyDown(key)
                time.sleep(delay)
                pyautogui.keyUp(key)
                print(f"长按按键{key}保持{delay}秒")
            else:
                pyautogui.press(key)
                time.sleep(retrytime)
                print("点击了一次，休息"+str(retrytime)+"秒")

def inputevent(inputvalue,retry):
    if retry==1:
        pyperclip.copy(inputvalue)
        pyautogui.hotkey('ctrl','v')
        time.sleep(0.5)

    elif retry>1:
        for i in range(int(retry)):
            pyperclip.copy(inputvalue)
            pyautogui.hotkey('ctrl','v')
            time.sleep(0.5)
        print(f"重复输入{inputvalue}一共{retry}次")
    elif retry ==-1:
        print("跳过该指令:")


#定义任务 对于参数有： 
# sheet1:字符型，Excel表格文件位置
# imgpath：字符型，图片文件位置
# skip：布尔型，是否使得retry=0

def mainWork(sheet1,imgpath,skip,time1,time2,saveimage,loopcount,retrytime):
    imagecount=1  #截图记数
    i = 1
    while i < sheet1.nrows:
        #取本行指令的操作类型
        cmdType = sheet1.row(i)[0]
        cmdType_str=["鼠标拖拽","单击","双击","右键","输入","等待","滚轮","移动","截图","键盘按键"]
        if cmdType.value in cmdType_str:
            cmdType.value=float(cmdType_str.index(cmdType.value))

        #1鼠标点击
        if cmdType.value == 1.0:
            #取图片名称
            img = imgpath+"\\"+sheet1.row(i)[1].value
            reTry = 1
            if skip:
                reTry=0
            elif sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1,"left",img,reTry,time1,time2,retrytime=retrytime)
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
            mouseClick(2,"left",img,reTry,time1,time2,retrytime=retrytime)
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
            mouseClick(1,"right",img,reTry,time1,time2,retrytime=retrytime)
            print("右键单击",img) 

        #4代表输入
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value
            #取重试次数
            reTry = 1
            if skip:
                reTry=0
            elif sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            inputevent(inputValue,reTry)
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

        
        #7代表移动鼠标
        elif cmdType.value == 7.0:  
            img=sheet1.row(i)[1].value   
            if "[" not in img:  #.png结尾则取图片名称
                img = imgpath+"\\"+ img 
                #取重试次数
                reTry = 1
                if skip:
                    reTry=0
                elif sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                    reTry = sheet1.row(i)[2].value 
                mousemove(img,reTry,retrytime)
                print("鼠标移动到图片\n"+img+" 的位置") 
            else:
                parts = readregion(img)  #否则读取坐标
                #取重试次数
                reTry = 1
                if skip:
                    reTry=0
                elif sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                    reTry = sheet1.row(i)[2].value 
                mousemove(img,reTry,retrytime,coor_x=parts[0],coor_y=parts[1])
                print(f"鼠标移动到({parts[0]},{parts[1]})的位置") 

        #8代表捕获截图
        elif cmdType.value == 8.0:
           region=sheet1.row(i)[1].value
           if region=="":
            pyautogui.screenshot(saveimage+"\\screenshot_"+loopcount+str(imagecount)+".png")
            imagecount=imagecount+1
           else:
            regionlist=readregion(region)
            pyautogui.screenshot(saveimage+"\\screenshot_"+loopcount+str(imagecount)+".png",region=regionlist)
            imagecount=imagecount+1
           print("截图已经保存到"+saveimage) 

        #9代表键盘按键（立即释放）,如果可选参数不为空，则表示长按该按键的时间
        elif cmdType.value == 9.0:
            aimkey=sheet1.row(i)[1].value
            delay=sheet1.row(i)[3].value
            reTry=1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            if delay=="":
                keyboardClick(aimkey,reTry,retrytime)
            else:
                keyboardClick(aimkey,reTry,retrytime,delay)
            print(f"按下{aimkey}键") 
        
        #0鼠标长按/拖拽
        elif cmdType.value ==0.0:
            #取图片名称
            img1 = imgpath+"\\"+sheet1.row(i)[1].value
            #可选参数如果是数则是长按
            if isinstance(sheet1.row(i)[3].value, float):
                img2=sheet1.row(i)[3].value
            else:
                img2 = imgpath+"\\"+sheet1.row(i)[3].value

            reTry = 1
            if skip:
                reTry=0
            elif sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            drag_or_hold(img1,img2,reTry,retrytime=retrytime)
            print("鼠标拖拽",img1)

        i += 1