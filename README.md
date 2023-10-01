# ATravelerRecognizeImage
图像识别旅者

| :exclamation: 免责声明<br>本脚本使用后果均为使用者个人承担<br>The consequence of using this script must be undertaken by user. |
| ------------------------------------------------------------------------------------------------------------------------- |

主要使用pyautogui的库实现，以及使用gooey库制作了GUI :grin:

## 功能
- [x] 支持图形化界面
  - [x] 自定义excel和图片路径
  - [x] 使用checkbox选择是否循环和跳过匹配不到的图片
  - [x] 设置菜单，显示关于文本
- [x] 指令功能
  - [x] 多次匹配不到一个指令可以跳过该指令
  - [x] 可以设置只执行一次的指令
- [x] 鼠标功能
  - [x] 单击
  - [x] 双击
  - [x] 右键
  - [x] 移动到指定图片的位置
  - [x] 移动到指定屏幕坐标
  - [x] 长按鼠标
  - [x] 长按鼠标拖拽
- [x] 输入功能
  - [x] 文本输入
  - [x] 键盘按键键入
  - [x] 按键可以长按
- [x] 等待功能
  - [x] 用户自定义每一步等待时长
- [x] 滚轮功能
- [x] 截图功能
  - [x] 截全屏
  - [x] 自定义范围
  - [x] 保存截图，命名为“字母+数字”组合
- [x] 用户自定义鼠标参数
- [x] 优化gui界面
- [x] 优化Excel表，添加了更加准确的数据检查。

## 界面预览

![sample](https://github.com/Violetmail/ATravelerRecognizeImage/assets/90465552/eb214b79-1b57-4796-9f6d-f1d360df02c0)

## 使用方法：
- 在commands.xlsx文件中配置每一步的指令，根据表格提示填入对应的内容(你也可以填入数字指令：0123456789，分别具备鼠标拖拽，单击，双击，右键单击，输入，等待，滚动，移动，截图，键盘按键的功能)。

- 关于重复：对于单一指令的重复，当重复过程中多次匹配失败则会跳出重复，执行下一条指令。
- 指令文件（commands.xlsx）表格添加了数据验证，方便指令填写。可以存在多个指令文件，每次选择相应的作业去完成。
- 点击ATRI.py文件运行，最好使用pyw，无黑窗启动。:heartpulse:

## 功能解释：
- 鼠标拖拽[0]:必须指定一个图片参数，使鼠标移动到该图片位置。如果可选参数是数字，进行鼠标长按操作，参数值为长按时间；可选参数为图片名，则进行拖拽操作，拖拽到可选参数的图片位置
- 左键单击[1]:鼠标单击
- 左键双击[2]:鼠标双击
- 右键单击[3]:鼠标右键单击
- 输入[4]:输入一段文本
- 等待[5]:time.sleep()
- 滚轮[6]:使用滚轮，正值前滚，负值后滚
- 移动鼠标[7]:移动鼠标到指定的图片位置。如果参数输入的不是图片名而是坐标，类似` [200,300]`,则会移动到坐标位置。
- 截图[8]：鼠标点击和移动默认为图片的中心位置，截图的区域指定内容必须类似为` [0,0,200,300]`'的坐标格式，分别代表矩形区域对角线的坐标
- 键盘按键[9]，默认为立即释放,如果可选参数不为空，则表示长按该按键的时间

## 效果演示
用户可以根据自定义的指令文件搭配出繁多的操作流程，接下来举例说明如何使用该程序打开浏览器，并输入一行js代码：

以下是指令表，指令类型数字代码和文字都可以。

| 指令类型 | 参数(图片名称.png、输入内容、等待时长/秒、截图范围) | 重复次数(-1代表只运行一次 0代表多次匹配不到可以跳过) | 可选参数(图片名称.png、时间/秒） | 备注               |
|------|-------------------------------|-------------------------------|---------------------|------------------|
| 2    | edge.png                      | -1                            |                     | 打开edge浏览器，只运行一次  |
| 1    | search.png                    | -1                            |                     | 点击搜索栏，只运行一次      |
| 4    | www.baidu.com                 | -1                            |                     | 输入百度网址，只运行一次     |
| 键盘按键 | enter                         |                               |                     | 输入回车             |
| 9    | F12                           | -1                            |                     | 进入开发者模式，只运行一次    |
| 1    | control.png                   |                               |                     | 点击控制台，准备输入代码     |
| 4    | console.log("Hello world!");  |                               |                     | 输入js代码           |
| 键盘按键 | enter                         |                               |                     | 输入回车，打印log       |
| 8    | [200,500,500,800]             |                               |                     | 区域截图，并保存         |
| 移动   | [200,500]                     |                               |                     | 将鼠标移动到指定位置       |
| 等待   | 5                             |                               |                     | 等待5秒，然后循环输入js    |
|      |                               |                               |                     |                  |


每一步的逻辑备注在注释里，这样就可以实现一直输入js，一直截图的效果，当然这只是一个简单的效果，更多复杂的搭配根据指令而异。

本人也用过作为游戏辅助程序用来挂机刷道具，用处多多，多多开发 :joy:

## 注意事项
- **文件夹必须放在英文路径下**
- python版本不要高于3.9，3.9版本改变了xlrd包中部分函数名的使用，在读取xlsx文件会异常。（解决方案：找到xlrd包的位置，打开`` xlsx.py ``文件,查找里面的`` getiterator``()函数，全部替换成`` iter()``，保存关闭后即可。）
- 注意coloerd包的版本，gooey库内用的colored包最高版本为1.4.4，colored包在2023.7进行更新到了2.0，而gooey还未适应。

## 安装依赖
在项目根目录shift+右键，打开Powershell窗口，执行：
```
pip install -r requirements.txt
```

## 更新日志
### v1.5
- 完善了最后的功能，比如拖拽鼠标，延迟按键等。
- 优化了excel，指令变为可选下拉栏。
- 优化了数据检查，为一些不合法参数添加了一些必要说明，比如截图的参数格式等。
- 丰富了预览文档，图示操作，并添加了一些使用技巧。
- 因为早期框架构思的不好，如今功能扩充困难，此为最终版本，不再做其他调整。

### v1.4
改变了单一指令重复执行的逻辑，如果在重复执行的过程中图片匹配失败10次，则会跳出循环，不再重复，进而执行下一条指令。防止卡壳。

实现了可以添加只执行一次的指令的功能。

### v1.3
没事做重新捡起来优化一下，改变了gui界面，优化了文件的格式和排布，增加了代码可读性。
同时增加了键盘按键的功能，不过目前只做了按下立即释放的功能，后续会增加长按(延迟释放)的功能。

### v1.2
添加了截图功能，可以自定义截图范围，图片保存的文件夹也可以自定义。
添加了自定义鼠标参数的功能，点击间隔和执行时鼠标移动持续时间，默认均为0.2

### v1.1
改了功能“7”为移动鼠标，判断功能后面再补。

主要添加了是否循环和匹配不到自动跳过的功能，图片也支持手动输入文件夹检索了。

### v1.0
这是一个基于图片匹配的自动操作脚本，可以实现部分软件操作的自动化。

我根据gooey库，制作了一个图形化界面，用来调用参数和用户选项。


