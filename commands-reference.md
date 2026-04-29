# UIBot 6.0 命令详细参考

本文档包含 UIBot 6.0 所有命令的详细说明、参数、返回值和示例代码。

## 目录

- [基本命令](#基本命令)
- [鼠标键盘](#鼠标键盘)
- [界面操作](#界面操作)
- [智能文档](#智能文档)
- [软件自动化](#软件自动化)
- [数据处理](#数据处理)
- [文件处理](#文件处理)
- [系统操作](#系统操作)
- [网络](#网络)
- [机器人指挥官](#机器人指挥官)

---

## 基本命令

### 数据转换

#### 转为逻辑数据
将数据转换为逻辑类型（True/False）

**语法**: `bRet = CBool(data)`

**参数**:
- data: 要转换的数据

**返回**: 布尔值

**示例**:
```vb
bRet = CBool(1)  ' 返回 True
bRet = CBool(0)  ' 返回 False
```

#### 转为整数数据
将数据转换为整数类型

**语法**: `iRet = CInt(data)`

**参数**:
- data: 要转换的数据

**返回**: 整数

**示例**:
```vb
iRet = CInt("123")  ' 返回 123
iRet = CInt(3.14)   ' 返回 3
```

#### 转为小数数据
将数据转换为小数（浮点数）类型

**语法**: `fRet = CFloat(data)`

**参数**:
- data: 要转换的数据

**返回**: 浮点数

**示例**:
```vb
fRet = CFloat("3.14")  ' 返回 3.14
fRet = CFloat(5)       ' 返回 5.0
```

#### 转为文字数据
将数据转换为字符串类型

**语法**: `sRet = CStr(data)`

**参数**:
- data: 要转换的数据

**返回**: 字符串

**示例**:
```vb
sRet = CStr(123)    ' 返回 "123"
sRet = CStr(True)   ' 返回 "True"
```

### 流程控制

#### 延时
延时等待指定毫秒后继续执行

**语法**: `Delay(ms)`

**参数**:
- ms: 延时时间（毫秒）

**示例**:
```vb
Delay(1000)  ' 延时 1 秒
Delay(500)   ' 延时 0.5 秒
```

#### 注释
用于给代码添加注释说明

**语法**: `Rem 注释内容` 或 `' 注释内容`

**示例**:
```vb
Rem 这是一条注释
' 这也是一条注释
```

---

## 鼠标键盘

### 鼠标操作

#### 点击目标
单击指定的界面元素

**语法**: `Mouse.Click(objElement, button, clickType, x, y)`

**参数**:
- objElement: 目标元素
- button: 鼠标按钮（"left"/"right"/"middle"）
- clickType: 点击类型（"single"/"double"）
- x, y: 相对坐标偏移

**示例**:
```vb
' 左键单击
Mouse.Click(@ui"按钮", "left", "single", 0, 0)

' 右键单击
Mouse.Click(@ui"按钮", "right", "single", 0, 0)

' 双击
Mouse.Click(@ui"按钮", "left", "double", 0, 0)
```

#### 模拟点击
模拟鼠标在指定坐标点击

**语法**: `Mouse.Action(action, x, y, button, clickType)`

**参数**:
- action: 操作类型（"点击"/"移动"/"拖动"）
- x, y: 屏幕坐标
- button: 鼠标按钮
- clickType: 点击类型

**示例**:
```vb
' 在坐标 (100, 200) 点击
Mouse.Action("点击", 100, 200, "left", "single")
```

#### 移动到目标上
将鼠标移动到指定元素上

**语法**: `Mouse.MoveTo(objElement, x, y)`

**参数**:
- objElement: 目标元素
- x, y: 相对坐标偏移

**示例**:
```vb
Mouse.MoveTo(@ui"按钮", 0, 0)
```

### 键盘操作

#### 输入文本
在指定元素中输入文本

**语法**: `Keyboard.InputText(objElement, text, clearBefore, sendEnter)`

**参数**:
- objElement: 目标元素
- text: 要输入的文本
- clearBefore: 输入前是否清空（True/False）
- sendEnter: 输入后是否按回车（True/False）

**示例**:
```vb
' 输入文本
Keyboard.InputText(@ui"输入框", "Hello World", True, False)

' 输入后按回车
Keyboard.InputText(@ui"搜索框", "UIBot", True, True)
```

#### 输入密码
在指定元素中输入密码

**语法**: `Keyboard.InputPwd(objElement, password, clearBefore)`

**参数**:
- objElement: 目标元素
- password: 密码
- clearBefore: 输入前是否清空

**示例**:
```vb
Keyboard.InputPwd(@ui"密码框", "mypassword", True)
```

#### 模拟按键
模拟键盘按键

**语法**: `Keyboard.Press(key, times)`

**参数**:
- key: 按键名称（如 "enter", "tab", "esc"）
- times: 按键次数

**示例**:
```vb
Keyboard.Press("enter", 1)  ' 按回车
Keyboard.Press("tab", 2)    ' 按两次 Tab
Keyboard.Press("esc", 1)    ' 按 ESC
```

---

## 界面操作

### 元素操作

#### 判断元素是否存在
判断界面元素是否存在

**语法**: `bRet = UiElement.Exists(objElement, timeout)`

**参数**:
- objElement: 目标元素
- timeout: 超时时间（秒）

**返回**: True/False

**示例**:
```vb
If UiElement.Exists(@ui"按钮", 5) Then
    TracePrint("元素存在")
Else
    TracePrint("元素不存在")
End If
```

#### 获取元素文本
获取元素的文本内容

**语法**: `sText = UiElement.GetText(objElement)`

**参数**:
- objElement: 目标元素

**返回**: 文本内容

**示例**:
```vb
sText = UiElement.GetText(@ui"标签")
TracePrint(sText)
```

#### 设置元素文本
设置元素的文本内容

**语法**: `UiElement.SetText(objElement, text)`

**参数**:
- objElement: 目标元素
- text: 要设置的文本

**示例**:
```vb
UiElement.SetText(@ui"输入框", "新内容")
```

#### 获取元素属性
获取元素的属性值

**语法**: `sValue = UiElement.GetAttribute(objElement, attrName)`

**参数**:
- objElement: 目标元素
- attrName: 属性名称

**返回**: 属性值

**示例**:
```vb
sValue = UiElement.GetAttribute(@ui"按钮", "name")
sValue = UiElement.GetAttribute(@ui"输入框", "value")
```

### 窗口管理

#### 获取活动窗口
获取当前活动窗口

**语法**: `objWindow = Window.GetActive()`

**返回**: 窗口对象

**示例**:
```vb
objWindow = Window.GetActive()
TracePrint(objWindow)
```

#### 设置活动窗口
将指定窗口设置为活动状态

**语法**: `Window.SetActive(objWindow)`

**参数**:
- objWindow: 窗口对象

**示例**:
```vb
Window.SetActive(@ui"窗口_记事本")
```

#### 关闭窗口
关闭指定窗口

**语法**: `Window.Close(objWindow)`

**参数**:
- objWindow: 窗口对象

**示例**:
```vb
Window.Close(@ui"窗口_记事本")
```

---

## 软件自动化

### 浏览器操作

#### 启动新的浏览器
启动一个新的浏览器

**语法**: `objBrowser = WebBrowser.Create(browserType, url, timeout)`

**参数**:
- browserType: 浏览器类型（"chrome"/"ie"/"firefox"/"edge"）
- url: 要打开的网址
- timeout: 超时时间（秒）

**返回**: 浏览器对象

**示例**:
```vb
objBrowser = WebBrowser.Create("chrome", "https://www.baidu.com", 30)
```

#### 打开网页
在浏览器中打开指定网址

**语法**: `WebBrowser.Navigate(objBrowser, url, timeout)`

**参数**:
- objBrowser: 浏览器对象
- url: 网址
- timeout: 超时时间

**示例**:
```vb
WebBrowser.Navigate(objBrowser, "https://www.google.com", 30)
```

#### 执行JS
在浏览器中执行 JavaScript 代码

**语法**: `WebBrowser.RunJS(objBrowser, jsCode, result)`

**参数**:
- objBrowser: 浏览器对象
- jsCode: JavaScript 代码
- result: 返回结果

**示例**:
```vb
' 获取页面标题
WebBrowser.RunJS(objBrowser, "document.title", result)

' 点击按钮
WebBrowser.RunJS(objBrowser, "document.getElementById('btn').click()", result)
```

---

## 数据处理

### 字符串操作

#### 分割字符串
使用分隔符分割字符串

**语法**: `arrRet = String.Split(text, delimiter)`

**参数**:
- text: 要分割的字符串
- delimiter: 分隔符

**返回**: 字符串数组

**示例**:
```vb
arrRet = String.Split("a,b,c", ",")
' 返回 ["a", "b", "c"]
```

#### 替换字符串
替换字符串中的内容

**语法**: `sRet = String.Replace(text, oldStr, newStr)`

**参数**:
- text: 原字符串
- oldStr: 要替换的内容
- newStr: 替换为的内容

**返回**: 替换后的字符串

**示例**:
```vb
sRet = String.Replace("Hello World", "World", "UIBot")
' 返回 "Hello UIBot"
```

### 数组操作

#### 获取数组长度
获取数组的元素数量

**语法**: `iLen = Array.Length(arr)`

**参数**:
- arr: 数组

**返回**: 数组长度

**示例**:
```vb
Dim arr = [1, 2, 3, 4, 5]
iLen = Array.Length(arr)  ' 返回 5
```

#### 在数组尾部添加元素
在数组末尾添加元素

**语法**: `Array.Push(arr, element)`

**参数**:
- arr: 数组
- element: 要添加的元素

**示例**:
```vb
Dim arr = [1, 2, 3]
Array.Push(arr, 4)
' arr 变为 [1, 2, 3, 4]
```

---

## 文件处理

### 文件操作

#### 读取文件
读取文件内容

**语法**: `sContent = File.Read(filePath, encoding)`

**参数**:
- filePath: 文件路径
- encoding: 编码格式（"utf-8"/"gbk"等）

**返回**: 文件内容

**示例**:
```vb
sContent = File.Read("C:\test.txt", "utf-8")
TracePrint(sContent)
```

#### 写入文件
写入内容到文件

**语法**: `File.Write(filePath, content, encoding)`

**参数**:
- filePath: 文件路径
- content: 要写入的内容
- encoding: 编码格式

**示例**:
```vb
File.Write("C:\output.txt", "Hello World", "utf-8")
```

#### 判断文件是否存在
判断文件是否存在

**语法**: `bRet = File.Exists(filePath)`

**参数**:
- filePath: 文件路径

**返回**: True/False

**示例**:
```vb
If File.Exists("C:\test.txt") Then
    TracePrint("文件存在")
End If
```

---

## 系统操作

### 应用程序管理

#### 启动应用程序
启动一个应用程序

**语法**: `iPID = App.Run(appPath, waitType, showType)`

**参数**:
- appPath: 应用程序路径
- waitType: 等待方式（0=不等待, 1=等待准备好, 2=等待退出）
- showType: 显示样式（0=隐藏, 1=默认, 3=最大化, 6=最小化）

**返回**: 进程 PID

**示例**:
```vb
iPID = App.Run("C:\Windows\notepad.exe", 1, 1)
```

#### 关闭应用
强制停止应用程序

**语法**: `App.Kill(processName)`

**参数**:
- processName: 进程名或 PID

**示例**:
```vb
App.Kill("notepad.exe")
App.Kill(1234)  ' 使用 PID
```

### 对话框

#### 消息框
弹出消息提示对话框

**语法**: `iRet = MsgBox(message, buttons, title)`

**参数**:
- message: 消息内容
- buttons: 按钮类型（0=确定, 1=确定/取消, 2=是/否等）
- title: 标题

**返回**: 用户点击的按钮

**示例**:
```vb
iRet = MsgBox("操作完成", 0, "提示")

iRet = MsgBox("是否继续？", 4, "确认")
If iRet = 6 Then  ' 6 表示点击了"是"
    TracePrint("用户选择继续")
End If
```

---

## 网络

### HTTP 操作

#### Get 获取数据
发送 HTTP GET 请求

**语法**: `HTTP.Get(url, headers, result)`

**参数**:
- url: 请求地址
- headers: 请求头（字典）
- result: 返回结果

**示例**:
```vb
Dim headers = {"User-Agent": "UIBot"}
HTTP.Get("https://api.example.com/data", headers, result)
TracePrint(result)
```

### 邮件操作

#### 发送邮件
发送邮件到指定邮箱

**语法**: `Mail.Send(smtpServer, port, username, password, from, to, subject, body, attachments)`

**参数**:
- smtpServer: SMTP 服务器
- port: 端口号
- username: 用户名
- password: 密码
- from: 发件人
- to: 收件人
- subject: 主题
- body: 正文
- attachments: 附件列表

**示例**:
```vb
Mail.Send("smtp.qq.com", 465, "user@qq.com", "password", _
          "user@qq.com", "receiver@example.com", _
          "测试邮件", "这是邮件正文", [])
```

---

## 最佳实践

### 1. 错误处理
始终使用 Try-Catch 处理可能出错的操作：

```vb
Try
    ' 可能出错的代码
    Mouse.Click(@ui"按钮")
Catch ex
    TracePrint("发生错误: " & ex.Message)
End Try
```

### 2. 等待元素
在操作元素前先等待元素出现：

```vb
If UiElement.Exists(@ui"按钮", 10) Then
    Mouse.Click(@ui"按钮")
Else
    TracePrint("元素未找到")
End If
```

### 3. 适当延时
在操作之间添加适当的延时：

```vb
Mouse.Click(@ui"按钮1")
Delay(500)  ' 等待 0.5 秒
Mouse.Click(@ui"按钮2")
```

### 4. 资源释放
使用完对象后及时释放：

```vb
' 关闭浏览器
WebBrowser.Close(objBrowser)

' 关闭 Excel
Excel.Close(objExcel)
```

---

**文档版本**: v1.0.0  
**适用版本**: UIBot 6.0.0.211215(64位)  
**更新时间**: 2024-01-15
