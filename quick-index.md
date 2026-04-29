# UIBot 6.0 快速索引

本文档提供快速命令查询索引，帮助您快速定位所需命令。

## 按功能快速查找

### 🖱️ 鼠标操作
- **点击** → `Mouse.Click()` - 点击界面元素
- **移动** → `Mouse.MoveTo()` - 移动鼠标到指定位置
- **拖动** → `Mouse.Action("拖动")` - 拖动元素

### ⌨️ 键盘操作
- **输入文本** → `Keyboard.InputText()` - 输入文本内容
- **输入密码** → `Keyboard.InputPwd()` - 输入密码
- **按键** → `Keyboard.Press()` - 模拟按键

### 🌐 浏览器操作
- **启动浏览器** → `WebBrowser.Create()` - 创建浏览器实例
- **打开网页** → `WebBrowser.Navigate()` - 导航到指定URL
- **执行JS** → `WebBrowser.RunJS()` - 执行JavaScript代码
- **获取标题** → `WebBrowser.GetTitle()` - 获取页面标题
- **下载文件** → `WebBrowser.Download()` - 下载文件
- **关闭浏览器** → `WebBrowser.Close()` - 关闭浏览器

### 📊 Excel 操作
- **打开Excel** → `Excel.Open()` - 打开Excel文件
- **创建Excel** → `Excel.Create()` - 创建新Excel
- **读取单元格** → `Excel.GetCell()` - 读取单元格值
- **写入单元格** → `Excel.SetCell()` - 写入单元格值
- **保存** → `Excel.Save()` / `Excel.SaveAs()` - 保存文件
- **关闭** → `Excel.Close()` - 关闭Excel

### 📄 文件操作
- **读取文件** → `File.Read()` - 读取文件内容
- **写入文件** → `File.Write()` - 写入文件内容
- **追加内容** → `File.Append()` - 追加到文件末尾
- **判断存在** → `File.Exists()` - 判断文件是否存在
- **复制文件** → `File.Copy()` - 复制文件
- **删除文件** → `File.Delete()` - 删除文件
- **重命名** → `File.Rename()` - 重命名文件
- **获取文件列表** → `File.GetFileList()` - 获取目录下文件列表

### 🔤 字符串操作
- **分割** → `String.Split()` - 分割字符串
- **替换** → `String.Replace()` - 替换字符串
- **查找** → `String.IndexOf()` / `InStr()` - 查找子串位置
- **截取** → `String.Substring()` - 截取子串
- **拼接** → `String.Concat()` / `&` - 拼接字符串
- **去空格** → `String.Trim()` - 去除首尾空格

### 📋 数组操作
- **获取长度** → `Array.Length()` - 获取数组长度
- **添加元素** → `Array.Push()` - 添加到末尾
- **删除元素** → `Array.Remove()` - 删除指定元素
- **拼接数组** → `Array.Join()` - 将数组元素拼接成字符串
- **排序** → `Array.Sort()` - 数组排序

### 🪟 窗口操作
- **获取活动窗口** → `Window.GetActive()` - 获取当前活动窗口
- **设置活动** → `Window.SetActive()` - 激活指定窗口
- **关闭窗口** → `Window.Close()` - 关闭窗口
- **最大化** → `Window.Maximize()` - 最大化窗口
- **最小化** → `Window.Minimize()` - 最小化窗口

### 🎯 元素操作
- **判断存在** → `UiElement.Exists()` - 判断元素是否存在
- **获取文本** → `UiElement.GetText()` - 获取元素文本
- **设置文本** → `UiElement.SetText()` - 设置元素文本
- **获取属性** → `UiElement.GetAttribute()` - 获取元素属性
- **获取子元素** → `UiElement.GetChildren()` - 获取子元素列表

### 📧 邮件操作
- **发送邮件** → `Mail.Send()` - 发送邮件
- **连接邮箱** → `Mail.Connect()` - 连接邮箱服务器
- **获取邮件列表** → `Mail.GetMailList()` - 获取邮件列表
- **下载附件** → `Mail.DownloadAttachment()` - 下载附件

### 🗄️ 数据库操作
- **创建连接** → `DB.Create()` - 创建数据库连接
- **查询所有** → `DB.QueryAll()` - 查询所有结果
- **查询单条** → `DB.QueryOne()` - 查询单条记录
- **执行SQL** → `DB.Execute()` - 执行SQL语句
- **关闭连接** → `DB.Close()` - 关闭数据库连接

### 🔍 OCR 识别
- **图片识别** → `OCR.ImageOCR()` - 识别图片文字
- **屏幕识别** → `OCR.ScreenOCR()` - 识别屏幕区域文字
- **获取全部文本** → `OCR.GetAllText()` - 获取所有识别文本
- **查找文本** → `OCR.FindText()` - 查找指定文本位置
- **点击文本** → `OCR.ClickText()` - 点击识别到的文本

### 🌐 HTTP 操作
- **GET请求** → `HTTP.Get()` - 发送GET请求
- **POST请求** → `HTTP.Post()` - 发送POST请求
- **下载文件** → `HTTP.Download()` - 下载文件

### ⏰ 时间日期
- **获取当前时间** → `Time.Now()` - 获取当前时间
- **格式化时间** → `Time.Format()` - 格式化时间字符串
- **时间计算** → `Time.Add()` - 时间加减运算

### 💬 对话框
- **消息框** → `MsgBox()` - 显示消息框
- **输入框** → `InputBox()` - 显示输入框

### 🖥️ 系统操作
- **启动应用** → `App.Run()` - 启动应用程序
- **关闭应用** → `App.Kill()` - 关闭应用程序
- **执行命令** → `System.Exec()` - 执行系统命令

## 按场景快速查找

### 场景1：网页自动化
```
启动浏览器 → 打开网页 → 等待加载 → 输入内容 → 点击按钮 → 获取结果
WebBrowser.Create() → WebBrowser.Navigate() → Delay() → 
Keyboard.InputText() → Mouse.Click() → UiElement.GetText()
```

### 场景2：Excel数据处理
```
打开Excel → 读取数据 → 处理数据 → 写入结果 → 保存关闭
Excel.Open() → Excel.GetCell() → [处理逻辑] → 
Excel.SetCell() → Excel.Save() → Excel.Close()
```

### 场景3：文件批量处理
```
获取文件列表 → 遍历文件 → 读取内容 → 处理内容 → 保存结果
File.GetFileList() → For Each → File.Read() → 
[处理逻辑] → File.Write()
```

### 场景4：自动登录
```
打开网站 → 输入用户名 → 输入密码 → 点击登录 → 验证成功
WebBrowser.Create() → Keyboard.InputText() → 
Keyboard.InputPwd() → Mouse.Click() → UiElement.Exists()
```

### 场景5：数据采集
```
打开页面 → 定位元素 → 提取数据 → 保存到Excel
WebBrowser.Create() → UiElement.GetChildren() → 
UiElement.GetText() → Excel.SetCell()
```

## 常用代码片段

### 智能等待元素
```vb
Dim iMaxWait = 30
Dim iWaited = 0
Do While Not UiElement.Exists(@ui"目标元素", 1)
    iWaited = iWaited + 1
    If iWaited >= iMaxWait Then Exit Do
    Delay(1000)
Loop
```

### 错误处理模板
```vb
Try
    ' 业务代码
Catch ex
    TracePrint("错误: " & ex.Message)
    ' 错误处理
Finally
    ' 清理资源
End Try
```

### 日志记录函数
```vb
Function WriteLog(sMessage)
    Dim sTime = Time.Format(Time.Now(), "yyyy-MM-dd HH:mm:ss")
    Dim sLog = "[" & sTime & "] " & sMessage & "\n"
    File.Append("C:\log.txt", sLog, "utf-8")
End Function
```

## 命令速查表

| 功能 | 命令 | 说明 |
|------|------|------|
| 延时 | `Delay(ms)` | 延时指定毫秒 |
| 打印 | `TracePrint(msg)` | 输出调试信息 |
| 类型转换 | `CInt()` `CStr()` `CBool()` | 数据类型转换 |
| 判断类型 | `IsNull()` `IsEmpty()` | 判断变量状态 |

## 快速参考链接

- 详细命令说明 → [commands-reference.md](commands-reference.md)
- 实战代码示例 → [examples.md](examples.md)
- 完整命令手册 → [uibot.md](uibot.md)

---

**提示**: 使用 Ctrl+F 快速搜索关键词
