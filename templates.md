# UIBot 6.0 代码模板库

本文档提供常用场景的代码模板，可直接复制使用并根据需求修改。

## 目录

- [网页自动化模板](#网页自动化模板)
- [Excel处理模板](#excel处理模板)
- [文件处理模板](#文件处理模板)
- [数据采集模板](#数据采集模板)
- [邮件自动化模板](#邮件自动化模板)
- [错误处理模板](#错误处理模板)
- [工具函数库](#工具函数库)

---

## 网页自动化模板

### 模板1：基础网页操作
```vb
' 基础网页操作模板
Dim objBrowser

Try
    ' 启动浏览器
    objBrowser = WebBrowser.Create("chrome", "https://www.example.com", 30)
    Delay(2000)
    
    ' 等待页面加载
    If UiElement.Exists(@ui"目标元素", 10) Then
        ' 执行操作
        Mouse.Click(@ui"按钮", "left", "single", 0, 0)
        Delay(1000)
        
        ' 获取结果
        Dim sResult = UiElement.GetText(@ui"结果元素")
        TracePrint("结果: " & sResult)
    Else
        TracePrint("元素未找到")
    End If
    
Catch ex
    TracePrint("错误: " & ex.Message)
Finally
    ' 清理资源
    If objBrowser <> Null Then
        WebBrowser.Close(objBrowser)
    End If
End Try
```

### 模板2：表单填写提交
```vb
' 表单填写提交模板
Dim objBrowser

objBrowser = WebBrowser.Create("chrome", "https://form.example.com", 30)
Delay(2000)

' 填写表单
Keyboard.InputText(@ui"姓名输入框", "张三", True, False)
Delay(300)

Keyboard.InputText(@ui"邮箱输入框", "zhangsan@example.com", True, False)
Delay(300)

Keyboard.InputText(@ui"电话输入框", "13800138000", True, False)
Delay(300)

' 选择下拉框
Mouse.Click(@ui"城市下拉框", "left", "single", 0, 0)
Delay(500)
Mouse.Click(@ui"选项_北京", "left", "single", 0, 0)
Delay(300)

' 提交表单
Mouse.Click(@ui"提交按钮", "left", "single", 0, 0)
Delay(2000)

' 验证提交结果
If UiElement.Exists(@ui"成功提示", 5) Then
    TracePrint("提交成功")
Else
    TracePrint("提交失败")
End If

WebBrowser.Close(objBrowser)
```

### 模板3：使用JS操作网页
```vb
' JS操作网页模板
Dim objBrowser
Dim sResult

objBrowser = WebBrowser.Create("chrome", "https://www.example.com", 30)
Delay(2000)

' 执行JS获取数据
Dim jsCode = @"
(function() {
    let data = [];
    let items = document.querySelectorAll('.item');
    items.forEach(item => {
        data.push({
            title: item.querySelector('.title').innerText,
            price: item.querySelector('.price').innerText
        });
    });
    return JSON.stringify(data);
})();
"@

WebBrowser.RunJS(objBrowser, jsCode, sResult)

If sResult <> "" Then
    Dim arrData = JSON.Parse(sResult)
    TracePrint("获取到 " & Array.Length(arrData) & " 条数据")
End If

WebBrowser.Close(objBrowser)
```

---

## Excel处理模板

### 模板4：读取Excel数据
```vb
' 读取Excel数据模板
Dim objExcel
Dim sFilePath = "C:\data.xlsx"

objExcel = Excel.Open(sFilePath, True, "")

' 读取数据到数组
Dim arrData = []
Dim iRow = 2  ' 从第2行开始（跳过标题）

Do While True
    Dim sValue = Excel.GetCell(objExcel, "A" & iRow)
    If sValue = "" Then Exit Do
    
    Dim objRow = {
        "name": Excel.GetCell(objExcel, "A" & iRow),
        "age": Excel.GetCell(objExcel, "B" & iRow),
        "city": Excel.GetCell(objExcel, "C" & iRow)
    }
    Array.Push(arrData, objRow)
    
    iRow = iRow + 1
Loop

Excel.Close(objExcel)
TracePrint("读取完成，共 " & Array.Length(arrData) & " 条数据")
```

### 模板5：写入Excel数据
```vb
' 写入Excel数据模板
Dim objExcel
Dim arrData = [
    {"name": "张三", "age": 25, "city": "北京"},
    {"name": "李四", "age": 30, "city": "上海"},
    {"name": "王五", "age": 28, "city": "广州"}
]

objExcel = Excel.Create(True, "")

' 写入表头
Excel.SetCell(objExcel, "A1", "姓名")
Excel.SetCell(objExcel, "B1", "年龄")
Excel.SetCell(objExcel, "C1", "城市")

' 写入数据
Dim iRow = 2
For Each objItem In arrData
    Excel.SetCell(objExcel, "A" & iRow, objItem["name"])
    Excel.SetCell(objExcel, "B" & iRow, objItem["age"])
    Excel.SetCell(objExcel, "C" & iRow, objItem["city"])
    iRow = iRow + 1
Next

' 保存文件
Dim sFileName = "output_" & Time.Format(Time.Now(), "yyyyMMdd_HHmmss") & ".xlsx"
Excel.SaveAs(objExcel, "C:\" & sFileName)
Excel.Close(objExcel)

TracePrint("数据已保存到: C:\" & sFileName)
```

### 模板6：Excel数据处理
```vb
' Excel数据处理模板
Dim objExcel
objExcel = Excel.Open("C:\data.xlsx", True, "")

Dim iRow = 2
Dim iProcessed = 0

Do While True
    Dim sValue = Excel.GetCell(objExcel, "A" & iRow)
    If sValue = "" Then Exit Do
    
    ' 数据处理逻辑
    Dim sProcessed = String.Replace(sValue, "旧值", "新值")
    sProcessed = String.Trim(sProcessed)
    
    ' 写入处理结果
    Excel.SetCell(objExcel, "B" & iRow, sProcessed)
    Excel.SetCell(objExcel, "C" & iRow, "已处理")
    
    iProcessed = iProcessed + 1
    iRow = iRow + 1
Loop

Excel.Save(objExcel)
Excel.Close(objExcel)
TracePrint("处理完成，共 " & iProcessed & " 条")
```

---

## 文件处理模板

### 模板7：批量重命名文件
```vb
' 批量重命名文件模板
Dim sFolderPath = "C:\Files"
Dim arrFiles = File.GetFileList(sFolderPath, "*.txt", False)

Dim iCount = 1
For Each sFile In arrFiles
    Dim sExt = File.GetExtension(sFile)
    Dim sNewName = "文档_" & String.Format("{0:D3}", iCount) & sExt
    
    Try
        File.Rename(sFile, sNewName)
        TracePrint("重命名: " & File.GetName(sFile) & " -> " & sNewName)
        iCount = iCount + 1
    Catch ex
        TracePrint("重命名失败: " & sFile & " - " & ex.Message)
    End Try
Next

TracePrint("批量重命名完成，共 " & (iCount - 1) & " 个文件")
```

### 模板8：文件内容批量处理
```vb
' 文件内容批量处理模板
Dim sFolderPath = "C:\TextFiles"
Dim arrFiles = File.GetFileList(sFolderPath, "*.txt", False)

Dim sOldText = "旧内容"
Dim sNewText = "新内容"
Dim iProcessed = 0

For Each sFile In arrFiles
    Try
        ' 读取文件
        Dim sContent = File.Read(sFile, "utf-8")
        
        ' 检查并替换
        If InStr(sContent, sOldText) > 0 Then
            sContent = String.Replace(sContent, sOldText, sNewText)
            File.Write(sFile, sContent, "utf-8")
            TracePrint("已处理: " & File.GetName(sFile))
            iProcessed = iProcessed + 1
        End If
    Catch ex
        TracePrint("处理失败: " & sFile & " - " & ex.Message)
    End Try
Next

TracePrint("批量处理完成，共处理 " & iProcessed & " 个文件")
```

---

## 数据采集模板

### 模板9：网页列表数据采集
```vb
' 网页列表数据采集模板
Dim objBrowser
Dim arrData = []

objBrowser = WebBrowser.Create("chrome", "https://www.example.com/list", 30)
Delay(3000)

' 获取列表项
Dim objElements = UiElement.GetChildren(@ui"列表容器", 1)

For Each objElement In objElements
    Try
        Dim objItem = {
            "title": UiElement.GetText(UiElement.GetChild(objElement, ".title")),
            "price": UiElement.GetText(UiElement.GetChild(objElement, ".price")),
            "date": UiElement.GetText(UiElement.GetChild(objElement, ".date"))
        }
        Array.Push(arrData, objItem)
    Catch ex
        TracePrint("提取失败: " & ex.Message)
    End Try
Next

' 保存到Excel
Dim objExcel = Excel.Create(True, "")
Excel.SetCell(objExcel, "A1", "标题")
Excel.SetCell(objExcel, "B1", "价格")
Excel.SetCell(objExcel, "C1", "日期")

Dim iRow = 2
For Each objItem In arrData
    Excel.SetCell(objExcel, "A" & iRow, objItem["title"])
    Excel.SetCell(objExcel, "B" & iRow, objItem["price"])
    Excel.SetCell(objExcel, "C" & iRow, objItem["date"])
    iRow = iRow + 1
Next

Dim sFileName = "采集数据_" & Time.Format(Time.Now(), "yyyyMMdd_HHmmss") & ".xlsx"
Excel.SaveAs(objExcel, "C:\" & sFileName)
Excel.Close(objExcel)

WebBrowser.Close(objBrowser)
TracePrint("采集完成，共 " & Array.Length(arrData) & " 条数据")
```

---

## 邮件自动化模板

### 模板10：发送邮件报告
```vb
' 发送邮件报告模板
Dim sSmtpServer = "smtp.qq.com"
Dim iPort = 465
Dim sUsername = "your_email@qq.com"
Dim sPassword = "your_password"
Dim sFrom = "your_email@qq.com"
Dim sTo = "receiver@example.com"

' 生成报告内容
Dim sSubject = "自动化报告 - " & Time.Format(Time.Now(), "yyyy-MM-dd")
Dim sBody = "报告摘要:\n\n"
sBody = sBody & "执行时间: " & Time.Format(Time.Now(), "yyyy-MM-dd HH:mm:ss") & "\n"
sBody = sBody & "处理记录: 100 条\n"
sBody = sBody & "成功: 98 条\n"
sBody = sBody & "失败: 2 条\n\n"
sBody = sBody & "详细信息请查看附件。"

' 附件列表
Dim arrAttachments = ["C:\report.xlsx", "C:\log.txt"]

' 发送邮件
Try
    Mail.Send(sSmtpServer, iPort, sUsername, sPassword, _
              sFrom, sTo, sSubject, sBody, arrAttachments)
    TracePrint("邮件发送成功")
Catch ex
    TracePrint("邮件发送失败: " & ex.Message)
End Try
```

---

## 错误处理模板

### 模板11：完整错误处理框架
```vb
' 完整错误处理框架模板
Dim sLogFile = "C:\Logs\automation_" & Time.Format(Time.Now(), "yyyyMMdd") & ".log"

' 日志记录函数
Function WriteLog(sLevel, sMessage)
    Dim sTime = Time.Format(Time.Now(), "yyyy-MM-dd HH:mm:ss")
    Dim sLogLine = "[" & sTime & "] [" & sLevel & "] " & sMessage & "\n"
    File.Append(sLogFile, sLogLine, "utf-8")
    TracePrint(sLogLine)
End Function

WriteLog("INFO", "=== 流程开始 ===")

Try
    WriteLog("INFO", "步骤1: 初始化")
    ' 初始化代码
    
    WriteLog("INFO", "步骤2: 执行主要业务")
    ' 主要业务代码
    
    WriteLog("INFO", "步骤3: 保存结果")
    ' 保存结果代码
    
    WriteLog("INFO", "=== 流程成功完成 ===")
    
Catch ex
    WriteLog("ERROR", "发生错误: " & ex.Message)
    WriteLog("ERROR", "错误位置: " & ex.Source)
    
    ' 发送错误通知
    Try
        Mail.Send("smtp.qq.com", 465, "your_email@qq.com", "password", _
                  "your_email@qq.com", "admin@example.com", _
                  "自动化流程错误通知", _
                  "流程执行失败\n\n错误信息: " & ex.Message & "\n错误位置: " & ex.Source, _
                  [sLogFile])
    Catch mailEx
        WriteLog("ERROR", "邮件发送失败: " & mailEx.Message)
    End Try
    
Finally
    WriteLog("INFO", "=== 流程结束 ===")
    ' 清理资源
End Try
```

---

## 工具函数库

### 函数1：智能等待元素
```vb
' 智能等待元素函数
Function WaitForElement(objElement, iTimeout)
    Dim iWaited = 0
    Do While Not UiElement.Exists(objElement, 1)
        iWaited = iWaited + 1
        If iWaited >= iTimeout Then
            Return False
        End If
        Delay(1000)
    Loop
    Return True
End Function

' 使用示例
If WaitForElement(@ui"按钮", 30) Then
    Mouse.Click(@ui"按钮")
Else
    TracePrint("元素等待超时")
End If
```

### 函数2：安全点击
```vb
' 安全点击函数
Function SafeClick(objElement, iTimeout)
    If WaitForElement(objElement, iTimeout) Then
        Try
            Mouse.Click(objElement, "left", "single", 0, 0)
            Return True
        Catch ex
            TracePrint("点击失败: " & ex.Message)
            Return False
        End Try
    Else
        TracePrint("元素未找到")
        Return False
    End If
End Function

' 使用示例
SafeClick(@ui"登录按钮", 10)
```

### 函数3：安全输入
```vb
' 安全输入函数
Function SafeInput(objElement, sText, iTimeout)
    If WaitForElement(objElement, iTimeout) Then
        Try
            Keyboard.InputText(objElement, sText, True, False)
            Return True
        Catch ex
            TracePrint("输入失败: " & ex.Message)
            Return False
        End Try
    Else
        TracePrint("元素未找到")
        Return False
    End If
End Function

' 使用示例
SafeInput(@ui"用户名输入框", "admin", 10)
```

### 函数4：重试机制
```vb
' 重试机制函数
Function RetryAction(funcAction, iMaxRetry)
    Dim iRetry = 0
    Do While iRetry < iMaxRetry
        Try
            funcAction()
            Return True
        Catch ex
            iRetry = iRetry + 1
            TracePrint("重试 " & iRetry & "/" & iMaxRetry & ": " & ex.Message)
            If iRetry < iMaxRetry Then
                Delay(2000)
            End If
        End Try
    Loop
    Return False
End Function
```

### 函数5：数据验证
```vb
' 数据验证函数
Function ValidateData(sValue, sType)
    Select Case sType
        Case "email"
            Return String.Match(sValue, "^[\w\.-]+@[\w\.-]+\.\w+$")
        Case "phone"
            Return String.Match(sValue, "^1[3-9]\d{9}$")
        Case "number"
            Return IsNumeric(sValue)
        Case "notempty"
            Return sValue <> "" And Not IsNull(sValue)
        Case Else
            Return True
    End Select
End Function

' 使用示例
If ValidateData("test@example.com", "email") Then
    TracePrint("邮箱格式正确")
End If
```

### 函数6：格式化输出
```vb
' 格式化输出函数
Function FormatOutput(sTemplate, objData)
    Dim sResult = sTemplate
    For Each sKey In objData.Keys
        sResult = String.Replace(sResult, "{" & sKey & "}", CStr(objData[sKey]))
    Next
    Return sResult
End Function

' 使用示例
Dim objData = {"name": "张三", "age": 25, "city": "北京"}
Dim sOutput = FormatOutput("姓名: {name}, 年龄: {age}, 城市: {city}", objData)
TracePrint(sOutput)
```

---

## 使用说明

1. **选择模板**: 根据需求选择合适的模板
2. **修改参数**: 修改模板中的路径、URL、元素定位等参数
3. **测试运行**: 在测试环境中运行并调试
4. **优化调整**: 根据实际情况调整延时、重试次数等参数

## 注意事项

- 所有路径使用绝对路径
- 元素定位需要在 UIBot 中录制
- 延时时间根据实际网络和系统性能调整
- 敏感信息（密码、密钥）不要硬编码在代码中
- 生产环境使用前务必充分测试

---

**文档版本**: v1.0.0  
**更新时间**: 2024-01-15
