# UIBot 6.0 实战示例集

本文档包含 UIBot 6.0 的实战应用示例，涵盖常见的自动化场景。

## 目录

- [网页自动化](#网页自动化)
- [Excel 数据处理](#excel-数据处理)
- [文件批量处理](#文件批量处理)
- [OCR 识别应用](#ocr-识别应用)
- [邮件自动化](#邮件自动化)
- [数据库操作](#数据库操作)
- [综合应用](#综合应用)

---

## 网页自动化

### 示例1：自动登录网站

```vb
' 自动登录示例
Dim objBrowser

' 启动浏览器并打开登录页面
objBrowser = WebBrowser.Create("chrome", "https://www.example.com/login", 30)
Delay(2000)

' 输入用户名
Keyboard.InputText(@ui"用户名输入框", "your_username", True, False)
Delay(500)

' 输入密码
Keyboard.InputPwd(@ui"密码输入框", "your_password", True)
Delay(500)

' 点击登录按钮
Mouse.Click(@ui"登录按钮", "left", "single", 0, 0)
Delay(3000)

' 验证登录成功
Dim sTitle
WebBrowser.GetTitle(objBrowser, sTitle)
If InStr(sTitle, "首页") > 0 Then
    TracePrint("登录成功")
Else
    TracePrint("登录失败")
End If
```

### 示例2：批量下载文件

```vb
' 批量下载文件示例
Dim objBrowser
Dim arrUrls = ["https://example.com/file1.pdf", _
               "https://example.com/file2.pdf", _
               "https://example.com/file3.pdf"]

objBrowser = WebBrowser.Create("chrome", "about:blank", 30)

For Each url In arrUrls
    ' 下载文件
    WebBrowser.Download(objBrowser, url, "C:\Downloads", 60)
    TracePrint("已下载: " & url)
    Delay(2000)
Next

WebBrowser.Close(objBrowser)
TracePrint("所有文件下载完成")
```

### 示例3：网页数据采集

```vb
' 网页数据采集示例
Dim objBrowser
Dim arrData = []

objBrowser = WebBrowser.Create("chrome", "https://www.example.com/products", 30)
Delay(3000)

' 获取所有商品名称
Dim objElements = UiElement.GetChildren(@ui"商品列表", 1)
For Each objElement In objElements
    Dim sName = UiElement.GetText(objElement)
    Array.Push(arrData, sName)
    TracePrint("商品: " & sName)
Next

' 保存到文件
Dim sContent = Array.Join(arrData, "\n")
File.Write("C:\products.txt", sContent, "utf-8")

WebBrowser.Close(objBrowser)
TracePrint("数据采集完成，共 " & Array.Length(arrData) & " 条")
```

### 示例4：BOSS直聘职位信息抓取（使用JS）

```vb
// BOSS直聘职位信息抓取脚本
Dim objBrowser
Dim sUrl = "https://www.zhipin.com/web/geek/jobs?city=101010100&experience=104&query=%E8%87%AA%E5%8A%A8%E5%8C%96%E5%B7%A5%E7%A8%8B%E5%B8%88%20RPA"
Dim arrJobs = []

// 启动 Edge 浏览器
TracePrint("启动浏览器...")
objBrowser = WebBrowser.Create("edge", sUrl, 30)
Delay(5000)

// 使用 JS 获取所有职位信息
TracePrint("开始抓取数据...")
Dim jsCode = @"
(function() {
    let jobs = [];
    let jobCards = document.querySelectorAll('.job-card-wrapper');
    
    jobCards.forEach(card => {
        let job = {};
        let titleEl = card.querySelector('.job-name');
        job.title = titleEl ? titleEl.innerText.trim() : '';
        let salaryEl = card.querySelector('.salary');
        job.salary = salaryEl ? salaryEl.innerText.trim() : '';
        let companyEl = card.querySelector('.company-name');
        job.company = companyEl ? companyEl.innerText.trim() : '';
        let tagsEl = card.querySelectorAll('.company-tag-list li');
        job.tags = Array.from(tagsEl).map(tag => tag.innerText.trim()).join(' | ');
        let locationEl = card.querySelector('.job-area');
        job.location = locationEl ? locationEl.innerText.trim() : '';
        let expEl = card.querySelector('.job-limit .tag-list li:nth-child(1)');
        job.experience = expEl ? expEl.innerText.trim() : '';
        let eduEl = card.querySelector('.job-limit .tag-list li:nth-child(2)');
        job.education = eduEl ? eduEl.innerText.trim() : '';
        let linkEl = card.querySelector('.job-card-left a');
        job.link = linkEl ? 'https://www.zhipin.com' + linkEl.getAttribute('href') : '';
        let hrEl = card.querySelector('.info-public');
        job.hr = hrEl ? hrEl.innerText.trim() : '';
        jobs.push(job);
    });
    return JSON.stringify(jobs);
})();
"@

Dim sResult
WebBrowser.RunJS(objBrowser, jsCode, sResult)

// 解析 JSON 结果
If sResult <> "" Then
    arrJobs = JSON.Parse(sResult)
    TracePrint("成功抓取 " & Array.Length(arrJobs) & " 条职位信息")
    
    // 保存到 Excel
    TracePrint("保存到 Excel...")
    Dim objExcel = Excel.Create(True, "")
    
    // 写入表头
    Excel.SetCell(objExcel, "A1", "职位名称")
    Excel.SetCell(objExcel, "B1", "薪资")
    Excel.SetCell(objExcel, "C1", "公司名称")
    Excel.SetCell(objExcel, "D1", "公司标签")
    Excel.SetCell(objExcel, "E1", "工作地点")
    Excel.SetCell(objExcel, "F1", "经验要求")
    Excel.SetCell(objExcel, "G1", "学历要求")
    Excel.SetCell(objExcel, "H1", "HR信息")
    Excel.SetCell(objExcel, "I1", "职位链接")
    
    // 写入数据
    Dim iRow = 2
    For Each objJob In arrJobs
        Excel.SetCell(objExcel, "A" & iRow, objJob["title"])
        Excel.SetCell(objExcel, "B" & iRow, objJob["salary"])
        Excel.SetCell(objExcel, "C" & iRow, objJob["company"])
        Excel.SetCell(objExcel, "D" & iRow, objJob["tags"])
        Excel.SetCell(objExcel, "E" & iRow, objJob["location"])
        Excel.SetCell(objExcel, "F" & iRow, objJob["experience"])
        Excel.SetCell(objExcel, "G" & iRow, objJob["education"])
        Excel.SetCell(objExcel, "H" & iRow, objJob["hr"])
        Excel.SetCell(objExcel, "I" & iRow, objJob["link"])
        iRow = iRow + 1
    Next
    
    // 保存文件
    Dim sFileName = "BOSS直聘_RPA职位_" & Time.Format(Time.Now(), "yyyyMMdd_HHmmss") & ".xlsx"
    Excel.SaveAs(objExcel, "C:\" & sFileName)
    Excel.Close(objExcel)
    TracePrint("数据已保存到: C:\" & sFileName)
Else
    TracePrint("未获取到数据")
End If

WebBrowser.Close(objBrowser)
TracePrint("抓取完成！")
```

---

## Excel 数据处理

### 示例4：读取 Excel 并处理数据

```vb
' 读取 Excel 并处理数据
Dim objExcel
Dim sFilePath = "C:\data.xlsx"

' 打开 Excel 文件
objExcel = Excel.Open(sFilePath, True, "")

' 读取数据
Dim iRow = 2  ' 从第2行开始（跳过标题）
Do While True
    Dim sName = Excel.GetCell(objExcel, "A" & iRow)
    If sName = "" Then Exit Do
    
    Dim iAge = Excel.GetCell(objExcel, "B" & iRow)
    Dim sCity = Excel.GetCell(objExcel, "C" & iRow)
    
    ' 处理数据
    TracePrint("姓名: " & sName & ", 年龄: " & iAge & ", 城市: " & sCity)
    
    ' 写入处理结果
    Excel.SetCell(objExcel, "D" & iRow, "已处理")
    
    iRow = iRow + 1
Loop

' 保存并关闭
Excel.Save(objExcel)
Excel.Close(objExcel)
TracePrint("处理完成")
```

### 示例5：Excel 数据对比

```vb
' Excel 数据对比示例
Dim objExcel1, objExcel2
Dim arrDiff = []

' 打开两个 Excel 文件
objExcel1 = Excel.Open("C:\file1.xlsx", True, "")
objExcel2 = Excel.Open("C:\file2.xlsx", True, "")

' 对比数据
Dim iRow = 2
Do While True
    Dim sValue1 = Excel.GetCell(objExcel1, "A" & iRow)
    Dim sValue2 = Excel.GetCell(objExcel2, "A" & iRow)
    
    If sValue1 = "" And sValue2 = "" Then Exit Do
    
    If sValue1 <> sValue2 Then
        Dim sDiff = "行 " & iRow & ": " & sValue1 & " <> " & sValue2
        Array.Push(arrDiff, sDiff)
        TracePrint(sDiff)
    End If
    
    iRow = iRow + 1
Loop

Excel.Close(objExcel1)
Excel.Close(objExcel2)

' 保存差异报告
If Array.Length(arrDiff) > 0 Then
    Dim sReport = Array.Join(arrDiff, "\n")
    File.Write("C:\diff_report.txt", sReport, "utf-8")
    TracePrint("发现 " & Array.Length(arrDiff) & " 处差异")
Else
    TracePrint("两个文件完全相同")
End If
```

---

## 文件批量处理

### 示例6：批量重命名文件

```vb
' 批量重命名文件示例
Dim sFolderPath = "C:\Files"
Dim arrFiles = File.GetFileList(sFolderPath, "*.txt", False)

Dim iCount = 1
For Each sFile In arrFiles
    Dim sOldName = File.GetName(sFile)
    Dim sNewName = "文档_" & iCount & ".txt"
    Dim sNewPath = sFolderPath & "\" & sNewName
    
    File.Rename(sFile, sNewName)
    TracePrint("重命名: " & sOldName & " -> " & sNewName)
    
    iCount = iCount + 1
Next

TracePrint("批量重命名完成，共 " & Array.Length(arrFiles) & " 个文件")
```

### 示例7：文件内容批量替换

```vb
' 文件内容批量替换示例
Dim sFolderPath = "C:\TextFiles"
Dim arrFiles = File.GetFileList(sFolderPath, "*.txt", False)

Dim sOldText = "旧内容"
Dim sNewText = "新内容"

For Each sFile In arrFiles
    ' 读取文件
    Dim sContent = File.Read(sFile, "utf-8")
    
    ' 替换内容
    If InStr(sContent, sOldText) > 0 Then
        sContent = String.Replace(sContent, sOldText, sNewText)
        
        ' 写回文件
        File.Write(sFile, sContent, "utf-8")
        TracePrint("已处理: " & File.GetName(sFile))
    End If
Next

TracePrint("批量替换完成")
```

---

## OCR 识别应用

### 示例8：图片文字识别

```vb
' 图片文字识别示例
Dim sImagePath = "C:\image.png"
Dim objResult

' 使用本地 OCR 识别
OCR.ImageOCR(sImagePath, objResult)

' 获取全部文本
Dim sAllText = OCR.GetAllText(objResult)
TracePrint("识别结果:\n" & sAllText)

' 获取每行文本
Dim arrLines = OCR.GetLineText(objResult)
For Each sLine In arrLines
    TracePrint("行: " & sLine)
Next

' 保存结果
File.Write("C:\ocr_result.txt", sAllText, "utf-8")
```

### 示例9：屏幕区域文字识别

```vb
' 屏幕区域文字识别示例
Dim objWindow = @ui"窗口_记事本"
Dim objResult

' 识别窗口内的文字
OCR.ScreenOCR(objWindow, objResult)

' 查找特定文本
Dim sTargetText = "总金额"
Dim objPos = OCR.FindText(objResult, sTargetText)

If objPos <> Null Then
    TracePrint("找到文本: " & sTargetText)
    TracePrint("位置: X=" & objPos["x"] & ", Y=" & objPos["y"])
    
    ' 点击该文本
    OCR.ClickText(objWindow, sTargetText)
Else
    TracePrint("未找到文本: " & sTargetText)
End If
```

---

## 邮件自动化

### 示例10：自动发送邮件报告

```vb
' 自动发送邮件报告示例
Dim sSmtpServer = "smtp.qq.com"
Dim iPort = 465
Dim sUsername = "your_email@qq.com"
Dim sPassword = "your_password"
Dim sFrom = "your_email@qq.com"
Dim sTo = "receiver@example.com"
Dim sSubject = "日报 - " & Time.Format(Time.Now(), "yyyy-MM-dd")

' 生成报告内容
Dim sBody = "今日工作总结:\n\n"
sBody = sBody & "1. 完成任务数: 10\n"
sBody = sBody & "2. 处理数据量: 1000 条\n"
sBody = sBody & "3. 异常情况: 无\n\n"
sBody = sBody & "报告生成时间: " & Time.Format(Time.Now(), "yyyy-MM-dd HH:mm:ss")

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

### 示例11：批量接收邮件并处理附件

```vb
' 批量接收邮件并处理附件
Dim objMail
Dim sServer = "imap.qq.com"
Dim iPort = 993
Dim sUsername = "your_email@qq.com"
Dim sPassword = "your_password"

' 连接邮箱
objMail = Mail.Connect(sServer, iPort, sUsername, sPassword, "imap", True)

' 获取邮件列表
Dim arrMails = Mail.GetMailList(objMail, "INBOX", 10)

For Each objMailItem In arrMails
    Dim sSubject = objMailItem["subject"]
    Dim sFrom = objMailItem["from"]
    
    TracePrint("邮件: " & sSubject & " (来自: " & sFrom & ")")
    
    ' 下载附件
    If objMailItem["hasAttachment"] Then
        Mail.DownloadAttachment(objMail, objMailItem, "C:\Attachments")
        TracePrint("已下载附件")
    End If
Next

' 断开连接
Mail.Disconnect(objMail)
TracePrint("邮件处理完成")
```

---

## 数据库操作

### 示例12：数据库查询与导出

```vb
' 数据库查询与导出示例
Dim objDB
Dim sConnStr = "Driver={SQL Server};Server=localhost;Database=TestDB;Uid=sa;Pwd=password;"

' 创建数据库连接
objDB = DB.Create("sqlserver", sConnStr)

' 执行查询
Dim sSQL = "SELECT * FROM Users WHERE Age > 18"
Dim objResult = DB.QueryAll(objDB, sSQL, [])

' 导出到 Excel
Dim objExcel = Excel.Create(True, "")
Dim iRow = 1

' 写入标题
Excel.SetCell(objExcel, "A1", "ID")
Excel.SetCell(objExcel, "B1", "姓名")
Excel.SetCell(objExcel, "C1", "年龄")

' 写入数据
For Each objRow In objResult
    iRow = iRow + 1
    Excel.SetCell(objExcel, "A" & iRow, objRow["ID"])
    Excel.SetCell(objExcel, "B" & iRow, objRow["Name"])
    Excel.SetCell(objExcel, "C" & iRow, objRow["Age"])
Next

' 保存文件
Excel.SaveAs(objExcel, "C:\users_export.xlsx")
Excel.Close(objExcel)

' 关闭数据库连接
DB.Close(objDB)
TracePrint("数据导出完成，共 " & Array.Length(objResult) & " 条记录")
```

---

## 综合应用

### 示例13：自动化工作流程

```vb
' 综合自动化工作流程示例
' 场景：从网站下载数据 -> 处理数据 -> 生成报告 -> 发送邮件

TracePrint("=== 开始自动化流程 ===")

' 步骤1：从网站下载数据
TracePrint("步骤1: 下载数据...")
Dim objBrowser = WebBrowser.Create("chrome", "https://data.example.com", 30)
Delay(2000)

' 登录
Keyboard.InputText(@ui"用户名", "admin", True, False)
Keyboard.InputPwd(@ui"密码", "password", True)
Mouse.Click(@ui"登录按钮")
Delay(3000)

' 下载数据文件
Mouse.Click(@ui"导出按钮")
Delay(5000)
WebBrowser.Close(objBrowser)
TracePrint("数据下载完成")

' 步骤2：处理数据
TracePrint("步骤2: 处理数据...")
Dim objExcel = Excel.Open("C:\Downloads\data.xlsx", True, "")
Dim iRow = 2
Dim iProcessed = 0

Do While True
    Dim sValue = Excel.GetCell(objExcel, "A" & iRow)
    If sValue = "" Then Exit Do
    
    ' 数据处理逻辑
    Dim sProcessed = String.Replace(sValue, "旧值", "新值")
    Excel.SetCell(objExcel, "B" & iRow, sProcessed)
    
    iProcessed = iProcessed + 1
    iRow = iRow + 1
Loop

Excel.SaveAs(objExcel, "C:\Reports\processed_data.xlsx")
Excel.Close(objExcel)
TracePrint("数据处理完成，共处理 " & iProcessed & " 条")

' 步骤3：生成报告
TracePrint("步骤3: 生成报告...")
Dim sReport = "数据处理报告\n"
sReport = sReport & "==================\n\n"
sReport = sReport & "处理时间: " & Time.Format(Time.Now(), "yyyy-MM-dd HH:mm:ss") & "\n"
sReport = sReport & "处理记录数: " & iProcessed & "\n"
sReport = sReport & "数据文件: C:\Reports\processed_data.xlsx\n"

File.Write("C:\Reports\report.txt", sReport, "utf-8")
TracePrint("报告生成完成")

' 步骤4：发送邮件
TracePrint("步骤4: 发送邮件...")
Try
    Mail.Send("smtp.qq.com", 465, "your_email@qq.com", "password", _
              "your_email@qq.com", "manager@example.com", _
              "数据处理报告 - " & Time.Format(Time.Now(), "yyyy-MM-dd"), _
              sReport, ["C:\Reports\processed_data.xlsx"])
    TracePrint("邮件发送成功")
Catch ex
    TracePrint("邮件发送失败: " & ex.Message)
End Try

TracePrint("=== 自动化流程完成 ===")
```

### 示例14：异常处理与日志记录

```vb
' 带完善异常处理的自动化流程
Dim sLogFile = "C:\Logs\automation_" & Time.Format(Time.Now(), "yyyyMMdd") & ".log"

Function WriteLog(sMessage)
    Dim sTime = Time.Format(Time.Now(), "yyyy-MM-dd HH:mm:ss")
    Dim sLogLine = "[" & sTime & "] " & sMessage & "\n"
    File.Append(sLogFile, sLogLine, "utf-8")
    TracePrint(sMessage)
End Function

WriteLog("=== 流程开始 ===")

Try
    ' 主要业务逻辑
    WriteLog("步骤1: 打开浏览器")
    Dim objBrowser = WebBrowser.Create("chrome", "https://www.example.com", 30)
    
    WriteLog("步骤2: 执行操作")
    ' ... 业务代码 ...
    
    WriteLog("步骤3: 关闭浏览器")
    WebBrowser.Close(objBrowser)
    
    WriteLog("=== 流程成功完成 ===")
    
Catch ex
    WriteLog("!!! 发生错误: " & ex.Message)
    WriteLog("错误位置: " & ex.Source)
    
    ' 发送错误通知邮件
    Try
        Mail.Send("smtp.qq.com", 465, "your_email@qq.com", "password", _
                  "your_email@qq.com", "admin@example.com", _
                  "自动化流程错误通知", _
                  "流程执行失败\n\n错误信息: " & ex.Message, [])
    Catch mailEx
        WriteLog("邮件发送失败: " & mailEx.Message)
    End Try
    
Finally
    WriteLog("=== 流程结束 ===")
End Try
```

---

## 使用技巧

### 1. 元素定位优化
```vb
' 使用多种定位方式提高稳定性
If UiElement.Exists(@ui"按钮_ID", 5) Then
    Mouse.Click(@ui"按钮_ID")
ElseIf UiElement.Exists(@ui"按钮_Name", 5) Then
    Mouse.Click(@ui"按钮_Name")
Else
    TracePrint("元素未找到")
End If
```

### 2. 智能等待
```vb
' 智能等待元素出现
Dim iMaxWait = 30
Dim iWaited = 0
Do While Not UiElement.Exists(@ui"目标元素", 1)
    iWaited = iWaited + 1
    If iWaited >= iMaxWait Then
        TracePrint("等待超时")
        Exit Do
    End If
    Delay(1000)
Loop
```

### 3. 批量操作优化
```vb
' 使用数组批量处理
Dim arrTasks = ["任务1", "任务2", "任务3"]
For Each sTask In arrTasks
    Try
        ' 处理任务
        TracePrint("处理: " & sTask)
    Catch ex
        TracePrint("任务失败: " & sTask & " - " & ex.Message)
        Continue  ' 继续处理下一个
    End Try
Next
```

---

**文档版本**: v1.0.0  
**适用版本**: UIBot 6.0.0.211215(64位)  
**更新时间**: 2024-01-15
