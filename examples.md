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
// 自动登录示例
Dim objBrowser

// 启动浏览器并打开登录页面
objBrowser = WebBrowser.Create("chrome", "https://www.example.com/login", 30)
Delay(2000)

// 输入用户名
Keyboard.InputText(@ui"用户名输入框", "your_username", True, False)
Delay(500)

// 输入密码
Keyboard.InputPwd(@ui"密码输入框", "your_password", True)
Delay(500)

// 点击登录按钮
Mouse.Click(@ui"登录按钮", "left", "single", 0, 0)
Delay(3000)

// 验证登录成功
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
// 批量下载文件示例
Dim objBrowser
Dim arrUrls = ["https://example.com/file1.pdf", _
               "https://example.com/file2.pdf", _
               "https://example.com/file3.pdf"]

objBrowser = WebBrowser.Create("chrome", "about:blank", 30)

For Each url In arrUrls
    // 下载文件
    WebBrowser.Download(objBrowser, url, "C:\Downloads", 60)
    TracePrint("已下载: " & url)
    Delay(2000)
Next

WebBrowser.Close(objBrowser)
TracePrint("所有文件下载完成")
```

### 示例3：网页数据采集

```vb
// 网页数据采集示例
Dim objBrowser
Dim arrData = []

objBrowser = WebBrowser.Create("chrome", "https://www.example.com/products", 30)
Delay(3000)

// 获取所有商品名称
Dim objElements = UiElement.GetChildren(@ui"商品列表", 1)
For Each objElement In objElements
    Dim sName = UiElement.GetText(objElement)
    Array.Push(arrData, sName)
    TracePrint("商品: " & sName)
Next

// 保存到文件
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
// 读取 Excel 并处理数据
Dim objExcel
Dim sFilePath = "C:\data.xlsx"

// 打开 Excel 文件
objExcel = Excel.Open(sFilePath, True, "")

// 读取数据
Dim iRow = 2  // 从第2行开始（跳过标题）
Do While True
    Dim sName = Excel.GetCell(objExcel, "A" & iRow)
    If sName = "" Then Exit Do
    
    Dim iAge = Excel.GetCell(objExcel, "B" & iRow)
    Dim sCity = Excel.GetCell(objExcel, "C" & iRow)
    
    // 处理数据
    TracePrint("姓名: " & sName & ", 年龄: " & iAge & ", 城市: " & sCity)
    
    // 写入处理结果
    Excel.SetCell(objExcel, "D" & iRow, "已处理")
    
    iRow = iRow + 1
Loop

// 保存并关闭
Excel.Save(objExcel)
Excel.Close(objExcel)
TracePrint("处理完成")
```

### 示例5：Excel 数据对比

```vb
// Excel 数据对比示例
Dim objExcel1, objExcel2
Dim arrDiff = []

// 打开两个 Excel 文件
objExcel1 = Excel.Open("C:\file1.xlsx", True, "")
objExcel2 = Excel.Open("C:\file2.xlsx", True, "")

// 对比数据
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

// 保存差异报告
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
// 批量重命名文件示例
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
// 文件内容批量替换示例
Dim sFolderPath = "C:\TextFiles"
Dim arrFiles = File.GetFileList(sFolderPath, "*.txt", False)

Dim sOldText = "旧内容"
Dim sNewText = "新内容"

For Each sFile In arrFiles
    // 读取文件
    Dim sContent = File.Read(sFile, "utf-8")
    
    // 替换内容
    If InStr(sContent, sOldText) > 0 Then
        sContent = String.Replace(sContent, sOldText, sNewText)
        
        // 写回文件
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
// 图片文字识别示例
Dim sImagePath = "C:\image.png"
Dim objResult

// 使用本地 OCR 识别
OCR.ImageOCR(sImagePath, objResult)

// 获取全部文本
Dim sAllText = OCR.GetAllText(objResult)
TracePrint("识别结果:\n" & sAllText)

// 获取每行文本
Dim arrLines = OCR.GetLineText(objResult)
For Each sLine In arrLines
    TracePrint("行: " & sLine)
Next

// 保存结果
File.Write("C:\ocr_result.txt", sAllText, "utf-8")
```

### 示例9：屏幕区域文字识别

```vb
// 屏幕区域文字识别示例
Dim objWindow = @ui"窗口_记事本"
Dim objResult

// 识别窗口内的文字
OCR.ScreenOCR(objWindow, objResult)

// 查找特定文本
Dim sTargetText = "总金额"
Dim objPos = OCR.FindText(objResult, sTargetText)

If objPos <> Null Then
    TracePrint("找到文本: " & sTargetText)
    TracePrint("位置: X=" & objPos["x"] & ", Y=" & objPos["y"])
    
    // 点击该文本
    OCR.ClickText(objWindow, sTargetText)
Else
    TracePrint("未找到文本: " & sTargetText)
End If
```

### 示例10：通用多票据识别（Mage AI）

**业务场景**：
通用多票据识别能够识别全电发票、普通发票、专用发票、出租车票、火车票、飞机行程单、财政发票等财务场景覆盖的30多种常见票据，并从中抽取出核心字段值。

财务场景是RPA+AI最广泛应用的场景。票据识别能将财务报销填单、审单流程中的重复劳动自动化，**提升工作效率**；同时由于财务场景容错性低、业务规则多，机器人能够保证稳定的工作状态，**减少人工原因导致的错误**。

**核心特点**：
- ✅ **种类丰富**：支持包括国税票种、地方票种及其他票种等30余种票据
- ✅ **自动分类**：模型能够自动识别是哪种票据
- ✅ **票据混贴**：支持一张图片上存在多个不同类型的票据，模型会把每张票据切分出来，分别返回对应的识别结果

**支持的票据类型（32种）**：

| 序号 | 票据类型 | type_key | 说明 |
|------|---------|----------|------|
| 1 | 增值税专用发票 | vat_special_invoice | 国税发票 |
| 2 | 增值税普通发票 | vat_common_invoice | 国税发票 |
| 3 | 增值税电子普通发票 | vat_electronic_invoice | 国税发票 |
| 4 | 增值税电子普通发票（通行费） | vat_electronic_toll_invoice | 国税发票 |
| 5 | 增值税电子专用发票 | vat_electronic_special_invoice | 国税发票 |
| 6 | 区块链电子发票 | blockchain_electronic_invoice | 地方票种 |
| 7 | 电子发票（增值税专用发票） | vat_electronic_special_invoice_new | 全电发票 |
| 8 | 电子发票（普通发票） | vat_electronic_invoice_new | 全电发票 |
| 9 | 出租车发票 | taxi_ticket | 交通票据 |
| 10 | 机票行程单 | air_transport | 交通票据 |
| 11 | 电子发票（航空运输电子客票行程单） | electronic_air_transport | 交通票据 |
| 12 | 火车票 | train_ticket | 交通票据 |
| 13 | 电子发票（铁路电子客票） | electronic_train_ticket | 交通票据 |
| 14 | 增值税普通发票（卷票） | vat_roll_invoice | 国税发票 |
| 15 | 机动车销售统一发票 | motor_vehicle_sale_invoice | 专用发票 |
| 16 | 二手车销售统一发票 | used_car_purchase_invoice | 专用发票 |
| 17 | 通用定额发票 | quota_invoice | 地方票种 |
| 18 | 通用机打发票 | general_machine_invoice | 地方票种 |
| 19 | 通用机打电子发票 | machine_printed_invoice | 地方票种 |
| 20 | 公路客运发票 | highway_passenger_invoice | 交通票据 |
| 21 | 增值税销货清单 | vat_invoice_sales_list | 国税发票 |
| 22 | 船运客票 | shipping_invoice | 交通票据 |
| 23 | 过路过桥费发票 | vehicle_toll | 交通票据 |
| 24 | 网约车行程单 | travel_transport | 交通票据 |
| 25 | 火车票退票费 | ticket_refund_fee | 交通票据 |
| 26 | 电子财政票据 | fiscal_paper_electronic | 财政票据 |
| 27 | 财政票据 | fiscal_paper | 财政票据 |
| 28 | 电子医疗票据 | medical_electronic_invoice | 医疗票据 |
| 29 | 医疗票据 | medical_invoice | 医疗票据 |
| 30 | 完税证明 | tax_clearance_certificate | 税务票据 |
| 31 | 海关缴费书 | customs_payment_form | 海关票据 |
| 32 | 海关报关单 | custom_declaration_form | 海关票据 |

**识别方式**：
- 方式1：屏幕识别 - 识别屏幕上显示的票据
- 方式2：图片识别 - 识别本地图片文件中的票据
- 方式3：PDF识别 - 识别PDF文件中的票据（支持多页）

#### 方式1：屏幕票据识别

```vb
// 屏幕票据识别示例 - 识别屏幕上显示的票据
Dim ctrl = @ui"票据显示窗口"  // 票据显示的窗口或区域
Dim range = Null  // 识别范围，Null 表示整个窗口
Dim config = {
    "app_id": "your_app_id",
    "app_key": "your_app_key",
    "timeout": 30
}
Dim timeout = 30

TracePrint("=== 开始屏幕票据识别 ===")

Try
    // 使用 Mage.ScreenOCRInvoice 识别屏幕上的票据
    With Each Mage.ScreenOCRInvoice(ctrl, range, config, timeout)
        // 获取票据类型
        Dim sInvoiceType = .ExtractInvoiceType()
        TracePrint("识别到票据类型: " & sInvoiceType)
        
        // 根据不同票据类型提取字段
        Select Case .ExtractInvoiceType()
            Case Alias("train_ticket", "火车票")
                TracePrint("--- 火车票信息 ---")
                Dim sDepartureStation = .ExtractInvoiceInfo("train_ticket", "departure_station")
                Dim sDepartureDate = .ExtractInvoiceInfo("train_ticket", "departure_date")
                Dim sArrivalStation = .ExtractInvoiceInfo("train_ticket", "arrival_station")
                Dim sPrice = .ExtractInvoiceInfo("train_ticket", "price")
                Dim sTrainNumber = .ExtractInvoiceInfo("train_ticket", "train_number")
                Dim sSeatNumber = .ExtractInvoiceInfo("train_ticket", "seat_number")
                
                TracePrint("出发站: " & sDepartureStation)
                TracePrint("出发日期: " & sDepartureDate)
                TracePrint("到达站: " & sArrivalStation)
                TracePrint("车次: " & sTrainNumber)
                TracePrint("座位号: " & sSeatNumber)
                TracePrint("票价: " & sPrice & " 元")
                
            Case Alias("air_transport", "机票行程单")
                TracePrint("--- 机票行程单信息 ---")
                Dim sPassengerName = .ExtractInvoiceInfo("air_transport", "passenger_name")
                Dim sFlightNumber = .ExtractInvoiceInfo("air_transport", "flight_number")
                Dim sFrom = .ExtractInvoiceInfo("air_transport", "from")
                Dim sTo = .ExtractInvoiceInfo("air_transport", "to")
                Dim sDate = .ExtractInvoiceInfo("air_transport", "date")
                Dim sFare = .ExtractInvoiceInfo("air_transport", "fare")
                
                TracePrint("乘客姓名: " & sPassengerName)
                TracePrint("航班号: " & sFlightNumber)
                TracePrint("出发地: " & sFrom)
                TracePrint("目的地: " & sTo)
                TracePrint("日期: " & sDate)
                TracePrint("票价: " & sFare & " 元")
                
            Case Alias("vat_special_invoice", "增值税专用发票")
                TracePrint("--- 增值税专用发票信息 ---")
                Dim sInvoiceCode = .ExtractInvoiceInfo("vat_special_invoice", "vat_invoice_daima_print")
                Dim sInvoiceNumber = .ExtractInvoiceInfo("vat_special_invoice", "vat_invoice_haoma_large_size")
                Dim sIssueDate = .ExtractInvoiceInfo("vat_special_invoice", "vat_invoice_issue_date")
                Dim sTotalAmount = .ExtractInvoiceInfo("vat_special_invoice", "vat_invoice_total_cover_tax_digits")
                Dim sSellerName = .ExtractInvoiceInfo("vat_special_invoice", "vat_invoice_seller_name")
                Dim sBuyerName = .ExtractInvoiceInfo("vat_special_invoice", "vat_invoice_payer_name")
                
                TracePrint("发票代码: " & sInvoiceCode)
                TracePrint("发票号码: " & sInvoiceNumber)
                TracePrint("开票日期: " & sIssueDate)
                TracePrint("价税合计: " & sTotalAmount & " 元")
                TracePrint("销售方: " & sSellerName)
                TracePrint("购买方: " & sBuyerName)
                
            Case Alias("taxi_ticket", "出租车发票")
                TracePrint("--- 出租车发票信息 ---")
                Dim sInvoiceNo = .ExtractInvoiceInfo("taxi_ticket", "invoice_no")
                Dim sDate = .ExtractInvoiceInfo("taxi_ticket", "date")
                Dim sTaxiNo = .ExtractInvoiceInfo("taxi_ticket", "taxi_no")
                Dim sSum = .ExtractInvoiceInfo("taxi_ticket", "sum")
                
                TracePrint("发票号码: " & sInvoiceNo)
                TracePrint("日期: " & sDate)
                TracePrint("车号: " & sTaxiNo)
                TracePrint("金额: " & sSum & " 元")
                
            Case Else
                TracePrint("识别到其他类型票据: " & sInvoiceType)
                // 可以根据需要添加更多票据类型的处理
        End Select
    End With
    
    TracePrint("=== 屏幕票据识别完成 ===")
    
Catch ex
    TracePrint("识别失败: " & ex.Message)
End Try
```

#### 方式2：图片票据识别

```vb
// 图片票据识别示例 - 识别本地图片文件中的票据
Dim imagePath = "C:\Invoices\train_ticket.jpg"  // 票据图片路径
Dim config = {
    "app_id": "your_app_id",
    "app_key": "your_app_key",
    "timeout": 30
}
Dim timeout = 30

TracePrint("=== 开始图片票据识别 ===")
TracePrint("图片路径: " & imagePath)

Try
    // 使用 Mage.ImageOCRInvoice 识别图片中的票据
    With Each Mage.ImageOCRInvoice(imagePath, config, timeout)
        // 获取票据类型
        Dim sInvoiceType = .ExtractInvoiceType()
        TracePrint("识别到票据类型: " & sInvoiceType)
        
        // 提取火车票信息
        If sInvoiceType = "train_ticket" Or sInvoiceType = "火车票" Then
            TracePrint("--- 火车票详细信息 ---")
            
            // 基本信息
            Dim sTicketNumber = .ExtractInvoiceInfo("train_ticket", "ticket_number")
            Dim sDepartureStation = .ExtractInvoiceInfo("train_ticket", "departure_station")
            Dim sArrivalStation = .ExtractInvoiceInfo("train_ticket", "arrival_station")
            Dim sDepartureDate = .ExtractInvoiceInfo("train_ticket", "departure_date")
            Dim sTrainNumber = .ExtractInvoiceInfo("train_ticket", "train_number")
            
            // 座位和价格信息
            Dim sSeatClass = .ExtractInvoiceInfo("train_ticket", "class")
            Dim sSeatNumber = .ExtractInvoiceInfo("train_ticket", "seat_number")
            Dim sPrice = .ExtractInvoiceInfo("train_ticket", "price")
            
            // 乘客信息
            Dim sPassengerName = .ExtractInvoiceInfo("train_ticket", "passenger_name")
            Dim sPassengerId = .ExtractInvoiceInfo("train_ticket", "passenger_id")
            
            // 输出识别结果
            TracePrint("票据号码: " & sTicketNumber)
            TracePrint("乘客姓名: " & sPassengerName)
            TracePrint("身份证号: " & sPassengerId)
            TracePrint("出发站: " & sDepartureStation)
            TracePrint("到达站: " & sArrivalStation)
            TracePrint("车次: " & sTrainNumber)
            TracePrint("日期: " & sDepartureDate)
            TracePrint("座位类型: " & sSeatClass)
            TracePrint("座位号: " & sSeatNumber)
            TracePrint("票价: " & sPrice & " 元")
            
            // 保存识别结果到 Excel
            Dim objExcel = Excel.Create(True, "")
            Excel.SetCell(objExcel, "A1", "字段名称")
            Excel.SetCell(objExcel, "B1", "识别结果")
            Excel.SetCell(objExcel, "A2", "票据号码")
            Excel.SetCell(objExcel, "B2", sTicketNumber)
            Excel.SetCell(objExcel, "A3", "乘客姓名")
            Excel.SetCell(objExcel, "B3", sPassengerName)
            Excel.SetCell(objExcel, "A4", "出发站")
            Excel.SetCell(objExcel, "B4", sDepartureStation)
            Excel.SetCell(objExcel, "A5", "到达站")
            Excel.SetCell(objExcel, "B5", sArrivalStation)
            Excel.SetCell(objExcel, "A6", "车次")
            Excel.SetCell(objExcel, "B6", sTrainNumber)
            Excel.SetCell(objExcel, "A7", "日期")
            Excel.SetCell(objExcel, "B7", sDepartureDate)
            Excel.SetCell(objExcel, "A8", "票价")
            Excel.SetCell(objExcel, "B8", sPrice)
            
            Dim sResultFile = "C:\Invoices\train_ticket_result_" & Time.Format(Time.Now(), "yyyyMMdd_HHmmss") & ".xlsx"
            Excel.SaveAs(objExcel, sResultFile)
            Excel.Close(objExcel)
            TracePrint("识别结果已保存到: " & sResultFile)
        Else
            TracePrint("当前示例仅处理火车票，识别到的类型为: " & sInvoiceType)
        End If
    End With
    
    TracePrint("=== 图片票据识别完成 ===")
    
Catch ex
    TracePrint("识别失败: " & ex.Message)
End Try
```

#### 方式3：PDF票据识别

```vb
// PDF票据识别示例 - 识别PDF文件中的票据（支持多页）
Dim config = {
    "app_id": "your_app_id",
    "app_key": "your_app_key",
    "timeout": 30
}
Dim pdfPath = "C:\Invoices\invoices.pdf"  // PDF文件路径
Dim password = ""  // PDF密码，无密码则为空字符串
Dim readAll = True  // True: 识别所有页，False: 识别指定页
Dim pages = "1-3"  // 指定页码范围，如 "1-3" 或 "1,3,5"
Dim interval = 1000  // 每页识别间隔（毫秒）
Dim timeout = 60

TracePrint("=== 开始PDF票据识别 ===")
TracePrint("PDF路径: " & pdfPath)

Try
    Dim iPageCount = 0
    Dim arrResults = []  // 存储所有识别结果
    
    // 使用 Mage.PDFOCRInvoice 识别PDF中的票据
    With Each Mage.PDFOCRInvoice(config, pdfPath, password, readAll, pages, interval, timeout)
        iPageCount = iPageCount + 1
        TracePrint("--- 第 " & iPageCount & " 页 ---")
        
        // 获取票据类型
        Dim sInvoiceType = .ExtractInvoiceType()
        TracePrint("票据类型: " & sInvoiceType)
        
        // 创建结果对象
        Dim objResult = {
            "page": iPageCount,
            "type": sInvoiceType,
            "data": {}
        }
        
        // 根据票据类型提取信息
        Select Case .ExtractInvoiceType()
            Case Alias("train_ticket", "火车票")
                objResult["data"]["departure_station"] = .ExtractInvoiceInfo("train_ticket", "departure_station")
                objResult["data"]["arrival_station"] = .ExtractInvoiceInfo("train_ticket", "arrival_station")
                objResult["data"]["departure_date"] = .ExtractInvoiceInfo("train_ticket", "departure_date")
                objResult["data"]["train_number"] = .ExtractInvoiceInfo("train_ticket", "train_number")
                objResult["data"]["price"] = .ExtractInvoiceInfo("train_ticket", "price")
                objResult["data"]["passenger_name"] = .ExtractInvoiceInfo("train_ticket", "passenger_name")
                
                TracePrint("出发站: " & objResult["data"]["departure_station"])
                TracePrint("到达站: " & objResult["data"]["arrival_station"])
                TracePrint("日期: " & objResult["data"]["departure_date"])
                TracePrint("车次: " & objResult["data"]["train_number"])
                TracePrint("票价: " & objResult["data"]["price"])
                
            Case Alias("vat_special_invoice", "增值税专用发票"), Alias("vat_common_invoice", "增值税普通发票")
                objResult["data"]["invoice_code"] = .ExtractInvoiceInfo(sInvoiceType, "vat_invoice_daima_print")
                objResult["data"]["invoice_number"] = .ExtractInvoiceInfo(sInvoiceType, "vat_invoice_haoma_large_size")
                objResult["data"]["issue_date"] = .ExtractInvoiceInfo(sInvoiceType, "vat_invoice_issue_date")
                objResult["data"]["total_amount"] = .ExtractInvoiceInfo(sInvoiceType, "vat_invoice_total_cover_tax_digits")
                objResult["data"]["seller_name"] = .ExtractInvoiceInfo(sInvoiceType, "vat_invoice_seller_name")
                
                TracePrint("发票代码: " & objResult["data"]["invoice_code"])
                TracePrint("发票号码: " & objResult["data"]["invoice_number"])
                TracePrint("开票日期: " & objResult["data"]["issue_date"])
                TracePrint("价税合计: " & objResult["data"]["total_amount"])
                
            Case Alias("taxi_ticket", "出租车发票")
                objResult["data"]["invoice_no"] = .ExtractInvoiceInfo("taxi_ticket", "invoice_no")
                objResult["data"]["date"] = .ExtractInvoiceInfo("taxi_ticket", "date")
                objResult["data"]["sum"] = .ExtractInvoiceInfo("taxi_ticket", "sum")
                
                TracePrint("发票号码: " & objResult["data"]["invoice_no"])
                TracePrint("日期: " & objResult["data"]["date"])
                TracePrint("金额: " & objResult["data"]["sum"])
                
            Case Else
                TracePrint("其他类型票据，请根据需要添加处理逻辑")
        End Select
        
        // 添加到结果数组
        Array.Push(arrResults, objResult)
    End With
    
    TracePrint("=== PDF票据识别完成 ===")
    TracePrint("共识别 " & iPageCount & " 页票据")
    
    // 将所有结果导出到 Excel
    If Array.Length(arrResults) > 0 Then
        Dim objExcel = Excel.Create(True, "")
        
        // 写入表头
        Excel.SetCell(objExcel, "A1", "页码")
        Excel.SetCell(objExcel, "B1", "票据类型")
        Excel.SetCell(objExcel, "C1", "关键信息")
        
        // 写入数据
        Dim iRow = 2
        For Each objResult In arrResults
            Excel.SetCell(objExcel, "A" & iRow, objResult["page"])
            Excel.SetCell(objExcel, "B" & iRow, objResult["type"])
            
            // 将数据对象转换为字符串
            Dim sInfo = ""
            For Each sKey In objResult["data"]
                sInfo = sInfo & sKey & ": " & objResult["data"][sKey] & "; "
            Next
            Excel.SetCell(objExcel, "C" & iRow, sInfo)
            
            iRow = iRow + 1
        Next
        
        // 保存文件
        Dim sResultFile = "C:\Invoices\pdf_invoices_result_" & Time.Format(Time.Now(), "yyyyMMdd_HHmmss") & ".xlsx"
        Excel.SaveAs(objExcel, sResultFile)
        Excel.Close(objExcel)
        TracePrint("所有识别结果已保存到: " & sResultFile)
    End If
    
Catch ex
    TracePrint("识别失败: " & ex.Message)
End Try
```

#### 批量票据识别与分类

```vb
' 批量票据识别与分类示例
Dim sFolderPath = "C:\Invoices\Batch"  ' 票据图片文件夹
Dim arrFiles = File.GetFileList(sFolderPath, "*.jpg|*.png|*.jpeg", False)
Dim config = {
    "app_id": "your_app_id",
    "app_key": "your_app_key",
    "timeout": 30
}

TracePrint("=== 开始批量票据识别 ===")
TracePrint("待识别文件数: " & Array.Length(arrFiles))

' 创建分类统计
Dim objStats = {
    "train_ticket": 0,
    "air_transport": 0,
    "taxi_ticket": 0,
    "vat_invoice": 0,
    "other": 0
}

' 创建 Excel 汇总表
Dim objExcel = Excel.Create(True, "")
Excel.SetCell(objExcel, "A1", "序号")
Excel.SetCell(objExcel, "B1", "文件名")
Excel.SetCell(objExcel, "C1", "票据类型")
Excel.SetCell(objExcel, "D1", "关键信息")
Excel.SetCell(objExcel, "E1", "金额")

Dim iRow = 2
Dim iSuccess = 0
Dim iFailed = 0

For Each sFile In arrFiles
    TracePrint("处理文件: " & File.GetName(sFile))
    
    Try
        With Each Mage.ImageOCRInvoice(sFile, config, 30)
            Dim sType = .ExtractInvoiceType()
            Dim sAmount = ""
            Dim sInfo = ""
            
            ' 根据类型提取关键信息
            Select Case sType
                Case Alias("train_ticket", "火车票")
                    objStats["train_ticket"] = objStats["train_ticket"] + 1
                    sInfo = .ExtractInvoiceInfo("train_ticket", "departure_station") & " -> " & _
                            .ExtractInvoiceInfo("train_ticket", "arrival_station")
                    sAmount = .ExtractInvoiceInfo("train_ticket", "price")
                    
                Case Alias("air_transport", "机票行程单")
                    objStats["air_transport"] = objStats["air_transport"] + 1
                    sInfo = .ExtractInvoiceInfo("air_transport", "from") & " -> " & _
                            .ExtractInvoiceInfo("air_transport", "to")
                    sAmount = .ExtractInvoiceInfo("air_transport", "fare")
                    
                Case Alias("taxi_ticket", "出租车发票")
                    objStats["taxi_ticket"] = objStats["taxi_ticket"] + 1
                    sInfo = "车号: " & .ExtractInvoiceInfo("taxi_ticket", "taxi_no")
                    sAmount = .ExtractInvoiceInfo("taxi_ticket", "sum")
                    
                Case Alias("vat_special_invoice", "增值税专用发票"), _
                     Alias("vat_common_invoice", "增值税普通发票")
                    objStats["vat_invoice"] = objStats["vat_invoice"] + 1
                    sInfo = "发票号: " & .ExtractInvoiceInfo(sType, "vat_invoice_haoma_large_size")
                    sAmount = .ExtractInvoiceInfo(sType, "vat_invoice_total_cover_tax_digits")
                    
                Case Else
                    objStats["other"] = objStats["other"] + 1
                    sInfo = "其他类型"
            End Select
            
            ' 写入 Excel
            Excel.SetCell(objExcel, "A" & iRow, iRow - 1)
            Excel.SetCell(objExcel, "B" & iRow, File.GetName(sFile))
            Excel.SetCell(objExcel, "C" & iRow, sType)
            Excel.SetCell(objExcel, "D" & iRow, sInfo)
            Excel.SetCell(objExcel, "E" & iRow, sAmount)
            
            iRow = iRow + 1
            iSuccess = iSuccess + 1
            TracePrint("识别成功: " & sType)
        End With
        
    Catch ex
        iFailed = iFailed + 1
        TracePrint("识别失败: " & ex.Message)
    End Try
    
    Delay(500)  ' 避免请求过快
Next

' 保存汇总表
Dim sResultFile = "C:\Invoices\batch_result_" & Time.Format(Time.Now(), "yyyyMMdd_HHmmss") & ".xlsx"
Excel.SaveAs(objExcel, sResultFile)
Excel.Close(objExcel)

' 输出统计信息
TracePrint("=== 批量识别完成 ===")
TracePrint("总文件数: " & Array.Length(arrFiles))
TracePrint("成功: " & iSuccess & ", 失败: " & iFailed)
TracePrint("--- 票据类型统计 ---")
TracePrint("火车票: " & objStats["train_ticket"])
TracePrint("机票: " & objStats["air_transport"])
TracePrint("出租车票: " & objStats["taxi_ticket"])
TracePrint("增值税发票: " & objStats["vat_invoice"])
TracePrint("其他: " & objStats["other"])
TracePrint("结果已保存到: " & sResultFile)
```

**使用说明**：
1. **配置要求**：需要配置 Mage AI 的 app_id 和 app_key（在来也科技平台申请）
2. **票据类型**：支持32种票据类型，详见上方表格中的 type_key
3. **字段提取**：不同票据类型支持的字段不同，常见字段包括：
   - 火车票：departure_station（出发站）、arrival_station（到达站）、departure_date（日期）、train_number（车次）、price（票价）、passenger_name（乘客姓名）、passenger_id（身份证号）、seat_number（座位号）
   - 增值税发票：vat_invoice_daima_print（发票代码）、vat_invoice_haoma_large_size（发票号码）、vat_invoice_issue_date（开票日期）、vat_invoice_total_cover_tax_digits（价税合计）、vat_invoice_seller_name（销售方名称）、vat_invoice_payer_name（购买方名称）
   - 机票行程单：passenger_name（乘客姓名）、flight_number（航班号）、from（出发地）、to（目的地）、date（日期）、fare（票价）
   - 出租车发票：invoice_no（发票号码）、date（日期）、taxi_no（车号）、sum（金额）
4. **性能优化**：
   - 建议添加适当的延迟（500-1000ms）避免请求过快
   - 对于大批量识别，建议分批处理并保存中间结果
   - PDF识别时可以指定页码范围，避免处理不必要的页面
5. **错误处理**：建议使用 Try-Catch 包裹识别代码，处理网络异常、识别失败等情况
6. **参考文档**：完整的字段列表和API说明请参考 [来也IDP官方文档](https://documents.laiye.com/idp-mage/docs/OCR/ocr_receipt)

**成功案例**：
某企业的财务报销业务特点是审核点极多（将近100个），审核点复杂，存在多票据交叉复核，人工验证出错率极高。通过RPA+AI重塑业务流程，企业降低了填单环节的人工输入成本，减少了审单环节的出错率。

---

## 邮件自动化

### 示例10：自动发送邮件报告

```vb
// 自动发送邮件报告示例
Dim sSmtpServer = "smtp.qq.com"
Dim iPort = 465
Dim sUsername = "your_email@qq.com"
Dim sPassword = "your_password"
Dim sFrom = "your_email@qq.com"
Dim sTo = "receiver@example.com"
Dim sSubject = "日报 - " & Time.Format(Time.Now(), "yyyy-MM-dd")

// 生成报告内容
Dim sBody = "今日工作总结:\n\n"
sBody = sBody & "1. 完成任务数: 10\n"
sBody = sBody & "2. 处理数据量: 1000 条\n"
sBody = sBody & "3. 异常情况: 无\n\n"
sBody = sBody & "报告生成时间: " & Time.Format(Time.Now(), "yyyy-MM-dd HH:mm:ss")

// 附件列表
Dim arrAttachments = ["C:\report.xlsx", "C:\log.txt"]

// 发送邮件
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
// 批量接收邮件并处理附件
Dim objMail
Dim sServer = "imap.qq.com"
Dim iPort = 993
Dim sUsername = "your_email@qq.com"
Dim sPassword = "your_password"

// 连接邮箱
objMail = Mail.Connect(sServer, iPort, sUsername, sPassword, "imap", True)

// 获取邮件列表
Dim arrMails = Mail.GetMailList(objMail, "INBOX", 10)

For Each objMailItem In arrMails
    Dim sSubject = objMailItem["subject"]
    Dim sFrom = objMailItem["from"]
    
    TracePrint("邮件: " & sSubject & " (来自: " & sFrom & ")")
    
    // 下载附件
    If objMailItem["hasAttachment"] Then
        Mail.DownloadAttachment(objMail, objMailItem, "C:\Attachments")
        TracePrint("已下载附件")
    End If
Next

// 断开连接
Mail.Disconnect(objMail)
TracePrint("邮件处理完成")
```

---

## 数据库操作

### 示例12：数据库查询与导出

```vb
// 数据库查询与导出示例
Dim objDB
Dim sConnStr = "Driver={SQL Server};Server=localhost;Database=TestDB;Uid=sa;Pwd=password;"

// 创建数据库连接
objDB = DB.Create("sqlserver", sConnStr)

// 执行查询
Dim sSQL = "SELECT * FROM Users WHERE Age > 18"
Dim objResult = DB.QueryAll(objDB, sSQL, [])

// 导出到 Excel
Dim objExcel = Excel.Create(True, "")
Dim iRow = 1

// 写入标题
Excel.SetCell(objExcel, "A1", "ID")
Excel.SetCell(objExcel, "B1", "姓名")
Excel.SetCell(objExcel, "C1", "年龄")

// 写入数据
For Each objRow In objResult
    iRow = iRow + 1
    Excel.SetCell(objExcel, "A" & iRow, objRow["ID"])
    Excel.SetCell(objExcel, "B" & iRow, objRow["Name"])
    Excel.SetCell(objExcel, "C" & iRow, objRow["Age"])
Next

// 保存文件
Excel.SaveAs(objExcel, "C:\users_export.xlsx")
Excel.Close(objExcel)

// 关闭数据库连接
DB.Close(objDB)
TracePrint("数据导出完成，共 " & Array.Length(objResult) & " 条记录")
```

---

## 综合应用

### 示例13：自动化工作流程

```vb
// 综合自动化工作流程示例
// 场景：从网站下载数据 -> 处理数据 -> 生成报告 -> 发送邮件

TracePrint("=== 开始自动化流程 ===")

// 步骤1：从网站下载数据
TracePrint("步骤1: 下载数据...")
Dim objBrowser = WebBrowser.Create("chrome", "https://data.example.com", 30)
Delay(2000)

// 登录
Keyboard.InputText(@ui"用户名", "admin", True, False)
Keyboard.InputPwd(@ui"密码", "password", True)
Mouse.Click(@ui"登录按钮")
Delay(3000)

// 下载数据文件
Mouse.Click(@ui"导出按钮")
Delay(5000)
WebBrowser.Close(objBrowser)
TracePrint("数据下载完成")

// 步骤2：处理数据
TracePrint("步骤2: 处理数据...")
Dim objExcel = Excel.Open("C:\Downloads\data.xlsx", True, "")
Dim iRow = 2
Dim iProcessed = 0

Do While True
    Dim sValue = Excel.GetCell(objExcel, "A" & iRow)
    If sValue = "" Then Exit Do
    
    // 数据处理逻辑
    Dim sProcessed = String.Replace(sValue, "旧值", "新值")
    Excel.SetCell(objExcel, "B" & iRow, sProcessed)
    
    iProcessed = iProcessed + 1
    iRow = iRow + 1
Loop

Excel.SaveAs(objExcel, "C:\Reports\processed_data.xlsx")
Excel.Close(objExcel)
TracePrint("数据处理完成，共处理 " & iProcessed & " 条")

// 步骤3：生成报告
TracePrint("步骤3: 生成报告...")
Dim sReport = "数据处理报告\n"
sReport = sReport & "==================\n\n"
sReport = sReport & "处理时间: " & Time.Format(Time.Now(), "yyyy-MM-dd HH:mm:ss") & "\n"
sReport = sReport & "处理记录数: " & iProcessed & "\n"
sReport = sReport & "数据文件: C:\Reports\processed_data.xlsx\n"

File.Write("C:\Reports\report.txt", sReport, "utf-8")
TracePrint("报告生成完成")

// 步骤4：发送邮件
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
// 带完善异常处理的自动化流程
Dim sLogFile = "C:\Logs\automation_" & Time.Format(Time.Now(), "yyyyMMdd") & ".log"

Function WriteLog(sMessage)
    Dim sTime = Time.Format(Time.Now(), "yyyy-MM-dd HH:mm:ss")
    Dim sLogLine = "[" & sTime & "] " & sMessage & "\n"
    File.Append(sLogFile, sLogLine, "utf-8")
    TracePrint(sMessage)
End Function

WriteLog("=== 流程开始 ===")

Try
    // 主要业务逻辑
    WriteLog("步骤1: 打开浏览器")
    Dim objBrowser = WebBrowser.Create("chrome", "https://www.example.com", 30)
    
    WriteLog("步骤2: 执行操作")
    // ... 业务代码 ...
    
    WriteLog("步骤3: 关闭浏览器")
    WebBrowser.Close(objBrowser)
    
    WriteLog("=== 流程成功完成 ===")
    
Catch ex
    WriteLog("!!! 发生错误: " & ex.Message)
    WriteLog("错误位置: " & ex.Source)
    
    // 发送错误通知邮件
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
// 使用多种定位方式提高稳定性
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
// 智能等待元素出现
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
// 使用数组批量处理
Dim arrTasks = ["任务1", "任务2", "任务3"]
For Each sTask In arrTasks
    Try
        // 处理任务
        TracePrint("处理: " & sTask)
    Catch ex
        TracePrint("任务失败: " & sTask & " - " & ex.Message)
        Continue  // 继续处理下一个
    End Try
Next
```

---

**文档版本**: v1.0.0  
**适用版本**: UIBot 6.0.0.211215(64位)  
**更新时间**: 2024-01-15
