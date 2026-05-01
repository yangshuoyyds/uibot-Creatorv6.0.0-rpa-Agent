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

### 示例11：通用卡证识别（Mage AI）

**业务场景**：
通用卡证识别能够识别身份证、银行卡、驾驶证、营业执照等20种常见卡证，并从中抽取出核心字段值。卡证识别可以广泛应用于企业和个人的资质审核，包括但不限于银行开户、尽职调查、一网通办等场景。

**核心特点**：
- ✅ **种类丰富**：识别身份证、营业执照、驾驶证、行驶证、户口本、银行卡、护照、结婚证、车辆登记证、车辆合格证、港澳台居民来往大陆通行证、开户许可证、组织机构代码证、房产证、不动产证、军官证、出生证明、临时身份证、社保卡、外国人永久居留证等20种卡证
- ✅ **双面识别**：支持卡证正反双面识别，自动判断正面反面

**支持的卡证类型（20种）**：

| 序号 | 卡证类型 | type_key | 说明 |
|------|---------|----------|------|
| 1 | 身份证 | id_card | 支持正反面识别 |
| 2 | 营业执照 | business_license | 支持副本和电子执照 |
| 3 | 驾驶证 | drvlicense | 支持主页和副页 |
| 4 | 行驶证 | vehlicense | 支持主页和副页 |
| 5 | 户口本 | family_register | 支持户主页和成员页 |
| 6 | 银行卡 | bank_card | 识别卡号、持有人等 |
| 7 | 护照 | passport | 支持中国护照 |
| 8 | 结婚证 | marriage_certificate | 识别双方信息 |
| 9 | 车辆登记证 | vehicle_registration_certificate | 车辆详细信息 |
| 10 | 车辆合格证 | vehicle_certificate | 车辆出厂信息 |
| 11 | 港澳台通行证 | mainland_travel_permit_hk_macao_taiwan | 往来大陆通行证 |
| 12 | 开户许可证 | opening_license | 银行开户信息 |
| 13 | 组织机构代码证 | organization_certificate | 机构代码信息 |
| 14 | 房产证 | house_property_owner_ship | 房产权属信息 |
| 15 | 不动产证 | real_estate | 不动产权信息 |
| 16 | 军官证 | military_certificate | 军人身份证明 |
| 17 | 出生证明 | birth_certificate | 新生儿出生信息 |
| 18 | 临时身份证 | temporary_id_card | 临时身份证明 |
| 19 | 社保卡 | social_security_cards | 社会保障卡 |
| 20 | 外国人永久居留证 | permanent_residence_permit | 外籍人士居留证 |

**识别方式**：
- 方式1：屏幕识别 - 识别屏幕上显示的卡证
- 方式2：图片识别 - 识别本地图片文件中的卡证
- 方式3：PDF识别 - 识别PDF文件中的卡证（支持多页）

#### 方式1：屏幕卡证识别

```vb
// 屏幕卡证识别示例 - 识别屏幕上显示的身份证
Dim iPID = ""
Dim config = {
    "Pubkey": "your_pubkey",
    "Secret": "your_secret",
    "Url": "https://mage.uibot.com.cn"
}

TracePrint("=== 开始屏幕卡证识别 ===")

// 打开身份证图片（使用系统默认图片查看器）
iPID = App.Start("C:\Cards\id_card.jpg", "0", "1")
Delay(2000)

Try
    // 使用 Mage.ScreenOCRCard 识别屏幕上的卡证
    With Mage.ScreenOCRCard(@ui"图片查看窗口", Null, config, 30000)
        // 获取卡证类型
        Dim sCardType = .ExtractCardType()
        TracePrint("识别到卡证类型: " & sCardType)
        
        // 根据不同卡证类型提取字段
        Select Case .ExtractCardType()
            Case Alias("id_card", "身份证")
                TracePrint("--- 身份证信息 ---")
                Dim sName = .ExtractCardInfo("id_card", "name")
                Dim sSex = .ExtractCardInfo("id_card", "sex")
                Dim sNationality = .ExtractCardInfo("id_card", "nationality")
                Dim sBirth = .ExtractCardInfo("id_card", "birth")
                Dim sAddress = .ExtractCardInfo("id_card", "address")
                Dim sIdNumber = .ExtractCardInfo("id_card", "id_number")
                Dim sIssueAuthority = .ExtractCardInfo("id_card", "issue_authority")
                Dim sValidateDate = .ExtractCardInfo("id_card", "validate_date")
                
                TracePrint("姓名: " & sName)
                TracePrint("性别: " & sSex)
                TracePrint("民族: " & sNationality)
                TracePrint("出生: " & sBirth)
                TracePrint("地址: " & sAddress)
                TracePrint("身份证号: " & sIdNumber)
                TracePrint("签发机关: " & sIssueAuthority)
                TracePrint("有效期限: " & sValidateDate)
                
            Case Alias("business_license", "营业执照")
                TracePrint("--- 营业执照信息 ---")
                Dim sCreditCode = .ExtractCardInfo("business_license", "BizLicenseCreditCode")
                Dim sCompanyName = .ExtractCardInfo("business_license", "BizLicenseCompanyName")
                Dim sOwnerName = .ExtractCardInfo("business_license", "BizLicenseOwnerName")
                Dim sRegCapital = .ExtractCardInfo("business_license", "BizLicenseRegCapital")
                Dim sStartTime = .ExtractCardInfo("business_license", "BizLicenseStartTime")
                
                TracePrint("统一社会信用代码: " & sCreditCode)
                TracePrint("企业名称: " & sCompanyName)
                TracePrint("法定代表人: " & sOwnerName)
                TracePrint("注册资本: " & sRegCapital)
                TracePrint("成立日期: " & sStartTime)
                
            Case Alias("bank_card", "银行卡")
                TracePrint("--- 银行卡信息 ---")
                Dim sCardNumber = .ExtractCardInfo("bank_card", "card_number")
                Dim sHolderName = .ExtractCardInfo("bank_card", "holder_name")
                Dim sIssuer = .ExtractCardInfo("bank_card", "issuer")
                Dim sValidate = .ExtractCardInfo("bank_card", "validate")
                
                TracePrint("卡号: " & sCardNumber)
                TracePrint("持卡人: " & sHolderName)
                TracePrint("发卡行: " & sIssuer)
                TracePrint("有效期: " & sValidate)
                
            Case Else
                TracePrint("识别到其他类型卡证: " & sCardType)
        End Select
    End With
    
    TracePrint("=== 屏幕卡证识别完成 ===")
    
Catch ex
    TracePrint("识别失败: " & ex.Message)
End Try

// 关闭图片查看器
App.Close(iPID)
```

#### 方式2：图片卡证识别

```vb
// 图片卡证识别示例 - 识别本地图片文件中的身份证
Dim imagePath = "C:\Cards\id_card.jpg"
Dim config = {
    "Pubkey": "your_pubkey",
    "Secret": "your_secret",
    "Url": "https://mage.uibot.com.cn"
}

TracePrint("=== 开始图片卡证识别 ===")
TracePrint("图片路径: " & imagePath)

Try
    // 使用 Mage.ImageOCRCard 识别图片中的卡证
    With Mage.ImageOCRCard(imagePath, config, 30000)
        // 获取卡证类型
        Dim sCardType = .ExtractCardType()
        TracePrint("识别到卡证类型: " & sCardType)
        
        // 提取身份证信息
        If sCardType = "id_card" Or sCardType = "身份证" Then
            TracePrint("--- 身份证详细信息 ---")
            
            // 基本信息
            Dim sName = .ExtractCardInfo("id_card", "name")
            Dim sSex = .ExtractCardInfo("id_card", "sex")
            Dim sNationality = .ExtractCardInfo("id_card", "nationality")
            Dim sBirth = .ExtractCardInfo("id_card", "birth")
            Dim sAddress = .ExtractCardInfo("id_card", "address")
            Dim sIdNumber = .ExtractCardInfo("id_card", "id_number")
            
            // 证件信息
            Dim sIssueAuthority = .ExtractCardInfo("id_card", "issue_authority")
            Dim sValidateDate = .ExtractCardInfo("id_card", "validate_date")
            
            // 输出识别结果
            TracePrint("姓名: " & sName)
            TracePrint("性别: " & sSex)
            TracePrint("民族: " & sNationality)
            TracePrint("出生: " & sBirth)
            TracePrint("地址: " & sAddress)
            TracePrint("身份证号: " & sIdNumber)
            TracePrint("签发机关: " & sIssueAuthority)
            TracePrint("有效期限: " & sValidateDate)
            
            // 保存识别结果到 Excel
            Dim objExcel = Excel.Create(True, "")
            Excel.SetCell(objExcel, "A1", "字段名称")
            Excel.SetCell(objExcel, "B1", "识别结果")
            Excel.SetCell(objExcel, "A2", "姓名")
            Excel.SetCell(objExcel, "B2", sName)
            Excel.SetCell(objExcel, "A3", "性别")
            Excel.SetCell(objExcel, "B3", sSex)
            Excel.SetCell(objExcel, "A4", "民族")
            Excel.SetCell(objExcel, "B4", sNationality)
            Excel.SetCell(objExcel, "A5", "出生")
            Excel.SetCell(objExcel, "B5", sBirth)
            Excel.SetCell(objExcel, "A6", "地址")
            Excel.SetCell(objExcel, "B6", sAddress)
            Excel.SetCell(objExcel, "A7", "身份证号")
            Excel.SetCell(objExcel, "B7", sIdNumber)
            Excel.SetCell(objExcel, "A8", "签发机关")
            Excel.SetCell(objExcel, "B8", sIssueAuthority)
            Excel.SetCell(objExcel, "A9", "有效期限")
            Excel.SetCell(objExcel, "B9", sValidateDate)
            
            Dim sResultFile = "C:\Cards\id_card_result_" & Time.Format(Time.Now(), "yyyyMMdd_HHmmss") & ".xlsx"
            Excel.SaveAs(objExcel, sResultFile)
            Excel.Close(objExcel)
            TracePrint("识别结果已保存到: " & sResultFile)
        Else
            TracePrint("当前示例仅处理身份证，识别到的类型为: " & sCardType)
        End If
    End With
    
    TracePrint("=== 图片卡证识别完成 ===")
    
Catch ex
    TracePrint("识别失败: " & ex.Message)
End Try
```

#### 方式3：PDF卡证识别

```vb
// PDF卡证识别示例 - 识别PDF文件中的卡证（支持多页）
Dim config = {
    "Pubkey": "your_pubkey",
    "Secret": "your_secret",
    "Url": "https://mage.uibot.com.cn"
}
Dim pdfPath = "C:\Cards\cards.pdf"
Dim password = ""  // PDF密码，无密码则为空字符串
Dim readAll = True  // True: 识别所有页，False: 识别指定页
Dim pages = "1-3"  // 指定页码范围，如 "1-3" 或 "1,3,5"
Dim interval = 1000  // 每页识别间隔（毫秒）
Dim timeout = 60

TracePrint("=== 开始PDF卡证识别 ===")
TracePrint("PDF路径: " & pdfPath)

Try
    Dim iPageCount = 0
    Dim arrResults = []  // 存储所有识别结果
    
    // 使用 Mage.PDFOCRCard 识别PDF中的卡证
    With Each Mage.PDFOCRCard(config, pdfPath, password, readAll, pages, interval, timeout)
        iPageCount = iPageCount + 1
        TracePrint("--- 第 " & iPageCount & " 页 ---")
        
        // 获取卡证类型
        Dim sCardType = .ExtractCardType()
        TracePrint("卡证类型: " & sCardType)
        
        // 创建结果对象
        Dim objResult = {
            "page": iPageCount,
            "type": sCardType,
            "data": {}
        }
        
        // 根据卡证类型提取信息
        Select Case .ExtractCardType()
            Case Alias("id_card", "身份证")
                objResult["data"]["name"] = .ExtractCardInfo("id_card", "name")
                objResult["data"]["sex"] = .ExtractCardInfo("id_card", "sex")
                objResult["data"]["nationality"] = .ExtractCardInfo("id_card", "nationality")
                objResult["data"]["birth"] = .ExtractCardInfo("id_card", "birth")
                objResult["data"]["address"] = .ExtractCardInfo("id_card", "address")
                objResult["data"]["id_number"] = .ExtractCardInfo("id_card", "id_number")
                
                TracePrint("姓名: " & objResult["data"]["name"])
                TracePrint("性别: " & objResult["data"]["sex"])
                TracePrint("民族: " & objResult["data"]["nationality"])
                TracePrint("身份证号: " & objResult["data"]["id_number"])
                
            Case Alias("business_license", "营业执照")
                objResult["data"]["credit_code"] = .ExtractCardInfo("business_license", "BizLicenseCreditCode")
                objResult["data"]["company_name"] = .ExtractCardInfo("business_license", "BizLicenseCompanyName")
                objResult["data"]["owner_name"] = .ExtractCardInfo("business_license", "BizLicenseOwnerName")
                objResult["data"]["reg_capital"] = .ExtractCardInfo("business_license", "BizLicenseRegCapital")
                
                TracePrint("企业名称: " & objResult["data"]["company_name"])
                TracePrint("统一社会信用代码: " & objResult["data"]["credit_code"])
                TracePrint("法定代表人: " & objResult["data"]["owner_name"])
                
            Case Alias("drvlicense", "驾驶证")
                objResult["data"]["name"] = .ExtractCardInfo("drvlicense", "name")
                objResult["data"]["license_number"] = .ExtractCardInfo("drvlicense", "driving_license_main_number")
                objResult["data"]["drive_type"] = .ExtractCardInfo("drvlicense", "drive_type")
                objResult["data"]["valid_period"] = .ExtractCardInfo("drvlicense", "valid_period")
                
                TracePrint("姓名: " & objResult["data"]["name"])
                TracePrint("证号: " & objResult["data"]["license_number"])
                TracePrint("准驾车型: " & objResult["data"]["drive_type"])
                
            Case Else
                TracePrint("其他类型卡证，请根据需要添加处理逻辑")
        End Select
        
        // 添加到结果数组
        Array.Push(arrResults, objResult)
    End With
    
    TracePrint("=== PDF卡证识别完成 ===")
    TracePrint("共识别 " & iPageCount & " 页卡证")
    
    // 将所有结果导出到 Excel
    If Array.Length(arrResults) > 0 Then
        Dim objExcel = Excel.Create(True, "")
        
        // 写入表头
        Excel.SetCell(objExcel, "A1", "页码")
        Excel.SetCell(objExcel, "B1", "卡证类型")
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
        Dim sResultFile = "C:\Cards\pdf_cards_result_" & Time.Format(Time.Now(), "yyyyMMdd_HHmmss") & ".xlsx"
        Excel.SaveAs(objExcel, sResultFile)
        Excel.Close(objExcel)
        TracePrint("所有识别结果已保存到: " & sResultFile)
    End If
    
Catch ex
    TracePrint("识别失败: " & ex.Message)
End Try
```

#### 批量卡证识别与分类

```vb
// 批量卡证识别与分类示例
Dim sFolderPath = "C:\Cards\Batch"  // 卡证图片文件夹
Dim arrFiles = File.GetFileList(sFolderPath, "*.jpg|*.png|*.jpeg", False)
Dim config = {
    "Pubkey": "your_pubkey",
    "Secret": "your_secret",
    "Url": "https://mage.uibot.com.cn"
}

TracePrint("=== 开始批量卡证识别 ===")
TracePrint("待识别文件数: " & Array.Length(arrFiles))

// 创建分类统计
Dim objStats = {
    "id_card": 0,
    "business_license": 0,
    "bank_card": 0,
    "drvlicense": 0,
    "other": 0
}

// 创建 Excel 汇总表
Dim objExcel = Excel.Create(True, "")
Excel.SetCell(objExcel, "A1", "序号")
Excel.SetCell(objExcel, "B1", "文件名")
Excel.SetCell(objExcel, "C1", "卡证类型")
Excel.SetCell(objExcel, "D1", "关键信息")

Dim iRow = 2
Dim iSuccess = 0
Dim iFailed = 0

For Each sFile In arrFiles
    TracePrint("处理文件: " & File.GetName(sFile))
    
    Try
        With Mage.ImageOCRCard(sFile, config, 30000)
            Dim sType = .ExtractCardType()
            Dim sInfo = ""
            
            // 根据类型提取关键信息
            Select Case sType
                Case Alias("id_card", "身份证")
                    objStats["id_card"] = objStats["id_card"] + 1
                    sInfo = "姓名: " & .ExtractCardInfo("id_card", "name") & _
                            ", 身份证号: " & .ExtractCardInfo("id_card", "id_number")
                    
                Case Alias("business_license", "营业执照")
                    objStats["business_license"] = objStats["business_license"] + 1
                    sInfo = "企业: " & .ExtractCardInfo("business_license", "BizLicenseCompanyName") & _
                            ", 信用代码: " & .ExtractCardInfo("business_license", "BizLicenseCreditCode")
                    
                Case Alias("bank_card", "银行卡")
                    objStats["bank_card"] = objStats["bank_card"] + 1
                    sInfo = "卡号: " & .ExtractCardInfo("bank_card", "card_number") & _
                            ", 持卡人: " & .ExtractCardInfo("bank_card", "holder_name")
                    
                Case Alias("drvlicense", "驾驶证")
                    objStats["drvlicense"] = objStats["drvlicense"] + 1
                    sInfo = "姓名: " & .ExtractCardInfo("drvlicense", "name") & _
                            ", 准驾车型: " & .ExtractCardInfo("drvlicense", "drive_type")
                    
                Case Else
                    objStats["other"] = objStats["other"] + 1
                    sInfo = "其他类型"
            End Select
            
            // 写入 Excel
            Excel.SetCell(objExcel, "A" & iRow, iRow - 1)
            Excel.SetCell(objExcel, "B" & iRow, File.GetName(sFile))
            Excel.SetCell(objExcel, "C" & iRow, sType)
            Excel.SetCell(objExcel, "D" & iRow, sInfo)
            
            iRow = iRow + 1
            iSuccess = iSuccess + 1
            TracePrint("识别成功: " & sType)
        End With
        
    Catch ex
        iFailed = iFailed + 1
        TracePrint("识别失败: " & ex.Message)
    End Try
    
    Delay(500)  // 避免请求过快
Next

// 保存汇总表
Dim sResultFile = "C:\Cards\batch_result_" & Time.Format(Time.Now(), "yyyyMMdd_HHmmss") & ".xlsx"
Excel.SaveAs(objExcel, sResultFile)
Excel.Close(objExcel)

// 输出统计信息
TracePrint("=== 批量识别完成 ===")
TracePrint("总文件数: " & Array.Length(arrFiles))
TracePrint("成功: " & iSuccess & ", 失败: " & iFailed)
TracePrint("--- 卡证类型统计 ---")
TracePrint("身份证: " & objStats["id_card"])
TracePrint("营业执照: " & objStats["business_license"])
TracePrint("银行卡: " & objStats["bank_card"])
TracePrint("驾驶证: " & objStats["drvlicense"])
TracePrint("其他: " & objStats["other"])
TracePrint("结果已保存到: " & sResultFile)
```

**使用说明**：
1. **配置要求**：需要配置 Mage AI 的 Pubkey 和 Secret（在来也科技平台申请）
2. **卡证类型**：支持20种卡证类型，详见上方表格中的 type_key
3. **字段提取**：不同卡证类型支持的字段不同，常见字段包括：
   - 身份证：name（姓名）、sex（性别）、nationality（民族）、birth（出生）、address（地址）、id_number（身份证号）、issue_authority（签发机关）、validate_date（有效期限）
   - 营业执照：BizLicenseCreditCode（统一社会信用代码）、BizLicenseCompanyName（企业名称）、BizLicenseOwnerName（法定代表人）、BizLicenseRegCapital（注册资本）、BizLicenseStartTime（成立日期）
   - 银行卡：card_number（卡号）、holder_name（持卡人）、issuer（发卡行）、validate（有效期）
   - 驾驶证：name（姓名）、driving_license_main_number（证号）、drive_type（准驾车型）、valid_period（有效期限）
4. **性能优化**：
   - 建议添加适当的延迟（500-1000ms）避免请求过快
   - 对于大批量识别，建议分批处理并保存中间结果
   - PDF识别时可以指定页码范围，避免处理不必要的页面
5. **错误处理**：建议使用 Try-Catch 包裹识别代码，处理网络异常、识别失败等情况
6. **参考文档**：完整的字段列表和API说明请参考 [来也IDP官方文档](https://documents.laiye.com/idp-mage/docs/OCR/ocr_card)

**成功案例**：
某医药企业的业务系统中维护了几万个客户，每个客户会提交事业单位法人证、医疗机构执业许可证等材料，以证明自己有采购疫苗的资质。新客户直接审核，老客户每年会进行年审。审核的业务流程主要是用卡证上识别的关键信息和业务系统中填写的字段进行比对。涉及到6个页面、3种客户类型，每个客户类型审核逻辑各不相同，引入RPA+AI能够极大的减少工作量、提升审核效率、减少犯错几率。

### 示例12：验证码识别（Mage AI）

**业务场景**：
验证码是RPA+AI场景中使用最高频的AI应用。RPA能够模拟人工进行鼠标键盘的操作，将办公流程自动化。通常在RPA操作业务系统的过程中，会遇到需要输入验证码的情况。比如：登录银行网银下载流水、进入增值税发票查验平台验证发票的真实性，都需要填写验证码信息。验证码识别能力可以识别纯英文、纯数字、英文数字组合、四则运算、滑块等验证码图片，让流程全自动化，不再需要人工介入。

**核心特点**：
- ✅ **支持多种验证码类型**：纯英文、纯数字、英文数字组合、四则运算、滑块验证码等
- ✅ **识别速度快**：响应速度快，识别结果秒回
- ✅ **高准确率**：针对常见验证码类型进行了优化训练

**识别方式**：
- 方式1：屏幕验证码识别 - 识别屏幕上显示的验证码
- 方式2：图片验证码识别 - 识别本地图片文件中的验证码

#### 方式1：屏幕验证码识别

```vb
// 屏幕验证码识别示例 - 识别网页上的验证码并自动填写
Dim hWeb = ""
Dim sVerifyCode = ""
Dim config = {
    "Pubkey": "your_pubkey",
    "Secret": "your_secret",
    "Url": "https://mage.uibot.com.cn"
}

TracePrint("=== 开始屏幕验证码识别 ===")

Try
    // 打开需要登录的网页（以交通银行企业网银为例）
    TracePrint("打开登录页面...")
    hWeb = WebBrowser.Create("chrome", "https://ebank.95559.com.cn/CEBS/cebs/logon.do", 30000)
    Delay(3000)
    
    // 使用 Mage.ScreenOCRVerifyCode 识别屏幕上的验证码
    TracePrint("识别验证码...")
    sVerifyCode = Mage.ScreenOCRVerifyCode(@ui"验证码图片元素", Null, config, 30000)
    
    TracePrint("识别到验证码: " & sVerifyCode)
    
    // 自动填写验证码
    If sVerifyCode <> "" Then
        Keyboard.InputText(@ui"验证码输入框", sVerifyCode, True, False)
        TracePrint("验证码已自动填写")
        
        // 继续填写其他登录信息
        Keyboard.InputText(@ui"用户名输入框", "your_username", True, False)
        Keyboard.InputPwd(@ui"密码输入框", "your_password", True)
        
        // 点击登录按钮
        Mouse.Click(@ui"登录按钮")
        Delay(3000)
        
        TracePrint("登录流程完成")
    Else
        TracePrint("验证码识别失败，请重试")
    End If
    
    TracePrint("=== 屏幕验证码识别完成 ===")
    
Catch ex
    TracePrint("识别失败: " & ex.Message)
Finally
    // 可选：关闭浏览器
    // WebBrowser.Close(hWeb)
End Try
```

#### 方式2：图片验证码识别

```vb
// 图片验证码识别示例 - 识别本地图片文件中的验证码
Dim imagePath = "C:\Captcha\verify_code.png"
Dim config = {
    "Pubkey": "your_pubkey",
    "Secret": "your_secret",
    "Url": "https://mage.uibot.com.cn"
}

TracePrint("=== 开始图片验证码识别 ===")
TracePrint("图片路径: " & imagePath)

Try
    // 使用 Mage.ImageOCRVerifyCode 识别图片中的验证码
    Dim sVerifyCode = Mage.ImageOCRVerifyCode(imagePath, config, 30000)
    
    If sVerifyCode <> "" Then
        TracePrint("识别到验证码: " & sVerifyCode)
        
        // 保存识别结果到文本文件
        Dim sResultFile = "C:\Captcha\verify_code_result.txt"
        File.Write(sResultFile, "验证码识别结果: " & sVerifyCode, "utf-8")
        TracePrint("识别结果已保存到: " & sResultFile)
    Else
        TracePrint("验证码识别失败")
    End If
    
    TracePrint("=== 图片验证码识别完成 ===")
    
Catch ex
    TracePrint("识别失败: " & ex.Message)
End Try
```

#### 批量验证码识别

```vb
// 批量验证码识别示例 - 识别文件夹中的多个验证码图片
Dim sFolderPath = "C:\Captcha\Batch"
Dim arrFiles = File.GetFileList(sFolderPath, "*.png|*.jpg|*.jpeg", False)
Dim config = {
    "Pubkey": "your_pubkey",
    "Secret": "your_secret",
    "Url": "https://mage.uibot.com.cn"
}

TracePrint("=== 开始批量验证码识别 ===")
TracePrint("待识别文件数: " & Array.Length(arrFiles))

// 创建 Excel 汇总表
Dim objExcel = Excel.Create(True, "")
Excel.SetCell(objExcel, "A1", "序号")
Excel.SetCell(objExcel, "B1", "文件名")
Excel.SetCell(objExcel, "C1", "识别结果")
Excel.SetCell(objExcel, "D1", "识别状态")

Dim iRow = 2
Dim iSuccess = 0
Dim iFailed = 0

For Each sFile In arrFiles
    TracePrint("处理文件: " & File.GetName(sFile))
    
    Try
        // 识别验证码
        Dim sVerifyCode = Mage.ImageOCRVerifyCode(sFile, config, 30000)
        
        If sVerifyCode <> "" Then
            // 写入 Excel
            Excel.SetCell(objExcel, "A" & iRow, iRow - 1)
            Excel.SetCell(objExcel, "B" & iRow, File.GetName(sFile))
            Excel.SetCell(objExcel, "C" & iRow, sVerifyCode)
            Excel.SetCell(objExcel, "D" & iRow, "成功")
            
            iSuccess = iSuccess + 1
            TracePrint("识别成功: " & sVerifyCode)
        Else
            Excel.SetCell(objExcel, "A" & iRow, iRow - 1)
            Excel.SetCell(objExcel, "B" & iRow, File.GetName(sFile))
            Excel.SetCell(objExcel, "C" & iRow, "")
            Excel.SetCell(objExcel, "D" & iRow, "失败")
            
            iFailed = iFailed + 1
            TracePrint("识别失败")
        End If
        
        iRow = iRow + 1
        
    Catch ex
        iFailed = iFailed + 1
        TracePrint("识别异常: " & ex.Message)
    End Try
    
    Delay(500)  // 避免请求过快
Next

// 保存汇总表
Dim sResultFile = "C:\Captcha\batch_result_" & Time.Format(Time.Now(), "yyyyMMdd_HHmmss") & ".xlsx"
Excel.SaveAs(objExcel, sResultFile)
Excel.Close(objExcel)

// 输出统计信息
TracePrint("=== 批量识别完成 ===")
TracePrint("总文件数: " & Array.Length(arrFiles))
TracePrint("成功: " & iSuccess & ", 失败: " & iFailed)
TracePrint("结果已保存到: " & sResultFile)
```

**使用说明**：
1. **配置要求**：需要配置 Mage AI 的 Pubkey 和 Secret（在来也科技平台申请）
2. **支持的验证码类型**：
   - 纯数字验证码（如：1234）
   - 纯英文验证码（如：ABCD）
   - 英文数字组合（如：A1B2）
   - 四则运算验证码（如：3+5=?）
   - 滑块验证码
3. **性能优化**：
   - 建议添加适当的延迟（500-1000ms）避免请求过快
   - 对于大批量识别，建议分批处理
4. **错误处理**：建议使用 Try-Catch 包裹识别代码，处理网络异常、识别失败等情况
5. **参考文档**：完整的API说明请参考 [来也IDP官方文档](https://documents.laiye.com/idp-mage/docs/OCR/ocr_captcha)

**应用场景**：
- 银行网银登录自动化
- 发票查验平台自动验证
- 政务系统自动登录
- 电商平台自动下单
- 各类需要验证码验证的业务系统自动化

### 示例13：印章识别（Mage AI）

**业务场景**：
印章识别能够识别合同、票据、卡证、表格文档上是否加盖过印章，并返回印章文字内容、所在位置、颜色。常用于合同审批、财务报销、资质审核等场景。主要支持公司用椭圆章、圆章、长方形章。

**核心特点**：
- ✅ **一图多章**：能够识别一张图片上的多个印章，在印章项目遮挡的情况下也能正确检测印章
- ✅ **印章颜色识别**：在审批流程中，需要核实印章是复印出来的还是新加盖的，通常需要看印章是黑白还是彩色的。印章识别能够准确返回印章颜色
- ✅ **支持多种印章类型**：椭圆章、圆章、长方形章等

**识别版本**：
- 标准版：基础印章识别功能
- 高级版：更高的识别准确率和更多的功能支持

#### 印章识别标准版

```vb
// 印章识别标准版示例
Dim stampResultJSON = ""
Dim config = {
    "Url": "https://mage.uibot.com.cn",
    "Pubkey": "your_pubkey_standard",
    "Secret": "your_secret_standard"
}
Dim imagePath = "C:\Stamps\contract_with_stamp.jpg"

TracePrint("=== 开始印章识别（标准版） ===")
TracePrint("图片路径: " & imagePath)

Try
    // 使用标准版印章识别
    stampResultJSON = laiyeUiBotMageV0.StampIdentify(config["Url"], config["Pubkey"], config["Secret"], imagePath)
    
    TracePrint("识别结果（JSON）:")
    TracePrint(stampResultJSON)
    
    // 解析 JSON 结果
    Dim objResult = JSON.Parse(stampResultJSON)
    
    If objResult["code"] = 0 Then
        Dim arrStamps = objResult["data"]["stamps"]
        TracePrint("识别到 " & Array.Length(arrStamps) & " 个印章")
        
        Dim iIndex = 1
        For Each objStamp In arrStamps
            TracePrint("--- 印章 " & iIndex & " ---")
            TracePrint("印章文字: " & objStamp["text"])
            TracePrint("印章颜色: " & objStamp["color"])
            TracePrint("位置信息: X=" & objStamp["x"] & ", Y=" & objStamp["y"] & ", W=" & objStamp["width"] & ", H=" & objStamp["height"])
            iIndex = iIndex + 1
        Next
        
        // 保存识别结果到 Excel
        Dim objExcel = Excel.Create(True, "")
        Excel.SetCell(objExcel, "A1", "序号")
        Excel.SetCell(objExcel, "B1", "印章文字")
        Excel.SetCell(objExcel, "C1", "印章颜色")
        Excel.SetCell(objExcel, "D1", "位置X")
        Excel.SetCell(objExcel, "E1", "位置Y")
        Excel.SetCell(objExcel, "F1", "宽度")
        Excel.SetCell(objExcel, "G1", "高度")
        
        Dim iRow = 2
        For Each objStamp In arrStamps
            Excel.SetCell(objExcel, "A" & iRow, iRow - 1)
            Excel.SetCell(objExcel, "B" & iRow, objStamp["text"])
            Excel.SetCell(objExcel, "C" & iRow, objStamp["color"])
            Excel.SetCell(objExcel, "D" & iRow, objStamp["x"])
            Excel.SetCell(objExcel, "E" & iRow, objStamp["y"])
            Excel.SetCell(objExcel, "F" & iRow, objStamp["width"])
            Excel.SetCell(objExcel, "G" & iRow, objStamp["height"])
            iRow = iRow + 1
        Next
        
        Dim sResultFile = "C:\Stamps\stamp_result_" & Time.Format(Time.Now(), "yyyyMMdd_HHmmss") & ".xlsx"
        Excel.SaveAs(objExcel, sResultFile)
        Excel.Close(objExcel)
        TracePrint("识别结果已保存到: " & sResultFile)
    Else
        TracePrint("识别失败: " & objResult["message"])
    End If
    
    TracePrint("=== 印章识别完成 ===")
    
Catch ex
    TracePrint("识别失败: " & ex.Message)
End Try
```

#### 印章识别高级版

```vb
// 印章识别高级版示例 - 更高的准确率
Dim stampResultJSON = ""
Dim config = {
    "Url": "https://mage.uibot.com.cn",
    "Pubkey": "your_pubkey_advanced",
    "Secret": "your_secret_advanced"
}
Dim imagePath = "C:\Stamps\contract_with_stamp.jpg"

TracePrint("=== 开始印章识别（高级版） ===")
TracePrint("图片路径: " & imagePath)

Try
    // 使用高级版印章识别
    stampResultJSON = laiyeUiBotMageV0.StampIdentify(config["Url"], config["Pubkey"], config["Secret"], imagePath)
    
    TracePrint("识别结果（JSON）:")
    TracePrint(stampResultJSON)
    
    // 解析 JSON 结果
    Dim objResult = JSON.Parse(stampResultJSON)
    
    If objResult["code"] = 0 Then
        Dim arrStamps = objResult["data"]["stamps"]
        TracePrint("识别到 " & Array.Length(arrStamps) & " 个印章")
        
        // 详细输出每个印章信息
        For Each objStamp In arrStamps
            TracePrint("--- 印章详细信息 ---")
            TracePrint("印章文字: " & objStamp["text"])
            TracePrint("印章类型: " & objStamp["type"])  // 圆章、椭圆章、方章
            TracePrint("印章颜色: " & objStamp["color"])  // 红色、蓝色、黑白等
            TracePrint("置信度: " & objStamp["confidence"])
            TracePrint("位置: (" & objStamp["x"] & ", " & objStamp["y"] & ")")
            TracePrint("尺寸: " & objStamp["width"] & " x " & objStamp["height"])
        Next
    Else
        TracePrint("识别失败: " & objResult["message"])
    End If
    
    TracePrint("=== 印章识别完成 ===")
    
Catch ex
    TracePrint("识别失败: " & ex.Message)
End Try
```

#### 批量印章识别与审核

```vb
// 批量印章识别与审核示例 - 用于合同审批流程
Dim sFolderPath = "C:\Stamps\Contracts"
Dim arrFiles = File.GetFileList(sFolderPath, "*.jpg|*.png|*.pdf", False)
Dim config = {
    "Url": "https://mage.uibot.com.cn",
    "Pubkey": "your_pubkey",
    "Secret": "your_secret"
}

TracePrint("=== 开始批量印章识别 ===")
TracePrint("待识别文件数: " & Array.Length(arrFiles))

// 创建 Excel 汇总表
Dim objExcel = Excel.Create(True, "")
Excel.SetCell(objExcel, "A1", "序号")
Excel.SetCell(objExcel, "B1", "文件名")
Excel.SetCell(objExcel, "C1", "印章数量")
Excel.SetCell(objExcel, "D1", "印章文字")
Excel.SetCell(objExcel, "E1", "印章颜色")
Excel.SetCell(objExcel, "F1", "审核状态")

Dim iRow = 2
Dim iSuccess = 0
Dim iFailed = 0

For Each sFile In arrFiles
    TracePrint("处理文件: " & File.GetName(sFile))
    
    Try
        // 识别印章
        Dim stampResultJSON = laiyeUiBotMageV0.StampIdentify(config["Url"], config["Pubkey"], config["Secret"], sFile)
        Dim objResult = JSON.Parse(stampResultJSON)
        
        If objResult["code"] = 0 Then
            Dim arrStamps = objResult["data"]["stamps"]
            Dim iStampCount = Array.Length(arrStamps)
            
            // 收集印章信息
            Dim sStampTexts = ""
            Dim sStampColors = ""
            Dim bHasColorStamp = False
            
            For Each objStamp In arrStamps
                sStampTexts = sStampTexts & objStamp["text"] & "; "
                sStampColors = sStampColors & objStamp["color"] & "; "
                
                // 检查是否有彩色印章（非黑白）
                If objStamp["color"] <> "黑白" And objStamp["color"] <> "灰色" Then
                    bHasColorStamp = True
                End If
            Next
            
            // 审核状态判断
            Dim sAuditStatus = ""
            If iStampCount = 0 Then
                sAuditStatus = "未盖章"
            ElseIf Not bHasColorStamp Then
                sAuditStatus = "疑似复印件"
            Else
                sAuditStatus = "审核通过"
            End If
            
            // 写入 Excel
            Excel.SetCell(objExcel, "A" & iRow, iRow - 1)
            Excel.SetCell(objExcel, "B" & iRow, File.GetName(sFile))
            Excel.SetCell(objExcel, "C" & iRow, iStampCount)
            Excel.SetCell(objExcel, "D" & iRow, sStampTexts)
            Excel.SetCell(objExcel, "E" & iRow, sStampColors)
            Excel.SetCell(objExcel, "F" & iRow, sAuditStatus)
            
            iSuccess = iSuccess + 1
            TracePrint("识别成功: " & iStampCount & " 个印章, 状态: " & sAuditStatus)
        Else
            Excel.SetCell(objExcel, "A" & iRow, iRow - 1)
            Excel.SetCell(objExcel, "B" & iRow, File.GetName(sFile))
            Excel.SetCell(objExcel, "F" & iRow, "识别失败")
            
            iFailed = iFailed + 1
            TracePrint("识别失败")
        End If
        
        iRow = iRow + 1
        
    Catch ex
        iFailed = iFailed + 1
        TracePrint("识别异常: " & ex.Message)
    End Try
    
    Delay(500)  // 避免请求过快
Next

// 保存汇总表
Dim sResultFile = "C:\Stamps\batch_result_" & Time.Format(Time.Now(), "yyyyMMdd_HHmmss") & ".xlsx"
Excel.SaveAs(objExcel, sResultFile)
Excel.Close(objExcel)

// 输出统计信息
TracePrint("=== 批量识别完成 ===")
TracePrint("总文件数: " & Array.Length(arrFiles))
TracePrint("成功: " & iSuccess & ", 失败: " & iFailed)
TracePrint("结果已保存到: " & sResultFile)
```

**使用说明**：
1. **配置要求**：
   - 标准版和高级版需要不同的 Pubkey 和 Secret
   - 在来也科技平台申请对应版本的密钥
2. **识别结果字段**：
   - text：印章文字内容
   - color：印章颜色（红色、蓝色、黑白等）
   - type：印章类型（圆章、椭圆章、方章）
   - x, y：印章位置坐标
   - width, height：印章尺寸
   - confidence：识别置信度（高级版）
3. **应用场景**：
   - 合同审批：检查合同是否加盖公章
   - 财务报销：验证报销单据上的印章
   - 资质审核：核实证照上的印章真实性
   - 文档归档：自动识别并分类带印章的文档
4. **审核规则**：
   - 未盖章：印章数量为0
   - 疑似复印件：只有黑白印章
   - 审核通过：有彩色印章
5. **性能优化**：
   - 建议添加适当的延迟（500-1000ms）避免请求过快
   - 对于大批量识别，建议分批处理
6. **参考文档**：完整的API说明请参考 [来也IDP官方文档](https://documents.laiye.com/idp-mage/docs/OCR/ocr_stamp)

**成功案例**：
在合同审批场景中，企业每天需要处理大量合同文档，人工检查印章不仅效率低下，还容易出错。通过RPA+印章识别AI，可以自动检测合同上是否加盖了公章，印章颜色是否为彩色（排除复印件），大大提升了审批效率和准确性。

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

## 智能文档处理平台 API 集成

### API 接口配置

**基础信息**：
- **域名**：`https://cloud.laiye.com/idp`
- **请求协议**：HTTPS
- **数据格式**：JSON

### 应用签名验证

使用应用签名验证方式调用 API，需要在 HTTP Header 中添加以下字段：

| Header Key | 描述 |
|------------|------|
| Api-Auth-pubkey | 用户创建应用的 pubkey |
| Api-Auth-timestamp | 当前时间戳（秒） |
| Api-Auth-nonce | 随机字符串 |
| Api-Auth-sign | 签名(signature)，生成规则：(Api-Auth-nonce+Api-Auth-timestamp+secret_key)的sha1值 |

### 示例14：生成 API 请求头（VBScript）

```vb
// 生成来也 IDP API 请求头
Function GenerateIDPHeader(sPubkey, sSecretKey)
    Dim objHeader, sTimestamp, sNonce, sSign, sSignSource
    
    // 创建 Header 字典
    Set objHeader = CreateObject("Scripting.Dictionary")
    
    // 生成时间戳（秒）
    sTimestamp = CStr(DateDiff("s", "1970-01-01 00:00:00", Now()))
    
    // 生成随机字符串（10位）
    sNonce = GenerateRandomString(10)
    
    // 生成签名
    sSignSource = sNonce & sTimestamp & sSecretKey
    sSign = SHA1Hash(sSignSource)
    
    // 添加到 Header
    objHeader.Add "Api-Auth-pubkey", sPubkey
    objHeader.Add "Api-Auth-timestamp", sTimestamp
    objHeader.Add "Api-Auth-nonce", sNonce
    objHeader.Add "Api-Auth-sign", sSign
    objHeader.Add "Content-Type", "application/json"
    
    Set GenerateIDPHeader = objHeader
End Function

// 生成随机字符串
Function GenerateRandomString(iLength)
    Dim sChars, sResult, i
    sChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    sResult = ""
    
    Randomize
    For i = 1 To iLength
        sResult = sResult & Mid(sChars, Int(Rnd() * Len(sChars)) + 1, 1)
    Next
    
    GenerateRandomString = sResult
End Function

// SHA1 哈希函数
Function SHA1Hash(sInput)
    Dim objSHA1, arrBytes, i, sHash
    
    // 使用 .NET 的 SHA1 类
    Set objSHA1 = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    arrBytes = objSHA1.ComputeHash_2(StringToByteArray(sInput))
    
    sHash = ""
    For i = 0 To UBound(arrBytes)
        sHash = sHash & Right("0" & Hex(arrBytes(i)), 2)
    Next
    
    SHA1Hash = LCase(sHash)
End Function

// 字符串转字节数组
Function StringToByteArray(sInput)
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2  // adTypeText
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.WriteText sInput
    objStream.Position = 0
    objStream.Type = 1  // adTypeBinary
    objStream.Position = 3  // 跳过 BOM
    StringToByteArray = objStream.Read
    objStream.Close
End Function
```

### 示例15：调用通用文字识别 API

```vb
// 调用来也 IDP 通用文字识别 API
Dim sPubkey = "your_pubkey_here"
Dim sSecretKey = "your_secret_key_here"
Dim sImagePath = "C:\test_image.jpg"
Dim sApiUrl = "https://cloud.laiye.com/idp/v1/mage/ocr/general"

TracePrint("=== 开始调用通用文字识别 API ===")

Try
    // 读取图片并转换为 Base64
    Dim objFSO, objFile, arrBytes, sBase64
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(sImagePath, 1)
    
    // 读取文件为二进制
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1  // adTypeBinary
    objStream.Open
    objStream.LoadFromFile sImagePath
    arrBytes = objStream.Read
    sBase64 = Base64Encode(arrBytes)
    objStream.Close
    
    // 生成请求头
    Dim objHeader
    Set objHeader = GenerateIDPHeader(sPubkey, sSecretKey)
    
    // 构建请求体
    Dim sRequestBody
    sRequestBody = "{""image"":""" & sBase64 & """}"
    
    // 发送 HTTP POST 请求
    Dim objHTTP, sResponse
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    objHTTP.Open "POST", sApiUrl, False
    
    // 设置请求头
    Dim sKey
    For Each sKey In objHeader.Keys
        objHTTP.setRequestHeader sKey, objHeader(sKey)
    Next
    
    // 发送请求
    objHTTP.Send sRequestBody
    
    // 获取响应
    sResponse = objHTTP.responseText
    TracePrint("API 响应: " & sResponse)
    
    // 解析 JSON 响应
    Dim objJSON, sText
    Set objJSON = ParseJSON(sResponse)
    
    If objJSON("code") = 0 Then
        sText = objJSON("data")("text")
        TracePrint("识别结果: " & sText)
        
        // 保存结果到文件
        Dim objOutFile
        Set objOutFile = objFSO.CreateTextFile("C:\ocr_result.txt", True)
        objOutFile.Write sText
        objOutFile.Close
        
        TracePrint("识别成功，结果已保存")
    Else
        TracePrint("识别失败: " & objJSON("message"))
    End If
    
    TracePrint("=== API 调用完成 ===")
    
Catch ex
    TracePrint("调用失败: " & ex.Message)
End Try
```

### 示例16：批量调用票据识别 API

```vb
// 批量调用来也 IDP 通用多票据识别 API
Dim sPubkey = "your_pubkey_here"
Dim sSecretKey = "your_secret_key_here"
Dim sFolderPath = "C:\Invoices"
Dim sApiUrl = "https://cloud.laiye.com/idp/v1/mage/ocr/bills"

TracePrint("=== 开始批量票据识别 ===")

// 获取文件列表
Dim arrFiles
arrFiles = File.GetFileList(sFolderPath, "*.jpg|*.png", False)
TracePrint("待识别文件数: " & Array.Length(arrFiles))

// 创建 Excel 汇总表
Dim objExcel
objExcel = Excel.Create(True, "")
Excel.SetCell(objExcel, "A1", "序号")
Excel.SetCell(objExcel, "B1", "文件名")
Excel.SetCell(objExcel, "C1", "票据类型")
Excel.SetCell(objExcel, "D1", "发票号码")
Excel.SetCell(objExcel, "E1", "金额")
Excel.SetCell(objExcel, "F1", "识别状态")

Dim iRow = 2
Dim iSuccess = 0
Dim iFailed = 0

For Each sFile In arrFiles
    TracePrint("处理文件: " & File.GetName(sFile))
    
    Try
        // 读取图片并转换为 Base64
        Dim objStream, sBase64
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1
        objStream.Open
        objStream.LoadFromFile sFile
        sBase64 = Base64Encode(objStream.Read)
        objStream.Close
        
        // 生成请求头
        Dim objHeader
        Set objHeader = GenerateIDPHeader(sPubkey, sSecretKey)
        
        // 构建请求体
        Dim sRequestBody
        sRequestBody = "{""image"":""" & sBase64 & """}"
        
        // 发送请求
        Dim objHTTP, sResponse
        Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        objHTTP.Open "POST", sApiUrl, False
        
        Dim sKey
        For Each sKey In objHeader.Keys
            objHTTP.setRequestHeader sKey, objHeader(sKey)
        Next
        
        objHTTP.Send sRequestBody
        sResponse = objHTTP.responseText
        
        // 解析响应
        Dim objJSON
        Set objJSON = ParseJSON(sResponse)
        
        If objJSON("code") = 0 Then
            Dim arrBills, objBill
            arrBills = objJSON("data")("bills")
            
            If Array.Length(arrBills) > 0 Then
                objBill = arrBills(0)
                
                // 写入 Excel
                Excel.SetCell(objExcel, "A" & iRow, iRow - 1)
                Excel.SetCell(objExcel, "B" & iRow, File.GetName(sFile))
                Excel.SetCell(objExcel, "C" & iRow, objBill("type"))
                Excel.SetCell(objExcel, "D" & iRow, objBill("invoice_number"))
                Excel.SetCell(objExcel, "E" & iRow, objBill("amount"))
                Excel.SetCell(objExcel, "F" & iRow, "成功")
                
                iSuccess = iSuccess + 1
                TracePrint("识别成功: " & objBill("type"))
            End If
        Else
            Excel.SetCell(objExcel, "A" & iRow, iRow - 1)
            Excel.SetCell(objExcel, "B" & iRow, File.GetName(sFile))
            Excel.SetCell(objExcel, "F" & iRow, "失败: " & objJSON("message"))
            
            iFailed = iFailed + 1
            TracePrint("识别失败")
        End If
        
        iRow = iRow + 1
        Delay(500)  // 避免请求过快
        
    Catch ex
        iFailed = iFailed + 1
        TracePrint("处理异常: " & ex.Message)
    End Try
Next

// 保存汇总表
Dim sResultFile
sResultFile = "C:\Invoices\batch_result_" & Time.Format(Time.Now(), "yyyyMMdd_HHmmss") & ".xlsx"
Excel.SaveAs(objExcel, sResultFile)
Excel.Close(objExcel)

TracePrint("=== 批量识别完成 ===")
TracePrint("总文件数: " & Array.Length(arrFiles))
TracePrint("成功: " & iSuccess & ", 失败: " & iFailed)
TracePrint("结果已保存到: " & sResultFile)
```

### API 限流说明

为了安全性和响应效率，来也 IDP 对 API 做了调用频率限流：

| AI能力 | URI | 限制规则 |
|--------|-----|----------|
| 通用文字识别 | /v1/mage/ocr/general | 企业版: 根据商务沟通确定<br>免费版: 每分钟6次 |
| 通用表格识别 | /v1/mage/ocr/table | 企业版: 根据商务沟通确定<br>免费版: 每分钟6次 |
| 通用卡证识别 | /v1/mage/ocr/license | 企业版: 根据商务沟通确定<br>免费版: 每分钟6次 |
| 通用多票据识别 | /v1/mage/ocr/bills | 企业版: 根据商务沟通确定<br>免费版: 每分钟6次 |
| 模板识别 | /v1/document/ocr/template | 企业版: 根据商务沟通确定<br>免费版: 每分钟6次 |

每次调用，在返回的 Response Headers 中会给出以下参数：
- `X-Ratelimit-Remaining`：当前时间窗口剩余请求
- `X-Ratelimit-Reset`：下次重置时间
- `X-Ratelimit-Limit`：当前时间窗口最大限流次数

### 常见错误码

| Code | 描述 |
|------|------|
| 0 | 正常 |
| 3 | 参数错误 |
| 8 | 资源耗尽（如请求体过大） |
| 10000 | 服务内部错误 |
| 10001 | Header解析错误 |
| 10002 | 签名验证失败 |
| 10003 | 参数不正确，应用不存在 |
| 10006 | 需要选择待识别的图片 |
| 10007 | 错误的文件类型 |
| 10008 | 格式不正确，只支持png,jpeg,jpg,bmp,tiff,pdf |
| 10009 | 文件尺寸不正确，文件的长宽需要在15和4096像素之间 |
| 10010 | 处理超时 |
| 10011 | 账号配额不足 |
| 10015 | 调用频率超限 |
| 10017 | 不支持加密的PDF |
| 10018 | 请求数据过大，请控制在10M以内 |

**参考文档**：[来也 IDP 接口文档](https://cloud.laiye.com/idp/docs/latest/docUnderstanding/backend/api.html)

---

**文档版本**: v1.1.0  
**适用版本**: UIBot 6.0.0.211215(64位)  
**更新时间**: 2025-05-02
