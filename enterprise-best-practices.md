# UIBot 企业级开发最佳实践

基于官方企业级流程模板的深度分析，本文档提供 UIBot 企业级 RPA 项目的最佳实践指南。

## 目录

- [企业级流程架构](#企业级流程架构)
- [核心设计模式](#核心设计模式)
- [配置管理](#配置管理)
- [异常处理策略](#异常处理策略)
- [重试机制](#重试机制)
- [日志管理](#日志管理)
- [数据队列处理](#数据队列处理)
- [企业级代码模板](#企业级代码模板)

---

## 企业级流程架构

### 标准流程结构

```
主流程 (main.prj)
├── 配置初始化 (配置初始化.task)
├── 目标程序初始化 (目标程序初始化.task)
├── 获取新数据 (获取新数据.task)
├── 执行流程 (子流程)
│   └── 流程块 (流程块.task)
├── 重试计数 (重试计数.task)
├── 异常处理 (异常处理.task)
└── 流程结束 (流程结束.task)
```

### 全局变量设计

```vb
' 核心全局变量
Dim g_bRetryStatus = False        ' 重试状态标识
Dim g_dicTransactionItem = Null   ' 当前事务项数据
Dim g_dicConfigData = {}          ' 全局配置数据
Dim g_iCount = 0                  ' 重试计数器
Dim g_iRetryNumber = 0            ' 最大重试次数
Dim g_bFirstRun = True            ' 首次运行标识
```

### 流程执行逻辑

```
开始
  ↓
配置初始化 → 读取Config.xlsx → 设置日志级别 → 初始化全局变量
  ↓
目标程序初始化 → 关闭残留进程 → 检查前置条件
  ↓
首次运行判断
  ├─ Yes → 获取新数据
  └─ No → 重试判断
           ├─ Yes → 执行流程（使用旧数据）
           └─ No → 获取新数据
  ↓
有数据判断
  ├─ Yes → 执行流程（子流程）
  │         ↓
  │       成功 → 获取新数据（循环）
  │         ↓
  │       失败 → 重试计数 → 目标程序初始化（重试）
  └─ No → 流程结束
  ↓
异常处理 → 记录日志 → 截图 → 通知
  ↓
流程结束 → 清理资源 → 上报统计
```

---

## 核心设计模式

### 1. REFramework 模式（Robotic Enterprise Framework）

UIBot 企业级模板采用类似 UiPath REFramework 的设计模式：

**四大核心阶段**：
1. **初始化阶段** (Initialization)
   - 配置初始化
   - 目标程序初始化
   - 环境检查

2. **获取事务阶段** (Get Transaction Data)
   - 从队列/数据库/Excel获取待处理数据
   - 判断是否有新数据

3. **处理事务阶段** (Process Transaction)
   - 执行具体业务逻辑
   - 子流程化处理

4. **结束阶段** (End Process)
   - 清理资源
   - 统计上报

### 2. 事务项模式

```vb
' 事务项数据结构
g_dicTransactionItem = {
    "TransactionID": "TXN20240115001",
    "UserName": "张三",
    "OrderNo": "ORD123456",
    "Amount": 1000,
    "Status": "Pending"
}
```

### 3. 配置驱动模式

所有配置集中在 `Config.xlsx` 中管理：

**工作表结构**：
- **全局设置**：日志级别、环境配置
- **常量设置**：重试次数、超时时间、路径配置
- **本地参数**：应用程序名称、特定参数

---

## 配置管理

### Config.xlsx 标准结构

#### 工作表1：全局设置
| 参数名 | 参数值 | 说明 |
|--------|--------|------|
| LogLevel | 2 | 日志级别：0=关闭,1=Error,2=Info,3=Debug |
| Environment | Production | 环境：Development/Test/Production |

#### 工作表2：常量设置
| 参数名 | 参数值 | 说明 |
|--------|--------|------|
| MaxRetryNumber | 2 | 最大重试次数 |
| ExScreenshotsFolderPath | \Screenshots | 异常截图保存路径 |
| DefaultTimeout | 30 | 默认超时时间（秒） |

#### 工作表3：本地参数
| 参数名 | 参数值 | 说明 |
|--------|--------|------|
| appsName | chrome.exe,notepad.exe | 需要关闭的应用进程名 |
| TargetURL | https://example.com | 目标网址 |

### 配置初始化完整代码

```vb
' 配置初始化标准模板
Dim config = {}
Dim objExcelWorkBook = ""
Dim sheetsName = ["全局设置", "常量设置", "本地参数"]

' 打开配置文件
objExcelWorkBook = Excel.OpenExcel(@res"Data\\Config.xlsx", True, "Excel", "", "")

' 读取所有工作表
For Each sheetName In sheetsName
    Dim iRowsCount = Excel.GetRowsCount(objExcelWorkBook, sheetName)
    If iRowsCount > 1
        Dim arrayRet = Excel.ReadRange(objExcelWorkBook, sheetName, "A2:B" & iRowsCount)
        config[sheetName] = {}
        For Each row In arrayRet
            If Trim(row[0], "") <> ""
                config[sheetName][Trim(row[0], "")] = Trim(row[1], "")
            End If
        Next
    Else
        Log.Warn("配置文件的[" & sheetName & "]工作表无配置信息！")
    End If
Next

Excel.CloseExcel(objExcelWorkBook, True)

' 设置日志级别
Dim logLevel = 2
If Not IsNull(config["全局设置"]["LogLevel"])
    logLevel = CInt(config["全局设置"]["LogLevel"])
End If
Log.SetLevel(logLevel)

' 初始化最大重试次数
If Not IsNull(config["常量设置"]["MaxRetryNumber"])
    g_iRetryNumber = CInt(config["常量设置"]["MaxRetryNumber"])
End If

' 初始化截图路径
Dim HomePath = Sys.GetHomePath()
Dim ExScreenshotsFolderPath = config["常量设置"]["ExScreenshotsFolderPath"]
If Not IsNull(ExScreenshotsFolderPath)
    ExScreenshotsFolderPath = Trim(ExScreenshotsFolderPath, " ")
    If Left(ExScreenshotsFolderPath, 1) = "\\"
        ExScreenshotsFolderPath = HomePath & ExScreenshotsFolderPath
        config["常量设置"]["ExScreenshotsFolderPath"] = ExScreenshotsFolderPath
    End If
    If Not File.FolderExists(ExScreenshotsFolderPath)
        File.CreateFolder(ExScreenshotsFolderPath)
    End If
End If

' 赋值给全局变量
g_dicConfigData = config
```

---

## 异常处理策略

### 三层异常处理架构

```vb
' 第一层：业务异常（Try-Catch捕获）
Try
    ' 业务代码
    Mouse.Click(@ui"按钮")
Catch Exception
    Log.Error("业务异常: " & Exception.Message)
    ' 业务异常处理，不会触发流程块异常虚线
End Try

' 第二层：流程块异常（异常虚线）
' 流程块外的代码发生异常，会触发异常虚线指向异常处理流程块

' 第三层：全局异常处理
' 在异常处理流程块中统一处理
```

### 异常处理标准模板

```vb
' 异常处理流程块
Log.Error("流程发生异常，正在处理！")

' 获取系统异常
Dim SystemException = $BlockInput
Log.Error("异常信息: " & SystemException.Message)
Log.Error("异常位置: " & SystemException.Source)

' 截取异常现场
Dim dTime = Time.Now()
Dim sTime = Time.Format(dTime, "yyyyMMdd_HHmmss")
Dim screenshotPath = g_dicConfigData["常量设置"]["ExScreenshotsFolderPath"]
Dim screenshotFile = screenshotPath & "\\Exception_" & sTime & ".png"

Try
    Screen.Capture(screenshotFile, 0, 0, 0, 0)
    Log.Info("已保存异常截图: " & screenshotFile)
Catch ex
    Log.Error("截图失败: " & ex.Message)
End Try

' 发送异常通知邮件
Try
    Dim sSubject = "RPA流程异常通知 - " & Time.Format(dTime, "yyyy-MM-dd HH:mm:ss")
    Dim sBody = "流程名称: " & g_dicConfigData["全局设置"]["ProcessName"] & "\n"
    sBody = sBody & "异常时间: " & Time.Format(dTime, "yyyy-MM-dd HH:mm:ss") & "\n"
    sBody = sBody & "异常信息: " & SystemException.Message & "\n"
    sBody = sBody & "异常位置: " & SystemException.Source & "\n"
    
    Mail.Send(
        g_dicConfigData["邮件设置"]["SMTPServer"],
        CInt(g_dicConfigData["邮件设置"]["Port"]),
        g_dicConfigData["邮件设置"]["Username"],
        g_dicConfigData["邮件设置"]["Password"],
        g_dicConfigData["邮件设置"]["From"],
        g_dicConfigData["邮件设置"]["To"],
        sSubject,
        sBody,
        [screenshotFile]
    )
    Log.Info("异常通知邮件已发送")
Catch mailEx
    Log.Error("邮件发送失败: " & mailEx.Message)
End Try

' 企业版：上传截图到Commander
' Upload.UploadScreenShot("Exception_" & sTime, screenshotPath, True, {"sDescribe": "流程异常截图"})

Log.Info("异常处理完毕！")
```

---

## 重试机制

### 智能重试策略

```vb
' 重试计数流程块
Dim sRet = ""
Dim dTime = Time.Now()
Dim sTime = Time.Format(dTime, "yyyyMMddHHmmss")
Dim shotsPath = g_dicConfigData["常量设置"]["ExScreenshotsFolderPath"]

If g_iCount < g_iRetryNumber
    ' 未超过最大重试次数
    g_bRetryStatus = True
    g_iCount = g_iCount + 1
    Log.Info("准备第 " & g_iCount & " 次重试，最大重试次数: " & g_iRetryNumber)
Else
    ' 超过最大重试次数
    g_iCount = 0
    g_bRetryStatus = False
    Log.Error("已达最大重试次数，事务项处理失败")
    
    ' 最后一次重试失败，截图保存
    If (shotsPath <> "") And (Not IsNull(shotsPath))
        Dim screenshotFile = shotsPath & "\\Retry_Failed_" & sTime & ".png"
        Screen.Capture(screenshotFile, 0, 0, 0, 0)
        Log.Info("最后一次重试失败截图: " & screenshotFile)
        
        ' 企业版：上传截图
        ' Upload.UploadScreenShot("Retry_Failed_" & sTime, shotsPath, True, {"sDescribe": "最后一次重试失败截图"})
    End If
End If
```

### 重试场景最佳实践

```vb
' 场景1：网络请求重试
Function HttpGetWithRetry(sUrl, iMaxRetry)
    Dim iRetry = 0
    Dim sResult = ""
    
    Do While iRetry < iMaxRetry
        Try
            HTTP.Get(sUrl, {}, sResult)
            Return sResult
        Catch ex
            iRetry = iRetry + 1
            Log.Warn("HTTP请求失败，重试 " & iRetry & "/" & iMaxRetry)
            If iRetry < iMaxRetry Then
                Delay(2000)  ' 等待2秒后重试
            End If
        End Try
    Loop
    
    Throw "HTTP请求失败，已达最大重试次数"
End Function

' 场景2：元素操作重试
Function ClickWithRetry(objElement, iMaxRetry)
    Dim iRetry = 0
    
    Do While iRetry < iMaxRetry
        Try
            If UiElement.Exists(objElement, 5) Then
                Mouse.Click(objElement, "left", "single", 0, 0)
                Return True
            End If
        Catch ex
            iRetry = iRetry + 1
            Log.Warn("点击失败，重试 " & iRetry & "/" & iMaxRetry)
            Delay(1000)
        End Try
    Loop
    
    Return False
End Function
```

---

## 日志管理

### 日志级别定义

```vb
' 日志级别
' 0 = 关闭日志
' 1 = Error（仅错误）
' 2 = Info（信息 + 错误）
' 3 = Debug（调试 + 信息 + 错误）

Log.SetLevel(2)
```

### 企业级日志实践

```vb
' 日志记录函数
Function WriteDetailLog(sLevel, sModule, sMessage, dicData)
    Dim sTime = Time.Format(Time.Now(), "yyyy-MM-dd HH:mm:ss.fff")
    Dim sLogLine = "[" & sTime & "] [" & sLevel & "] [" & sModule & "] " & sMessage
    
    ' 添加数据信息
    If Not IsNull(dicData) Then
        sLogLine = sLogLine & " | Data: " & JSON.Stringify(dicData)
    End If
    
    ' 输出到控制台
    Select Case sLevel
        Case "ERROR"
            Log.Error(sLogLine)
        Case "WARN"
            Log.Warn(sLogLine)
        Case "INFO"
            Log.Info(sLogLine)
        Case "DEBUG"
            Log.Debug(sLogLine)
    End Select
    
    ' 写入日志文件
    Dim sLogFile = g_dicConfigData["常量设置"]["LogPath"] & "\\RPA_" & Time.Format(Time.Now(), "yyyyMMdd") & ".log"
    File.Append(sLogFile, sLogLine & "\n", "utf-8")
End Function

' 使用示例
WriteDetailLog("INFO", "配置初始化", "配置加载成功", g_dicConfigData)
WriteDetailLog("ERROR", "数据处理", "数据验证失败", {"OrderNo": "ORD123", "Error": "金额为空"})
```

### 关键节点日志

```vb
' 流程开始
Log.Info("========== 流程开始 ==========")
Log.Info("流程名称: " & g_dicConfigData["全局设置"]["ProcessName"])
Log.Info("执行时间: " & Time.Format(Time.Now(), "yyyy-MM-dd HH:mm:ss"))

' 配置初始化
Log.Info("配置初始化开始")
Log.Info("日志级别: " & g_dicConfigData["全局设置"]["LogLevel"])
Log.Info("最大重试次数: " & g_iRetryNumber)

' 事务处理
Log.Info("开始处理事务项: " & g_dicTransactionItem["TransactionID"])

' 流程结束
Log.Info("流程执行时长: " & $Flow.ElaspedTime & " 毫秒")
Log.Info("========== 流程结束 ==========")
```

---

## 数据队列处理

### 获取新数据标准模板

```vb
' 获取新数据流程块
g_bFirstRun = False  ' 更新首次运行标识
g_dicTransactionItem = Null  ' 初始化事务项

' 方式1：从Commander数据队列获取（企业版）
Try
    Dim queueItem = Queue.GetQueueItem("订单处理队列")
    If Not IsNull(queueItem) Then
        g_dicTransactionItem = {
            "QueueItemID": queueItem["id"],
            "OrderNo": queueItem["data"]["OrderNo"],
            "CustomerName": queueItem["data"]["CustomerName"],
            "Amount": queueItem["data"]["Amount"]
        }
        Log.Info("从队列获取数据: " & g_dicTransactionItem["OrderNo"])
    End If
Catch ex
    Log.Error("队列获取失败: " & ex.Message)
End Try

' 方式2：从Excel获取
If IsNull(g_dicTransactionItem) Then
    Try
        Dim objExcel = Excel.Open(g_dicConfigData["本地参数"]["DataFilePath"], True, "")
        Dim iRow = g_dicConfigData["运行时"]["CurrentRow"]
        
        Dim sOrderNo = Excel.GetCell(objExcel, "A" & iRow)
        If sOrderNo <> "" Then
            g_dicTransactionItem = {
                "RowNumber": iRow,
                "OrderNo": sOrderNo,
                "CustomerName": Excel.GetCell(objExcel, "B" & iRow),
                "Amount": Excel.GetCell(objExcel, "C" & iRow)
            }
            g_dicConfigData["运行时"]["CurrentRow"] = iRow + 1
            Log.Info("从Excel获取数据: 第" & iRow & "行")
        End If
        
        Excel.Close(objExcel)
    Catch ex
        Log.Error("Excel读取失败: " & ex.Message)
    End Try
End If

' 方式3：从数据库获取
If IsNull(g_dicTransactionItem) Then
    Try
        Dim objDB = DB.Create("sqlserver", g_dicConfigData["数据库"]["ConnectionString"])
        Dim sSQL = "SELECT TOP 1 * FROM Orders WHERE Status='Pending' ORDER BY CreateTime"
        Dim objResult = DB.QueryOne(objDB, sSQL, [])
        
        If Not IsNull(objResult) Then
            g_dicTransactionItem = {
                "OrderID": objResult["OrderID"],
                "OrderNo": objResult["OrderNo"],
                "CustomerName": objResult["CustomerName"],
                "Amount": objResult["Amount"]
            }
            Log.Info("从数据库获取数据: " & g_dicTransactionItem["OrderNo"])
        End If
        
        DB.Close(objDB)
    Catch ex
        Log.Error("数据库查询失败: " & ex.Message)
    End Try
End If

' 方式4：从WebAPI获取
If IsNull(g_dicTransactionItem) Then
    Try
        Dim sApiUrl = g_dicConfigData["API"]["BaseURL"] & "/api/orders/next"
        Dim sResult = ""
        HTTP.Get(sApiUrl, {"Authorization": "Bearer " & g_dicConfigData["API"]["Token"]}, sResult)
        
        Dim objData = JSON.Parse(sResult)
        If objData["success"] Then
            g_dicTransactionItem = objData["data"]
            Log.Info("从API获取数据: " & g_dicTransactionItem["OrderNo"])
        End If
    Catch ex
        Log.Error("API请求失败: " & ex.Message)
    End Try
End If
```

---

## 企业级代码模板

### 完整企业级流程模板

```vb
' ========== 主流程 ==========

' 1. 配置初始化
Call InitializeConfig()

' 2. 目标程序初始化
Call InitializeApplication()

' 3. 主循环
Do While True
    ' 获取新数据
    If g_bFirstRun Or Not g_bRetryStatus Then
        Call GetTransactionData()
    End If
    
    ' 判断是否有数据
    If IsNull(g_dicTransactionItem) Then
        Log.Info("没有待处理数据，流程结束")
        Exit Do
    End If
    
    ' 处理事务
    Try
        Call ProcessTransaction()
        Log.Info("事务处理成功: " & g_dicTransactionItem["TransactionID"])
        g_iCount = 0  ' 重置重试计数
    Catch ex
        Log.Error("事务处理失败: " & ex.Message)
        Call HandleException(ex)
        
        ' 重试判断
        If g_iCount < g_iRetryNumber Then
            g_bRetryStatus = True
            g_iCount = g_iCount + 1
            Log.Info("准备重试，当前次数: " & g_iCount)
        Else
            Log.Error("超过最大重试次数，跳过此事务")
            g_iCount = 0
            g_bRetryStatus = False
        End If
    End Try
Loop

' 4. 流程结束
Call FinalizeProcess()

' ========== 函数定义 ==========

' 配置初始化函数
Function InitializeConfig()
    ' 实现配置初始化逻辑
End Function

' 目标程序初始化函数
Function InitializeApplication()
    ' 实现应用初始化逻辑
End Function

' 获取事务数据函数
Function GetTransactionData()
    ' 实现数据获取逻辑
End Function

' 处理事务函数
Function ProcessTransaction()
    ' 实现业务处理逻辑
End Function

' 异常处理函数
Function HandleException(ex)
    ' 实现异常处理逻辑
End Function

' 流程结束函数
Function FinalizeProcess()
    ' 实现清理逻辑
End Function
```

---

## 最佳实践总结

### 1. 架构设计
- ✅ 采用 REFramework 模式
- ✅ 配置与代码分离
- ✅ 主流程与子流程分离
- ✅ 全局变量统一管理

### 2. 异常处理
- ✅ 三层异常处理机制
- ✅ 异常截图保存
- ✅ 异常通知机制
- ✅ 详细的异常日志

### 3. 重试机制
- ✅ 可配置的重试次数
- ✅ 智能重试判断
- ✅ 重试状态管理
- ✅ 失败截图保存

### 4. 日志管理
- ✅ 分级日志记录
- ✅ 关键节点日志
- ✅ 日志文件持久化
- ✅ 日志级别可配置

### 5. 数据处理
- ✅ 多数据源支持
- ✅ 事务项模式
- ✅ 数据验证
- ✅ 状态更新

---

**文档版本**: v1.0.0  
**更新时间**: 2024-01-15  
**基于**: UIBot 6.0 企业级流程模板
