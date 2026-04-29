# UIBot 流程设计模式

本文档基于企业级流程模板，总结 UIBot RPA 开发中的常用设计模式和架构方案。

## 目录

- [REFramework 模式](#reframework-模式)
- [状态机模式](#状态机模式)
- [线性流程模式](#线性流程模式)
- [事务处理模式](#事务处理模式)
- [配置驱动模式](#配置驱动模式)
- [模块化设计](#模块化设计)

---

## REFramework 模式

### 概述

REFramework（Robotic Enterprise Framework）是企业级 RPA 的标准架构模式，适用于需要处理大量事务的场景。

### 架构图

```
┌─────────────────────────────────────────────────────────┐
│                    REFramework 流程                      │
├─────────────────────────────────────────────────────────┤
│                                                          │
│  ┌──────────────┐                                       │
│  │  初始化阶段   │                                       │
│  │ Initialization│                                       │
│  └──────┬───────┘                                       │
│         │                                                │
│         ▼                                                │
│  ┌──────────────┐      ┌──────────────┐               │
│  │ 获取事务数据  │◄─────┤  处理成功     │               │
│  │Get Transaction│      │              │               │
│  └──────┬───────┘      └──────▲───────┘               │
│         │                     │                         │
│         │ 有数据              │                         │
│         ▼                     │                         │
│  ┌──────────────┐            │                         │
│  │  处理事务     │────────────┘                         │
│  │Process Trans. │                                      │
│  └──────┬───────┘                                       │
│         │                                                │
│         │ 异常                                           │
│         ▼                                                │
│  ┌──────────────┐      ┌──────────────┐               │
│  │  异常处理     │──────►  重试判断     │               │
│  │Exception      │      │              │               │
│  └──────────────┘      └──────┬───────┘               │
│                               │                         │
│                               │ 重试                    │
│                               └─────────────────┐       │
│                                                 │       │
│         │ 无数据                                │       │
│         ▼                                       ▼       │
│  ┌──────────────┐                      ┌──────────────┐│
│  │  流程结束     │                      │目标程序初始化 ││
│  │End Process    │                      │              ││
│  └──────────────┘                      └──────────────┘│
│                                                          │
└─────────────────────────────────────────────────────────┘
```

### 核心组件

#### 1. 初始化阶段

```vb
' 配置初始化
Function InitConfig()
    ' 读取配置文件
    g_dicConfigData = LoadConfigFromExcel(@res"Data\\Config.xlsx")
    
    ' 设置日志级别
    Log.SetLevel(g_dicConfigData["全局设置"]["LogLevel"])
    
    ' 初始化全局变量
    g_iCount = 0
    g_iRetryNumber = g_dicConfigData["常量设置"]["MaxRetryNumber"]
    g_bRetryStatus = False
    g_dicTransactionItem = Null
    g_bFirstRun = True
    
    Log.Info("配置初始化完成")
End Function

' 应用初始化
Function InitApplication()
    ' 关闭残留进程
    Dim appArr = Split(g_dicConfigData["本地参数"]["appsName"], ",")
    For Each appName In appArr
        If App.GetStatus(appName) Then
            App.Kill(appName)
            Log.Info("已关闭进程: " & appName)
        End If
    Next
    
    ' 启动目标应用
    ' App.Run(...)
    
    Log.Info("应用初始化完成")
End Function
```

#### 2. 获取事务数据

```vb
Function GetTransactionData()
    g_bFirstRun = False
    g_dicTransactionItem = Null
    
    ' 从数据源获取下一条待处理数据
    ' 可以是：队列、数据库、Excel、API等
    
    ' 示例：从队列获取
    Try
        Dim queueItem = Queue.GetQueueItem("订单队列")
        If Not IsNull(queueItem) Then
            g_dicTransactionItem = queueItem["data"]
            Log.Info("获取事务: " & g_dicTransactionItem["ID"])
        Else
            Log.Info("队列为空，无待处理数据")
        End If
    Catch ex
        Log.Error("获取事务失败: " & ex.Message)
        Throw ex
    End Try
End Function
```

#### 3. 处理事务

```vb
Function ProcessTransaction()
    Log.Info("开始处理事务: " & g_dicTransactionItem["ID"])
    
    Try
        ' 执行具体业务逻辑
        ' 1. 打开应用/网页
        ' 2. 输入数据
        ' 3. 执行操作
        ' 4. 获取结果
        ' 5. 更新状态
        
        ' 示例业务逻辑
        Dim objBrowser = WebBrowser.Create("chrome", g_dicTransactionItem["URL"], 30)
        ' ... 业务操作 ...
        WebBrowser.Close(objBrowser)
        
        ' 更新事务状态为成功
        UpdateTransactionStatus(g_dicTransactionItem["ID"], "Success")
        
        Log.Info("事务处理成功")
        Return True
        
    Catch ex
        Log.Error("事务处理失败: " & ex.Message)
        Throw ex
    End Try
End Function
```

#### 4. 异常处理与重试

```vb
Function HandleException(ex)
    Log.Error("异常信息: " & ex.Message)
    
    ' 截图保存
    Dim screenshotPath = SaveExceptionScreenshot()
    
    ' 判断是否重试
    If g_iCount < g_iRetryNumber Then
        g_bRetryStatus = True
        g_iCount = g_iCount + 1
        Log.Info("准备第 " & g_iCount & " 次重试")
        Return "Retry"
    Else
        ' 超过重试次数，标记为失败
        UpdateTransactionStatus(g_dicTransactionItem["ID"], "Failed")
        g_iCount = 0
        g_bRetryStatus = False
        Log.Error("超过最大重试次数，事务失败")
        Return "Failed"
    End If
End Function
```

### 适用场景

- ✅ 批量数据处理（订单、发票、报表等）
- ✅ 需要重试机制的流程
- ✅ 需要异常恢复的流程
- ✅ 企业级生产环境

---

## 状态机模式

### 概述

状态机模式将流程分解为多个状态，每个状态执行特定操作，根据条件转换到下一个状态。

### 架构图

```
     ┌─────────┐
     │  开始    │
     └────┬────┘
          │
          ▼
     ┌─────────┐
     │ 状态1    │
     │ 登录     │
     └────┬────┘
          │
          ▼
     ┌─────────┐
     │ 状态2    │
     │ 查询     │
     └────┬────┘
          │
          ▼
     ┌─────────┐
     │ 状态3    │
     │ 处理     │
     └────┬────┘
          │
          ▼
     ┌─────────┐
     │ 状态4    │
     │ 保存     │
     └────┬────┘
          │
          ▼
     ┌─────────┐
     │  结束    │
     └─────────┘
```

### 实现代码

```vb
' 状态机模式实现
Dim sCurrentState = "Init"
Dim bContinue = True

Do While bContinue
    Select Case sCurrentState
        Case "Init"
            Log.Info("状态: 初始化")
            If InitializeSystem() Then
                sCurrentState = "Login"
            Else
                sCurrentState = "Error"
            End If
            
        Case "Login"
            Log.Info("状态: 登录")
            If LoginToSystem() Then
                sCurrentState = "Query"
            Else
                sCurrentState = "Error"
            End If
            
        Case "Query"
            Log.Info("状态: 查询数据")
            If QueryData() Then
                sCurrentState = "Process"
            Else
                sCurrentState = "End"
            End If
            
        Case "Process"
            Log.Info("状态: 处理数据")
            If ProcessData() Then
                sCurrentState = "Save"
            Else
                sCurrentState = "Error"
            End If
            
        Case "Save"
            Log.Info("状态: 保存结果")
            If SaveResult() Then
                sCurrentState = "Query"  ' 继续查询下一条
            Else
                sCurrentState = "Error"
            End If
            
        Case "Error"
            Log.Error("状态: 错误处理")
            HandleError()
            sCurrentState = "End"
            
        Case "End"
            Log.Info("状态: 结束")
            bContinue = False
            
        Case Else
            Log.Error("未知状态: " & sCurrentState)
            bContinue = False
    End Select
Loop
```

### 适用场景

- ✅ 流程有明确的状态转换
- ✅ 需要根据条件跳转到不同状态
- ✅ 复杂的业务流程控制

---

## 线性流程模式

### 概述

线性流程模式按顺序执行一系列步骤，适用于简单的自动化任务。

### 架构图

```
开始 → 步骤1 → 步骤2 → 步骤3 → 步骤4 → 结束
```

### 实现代码

```vb
' 线性流程模式
Log.Info("========== 流程开始 ==========")

Try
    ' 步骤1：打开应用
    Log.Info("步骤1: 打开应用")
    Dim objBrowser = WebBrowser.Create("chrome", "https://example.com", 30)
    Delay(2000)
    
    ' 步骤2：登录
    Log.Info("步骤2: 登录系统")
    Keyboard.InputText(@ui"用户名", "admin", True, False)
    Keyboard.InputPwd(@ui"密码", "password", True)
    Mouse.Click(@ui"登录按钮")
    Delay(3000)
    
    ' 步骤3：执行操作
    Log.Info("步骤3: 执行操作")
    Mouse.Click(@ui"菜单")
    Delay(1000)
    Mouse.Click(@ui"功能按钮")
    Delay(2000)
    
    ' 步骤4：获取结果
    Log.Info("步骤4: 获取结果")
    Dim sResult = UiElement.GetText(@ui"结果文本")
    Log.Info("结果: " & sResult)
    
    ' 步骤5：保存数据
    Log.Info("步骤5: 保存数据")
    File.Write("C:\result.txt", sResult, "utf-8")
    
    ' 步骤6：关闭应用
    Log.Info("步骤6: 关闭应用")
    WebBrowser.Close(objBrowser)
    
    Log.Info("========== 流程成功完成 ==========")
    
Catch ex
    Log.Error("流程执行失败: " & ex.Message)
    ' 清理资源
    If objBrowser <> Null Then
        WebBrowser.Close(objBrowser)
    End If
End Try
```

### 适用场景

- ✅ 简单的自动化任务
- ✅ 步骤固定，无复杂分支
- ✅ 一次性执行的流程

---

## 事务处理模式

### 概述

事务处理模式将数据处理分为独立的事务单元，每个事务独立处理，失败不影响其他事务。

### 架构图

```
┌─────────────────────────────────────┐
│         事务处理循环                 │
├─────────────────────────────────────┤
│                                      │
│  ┌──────────────────────────────┐  │
│  │  获取事务列表                 │  │
│  └────────────┬─────────────────┘  │
│               │                     │
│               ▼                     │
│  ┌──────────────────────────────┐  │
│  │  For Each 事务 In 事务列表    │  │
│  └────────────┬─────────────────┘  │
│               │                     │
│               ▼                     │
│  ┌──────────────────────────────┐  │
│  │  Try                          │  │
│  │    处理单个事务                │  │
│  │    标记为成功                  │  │
│  │  Catch                        │  │
│  │    记录错误                    │  │
│  │    标记为失败                  │  │
│  │  End Try                      │  │
│  └────────────┬─────────────────┘  │
│               │                     │
│               ▼                     │
│  ┌──────────────────────────────┐  │
│  │  Next                         │  │
│  └──────────────────────────────┘  │
│                                      │
└─────────────────────────────────────┘
```

### 实现代码

```vb
' 事务处理模式
Function ProcessTransactions()
    ' 获取所有待处理事务
    Dim arrTransactions = GetAllTransactions()
    Dim iSuccess = 0
    Dim iFailed = 0
    
    Log.Info("共有 " & Array.Length(arrTransactions) & " 个事务待处理")
    
    ' 逐个处理事务
    For Each transaction In arrTransactions
        Try
            Log.Info("处理事务: " & transaction["ID"])
            
            ' 处理单个事务
            ProcessSingleTransaction(transaction)
            
            ' 标记为成功
            UpdateStatus(transaction["ID"], "Success")
            iSuccess = iSuccess + 1
            
            Log.Info("事务处理成功: " & transaction["ID"])
            
        Catch ex
            ' 记录错误
            Log.Error("事务处理失败: " & transaction["ID"] & " - " & ex.Message)
            
            ' 标记为失败
            UpdateStatus(transaction["ID"], "Failed", ex.Message)
            iFailed = iFailed + 1
            
            ' 继续处理下一个事务
            Continue
        End Try
    Next
    
    ' 输出统计
    Log.Info("处理完成 - 成功: " & iSuccess & ", 失败: " & iFailed)
End Function

' 处理单个事务
Function ProcessSingleTransaction(transaction)
    ' 具体业务逻辑
    ' ...
End Function
```

### 适用场景

- ✅ 批量数据处理
- ✅ 每条数据独立处理
- ✅ 需要统计成功/失败数量

---

## 配置驱动模式

### 概述

配置驱动模式将流程参数、业务规则等配置化，通过修改配置文件即可调整流程行为，无需修改代码。

### 配置文件结构

```
Config.xlsx
├── 全局设置
│   ├── LogLevel = 2
│   ├── Environment = Production
│   └── ProcessName = 订单处理流程
├── 常量设置
│   ├── MaxRetryNumber = 3
│   ├── DefaultTimeout = 30
│   └── ExScreenshotsFolderPath = \Screenshots
├── 本地参数
│   ├── appsName = chrome.exe,excel.exe
│   ├── TargetURL = https://example.com
│   └── DataFilePath = C:\Data\orders.xlsx
└── 业务规则
    ├── MinAmount = 100
    ├── MaxAmount = 10000
    └── ApprovalRequired = True
```

### 实现代码

```vb
' 配置驱动模式
Function LoadConfig()
    Dim config = {}
    Dim objExcel = Excel.OpenExcel(@res"Data\\Config.xlsx", True, "Excel", "", "")
    
    ' 读取所有配置表
    Dim sheets = ["全局设置", "常量设置", "本地参数", "业务规则"]
    For Each sheet In sheets
        config[sheet] = {}
        Dim iRows = Excel.GetRowsCount(objExcel, sheet)
        Dim arrData = Excel.ReadRange(objExcel, sheet, "A2:B" & iRows)
        
        For Each row In arrData
            If row[0] <> "" Then
                config[sheet][row[0]] = row[1]
            End If
        Next
    Next
    
    Excel.CloseExcel(objExcel, True)
    Return config
End Function

' 使用配置
Dim g_config = LoadConfig()

' 根据配置执行不同逻辑
If g_config["业务规则"]["ApprovalRequired"] = "True" Then
    ' 需要审批
    SendForApproval()
Else
    ' 直接处理
    ProcessDirectly()
End If

' 使用配置的超时时间
Dim iTimeout = CInt(g_config["常量设置"]["DefaultTimeout"])
WebBrowser.Create("chrome", g_config["本地参数"]["TargetURL"], iTimeout)
```

### 适用场景

- ✅ 需要频繁调整参数
- ✅ 多环境部署（开发/测试/生产）
- ✅ 业务规则经常变化

---

## 模块化设计

### 概述

模块化设计将流程拆分为独立的功能模块（子流程），提高代码复用性和可维护性。

### 架构图

```
主流程
├── 模块1: 登录模块 (子流程)
├── 模块2: 数据查询模块 (子流程)
├── 模块3: 数据处理模块 (子流程)
├── 模块4: 结果保存模块 (子流程)
└── 模块5: 通知模块 (子流程)
```

### 实现代码

```vb
' 主流程
Log.Info("========== 主流程开始 ==========")

' 调用登录模块
Dim loginResult = SubFlow.Run("登录模块.flow", {
    "username": "admin",
    "password": "password"
})

If loginResult["success"] Then
    ' 调用数据查询模块
    Dim queryResult = SubFlow.Run("数据查询模块.flow", {
        "startDate": "2024-01-01",
        "endDate": "2024-01-31"
    })
    
    ' 调用数据处理模块
    Dim processResult = SubFlow.Run("数据处理模块.flow", {
        "data": queryResult["data"]
    })
    
    ' 调用结果保存模块
    SubFlow.Run("结果保存模块.flow", {
        "result": processResult["result"],
        "filePath": "C:\output.xlsx"
    })
    
    ' 调用通知模块
    SubFlow.Run("通知模块.flow", {
        "message": "流程执行成功",
        "recipient": "admin@example.com"
    })
Else
    Log.Error("登录失败")
End If

Log.Info("========== 主流程结束 ==========")
```

### 子流程示例

```vb
' 登录模块.flow
' 输入参数: inUsername, inPassword
' 输出参数: outSuccess, outMessage

Try
    objBrowser = WebBrowser.Create("chrome", "https://example.com/login", 30)
    Keyboard.InputText(@ui"用户名", inUsername, True, False)
    Keyboard.InputPwd(@ui"密码", inPassword, True)
    Mouse.Click(@ui"登录按钮")
    Delay(3000)
    
    If UiElement.Exists(@ui"首页标识", 5) Then
        outSuccess = True
        outMessage = "登录成功"
    Else
        outSuccess = False
        outMessage = "登录失败"
    End If
    
Catch ex
    outSuccess = False
    outMessage = "登录异常: " & ex.Message
End Try
```

### 适用场景

- ✅ 复杂的大型流程
- ✅ 需要代码复用
- ✅ 团队协作开发

---

## 设计模式对比

| 模式 | 复杂度 | 适用场景 | 优点 | 缺点 |
|------|--------|---------|------|------|
| REFramework | 高 | 企业级批量处理 | 健壮、可恢复 | 学习成本高 |
| 状态机 | 中 | 复杂流程控制 | 灵活、清晰 | 状态多时复杂 |
| 线性流程 | 低 | 简单任务 | 简单、直观 | 不适合复杂场景 |
| 事务处理 | 中 | 批量数据处理 | 独立、可统计 | 需要数据源支持 |
| 配置驱动 | 中 | 参数频繁变化 | 灵活、易维护 | 需要配置管理 |
| 模块化 | 中 | 大型项目 | 复用、协作 | 需要规划设计 |

---

## 最佳实践建议

### 1. 选择合适的模式
- 简单任务 → 线性流程
- 批量处理 → REFramework + 事务处理
- 复杂控制 → 状态机
- 大型项目 → 模块化 + 配置驱动

### 2. 组合使用
- REFramework + 配置驱动
- 模块化 + 事务处理
- 状态机 + 配置驱动

### 3. 设计原则
- ✅ 单一职责：每个模块只做一件事
- ✅ 开闭原则：对扩展开放，对修改关闭
- ✅ 依赖倒置：依赖配置而非硬编码
- ✅ 接口隔离：子流程接口清晰简洁

---

**文档版本**: v1.0.0  
**更新时间**: 2024-01-15  
**基于**: UIBot 6.0 企业级流程模板
