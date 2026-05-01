---
name: uibot
description: UIBot 6.0 命令手册 - 包含 400+ 个 RPA 自动化命令的完整参考
version: 1.0.0
---

# UIBot 6.0 命令手册

本 Skill 包含 UIBot 6.0.0.211215(64位) 的完整命令参考，涵盖 10 大类 400+ 个命令。

## 使用方法

直接询问命令功能或需求，例如：
- "如何点击按钮？"
- "怎么读取 Excel？"
- "如何发送邮件？"

---

## 编程基础知识

### 数据与数据类型

**数据**是计算机处理的基本单位。在RPA中，数据可以是数字、文字、日期等。

**数据类型**决定了数据的性质和可进行的操作：

| 数据类型 | 说明 | 示例 |
|---------|------|------|
| **整数型** | 不带小数点的数字 | `100`、`-50`、`&HFF`（十六进制） |
| **浮点数型** | 带小数点的数字 | `3.14`、`0.01`、`1E-2`（科学计数法） |
| **布尔型** | 逻辑真假值 | `True`、`False` |
| **字符串型** | 文本内容 | `"Hello"`、`'来也科技'` |
| **空值型** | 表示无值 | `Null` |
| **数组型** | 有序元素集合 | `[1, 2, 3]`、`["a", "b", "c"]` |
| **字典型** | 键值对集合 | `{"name": "张三", "age": 25}` |

### 变量与常量

**变量**是存储数据的容器，其值可以改变：

```vb
Dim a = 100              // 定义整数变量
Dim name = "张三"        // 定义字符串变量
Dim isValid = True       // 定义布尔变量
Dim arr = [1, 2, 3]      // 定义数组变量
```

**常量**的值在定义后不能改变：

```vb
Const PI = 3.14159       // 定义常量
Const MAX_COUNT = 100    // 定义常量
```

**变量命名规则**：
- 可使用字母、数字、下划线、汉字
- 不能以数字开头
- 不区分大小写（`abc` 和 `ABC` 是同一个变量）

### 运算符与表达式

**算术运算符**：

| 运算符 | 说明 | 示例 |
|--------|------|------|
| `+` | 加法 | `5 + 3` 结果为 `8` |
| `-` | 减法/求负 | `5 - 3` 结果为 `2` |
| `*` | 乘法 | `5 * 3` 结果为 `15` |
| `/` | 除法 | `6 / 3` 结果为 `2` |
| `^` | 求幂 | `2 ^ 3` 结果为 `8` |
| `Mod` | 取余数 | `7 Mod 3` 结果为 `1` |

**比较运算符**：

| 运算符 | 说明 | 示例 |
|--------|------|------|
| `=` | 等于 | `5 = 5` 结果为 `True` |
| `<>` | 不等于 | `5 <> 3` 结果为 `True` |
| `>` | 大于 | `5 > 3` 结果为 `True` |
| `<` | 小于 | `3 < 5` 结果为 `True` |
| `>=` | 大于等于 | `5 >= 5` 结果为 `True` |
| `<=` | 小于等于 | `3 <= 5` 结果为 `True` |

**逻辑运算符**：

| 运算符 | 说明 | 示例 |
|--------|------|------|
| `And` | 逻辑与 | `True And False` 结果为 `False` |
| `Or` | 逻辑或 | `True Or False` 结果为 `True` |
| `Not` | 逻辑非 | `Not True` 结果为 `False` |

**字符串运算符**：

| 运算符 | 说明 | 示例 |
|--------|------|------|
| `&` | 连接字符串 | `"Hello" & "World"` 结果为 `"HelloWorld"` |

### 条件判断

**If 语句**用于根据条件执行不同的代码：

```vb
// 单分支
If Time.Hour() > 18
    TracePrint("下班时间")
End If

// 双分支
If Time.Hour() > 18
    TracePrint("下班时间")
Else
    TracePrint("上班时间")
End If

// 多分支
If score >= 90
    TracePrint("优秀")
ElseIf score >= 80
    TracePrint("良好")
ElseIf score >= 60
    TracePrint("及格")
Else
    TracePrint("不及格")
End If
```

**Select Case 语句**用于多值选择：

```vb
Select Case Time.Month()
    Case 1, 3, 5, 7, 8, 10, 12
        DayOfMonth = 31
    Case 4, 6, 9, 11
        DayOfMonth = 30
    Case Else
        DayOfMonth = 28
End Select
```

### 循环

**For 循环**（计次循环）：

```vb
// 从1循环到100
For i = 1 To 100
    TracePrint(i)
Next

// 指定步长
For i = 0 To 10 Step 2
    TracePrint(i)  // 输出 0, 2, 4, 6, 8, 10
Next
```

**For Each 循环**（遍历循环）：

```vb
// 遍历数组
Dim arr = [10, 20, 30]
For Each item In arr
    TracePrint(item)
Next

// 遍历字典
Dim dict = {"name": "张三", "age": 25}
For Each key, value In dict
    TracePrint(key & ": " & value)
Next
```

**Do...Loop 循环**（条件循环）：

```vb
// 前置条件判断
Dim i = 1
Do While i <= 5
    TracePrint(i)
    i = i + 1
Loop

// 后置条件判断
Dim j = 1
Do
    TracePrint(j)
    j = j + 1
Loop Until j > 5

// 无限循环（需要用Break跳出）
Do
    If 某条件 Then
        Break  // 跳出循环
    End If
Loop
```

**循环控制语句**：

```vb
// Break - 跳出循环
For i = 1 To 10
    If i = 5 Then
        Break  // 当i等于5时跳出循环
    End If
    TracePrint(i)
Next

// Continue - 跳过本次循环
For i = 1 To 10
    If i Mod 2 = 0 Then
        Continue  // 跳过偶数
    End If
    TracePrint(i)  // 只输出奇数
Next
```

### 函数

**函数定义**：

```vb
// 无参数函数
Function SayHello()
    TracePrint("Hello, World!")
End Function

// 有参数函数
Function Add(x, y)
    Return x + y
End Function

// 带默认值的参数
Function Greet(name, greeting = "你好")
    Return greeting & ", " & name
End Function
```

**函数调用**：

```vb
// 调用无参数函数
SayHello()

// 调用有参数函数
result = Add(10, 20)
TracePrint(result)  // 输出 30

// 使用默认参数
msg1 = Greet("张三")           // 输出 "你好, 张三"
msg2 = Greet("李四", "早上好")  // 输出 "早上好, 李四"
```

### 数组操作

```vb
// 创建数组
Dim arr = [1, 2, 3, 4, 5]

// 访问元素（索引从0开始）
TracePrint(arr[0])  // 输出 1
TracePrint(arr[2])  // 输出 3

// 修改元素
arr[1] = 20
TracePrint(arr[1])  // 输出 20

// 多维数组
Dim matrix = [[1, 2], [3, 4], [5, 6]]
TracePrint(matrix[0][1])  // 输出 2
```

### 字典操作

```vb
// 创建字典
Dim person = {
    "name": "张三",
    "age": 25,
    "city": "北京"
}

// 访问元素
TracePrint(person["name"])  // 输出 "张三"

// 修改元素
person["age"] = 26

// 添加新元素
person["job"] = "工程师"

// 遍历字典
For Each key, value In person
    TracePrint(key & ": " & value)
Next
```

### 异常处理

```vb
// 基本异常处理
Try
    // 可能出错的代码
    result = 10 / 0
Catch ex
    // 处理异常
    TracePrint("发生错误: " & ex["Message"])
End Try

// 带重试的异常处理
Try 3  // 最多重试3次
    // 可能出错的代码
    Mouse.Click(@ui"按钮")
Catch ex
    TracePrint("重试3次后仍然失败")
Else
    TracePrint("操作成功")
End Try

// 手动抛出异常
If age < 0 Then
    Throw "年龄不能为负数"
End If
```

### 模块导入

```vb
// 导入其他流程块
Import MyModule

// 调用模块中的函数
MyModule.MyFunction()

// 直接运行模块
MyModule()
```

---

## BotScript 语言参考

### 语言概述

BotScript 是来也科技专为 RPA 开发设计的编程语言，具有以下特点：
- **简单易学**：接近自然语言，易于理解
- **动态类型**：变量类型可在运行时改变
- **不区分大小写**：变量名、关键字均不区分大小写
- **专为 RPA**：针对自动化场景优化设计

### 基本结构

**文件格式**：
- 纯文本格式，UTF-8 编码
- 扩展名通常为 `.task`
- 从第一行开始执行

**语句规则**：
- 一行一个语句（推荐）
- 多个语句用冒号 `:` 分隔
- 折行用反斜杠 `\` 或在逗号、运算符后直接换行

**注释**：
```vb
// 单行注释

/*
多行注释
可以跨越多行
*/
```

### 数据类型详解

**整数型**：
```vb
Dim a = 100        // 十进制
Dim b = &HFF       // 十六进制（255）
Dim c = &h10       // 十六进制（16）
```

**浮点数型**：
```vb
Dim x = 3.14       // 常规表示
Dim y = 1E-2       // 科学计数法（0.01）
Dim z = 2.5e3      // 科学计数法（2500）
```

**布尔型**：
```vb
Dim flag1 = True   // 真
Dim flag2 = FALSE  // 假（不区分大小写）
```

**字符串型**：
```vb
Dim s1 = "Hello"           // 双引号
Dim s2 = 'World'           // 单引号
Dim s3 = "多行
字符串"                     // 可以直接换行

// 转义字符
Dim s4 = "制表符:\t换行:\n引号:\"反斜杠:\\"

// 长字符串（无需转义）
Dim s5 = '''
这是一个长字符串
可以包含 "双引号" 和 '单引号'
无需转义
'''
```

**数组型**：
```vb
Dim arr1 = [1, 2, 3, 4, 5]                    // 一维数组
Dim arr2 = ["a", "b", "c"]                    // 字符串数组
Dim arr3 = [1, "text", True, Null]            // 混合类型数组
Dim arr4 = [[1, 2], [3, 4], [5, 6]]           // 二维数组
Dim arr5 = [[[1, 2], [3, 4]], [[5, 6], [7, 8]]]  // 三维数组
```

**字典型**：
```vb
Dim dict1 = {"name": "张三", "age": 25}
Dim dict2 = {
    "id": 1001,
    "info": {
        "city": "北京",
        "phone": "13800138000"
    },
    "tags": ["VIP", "活跃"]
}
```

**空值型**：
```vb
Dim empty = Null   // 空值（不区分大小写）
```

### 变量与常量详解

**变量定义**：
```vb
Dim a                      // 定义变量，不赋值
Dim b = 100                // 定义并赋值
Dim c, d = 200             // 定义多个变量
Dim e = 1, f = 2, g = 3    // 定义多个变量并赋值
```

**常量定义**：
```vb
Const PI = 3.14159         // 必须在定义时赋值
Const MAX = 100, MIN = 0   // 定义多个常量
```

**变量作用域**：
- **局部变量**：在函数内定义，函数退出时清空
- **流程块级变量**：在函数外定义，整个流程块运行期间有效

**引用赋值**：
```vb
Dim a = [1, 2, 3]
Dim b = a              // b 是 a 的引用（别名）
b[0] = 100             // 修改 b 会影响 a
TracePrint(a[0])       // 输出 100

Dim c = a[0]           // c 是值拷贝
c = 200                // 修改 c 不影响 a
TracePrint(a[0])       // 仍然输出 100
```

### 运算符优先级

从高到低：
1. `( )` 圆括号
2. `^` 求幂
3. `-` 负号（一元）
4. `*` `/` `Mod` 乘除取余
5. `+` `-` 加减
6. `&` 字符串连接
7. `=` `<>` `<` `>` `<=` `>=` 比较运算符
8. `Not` 逻辑非
9. `And` 逻辑与
10. `Or` 逻辑或

**示例**：
```vb
result = 2 + 3 * 4        // 结果为 14（先乘后加）
result = (2 + 3) * 4      // 结果为 20（括号优先）
result = 2 ^ 3 * 4        // 结果为 32（先幂后乘）
```

### 条件分支详解

**If 语句完整语法**：
```vb
If 条件1
    语句块1
ElseIf 条件2
    语句块2
ElseIf 条件3
    语句块3
Else
    语句块4
End If
```

**Select Case 语句**：
```vb
Select Case 表达式
    Case 值1, 值2, 值3
        语句块1
    Case 值4, 值5
        语句块2
    Case Else
        语句块3
End Select

// 示例
Select Case status
    Case "success", "ok"
        TracePrint("成功")
    Case "error", "fail"
        TracePrint("失败")
    Case Else
        TracePrint("未知状态")
End Select
```

### 循环详解

**Do...Loop 五种形式**：

1. **前置条件成立则循环**：
```vb
Do While 条件
    语句块
Loop
```

2. **前置条件不成立则循环**：
```vb
Do Until 条件
    语句块
Loop
```

3. **后置条件成立则循环**：
```vb
Do
    语句块
Loop While 条件
```

4. **后置条件不成立则循环**：
```vb
Do
    语句块
Loop Until 条件
```

5. **无限循环**：
```vb
Do
    语句块
    If 退出条件 Then
        Break
    End If
Loop
```

**For 循环详解**：
```vb
// 基本形式
For i = 1 To 10
    TracePrint(i)
Next

// 指定步长
For i = 10 To 1 Step -1
    TracePrint(i)  // 倒序输出
Next

// 浮点数循环
For x = 0.0 To 1.0 Step 0.1
    TracePrint(x)
Next
```

**For Each 循环详解**：
```vb
// 遍历数组（只获取值）
Dim arr = [10, 20, 30]
For Each value In arr
    TracePrint(value)
Next

// 遍历数组（获取索引和值）
For Each index, value In arr
    TracePrint("arr[" & index & "] = " & value)
Next

// 遍历字典（获取键和值）
Dim dict = {"a": 1, "b": 2, "c": 3}
For Each key, value In dict
    TracePrint(key & " => " & value)
Next
```

**循环控制**：
```vb
// Break - 立即跳出循环
For i = 1 To 100
    If i > 10 Then
        Break
    End If
    TracePrint(i)
Next

// Continue - 跳过本次循环，继续下一次
For i = 1 To 10
    If i Mod 2 = 0 Then
        Continue  // 跳过偶数
    End If
    TracePrint(i)
Next

// Exit - 结束整个流程
For i = 1 To 100
    If 发生严重错误 Then
        Exit  // 结束整个流程
    End If
Next
```

### 函数详解

**函数定义语法**：
```vb
Function 函数名(参数1, 参数2 = 默认值)
    语句块
    Return 返回值
End Function
```

**参数类型**：
```vb
// 必需参数
Function Add(x, y)
    Return x + y
End Function

// 可选参数（带默认值）
Function Greet(name, greeting = "你好")
    Return greeting & ", " & name
End Function

// 混合参数
Function Calculate(x, y, operator = "+")
    Select Case operator
        Case "+"
            Return x + y
        Case "-"
            Return x - y
        Case "*"
            Return x * y
        Case "/"
            Return x / y
    End Select
End Function
```

**函数调用方式**：
```vb
// 方式1：带括号，可获取返回值
result = Add(10, 20)

// 方式2：不带括号，不关心返回值
TracePrint "Hello"

// 方式3：函数作为变量
Dim myFunc = Add
result = myFunc(5, 3)
```

**函数参数传递**：
```vb
// 简单类型：值传递
Function ModifyValue(x)
    x = 100
End Function

Dim a = 10
ModifyValue(a)
TracePrint(a)  // 仍然是 10

// 复合类型：引用传递
Function ModifyArray(arr)
    arr[0] = 100
End Function

Dim b = [1, 2, 3]
ModifyArray(b)
TracePrint(b[0])  // 变成了 100
```

### 异常处理详解

**基本语法**：
```vb
Try
    // 可能出错的代码
Catch 异常变量
    // 处理异常
Else
    // 没有异常时执行
End Try
```

**异常变量结构**：
```vb
Try
    // 可能出错的代码
    result = 10 / 0
Catch ex
    TracePrint("文件: " & ex["File"])
    TracePrint("行号: " & ex["Line"])
    TracePrint("信息: " & ex["Message"])
End Try
```

**带重试的异常处理**：
```vb
Try 5  // 最多重试5次
    Mouse.Click(@ui"按钮")
Catch ex
    TracePrint("重试5次后仍然失败: " & ex["Message"])
Else
    TracePrint("操作成功")
End Try
```

**手动抛出异常**：
```vb
Function Divide(x, y)
    If y = 0 Then
        Throw "除数不能为零"
    End If
    Return x / y
End Function

Try
    result = Divide(10, 0)
Catch ex
    TracePrint("捕获异常: " & ex["Message"])
End Try
```

**异常处理最佳实践**：
```vb
// 针对特定操作使用重试
Try 3
    // 界面操作可能因卡顿失败
    Mouse.Click(@ui"按钮")
Catch ex
    // 记录日志
    TracePrint("点击失败: " & ex["Message"])
    // 采取备用方案
    Keyboard.Press("enter")
End Try

// 资源清理
Dim objExcel = Null
Try
    objExcel = Excel.Open("data.xlsx")
    // 处理数据
Catch ex
    TracePrint("处理失败: " & ex["Message"])
Finally
    // 确保关闭Excel（即使出错也执行）
    If objExcel <> Null Then
        Excel.Close(objExcel)
    End If
End Try
```

### 模块化编程

**导入模块**：
```vb
// 导入模块（文件名为 MyModule.task）
Import MyModule

// 调用模块中的函数
result = MyModule.Calculate(10, 20)

// 直接运行模块
MyModule()
```

**模块示例**：

文件：`MathUtils.task`
```vb
// 定义函数
Function Add(x, y)
    Return x + y
End Function

Function Multiply(x, y)
    Return x * y
End Function

// 模块级变量
Dim PI = 3.14159
```

文件：`Main.task`
```vb
// 导入模块
Import MathUtils

// 使用模块函数
result1 = MathUtils.Add(10, 20)
result2 = MathUtils.Multiply(5, 6)

// 访问模块变量
TracePrint("PI = " & MathUtils.PI)
```

### 编程技巧

**智能等待**：
```vb
// 等待元素出现
Dim maxWait = 30
Dim waited = 0
Do While Not UiElement.Exists(@ui"目标元素", 1)
    waited = waited + 1
    If waited >= maxWait Then
        Throw "等待超时"
    End If
    Delay(1000)
Loop
```

**批量操作**：
```vb
// 使用数组批量处理
Dim tasks = ["任务1", "任务2", "任务3"]
For Each task In tasks
    Try
        ProcessTask(task)
        TracePrint("完成: " & task)
    Catch ex
        TracePrint("失败: " & task & " - " & ex["Message"])
        Continue
    End Try
Next
```

**日志记录**：
```vb
Function WriteLog(message)
    Dim timestamp = Time.Format(Time.Now(), "yyyy-MM-dd HH:mm:ss")
    Dim logFile = "C:\Logs\automation.log"
    Dim logLine = "[" & timestamp & "] " & message & "\n"
    File.Append(logFile, logLine, "utf-8")
    TracePrint(message)
End Function

// 使用日志
WriteLog("流程开始")
Try
    // 执行任务
    WriteLog("任务执行成功")
Catch ex
    WriteLog("任务失败: " & ex["Message"])
End Try
WriteLog("流程结束")
```

**配置管理**：
```vb
// 从配置文件读取设置
Dim configFile = "config.json"
Dim configText = File.Read(configFile, "utf-8")
Dim config = JSON.Parse(configText)

// 使用配置
Dim serverUrl = config["server"]["url"]
Dim timeout = config["server"]["timeout"]
Dim username = config["credentials"]["username"]
```

---

## 命令分类

1. **基本命令** - 数据转换、变量操作、流程控制
2. **鼠标键盘** - 鼠标点击、键盘输入
3. **界面操作** - 元素操作、窗口管理、图像识别
4. **智能文档** - OCR、表格识别、票据识别
5. **软件自动化** - 浏览器、Excel、Word、数据库
6. **数据处理** - 字符串、数组、正则、时间
7. **文件处理** - 文件读写、CSV、PDF
8. **系统操作** - 系统命令、应用管理、对话框
9. **网络** - HTTP、FTP、邮件
10. **机器人指挥官** - 数据队列、表单协同

---


## 获取应用运行状态

**说明**: 获取应用的运行状态，如果应用仍在运行返回 true，如果应用已经退出返回 false  

**原型**: `bRet=App.GetStatus(sName)`  

**参数**:  
- **sName** (True) [string] 默认:"" - 应用程序进程名或进程PID  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*******************************获取应用运行状态********************************** 命令原型： bRet=App.GetStatus("") 入参： sName--应用程序进程名或进程PID 出参： bRet--命令运行后的结果 注意事项： 无 ********************************************************************************/ Dim sName = "Chrome.exe" bRet=App.GetStatus(sName) TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/App_图片/App_GetStatus.png)  

---

## 关闭应用

**说明**: 强制停止应用程序的运行（结束进程）  

**原型**: `App.Kill(sName)`  

**参数**:  
- **sName** (True) [string] 默认:"" - 应用程序进程名或进程PID  

**示例**:  
```
/*********************************关闭应用************************************ 命令原型 App.Kill("") 入参： sName--应用程序进程名或进程PID 出参： 无 注意事项： 无 ****************************************************************************/ Dim sName = "YoudaoDict.exe" App.Kill(sName)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/App_图片/App_Kill.png)  

---

## 启动应用程序

**说明**: 启动一个应用程序，返回应用程序的 PID  

**原型**: `iPID = App.Run(sPath, iWait, iShow)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 应用程序文件路径  
- **iWait** (True) [enum] 默认:0 - 等待方式  
- **iShow** (True) [enum] 默认:1 - 程序启动后的显示样式（不一定生效）  

**返回**: iPID，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************************启动应用程序*********************************** 命令原型： iPID = App.Run(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;, 0, 1) 入参： sPath--应用程序文件路径 iWait-等待方式：0是等待、1是等待应用程序准备好、2是等待应用程序执行到退出 iShow-程序启动后的显示样式（不一定生效）：0是隐藏、1是默认、3是最大化、6是最小化 出参： iPID--命令运行后的结果 注意事项： 等待方式默认不等待、显示样式为初始状态，可以切换至可视化界面，在对应属性栏进行选择 ********************************************************************************/ Dim sPath = &#x27;&#x27;&#x27;D:\app\Dict\YoudaoDict.exe&#x27;&#x27;&#x27; Dim iWait=0 Dim iShow=1 iPID = App.Run(sPath, iWait, iShow) TracePrint(iPID)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/App_图片/App_Run.png)  

---

## 打开文件或网址

**说明**: 使用系统关联打开一个文件或网站URL  

**原型**: `iPID = App.Start(sPath, iWait, iShow)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要打开的文件路径或网站URL  
- **iWait** (True) [enum] 默认:0 - 等待方式  
- **iShow** (True) [enum] 默认:1 - 程序启动后的显示样式（不一定生效）  

**返回**: iPID，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************打开文件或网址********************************** 命令原型： iPID = App.Start(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;, 0, 1) 入参： sPath--应用程序文件路径 iWait-等待方式：0是等待、1是等待应用程序准备好、2是等待应用程序执行到退出 iShow-程序启动后的显示样式（不一定生效）：0是隐藏、1是默认、3是最大化、6是最小化 出参： iPID--命令运行后的结果 注意事项： 等待方式默认不等待、显示样式为初始状态，可以切换至可视化界面，在对应属性栏进行选择 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;D:\App\Firefox\Firefox.exe&#x27;&#x27;&#x27; Dim iWait=0 Dim iShow=1 iPID = App.Start(sPath, iWait, iShow)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/App_图片/App_Start.png)  

---

## 创建多维数组

**说明**: 根据指定的维数创建一维、二维、三维或更高维的数组  

**原型**: `arrayRet = Array(arrSize,defaultValue)`  

**参数**:  
- **arrSize** (True) [expression] 默认:[5,5] - 如果该属性是一个整数N，则新建一个包含N个元素的一维数组；如果该属性是数组 [M, N] ，则新建一个包含M N个元素的二维数组；如果该属性是数组 [M, N, Q] ，则新建一个包含M N*Q个元素的三维数组；总之当该属性为数组时，且数组中的元素个数为X，则输出是一个X维数组，其中的每一维的个数由属性中对应位置的整数来表示  
- **defaultValue** (True) [expression] 默认:null - 被创建的高维数组中的元素全部由此值填充，默认为null  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************创建多维数组************************ 命令原型: arrayRet=Array(arrSize,defaultValue) 入参： arrSize--如果该属性是一个整数N，则新建一个包含N个元素的一维数组；如果该属性是数组[M, N]，则新建一个包含MN个元素的二维数组；如果该属性是数组[M, N, Q]，则新建一个包含MN*Q个元素的三维数组；总之当该属性为数组时，且数组中的元素个数为X，则输出是一个X维数组，其中的每一维的个数由属性中对应位置的整数来表示。 defaultValue--被创建的高维数组中的元素全部由此值填充，默认为null。 出参： arrayRet--函数调用的输出保存到的变量。 注意事项: 根据指定的维数创建一维、二维、三维或更高维的数组。 ***********************************************************/ arrayRet=Array([2,2],0)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Array_图片/Array.png)  

---

## 合并数组

**说明**: 合并两个数组  

**原型**: `arrRet = concat(srcArray,distArray)`  

**参数**:  
- **srcArray** (True) [expression] 默认:[] - 需要合并的数组  
- **distArray** (True) [expression] 默认:[] - 需要合并的数组  

**返回**: arrRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************合并数组************************ 命令原型: arrRet = concat(srcArray,distArray) 入参： srcArray--需要合并的数组。 distArray--需要合并的数组。 出参： arrRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ arrRet=concat(["a", "b"],["c", "d"])
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Array_图片/Concat.png)  

---

## 过滤数组数据

**说明**: 过滤数组的中的字符串，可选择是否保留过滤文字  

**原型**: `arrRet = Filter(arrParam,sFilter,bInclude)`  

**参数**:  
- **arrParam** (True) [expression] 默认:[] - 要进行过滤的数组  
- **sFilter** (True) [string] 默认:"" - 过滤使用的字符串，对数组元素逐个进行匹配到，当输入 null 时 ，将会过滤数组中所有的非字符串和空字符串  
- **bInclude** (True) [boolean] 默认:True - 使用过滤内容在目标数组中进行匹配，当匹配不到数组元素时，若保留过滤文字则返回空数组，若不保留过滤文字则返回原目标数组；当能匹配到数组元素时，若保留过滤文字则返回匹配到所有元素组成的数组，若不保留过滤文字则返回去除匹配到的元素后所有元素组成的数组  

**返回**: arrRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************过滤数组数据************************ 命令原型: arrRet=Filter(arrParam,sFilter,bInclude) 入参： arrParam--要进行过滤的数组。 sFilter--过滤使用的字符串，对数组元素逐个进行匹配到，当输入 null 时 ，将会过滤数组中所有的非字符串和空字符串。 bInclude--使用过滤内容在目标数组中进行匹配，当匹配不到数组元素时，若保留过滤文字则返回空数组，若不保留过滤文字则返回原目标数组；当能匹配到数组元素时，若保留过滤文字则返回匹配到所有元素组成的数组，若不保留过滤文字则返回去除匹配到的元素后所有元素组成的数组。 出参： arrRet--函数调用的输出保存到的变量。 注意事项: 过滤数组中的字符串，可选择是否保留过滤文字 ***********************************************************/ arrRet=Filter(["UiBot","RPA","UiBot123"],"Bot",True)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Array_图片/Filter.png)  

---

## 插入元素

**说明**: 在数组指定位置添加一个元素  

**原型**: `arrRet = insert(array,postion,item)`  

**参数**:  
- **array** (True) [expression] 默认:[] - 要插入元素的数组  
- **postion** (True) [number] 默认:0 - 插入元素的位置  
- **item** (True) [expression] 默认:"" - 要插入到数组的元素  

**返回**: arrRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************插入元素************************ 命令原型: arrRet=insert(array,postion,item) 入参： array--要插入元素的数组。 postion--插入元素的位置。 item--要插入到数组的元素。 出参： arrRet--函数调用的输出保存到的变量。 注意事项: 在数组指定位置添加一个元素。 ***********************************************************/ arrRet=insert([1,3,4],1,2)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Array_图片/Insert.png)  

---

## 将数组合并为字符串

**说明**: 将数组拼接成字符串，使用指定的分隔符分割数组元素  

**原型**: `sRet = Join(arrData,sSeparator)`  

**参数**:  
- **arrData** (True) [expression] 默认:[] - 要进行合并的数组  
- **sSeparator** (True) [string] 默认:"," - 合并数组时使用的分隔符  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************将数组合并为字符串************************ 命令原型: sRet=Join(arrData,sSeparator) 入参： arrData--要进行合并的数组。 sSeparator--合并数组时使用的分隔符。 sRet--函数调用的输出保存到的变量。 注意事项: 无。 ***********************************************************/ sRet=Join(["RPA","UiBot"],",")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Array_图片/Join.png)  

---

## 删除并返回最后元素

**说明**: 删除并返回数组的最后一个元素  

**原型**: `item = pop(array)`  

**参数**:  
- **array** (True) [expression] 默认:[] - 需要删除并返回数组的最后一个元素的数组  

**返回**: item，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************删除并返回最后元素************************ 命令原型: item=pop(array) 入参： array--需要删除并返回数组的最后一个元素的数组。 出参： item--函数调用的输出保存到的变量。 注意事项: 无。 ***********************************************************/ item=pop(["a","b","c"])
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Array_图片/Pop.png)  

---

## 在数组尾部添加元素

**说明**: 在数组尾部添加元素并返回数组  

**原型**: `arrRet = push(array,item)`  

**参数**:  
- **array** (True) [expression] 默认:[] - 要添加元素的数组  
- **item** (True) [expression] 默认:"" - 要添加到数组的元素  

**返回**: arrRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************在数组尾部添加元素************************ 命令原型: arrRet = push(array,item) 入参： array--要添加元素的数组。 item--要添加到数组的元素。 出参： arrRet--函数调用的输出保存到的变量。 注意事项: 无。 ***********************************************************/ arrRet=push(["a","b"],"c")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Array_图片/Push.png)  

---

## 删除并返回第一个元素

**说明**: 删除并返回数组的第一个元素  

**原型**: `item = Shift(array)`  

**参数**:  
- **array** (True) [expression] 默认:[] - 需要删除并返回第一个元素的数组  

**返回**: item，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************删除并返回第一个元素************************ 命令原型: item=Shift(array) 入参： array--需要删除并返回第一个元素的数组。 出参： item--函数调用的输出保存到的变量。 注意事项: 无。 ***********************************************************/ item=Shift(["a","b","c"])
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Array_图片/Shift.png)  

---

## 截取数组

**说明**: 截取数组从指定位置开始到指定位置结束的元素，返回数组  

**原型**: `arrRet = splice(array,begin,end)`  

**参数**:  
- **array** (True) [expression] 默认:[] - 需要截取数组元素的数组  
- **begin** (True) [number] 默认:0 - 要截取元素的开始位置  
- **end** (True) [number] 默认:0 - 要截取元素的结束位置  

**返回**: arrRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************截取数组************************ 命令原型: arrRet=splice(array,begin,end) 入参： array--需要截取数组元素的数组。 begin--要截取元素的开始位置。 end--要截取元素的结束位置。 出参： arrRet--函数调用的输出保存到的变量。 注意事项: 无。 ***********************************************************/ arrRet=splice(["a","b","c","d"],0,1)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Array_图片/Splice.png)  

---

## 获取数组最大下标

**说明**: 获取数组的元素数量（下标）  

**原型**: `iRet = UBound(arrData)`  

**参数**:  
- **arrData** (True) [expression] 默认:[] - 要操作的数组  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************获取数组最大下标************************ 命令原型: iRet = UBound(arrData) 入参： arrData--要操作的数组。 出参： iRet--函数调用的输出保存到的变量。 注意事项: 无。 ***********************************************************/ iRet=UBound(["a","b","c"])
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Array_图片/UBound.png)  

---

## 在数组头部添加元素

**说明**: 在数组头部添加元素并返回数组  

**原型**: `arrRet = Unshift(array,item)`  

**参数**:  
- **array** (True) [expression] 默认:[] - 要添加元素的数组  
- **item** (True) [expression] 默认:"" - 要添加到数组的元素  

**返回**: arrRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************在数组头部添加元素************************ 命令原型: arrRet=Unshift(array,item) 入参： array--要添加元素的数组。 item--要添加到数组的元素。 出参： arrRet--函数调用的输出保存到的变量。 注意事项: 无。 ***********************************************************/ arrRet=Unshift(["b","c"],"a")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Array_图片/Unshift.png)  

---

## 转为逻辑数据

**说明**: 将数据转换为逻辑类型  

**原型**: `bRet = CBool(varData)`  

**参数**:  
- **varData** (True) [expression] 默认:PrevResult - 要进行转换的数据  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************转为逻辑数据*************************************** 命令原型： bRet = CBool(varData) 入参： varData -- 目标数据 出参： bRet -- 返回布尔类型的值，true和false 注意事项： 将数据转为逻辑类型 ********************************************************************************/ TracePrint("将目标数据转换为布尔值") bRet = CBool("Hello UiBot") TracePrint("将字符串Hello UiBot转换为布尔值，结果为：") TracePrint(bRet) bRet = CBool(0) TracePrint("将数字0转换为布尔值，结果为：") TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_CBool.png)  

---

## 转为整数数据

**说明**: 将数据转换为整数类型  

**原型**: `iRet = CInt(varData)`  

**参数**:  
- **varData** (True) [expression] 默认:PrevResult - 要进行转换的数据  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************转为整数数据*************************************** 命令原型： iRet = CInt(varData) 入参： varData--目标数据 出参： 无 注意事项： 该命令会将数字类型和类数字类型的字符串四舍五入为整数数据 ********************************************************************************/ TracePrint("将小数1.5四舍五入") dRet = CInt(1.5) TracePrint("转换后数据类型为："& dRet) TracePrint("将字符串1.5四舍五入") dRet = CInt("1.5") TracePrint("转换后数据类型为："& dRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_CInt.png)  

---

## 复制数据

**说明**: 复制数据，可以用来复制字典和数组  

**原型**: `varRet = Clone(varData)`  

**参数**:  
- **varData** (True) [expression] 默认:$PrevResult - 要进行复制的数据  

**返回**: varRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************复制数据*************************************** 命令原型： varRet = Clone(varData) 入参： varData--目标数据 出参： 无 注意事项： 复制数据，可用来复制数组和字典 ********************************************************************************/ a = [1,2,3] varRet = Clone(a) TracePrint("另数组a为[1,2,3],复制后赋值给varRet,varRet值为：") TracePrint(varRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_Clone.png)  

---

## 转为小数数据

**说明**: 将数据转换为小数（浮点数）类型  

**原型**: `dRet = CNumber(varData)`  

**参数**:  
- **varData** (True) [expression] 默认:PrevResult - 要进行转换的数据  

**返回**: dRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************转为小数数据*************************************** 命令原型： dRet = CNumber(varData) 入参： varData--目标数据 出参： 无 注意事项： 该命令会将数字类型和类数字类型的字符串四舍五入为整数数据 ********************************************************************************/ TracePrint("将字符串‘1.5’转换为小数") TracePrint("转换前数据类型为："&Type("1.5")) dRet = CNumber("1.5") TracePrint("转换后数据类型为："&type(dRet))
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_CNumber.png)  

---

## 垃圾回收

**说明**: 回收不再使用的内存空间  

**原型**: `CollectGarbage()`  

**参数**:  
- **无** (无) [无] 默认:无 - 回收不再使用的内存空间  

**示例**:  
```
/*********************************垃圾回收*************************************** 命令原型： CollectGarbage() 入参： 无 出参： 无 注意事项： 回收不再使用的内存空间 ********************************************************************************/ CollectGarbage()
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_CollectGarbage.png)  

---

## 注释

**说明**: 用于给其他命令做注释说明，运行时没有任何效果  

**原型**: `Rem sText`  

**参数**:  
- **sText** (True) [rem] 默认:无 - 需要显示的注释内容  

**示例**:  
```
/*********************************注释*************************************** 命令原型： Rem sText 入参： sText -- 需要显示的注释内容 出参： 无 注意事项： 注释的内容不会在流程中编译执行 ********************************************************************************/ Rem "这是一行注释"
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_Comment.png)  

---

## 转为文字数据

**说明**: 将数据转换为文字（字符串）类型  

**原型**: `sRet = CStr(varData)`  

**参数**:  
- **varData** (True) [expression] 默认:$PrevResult - 要进行转换的数据  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************转为文字数据*************************************** 命令原型： sRet = CStr(varData) 入参： varData--要进行转换的数据 出参： 无 注意事项： 将数据转换为文字（字符串）类型 ********************************************************************************/ TracePrint("将数字1转换为文字类型") TracePrint("转换前数据类型为："&Type(1)) dRet = CStr(1) TracePrint("转换后数据类型为："&type(dRet))
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_CStr.png)  

---

## 转为十进制数字

**说明**: 将整数、浮点数或字符串转为十进制数字，默认保存28个整数和小数，超过28位会在28位处截断  

**原型**: `deRet = Decimal(varData)`  

**参数**:  
- **varData** (True) [expression] 默认:$PrevResult - 要进行转换的数据，建议传入数字字符串类型  

**返回**: deRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************转为十进制数字*************************************** 命令原型： deRet = Decimal(varData) 入参： varData -- 要进行转换的数据，建议传入数字字符串类型 出参： deRet -- 十进制数字 注意事项： 将整数、浮点数或字符串转为十进制数字，默认保存28个整数和小数，超过28位会在28位处截断 ********************************************************************************/ TracePrint("将字符串12转换为十进制数字") varData="12" deRet = Decimal(varData) TracePrint("转换后结果为：") TracePrint(deRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_Decimal.png)  

---

## 延时

**说明**: 延时等待 ms 毫秒后继续执行之后的代码  

**原型**: `Delay(ms)`  

**参数**:  
- **ms** (True) [number] 默认:1000 - 延时等待的时间（毫秒，1 秒等于 1000 毫秒）  

**示例**:  
```
/*********************************延时*************************************** 命令原型： Delay(ms) 入参： ms -- 正整数 出参： 无 注意事项： 延时等待 ms 毫秒后继续执行之后的代码 ********************************************************************************/ dTime = Time.Now() sRet = Time.Format(dTime,"yyyy-mm-dd hh:mm:ss") TracePrint("当前时间为:"&sRet) Delay(1000) dTime = Time.Now() sRet = Time.Format(dTime,"yyyy-mm-dd hh:mm:ss") TracePrint("延时一秒后，当前时间为"&sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_Delay.png)  

---

## 子程序

**说明**: 创建一个可复用的子程序  

**原型**: `Function name(属性1) End Function`  

**参数**:  
- **name** (True) [id] 默认:命令名 - 定义子程序的命令名  

**示例**:  
```
/*********************************子程序*************************************** 命令原型： Function name(prop) End Function 入参： name--定义子程序的命令名 prop--设置命令的属性 出参： 无 注意事项： 无 ********************************************************************************/ TracePrint ("定义个一个叫test的子程序，传入两个参数分别为（a、b），返回两个数据的和") Function test(a,b) Return a+b End Function TracePrint ("调用test子程序，传入参数（1，2）") test(1,2) TracePrint ("执行结果："&test(1,2))
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_Function.png)  

---

## 取十六进制

**说明**: 获取一个整数的十六进制表现形式  

**原型**: `sRet = Hex(iData)`  

**参数**:  
- **iData** (True) [number] 默认:0 - 要进行转换的整数  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取十六进制*************************************** 命令原型： sRet = Hex(iData) 入参： iData -- 要进行转换的整数 出参： sRet -- 十六进制数字 注意事项： 无 ********************************************************************************/ a = 111 sRet = Hex(a) TracePrint("取整数111的十六进制，结果为："&sRet) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_Hex.png)  

---

## 是否为数组

**说明**: 判断一个数据是否为数组  

**原型**: `bRet = IsArray(varData)`  

**参数**:  
- **varData** (True) [expression] 默认:$PrevResult - 要进行判断的数据  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************是否为数组*************************************** 命令原型： bRet = IsArray(varData) 入参： varData -- 要进行判断的数据 出参： sRet -- 返回是否为数组的结果，true和false 注意事项： 字符串虽然可以遍历，但是不算数组 ********************************************************************************/ a = [1,2,3] bRet = IsArray(a) TracePrint("令变量a为[1,2,3]判断变量a是否为数组，结果为：") TracePrint(bRet) a = "Hello UiBot" bRet = IsArray(a) TracePrint("令变量a为Hello UiBot判断变量a是否为数组，结果为：") TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_IsArray.png)  

---

## 是否为字典

**说明**: 判断一个数据是否为字典  

**原型**: `bRet = IsDictionary(varData)`  

**参数**:  
- **varData** (True) [expression] 默认:$PrevResult - 要进行判断的数据  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************是否为字典*************************************** 命令原型： Ret = IsDictionary(varData) 入参： varData -- 要进行判断的数据 出参： sRet -- 返回是否为字典的结果，true和false 注意事项： Json字符串可以通过命令转换为字典 ********************************************************************************/ a = {"a":1,"b":2,"c":3} bRet = IsDictionary(a) TracePrint("令变量a为{&#x27;a&#x27;:1,&#x27;b&#x27;:2,&#x27;c&#x27;:3}判断变量a是否为字典，结果为：") TracePrint(bRet) a = "Hello UiBot" bRet = IsDictionary(a) TracePrint("令变量a为Hello UiBot判断变量a是否为字典，结果为：") TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_IsDictionary.png)  

---

## 是否为空值

**说明**: 判断一个数据是否为空值（NULL）  

**原型**: `bRet = IsNull(varData)`  

**参数**:  
- **varData** (True) [expression] 默认:$PrevResult - 要进行判断的数据  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************是否为空值*************************************** 命令原型： bRet = IsNull(varData) 入参： varData -- 要进行判断的数据 出参： sRet -- 返回是否为空值的结果，true和false 注意事项： 无 ********************************************************************************/ a = Null bRet = IsNull(a) TracePrint("令变量a为null,判断是否为空，结果为：") TracePrint(bRet) a = "Hello UiBot" bRet = IsNull(a) TracePrint("令变量a为Hello UiBot判断变量a是否为空，结果为：") TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_IsNull.png)  

---

## 是否为数值

**说明**: 判断一个数据是否为数值（可转换为小数或整数）  

**原型**: `bRet = IsNumeric(varData)`  

**参数**:  
- **varData** (True) [expression] 默认:$PrevResult - 要进行判断的数据  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************是否为数值*************************************** 命令原型： bRet = IsNumeric(varData) 入参： varData -- 要进行判断的数据 出参： sRet -- 返回是否为数值的结果，true和false 注意事项： 无 ********************************************************************************/ a = "111" bRet = IsNumeric(a) TracePrint("令变量a为111,判断是否为数值，结果为：") TracePrint(bRet) a = "Hello UiBot" bRet = IsNumeric(a) TracePrint("令变量a为Hello UiBot判断变量a是否为数值，结果为：") TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_IsNumeric.png)  

---

## 获取长度

**说明**: 获取字符串或数组的长度  

**原型**: `iRet = Len(varData)`  

**参数**:  
- **varData** (True) [expression] 默认:$PrevResult - 要操作的字符串或数组  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取字符串或数组的长度*************************************** 命令原型： iRet = Len(varData) 入参： varData -- 字符串或数组 出参： sRet -- 返回字符串长度 注意事项： 即可以获取字符串长度又可以获取数组长度 ********************************************************************************/ a = "Hello World" bRet = Len(a) TracePrint("字符串变量a的长度为：") TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_Len.png)  

---

## 取八进制

**说明**: 获取一个整数的八进制表现形式  

**原型**: `sRet = Oct(iData)`  

**参数**:  
- **iData** (True) [number] 默认:0 - 要进行转换的整数  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取八进制*************************************** 命令原型： sRet = Oct(iData) 入参： iData -- 要进行转换的整数 出参： sRet -- 指定数据的八进制值 注意事项： 无 ********************************************************************************/ a = 111 sRet = Oct(111) TracePrint("取整数111的八进制，结果为："&sRet) TracePrint sRet
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_Oct.png)  

---

## 删除指定元素

**说明**: 删除数组中指定位置（从0开始计数）的元素，或者删除字典中指定键名的元素  

**原型**: `Remove(varData,index)`  

**参数**:  
- **varData** (True) [expression] 默认:$PrevResult - 输入需要删除其中元素的数组或者字典  
- **index** (True) [number] 默认:0 - 输入数组中的位置（从0开始计数），或输入字典中的键名  

**示例**:  
```
/*********************************删除指定元素*************************************** 命令原型： Remove(varData,index) 入参： varData--输入需要删除其中元素的数组或者字典。 index--输入数组中的位置（从0开始计数），或输入字典中的键名。 出参： 无 注意事项： 该命令无返回值，会修改原本的数组 ********************************************************************************/ Dim varData=[1,2,3,4,2] Remove(varData,4) TracePrint "删除数组varData中的第5个元素：" TracePrint varData
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_Remove.png)  

---

## 取随机数

**说明**: 获取一个 0 - 1 之间的随机数  

**原型**: `dRet = Rnd()`  

**参数**:  
- **无** (无) [无] 默认:无 - 无  

**返回**: dRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取随机数*************************************** 命令原型： dRet = Rnd() 入参： 无 出参： dRet -- 随机数 注意事项： 无 ********************************************************************************/ dRet = Rnd() TracePrint("获取一个0-1之间的随机数，结果为："&Cstr(dRet))
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_Rnd.png)  

---

## 输出调试信息

**说明**: 输出调试信息  

**原型**: `TracePrint(sText)`  

**参数**:  
- **sText** (True) [expression] 默认:$PrevResult - 调试信息内容  

**示例**:  
```
/*********************************输出调试信息*************************************** 命令原型： TracePrint(sText) 入参： sText = 想要输出的内容,可以是变量也可以是其他数据类型 出参： 无 注意事项： 无 ********************************************************************************/ TracePrint("Hello UiBot")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_TracePrint.png)  

---

## 获取变量类型

**说明**: 获取变量的类型，根据类型返回不同的字符串值:int，float，Decimal，string，bool，null，array，dictionary，function，object，unknown  

**原型**: `sRet = Type(varData)`  

**参数**:  
- **varData** (True) [expression] 默认:$PrevResult - 要进行判断的数据  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取变量类型*************************************** 命令原型： sRet = Type(varData) 入参： varData -- 要进行判断的数据 出参： sRet -- 变量类型 注意事项： 获取变量的类型，根据不同类型返回string,int,float,bool...等 ********************************************************************************/ a = [1,2,3] bRet = Type(a) TracePrint("[1,2,3]的变量类型为："&bRet) a = "Hello UiBot" bRet = Type(a) TracePrint("Hello UiBot的变量类型为："&bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_Type.png)  

---

## 单元测试块

**说明**: 单元测试块中可以添加任何命令，单元测试块中的命令只在单独运行流程块时有效，运行时会先执行单元测试块中的命令，然后再执行流程块中本身的命令  

**原型**: `UnitTest End UnitTest`  

**参数**:  
- **name** (True) [id] 默认:命令名 - 定义单元测试块的命令名  

**示例**:  
```
/*********************************单元测试块*************************************** 命令原型： UnitTest End UnitTest 入参： name -- 定义单元测试块的命令名 出参： 无 注意事项： 单元测试块中可以添加任何命令，单元测试块中的命令只在单独运行流程块时有效，运行时会先执行单元测试块中的命令，然后再执行流程块中本身的命令。 ********************************************************************************/ TracePrint("分别打印123，456，789，其中456在单元测试块中，结果为：") TracePrint(123) UnitTest Delay(1000) TracePrint(456) End UnitTest TracePrint(789)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Base_图片/Base_UnitTestBlock.gif)  

---

## 读取剪贴板文本

**说明**: 读取剪贴板文本  

**原型**: `sRet = Clipboard.GetText()`  

**参数**:  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************读取剪贴板文本*************************************** 命令原型： sRet = Clipboard.GetText() 入参： 无 出参： sRet--将命令运行后的结果赋值给此变量。 注意事项： 1.该命令只能用于读取剪贴板中的文字数据，如果存在图片则不进行读取。 2.该命令读取文字后只保留换行空格等格式，对文字大小颜色无法读取。 ********************************************************************************/ Dim sRet = "" sRet = Clipboard.GetText() TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Clipboard_图片/Clipboard_GetText.png)  

---

## 保存剪贴板图像

**说明**: 将剪贴板中的图像数据保存到指定路径  

**原型**: `Clipboard.SaveImage(sPath)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要将剪贴板的图像保存到的路径  

**示例**:  
```
/*********************************保存剪贴板图像*************************************** 命令原型： Clipboard.SaveImage(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： sPath--文件保存路径。注：要将剪贴板的图像保存到的路径 注意事项： 1.该方法只支持保存*.jpeg、*.jpg、*.png、*.bmp、*.tif、*.tiff几种图片格式。 2.如果在剪贴板中不存在图片，方法不会出错但是也不生成任何图片文件。 ********************************************************************************/ Clipboard.SaveImage(@res&#x27;123.png&#x27;)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Clipboard_图片/Clipboard_SaveImage.png)  

---

## 图片设置到剪贴板

**说明**: 将一副图片放入剪贴板  

**原型**: `Clipboard.SetImage(sPath)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要放入剪贴板的文件路径，多个文件可以使用多行分开  

**示例**:  
```
/*********************************图片设置到剪贴板*************************************** 命令原型： Clipboard.SetImage(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： sPath--文件路径。注：要放入剪贴板的文件路径，多个文件可以使用多行分开 注意事项： 1.该方法只支持保存*.jpeg、*.jpg、*.png、*.bmp、*.tif、*.tiff几种图片格式。 ********************************************************************************/ Clipboard.SetImage(@res"8c2637f0-7daa-11ec-ac0d-372b37312174.png")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Clipboard_图片/Clipboard_SetImage.png)  

---

## 设置剪贴板文本

**说明**: 设置剪贴板文本  

**原型**: `Clipboard.SetText(sText)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 新的剪贴板文本内容  

**示例**:  
```
/*********************************设置剪贴板文本*************************************** 命令原型： Clipboard.SetText("") 入参： sText--剪贴板内容。注：新的剪贴板文本内容 注意事项： 1.该命令只能将文字设置到剪贴板中。 2.该命令没有字符串长度限制。 ********************************************************************************/ Clipboard.SetText("剪贴板内容")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Clipboard_图片/Clipboard_SetText.png)  

---

## 添加凭据

**说明**: 添加一个凭据，添加的凭据可在Windows系统的凭据管理器中查看  

**原型**: `bRet = Credential.Add(cName,userName,password,cType,sType)`  

**参数**:  
- **cName** (True) [string] 默认:"" - 该凭据的名称，用于区分每个凭据  
- **userName** (True) [string] 默认:"" - 设置凭据的用户名  
- **password** (True) [string] 默认:"" - 设置凭据的访问密码  
- **cType** (True) [enum] 默认:"normal" - 凭据所属的类型  
- **sType** (True) [enum] 默认:"enterprise" - 凭据保存的类型  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************添加凭据*************************************** 命令原型： bRet = Credential.Add("","","","normal","enterprise") 入参： cName--凭据名。注：该凭据的名称，用于区分每个凭据 userName--用户名。注：设置凭据的用户名 password--密码。注：设置凭据的访问密码 cType--凭据类型。注：凭据所属的类型 sType--保存类型。注：凭据保存的类型 出参： bRet--函数调用的输出保存到的变量。 注意事项： 1.如果对相同名称的凭据进行重复添加则进行覆盖，请谨慎使用。 2.windows凭据是用于windows客户端自用的凭据比如远程连接，共享文件这种。普通凭据用于web，第三方的客户端。 ********************************************************************************/ Dim bRet = "" bRet = Credential.Add("命令库6.0","laiye","test","normal","enterprise") TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Credential_图片/Credential_Add.png)  

---

## 删除凭据

**说明**: 删除指定凭据  

**原型**: `bRet = Credential.Delete(cName,cType)`  

**参数**:  
- **cName** (True) [string] 默认:"" - 要删除凭据的凭据名  
- **cType** (True) [enum] 默认:"normal" - 凭据所属的类型  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************删除凭据*************************************** 命令原型： bRet = Credential.Delete("","normal") 入参： cName--凭据名。注：该凭据的名称，用于区分每个凭据 cType--凭据类型。注：凭据所属的类型 出参： bRet--函数调用的输出保存到的变量。 注意事项： 1.本地windows要存在该凭据，不存在会返回false。 2.在删除凭据时注意要删除的凭据属于Windows凭据还是普通凭据，避免删除错误。 ********************************************************************************/ Dim bRet = "" bRet = Credential.Delete("命令库6.0","normal") TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Credential_图片/Credential_Delete.png)  

---

## 获取凭据

**说明**: 获取凭据，内容以{"username":"用户名","password":"密码"}格式返回，获取到的普通凭据密码会进行加密保护；Windows凭据无法获取到密码  

**原型**: `objCredential = Credential.Get(cName,cType)`  

**参数**:  
- **cName** (True) [string] 默认:"" - 要获取凭据的凭据名  
- **cType** (True) [enum] 默认:"normal" - 凭据所属的类型  

**返回**: objCredential，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取凭据*************************************** 命令原型： objCredential = Credential.Get("","normal") 入参： cName--凭据名。注：该凭据的名称，用于区分每个凭据 cType--凭据类型。注：凭据所属的类型 出参： objCredential--函数调用的输出保存到的变量。 注意事项： 1.windows本地中要存在该凭据，不存在会返回空值。 2.windows凭据是用于windows客户端自用的凭据比如远程连接，共享文件等；普通凭据用于web，第三方的客户端。 3.windows凭据在获取时会密码会返回为空，普通凭据则不会。 ********************************************************************************/ Dim objCredential = "" objCredential = Credential.Get("命令库6.0","normal") TracePrint(objCredential)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Credential_图片/Credential_Get.png)  

---

## 打开CSV文件

**说明**: 打开CSV文件，返回数据表对象  

**原型**: `arrayRet = CSV.Open(sPath,optionArgs)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要打开的CSV文件路径  
- **encoding** (False) [enum] 默认:"auto" - 文件字符集编码  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************************打开CSV文件*********************************** 命令原型： arrayRet = CSV.Open(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,{"encoding":"auto"}) 入参： sPath--要打开的CSV文件路径 出参： arrayRet--命令运行后的结果 注意事项： 字符集编码默认自动识别，还支持GBK（ANSI）、UTF-8、UNICODE、带有BOM的UTF-8，可以切换至可视化界面，在对应属性栏进行选择 ********************************************************************************/ Dim sPath = &#x27;&#x27;&#x27;C:\tempFolder\test.csv&#x27;&#x27;&#x27; arrayRet = CSV.Open(sPath,{"encoding":"auto"}) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/CSV_图片/CSV_Open.png)  

---

## 保存CSV文件

**说明**: 保存CSV文件  

**原型**: `CSV.Save(objData,sPath,optionArgs)`  

**参数**:  
- **objData** (True) [expression] 默认:objData - 要保存的数据表对象，可以是使用 CSV.Open 打开的数据表对象，或数据库返回的数据表对象（会将字段名写为第一行）  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要将CSV文件保存到的路径，传递为空字符串则保存到原始路径，否则另存为到新的位置，如果数据表不是使用 CSV.Open 打开的，这项属性填写空字符串会导致出错  
- **encoding** (False) [enum] 默认:"gbk" - 文件编码，传递为 "ansi" 时使用ANSI编码，传递为 "utf8" 时使用utf-8编码，传递为 "unicode" 时使用 utf-16 编码，传递为 "带有 BOM 的 UTF-8" 时使用 utf-8-sig 编码  

**示例**:  
```
/**********************************保存CSV文件*********************************** 命令原型： CSV.Save(objData,&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,{"encoding": "gbk"}) 入参： objData--要保存的数据表对象 sPath--要将CSV文件保存到的路径 出参： 无 注意事项： 如果保存到的路径文件已存在，会直接覆盖文件，要防止数据丢失的风险；不存在则新建文件保存数据; 字符集编码默认自动识别，还支持GBK（ANSI）、UTF-8、UNICODE、带有BOM的UTF-8，可以切换至可视化界面，在对应属性栏进行选择 ********************************************************************************/ Dim objData= [["name","age"],["Bieber","27"]] Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder\test.csv&#x27;&#x27;&#x27; CSV.Save(objData, sPath,{"encoding": "gbk"})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/CSV_图片/CSV_Save.png)  

---

## 关闭连接

**说明**: 关闭数据库连接  

**原型**: `Database.CloseDB(objDatabase)`  

**参数**:  
- **objDatabase** (True) [expression] 默认:objDatabase - 数据库对象，使用“创建数据库对象“命令（Database.CreateDB）创建的数据库对象  

**示例**:  
```
/*********************************关闭连接*************************************** 命令原型： Database.CloseDB(objDatabase) 入参： objDatabase--数据库对象，使用“创建数据库对象“命令（Database.CreateDB）创建的数据库对象 注意事项： 数据库连接使用完后记得及时关闭 连接命令的charset参数需要与服务端保持一致 **********************************************************************************/ Dim objDatabase Dim ip,port,username,password,db //*********************************MySQL*************************************** // 连接MySQL数据库 ip = "127.0.0.1" port = "3306" username = "root" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "test" objDatabase = Database.CreateDB("MySQL", { "host": ip, "port": port, "user": username, "password": password, "database": db, "charset": "utf8" }) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************PostgreSQL*************************************** // 连接PostGreSQL数据库 ip = "127.0.0.1" port = "5432" username = "root" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "postgres" objDatabase = Database.CreateDB("PostgreSQL", { "host": ip, "port": port, "user": username, "password": password, "database": db }) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************Sqlite3*************************************** // 连接Sqlite3数据库 objDatabase = Database.CreateDB("Sqlite3", {"filepath": &#x27;&#x27;&#x27;D:\工作文档\sqlite\test.db&#x27;&#x27;&#x27;}) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************SQLServer*************************************** // 连接SQLServer数据库 ip = "127.0.0.1" port = "1433" username = "SA" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "TestDB" objDatabase = Database.CreateDB("PostgreSQL", { "host": ip, "port": port, "user": username, "password": password, "database": db, "charset": "utf8" }) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************Oracle*************************************** // 连接Oracle数据库 ip = "127.0.0.1" port = "1521" username = "oracle" password = "OvUZIny9qZUrgJE0ho2tnQ==" service_name = "" sid = "helowin" objDatabase = Database.CreateDB("Oracle", { "host": ip, "port": port, "user": username, "password": password, "service_name": service_name, "sid": sid, "charset": "utf8" }) // 关闭数据库连接 Database.CloseDB(objDatabase)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Database_图片/CloseDB.png)  

---

## 创建数据库对象

**说明**: 创建数据库对象  

**原型**: `objDatabase = Database.CreateDB(dbtype,dbDict)`  

**参数**:  
- **dbtype** (True) [enum] 默认:"MySQL" - 数据库的类型  
- **dbDict** (True) [multiDictionary] 默认:{ "host": "", "port": "3306", "user": "","password": "","database": "","charset": "utf8" } - 连接数据库的配置字典  

**返回**: objDatabase，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************创建数据库对象*************************************** 命令原型： Database.CloseDB(objDatabase) 入参： dbtype--数据库的类型 dbDict--连接数据库的配置字典 objDatabase--命令运行后的结果 注意事项： 数据库连接使用完后记得及时关闭 连接命令的charset参数需要与服务端保持一致 Oracle数据库前置条件： 1.安装oracle客户端(这里假设安装到C:\instantclient_11_2，下面配置需要替换为实际路径) 2.在“环境变量”的“系统变量”中增加： ORACLE_HOME = C:\instantclient_11_2 TNS_ADMIN = C:\instantclient_11_2 NLS_LANG = SIMPLIFIED CHINESE_CHINA.ZHS16GBK 3.修改Path变量，在后面添加 C:\instantclient_11_2 4.在C:\instantclient_11_2 新建一个tnsnames.ora文件，增加自己的数据库别名配置。 示例如下： MyDB=(DESCRIPTION=(ADDRESS= (PROTOCOL = TCP)(HOST= 172.16.1.16)(PORT = 1521)) (CONNECT_DATA=(SERVER=DEDICATED) (SERVICE_NAME=ora10g) ) ) 修改HOST、PORT、SERVICE_NAME与Oracle服务端对应 **********************************************************************************/ Dim objDatabase Dim ip,port,username,password,db //*********************************MySQL*************************************** // 连接MySQL数据库 ip = "127.0.0.1" port = "3306" username = "root" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "test" objDatabase = Database.CreateDB("MySQL", { "host": ip, "port": port, "user": username, "password": password, "database": db, "charset": "utf8" }) //*********************************PostgreSQL*************************************** // 连接PostGreSQL数据库 ip = "127.0.0.1" port = "5432" username = "root" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "postgres" objDatabase = Database.CreateDB("PostgreSQL", { "host": ip, "port": port, "user": username, "password": password, "database": db }) //*********************************Sqlite3*************************************** // 连接Sqlite3数据库 objDatabase = Database.CreateDB("Sqlite3", {"filepath": &#x27;&#x27;&#x27;D:\工作文档\sqlite\test.db&#x27;&#x27;&#x27;}) //*********************************SQLServer*************************************** // 连接SQLServer数据库 ip = "127.0.0.1" port = "1433" username = "SA" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "TestDB" objDatabase = Database.CreateDB("PostgreSQL", { "host": ip, "port": port, "user": username, "password": password, "database": db, "charset": "utf8" }) //*********************************Oracle*************************************** // 连接Oracle数据库 ip = "127.0.0.1" port = "1521" username = "oracle" password = "OvUZIny9qZUrgJE0ho2tnQ==" service_name = "" sid = "helowin" objDatabase = Database.CreateDB("Oracle", { "host": ip, "port": port, "user": username, "password": password, "service_name": service_name, "sid": sid, "charset": "utf8" })
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Database_图片/CreateDB.png)  

---

## 批量执行SQL语句

**说明**: 批量执行SQL语句，返回影响的结果数（select语句暂不支持返回操作影响行数），根据SQL属性进行批量。SQL语句占位符：MySQL、SQLServer、PostgreSQL使用%s，Sqlite3使用?，Oracle使用:1  

**原型**: `iRet = Database.ExecuteBatchSQL(objDatabase ,sql, optionArgs)`  

**参数**:  
- **objDatabase** (True) [expression] 默认:objDatabase - 数据库对象，使用“创建数据库对象“命令（Database.CreateDB）创建的数据库对象  
- **sql** (True) [string] 默认:"" - 增删改SQL语句  
- **args** (False) [expression] 默认:[] - SQL语句参数，遍历参数的二维数组循环执行SQL语句，SQL语句占位符：MySQL 、SQLServer、PostgreSQL都使用%s，Sqlite3使用?，Oracle使用:1  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************批量执行SQL语句*************************************** 命令原型： iRet = Database.ExecuteBatchSQL(objDatabase ,sql, args) 入参： objDatabase--数据库对象，使用“创建数据库对象“命令（Database.CreateDB）创建的数据库对象 sql--增删改SQL语句 args--SQL语句参数，遍历参数的二维数组循环执行SQL语句，SQL语句占位符：MySQL 、SQLServer、PostgreSQL都使用%s，Sqlite3使用?，Oracle使用:1 出参: iRet--命令运行后的结果 注意事项： 数据库连接使用完后记得及时关闭 连接命令的charset参数需要与服务端保持一致 Oracle数据库的SQL与其他数据库略微有区别，使用Oracle时，SQL语句中表名、列名需要使用双引号，SQL尾部不能带‘;’分号 **********************************************************************************/ Dim objDatabase,sql,iRet Dim ip,port,username,password,db //*********************************MySQL*************************************** // 连接MySQL数据库 ip = "127.0.0.1" port = "3306" username = "root" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "test" objDatabase = Database.CreateDB("MySQL", { "host": ip, "port": port, "user": username, "password": password, "database": db, "charset": "utf8" }) // 批量插入 sql = "INSERT IGNORE INTO `test` (`name`, `age`, `aa`, `bb`) VALUES (%s, %s, %s, %s);" iRet = Database.ExecuteBatchSQL(objDatabase ,sql, {"args": [["test3", 12, 2, 2], ["test4", 12, 2, 2]]}) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************PostgreSQL*************************************** // 连接PostGreSQL数据库 ip = "127.0.0.1" port = "5432" username = "root" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "postgres" objDatabase = Database.CreateDB("PostgreSQL", { "host": ip, "port": port, "user": username, "password": password, "database": db }) // 执行批量插入语句 sql = "INSERT INTO test (name, age, id) VALUES (%s, %s, %s);" iRet = Database.ExecuteBatchSQL(objDatabase ,sql, {"args": [["test3", 12, 4],["test4", 13, 5]]}) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************Sqlite3*************************************** // 连接Sqlite3数据库 objDatabase = Database.CreateDB("Sqlite3", {"filepath": &#x27;&#x27;&#x27;D:\工作文档\sqlite\test.db&#x27;&#x27;&#x27;}) // 执行批量插入语句 sql = "INSERT INTO test (name, age, id) VALUES (?, ?, ?);" iRet = Database.ExecuteBatchSQL(objDatabase ,sql, {"args": [["test3", 12, 6],["test4", 13, 7]]}) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************SQLServer*************************************** // 连接SQLServer数据库 ip = "127.0.0.1" port = "1433" username = "SA" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "TestDB" objDatabase = Database.CreateDB("PostgreSQL", { "host": ip, "port": port, "user": username, "password": password, "database": db, "charset": "utf8" }) // 执行批量插入语句 sql = "INSERT INTO Inventory (name, quantity, id) VALUES (%s, %s, %s);" iRet = Database.ExecuteBatchSQL(objDatabase ,sql, {"args": [["test3", 12, 3],["test4", 13, 4]]}) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************Oracle*************************************** // 连接Oracle数据库 ip = "127.0.0.1" port = "1521" username = "oracle" password = "OvUZIny9qZUrgJE0ho2tnQ==" service_name = "" sid = "helowin" objDatabase = Database.CreateDB("Oracle", { "host": ip, "port": port, "user": username, "password": password, "service_name": service_name, "sid": sid, "charset": "utf8" }) // 执行批量插入语句 sql = &#x27;&#x27;&#x27;insert into "student"("id","name","age") VALUES (:1, :1, :1)&#x27;&#x27;&#x27; iRet = Database.ExecuteBatchSQL(objDatabase ,sql, {"args": [[1, "test1", 11],[2, "test2", 22]]}) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Database_图片/ExecuteBatchSQL.png)  

---

## 执行SQL语句

**说明**: 执行SQL语句，返回影响的结果数（select语句暂不支持返回操作影响行数）。SQL语句占位符：MySQL、SQLServer、PostgreSQL使用%s，Sqlite3使用?，Oracle使用:1  

**原型**: `iRet = Database.ExecuteSQL(objDatabase ,sql, optionArgs)`  

**参数**:  
- **objDatabase** (True) [expression] 默认:objDatabase - 数据库对象，使用“创建数据库对象“命令（Database.CreateDB）创建的数据库对象  
- **sql** (True) [string] 默认:"" - 增删改SQL语句  
- **args** (False) [expression] 默认:[] - SQL语句参数，SQL语句占位符：MySQL和SQLServer使用%s，Sqlite3使用?，Oracle使用:1  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************执行SQL语句*************************************** 命令原型： iRet = Database.ExecuteSQL(objDatabase ,sql, args) 入参： objDatabase--数据库对象，使用“创建数据库对象“命令（Database.CreateDB）创建的数据库对象 sql--增删改SQL语句 args--SQL语句参数，遍历参数的二维数组循环执行SQL语句，SQL语句占位符：MySQL 、SQLServer、PostgreSQL都使用%s，Sqlite3使用?，Oracle使用:1 出参: iRet--命令运行后的结果 注意事项： 数据库连接使用完后记得及时关闭 连接命令的charset参数需要与服务端保持一致 Oracle数据库的SQL与其他数据库略微有区别，使用Oracle时，SQL语句中表名、列名需要使用双引号，SQL尾部不能带‘;’分号 **********************************************************************************/ Dim objDatabase,sql,iRet Dim ip,port,username,password,db //*********************************MySQL*************************************** // 连接MySQL数据库 ip = "127.0.0.1" port = "3306" username = "root" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "test" objDatabase = Database.CreateDB("MySQL", { "host": ip, "port": port, "user": username, "password": password, "database": db, "charset": "utf8" }) // 执行插入语句 sql = "INSERT IGNORE INTO `test` (`name`, `age`, `aa`, `bb`) VALUES (%s, %s, %s, %s);" iRet = Database.ExecuteSQL(objDatabase ,sql, {"args": ["test2", 12, 2, 2]}) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************PostgreSQL*************************************** // 连接PostGreSQL数据库 ip = "127.0.0.1" port = "5432" username = "root" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "postgres" objDatabase = Database.CreateDB("PostgreSQL", { "host": ip, "port": port, "user": username, "password": password, "database": db }) // 执行插入语句 sql = "INSERT INTO test (name, age, id) VALUES (%s, %s, %s);" iRet = Database.ExecuteSQL(objDatabase ,sql, {"args": ["test", 12, 3]}) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************Sqlite3*************************************** // 连接Sqlite3数据库 objDatabase = Database.CreateDB("Sqlite3", {"filepath": &#x27;&#x27;&#x27;D:\工作文档\sqlite\test.db&#x27;&#x27;&#x27;}) // 执行插入语句 sql = "INSERT INTO test (name, age, id) VALUES (?, ?, ?);" iRet = Database.ExecuteSQL(objDatabase ,sql, {"args": ["test", 12, 3]}) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************SQLServer*************************************** // 连接SQLServer数据库 ip = "127.0.0.1" port = "1433" username = "SA" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "TestDB" objDatabase = Database.CreateDB("PostgreSQL", { "host": ip, "port": port, "user": username, "password": password, "database": db, "charset": "utf8" }) // 执行插入语句 sql = "INSERT INTO Inventory (name, quantity, id) VALUES (%s, %s, %s);" iRet = Database.ExecuteSQL(objDatabase ,sql, {"args": ["test5", 130, 5]}) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************Oracle*************************************** // 连接Oracle数据库 ip = "127.0.0.1" port = "1521" username = "oracle" password = "OvUZIny9qZUrgJE0ho2tnQ==" service_name = "" sid = "helowin" objDatabase = Database.CreateDB("Oracle", { "host": ip, "port": port, "user": username, "password": password, "service_name": service_name, "sid": sid, "charset": "utf8" }) // 执行批量插入语句 sql = &#x27;&#x27;&#x27;insert into "student"("id","name","age") VALUES (:1, :1, :1)&#x27;&#x27;&#x27; iRet = Database.ExecuteSQL(objDatabase ,sql, {"args": [1, "test1", 11]}) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Database_图片/ExecuteSQL.png)  

---

## 执行全SQL查询

**说明**: 执行查询SQL语句，返回查询到所有结果。SQL语句占位符：MySQL、SQLServer、PostgreSQL使用%s，Sqlite3使用?，Oracle使用:1  

**原型**: `iRet = Database.QueryAll(objDatabase ,sql ,optionArgs)`  

**参数**:  
- **objDatabase** (True) [expression] 默认:objDatabase - 数据库对象，使用“创建数据库对象“命令（Database.CreateDB）创建的数据库对象  
- **sql** (True) [string] 默认:"" - 查询SQL语句  
- **rdict** (False) [boolean] 默认:False - 是否返回字典  
- **args** (False) [expression] 默认:[] - SQL语句参数，SQL语句占位符：MySQL和SQLServer使用%s，Sqlite3使用?，Oracle使用:1  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************执行全SQL查询*************************************** 命令原型： iRet = Database.QueryAll(objDatabase ,sql ,optionArgs) 入参1： objDatabase--数据库对象，使用“创建数据库对象“命令（Database.CreateDB）创建的数据库对象 入参2: sql--查询SQL语句 入参3: rdict--是否返回字典 入参4: args--SQL语句参数，遍历参数的二维数组循环执行SQL语句，SQL语句占位符：MySQL 、SQLServer、PostgreSQL都使用%s，Sqlite3使用?，Oracle使用:1 出参: iRet--命令运行后的结果 注意事项： 数据库连接使用完后记得及时关闭 连接命令的charset参数需要与服务端保持一致 Oracle数据库的SQL与其他数据库略微有区别，使用Oracle时，SQL语句中表名、列名需要使用双引号，SQL尾部不能带‘;’分号 **********************************************************************************/ Dim objDatabase,sql,iRet Dim ip,port,username,password,db //*********************************MySQL*************************************** // 连接MySQL数据库 ip = "127.0.0.1" port = "3306" username = "root" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "test" objDatabase = Database.CreateDB("MySQL", { "host": ip, "port": port, "user": username, "password": password, "database": db, "charset": "utf8" }) // 执行查询语句 sql = "select * from test" iRet = Database.QueryAll(objDatabase ,sql) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************PostgreSQL*************************************** // 连接PostGreSQL数据库 ip = "127.0.0.1" port = "5432" username = "root" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "postgres" objDatabase = Database.CreateDB("PostgreSQL", { "host": ip, "port": port, "user": username, "password": password, "database": db }) // 执行查询语句 sql = "select * from test" iRet = Database.QueryAll(objDatabase ,sql) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************Sqlite3*************************************** // 连接Sqlite3数据库 objDatabase = Database.CreateDB("Sqlite3", {"filepath": &#x27;&#x27;&#x27;D:\工作文档\sqlite\test.db&#x27;&#x27;&#x27;}) // 执行查询语句 sql = "select * from test" iRet = Database.QueryAll(objDatabase ,sql) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************SQLServer*************************************** // 连接SQLServer数据库 ip = "127.0.0.1" port = "1433" username = "SA" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "TestDB" objDatabase = Database.CreateDB("PostgreSQL", { "host": ip, "port": port, "user": username, "password": password, "database": db, "charset": "utf8" }) // 执行查询语句 sql = "select * from Inventory" iRet = Database.QueryAll(objDatabase ,sql) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************Oracle*************************************** // 连接Oracle数据库 ip = "127.0.0.1" port = "1521" username = "oracle" password = "OvUZIny9qZUrgJE0ho2tnQ==" service_name = "" sid = "helowin" objDatabase = Database.CreateDB("Oracle", { "host": ip, "port": port, "user": username, "password": password, "service_name": service_name, "sid": sid, "charset": "utf8" }) // 执行全查询语句 sql = &#x27;&#x27;&#x27;select * from "student"&#x27;&#x27;&#x27; iRet = Database.QueryAll(objDatabase ,sql, {"rdict": false, "args": []}) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Database_图片/QueryAll.png)  

---

## 执行单SQL查询

**说明**: 执行查询SQL语句，返回查询到的第一行结果。SQL语句占位符：MySQL、SQLServer、PostgreSQL使用%s，Sqlite3使用?，Oracle使用:1  

**原型**: `iRet = Database.QueryOne(objDatabase ,sql, optionArgs)`  

**参数**:  
- **objDatabase** (True) [expression] 默认:objDatabase - 数据库对象，使用“创建数据库对象“命令（Database.CreateDB）创建的数据库对象  
- **sql** (True) [string] 默认:"" - 查询SQL语句  
- **rdict** (False) [boolean] 默认:False - 是否返回字典  
- **args** (False) [expression] 默认:[] - SQL语句参数，SQL语句占位符：MySQL和SQLServer使用%s，Sqlite3使用?，Oracle使用:1  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************执行单SQL查询*************************************** 命令原型： iRet = Database.QueryOne(objDatabase ,sql, optionArgs) 入参： objDatabase--数据库对象，使用“创建数据库对象“命令（Database.CreateDB）创建的数据库对象 sql--查询SQL语句 rdict--是否返回字典 args--SQL语句参数，遍历参数的二维数组循环执行SQL语句，SQL语句占位符：MySQL 、SQLServer、PostgreSQL都使用%s，Sqlite3使用?，Oracle使用:1 出参: iRet--命令运行后的结果 注意事项： 数据库连接使用完后记得及时关闭 连接命令的charset参数需要与服务端保持一致 Oracle数据库的SQL与其他数据库略微有区别，使用Oracle时，SQL语句中表名、列名需要使用双引号，SQL尾部不能带‘;’分号 **********************************************************************************/ Dim objDatabase,sql,iRet Dim ip,port,username,password,db //*********************************MySQL*************************************** // 连接MySQL数据库 ip = "127.0.0.1" port = "3306" username = "root" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "test" objDatabase = Database.CreateDB("MySQL", { "host": ip, "port": port, "user": username, "password": password, "database": db, "charset": "utf8" }) // 执行单SQL查询 sql = "select * from test" iRet = Database.QueryOne(objDatabase ,sql) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************PostgreSQL*************************************** // 连接PostGreSQL数据库 ip = "127.0.0.1" port = "5432" username = "root" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "postgres" objDatabase = Database.CreateDB("PostgreSQL", { "host": ip, "port": port, "user": username, "password": password, "database": db }) // 执行单SQL查询 sql = "select * from test" iRet = Database.QueryOne(objDatabase ,sql) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************Sqlite3*************************************** // 连接Sqlite3数据库 objDatabase = Database.CreateDB("Sqlite3", {"filepath": &#x27;&#x27;&#x27;D:\工作文档\sqlite\test.db&#x27;&#x27;&#x27;}) // 执行单SQL查询 sql = "select * from test" iRet = Database.QueryOne(objDatabase ,sql) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************SQLServer*************************************** // 连接SQLServer数据库 ip = "127.0.0.1" port = "1433" username = "SA" password = "rg+d2Wr8T+Dv10iQBk7VUw==" db = "TestDB" objDatabase = Database.CreateDB("PostgreSQL", { "host": ip, "port": port, "user": username, "password": password, "database": db, "charset": "utf8" }) // 执行查询语句 sql = "select * from Inventory" iRet = Database.QueryOne(objDatabase ,sql) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase) //*********************************Oracle*************************************** // 连接Oracle数据库 ip = "127.0.0.1" port = "1521" username = "oracle" password = "OvUZIny9qZUrgJE0ho2tnQ==" service_name = "" sid = "helowin" objDatabase = Database.CreateDB("Oracle", { "host": ip, "port": port, "user": username, "password": password, "service_name": service_name, "sid": sid, "charset": "utf8" }) // 执行查询语句 sql = &#x27;&#x27;&#x27;select * from "student"&#x27;&#x27;&#x27; iRet = Database.QueryOne(objDatabase ,sql, {"rdict": false, "args": []}) TracePrint(iRet) // 关闭数据库连接 Database.CloseDB(objDatabase)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Database_图片/QueryOne.png)  

---

## 从队列取出

**说明**: 当Creator开发者、人机交互Worker用户、Commander任务创建者获得指定队列的权限时，可从该队列中取出数据  

**原型**: `sRet = DataQueue.PullEx(queName)`  

**参数**:  
- **queName** (True) [string] 默认:"" - 指定被取出数据的队列名称，可从Commander中获得  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************从队列取出*************************************** 命令原型： sRet = DataQueue.PullEx(queName) 入参： queName--指定被取出数据的队列名称，可从Commander中获得 出参: sRet--命令运行后的结果 **********************************************************************************/ Dim sRet // 向test队列取出元素 sRet = DataQueue.PullEx("test") Traceprint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/DataQueue_图片/PullEx.png)  

---

## 放入队列

**说明**: 当Creator开发者、人机交互Worker用户、Commander任务创建者获得指定队列的权限时，可将数据放入该队列  

**原型**: `DataQueue.PushEx(queName,item)`  

**参数**:  
- **queName** (True) [string] 默认:"" - 指定被放入数据的队列名称，可从Commander中获得  
- **item** (True) [string] 默认:"" - 放入队列的数据  

**示例**:  
```
/*********************************放入队列*************************************** 命令原型： DataQueue.PushEx(queName,item) 入参： queName--指定被取出数据的队列名称，可从Commander中获得 item--放入队列的数据 注意事项: 请勿放入不存在的队列 **********************************************************************************/ // 向test队列推入"123" DataQueue.PushEx("test","123")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/DataQueue_图片/PushEx.png)  

---

## 增加列

**说明**: 为数据表增加一列数据  

**原型**: `Datatable.AddColumn(dtTable,column,iIndex,objDefaultValue)`  

**参数**:  
- **dtTable** (True) [expression] 默认:objDatatable - 需要进行筛选的数据表  
- **column** (True) [string] 默认:"" - 增加数据列的列名  
- **iIndex** (True) [expression] 默认:null - 增加到数据表中的列的位置，如果为null则增加到最后一列  
- **objDefaultValue** (True) [string] 默认:"" - 填充列的值，可以是数组或者单个值  

**示例**:  
```
/*********************************增加列*************************************** 命令原型： Datatable.AddColumn(dtTable,column,iIndex,objDefaultValue) 入参： dtTable--需要进行筛选的数据表 column--增加数据列的列名 iIndex--增加到数据表中的列的位置，如果为null则增加到最后一列 objDefaultValue--填充列的值，可以是数组或者单个值 **********************************************************************************/ Dim aryData,aryColumns,objDatatable // 定义二维数组 aryData = [["a", 1], ["b", 2], ["c", 3], ["d", 1]] aryColumns = ["letter", "number"] // 构建数据表 objDatatable = Datatable.BuildDataTable(aryData,aryColumns) // 增加列 Datatable.AddColumn(objDatatable,"other",null,"123") TracePrint(objDatatable)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Datatable_图片/AddColumn.png)  

---

## 构建数据表

**说明**: 构建数据表  

**原型**: `objDatatable = Datatable.BuildDataTable(aryData,aryColumns)`  

**参数**:  
- **aryData** (True) [expression] 默认:[] - 要改造数据表的数据，一般是一个二维数组。可以从Excel读取或者使用数据抓取功能进行抓取  
- **aryColumns** (True) [expression] 默认:[] - 数据表列头  

**返回**: objDatatable，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************构建数据表*************************************** 命令原型： objDatatable = Datatable.BuildDataTable(aryData,aryColumns) 入参： aryData--要改造数据表的数据，一般是一个二维数组。可以从Excel读取或者使用数据抓取功能进行抓取 aryColumns--数据表列头 出参: objDatatable--命令运行后的结果 **********************************************************************************/ Dim aryData,aryColumns,objDatatable // 定义二维数组 aryData = [["a", 1], ["b", 2], ["c", 3]] aryColumns = ["letter", "number"] // 构建数据表 objDatatable = Datatable.BuildDataTable(aryData,aryColumns) TracePrint(objDatatable)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Datatable_图片/BuildDataTable.png)  

---

## 比较数据表

**说明**: 比较两个数据表的内容是否一致(只比较数据)  

**原型**: `bRet = Datatable.CompareDataTable(dtSrcTable,dtDistTable)`  

**参数**:  
- **dtSrcTable** (True) [expression] 默认:dtSrcTable - 比较两个数据表的内容是否一致，一个称之为源数据表，另一个称之为目标数据表  
- **dtDistTable** (True) [expression] 默认:dtDistTable - 比较两个数据表的内容是否一致，一个称之为源数据表，另一个称之为目标数据表  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************比较数据表*************************************** 命令原型： bRet = Datatable.CompareDataTable(dtSrcTable,dtDistTable) 入参： dtSrcTable--比较两个数据表的内容是否一致，一个称之为源数据表，另一个称之为目标数据表 dtDistTable--比较两个数据表的内容是否一致，一个称之为源数据表，另一个称之为目标数据表 出参: bRet--命令运行后的结果 **********************************************************************************/ Dim aryData,aryColumns,objDatatable Dim aryData2,aryColumns2,objDatatable2 Dim bRet // 构建源数据表 aryData = [["a", 1], ["b", 2], ["c", 3], ["d", 1]] aryColumns = ["letter", "number"] objDatatable = Datatable.BuildDataTable(aryData,aryColumns) // 构建目标数据表 aryData2 = [["a", 1], ["b", 2]] aryColumns2 = ["letter", "number"] objDatatable2 = Datatable.BuildDataTable(aryData2,aryColumns2) //比较数据表 bRet = Datatable.CompareDataTable(objDatatable,objDatatable2) TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Datatable_图片/CompareDataTable.png)  

---

## 转换列类型

**说明**: 转换数据表列的数据类型  

**原型**: `Datatable.ConvertColumnDataType(dtTable,column,strType,bRaiseExcept,defaultValue)`  

**参数**:  
- **dtTable** (True) [expression] 默认:objDatatable - 需要进行筛选的数据表  
- **column** (True) [string] 默认:"" - 要转换的数据列，可以写单个列名，也可以使用数组形式一次写多个列  
- **strType** (True) [enum] 默认:"float" - 要转换的目标数据类型  
- **bRaiseExcept** (True) [boolean] 默认:False - 当转换失败的时候是否抛出异常  
- **defaultValue** (True) [expression] 默认:null - 当存在转换失败的值且设置为不抛出异常时，该失败的值可被统一转换为填充值  

**示例**:  
```
/*********************************转换列类型*************************************** 命令原型： Datatable.ConvertColumnDataType(dtTable,column,strType,bRaiseExcept,defaultValue) 入参: dtTable--需要进行筛选的数据表 column--要转换的数据列，可以写单个列名，也可以使用数组形式一次写多个列 strType--要转换的目标数据类型 bRaiseExcept--当转换失败的时候是否抛出异常 defaultValue--当存在转换失败的值且设置为不抛出异常时，该失败的值可被统一转换为填充值 **********************************************************************************/ Dim aryData,aryColumns,objDatatable // 构建数据表 aryData = [["a", 1], ["b", 2], ["c", 3], ["d", 1]] aryColumns = ["letter", "number"] objDatatable = Datatable.BuildDataTable(aryData,aryColumns) // 转换列类型 Datatable.ConvertColumnDataType(objDatatable,"number","float",false,null) TracePrint(objDatatable)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Datatable_图片/ConvertColumnDataType.png)  

---

## 复制到剪贴板

**说明**: 将数据表的内容复制到剪贴板  

**原型**: `Datatable.DataTableToClipboard(dtTable)`  

**参数**:  
- **dtTable** (True) [expression] 默认:objDatatable - 需要进行筛选的数据表  

**示例**:  
```
/*********************************复制到剪贴板*************************************** 命令原型： Datatable.DataTableToClipboard(dtTable) 入参: dtTable--需要进行筛选的数据表 **********************************************************************************/ Dim aryData,aryColumns,objDatatable // 构建数据表 aryData = [["a", 1], ["b", 2], ["c", 3], ["d", 1]] aryColumns = ["letter", "number"] objDatatable = Datatable.BuildDataTable(aryData,aryColumns) // 复制到剪贴板 Datatable.DataTableToClipboard(objDatatable)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Datatable_图片/DataTableToClipboard.png)  

---

## 数据表去重

**说明**: 去除数据表中重复的行  

**原型**: `objDatatable = Datatable.DropDuplicatesDataTable(dtTable,aryColumns,strKeep)`  

**参数**:  
- **dtTable** (True) [expression] 默认:objDatatable - 需要进行筛选的数据表  
- **aryColumns** (True) [expression] 默认:[] - 需要去重并且保留的列  
- **strKeep** (True) [enum] 默认:"first" - 需要去重并且保留的列  

**返回**: objDatatable，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************数据表去重*************************************** 命令原型： objDatatable = Datatable.DropDuplicatesDataTable(dtTable,aryColumns,strKeep) 入参: dtTable--需要进行筛选的数据表 aryColumns--需要去重并且保留的列 strKeep--需要去重并且保留的列 出参: objDatatable--命令运行后的结果 **********************************************************************************/ Dim aryData,aryColumns,objDatatable // 构建数据表 aryData = [["a", 1], ["b", 2], ["c", 3], ["d", 1]] aryColumns = ["letter", "number"] objDatatable = Datatable.BuildDataTable(aryData,aryColumns) // 数据表去重 objDatatable = Datatable.DropDuplicatesDataTable(objDatatable,["number"],"last") TracePrint(objDatatable)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Datatable_图片/DropDuplicatesDataTable.png)  

---

## 获取数据表列名

**说明**: 以数组形式返回数据表的所有列名  

**原型**: `arrayColumns = Datatable.GetColumns(dtTable)`  

**参数**:  
- **dtTable** (True) [expression] 默认:objDatatable - 需要进行筛选的数据表  

**返回**: arrayColumns，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取数据表列名*************************************** 命令原型： arrayColumns = Datatable.GetColumns(dtTable) 入参: dtTable--需要进行筛选的数据表 出参: arrayColumns--命令运行后的结果 **********************************************************************************/ Dim aryData,aryColumns,objDatatable // 构建数据表 aryData = [["a", 1], ["b", 2], ["c", 3], ["d", 1]] aryColumns = ["letter", "number"] objDatatable = Datatable.BuildDataTable(aryData,aryColumns) // 获取数据表列名 arrayColumns = Datatable.GetColumns(objDatatable) TracePrint(arrayColumns)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Datatable_图片/GetColumns.png)  

---

## 转换为数组

**说明**: 将数据表转换为数组  

**原型**: `objDatatable = Datatable.GetDataTableByArray(dtTable,hasHead)`  

**参数**:  
- **dtTable** (True) [expression] 默认:objDatatable - 需要进行筛选的数据表  
- **hasHead** (True) [boolean] 默认:False - 转为数组之后是否包含表头  

**返回**: objDatatable，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************转换为数组*************************************** 命令原型： objDatatable = Datatable.GetDataTableByArray(dtTable,hasHead) 入参: dtTable--需要进行筛选的数据表 hasHead--转为数组之后是否包含表头 出参: objDatatable--命令运行后的结果 **********************************************************************************/ Dim aryData,aryColumns,objDatatable // 构建数据表 aryData = [["a", 1], ["b", 2], ["c", 3], ["d", 1]] aryColumns = ["letter", "number"] objDatatable = Datatable.BuildDataTable(aryData,aryColumns) // 转换为数组 objDatatable = Datatable.GetDataTableByArray(objDatatable,true) TracePrint(objDatatable)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Datatable_图片/GetDataTableByArray.png)  

---

## 获取行列数

**说明**: 以数组的形式返回数据表的行列数，形式为 [行总数，列总数]  

**原型**: `arrayShape = Datatable.GetDataTableShape(dtTable)`  

**参数**:  
- **dtTable** (True) [expression] 默认:objDatatable - 需要进行筛选的数据表  

**返回**: arrayShape，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取行列数*************************************** 命令原型： arrayShape = Datatable.GetDataTableShape(dtTable) 入参: dtTable--需要进行筛选的数据表 出参: arrayShape--命令运行后的结果 **********************************************************************************/ Dim aryData,aryColumns,objDatatable,arrayShape // 构建数据表 aryData = [["a", 1], ["b", 2], ["c", 3], ["d", 1]] aryColumns = ["letter", "number"] objDatatable = Datatable.BuildDataTable(aryData,aryColumns) // 获取行列数 arrayShape = Datatable.GetDataTableShape(objDatatable) TracePrint(arrayShape)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Datatable_图片/GetDataTableShape.png)  

---

## 合并数据表

**说明**: 将两个数据表按照指定的连接方式合并  

**原型**: `dtTable = Datatable.MergeDataTable(dtLeftTable,dtRightTable,strHow,strLeftKey,strRightKey,bSort)`  

**参数**:  
- **dtLeftTable** (True) [expression] 默认:leftTable - 要合并的数据表1  
- **dtRightTable** (True) [expression] 默认:rightTable - 要合并的数据表2  
- **strHow** (True) [enum] 默认:"inner" - 合并两表的连接方式  
- **strLeftKey** (True) [string] 默认:"" - 左表用来做为合并依据的列名  
- **strRightKey** (True) [string] 默认:"" - 右表用来做为合并依据的列名  
- **bSort** (True) [boolean] 默认:False - 设置为True时表示合并时会根据给定的列值(也就是前面的left_on这种指定的列的值)来进行排序后再输出  

**返回**: dtTable，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************合并数据表*************************************** 命令原型： dtTable = Datatable.MergeDataTable(dtLeftTable,dtRightTable,strHow,strLeftKey,strRightKey,bSort) 入参: dtLeftTable--要合并的数据表1 dtRightTable--要合并的数据表2 strHow--合并两表的连接方式 strLeftKey--左表用来做为合并依据的列名 strRightKey--右表用来做为合并依据的列名 bSort--设置为True时表示合并时会根据给定的列值(也就是前面的left_on这种指定的列的值)来进行排序后再输出 出参: dtTable--命令运行后的结果 **********************************************************************************/ Dim aryData,aryColumns,objDatatable Dim aryData2,aryColumns2,objDatatable2 Dim dtTable // 构建数据表1 aryData = [["a", 1], ["b", 2], ["c", 3], ["d", 1]] aryColumns = ["letter", "number"] objDatatable = Datatable.BuildDataTable(aryData,aryColumns) // 构建数据表2 aryData2 = [["a", 19], ["b", 20]] aryColumns2 = ["letter", "age"] objDatatable2 = Datatable.BuildDataTable(aryData2,aryColumns2) // 合并数据表--外连接 dtTable = Datatable.MergeDataTable(objDatatable,objDatatable2,"outer","letter","letter",false) TracePrint(dtTable) // 合并数据表--内连接 dtTable = Datatable.MergeDataTable(objDatatable,objDatatable2,"inner","letter","letter",false) TracePrint(dtTable) // 合并数据表--左连接 dtTable = Datatable.MergeDataTable(objDatatable,objDatatable2,"left","letter","letter",false) TracePrint(dtTable) // 合并数据表--右连接 dtTable = Datatable.MergeDataTable(objDatatable,objDatatable2,"right","letter","letter",false) TracePrint(dtTable)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Datatable_图片/MergeDataTable.png)  

---

## 修改列名

**说明**: 修改数据表的列名  

**原型**: `Datatable.ModfiyColumns(dtTable,aryColumns)`  

**参数**:  
- **dtTable** (True) [expression] 默认:objDatatable - 需要进行筛选的数据表  
- **aryColumns** (True) [expression] 默认:[] - 将使用此数组替换原数据列名  

**示例**:  
```
/*********************************修改列名*************************************** 命令原型： Datatable.ModfiyColumns(dtTable,aryColumns) 入参: dtTable--需要进行筛选的数据表 aryColumns--将使用此数组替换原数据列名 **********************************************************************************/ Dim aryData,aryColumns,objDatatable // 构建数据表 aryData = [["a", 1], ["b", 2], ["c", 3], ["d", 1]] aryColumns = ["letter", "number"] objDatatable = Datatable.BuildDataTable(aryData,aryColumns) // 修改列名 Datatable.ModfiyColumns(objDatatable,["l", "n"]) TracePrint(objDatatable)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Datatable_图片/ModfiyColumns.png)  

---

## 数据筛选

**说明**: 使用表达式对数据表的数据进行筛选  

**原型**: `objDatatable = Datatable.QueryDataTable(dtTable,strQueryExpress)`  

**参数**:  
- **dtTable** (True) [expression] 默认:objDatatable - 需要进行筛选的数据表  
- **strQueryExpress** (True) [string] 默认:"" - 筛选数据的条件，如：column.str.contains(&#x27;Laiye RPA&#x27;) and column1>1,代表列&#x27;column&#x27;包含&#x27;Laiye RPA&#x27;，并且列&#x27;column1&#x27;大于1的行数据  

**返回**: objDatatable，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************数据筛选*************************************** 命令原型： objDatatable = Datatable.QueryDataTable(dtTable,strQueryExpress) 入参: dtTable--需要进行筛选的数据表 strQueryExpress--筛选数据的条件，如：column.str.contains(&#x27;Laiye RPA&#x27;) and column1>1,代表列&#x27;column&#x27;包含&#x27;Laiye RPA&#x27;，并且列&#x27;column1&#x27;大于1的行数据 注意事项: 如果条件需要使用变量，建议先使用可视化设置筛选条件，再到源代码中修改，拼接变量 **********************************************************************************/ Dim aryData,aryColumns,objDatatable,iNum // 构建数据表 aryData = [["a", 1], ["b", 2], ["c", 3]] aryColumns = ["letter", "number"] objDatatable = Datatable.BuildDataTable(aryData,aryColumns) // 数据筛选，条件为固定字符串 objDatatable2 = Datatable.QueryDataTable(objDatatable,"number>1") TracePrint(objDatatable2) // 数据筛选，条件拼接变量 iNum = 2 objDatatable3 = Datatable.QueryDataTable(objDatatable,"number>"&iNum) TracePrint(objDatatable3)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Datatable_图片/QueryDataTable.png)  

---

## 选择数据列

**说明**: 选择数据表中的数据列，返回一个新的数据表  

**原型**: `objDatatable = Datatable.SelectDataTableColumns(dtTable,aryColumns)`  

**参数**:  
- **dtTable** (True) [expression] 默认:objDatatable - 需要进行筛选的数据表  
- **aryColumns** (True) [expression] 默认:[] - 需要保留的列  

**返回**: objDatatable，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************选择数据列*************************************** 命令原型： objDatatable = Datatable.SelectDataTableColumns(dtTable,aryColumns) 入参: dtTable--需要进行筛选的数据表 aryColumns--需要保留的列 出参: objDatatable--命令运行后的结果 **********************************************************************************/ Dim aryData,aryColumns,objDatatable // 构建数据表 aryData = [["a", 1], ["b", 2], ["c", 3]] aryColumns = ["letter", "number"] objDatatable = Datatable.BuildDataTable(aryData,aryColumns) // 选择数据列 objDatatable = Datatable.SelectDataTableColumns(objDatatable,["letter"]) TracePrint(objDatatable)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Datatable_图片/SelectDataTableColumns.png)  

---

## 数据切片

**说明**: 数据切片  

**原型**: `objDatatable = Datatable.SliceDataTable(dtTable,aryRows,aryColumns)`  

**参数**:  
- **dtTable** (True) [expression] 默认:objDatatable - 需要切片的源数据表  
- **aryRows** (True) [expression] 默认:[] - 数据的行切片，数组形式，第一个元素代表起始行，后一个代表截止行号  
- **aryColumns** (True) [expression] 默认:[] - 数据表列头  

**返回**: objDatatable，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************数据切片*************************************** 命令原型： objDatatable = Datatable.SliceDataTable(dtTable,aryRows,aryColumns) 入参: dtTable--需要进行筛选的数据表 aryRows--数据的行切片，数组形式，第一个元素代表起始行，后一个代表截止行号 aryColumns--数据表列头 出参: objDatatable--命令运行后的结果 **********************************************************************************/ Dim aryData,aryColumns,objDatatable // 构建数据表 aryData = [["a", 1], ["b", 2], ["c", 3]] aryColumns = ["letter", "number"] objDatatable = Datatable.BuildDataTable(aryData,aryColumns) // 数据切片 objDatatable = Datatable.SliceDataTable(objDatatable,[0,1],["letter"]) TracePrint(objDatatable)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Datatable_图片/SliceDataTable.png)  

---

## 数据表排序

**说明**: 对数据表的指定列进行排序  

**原型**: `dtTable = Datatable.SortDataTable(dataTable,columns,bAscSort)`  

**参数**:  
- **dataTable** (True) [expression] 默认:objDatatable - 需要进行筛选的数据表  
- **columns** (True) [string] 默认:"" - 填入需要排序的数据列头，多列排序使用数组，填入 ["列1","列2","列3"] 则同时排序列头为"列1","列2","列3"的列  
- **bAscSort** (True) [boolean] 默认:True - 是否进行升序排序，选择否则进行降序排序  

**返回**: dtTable，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************数据表排序*************************************** 命令原型： dtTable = Datatable.SortDataTable(dataTable,columns,bAscSort) 入参: dataTable--需要进行筛选的数据表 columns--填入需要排序的数据列头，多列排序使用数组，填入["列1","列2","列3"]则同时排序列头为"列1","列2","列3"的列 bAscSort--是否进行升序排序，选择否则进行降序排序 出参: dtTable--命令运行后的结果 **********************************************************************************/ Dim aryData,aryColumns,objDatatable,dtTable // 构建数据表 aryData = [["a", 1], ["b", 2], ["c", 3], ["d", 1]] aryColumns = ["letter", "number"] objDatatable = Datatable.BuildDataTable(aryData,aryColumns) // 数据表排序 dtTable = Datatable.SortDataTable(objDatatable,"number",false) TracePrint(dtTable)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Datatable_图片/SortDataTable.png)  

---

## 输入对话框

**说明**: 弹出输入对话框，返回用户在对话框中输入的内容  

**原型**: `sRet = Dialog.InputBox(sText,sTitle,sDefault,bNumberOnly)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 对话框中显示的消息提示内容  
- **sTitle** (True) [string] 默认:"Laiye RPA" - 对话框标题  
- **sDefault** (True) [string] 默认:"" - 输入对话框的默认文字内容  
- **bNumberOnly** (True) [boolean] 默认:False - 输入对话框的默认文字内容  
- **### 返回结果** () [] 默认: -   
- **sRet，将命令运行后的结果赋值给此变量。** () [] 默认: -   
- **### 运行实例** () [] 默认: -   

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Dialog_图片/Dialog_InputBox.png)  

---

## 消息框

**说明**: 弹出消息提示对话框，返回用户点击的按钮  

**原型**: `iRet = Dialog.MsgBox(sText,sTitle,iStyle,sCommand,iTimeout)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 对话框中显示的消息提示内容  
- **sTitle** (True) [string] 默认:"Laiye RPA" - 对话框标题  
- **iStyle** (True) [enum] 默认:0 - 对话框的按钮样式  
- **sCommand** (True) [enum] 默认:1 - 对话框显示的图标  
- **iTimeout** (True) [number] 默认:0 - 超时时间（毫秒），0代表不使用超时时间  
- **### 返回结果** () [] 默认: -   
- **iRet，将命令运行后的结果赋值给此变量。** () [] 默认: -   
- **### 运行实例** () [] 默认: -   

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Dialog_图片/Dialog_MsgBox.png)  

---

## 消息通知

**说明**: 弹出消息通知对话框  

**原型**: `Dialog.Notify(sMessage, sTitle, iIcon)`  

**参数**:  
- **sMessage** (True) [string] 默认:"" - 通知对话框的消息内容  
- **sTitle** (True) [string] 默认:"Laiye RPA" - 对话框标题  
- **iIcon** (True) [enum] 默认:0 - 对话框图标  
- **### 运行实例** () [] 默认: -   

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Dialog_图片/Dialog_Notify.png)  

---

## 打开文件对话框

**说明**: 弹出打开文件对话框  

**原型**: `sRet = Dialog.OpenFile(sDefaultPath,sFilter,sTitle)`  

**参数**:  
- **sDefaultPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 对话框显示时定位到的默认目录，传递为 null 则定位到上次操作的目录  
- **sFilter** (True) [string] 默认:"文本文档 (txt、log) - .txt; .log  
- **sTitle** (True) [string] 默认:"Laiye RPA" - 对话框标题  
- **### 返回结果** () [] 默认: -   
- **sRet，将命令运行后的结果赋值给此变量。** () [] 默认: -   
- **### 运行实例** () [] 默认: -   

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Dialog_图片/Dialog_OpenFile.png)  

---

## 打开文件对话框 [多选]

**说明**: 弹出打开文件对话框，对话框中可以选择多个文件  

**原型**: `arrRet = Dialog.OpenFiles(sDefaultPath,sFilter,sTitle)`  

**参数**:  
- **sDefaultPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 对话框显示时定位到的默认目录，传递为 null 则定位到上次操作的目录  
- **sFilter** (True) [string] 默认:"文本文档 (txt、log) - .txt; .log  
- **sTitle** (True) [string] 默认:"Laiye RPA" - 对话框标题  
- **### 返回结果** () [] 默认: -   
- **arrRet，将命令运行后的结果赋值给此变量。** () [] 默认: -   
- **### 运行实例** () [] 默认: -   

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Dialog_图片/Dialog_OpenFiles.png)  

---

## 保存文件对话框

**说明**: 弹出保存文件对话框  

**原型**: `sRet = Dialog.SaveFile(sDefaultPath,sFilter,sTitle)`  

**参数**:  
- **sDefaultPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 对话框显示时定位到的默认目录，传递为 null 则定位到上次操作的目录  
- **sFilter** (True) [string] 默认:"文本文档 (txt、log) - .txt; .log  
- **sTitle** (True) [string] 默认:"Laiye RPA" - 对话框标题  
- **### 返回结果** () [] 默认: -   
- **sRet，将命令运行后的结果赋值给此变量。** () [] 默认: -   
- **### 运行实例** () [] 默认: -   

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Dialog_图片/Dialog_SaveFile.png)  

---

## 选项圆盘

**说明**: 弹出选项圆盘，圆盘中最多支持8个选项，每个选项代表一个索引值(从0开始依次至7)，选择某个选项后选项圆盘会关闭并返回其索引值，按Esc键可直接关闭选项圆盘，此时返回的索引值为 -1。流程中可以基于索引值进行后续的逻辑分支设计，比如索引值为1时，流程往子流程A执行，为5时，流程往子流程B执行等等  

**原型**: `iRet = Dialog.TurnTable(arrOption,tips)`  

**参数**:  
- **arrOption** (True) [expression] 默认:["京","沪","津","闽","湘","粤","港","澳"] - 选项配置为包含1-8个元素的数组(如果超过8个元素，后面的内容会被忽略)，当数组为 ["京","沪","津","闽","湘","粤","港","澳"] 的格式时，即数组的元素为字符串，圆盘选项名称显示为数组元素的前4个汉字或英文字母；当数组为 [ ["京", "第一个选项"] , ["沪", "第二个选项"] , ["粤", "第三个选项"] ]) 的格式时，即数组的元素为一维数组，鼠标移动到选项名称上会显示提示文字，如“第一个选项”  
- **tips** (True) [string] 默认:"" - 鼠标移动到圆盘的中心时，显示的文字提示，默认为空  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************选项圆盘*************************************** 命令原型： Dialog.TurnTable(["京","沪","津","闽","湘","粤","港","澳"],"") 入参： arrOption -- 选项配置为包含1-8个元素的数组(如果超过8个元素，后面的内容会被忽略)，当数组为["京","沪","津","闽","湘","粤","港","澳"] 的格式时，即数组的元素为字符串，圆盘选项名称显示为数组元素的前4个汉字或英文字母；当数组为 [["京", "第一个选项"], ["沪", "第二个选项"], ["粤", "第三个选项"]]) 的格式时，即数组的元素为一维数组，鼠标移动到选项名称上会显示提示文字，如“第一个选项” tips -- 鼠标移动到圆盘的中心时，显示的文字提示，默认为空 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 输出结果为用户选择选项的下标 **********************************************************************************/ Dim iRet iRet = Dialog.TurnTable(["京","沪","津","闽","湘","粤","港","澳"],"这里是提示文字") TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Dialog_图片/Dialog_TurnTable.png)  

---

## 自定义对话框

**说明**: 打开自定义对话框  

**原型**: `dictRet = Dialog.UDFDialog(strTitle,dictUIFilePath,sDefaultJson,optionArgs)`  

**参数**:  
- **strTitle** (True) [string] 默认:"" - 显示对话框的标题  
- **dictUIFilePath** (True) [path] 默认:"" - 设计自定义的表单结构  
- **sDefaultJson** (True) [expression] 默认:{ } - 设置自定义表单控件默认值，采用JSON格式，可传入变量或表达式  
- **iTimeout** (False) [number] 默认:0 - 对话框显示时间，默认为0则永远显示  
- **strTimoutClick** (False) [enum] 默认:"ok" - 对话框到达显示时间之后会触发点击的按钮，当显示时间为0时，此项无论是何值都不会有任何效果  
- **bInterruptTimeout** (False) [boolean] 默认:True - 当用户有对话框表单操作时，则中断超时自动点击按钮操作，默认为是  

**返回**: dictRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************自定义对话框*************************************** 命令原型： Dialog.UDFDialog("","",{},{"iTimeout": 0, "strTimoutClick": "ok", "bInterruptTimeout": true}) 入参： strTitle -- 显示对话框的标题 dictUIFilePath -- 设计自定义的表单结构 sDefaultJson -- 设置自定义表单控件默认值，采用JSON格式，可传入变量或表达式 iTimeout -- 对话框显示时间，默认为0则永远显示 strTimoutClick -- 对话框到达显示时间之后会触发点击的按钮，当显示时间为0时，此项无论是何值都不会有任何效果 bInterruptTimeout -- 当用户有对话框表单操作时，则中断超时自动点击按钮操作，默认为是 出参： dictRet -- 将命令运行后的结果赋值给此变量 注意事项： 输出结果为用户自定的对话框和内容的字典 **********************************************************************************/ Dim dictRet dictRet = Dialog.UDFDialog("这个是对话框标题",@res"1654582428966.json",{"这个是文本框标题":"这个是默认值"},{"iTimeout": 0, "strTimoutClick": "ok", "bInterruptTimeout": true}) TracePrint(dictRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Dialog_图片/Dialog_UDFDialog.png)  

---

## 激活Excel工作簿窗口

**说明**: 激活指定的Excel工作簿窗口  

**原型**: `Excel.ActiveBook(objExcelWorkBook)`  

**参数**:  
- **objExcelWorkBook** (True) [expression] 默认:objExcelWorkBook - 使用 "打开Excel工作簿"命令（Excel.OpenExcel） 或 "绑定Excel工作簿" 命令（Excel.BindBook）返回的工作簿对象  

**示例**:  
```
/*********************************激活Excel工作簿窗口*************************************** 命令原型： Excel.ActiveBook(objExcelWorkBook) 入参： objExcelWorkBook--Excel工作簿对象（使用 "打开Excel"命令（Excel.OpenExcel） 打开的工作簿或使用"绑定Excel"命令（Excel.BindBook）绑定的工作簿对象）。 注意事项： 该命令不能单独使用，需配合 "打开Excel"命令（Excel.OpenExcel） 或"绑定Excel"命令（Excel.BindBook）一起使用才能正常使用，单独使用则会报错。 **********************************************************************************/ Dim objExcelWorkBook = "" Dim objExcelWorkBook2 = "" objExcelWorkBook = Excel.OpenExcel(@res"测试.xlsx",True,"Excel","","") objExcelWorkBook2 = Excel.OpenExcel(@res"空白文件.xlsx",True,"Excel","","") Excel.ActiveBook(objExcelWorkBook) TracePrint "激活Excel工作簿窗口：已将Excel对象&#x27;objExcelWorkBook&#x27;激活并将窗口置顶" Excel.CloseExcel(objExcelWorkBook,False) Excel.CloseExcel(objExcelWorkBook2,False)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Excel_图片/Excel_ActiveBook.png)  

---

## 追加写入文件

**说明**: 指定一个文件路径，将内容写入到路径对应文件的末尾，不会覆盖文件中原有的内容  

**原型**: `File.Append(sPath,sText,sCharset)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需要写入的文件路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  
- **sText** (True) [string] 默认:"" - 写入的文件内容  
- **sCharset** (True) [enum] 默认:"gbk" - 文件编码，传递为 "auto" 时自动判断编码，传递为 "ansi" 时使用ANSI编码，传递为 "utf8" 时使用utf-8编码，传递为 "unicode" 时使用 utf-16 编码  

**示例**:  
```
/********************************追加写入文件************************************ 命令原型： File.Append(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"","gbk") 入参： sPath--需要写入的文件路径 sTest--写入的文件内容 出参： 无 注意事项： 默认字符集编码为gbk，可以切换至可视化界面，在对应属性栏选择其他字符集编码， 如果指定路径的文件不存在，会自动新建文件后追加写入 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder\data.txt&#x27;&#x27;&#x27; Dim sText="RPA" File.Append(sPath,sText,"gbk")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_Append.png)  

---

## 获取名称

**说明**: 指定一个文件或文件夹路径，获取路径对应文件或文件夹的名称  

**原型**: `sName = File.BaseName(sPath,bExt)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 文件或文件夹的路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  
- **bExt** (True) [boolean] 默认:False - 包含文件名称的扩展名，此属性只有在路径为文件时才会生效  

**返回**: sName，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************************获取名称************************************** 命令原型： sName = File.BaseName(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,false) 入参： sPath--文件或文件夹的路径 出参： sName--命令运行后的结果 注意事项： 默认不包含文件名称的扩展名，可以切换至可视化界面，在对应属性栏选择是，即包含文件扩展名 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder\data.txt&#x27;&#x27;&#x27; sName = File.BaseName(sPath,True) TracePrint(sName)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_BaseName.png)  

---

## 压缩文件或文件夹

**说明**: 将指定的文件或文件夹压缩成.zip文件，若存在同名文件则直接覆盖  

**原型**: `sRet = File.Compress(sPath,sZipPath,optionArgs)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需压缩的文件或文件夹路径。支持输入字符串和数组类型，输入字符串表示单文件或文件夹，输入数组代表的是文件路径的集合（比如选中了相同父路径下的多个文件）  
- **sZipPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 压缩后的zip文件路径。如@res"abc.zip"  
- **sPassword** (False) [string] 默认:"" - 设置压缩文件密码  
- **sAlgorithm** (False) [enum] 默认:"standard" - 选择压缩算法的级别  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************压缩文件或文件夹********************************* 命令原型： sRet = File.Compress(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,{"sPassword":&#x27;&#x27;,"sAlgorithm":"standard"}) 入参： sPath--需压缩的文件或文件夹路径 sZipPath--压缩后的zip文件路径 sPassword-设置压缩文件密码 出参： sRet--命令运行后的结果 注意事项： 压缩文件密码、压缩算法级别都为可选项，可以切换至可视化界面，在对应属性栏进行设置和选择 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder\someFiles&#x27;&#x27;&#x27; Dim sZipPath=&#x27;&#x27;&#x27;C:\tempFolder\all.zip&#x27;&#x27;&#x27; Dim sPassword=&#x27;1234&#x27; sRet = File.Compress(sPath,sZipPath,{"sPassword":sPassword,"sAlgorithm":"standard"}) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_Compress.png)  

---

## 复制文件

**说明**: 指定一个文件路径，将该文件复制到指定路径下  

**原型**: `File.CopyFile(sPathSrc,sPathDst,bOverWrite)`  

**参数**:  
- **sPathSrc** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 被复制的文件路径，可为绝对路径，也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  
- **sPathDst** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 复制到的文件夹路径，可为绝对路径，也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  
- **bOverWrite** (True) [boolean] 默认:False - 复制遇到同名文件时可选择是否替换，默认为否  

**示例**:  
```
/***********************************复制文件************************************* 命令原型： File.CopyFile(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,false) 入参： sPathSrc--被复制文件路径 sPathDst--复制到的文件夹路径 bOverWrite-复制遇到同名文件时可选择是否替换 出参： 无 注意事项： 默认复制遇到同名文件时不进行替换，可以切换至可视化界面，在对应属性栏选择替换 ********************************************************************************/ Dim sPathSrc=&#x27;&#x27;&#x27;C:\tempFolder\data.txt&#x27;&#x27;&#x27; Dim sPathDst=&#x27;&#x27;&#x27;C:\tempFolder\copyFolder&#x27;&#x27;&#x27; Dim bOverWrite=false File.CopyFile(sPathSrc,sPathDst,bOverWrite)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_CopyFile.png)  

---

## 复制文件夹

**说明**: 指定一个文件夹路径，将该文件夹下的所有内容复制至指定路径下  

**原型**: `File.CopyFolder(sPathSrc,sPathDst,bOverWrite)`  

**参数**:  
- **sPathSrc** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 被复制的文件夹路径，可为绝对路径，也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  
- **sPathDst** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 复制到的文件夹路径，可为绝对路径，也可使用@res"路径"形式表示当前流程res文件夹下的路径，如果不存在则自动创建该文件夹。路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  
- **bOverWrite** (True) [boolean] 默认:False - 复制遇到同名文件夹时可选择是否替换，默认为否  

**示例**:  
```
/***********************************复制文件夹************************************ 命令原型： File.CopyFolder(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,false) 入参： sPathSrc--被复制的文件夹路径 sPathDst--复制到的文件夹路径 bOverWrite-复制遇到同名文件时可选择是否替换 出参： 无 注意事项： 默认复制遇到同名文件夹时不进行替换，可以切换至可视化界面，在对应属性栏选择替换 ********************************************************************************/ Dim sPathSrc=&#x27;&#x27;&#x27;C:\tempFolder\dataFolder&#x27;&#x27;&#x27; Dim sPathDst=&#x27;&#x27;&#x27;C:\tempFolder\copyFolder&#x27;&#x27;&#x27; Dim bOverWrite=false File.CopyFolder(sPathSrc,sPathDst,bOverWrite)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_CopyFolder.png)  

---

## 创建文件夹

**说明**: 按指定的路径创建文件夹  

**原型**: `File.CreateFolder(sPath)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需要创建的文件夹路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  

**示例**:  
```
/***********************************创建文件夹*********************************** 命令原型： File.CreateFolder(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： sPath--需要创建的文件夹路径 出参： 无 注意事项： 建议先判断该路径对应的文件夹是否存在，如果存在先删除该文件夹再创建，否则会报错 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder\test&#x27;&#x27;&#x27; File.CreateFolder(sPath)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_CreateFolder.png)  

---

## 解压zip文件

**说明**: 将指定的.zip文件解压到指定文件夹，若存在同名文件则直接覆盖  

**原型**: `arrRet = File.Decompression(sPath,sZipPath,optionArgs)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需解压的文件路径  
- **sZipPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 解压到指定文件夹  
- **sPassword** (False) [string] 默认:"" - 设置解压文件密码  

**返回**: arrRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/***********************************解压zip文件*********************************** 命令原型： arrRet = File.Decompression(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,{"sPassword":&#x27;&#x27;}) 入参： sPath--需解压的文件路径 sZipPath--解压到指定文件夹 sPassword-设置解压文件密码 出参： sRet--命令运行后的结果 注意事项： 设置解压密码为可选项，可以切换至可视化界面，在对应属性栏进行设置 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder\all.zip&#x27;&#x27;&#x27; Dim sZipPath=&#x27;&#x27;&#x27;C:\tempFolder&#x27;&#x27;&#x27; Dim sPassword="1234" arrRet = File.Decompression(sPath,sZipPath,{"sPassword":sPassword}) TracePrint(arrRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_Decompression.png)  

---

## 删除文件

**说明**: 指定一个文件路径，删除路径对应文件  

**原型**: `File.DeleteFile(sPath)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需要删除的文件路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  

**示例**:  
```
/************************************删除文件************************************* 命令原型： File.DeleteFile(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： sPath--需要删除的文件路径 出参： 无 注意事项： 建议先判断该路径对应的文件是否存在，如果存在删除文件，不存在则会报错 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder\all.zip&#x27;&#x27;&#x27; File.DeleteFile(sPath)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_DeleteFile.png)  

---

## 删除文件夹

**说明**: 指定一个文件夹路径，删除路径对应的文件夹  

**原型**: `File.DeleteFolder(sPath)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需要删除的文件夹路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  

**示例**:  
```
/***********************************删除文件夹************************************ 命令原型： File.DeleteFolder(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： sPath--需要删除的文件夹路径 出参： 无 注意事项： 建议先判断该路径对应的文件夹是否存在，如果存在删除文件夹，不存在则会报错 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder\test&#x27;&#x27;&#x27; File.DeleteFolder(sPath)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_DeleteFolder.png)  

---

## 获取文件或文件夹列表

**说明**: 指定一个文件夹路径，获取路径对应的文件夹内的文件或文件夹列表，列表按英文字母A-Z的顺序排序，以数组的形式返回  

**原型**: `arrayRet = File.DirFileOrFolder(sPath,sFilter,optionArgs)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需要获取列表的文件夹路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  
- **sFilter** (True) [enum] 默认:"fileandfolder" - 获取到的列表内容  
- **hasPath** (False) [boolean] 默认:True - 选择是，返回列表包含的每一项是文件或文件夹的绝对路径，选择否，返回列表包含的每一项是文件或文件夹的名称  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*****************************获取文件或文件夹列表********************************* 命令原型： arrayRet = File.DirFileOrFolder(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"fileandfolder",{"hasPath":true}) 入参： sPath--需要获取列表的文件夹路径 sFilter-获取到的列表内容 出参： sName--命令运行后的结果 注意事项： 获取到的列表内容以及是否返回全路径，可以切换至可视化界面，在对应属性栏选择 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder&#x27;&#x27;&#x27; Dim sFilter="fileandfolder" arrayRet = File.DirFileOrFolder(sPath,sFilter,{"hasPath":true}) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_DirFileOrFolder.png)  

---

## 获取文件扩展名

**说明**: 指定一个文件路径，获取路径对应文件的文件扩展名  

**原型**: `sNameExtension = File.ExtensionName(sPath)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 文件的路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  

**返回**: sNameExtension，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/********************************获取文件扩展名*********************************** 命令原型： sNameExtension = File.ExtensionName(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： sPath--文件的路径 出参： sNameExtension--命令运行后的结果 注意事项： 建议先判断该路径对应的文件是否存在，如果存在获取文件扩展名，不存在则会报错 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder\hello.txt&#x27;&#x27;&#x27; sNameExtension = File.ExtensionName(sPath) TracePrint(sNameExtension)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_ExtensionName.png)  

---

## 判断文件是否存在

**说明**: 指定一个文件路径，判断路径对应的文件是否存在。返回布尔值，True表示存在，False表示不存在  

**原型**: `bRet = File.FileExists(sPath)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需要判断的文件路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************判断文件是否存在********************************* 命令原型： bRet = File.FileExists(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： sPath--需要判断的文件路径 出参： bRet--命令运行后的结果 注意事项： 无 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder\hello.txt&#x27;&#x27;&#x27; bRet = File.FileExists(sPath) TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_FileExists.png)  

---

## 获取文件大小

**说明**: 指定一个文件路径，以字节为单位获取路径对应文件的大小  

**原型**: `iRet = File.FileSize(sPath)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 文件路径  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/********************************获取文件大小************************************ 命令原型： iRet = File.FileSize(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： sPath--文件路径 出参： iRet--命令运行后的结果 注意事项： 建议先判断该路径对应的文件是否存在，如果存在获取文件大小，不存在则会报错 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder\hello.txt&#x27;&#x27;&#x27; iRet = File.FileSize(sPath) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_FileSize.png)  

---

## 判断文件夹是否存在

**说明**: 指定一个文件夹路径，判断路径对应的文件夹是否存在。返回布尔值，True表示存在，False表示不存在  

**原型**: `bRet = File.FolderExists(sPath)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需要判断的文件夹路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/********************************判断文件夹是否存在******************************** 命令原型： bRet = File.FolderExists(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： sPath--需要判断的文件夹路径 出参： bRet--命令运行后的结果 注意事项： 无 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder&#x27;&#x27;&#x27; bRet = File.FolderExists(sPath) TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_FolderExists.png)  

---

## 获取文件夹大小

**说明**: 指定一个文件夹路径，以字节为单位获取路径对应文件夹中包含的所有文件的大小  

**原型**: `iRet = File.FolderSize(sPath)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需要获取大小的文件夹路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/********************************获取文件夹大小********************************** 命令原型： iRet = File.FolderSize(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： sPath--需要获取大小的文件夹路径 出参： iRet--命令运行后的结果 注意事项： 建议先判断该路径对应的文件夹是否存在，如果存在获取文件夹大小，不存在则会报错 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder&#x27;&#x27;&#x27; iRet = File.FolderSize(sPath) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_FolderSize.png)  

---

## 判断路径是否为文件

**说明**: 指定一个文件路径或变量，判断为文件则返回 True ,不为文件则返回 False，如果为非法情况，则抛出异常并中止  

**原型**: `bIsFile = File.IsFile(sPath)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 文件的路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  

**返回**: bIsFile，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/********************************判断路径是否为文件******************************* 命令原型： bIsFile = File.IsFile(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： sPath--文件的路径 出参： bIsFile--命令运行后的结果 注意事项： 建议先判断该路径对应的文件是否存在，如果存在进行判断，不存在则会报错 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder\hello.txt&#x27;&#x27;&#x27; bIsFile = File.IsFile(sPath) TracePrint(bIsFile)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_IsFile.png)  

---

## 判断路径是否为文件夹

**说明**: 指定一个文件夹路径或变量，判断为文件夹则返回 True ,不为文件夹则返回 False，如果为非法情况，则抛出异常并中止  

**原型**: `bIsDirectory = File.IsFolder(sPath)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 文件夹的路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  

**返回**: bIsDirectory，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*******************************判断路径是否为文件夹****************************** 命令原型： bIsDirectory = File.IsFolder(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： sPath--文件夹的路径 出参： bIsDirectory--命令运行后的结果 注意事项： 建议先判断该路径对应的文件夹是否存在，如果存在进行判断，不存在则会报错 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder&#x27;&#x27;&#x27; bIsDirectory = File.IsFolder(sPath) TracePrint(bIsDirectory)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_IsFolder.png)  

---

## 移动文件

**说明**: 将文件移动到指定的路径  

**原型**: `File.MoveFile(sPathSrc,sPathDst,bOverWrite)`  

**参数**:  
- **sPathSrc** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要移动的文件原始路径  
- **sPathDst** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要将文件移动到的目标路径  
- **bOverWrite** (True) [boolean] 默认:False - 如果目标文件已存在是否覆盖，传递为 true 则覆盖文件，传递为 false 则函数执行失败  

**示例**:  
```
/***********************************移动文件************************************* 命令原型： File.MoveFile(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,false) 入参： sPathSrc--要移动的文件原始路径 sPathDst--要将文件移动到的目标路径 bOverWrite-如果目标文件已存在是否覆盖，传递为 true 则覆盖文件，传递为 false 则函数执行失败 出参： 无 注意事项： 目标文件已存在默认不会覆盖，可以切换至可视化界面，在对应属性栏选择， 建议先判断要移动的文件原始路径和要将文件移动到的目标路径两者是否存在，如果存在移动文件，只要有一项不存在则会报错 ********************************************************************************/ Dim sPathSrc=&#x27;&#x27;&#x27;C:\tempFolder\hello.txt&#x27;&#x27;&#x27; Dim sPathDst=&#x27;&#x27;&#x27;C:\tempFolder\dataFolder&#x27;&#x27;&#x27; Dim bOverWrite=false File.MoveFile(sPathSrc,sPathDst,bOverWrite)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_MoveFile.png)  

---

## 移动文件夹

**说明**: 将文件夹移动到指定的路径  

**原型**: `File.MoveFolder(sPathSrc,sPathDst,bOverWrite)`  

**参数**:  
- **sPathSrc** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需要移动的文件夹路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  
- **sPathDst** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需要移动到的文件夹路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  
- **bOverWrite** (True) [boolean] 默认:False - 移动时遇到同名文件或文件夹时可选择是否替换  

**示例**:  
```
/***********************************移动文件夹************************************ 命令原型： File.MoveFolder(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,false) 入参： sPathSrc--需要移动的文件夹路径 sPathDst--需要移动到的文件夹路径 bOverWrite-移动时遇到同名文件或文件夹时可选择是否替换 出参： 无 注意事项： 目标文件夹已存在默认不会覆盖，可以切换至可视化界面，在对应属性栏选择， 建议先判断要移动的文件夹原始路径和要将文件夹移动到的目标路径两者是否存在，如果存在移动文件，只要有一项不存在则会报错 ********************************************************************************/ Dim sPathSrc=&#x27;&#x27;&#x27;C:\tempFolder\test&#x27;&#x27;&#x27; Dim sPathDst=&#x27;&#x27;&#x27;C:\tempFolder\dataFolder&#x27;&#x27;&#x27; Dim bOverWrite=false File.MoveFolder(sPathSrc,sPathDst,bOverWrite)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_MoveFolder.png)  

---

## 获取父级路径

**说明**: 指定一个文件或文件夹路径，获取路径对应文件或文件夹的父级路径  

**原型**: `sPath = File.ParentPath(filePath)`  

**参数**:  
- **filePath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 文件或文件夹的路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  

**返回**: sPath，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取父级路径************************************ 命令原型： sPath = File.ParentPath(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： filePath--文件或文件夹的路径 出参： sPath--命令运行后的结果 注意事项： 建议先判断该路径对应的文件夹或文件是否存在，如果存在获取父级路径，不存在则会报错 ********************************************************************************/ Dim filePath=&#x27;&#x27;&#x27;C:\tempFolder\hello.txt&#x27;&#x27;&#x27; sPath = File.ParentPath(filePath) TracePrint(sPath)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_ParentPath.png)  

---

## 读取文件

**说明**: 指定一个文件路径，读取路径所对应文件的内容  

**原型**: `sRet = File.Read(sPath,sCharset)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需要读取的文件路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  
- **sCharset** (True) [enum] 默认:"auto" - 文件字符集编码  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************************读取文件************************************** 命令原型： sRet = File.Read(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"auto") 入参： sPath--需要读取的文件路径 sCharset-文件字符集编码 出参： sRet--命令运行后的结果 注意事项： 默认字符集编码为自动识别，可以切换至可视化界面，在对应属性栏选择字符集编码， 建议先判断该路径对应的文件是否存在，如果存在读取文件，不存在则会报错 ********************************************************************************/ Dim sCharset="auto" Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder\hello.txt&#x27;&#x27;&#x27; sRet = File.Read(sPath,sCharset) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_Read.png)  

---

## 重命名

**说明**: 指定一个文件或文件夹路径，把该路径对应的文件或文件夹的名称进行重命名。如果指定的路径非法或者重命名时已有同名，则抛出异常并中止  

**原型**: `File.RenameEx(sPathSrc,sNewName)`  

**参数**:  
- **sPathSrc** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需要重命名的文件或文件夹路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  
- **sNewName** (True) [string] 默认:"" - 将指定路径中对应的文件或文件夹的名称进行修改后的名称  

**示例**:  
```
/************************************重命名************************************** 命令原型： File.RenameEx(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"") 入参： sPathSrc--需要重命名的文件或文件夹路径 sNewName--将指定路径中对应的文件或文件夹的名称进行修改后的名称 出参： 无 注意事项： 建议先判断该路径对应的文件或文件夹是否存在，如果存在重命名，不存在则会报错 ********************************************************************************/ //重命名文件夹 Dim sPathSrc=&#x27;&#x27;&#x27;C:\tempFolder\dataFolder&#x27;&#x27;&#x27; Dim sNewName="newDataFolder" File.RenameEx(sPathSrc,sNewName) //重命名文件 Dim sPathSrc=&#x27;&#x27;&#x27;C:\tempFolder\dataFile.txt&#x27;&#x27;&#x27; Dim sNewName="newDataFile.txt" File.RenameEx(sPathSrc,sNewName)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_RenameEx.png)  

---

## 查找

**说明**: 指定一个文件夹路径，在路径对应的文件夹内查找指定的文件或文件夹。支持模糊匹配和通配符，通配符请使用*号  

**原型**: `arrayRet = File.SearchFile(sPath,sFileName,bIsDeep)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需要查找的文件夹路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  
- **sFileName** (True) [string] 默认:"" - 要查找的文件名，支持通配符*  
- **bIsDeep** (True) [boolean] 默认:True - 是否深度查找，如果选择是，则会在给定的路径下的文件夹里面继续查找，一直查找到末端文件  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/***************************************查找************************************* 命令原型： arrayRet = File.SearchFile(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"",true) 入参： sPath--需要查找的文件夹路径 sFileName--要查找的文件名，支持通配符* bIsDeep-是否深度查找，如果选择是，则会在给定的路径下的文件夹里面继续查找，一直查找到末端文件 出参： arrayRet--命令运行后的结果 注意事项： 默认为深度查找，可以切换至可视化界面，在对应属性栏选择 建议先判断需要查找的文件夹路径是否存在，如果存在查找文件或文件夹，不存在则会报错 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder&#x27;&#x27;&#x27; Dim sFileName="hello.txt" Dim bIsDeep=true arrayRet = File.SearchFile(sPath,sFileName,bIsDeep) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_SearchFile.png)  

---

## 写入文件

**说明**: 指定一个文件路径，将内容写入到路径所对应的文件中，会将文件中原有的内容全部覆盖  

**原型**: `File.WriteFile(sPath,sText,sCharset)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 需要写入的文件路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  
- **sText** (True) [string] 默认:"" - 写入的文件内容  
- **sCharset** (True) [enum] 默认:"gbk" - 文件编码，传递为 "auto" 时自动判断编码，传递为 "ansi" 时使用ANSI编码，传递为 "utf8" 时使用utf-8编码，传递为 "unicode" 时使用 utf-16 编码  

**示例**:  
```
/**********************************写入文件************************************** 命令原型： File.WriteFile(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"","gbk") 入参： sPath--需要读取的文件路径 sText--写入的文件内容 sCharset-文件编码 出参： 无 注意事项： 默认字符集编码为gbk，可以切换至可视化界面，在对应属性栏选择其他字符集编码， 如果指定路径的文件不存在，会自动新建文件后写入 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;C:\tempFolder\hello.txt&#x27;&#x27;&#x27; Dim sText="IDP" Dim sCharset="gbk" File.WriteFile(sPath,sText,sCharset)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/File_图片/File_WriteFile.png)  

---

## 获取表单填写结果的内容

**说明**: 人机协同中心将指定表单填写结果返回后，通过指定字段名称可以获取该字段值  

**原型**: `sRet = FormCollaboration.GetFormKeyValue(jsonRet,keyList)`  

**参数**:  
- **jsonRet** (True) [expression] 默认:jsonRet - 运行"发送表单并等待填写结果"命令后返回的表单填写结果  
- **keyList** (True) [cascade] 默认:[] - 通过指定字段名称，从指定的表单填写结果中获取该字段值  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取表单填写结果的内容***************************** 命令原型： sRet = FormCollaboration.GetFormKeyValue(jsonRet,keyList) 入参: jsonRet--运行"发送表单并等待填写结果"命令后返回的表单填写结果 keyList--通过指定字段名称，从指定的表单填写结果中获取该字段值 出参: sRet--命令运行后的结果 **********************************************************************************/ Dim jsonRet,sRet // 发送表单并返回等待结果 jsonRet = FormCollaboration.SendFormData({"action_id":215,"form_data":[{"field_id":"input_zdsnczzq","input_value":""}]},@ui"窗口_FolderView",{"x":1236,"y":1275,"width":390,"height":168}) // 获取表单填写结果的内容 sRet = FormCollaboration.GetFormKeyValue(jsonRet,[215,"input_zdsnczzq"]) TracePrint sRet
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/FormCollaboration_图片/GetFormKeyValue.png)  

---

## 发送表单并等待填写结果

**说明**: 触发人机协同中心发送指定表单消息(表单内部已指定接收人员)，并等待该表单的填写结果返回  

**原型**: `jsonRet = FormCollaboration.SendFormData(formDict,objElement,objRect)`  

**参数**:  
- **formDict** (True) [dictionary] 默认:{ } - 从“人机协同中心”中选择“表单输入”类型的“协同动作”，并设置表单字段的数据绑定。须登录 Commander  
- **objElement** (True) [decorator] 默认:@ui"" - 通过鼠标选取的界面元素，包含窗口、元素等信息，或者从界面库中选择已有的界面元素  
- **objRect** (True) [dictionary] 默认:{ "x":0,"y":0,"width":0,"height":0 } - 对指定界面元素截图的范围，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则截取该元素的全区域，否则以该元素的左上角为坐标原点，根据高宽进行截图  

**返回**: jsonRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************发送表单并等待填写结果***************************** 命令原型： jsonRet = FormCollaboration.SendFormData(formDict,objElement,objRect) 入参: formDict--从“人机协同中心”中选择“表单输入”类型的“协同动作”，并设置表单字段的数据绑定。须登录 Commander objElement--通过鼠标选取的界面元素，包含窗口、元素等信息，或者从界面库中选择已有的界面元素 objRect--对指定界面元素截图的范围，如果范围传递为 {"x":0,"y":0,"width":0,"height":0}，则截取该元素的全区域，否则以该元素的左上角为坐标原点，根据高宽进行截图 出参: objJSON--命令运行后的结果 **********************************************************************************/ Dim jsonRet jsonRet = FormCollaboration.SendFormData({"action_id":215,"form_data":[{"field_id":"input_zdsnczzq","input_value":""}]},@ui"窗口_FolderView",{"x":1236,"y":1275,"width":390,"height":168})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/FormCollaboration_图片/SendFormData.png)  

---

## 连接FTP服务器

**说明**: 连接指定属性的FTP服务器，并作为操控对象  

**原型**: `objFTP = FTP.Connect(sHostName,sUser,sPassword,bSsl,nPort,bAnonymous,nFTPSMode,optionArgs)`  

**参数**:  
- **sHostName** (True) [string] 默认:"" - 服务器地址  
- **sUser** (True) [string] 默认:"" - 登录用户名  
- **sPassword** (True) [string] 默认:"" - 登录密码  
- **bSsl** (True) [boolean] 默认:False - 使用SSL加密连接，默认为否  
- **nPort** (True) [number] 默认:21 - 连接端口，如21、22  
- **bAnonymous** (True) [boolean] 默认:False - 可匿名登录，默认为否  
- **nFTPSMode** (True) [enum] 默认:1 - FTP连接模式  
- **sftp** (False) [boolean] 默认:False - 使用SFTP协议加密连接，默认为否  
- **private_key** (False) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 密钥文件存储路径  

**返回**: objFTP，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************连接FTP服务器***************************** 命令原型： objFTP = FTP.Connect(sHostName,sUser,sPassword,bSsl,nPort,bAnonymous,nFTPSMode,optionArgs) 入参: sHostName--服务器地址 sUser--登录用户名 sPassword--登录密码 bSsl--使用SSL加密连接，默认为否 nPort--连接端口，如21、22 bAnonymous--可匿名登录，默认为否 nFTPSMode--FTP连接模式 sftp--使用SFTP协议加密连接，默认为否 private_key--密钥文件存储路径 出参: objFTP--命令运行后的结果 **********************************************************************************/ Dim objFTP,bRet objFTP = FTP.Connect("*.*.*.*","test","UjsIVVKa/FTORllo0BElZg==",false,21,false,1,{"sftp":false,"private_key":&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/FTP_图片/Connect.png)  

---

## 变量赋值

**说明**: 将等式左边的变量修改为右边的值，等式右侧可以为立即数值、表达式、命令输出等  

**原型**: `varName = varValue`  

**参数**:  
- **varName** (True) [id] 默认:temp - 变量名  
- **varValue** (True) [expression] 默认:"" - 变量的值  

**示例**:  
```
/*********************************变量赋值******************************** 命令原型： varName = varValue 入参： varName -- 变量名 varValue -- 变量的值 出参： 无 注意事项： 将等式左边的变量修改为右边的值，等式右侧可以为立即数值、表达式、命令输出等 ********************************************************************************/ Dim a="UiBot" TracePrint "给变量a赋值UiBot"
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_Assign.png)  

---

## 跳出循环

**说明**: 跳出 计次循环、遍历数据  

**原型**: `Break`  

**参数**:  
- **无** (无) [无] 默认:无 - 无  

**示例**:  
```
/*********************************跳出循环******************************** 命令原型： Break 入参： 无 出参： 无 注意事项： 跳出 计次循环、遍历数据，后面的循环将不会执行 ********************************************************************************/ TracePrint("令数组a = [&#x27;U&#x27;,&#x27;i&#x27;,&#x27;1&#x27;,&#x27;B&#x27;,&#x27;o&#x27;,&#x27;T&#x27;],遍历数组a,当value为1时退出循环") a = ["U","i","1","B","o","T"] For Each value In a If value = "1" Break End If TracePrint value Next
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_Break.png)  

---

## 如果分支条件符合判断条件则执行分支语句块

**说明**: 独立的 Case 条件分支，仅能与 Select Case 语句组合使用，且可添加多个条件分支  

**原型**: `Case expression`  

**参数**:  
- **expression** (True) [expression] 默认:条件1 - 进行判断的表达式  

**示例**:  
```
/*********************************条件分支******************************** 命令原型： Case expression 入参： 无 出参： 无 注意事项： 仅能与Select Case 语句组合使用。Select Case的表达式满足后续Case的所有条件将执行该Case分支的逻辑 ********************************************************************************/ a = 1 Select Case a + 1 Case 1 TracePrint("This is Case 1") Case 2 TracePrint("This is Case 2") Case 3 TracePrint("This is Case 3") End Select
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_Case.png)  

---

## 选择其他分支条件执行其他分支语句块

**说明**: 当所有预设的 Case 分支都无法与判断条件匹配时，可以使用 Case Else 分支，用来代表其他情况  

**原型**: `Case Else`  

**参数**:  
- **无** (无) [无] 默认:无 - 代表其他情况的分支  

**示例**:  
```
/*********************************选择其他分支条件执行其他分支语句块*************************************** 命令原型： dRet = CNumber(varData) 入参： 无 出参： 无 注意事项： 仅能与Select Case 语句组合使用，所有预设的 Case 分支都无法与判断条件匹配时Case分支的逻辑 ********************************************************************************/ a = 4 Select Case a + 1 Case 1 TracePrint("This is Case 1") Case 2 TracePrint("This is Case 2") Case 3 TracePrint("This is Case 3") Case Else TracePrint("This is Case Else") End Select
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_CaseElse.png)  

---

## 继续循环

**说明**: 跳过当次循环，进入当前循环语句的下一轮循环  

**原型**: `Continue`  

**参数**:  
- **无** (无) [无] 默认:无 - 进入当前循环的下一轮  

**示例**:  
```
/*********************************继续循环*************************************** 命令原型： Continue 入参： 无 出参： 无 注意事项： 跳过当次循环，进入当前循环语句的下一轮循环 ********************************************************************************/ TracePrint("令数组a = [&#x27;U&#x27;,&#x27;i&#x27;,&#x27;1&#x27;,&#x27;B&#x27;,&#x27;o&#x27;,&#x27;T&#x27;],遍历数组a,当value为1时跳过本次循环，继续下一次循环，并打印value值") a = ["U","i","1","B","o","T"] For Each value In a If value = "1" Continue End If TracePrint value Next
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_Continue.png)  

---

## 无限循环执行操作

**说明**: 无停止条件，一直循环执行操作  

**原型**: `Do Loop`  

**参数**:  
- **无** (无) [无] 默认:无 - 无停止条件，将会一直执行操作  

**示例**:  
```
/*********************************无限循环执行操作******************************** 命令原型： Do Loop 入参： 无 出参： 无 注意事项： 没有退出条件，循环将一直执行下去，消耗系统资源，不推荐使用 ********************************************************************************/ TracePrint("无限执行a加1计算") a = 1 Do a = a+1 Loop
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_DoLoop.png)  

---

## 先执行操作当后置条件不成立时继续循环执行操作

**说明**: 先执行语句块 (Block)，使用 Until 时，当条件表达式 (expression) 的值为假(不成立)时则继续执行语句块 (Block)，当条件表达式 (expression) 的值为真(成立)时退出循环  

**原型**: `Do Loop Until expression`  

**参数**:  
- **expression** (True) [expression] 默认:后置条件 - 进行判断的表达式  

**示例**:  
```
/*********************************先执行语句块当条件表达式为真的时候退出循环******************************** 命令原型： Do Loop Until expression 入参： 无 出参： 无 注意事项： 先执行语句块 (Block)，使用 Until 时，当条件表达式 (expression) 的值为假(不成立)时则继续执行语句块 (Block)，当条件表达式 (expression) 的值为真(成立)时退出循环 ********************************************************************************/ TracePrint("令a=1,执行a自增1，直到a等于10，退出循环") a = 1 Do a = a + 1 Loop Until a=10 TracePrint(a)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_DoLoopUntil.png)  

---

## 先执行操作当后置条件成立时继续循环执行操作

**说明**: 先执行语句块 (Block)，使用 While 时，当条件表达式 (expression) 的值为真则继续执行语句块 (Block)，当条件表达式 (expression) 的值为假时退出循环  

**原型**: `Do Loop While expression`  

**参数**:  
- **expression** (True) [expression] 默认:后置条件 - 进行判断的表达式  

**示例**:  
```
/*********************************后置条件成立时继续循环******************************** 命令原型： Do Loop Until expression 入参： 无 出参： 无 注意事项： 使用 While 时，执行 Block 语句块，当 expression 为真时退出循环。 ********************************************************************************/ TracePrint("令a=1，令a自增1，此时a = 2,符合循环条件a=2，退出循环") a = 1 Do a = a + 1 TracePrint a Loop While a = 2
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_DoLoopWhile.png)  

---

## 当前置条件不成立时循环执行操作

**说明**: 使用 Until 时，当条件表达式 (expression) 的值为假(不成立)时则继续执行语句块 (Block)，当条件表达式 (expression) 的值为真(成立)时退出循环  

**原型**: `Do Until expression Loop`  

**参数**:  
- **expression** (True) [expression] 默认:前置条件 - 进行判断的表达式  

**示例**:  
```
/*********************************条件不成立时循环******************************** 命令原型： Do Until expression Loop 入参： 无 出参： 无 注意事项： 使用 Until 时，当 expression 的值为真则继续执行 Block 语句块，当 expression 为假时退出循环。 ********************************************************************************/ TracePrint("令a=1,当a不等于10时，执行a自增1，直到a等于10，退出循环") a = 1 Do Until a=10 a = a + 1 Loop TracePrint(a)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_DoUntilLoop.png)  

---

## 当前置条件成立时循环执行操作

**说明**: 使用 While 时，当条件表达式 (expression) 的值为真则继续执行语句块 (Block)，当条件表达式 (expression) 的值为假时退出循环  

**原型**: `Do While expression Loop`  

**参数**:  
- **expression** (True) [expression] 默认:前置条件 - 进行判断的表达式  

**示例**:  
```
/*********************************条件循环******************************** 命令原型： Do While expression Loop 入参： 无 出参： 无 注意事项： 使用 While 时，当 expression 的值为真则继续执行 Block 语句块，当 expression 为假时退出循环。 ********************************************************************************/ TracePrint("令a=1，当a=1时开始条件循环，循环内令a自增1，此时a = 2,不符合循环条件，退出循环") a = 1 Do While a = 1 a = a + 1 TracePrint a Loop
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_DoWhileLoop.png)  

---

## 否则执行后续操作

**说明**: 否则如果表达式 expression 为真，则执行 Yes Block 语句块，否则执行 No Block 语句块。ElseIf 语句可以出现多次，对应更多的条件分支，ElseIf、Else 语句如果不需要可以不编写，对应的语句块也不需要编写  

**原型**: `Else`  

**参数**:  
- **无** (无) [无] 默认:无 - 无法单独使用，一般配合IF是使用  

**示例**:  
```
/*********************************条件分支******************************** 命令原型： Else 入参： 无 出参： 无 注意事项： 如果表达式 expression 为真，则执行 Yes Block 语句块，否则执行 No Block 语句块。ElseIf 语句可以出现多次，对应更多的条件分支，ElseIf、Else 语句如果不需要可以不编写，对应的语句块也不需要编写. ********************************************************************************/ TracePrint("令变量a为UiBot,条件分支判断a是否等于UiBot，是则打印yes,否则打印no为：") a = "UiBot" If a = "UiBot" TracePrint ("yes") Else TracePrint ("no") End If
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_CaseElse.png)  

---

## 否则如果条件成立则执行后续操作

**说明**: 否则如果条件表达式 (expression) 为真，则执行语句块 (Block) ，否则包含的语句块不执行。ElseIf 语句可以出现多次，对应更多的条件分支，但仅在If/End块内使用  

**原型**: `ElseIf expression`  

**参数**:  
- **expression** (True) [expression] 默认:条件成立 - 进行判断的表达式  

**示例**:  
```
/*********************************条件分支******************************** 命令原型： ElseIf expression 入参： 无 出参： 无 注意事项： 如果表达式 expression 为真，则执行 Yes Block 语句块，否则执行 No Block 语句块。ElseIf 语句可以出现多次，对应更多的条件分支，ElseIf、Else 语句如果不需要可以不编写，对应的语句块也不需要编写。 ********************************************************************************/ TracePrint("令变量a为UiBot,条件分支判断a是否等于1,然后判断a是否等于UiBot执行相应的语句块") a = "1" If a = 1 TracePrint "yes" ElseIf a = "UiBot" TracePrint "no" End If
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_ElseIf.png)  

---

## 退出流程

**说明**: 退出流程  

**原型**: `Exit()`  

**参数**:  
- **无** (无) [无] 默认:无 - 无  

**示例**:  
```
/*********************************退出流程******************************** 命令原型： ElseIf expression 入参： 无 出参： 无 注意事项： 退出流程，该代码后的流程无法执行到 ********************************************************************************/ TracePrint("令a为UiBot，条件分支，如果a=UiBot时退出流程") a = "UiBot" If a = "UiBot" exit() End If TracePrint("无法到达代码")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_Exit.png)  

---

## 依次读取字典中每对键值针对取到的键值对进行操作

**说明**: 循环遍历 Collection 的每一条数据，将数据的名称（对应字典时）或索引（对应数组时）代入 key，将数据的值代入 value 后执行 Block 语句块  

**原型**: `For Each key, value In dataStruct Next`  

**参数**:  
- **key** (True) [id] 默认:key - 遍历键  
- **value** (True) [id] 默认:value - 遍历值  
- **dataStruct** (True) [expression] 默认:dictVar - 要遍历的数据结构  

**示例**:  
```
/*********************************遍历字典******************************** 命令原型： For Each key, value In dataStruct Next 入参： key -- 遍历键 value -- 遍历值 dataStruct -- 要遍历的数据结构 出参： 无 注意事项： 循环遍历 Collection 的每一条数据，将数据的名称（对应字典时）或索引（对应数组时）代入 key，将数据的值代入 value 后执行 Block 语句块 ********************************************************************************/ TracePrint("令字典为a = {&#x27;a&#x27;:&#x27;U&#x27;,&#x27;b&#x27;:&#x27;i&#x27;,&#x27;c&#x27;:&#x27;B&#x27;,&#x27;d&#x27;:&#x27;o&#x27;,&#x27;e&#x27;:&#x27;T&#x27;},遍历打印每一个key,value") a = {&#x27;a&#x27;:&#x27;U&#x27;,&#x27;b&#x27;:&#x27;i&#x27;,&#x27;c&#x27;:&#x27;B&#x27;,&#x27;d&#x27;:&#x27;o&#x27;,&#x27;e&#x27;:&#x27;T&#x27;} For Each key,value In a TracePrint "key值为："&key&" value值为："&value Next
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_ForEachKeyValueNext.png)  

---

## 依次读取数组中每个元素针对取到的元素进行操作

**说明**: 循环遍历 Collection 的每一条数据，将数据的名称（对应字典时）或索引（对应数组时）代入 key，将数据的值代入 value 后执行 Block 语句块  

**原型**: `For Each value In dataStruct Next`  

**参数**:  
- **value** (True) [id] 默认:value - 遍历值  
- **dataStruct** (True) [expression] 默认:arrayRet - 要遍历的数据结构  

**示例**:  
```
/*********************************遍历数组******************************** 命令原型： For Each value In dataStruct Next 入参： value -- 遍历值 dataStruct -- 要遍历的数据结构 出参： 无 注意事项： 循环遍历 Collection 的每一条数据，将数据的名称（对应字典时）或索引（对应数组时）代入 key，将数据的值代入 value 后执行 Block 语句块 ********************************************************************************/ TracePrint("令数组a = [&#x27;U&#x27;,&#x27;i&#x27;,&#x27;B&#x27;,&#x27;o&#x27;,&#x27;T&#x27;],遍历打印每一个值") a = [&#x27;U&#x27;,&#x27;i&#x27;,&#x27;B&#x27;,&#x27;o&#x27;,&#x27;T&#x27;] For Each value In a TracePrint value Next
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_ForEachNext.png)  

---

## 从初始值开始按步长计数继续计数直至到结束值停止

**说明**: 将索引代入循环，执行for 至 next 之间的语句块，索引的值从 start 到 end 次，step为步长，每次循环可以使索引增加 step 对应的数量，而不再是 1  

**原型**: `For index = beginValue To endValue step stepValue Next`  

**参数**:  
- **index** (True) [id] 默认:i - 步长索引  
- **beginValue** (True) [number] 默认:0 - 初始值  
- **endValue** (True) [number] 默认:10 - 结束值  
- **stepValue** (True) [number] 默认:1 - 步长  

**示例**:  
```
/*********************************计次循环******************************** 命令原型： For Each value In dataStruct Next 入参： index -- 步长索引 beginValue -- 初始值 endValue -- 结束值 stepValue -- 步长 出参： 无 注意事项： 将索引代入循环，执行for 至 next 之间的语句块，索引的值从 start 到 end 次，step为步进，每次循环可以使索引增加 step 对应的数量，而不再是 1。 ********************************************************************************/ TracePrint("令i为0，步长为1，自增到10，打印i的值，结果为：") For i = 0 To 10 Step 1 TracePrint i Next
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_ForToNext.png)  

---

## 如果条件成立则执行后续操作

**说明**: 如果表达式 expression 为真，则执行 Yes Block 语句块，否则执行 No Block 语句块。ElseIf 语句可以出现多次，对应更多的条件分支，ElseIf、Else 语句如果不需要可以不编写，对应的语句块也不需要编写  

**原型**: `If expression End If`  

**参数**:  
- **expression** (True) [expression] 默认:条件成立 - 进行判断的表达式  

**示例**:  
```
/*********************************条件分支******************************** 命令原型： If expression End If 入参： 无 出参： 无 注意事项： 如果表达式 expression 为真，则执行 Yes Block 语句块，否则执行 No Block 语句块。ElseIf 语句可以出现多次，对应更多的条件分支，ElseIf、Else 语句如果不需要可以不编写，对应的语句块也不需要编写。 ********************************************************************************/ TracePrint("令变量a为UiBot,条件分支判断a是否等于1，是则打印yes,否则打印no,结果为：") a = "UiBot" If a = 1 TracePrint "yes" Else TracePrint "no" End If
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_ElseIf.png)  

---

## 跳出返回

**说明**: 跳出流程块并输出  

**原型**: `Return value`  

**参数**:  
- **value** (True) [expression] 默认:retValue - 返回到流程视图的值  

**示例**:  
```
/*********************************跳出返回******************************** 命令原型： Return value 入参： 无 出参： 无 注意事项： 跳出流程块并返回值，Return后面的代码不会执行,并且流程块结束，返回值可以在与流程图相连的下一个流程中使用 ********************************************************************************/ TracePrint("令a为UiBot，条件分支，如果a=UiBot时跳出返回") a = "UiBot" If a = "UiBot" Return a End If
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_Return.jpg)  

---

## 根据判断条件从多个分支中选择出一个

**说明**: 根据一定的条件，选择多个分支中的一个。先计算 Select Case 后面的表达式，然后判断是否有某个 Case 分支和这个表达式的值是一致的，若有则执行 Case 分支的语句块，若无则执行 Case Else (如果有) 分支的语句块  

**原型**: `Select Case expression Case 条件1 End Select`  

**参数**:  
- **expression** (True) [expression] 默认:判断条件 - 进行判断的表达式  

**示例**:  
```
/*********************************条件分支语句块******************************** 命令原型： Select Case expression Case 条件1 End Select 入参： 无 出参： 无 注意事项： 推荐使用if条件分支语句 ********************************************************************************/ TracePrint("令a为1，根据a+1选择，如果a+1是2，则执行Case 2下的流程块") a = 1 Select Case a+1 Case 2 TracePrint(2) Case 3 TracePrint(3) End Select
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_SelectCase.png)  

---

## 抛出异常

**说明**: 抛出一个自定义异常，当前程序的运行将被停止（Throw之后的语句将不会执行），可使用异常捕获（Catch）命令将抛出的自定义异常捕获  

**原型**: `Throw message`  

**参数**:  
- **message** (True) [string] 默认:"" - 抛出内容  

**示例**:  
```
/*********************************抛出异常******************************** 命令原型： Select Case expression Case 条件1 End Select 入参： 无 出参： 无 注意事项： 抛出一个自定义异常，当前程序的运行将被停止（Throw之后的语句将不会执行），可使用异常捕获（Catch）命令将抛出的自定义异常捕获。 ********************************************************************************/ TracePrint "UB是不能用数组拼接字符串，自定义抛出异常" Try TracePrint []&"a" Catch error TracePrint error Throw "语法错误" End Try
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_Throw.png)  

---

## 尝试执行操作发生异常则进行处理

**说明**: 使用try catch处理可能引发异常的代码块  

**原型**: `Try Catch 变量名 End Try`  

**参数**:  
- **无** (无) [无] 默认:无 - 无  

**示例**:  
```
/*********************************异常捕获******************************** 命令原型： Try Catch 变量名 End Try 入参： 无 出参： 无 注意事项： 使用try catch处理可能引发异常的代码块。 ********************************************************************************/ Try TracePrint []&"a" Catch error TracePrint error End Try
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_TryCatch.png)  

---

## 尝试执行操作且可以重试N次重试后仍发生异常则进行处理

**说明**: 使用try catch处理可能引发异常的代码块  

**原型**: `Try count Catch 变量名 End Try`  

**参数**:  
- **count** (True) [number] 默认:3 - 值必须为大于等于1的整数。如果发生了异常，则自动回到Try的地方重试，重试后还有异常，才会跳至Catch后的语句执行  

**示例**:  
```
/*********************************异常重试******************************** 命令原型： Try count Catch 变量名 End Try 入参： 无 出参： 无 注意事项： 使用try catch处理可能引发异常的代码块 ********************************************************************************/ TracePrint ("重试执行三次try中的内容，如出现异常之后再执行Catch中的内容") Try 3 TracePrint "a" TracePrint []&"a" Catch e TracePrint error Else TracePrint "123" End Try
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Grammar_图片/Grammar_TryNCatch.gif)  

---

## Get获取数据

**说明**: HTTP.Get 获取网络数据  

**原型**: `sRet = HTTP.Get(sURL, sForm, iTimeout)`  

**参数**:  
- **sURL** (True) [string] 默认:"" - Get页面的链接地址  
- **sForm** (True) [expression] 默认:{ } - Get时传递的表单数据，可以是字符串或字典  
- **iTimeout** (True) [number] 默认:60000 - 超时时间（毫秒）  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************Get获取数据*************************************** 命令原型： sRet = HTTP.Get("", {}, 60000) 入参： sURL--链接地址。注：Get页面的链接地址 sForm--表单数据。注：Get时传递的表单数据，可以是字符串或字典 iTimeout--超时时间（毫秒）。注：超时时间（毫秒） 出参： sRet--函数调用的输出保存到的变量。 注意事项： 1.要保证机器能够访问到网站所在网络。 2.表单使用如下：（原网址）https://www.baidu.com/s?ie=UTF-8&wd=get%E8%AF%B7%E6%B1%82%E9%99%84%E5%8A%A0%E8%A1%A8%E5%8D%95 表单内容为?号后面的所有参数，一个=号为一个键值对，如上：{&#x27;ie&#x27;:&#x27;UTF-8&#x27;,&#x27;wd&#x27;:&#x27;get%E8%AF%B7%E6%B1%82%E9%99%84%E5%8A%A0%E8%A1%A8%E5%8D%95&#x27;} ********************************************************************************/ Dim sRet = "" sRet = HTTP.Get("https://www.baidu.com/s?", {&#x27;ie&#x27;:&#x27;UTF-8&#x27;, &#x27;wd&#x27;:&#x27;get%E8%AF%B7%E6%B1%82%E9%99%84%E5%8A%A0%E8%A1%A8%E5%8D%95&#x27;}, 60000) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/HTTP_图片/HTTP_Get.png)  

---

## 删除邮件

**说明**: 删除指定邮件消息  

**原型**: `IBMNotes.Delete(pwd,messageJson)`  

**参数**:  
- **pwd** (True) [string] 默认:"" - IBM Notes上绑定邮箱的登录密码  
- **messageJson** (True) [expression] 默认:{ } - 邮件列表中的邮件对象  

**示例**:  
```
/*********************************删除邮件***************************** 命令原型： IBMNotes.Delete(pwd,messageJson) 入参: pwd--IBM Notes上绑定邮箱的登录密码 messageJson--邮件列表中的邮件对象 注意事项 无 **********************************************************************************/ Dim arrayRet // 获取邮件列表 arrayRet = IBMNotes.GetMailList("123456","$Inbox",0,30000) Traceprint arrayRet // 删除邮件 IBMNotes.Delete("123456", arrayRet[0])
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/IBM-Notes_图片/Delete.png)  

---

## 点击图像

**说明**: 在指定范围内搜索图像并单击它  

**原型**: `Image.Click(objUiElement,objRect,sImagePath,iAccuracy,iButton,iType,iTimeOut, optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **objRect** (True) [dictionary] 默认:{ "x": 0, "y": 0, "width": 0, "height": 0 } - 需要查找的范围，程序会在控件这个范围内进行文字识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的文字识别  
- **sImagePath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要查找的图片路径，一般在res文件夹  
- **iAccuracy** (True) [number] 默认:0.9 - 查找图片时使用的相似度，相似度范围从 0.5 - 1.0，表示 50% - 100% 相似  
- **iButton** (True) [enum] 默认:"left" - 鼠标按键 { left:左键, right:右键, middle:中键 }  
- **iType** (True) [enum] 默认:"click" - 点击类型 { click:单击, dbclick:双击, down:按下, up:弹起 }  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sCursorPosition** (False) [enum] 默认:"Center" - 描述添加OffsetX和OffsetY属性的偏移量的光标起点。可以使用以下选项：TopLeft，TopRight，BottomLeft，BottomRight和Center。默认选项是Center  
- **iCursorOffsetX** (False) [number] 默认:10 - 根据在“位置”字段中选择的选项，光标位置的水平位移  
- **iCursorOffsetY** (False) [number] 默认:10 - 根据在“位置”字段中选择的选项，光标位置的垂直位移  
- **sKeyModifiers** (False) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **sSimulate** (False) [enum] 默认:"simulate" - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：模拟操作(simulate)  
- **sMatchType** (False) [enum] 默认:"GrayMatch" - 指定查找图像的匹配方式，“灰度匹配”速度快，但在极端情况下可能会匹配失败，“彩色匹配”相对“灰度匹配”更精准但匹配速度稍慢  
- **iSerialNo** (False) [number] 默认:1 - 指定图像匹配到多个目标时的序号，序号为从1开始的正整数，在屏幕上从左到右从上到下依次递增，匹配到最靠近屏幕左上角的目标序号为1  

**示例**:  
```
/************************点击图像************************ 命令原型: Image.Click(objUiElement,objRect,sImagePath,iAccuracy,iButton,iType,iTimeOut, optionArgs) 入参: objUiElement--目标元素 objRect--识别范围 sImagePath--查找图片 iAccuracy--相似度 iButton--鼠标点击(左键/右键/中键) iType--点击类型(单击/双击/按下/弹起) iTimeOut--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时/激活窗口/光标位置/横坐标偏移/纵坐标偏移/辅助按键/操作类型/匹配方式/匹配序号).Type:Dict 出参： 无 注意事项: 必须选定目标 *******************************************************/ Image.Click({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"百度一下，你就知道 - Google Chrome","app":"chrome"}]},{"x": 0, "y": 0, "width": 0, "height": 0},@res"c6121320-d4c2-11ec-adbf-19b54b1b1b3f.png",0.9,"left","click",10000, {"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200, "bSetForeground": true, "sCursorPosition": "Center", "iCursorOffsetX": 10, "iCursorOffsetY": 10, "sKeyModifiers": [],"sSimulate": "simulate","sMatchType":"GrayMatch", "iSerialNo": 1})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Image_图片/Image_Click.png)  

---

## 判断图像是否存在

**说明**: 在指定范围内查找图像，成功返回 true，失败返回 false  

**原型**: `bRet = Image.Exists(objUiElement,objRect,sImagePath,iAccuracy,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **objRect** (True) [dictionary] 默认:{ "x": 0, "y": 0, "width": 0, "height": 0 } - 需要查找的范围，程序会在控件这个范围内进行文字识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的文字识别  
- **sImagePath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要查找的图片路径，一般在res文件夹  
- **iAccuracy** (True) [number] 默认:0.9 - 查找图片时使用的相似度，相似度范围从 0.5 - 1.0，表示 50% - 100% 相似  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sMatchType** (False) [enum] 默认:"GrayMatch" - 指定查找图像的匹配方式，“灰度匹配”速度快，但在极端情况下可能会匹配失败，“彩色匹配”相对“灰度匹配”更精准但匹配速度稍慢  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************判断图像是否存在************************ 命令原型: bRet = Image.Exists(objUiElement,objRect,sImagePath,iAccuracy,iTimeOut,optionArgs) 入参: objUiElement--目标元素 objRect--识别范围 sImagePath--查找图片 iAccuracy--相似度 iTimeOut--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时/激活窗口/匹配方式).Type:Dict 出参: bRet--函数调用的输出保存到的变量 注意事项: 必须选定目标 ***********************************************************/ bRet = Image.Exists({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"百度一下，你就知道 - Google Chrome","app":"chrome"}]},{"x": 0, "y": 0, "width": 0, "height": 0},@res"a8c41750-d4c2-11ec-adbf-19b54b1b1b3f.png",0.9,10000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200, "bSetForeground": true, "sMatchType":"GrayMatch"})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Image_图片/Image_Exists.png)  

---

## 查找图像

**说明**: 在指定范围内查找图像  

**原型**: `objPoint = Image.Find(objUiElement,objRect,sImagePath,iAccuracy,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **objRect** (True) [dictionary] 默认:{ "x": 0, "y": 0, "width": 0, "height": 0 } - 需要查找的范围，程序会在控件这个范围内进行文字识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的文字识别  
- **sImagePath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要查找的图片路径，一般在res文件夹  
- **iAccuracy** (True) [number] 默认:0.9 - 查找图片时使用的相似度，相似度范围从 0.5 - 1.0，表示 50% - 100% 相似  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sMatchType** (False) [enum] 默认:"GrayMatch" - 指定查找图像的匹配方式，“灰度匹配”速度快，但在极端情况下可能会匹配失败，“彩色匹配”相对“灰度匹配”更精准但匹配速度稍慢  
- **iSerialNo** (False) [number] 默认:1 - 指定图像匹配到多个目标时的序号，序号为从1开始的正整数，在屏幕上从左到右从上到下依次递增，匹配到最靠近屏幕左上角的目标序号为1  

**返回**: objPoint，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************查找图像************************ 命令原型: objPoint = Image.Find(objUiElement,objRect,sImagePath,iAccuracy,iTimeOut,optionArgs) 入参: objUiElement--目标元素 objRect--识别范围 sImagePath--查找图片 iAccuracy--相似度 iTimeOut--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时/激活窗口/匹配方式/匹配序号).Type:Dict 出参: objPoint--函数调用的输出保存到的变量 注意事项: 必须选定目标 ***********************************************************/ objPoint = Image.Find({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"百度一下，你就知道 - Google Chrome","app":"chrome"}]},{"x": 0, "y": 0, "width": 0, "height": 0},@res"819da920-d4c2-11ec-adbf-19b54b1b1b3f.png",0.9,10000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200, "bSetForeground": true, "sMatchType":"ColorMatch", "iSerialNo": 1})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Image_图片/Image_Find.png)  

---

## 鼠标移动到图像上

**说明**: 在指定范围内搜索图像并将鼠标指针移动到图像之上  

**原型**: `Image.Hover(objUiElement,objRect,sImagePath,iAccuracy,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **objRect** (True) [dictionary] 默认:{ "x": 0, "y": 0, "width": 0, "height": 0 } - 需要查找的范围，程序会在控件这个范围内进行文字识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的文字识别  
- **sImagePath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要查找的图片路径，一般在res文件夹  
- **iAccuracy** (True) [number] 默认:0.9 - 查找图片时使用的相似度，相似度范围从 0.5 - 1.0，表示 50% - 100% 相似  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sCursorPosition** (False) [enum] 默认:"Center" - 描述添加OffsetX和OffsetY属性的偏移量的光标起点。可以使用以下选项：TopLeft，TopRight，BottomLeft，BottomRight和Center。默认选项是Center  
- **iCursorOffsetX** (False) [number] 默认:10 - 根据在“位置”字段中选择的选项，光标位置的水平位移  
- **iCursorOffsetY** (False) [number] 默认:10 - 根据在“位置”字段中选择的选项，光标位置的垂直位移  
- **sKeyModifiers** (False) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **sSimulate** (False) [enum] 默认:"simulate" - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：模拟操作(simulate)  
- **sMatchType** (False) [enum] 默认:"GrayMatch" - 指定查找图像的匹配方式，“灰度匹配”速度快，但在极端情况下可能会匹配失败，“彩色匹配”相对“灰度匹配”更精准但匹配速度稍慢  
- **iSerialNo** (False) [number] 默认:1 - 指定图像匹配到多个目标时的序号，序号为从1开始的正整数，在屏幕上从左到右从上到下依次递增，匹配到最靠近屏幕左上角的目标序号为1  

**示例**:  
```
/************************鼠标移动到图像上************************ 命令原型: Image.Hover(objUiElement,objRect,sImagePath,iAccuracy,iTimeOut,optionArgs) 入参: objUiElement--目标元素 objRect--识别范围 sImagePath--查找图片 iAccuracy--相似度 iTimeOut--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时/激活窗口/光标位置/横坐标偏移/纵坐标偏移/辅助按键/操作类型/匹配方式/匹配序号).Type:Dict 出参： 无 注意事项: 必须选定目标 ***********************************************************/ Image.Hover({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"百度一下，你就知道 - Google Chrome","app":"chrome"}]},{"x": 0, "y": 0, "width": 0, "height": 0},@res"895d6a20-d4c1-11ec-adbf-19b54b1b1b3f.png",0.9,10000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200, "bSetForeground": true, "sCursorPosition": "Center", "iCursorOffsetX": 10, "iCursorOffsetY": 10, "sKeyModifiers": [],"sSimulate": "simulate","sMatchType":"GrayMatch", "iSerialNo": 1})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Image_图片/Image_Hover.png)  

---

## 等待图像

**说明**: 等待图片显示或消失  

**原型**: `Image.Wait(objUiElement,objRect,sImagePath,iAccuracy,iType,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **objRect** (True) [dictionary] 默认:{ "x": 0, "y": 0, "width": 0, "height": 0 } - 需要查找的范围，程序会在控件这个范围内进行文字识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的文字识别  
- **sImagePath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要查找的图片路径，一般在res文件夹  
- **iAccuracy** (True) [number] 默认:0.9 - 查找图片时使用的相似度，相似度范围从 0.5 - 1.0，表示 50% - 100% 相似  
- **iType** (True) [enum] 默认:"show" - 等待方式，可以设置为等待图像显示后结束或等待图片消失后结束  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sMatchType** (False) [enum] 默认:"GrayMatch" - 指定查找图像的匹配方式，“灰度匹配”速度快，但在极端情况下可能会匹配失败，“彩色匹配”相对“灰度匹配”更精准但匹配速度稍慢  

**示例**:  
```
/************************等待图像************************ 命令原型: Image.Wait(objUiElement,objRect,sImagePath,iAccuracy,iType,iTimeOut,optionArgs) 入参: objUiElement--目标元素 objRect--识别范围 sImagePath--查找图片 iAccuracy--相似度 iType--等待方式 iTimeOut--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时/激活窗口/匹配方式).Type:Dict 出参： 无 注意事项: 必须选定目标 ***********************************************************/ Image.Wait({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"百度一下，你就知道 - Google Chrome","app":"chrome"}]},{"x": 0, "y": 0, "width": 0, "height": 0},@res"fd90ec40-d4c2-11ec-adbf-19b54b1b1b3f.png",0.9,"show",10000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200, "bSetForeground": true, "sMatchType":"GrayMatch"})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Image_图片/Image_Wait.png)  

---

## 断开邮箱连接

**说明**: 断开IMAP连接  

**原型**: `IMAP.Close(objIMAP)`  

**参数**:  
- **objIMAP** (True) [expression] 默认:objIMAP - 由"连接邮箱"命令返回的可操控连接对象  

**示例**:  
```
/*********************************断开邮箱连接*************************************** 命令原型： IMAP.Close(objIMAP) 入参: objIMAP--由"连接邮箱"命令返回的可操控连接对象 注意事项: 邮箱使用完后及时关闭连接 **********************************************************************************/ Dim objIMAP // 连接邮箱 objIMAP = IMAP.Connect("imap.qq.com","***@qq.com","*****",143,false,"***@qq.com") // 断开邮箱连接 IMAP.Close(objIMAP)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/IMAP_图片/Close.png)  

---

## 删除键

**说明**: 删除 INI 配置文件下指定小节的指定键  

**原型**: `INI.DeleteKey(sFile, sSection, sKey)`  

**参数**:  
- **sFile** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - INI 配置文件所在路径  
- **sSection** (True) [string] 默认:"" - 要删除 INI 配置文件的小节名字  
- **sKey** (True) [string] 默认:"" - 要删除 INI 配置文件的键名  

**示例**:  
```
/***********************************删除键*************************************** 命令原型： INI.DeleteKey(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;, "", "") 入参： sFile--INI配置文件所在路径 sSection--要删除INI配置文件的小节名字 sKey--要删除INI配置文件的键名 出参： 无 注意事项： 建议先判断该路径对应的文件是否存在，如果存在删除指定键，不存在则会报错 ********************************************************************************/ Dim sFile=&#x27;&#x27;&#x27;C:\tempFolder\data.ini&#x27;&#x27;&#x27; Dim sSection="RPA" Dim sKey="开发工具" INI.DeleteKey(sFile,sSection,sKey)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/INI_图片/INI_DeleteKey.png)  

---

## 删除小节

**说明**: 删除 INI 配置文件下的指定小节  

**原型**: `INI.DeleteSection(sFile, sSection)`  

**参数**:  
- **sFile** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - INI 配置文件所在路径  
- **sSection** (True) [string] 默认:"" - 要删除 INI 配置文件的小节名字  

**示例**:  
```
/***********************************删除小节************************************* 命令原型： INI.DeleteSection(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;, "") 入参： sFile--INI配置文件所在路径 sSection--要删除INI配置文件的小节名字 出参： 无 注意事项： 建议先判断该路径对应的文件是否存在，如果存在删除指定小节，不存在则会报错 ********************************************************************************/ Dim sFile=&#x27;&#x27;&#x27;C:\tempFolder\data.ini&#x27;&#x27;&#x27; Dim sSection="RPA" INI.DeleteSection(sFile, sSection)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/INI_图片/INI_DeleteSection.png)  

---

## 枚举键

**说明**: 枚举 INI 配置文件中指定小节下的所有键  

**原型**: `dictRet = INI.EnumKey(sFile, sSection)`  

**参数**:  
- **sFile** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - INI 配置文件所在路径  
- **sSection** (True) [string] 默认:"" - 要访问 INI 配置文件的小节名字  

**返回**: dictRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/***********************************枚举键************************************** 命令原型： dictRet = INI.EnumKey(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;, "") 入参： sFile--INI配置文件所在路径 sSection--要访问INI配置文件的小节名字 出参： dictRet--命令运行后的结果 注意事项： 建议先判断该路径对应的文件是否存在，如果存在枚举指定小节所有键，不存在则会报错 ********************************************************************************/ Dim sFile=&#x27;&#x27;&#x27;C:\tempFolder\data.ini&#x27;&#x27;&#x27; Dim sSection="RPA" dictRet = INI.EnumKey(sFile,sSection) TracePrint(dictRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/INI_图片/INI_EnumKey.png)  

---

## 枚举小节

**说明**: 枚举 INI 配置文件中的所有小节  

**原型**: `dictRet = INI.EnumSection(sFile)`  

**参数**:  
- **sFile** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - INI 配置文件所在路径  

**返回**: dictRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/***********************************枚举小节************************************* 命令原型： dictRet = INI.EnumSection(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： sFile--INI配置文件所在路径 出参： dictRet--命令运行后的结果 注意事项： 建议先判断该路径对应的文件是否存在，如果存在枚举所有小节，不存在则会报错 ********************************************************************************/ Dim sFile=&#x27;&#x27;&#x27;C:\tempFolder\data.ini&#x27;&#x27;&#x27; dictRet = INI.EnumSection(sFile) TracePrint(dictRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/INI_图片/INI_EnumSection.png)  

---

## 读键值

**说明**: 读取 INI 文件指定小节下的键值  

**原型**: `sRet = INI.Read(sFile, sSection, sKey, sDefault)`  

**参数**:  
- **sFile** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - INI 配置文件所在路径  
- **sSection** (True) [string] 默认:"" - 要访问 INI 配置文件的小节名字  
- **sKey** (True) [string] 默认:"" - 要访问 INI 配置文件的键名  
- **sDefault** (True) [string] 默认:"" - 当 INI 配置文件键名不存在时，返回的默认内容  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/***********************************读键值************************************** 命令原型： sRet = INI.Read(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;, "", "", "") 入参： sFile--INI配置文件所在路径 sSectuion--要访问INI配置文件的小节名字 sKey--要访问INI配置文件的键名 sDefault--当INI配置文件键名不存在时，返回的默认内容 出参： sRet--命令运行后的结果 注意事项： 建议先判断该路径对应的文件是否存在，如果存在读键值，不存在则会报错 ********************************************************************************/ Dim sFile=&#x27;&#x27;&#x27;C:\tempFolder\data.ini&#x27;&#x27;&#x27; Dim sSection="RPA" Dim sKey="开发工具" Dim sDefault="Null" sRet = INI.Read(sFile, sSection, sKey, sDefault) TracePrint sRet
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/INI_图片/INI_Read.png)  

---

## 写键值

**说明**: 在指定的 INI 文件中对指定的小节名下写入键值对，若指定的 INI 文件或小节名不存在则自动创建  

**原型**: `INI.Write(sFile, sSection, sKey, sValue)`  

**参数**:  
- **sFile** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - INI 配置文件所在路径  
- **sSection** (True) [string] 默认:"" - 要访问 INI 配置文件的小节名字  
- **sKey** (True) [string] 默认:"" - INI 文件中被写入的键值对中的键名，若为空字符串，则此键值对不被写入  
- **sValue** (True) [string] 默认:"" - INI 文件中被写入的键值对中的键值  

**示例**:  
```
/***********************************写键值************************************** 命令原型： INI.Write(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;, "", "", "") 入参： sFile--INI配置文件所在路径 sSectuion--要访问INI配置文件的小节名字 sKey--INI文件中被写入的键值对中的键名，若为空字符串，则此键值对不被写入 sValue--INI文件中被写入的键值对中的键值 出参： 无 注意事项： 如果该路径对应的文件不存在，会自动创建文件，再写入键值 ********************************************************************************/ Dim sFile=&#x27;&#x27;&#x27;C:\tempFolder\data.ini&#x27;&#x27;&#x27; Dim sSection="RPA" Dim sKey="控制面板" Dim sValue="UiBotCommander" INI.Write(sFile,sSection,sKey,sValue)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/INI_图片/INI_Write.png)  

---

## JSON字符串转换为对象

**说明**: 将JSON字符串转换成JSON对象  

**原型**: `objJSON = JSON.Parse(strJSON)`  

**参数**:  
- **strJSON** (True) [string] 默认:"" - 要转换成JSON的字符串  

**返回**: objJSON，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************JSON字符串转换为对象*************************************** 命令原型： objJSON = JSON.Parse(strJSON) 入参: strJSON--要转换成JSON的字符串 出参: objJSON--命令运行后的结果 **********************************************************************************/ Dim objJSON // JSON字符串转换为对象 objJSON = JSON.Parse("{\"test\" : 123 }") TracePrint(objJSON)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/JSON_图片/Parse.png)  

---

## 输入文本

**说明**: 自由输入文本  

**原型**: `Keyboard.Input(sText,optionArgs)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 输入的内容  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **sSimulate** (False) [enum] 默认:"message" - 可选择操作类型为：模拟操作(simulate)、系统消息(message)，默认选择：系统消息(message)  

**示例**:  
```
/*****************************输入文本******************************** 命令原型: Keyboard.Input(sText,optionArgs) 入参: sText--输入内容 optionArgs--可选参数(包括:执行后延时/执行前延时/操作类型).Type:Dict 出参： 无 注意事项: 模拟操作：指通过调用系统api mouseevent等实现鼠标操作，会实际移动光标。 系统消息：指发送鼠标消息到目标元素，不移动光标。 *********************************************************************/ Keyboard.Input("UiBot",{"iDelayAfter": 300, "iDelayBefore": 200, "sSimulate": "message"})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Keyboard_图片/Keyboard_Input.png)  

---

## 输入密码

**说明**: 输入密码  

**原型**: `Keyboard.InputPassword(password,optionArgs)`  

**参数**:  
- **password** (True) [string] 默认:"" - 对应要一次模拟输入的内容  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **sSimulate** (False) [enum] 默认:"message" - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：系统消息(message)  

**示例**:  
```
/*****************************输入密码******************************** 命令原型: Keyboard.InputPassword(password,optionArgs) 入参: password--密码 optionArgs--可选参数(包括:执行后延时/执行前延时/操作类型).Type:Dict 出参： 无 注意事项: 模拟操作：指通过调用系统api mouseevent等实现鼠标操作，会实际移动光标。 系统消息：指发送鼠标消息到目标元素，不移动光标。 后台操作：可以理解为调用了一次元素的鼠标响应回调函数。 *********************************************************************/ Keyboard.InputPassword("gicCYW79cWCCaM5irWDnsQ==",{"iDelayAfter": 300, "iDelayBefore": 200, "sSimulate": "message"})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Keyboard_图片/Keyboard_InputPassword.png)  

---

## 在目标中输入密码

**说明**: 在指定的界面元素中输入密码  

**原型**: `Keyboard.InputPwd(objUiElement,sPwd,bEmptyField,iDelayBetweenKeys,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **sPwd** (True) [string] 默认:"" - 要在指定的界面元素中写入的文本  
- **bEmptyField** (True) [boolean] 默认:True - 写入文本之前是否清空输入框  
- **iDelayBetweenKeys** (True) [number] 默认:20 - 两次输入之间的间隔，（仅在“操作类型”属性中选项为”模拟操作“时生效）低于20毫秒时会自动转为20毫秒，间隔的值过小有可能导致输入时丢字，与机器的性能有关  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:500 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是500毫秒，过短可能导致漏输入情况  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sSimulate** (False) [enum] 默认:"message" - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：系统消息(message)  
- **bClickBeforeInput** (False) [boolean] 默认:False - 找到目标后先点击目标再输入内容  

**示例**:  
```
/*****************************在目标输入密码******************************** 命令原型: Keyboard.InputPwd(objUiElement,sPwd,bEmptyField,iDelayBetweenKeys,iTimeOut,optionArgs) 入参: objUiElement--目标元素 sPwd--密码 bEmptyField--清空原内容 iDelayBetweenKeys--键入间隔(ms) iTimeOut--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时/激活窗口/辅助按键/操作类型/输入前点击).Type:Dict 出参： 无 注意事项: 模拟操作：指通过调用系统api mouseevent等实现鼠标操作，会实际移动光标。 系统消息：指发送鼠标消息到目标元素，不移动光标。 后台操作：可以理解为调用了一次元素的鼠标响应回调函数。 *********************************************************************/ Keyboard.InputPwd(@ui"输入控件<input>2","gicCYW79cWCCaM5irWDnsQ==",true,20,10000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 500, "bSetForeground": true, "sSimulate": "message", "bValidate": false, "bClickBeforeInput": false})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Keyboard_图片/Keyboard_InputPwd.png)  

---

## 在目标中输入

**说明**: 在指定的界面元素中输入文本  

**原型**: `Keyboard.InputText(objUiElement,sText,bEmptyField,iDelayBetweenKeys,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **sText** (True) [string] 默认:"" - 要在指定的界面元素中写入的文本  
- **bEmptyField** (True) [boolean] 默认:True - 写入文本之前是否清空输入框  
- **iDelayBetweenKeys** (True) [number] 默认:20 - 两次输入之间的间隔，（仅在“操作类型”属性中选项为”模拟操作“时生效）低于20毫秒时会自动转为20毫秒，间隔的值过小有可能导致输入时丢字，与机器的性能有关  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:500 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是500毫秒，过短可能导致漏输入情况  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sSimulate** (False) [enum] 默认:"message" - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：系统消息(message)  
- **bValidate** (False) [boolean] 默认:False - 将 “写入文本”属性内容与实际输入内容进行比较，内容相同继续运行，内容不同抛出异常  
- **bClickBeforeInput** (False) [boolean] 默认:False - 找到目标后先点击目标再输入内容  

**示例**:  
```
/******************************在目标中输入******************************* 命令原型: Keyboard.InputText(objUiElement,sText,bEmptyField,iDelayBetweenKeys,iTimeOut,optionArgs) 入参: objUiElement--目标元素 sText--写入文本 bEmptyField--清空原内容 iDelayBetweenKeys--键入间隔(ms) iTimeOut--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时/激活窗口/操作类型/验证写入文本/输入前点击).Type:Dict 出参： 无 注意事项: 模拟操作：指通过调用系统api mouseevent等实现鼠标操作，会实际移动光标。 系统消息：指发送鼠标消息到目标元素，不移动光标。 后台操作：可以理解为调用了一次元素的鼠标响应回调函数。 *********************************************************************/ Keyboard.InputText(@ui"输入控件<input>","UiBot",true,20,10000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 500, "bSetForeground": true, "sSimulate": "message", "bValidate": false, "bClickBeforeInput": false})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Keyboard_图片/Keyboard_InputText.png)  

---

## 模拟按键

**说明**: 模拟键盘按键  

**原型**: `Keyboard.Press(sKey, sType, sKeyModifiers,optionArgs)`  

**参数**:  
- **sKey** (True) [enum] 默认:"Enter" - 对应要一次模拟输入的内容  
- **sType** (True) [enum] 默认:"press" - 按键的类型为：单击(press)、按下(down)、弹起(up)  
- **sKeyModifiers** (True) [set] 默认:[] - 触发按键动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **sSimulate** (False) [enum] 默认:"simulate" - 可选择操作类型为：模拟操作(simulate)、系统消息(message)，默认选择：模拟操作(simulate)  

**示例**:  
```
/******************************模拟按键******************************* 命令原型: Keyboard.Press(sKey, sType, sKeyModifiers,optionArgs) 入参: sKey--模拟按键 sType--按键类型(单击/双击/按下/弹起) sKeyModifiers--辅助按键 optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时/操作类型).Type:Dict 出参： 无 注意事项: 模拟操作：指通过调用系统api mouseevent等实现鼠标操作，会实际移动光标。 系统消息：指发送鼠标消息到目标元素，不移动光标。 *********************************************************************/ Keyboard.Press("Enter", "press", [],{"iDelayAfter": 300, "iDelayBefore": 200, "sSimulate": "simulate"})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Keyboard_图片/Keyboard_Press.png)  

---

## 在目标中按键

**说明**: 在指定的界面元素中输入按键  

**原型**: `Keyboard.PressKey(objUiElement,sKey,iDelayBetweenKeys,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **sKey** (True) [enum] 默认:"Enter" - 对应要一次模拟输入的内容  
- **iDelayBetweenKeys** (True) [number] 默认:20 - 两次输入之间的间隔，（仅在“操作类型”属性中选项为”模拟操作“时生效）低于20毫秒时会自动转为20毫秒，间隔的值过小有可能导致输入时丢字，与机器的性能有关  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:500 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是500毫秒，过短可能导致漏输入情况  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sKeyModifiers** (False) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **sSimulate** (False) [enum] 默认:"simulate" - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：系统消息(message)  
- **bClickBeforeInput** (False) [boolean] 默认:False - 找到目标后先点击目标再输入内容  

**示例**:  
```
/******************************在目标中按键******************************* 命令原型: Keyboard.PressKey(objUiElement,sKey,iDelayBetweenKeys,iTimeOut,optionArgs) 入参: objUiElement--目标元素 sKey--模拟按键 iDelayBetweenKeys--键入间隔(ms) iTimeOut--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时/激活窗口/辅助按键/操作类型/输入前点击).Type:Dict 出参： 无 注意事项: 模拟操作：指通过调用系统api mouseevent等实现鼠标操作，会实际移动光标。 系统消息：指发送鼠标消息到目标元素，不移动光标。 后台操作：可以理解为调用了一次元素的鼠标响应回调函数。 *********************************************************************/ // 在指定目标中输入"UiBot" Keyboard.InputText(@ui"输入控件<input>","UiBot",true,20,10000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 500, "bSetForeground": true, "sSimulate": "message", "bValidate": false, "bClickBeforeInput": false}) // 模拟发送"enter"按键 Keyboard.PressKey(@ui"输入控件<input>1","Enter",20,10000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200, "bSetForeground": true, "sSimulate": "simulate", "sKeyModifiers": [], "bClickBeforeInput": false})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Keyboard_图片/Keyboard_PressKey.png)  

---

## 输入文本

**说明**: 使用UiBot KeyBox输入文本，仅支持英文、数字和半角字符  

**原型**: `KeyBox.Input(uuid,sText,Delay)`  

**参数**:  
- **uuid** (True) [string] 默认:"" - UiBot KeyBox的唯一标识，用于区分不同的UiBot KeyBox  
- **sText** (True) [string] 默认:"" - 使用UiBot KeyBox输入的内容  
- **Delay** (True) [number] 默认:100 - 两次输入之间的延时间隔，间隔过小可能导致输入时丢字，与系统性能有关  

**示例**:  
```
/***************************输入文本********************************** 命令原型: KeyBox.Input(uuid,sText,Delay) 入参： uuid--设备号 sText--输入内容 Delay--键入间隔(ms) 出参： 无 注意事项: 无 *********************************************************************/ KeyBox.Input("20000008","123",100)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/KeyBox_图片/KeyBox_Input.png)  

---

## 鼠标点击OCR文本

**说明**: 使用本地OCR对窗口范围内进行指定文字识别，如果识别到指定文字就点击它。调用时不需要访问网络，没有调用频次的限制，但需要消耗一定的系统资源  

**原型**: `LocalOCR.Click(objUiElement,objRect,sText,iRule,iOccurrence,iButton,iType,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **objRect** (True) [dictionary] 默认:{ "x":0,"y":0,"width":0,"height":0 } - 需要进行OCR文字识别的范围，程序会在控件这个范围内进行文字识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的文字识别  
- **sText** (True) [string] 默认:"" - 查找元素时使用的文本  
- **iRule** (True) [enum] 默认:"instr" - 查找文本时使用的规则  
- **iOccurrence** (True) [number] 默认:1 - 如果“文本”字段中的字符串在指示的界面元素中出现多次，请在此处指定要单击的出现次数。例如，如果字符串出现4次并且您要单击第一个匹配项，请在此字段中写入1  
- **iButton** (True) [enum] 默认:"left" - 鼠标按键 { left:左键, right:右键, middle:中键 }  
- **iType** (True) [enum] 默认:"click" - 点击类型 { click:单击, dbclick:双击, down:按下, up:弹起 }  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sCursorPosition** (False) [enum] 默认:"Center" - 描述添加OffsetX和OffsetY属性的偏移量的光标起点。可以使用以下选项：TopLeft，TopRight，BottomLeft，BottomRight和Center。默认选项是Center  
- **iCursorOffsetX** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的水平位移  
- **iCursorOffsetY** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的垂直位移  
- **sKeyModifiers** (False) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **sSimulate** (False) [enum] 默认:"simulate" - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：模拟操作(simulate)  

**示例**:  
```
/*********************************鼠标点击OCR文本*************************************** 命令原型： LocalOCR.Click(@ui"",{"x":0,"y":0,"width":0,"height":0},"","instr",1,"left","click",10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate"}) 入参： objUiElement--识别目标。 objRect--目标识别范围。 sText--查找文本。注：查找元素时使用的文本 iRule--查找规则。注：查找文本时使用的规则 iOccurrence--相似结果位置。 iButton--鼠标点击。注：鼠标按键 {left:左键, right:右键, middle:中键} iType--点击类型。注：点击类型 {click:单击, dbclick:双击, down:按下, up:弹起} iTimeOut--超时时间。注：指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒） optionArgs--可选参数(包括:错误继续执行、执行后延时、执行前延时、激活窗口、光标位置、横坐标偏移、纵坐标偏移、辅助按键、操作类型).Type:Dict 注意事项： 1.要保证操作的页面为打开状态，否则会报错。 2.调用时不需要使用外部网络，但是需要消耗一部分系统资源。 ********************************************************************************/ LocalOCR.Click(@ui"文本<span>_鼠标点击OCR文本1",{"x":0,"y":0,"width":0,"height":0},"文本","instr",1,"left","click",10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate"})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/LocalOCR_图片/LocalOCR_Click.png)  

---

## 查找OCR文本位置

**说明**: 使用本地OCR在窗口范围内查找指定文本的位置，成功则返回字典类型的文本位置，失败则引发异常。调用时不需要访问网络，没有调用频次的限制，但需要消耗一定的系统资源  

**原型**: `objPoint = LocalOCR.Find(objElement, objRect,sText, sRule, iOccurrence, iTimeOut, optionArgs)`  

**参数**:  
- **objElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **objRect** (True) [dictionary] 默认:{ "x":0,"y":0,"width":0,"height":0 } - 需要进行OCR文字识别的范围，程序会在控件这个范围内进行文字识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的文字识别  
- **sText** (True) [string] 默认:"" - 查找元素时使用的文本  
- **sRule** (True) [enum] 默认:"instr" - 查找文本时使用的规则  
- **iOccurrence** (True) [number] 默认:1 - 如果“文本”字段中的字符串在指示的界面元素中出现多次，请在此处指定要单击的出现次数。例如，如果字符串出现4次并且您要单击第一个匹配项，请在此字段中写入1  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  

**返回**: objPoint，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************查找OCR文本位置*************************************** 命令原型： objPoint = LocalOCR.Find(@ui"", {"x":0,"y":0,"width":0,"height":0},"", "instr", 1, 10000, {"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true}) 入参： objUiElement--识别目标。 objRect--目标识别范围。默认值:{"x":0,"y":0,"width":0,"height":0} sText--查找文本。注：查找元素时使用的文本 sRule--查找规则。注：查找文本时使用的规则 iOccurrence--相似结果位置。 iTimeOut--超时时间。注：指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒） optionArgs--可选参数(包括:错误继续执行、执行后延时、执行前延时、激活窗口).Type:Dict 出参： objPoint--函数调用的输出保存到的变量。 注意事项： 1.要保证操作的页面为打开状态，否则会报错。 2.调用时不需要使用外部网络，但是需要消耗一部分系统资源。 ********************************************************************************/ Dim objPoint = "" objPoint = LocalOCR.Find(@ui"文本<span>_查找OCR文本位置2", {"x":0,"y":0,"width":0,"height":0},"文本", "instr", 1, 10000, {"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true}) TracePrint(objPoint)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/LocalOCR_图片/LocalOCR_Find.png)  

---

## 鼠标移动到OCR文本上

**说明**: 使用本地OCR对窗口范围内进行指定文字识别，如果识别到指定文字将光标移动到文本所在的位置。调用时不需要访问网络，没有调用频次的限制，但需要消耗一定的系统资源  

**原型**: `LocalOCR.Hover(objUiElement,objRect,sText,sRule,iOccurrence,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **objRect** (True) [dictionary] 默认:{ "x":0,"y":0,"width":0,"height":0 } - 需要进行OCR文字识别的范围，程序会在控件这个范围内进行文字识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的文字识别  
- **sText** (True) [string] 默认:"" - 查找元素时使用的文本  
- **sRule** (True) [enum] 默认:"instr" - 查找文本时使用的规则  
- **iOccurrence** (True) [number] 默认:1 - 如果“文本”字段中的字符串在指示的界面元素中出现多次，请在此处指定要单击的出现次数。例如，如果字符串出现4次并且您要单击第一个匹配项，请在此字段中写入1  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sCursorPosition** (False) [enum] 默认:"Center" - 描述添加OffsetX和OffsetY属性的偏移量的光标起点。可以使用以下选项：TopLeft，TopRight，BottomLeft，BottomRight和Center。默认选项是Center  
- **iCursorOffsetX** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的水平位移  
- **iCursorOffsetY** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的垂直位移  
- **sKeyModifiers** (False) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **sSimulate** (False) [enum] 默认:"simulate" - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：模拟操作(simulate)  

**示例**:  
```
/*********************************鼠标移动到OCR文本上*************************************** 命令原型： LocalOCR.Hover(@ui"",{"x":0,"y":0,"width":0,"height":0},"","instr",1,10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate"}) 入参： objUiElement--识别目标。 objRect--目标识别范围。默认值:{"x":0,"y":0,"width":0,"height":0} sText--查找文本。注：查找元素时使用的文本 sRule--查找规则。注：查找文本时使用的规则 iOccurrence--相似结果位置。 iTimeOut--超时时间。注：指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒） optionArgs--选填参数。注：该参数为多个选填参数，其中包括错误继续执行、执行后延时、执行前延时、激活窗口、光标位置、横坐标偏移、纵坐标偏移、辅助按键、操作类型 注意事项： 1.要保证操作的页面为打开状态，否则会报错。 2.调用时不需要使用外部网络，但是需要消耗一部分系统资源。 3.调用时限没有限制，但是因为占用系统资源，所以在使用时尽量设置延时调用。 ********************************************************************************/ LocalOCR.Hover(@ui"文本<span>_鼠标移动到OCR文本上1",{"x":0,"y":0,"width":0,"height":0},"文本","instr",1,10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate"})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/LocalOCR_图片/LocalOCR_Hover.png)  

---

## 图像OCR识别

**说明**: 使用本地OCR识别指定图像文件的文本内容。调用时不需要访问网络，没有调用频次的限制，但需要消耗一定的系统资源  

**原型**: `sText = LocalOCR.ImageOCR(sFileName,sOcrType, iTimeOut)`  

**参数**:  
- **sFileName** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要识别的图片路径，支持jpg、jpeg、gif、bmp、png格式  
- **sOcrType** (True) [enum] 默认:"SceneText" - 待识别的文本类型  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  

**返回**: sText，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************图像OCR识别*************************************** 命令原型： sText = LocalOCR.ImageOCR(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"SceneText", 10000) 入参： sFileName--识别文件的位置。注：要识别的图片路径 sOcrType--文本场景。注：待识别的文本类型 iTimeOut--超时时间。注：指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒） 出参： sText--函数调用的输出保存到的变量。 注意事项： 1.调用时不需要使用外部网络，但是需要消耗一部分系统资源。 2.要保证读取的图片存在于本地。 3.注意文本类型的选择。 4.调用时限没有限制，但是因为占用系统资源，所以在使用时尽量设置延时调用。 ********************************************************************************/ Dim sText = "" sText = LocalOCR.ImageOCR(@res"1643099897(1).jpg","SceneText", 10000) TracePrint(sText)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/LocalOCR_图片/LocalOCR_ImageOCR.png)  

---

## 屏幕OCR识别

**说明**: 使用本地OCR识别屏幕指定窗口范围内的文本内容。调用时不需要访问网络，没有调用频次的限制，但需要消耗一定的系统资源  

**原型**: `sText = LocalOCR.ScreenOCR(objElement,objRect,sOcrType,iTimeOut,optionArgs)`  

**参数**:  
- **objElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **objRect** (True) [dictionary] 默认:{ "x":0,"y":0,"width":0,"height":0 } - 需要进行OCR文字识别的范围，程序会在控件这个范围内进行文字识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的文字识别  
- **sOcrType** (True) [enum] 默认:"SceneText" - 待识别的文本类型  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  

**返回**: sText，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************屏幕OCR识别*************************************** 命令原型： sText = LocalOCR.ScreenOCR(@ui"",{"x":0,"y":0,"width":0,"height":0},"SceneText",10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true}) 入参： objUiElement--识别目标。 objRect--目标识别范围。默认值:{"x":0,"y":0,"width":0,"height":0} sOcrType--文本场景。注：待识别的文本类型 iTimeOut--超时时间。注：指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒） optionArgs--可选参数(包括:错误继续执行、执行后延时、执行前延时、激活窗口).Type:Dict 出参： sText--函数调用的输出保存到的变量。 注意事项： 1.调用时不需要使用外部网络，但是需要消耗一部分系统资源。 2.注意文本类型的选择。 3.调用时限没有限制，但是因为占用系统资源，所以在使用时尽量设置延时调用。 ********************************************************************************/ Dim sText = "" sText = LocalOCR.ScreenOCR(@ui"文本<span>_屏幕OCR识别",{"x":0,"y":0,"width":0,"height":0},"SceneText",10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true}) TracePrint(sText)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/LocalOCR_图片/LocalOCR_ScreenOCR.png)  

---

## 写入调试日志

**说明**: 写入调试(Debug)信息至日志文件中  

**原型**: `Log.Debug(sContent)`  

**参数**:  
- **sContent** (True) [string] 默认:"" - 要输出的具体信息  

**示例**:  
```
/*********************************写入调试日志*************************************** 命令原型： Log.Debug(sContent) 入参： content -- 日志内容 出参： 无 注意事项： 无 ********************************************************************************/ Log.Debug("这是一条调试日志，在流程调试时记录调试信息") TracePrint("这是一条级别为debug的日志，在流程调试运行时使用")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Log_图片/Log_Debug.png)  

---

## 写入错误日志

**说明**: 写入错误日志  

**原型**: `Log.Error(sContent)`  

**参数**:  
- **sContent** (True) [string] 默认:"" - 要输出的具体信息  

**示例**:  
```
/*********************************输出错误日志*************************************** 命令原型： Log.Error(sContent) 入参： content -- 日志内容 出参： 无 注意事项： 无 ********************************************************************************/ Log.Error("这是一条错误日志，请在流程运行未达到预期，需要停止时使用") TracePrint("这是一条级别为error的日志，在流程出现错误，需要立即停止时使用")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Log_图片/Log_Error.png)  

---

## 写入一般日志信息

**说明**: 写入一般日志信息  

**原型**: `Log.Info(sContent)`  

**参数**:  
- **sContent** (True) [string] 默认:"" - 要输出的具体信息  

**示例**:  
```
/*********************************输出一般日志信息*************************************** 命令原型： Log.Info(sContent) 入参： content -- 日志内容 出参： 无 注意事项： 一般的日志信息，标明流程运行到哪个阶段 ********************************************************************************/ Log.info("这是一条一般日志，请在流程符合预期时用做与标记流程运行到哪个阶段") TracePrint("这是一条级别为info的日志，在流程正常运行，做为标记流程运行到哪一个阶段时使用")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Log_图片/Log_Info.png)  

---

## 设置日志级别

**说明**: 设置某个级别以上的日志才需要输出  

**原型**: `Log.SetLevel(iLevel)`  

**参数**:  
- **iLevel** (True) [enum] 默认:2 - 整数型，可以取0至3，分别代表错误、警告、一般信息和调试信息。当设为0时，只输出错误；当设为1时，输出错误和警告；当设为2时，输出一般信息、错误和警告；当设为3时，除了一般信息、错误和警告之外，用TracePrint输出的调试信息也会同时输出到日志  

**示例**:  
```
/*********************************设置日志级别*************************************** 命令原型： Log.SetLevel(iLevel) 入参： level -- 日志级别 出参： 无 注意事项： 设置日志级别，在某个级别以上日志才会被显示 ********************************************************************************/ Log.SetLevel(2) TracePrint("日志级别设置为2级，即只有info,warning，error三个级别会显示")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Log_图片/Log_SetLevel.png)  

---

## 写入警告日志

**说明**: 写入警告日志  

**原型**: `Log.Warn(sContent)`  

**参数**:  
- **sContent** (True) [string] 默认:"" - 要输出的具体信息  

**示例**:  
```
/*********************************输出警告日志*************************************** 命令原型： Log.Warn(sContent) 入参： content -- 日志内容 出参： 无 注意事项： 当流程出现预期外的情况，但不影响流程运行时使用 ********************************************************************************/ Log.Warn("这是一条警告日志，请在流程出现一些预期内的意外，但不影响流程运行的时候使用") TracePrint("这是一条级别为warning的日志，在出现预期内的意外，但不影响流程运行时使用")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Log_图片/Log_Warn.png)  

---

## 提取地址信息

**说明**: 先循环遍历地址标准化命令返回结果，然后从遍历结果中提取指定类型的地址信息  

**原型**: `sRet = Mage.ExtractAddress(value, type)`  

**参数**:  
- **value** (True) [expression] 默认:value - 使用循环遍历标准化地址结果的值  
- **type** (True) [enum] 默认:"whole_address" - 选择提取的地址类型  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************提取地址信息********************** 命令原型: sRet = Mage.ExtractAddress(value, "whole_address") 入参: value--地址标准化结果 type--地址类型 出参: sRet--函数调用的输出保存到的变量 注意事项： 需要配合地址标准化命令（NLPAddressStandard）输出结果使用 ****************************************************/ Rem 测试数据 Dim value = {"address" : "","ai_function" : "nlp_addr_std","city" : "","district" : "浦东新区","length" : 0,"poi_name" : "","province" : "上海市","start_pos" : 0,"subdistrict" : "张江镇"} Dim sRet="" // 输出结果 sRet = Mage.ExtractAddress(value, "whole_address") Traceprint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageAddressStd_图片/Mage_ExtractAddress.png)  

---

## 地址标准化

**说明**: 将地址进行标准化，支持输入多个地址，以\n隔开，返回数组  

**原型**: `arrayRet = Mage.NLPAddressStandard(address,config,time)`  

**参数**:  
- **address** (True) [expression] 默认:"" - 待标准化地址的信息，支持输入多个地址，以\n隔开  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************地址标准化********************** 命令原型: arrayRet = Mage.NLPAddressStandard("",{},30000) 入参: address--待标准化地址 config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int 出参: arrayRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim address="上海市浦东新区张江镇" // 待标准化地址 Dim arrayRet="" // 输出结果 arrayRet = Mage.NLPAddressStandard(address,{"Pubkey":"KJ55tRgUjGxJAJ20SzpNtza0","Secret":"urfWfYHwQhyR8uVkpTWhjjEX6gpZKRyg","Url":"https://mage.uibot.com.cn"},30000) Traceprint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageAddressStd_图片/Mage_NLPAddressStandard.png)  

---

## 获取票据内容

**说明**: 获取通用多票据识别结果中的票据内容  

**原型**: `sRet = Mage.ExtractInvoiceInfo(value,invoice_type,invoice_key)`  

**参数**:  
- **value** (True) [expression] 默认:value - 使用"屏幕多票据识别"、"图像多票据识别"、"PDF多票据识别"等命令输出到的变量并循环遍历的值  
- **invoice_type** (True) [enum] 默认:"" - 选择需要获取的票据类型  
- **invoice_key** (True) [enum] 默认:"" - 选择获取票据类型下的字段  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取票据内容********************** 命令原型: sRet = Mage.ExtractInvoiceInfo(value,"","") 入参: value--票据识别结果 invoice_type--提取类型 invoice_key--提取字段 出参: sRet--函数调用的输出保存到的变量 ****************************************************/ Rem 测试数据 Dim value = {"ai_function" : "ocr_invoice","class" : 3,"goods" : [],"image_angle" : 0,"items" : [{"description" : "火车票红色编码","key" : "ticket_number","positions" : [{"x" : 46,"y" : 25},{"x" : 195,"y" : 25},{"x" : 195,"y" : 59},{"x" : 46,"y" : 59}],"value" : "H014366"},{"description" : "检票口","key" : "boarding_gate","positions" : [{"x" : 0,"y" : 0},{"x" : 0,"y" : 0},{"x" : 0,"y" : 0},{"x" : 0,"y" : 0}],"value" : ""},{"description" : "检票口","key" : "check","positions" : [{"x" : 0,"y" : 0},{"x" : 0,"y" : 0},{"x" : 0,"y" : 0},{"x" : 0,"y" : 0}],"value" : ""},{"description" : "出发地","key" : "departure_station","positions" : [{"x" : 70,"y" : 67},{"x" : 203,"y" : 67},{"x" : 203,"y" : 108},{"x" : 70,"y" : 108}],"value" : "北京北站"},{"description" : "车次号","key" : "train_number","positions" : [{"x" : 257,"y" : 74},{"x" : 374,"y" : 74},{"x" : 374,"y" : 109},{"x" : 257,"y" : 109}],"value" : "G9103"},{"description" : "目的地","key" : "arrival_station","positions" : [{"x" : 391,"y" : 67},{"x" : 523,"y" : 67},{"x" : 523,"y" : 108},{"x" : 391,"y" : 108}],"value" : "张家口站"},{"description" : "乘车时间","key" : "departure_date","positions" : [{"x" : 49,"y" : 133},{"x" : 343,"y" : 133},{"x" : 343,"y" : 165},{"x" : 49,"y" : 165}],"value" : "2019-12-30 17:53"},{"description" : "座位号","key" : "seat_number","positions" : [{"x" : 382,"y" : 135},{"x" : 505,"y" : 135},{"x" : 505,"y" : 164},{"x" : 382,"y" : 164}],"value" : "05车06C号"},{"description" : "价格","key" : "price","positions" : [{"x" : 82,"y" : 173},{"x" : 151,"y" : 172},{"x" : 151,"y" : 195},{"x" : 82,"y" : 195}],"value" : "110.0"},{"description" : "座位类别","key" : "class","positions" : [{"x" : 426,"y" : 167},{"x" : 531,"y" : 167},{"x" : 531,"y" : 196},{"x" : 426,"y" : 196}],"value" : "多功能座"},{"description" : "乘客身份证","key" : "passenger_id","positions" : [{"x" : 51,"y" : 260},{"x" : 398,"y" : 260},{"x" : 398,"y" : 293},{"x" : 51,"y" : 293}],"value" : "4330261954****0012"},{"description" : "乘客名称","key" : "passenger_name","positions" : [{"x" : 51,"y" : 260},{"x" : 398,"y" : 260},{"x" : 398,"y" : 293},{"x" : 51,"y" : 293}],"value" : "刘建华"},{"description" : "火车票ID","key" : "ticket_id","positions" : [{"x" : 57,"y" : 360},{"x" : 399,"y" : 360},{"x" : 399,"y" : 389},{"x" : 57,"y" : 389}],"value" : "65678301011231H014366"},{"description" : "发票代码","key" : "code","positions" : [{"x" : 57,"y" : 360},{"x" : 399,"y" : 360},{"x" : 399,"y" : 389},{"x" : 57,"y" : 389}],"value" : "65678301011231"},{"description" : "火车票红色编码","key" : "number","positions" : [{"x" : 57,"y" : 360},{"x" : 399,"y" : 360},{"x" : 399,"y" : 389},{"x" : 57,"y" : 389}],"value" : "H014366"}],"kind" : 2,"page_number" : 1,"rotated_image_height" : 408,"rotated_image_width" : 632,"type" : 20,"type_description" : "火车票","type_key" : "train_ticket"} // 测试数据 Dim sRet="" // 输出结果 sRet = Mage.ExtractInvoiceInfo(value,"train_ticket","passenger_name") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageBill_图片/Mage_ExtractInvoiceInfo.png)  

---

## 获取票据类型

**说明**: 获取通用多票据识别结果中的票据类型  

**原型**: `sRet = Mage.ExtractInvoiceType(value)`  

**参数**:  
- **value** (True) [expression] 默认:value - 使用"屏幕多票据识别"、"图像多票据识别"、"PDF多票据识别"等命令输出到的变量并循环遍历的值  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取票据类型********************** 命令原型: sRet = Mage.ExtractInvoiceType(value) 入参: value--票据识别结果 出参: sRet--函数调用的输出保存到的变量 ****************************************************/ Rem 测试数据 Dim value = {"ai_function" : "ocr_invoice","class" : 3,"goods" : [],"image_angle" : 0,"items" : [{"description" : "火车票红色编码","key" : "ticket_number","positions" : [{"x" : 46,"y" : 25},{"x" : 195,"y" : 25},{"x" : 195,"y" : 59},{"x" : 46,"y" : 59}],"value" : "H014366"},{"description" : "检票口","key" : "boarding_gate","positions" : [{"x" : 0,"y" : 0},{"x" : 0,"y" : 0},{"x" : 0,"y" : 0},{"x" : 0,"y" : 0}],"value" : ""},{"description" : "检票口","key" : "check","positions" : [{"x" : 0,"y" : 0},{"x" : 0,"y" : 0},{"x" : 0,"y" : 0},{"x" : 0,"y" : 0}],"value" : ""},{"description" : "出发地","key" : "departure_station","positions" : [{"x" : 70,"y" : 67},{"x" : 203,"y" : 67},{"x" : 203,"y" : 108},{"x" : 70,"y" : 108}],"value" : "北京北站"},{"description" : "车次号","key" : "train_number","positions" : [{"x" : 257,"y" : 74},{"x" : 374,"y" : 74},{"x" : 374,"y" : 109},{"x" : 257,"y" : 109}],"value" : "G9103"},{"description" : "目的地","key" : "arrival_station","positions" : [{"x" : 391,"y" : 67},{"x" : 523,"y" : 67},{"x" : 523,"y" : 108},{"x" : 391,"y" : 108}],"value" : "张家口站"},{"description" : "乘车时间","key" : "departure_date","positions" : [{"x" : 49,"y" : 133},{"x" : 343,"y" : 133},{"x" : 343,"y" : 165},{"x" : 49,"y" : 165}],"value" : "2019-12-30 17:53"},{"description" : "座位号","key" : "seat_number","positions" : [{"x" : 382,"y" : 135},{"x" : 505,"y" : 135},{"x" : 505,"y" : 164},{"x" : 382,"y" : 164}],"value" : "05车06C号"},{"description" : "价格","key" : "price","positions" : [{"x" : 82,"y" : 173},{"x" : 151,"y" : 172},{"x" : 151,"y" : 195},{"x" : 82,"y" : 195}],"value" : "110.0"},{"description" : "座位类别","key" : "class","positions" : [{"x" : 426,"y" : 167},{"x" : 531,"y" : 167},{"x" : 531,"y" : 196},{"x" : 426,"y" : 196}],"value" : "多功能座"},{"description" : "乘客身份证","key" : "passenger_id","positions" : [{"x" : 51,"y" : 260},{"x" : 398,"y" : 260},{"x" : 398,"y" : 293},{"x" : 51,"y" : 293}],"value" : "4330261954****0012"},{"description" : "乘客名称","key" : "passenger_name","positions" : [{"x" : 51,"y" : 260},{"x" : 398,"y" : 260},{"x" : 398,"y" : 293},{"x" : 51,"y" : 293}],"value" : "刘建华"},{"description" : "火车票ID","key" : "ticket_id","positions" : [{"x" : 57,"y" : 360},{"x" : 399,"y" : 360},{"x" : 399,"y" : 389},{"x" : 57,"y" : 389}],"value" : "65678301011231H014366"},{"description" : "发票代码","key" : "code","positions" : [{"x" : 57,"y" : 360},{"x" : 399,"y" : 360},{"x" : 399,"y" : 389},{"x" : 57,"y" : 389}],"value" : "65678301011231"},{"description" : "火车票红色编码","key" : "number","positions" : [{"x" : 57,"y" : 360},{"x" : 399,"y" : 360},{"x" : 399,"y" : 389},{"x" : 57,"y" : 389}],"value" : "H014366"}],"kind" : 2,"page_number" : 1,"rotated_image_height" : 408,"rotated_image_width" : 632,"type" : 20,"type_description" : "火车票","type_key" : "train_ticket"} // 测试数据 Dim sRet="" // 输出结果 sRet = Mage.ExtractInvoiceType(value) Traceprint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageBill_图片/Mage_ExtractInvoiceType.png)  

---

## 图像多票据识别

**说明**: 使用 Laiye Intelligent Document Processing 识别指定图像的多种票据，识别结果返回数组  

**原型**: `arrayRet = Mage.ImageOCRInvoice(path,config,time)`  

**参数**:  
- **path** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要识别图像的路径，支持 jpeg、jpg、png、bmp、tif、tiff 等格式，PDF格式仅支持首页识别  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************图像多票据识别********************** 命令原型: arrayRet = Mage.ImageOCRInvoice(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,{},30000) 入参: path--待识别票据路径 config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int 出参: arrayRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim path=&#x27;&#x27;&#x27;&#x27;&#x27;&#x27; // 待识别票据路径 Dim arrayRet="" // 输出结果 arrayRet = Mage.ImageOCRInvoice(path,{"Pubkey":"UsCqnK8TOHBosdYxkuC6Zmop","Secret":"TAfkRueOzvUeJy47jE7xEpjJ6qapyA4u","Url":"https://demo.laiye.com:8082"},30000) Traceprint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageBill_图片/Mage_ImageOCRInvoice.png)  

---

## PDF多票据识别

**说明**: 将 PDF 指定的页码通过 Laiye Intelligent Document Processing 通用多票据识别，返回结果数组。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额  

**原型**: `arrayRet = Mage.PDFOCRInvoice(config, path,password,all_pg_state,page_cfg,sleepTime,time)`  

**参数**:  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **path** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - PDF文件路径  
- **password** (True) [string] 默认:"" - PDF文件密码，无密码不需要填写  
- **all_pg_state** (True) [boolean] 默认:False - 当全部页码设为"是"，则识别全部且指定页码输入无效。设为否时，可指定页码识别  
- **page_cfg** (True) [expression] 默认:[ [1,2] ] - 支持正整数和数组格式，如输入2，则识别第2页；如输入 [1,3,5] ，则识别第1,3,5页；如输入[1, [6,9] ,4]，则识别1,4页和第6到第9页。当识别全部页码设为"是"，则识别指定页码的输入失效。超出PDF页码总数的部分会报错，页码重叠部分仅识别1次  
- **sleepTime** (True) [number] 默认:10000 - 识别PDF每页的间隔时长（以毫秒为单位），默认10000毫秒(10秒)。识别页数较多，间隔较短可能会导致调用频率超限错误  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************PDF多票据识别********************** 命令原型: arrayRet = Mage.PDFOCRInvoice({}, &#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"",false,[[1,2]],10000,30000) 入参: config--mage配置,需配置Pubkey和Secret.Type:Dict path--待识别包含票据的PDF文件路径 password--密码 all_pg_state--是否识别全部页 page_cfg--识别指定页码 sleepTime--间隔时间.默认单位:毫秒.Type:Int time--超时时间.默认单位:毫秒.Type:Int 出参: arrayRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim path=&#x27;&#x27;&#x27;&#x27;&#x27;&#x27; // 待识别包含票据的PDF文件路径 Dim arrayRet="" // 输出结果 arrayRet = Mage.PDFOCRInvoice({"Pubkey":"UsCqnK5TOHBosdYxkuC4Zmop","Secret":"TAfkRueOzvUeJy42jE7xEpjJ6qapyA3u","Url":"https://demo.laiye.com:8082"}, path,"",false,[[1,2]],10000,30000) Traceprint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageBill_图片/Mage_PDFOCRInvoice.png)  

---

## 屏幕多票据识别

**说明**: 使用 Laiye Intelligent Document Processing 识别指定屏幕范围的多种票据，识别结果返回数组  

**原型**: `arrayRet = Mage.ScreenOCRInvoice(target,rect,config,time,optionArgs)`  

**参数**:  
- **target** (True) [decorator] 默认:@ui"" - 通过鼠标选取或截取需要识别的目标屏幕范围。包含窗口、元素、范围等信息  
- **rect** (True) [dictionary] 默认:{ "x": 0, "y": 0, "width": 0, "height": 0 } - 需要查找的范围，程序会在控件这个范围内进行识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的识别  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************屏幕多票据识别********************** 命令原型: arrayRet = Mage.ScreenOCRInvoice(@ui"",{"x": 0, "y": 0, "width": 0, "height": 0},{},30000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200}) 入参: target--目标元素 rect--默认识别范围 config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参: arrayRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim arrayRet="" // 输出结果 arrayRet = Mage.ScreenOCRInvoice(@ui"图像<img>3",{"x": 0, "y": 0, "width": 0, "height": 0},{"Pubkey":"UsCqnK5TOHBosdYxkuC6Zmop","Secret":"TAfkRueOzvUeJy42jE7xEpjJ6qapyA4u","Url":"https://demo.laiye.com:8082"},30000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200}) Traceprint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageBill_图片/Mage_ScreenOCRInvoice.png)  

---

## 获取卡证内容

**说明**: 获取通用卡证识别结果中的卡证内容  

**原型**: `sRet = Mage.ExtractCardInfo(jsonRet,invoice_type,invoice_key)`  

**参数**:  
- **jsonRet** (True) [expression] 默认:jsonRet - 使用"屏幕卡证识别"、"图像卡证识别"命令输出到的变量。如是"PDF卡证识别"命令输出到的变量，则需使用遍历数组的值  
- **invoice_type** (True) [enum] 默认:"" - 选择需要获取的卡证类型  
- **invoice_key** (True) [enum] 默认:"" - 选择获取卡证类型下的字段  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取卡证内容********************** 命令原型: sRet = Mage.ExtractCardInfo(jsonRet,"","") 入参: jsonRet--卡证识别结果 invoice_type--提取类型 invoice_key--提取字段 出参: sRet--函数调用的输出保存到的变量 ****************************************************/ Dim jsonRet=&#x27;&#x27; // 卡证识别结果,如:银行卡等 Dim sRet="" // 输出结果 sRet = Mage.ExtractCardInfo(jsonRet,"bank_card","card_number") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageCard_图片/Mage_ExtractCardInfo.png)  

---

## 获取卡证类型

**说明**: 获取通用卡证识别结果中的卡证类型  

**原型**: `sRet = Mage.ExtractCardType(jsonRet)`  

**参数**:  
- **jsonRet** (True) [expression] 默认:jsonRet - 使用"屏幕卡证识别"、"图像卡证识别"命令输出到的变量。如是"PDF卡证识别"命令输出到的变量，则需使用遍历数组的值  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获证卡证类型********************** 命令原型: sRet = Mage.ExtractCardType(jsonRet) 入参: jsonRet--卡证识别结果 出参: sRet--函数调用的输出保存到的变量 ****************************************************/ Rem 测试数据 Dim jsonRet={"ai_function" : "ocr_card","msg_id" : "13c26ea253363b8529ad72d90aa114e0","page_number" : 1,"result" : {"image_angle" : 0,"img_id" : "c7pa5npr8eh8o888obp0","items" : [],"rotated_image_height" : 842,"rotated_image_width" : 594,"type" : 20,"type_description" : "其它","type_key" : "other"}} // 测试数据 Dim sRet="" // 输出结果 sRet = Mage.ExtractCardType(jsonRet) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageCard_图片/Mage_ExtractCardType.png)  

---

## 图像卡证识别

**说明**: 使用 Laiye Intelligent Document Processing 识别指定图像的卡证，识别结果返回 JSON 格式  

**原型**: `jsonRet = Mage.ImageOCRCard(path,config,time)`  

**参数**:  
- **path** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要识别图像的路径，支持 jpeg、jpg、png、bmp、tif、tiff 等格式，PDF格式仅支持首页识别  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: jsonRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************图像卡证识别********************** 命令原型: jsonRet = Mage.ImageOCRCard(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,{},30000) 入参: path--待识别图像卡证路径 config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int 出参: jsonRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim path=&#x27;&#x27;&#x27;&#x27;&#x27;&#x27; // 待识别图像卡证路径 Dim jsonRet="" // 输出结果 jsonRet = Mage.ImageOCRCard(path,{"Pubkey":"3KhCyuXQutibxeEjOMDPyTxg","Secret":"jBxamqfDDzuY3zvYwOwG5TwVBVLdAbFW","Url":"https://demo.laiye.com:8082"},30000) TracePrint(jsonRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageCard_图片/Mage_ImageOCRCard.png)  

---

## PDF卡证识别

**说明**: 将 PDF 指定的页码通过 Laiye Intelligent Document Processing 通用卡证识别，返回结果数组。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额  

**原型**: `arrayRet = Mage.PDFOCRCard(config, path,password,all_pg_state,page_cfg,sleepTime,time)`  

**参数**:  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **path** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - PDF文件路径  
- **password** (True) [string] 默认:"" - PDF文件密码，无密码不需要填写  
- **all_pg_state** (True) [boolean] 默认:False - 当全部页码设为"是"，则识别全部且指定页码输入无效。设为否时，可指定页码识别  
- **page_cfg** (True) [expression] 默认:[ [1,2] ] - 支持正整数和数组格式，如输入2，则识别第2页；如输入 [1,3,5] ，则识别第1,3,5页；如输入[1, [6,9] ,4]，则识别1,4页和第6到第9页。当识别全部页码设为"是"，则识别指定页码的输入失效。超出PDF页码总数的部分会报错，页码重叠部分仅识别1次  
- **sleepTime** (True) [number] 默认:10000 - 识别PDF每页的间隔时长（以毫秒为单位），默认10000毫秒(10秒)。识别页数较多，间隔较短可能会导致调用频率超限错误  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************PDF卡证识别********************** 命令原型: arrayRet = Mage.PDFOCRCard({}, &#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"",false,[[1,2]],10000,30000) 入参: config--mage配置,需配置Pubkey和Secret.Type:Dict path--待识别包含图像卡证的PDF文件路径 password--密码 all_pg_state--是否识别全部页 page_cfg--识别指定页码 sleepTime--间隔时间.默认单位:毫秒.Type:Int time--超时时间.默认单位:毫秒.Type:Int 出参: arrayRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim path=&#x27;&#x27;&#x27;&#x27;&#x27;&#x27; // 待识别包含图像卡证的PDF文件路径 Dim arrayRet="" // 输出结果 arrayRet = Mage.PDFOCRCard({"Pubkey":"2KhCyuXQutibxeEjOMDPyTxg","Secret":"jBxamqfDDzuY5zvYwOwG5TwVBVLdAbFW","Url":"https://demo.laiye.com:8082"}, path,"",false,[[1,2]],10000,30000) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageCard_图片/Mage_PDFOCRCard.png)  

---

## 屏幕卡证识别

**说明**: 使用 Laiye Intelligent Document Processing 识别指定屏幕范围的卡证，识别结果以 JSON 格式返回  

**原型**: `jsonRet = Mage.ScreenOCRCard(target,rect,config,time,optionArgs)`  

**参数**:  
- **target** (True) [decorator] 默认:@ui"" - 通过鼠标选取或截取需要识别的目标屏幕范围。包含窗口、元素、范围等信息  
- **rect** (True) [dictionary] 默认:{ "x": 0, "y": 0, "width": 0, "height": 0 } - 需要查找的范围，程序会在控件这个范围内进行识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的识别  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************屏幕卡证识别********************** 命令原型: jsonRet = Mage.ScreenOCRCard(@ui"",{"x": 0, "y": 0, "width": 0, "height": 0},{},30000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200}) 入参: target--目标元素 rect--默认识别范围 config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参: jsonRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim jsonRet="" // 输出结果 jsonRet = Mage.ScreenOCRCard(@ui"图像<img>3",{"x": 0, "y": 0, "width": 0, "height": 0},{"Pubkey":"3KhCyuXQutibxeEjOMDPyTxg","Secret":"jBxamqfDDzuY3zvYwOwG5TwVBVLdAbFW","Url":"https://demo.laiye.com:8082"},30000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200}) TracePrint(jsonRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageCard_图片/Mage_ScreenOCRCard.png)  

---

## 获取剩余配额

**说明**: 获取 Laiye Intelligent Document Processing 指定能力的剩余配额数。可用于提前预判额度  

**原型**: `iRet = Mage.QuerySurplusQuota(config,time)`  

**参数**:  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取剩余配额********************** 命令原型: iRet = Mage.QuerySurplusQuota({},30000) 入参: config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int 出参: iRet--函数调用的输出保存到的变量 注意事项： 需要获取期望查询剩余配额的mage对应的Key/Secret和URL ****************************************************/ Dim iRet="" // 输出结果 iRet = Mage.QuerySurplusQuota({"Pubkey":"sCXf4tfmGpq8um0rY9MOvApD","Secret":"KLAhZgzHVqb975HAywi5sAhbxkakSHGx","Url":"https://demo.laiye.com:8082"},30000) Traceprint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageQuota_图片/Mage_QuerySurplusQuota.png)  

---

## 提取印章信息

**说明**: 从印章识别结果中提取指定的印章信息，提取结果为数组格式  

**原型**: `arrayRet = Mage.ExtractStampInfo(ocrResult,field)`  

**参数**:  
- **ocrResult** (True) [expression] 默认:jsonRet - 使用“图像印章识别”、“屏幕印章识别”、“PDF印章识别”命令输出的识别结果  
- **field** (True) [enum] 默认:"text" - 提取印章信息中的字段，分别有文字、颜色、形状、位置  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************提取印章信息********************** 命令原型: arrayRet = Mage.ExtractStampInfo(jsonRet,"text") 入参: ocrResult--印章识别结果 field--提取字段 出参: arrayRet--函数调用的输出保存到的变量 ****************************************************/ Rem 测试数据 Dim jsonRet={"ai_function" : "ocr_stamp","img_id" : "","msg_id" : "58004d488076a4ab52d3dbc5f3451736","stamps" : [{"color" : "OTHERS","color_description" : "其它","confidence" : 1,"positions" : [{"x" : 3,"y" : 3},{"x" : 609,"y" : 3},{"x" : 3,"y" : 609},{"x" : 609,"y" : 609}],"shape" : "OTHERS","shape_description" : "其它","text" : "某某某科技股份有限公司"}]} // 测试数据 Dim arrayRet="" // 输出结果 arrayRet = Mage.ExtractStampInfo(jsonRet,"text") TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageStamp_图片/Mage_ExtractStampInfo.png)  

---

## 图像印章识别

**说明**: 使用 Laiye Intelligent Document Processing 识别指定图像中的印章信息，识别结果为JSON格式  

**原型**: `jsonRet = Mage.ImageOCRStamp(filepath,config,time)`  

**参数**:  
- **filepath** (True) [path] 默认:"" - 待识别图像的存放路径。支持jpeg、jpg、png、pdf、bmp、tiff格式，图像文件大小不能超过10M  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: jsonRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************图像印章识别********************** 命令原型: jsonRet = Mage.ImageOCRStamp("",{},30000) 入参: filepath--待识别印章图片路径 config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int 出参: jsonRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim filepath=&#x27;&#x27;&#x27;&#x27;&#x27;&#x27; // 待识别印章图片路径 Dim jsonRet="" // 输出结果 jsonRet = Mage.ImageOCRStamp(filepath,{"Pubkey":"XDDpJLuf57aLAYb49WAu3ise","Secret":"MWikeR0v3TbwdYwTCcPc44aGywaybKmJ","Url":"https://mage.uibot.com.cn"},30000) TracePrint(jsonRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageStamp_图片/Mage_ImageOCRStamp.png)  

---

## PDF印章识别

**说明**: 使用 Laiye Intelligent Document Processing 识别 PDF 文件中指定页码区域内的印章信息，识别结果为JSON格式。在识别多页过程中如果其中一页失败，则会引发异常，且会消耗配额  

**原型**: `jsonRet = Mage.PDFOCRStamp(filepath,config,all_pg_state,page_cfg,sleepTime,time,optionArgs)`  

**参数**:  
- **filepath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 待识别PDF文件的存放路径  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **all_pg_state** (True) [boolean] 默认:True - 对PDF文件中指定的页码区域进行识别，默认识别全部页  
- **page_cfg** (True) [expression] 默认:[ [1,3] ] - 支持正整数和数组格式，如输入2，则识别第2页；如输入 [1,3,5] ，则识别第1,3,5页；如输入[1, [6,9] ,4]，则识别1,4页和第6到第9页。当识别全部页码设为"是"，则识别指定页码的输入失效。超出PDF页码总数的部分会报错，页码重叠部分仅识别1次  
- **sleepTime** (True) [number] 默认:10000 - 对PDF文件中每页的间隔时长（以毫秒为单位），默认10000毫秒(10秒)。识别页数较多，间隔较短可能会引发调用频率超限异常  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  
- **password** (False) [string] 默认:"" - 仅需要提供PDF文件密码时才填写  

**返回**: jsonRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************PDF印章识别********************** 命令原型: jsonRet = Mage.PDFOCRStamp(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,{},true,[[1,3]],10000,30000,{"password":""}) 入参: filepath--待识别包含印章图片的PDF文件路径 config--mage配置,需配置Pubkey和Secret.Type:Dict all_pg_state--是否识别全部页 page_cfg--指定页码区域 sleepTime--间隔时间.默认单位:毫秒.Type:Int time--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:密码).Type:Dict 出参: jsonRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim filepath=&#x27;&#x27;&#x27;&#x27;&#x27;&#x27; // 待识别包含印章图片的PDF文件路径 Dim jsonRet="" // 输出结果 jsonRet = Mage.PDFOCRStamp(filepath,{"Pubkey":"XDDpJLuf57aLAYb69WAu2ise","Secret":"MWikeR0v3TbwdYwTCcPc46aGywaybKmJ","Url":"https://mage.uibot.com.cn"},true,[[1,3]],10000,30000,{"password":""}) TracePrint(jsonRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageStamp_图片/Mage_PDFOCRStamp.png)  

---

## 屏幕印章识别

**说明**: 使用 Laiye Intelligent Document Processing 识别指定屏幕范围内的印章信息，识别结果为JSON格式  

**原型**: `jsonRet = Mage.ScreenOCRStamp(target,rect,config,time)`  

**参数**:  
- **target** (True) [decorator] 默认:@ui"" - 通过鼠标选取或截取需要识别的目标屏幕范围。包含窗口、元素、范围等信息  
- **rect** (True) [dictionary] 默认:{ "height":0,"width":0,"x":0,"y":0 } - 需要查找的范围，程序会在控件这个范围内进行识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的识别  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: jsonRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************屏幕印章识别********************** 命令原型: jsonRet = Mage.ScreenOCRStamp(@ui"",{"height":0,"width":0,"x":0,"y":0},{},30000) 入参: target--目标元素,该示例中使用的是百度中搜索的印章图片中的元素 rect--识别范围 config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int 出参: jsonRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim jsonRet="" // 输出结果 jsonRet = Mage.ScreenOCRStamp(@ui"图像<img>1",{"height":0,"width":0,"x":0,"y":0},{"Pubkey":"XDDpJLuf57aLAYb69WAu2ise","Secret":"MWikeR0v3TbwdYwTCcPc45aGywaybKmJ","Url":"https://mage.uibot.com.cn"},30000) TracePrint(jsonRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageStamp_图片/Mage_ScreenOCRStamp.png)  

---

## 获取所有表格

**说明**: 获取表格识别结果中的所有表格信息（不包含非表格文字），返回表格对象的数组  

**原型**: `arrayRet = Mage.ExtractAllTables(jsonRet)`  

**参数**:  
- **jsonRet** (True) [expression] 默认:jsonRet - 使用"屏幕表格识别"、"图像表格识别"、"PDF表格识别"命令输出到的变量  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取所有表格********************** 命令原型: arrayRet = Mage.ExtractAllTables(jsonRet) 入参: jsonRet--表格识别结果 出参: arrayRet :函数调用的输出保存到的变量 ****************************************************/ Dim jsonRet=&#x27;&#x27; // 表格识别结果,需使用自行识别指定图片后的结果进行赋值 Dim arrayRet="" // 输出结果 arrayRet = Mage.ExtractAllTables(jsonRet) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTable_图片/Mage_ExtractAllTables.png)  

---

## 获取非表格文字

**说明**: 获取表格识别结果中的非表格文字信息  

**原型**: `arrayRet = Mage.ExtractOutsideTableText(jsonRet)`  

**参数**:  
- **jsonRet** (True) [expression] 默认:jsonRet - 使用"屏幕表格识别"、"图像表格识别"、"PDF表格识别"命令输出到的变量  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取非表格文字********************** 命令原型: arrayRet = Mage.ExtractOutsideTableText(jsonRet) 入参: jsonRet--表格识别结果 出参: jsonRet:函数调用的输出保存到的变量 ****************************************************/ Dim jsonRet=&#x27;&#x27; // 表格识别结果,需使用自行识别指定图片后的结果进行赋值 Dim arrayRet="" // 输出结果 arrayRet = Mage.ExtractOutsideTableText(jsonRet) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTable_图片/Mage_ExtractOutsideTableText.png)  

---

## 获取指定表格

**说明**: 获取表格识别结果中的指定表格信息，返回表格对象，该对象为二维数组  

**原型**: `objTableData = Mage.ExtractSingleTable(jsonRet, talbe_id)`  

**参数**:  
- **jsonRet** (True) [expression] 默认:jsonRet - 使用"屏幕表格识别"、"图像表格识别"、"PDF表格识别"命令输出到的变量  
- **talbe_id** (True) [number] 默认:0 - 指定表格识别结果中的表格索引（从0开始）  

**返回**: objTableData，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取指定表格********************** 命令原型: objTableData = Mage.ExtractSingleTable(jsonRet, 0) 入参: jsonRet--表格识别结果 talbe_id--表格索引 出参: objTableData:函数调用的输出保存到的变量 ****************************************************/ Dim jsonRet=&#x27;&#x27; // 表格识别结果,需使用自行识别指定图片后的结果进行赋值 Dim objTableData="" // 输出结果 objTableData = Mage.ExtractSingleTable(jsonRet, 0) TracePrint(objTableData)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTable_图片/Mage_ExtractSingleTable.png)  

---

## 获取表格单元格

**说明**: 从表格对象中获取指定表格单元格信息，返回字符串  

**原型**: `sRet = Mage.ExtractSingleTableCell(objTableData,row,col)`  

**参数**:  
- **objTableData** (True) [expression] 默认:objTableData - 使用"获取指定表格"命令输出到的变量或"获取全部表格"命令指定表格索引的变量（如：arrayRet [1] ，表示表格索引为2的表格）  
- **row** (True) [number] 默认:1 - 指定表格对象的单元格行号（从1开始）  
- **col** (True) [number] 默认:1 - 指定表格对象的单元格列号（从1开始）  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取表格单元格********************** 命令原型: sRet = Mage.ExtractSingleTableCell(objTableData,1,1) 入参: objTableData--表格对象 row--行号 col--列号 出参: sRet:函数调用的输出保存到的变量 ****************************************************/ Dim objTableData=&#x27;&#x27; // 表格对象,使用"获取指定表格"命令输出到的变量或"获取全部表格"命令指定表格索引的变量（如：arrayRet[1]，表示表格索引为2的表格） Dim sRet="" // 输出结果 sRet = Mage.ExtractSingleTableCell(objTableData,1,1) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTable_图片/Mage_ExtractSingleTableCell.png)  

---

## 获取表格列

**说明**: 从表格对象中获取指定表格整列信息，返回一维数组  

**原型**: `arrayRet = Mage.ExtractSingleTableCol(objTableData,col)`  

**参数**:  
- **objTableData** (True) [expression] 默认:objTableData - 使用"获取指定表格"命令输出到的变量或"获取全部表格"命令指定表格索引的变量（如：arrayRet [1] ，表示表格索引为2的表格）  
- **col** (True) [number] 默认:1 - 指定表格对象的列号（从1开始）  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取表格列********************** 命令原型: arrayRet = Mage.ExtractSingleTableCol(objTableData,1) 入参: objTableData--表格对象 col--列号 出参: arrayRet:函数调用的输出保存到的变量 ****************************************************/ Dim objTableData=&#x27;&#x27; // 表格对象,使用"获取指定表格"命令输出到的变量或"获取全部表格"命令指定表格索引的变量（如：arrayRet[1]，表示表格索引为2的表格） Dim arrayRet="" // 输出结果 arrayRet = Mage.ExtractSingleTableCol(objTableData,1) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTable_图片/Mage_ExtractSingleTableCol.png)  

---

## 获取表格列数

**说明**: 从表格对象中获取表格的列数，返回数字  

**原型**: `iRet = Mage.ExtractSingleTableColNum(objTableData)`  

**参数**:  
- **objTableData** (True) [expression] 默认:objTableData - 使用"获取指定表格"命令输出到的变量或"获取全部表格"命令指定表格索引的变量（如：arrayRet [1] ，表示表格索引为2的表格）  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取表格列数********************** 命令原型: iRet = Mage.ExtractSingleTableColNum(objTableData) 入参: objTableData--表格对象 出参: iRet:函数调用的输出保存到的变量 ****************************************************/ Dim objTableData=&#x27;&#x27; // 表格对象,使用"获取指定表格"命令输出到的变量或"获取全部表格"命令指定表格索引的变量（如：arrayRet[1]，表示表格索引为2的表格） Dim iRet="" // 输出结果 iRet = Mage.ExtractSingleTableColNum(objTableData) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTable_图片/Mage_ExtractSingleTableColNum.png)  

---

## 获取表格行

**说明**: 从表格对象中获取指定表格整行信息，返回一维数组  

**原型**: `arrayRet = Mage.ExtractSingleTableRow(objTableData,row)`  

**参数**:  
- **objTableData** (True) [expression] 默认:objTableData - 使用"获取指定表格"命令输出到的变量或"获取全部表格"命令指定表格索引的变量（如：arrayRet [1] ，表示表格索引为2的表格）  
- **row** (True) [number] 默认:1 - 指定表格对象的行号（从1开始）  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取表格行********************** 命令原型: arrayRet = Mage.ExtractSingleTableRow(objTableData,1) 入参: objTableData--表格对象 row--行号 出参: arrayRet:函数调用的输出保存到的变量 ****************************************************/ Dim objTableData=&#x27;&#x27; // 表格对象,使用"获取指定表格"命令输出到的变量或"获取全部表格"命令指定表格索引的变量（如：arrayRet[1]，表示表格索引为2的表格） Dim arrayRet="" // 输出结果 arrayRet = Mage.ExtractSingleTableRow(objTableData,1) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTable_图片/Mage_ExtractSingleTableRow.png)  

---

## 获取表格行数

**说明**: 从表格对象中获取表格的行数，返回数字  

**原型**: `iRet = Mage.ExtractSingleTableRowNum(objTableData)`  

**参数**:  
- **objTableData** (True) [expression] 默认:objTableData - 使用"获取指定表格"命令输出到的变量或"获取全部表格"命令指定表格索引的变量（如：arrayRet [1] ，表示表格索引为2的表格）  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取表格行数********************** 命令原型: iRet = Mage.ExtractSingleTableRowNum(objTableData) 入参: objTableData--表格对象 出参: iRet:函数调用的输出保存到的变量 ****************************************************/ Dim objTableData=&#x27;&#x27; // 表格对象,使用"获取指定表格"命令输出到的变量或"获取全部表格"命令指定表格索引的变量（如：arrayRet[1]，表示表格索引为2的表格） Dim iRet="" // 输出结果 iRet = Mage.ExtractSingleTableRowNum(objTableData) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTable_图片/Mage_ExtractSingleTableRowNum.png)  

---

## 获取表格区域

**说明**: 从表格对象中获取区域信息，返回二维数组  

**原型**: `arrayRet = Mage.ExtractTableRegion(objTableData, start_row,start_col,end_row,end_col)`  

**参数**:  
- **objTableData** (True) [expression] 默认:objTableData - 使用"获取指定表格"命令输出到的变量或"获取全部表格"命令指定表格索引的变量（如：arrayRet [1] ，表示表格索引为2的表格）  
- **start_row** (True) [number] 默认:1 - 指定表格对象的开始单元格行号（从1开始）  
- **start_col** (True) [number] 默认:1 - 指定表格对象的开始单元格列号（从1开始）  
- **end_row** (True) [number] 默认:2 - 指定表格对象的结束单元格行号（从1开始）  
- **end_col** (True) [number] 默认:2 - 指定表格对象的结束单元格列号（从1开始）  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取表格区域********************** 命令原型: arrayRet = Mage.ExtractTableRegion(objTableData, 1,1,2,2) 入参: objTableData--表格对象 start_row--开始行号 start_col--开始列号 end_row--结束行号 end_col--结束列号 出参: arrayRet:函数调用的输出保存到的变量 ****************************************************/ Dim objTableData=&#x27;&#x27; // 表格对象,使用"获取指定表格"命令输出到的变量或"获取全部表格"命令指定表格索引的变量（如：arrayRet[1]，表示表格索引为2的表格） Dim arrayRet="" // 输出结果 arrayRet = Mage.ExtractTableRegion(objTableData, 1,1,2,2) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTable_图片/Mage_ExtractTableRegion.png)  

---

## 获取表格数

**说明**: 获取表格识别结果中的所有表格数量（不包含非表格文字），返回数字  

**原型**: `iRet = Mage.ExtractTablesNum(jsonRet)`  

**参数**:  
- **jsonRet** (True) [expression] 默认:jsonRet - 使用"屏幕表格识别"、"图像表格识别"、"PDF表格识别"命令输出到的变量  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取表格数********************** 命令原型: iRet = Mage.ExtractTablesNum(jsonRet) 入参: jsonRet--表格识别结果 出参: iRet:函数调用的输出保存到的变量 ****************************************************/ Dim jsonRet=&#x27;&#x27; // 表格识别结果,需使用自行识别指定图片后的结果进行赋值 Dim iRet="" // 输出结果 iRet = Mage.ExtractTablesNum(jsonRet) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTable_图片/Mage_ExtractTablesNum.png)  

---

## 提取表格结果至Excel

**说明**: 将 "屏幕表格识别"、"图像表格识别"、"PDF表格识别"命令的识别结果直接提取至Excel文件中  

**原型**: `Mage.ExtractTablesToExcel(jsonRet,filter_text,sPath, appType)`  

**参数**:  
- **jsonRet** (True) [expression] 默认:jsonRet - 使用"屏幕表格识别"、"图像表格识别"、"PDF表格识别"命令输出到的变量  
- **filter_text** (True) [boolean] 默认:False - 过滤识别结果中的非表格文本。选择"否"则将完整识别结果写入Excel中的Sheet1页，选择"是"则将识别的每个表格按顺序分别写入Excel的单个Sheet页  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - Excel工作簿文件路径，如果指定路径不存在对应文件，该命令将在此路径创建该文件  
- **appType** (True) [enum] 默认:"Excel" - 使用Excel或者WPS打开  

**示例**:  
```
/**********************提取表格结果至Excel********************** 命令原型: Mage.ExtractTablesToExcel(jsonRet,false,&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;, "Excel") 入参: jsonRet--表格识别结果 filter_text--过滤非表格文本 sPath--文件路径 appType--打开方式 ****************************************************/ Dim jsonRet=&#x27;&#x27; // 表格识别结果,需使用自行识别指定图片后的结果进行赋值 Dim sPath=&#x27;&#x27;&#x27;&#x27;&#x27;&#x27; // 文件保存路径 Mage.ExtractTablesToExcel(jsonRet,false,sPath, "Excel")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTable_图片/Mage_ExtractTablesToExcel.png)  

---

## 图像表格识别

**说明**: 使用 Laiye Intelligent Document Processing 识别指定图像的多个表格，识别结果返回 JSON 格式  

**原型**: `jsonRet = Mage.ImageOCRTable(path,config, time)`  

**参数**:  
- **path** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要识别图像的路径，支持 jpeg、jpg、png、bmp、tif、tiff 等格式，PDF格式仅支持首页识别  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: jsonRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************图像表格识别********************** 命令原型: jsonRet = Mage.ImageOCRTable(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,{}, 30000) 入参: path--待识别图片的路径.Type:String config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int 出参: jsonRet:函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim path=&#x27;&#x27;&#x27;&#x27;&#x27;&#x27; // 待识别图片的路径 Dim jsonRet="" // 输出结果 jsonRet = Mage.ImageOCRTable(path,{"Pubkey":"SXX2ZbKqndP3QGhVyZM30eqh","Secret":"zidgGiVY2JzxoYMH2BB6o7YxBS97Xyv6","Url":"https://demo.laiye.com:8082"}, 30000) TracePrint(jsonRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTable_图片/Mage_ImageOCRTable.png)  

---

## PDF表格识别

**说明**: 将 PDF 指定的页码通过 Laiye Intelligent Document Processing 通用表格识别，识别结果返回 JSON 格式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额  

**原型**: `jsonRet = Mage.PDFOCRTable(config, path,password,all_pg_state,page_cfg,sleepTime,time)`  

**参数**:  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **path** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - PDF文件路径  
- **password** (True) [string] 默认:"" - PDF文件密码，无密码不需要填写  
- **all_pg_state** (True) [boolean] 默认:False - 当全部页码设为"是"，则识别全部且指定页码输入无效。设为否时，可指定页码识别  
- **page_cfg** (True) [expression] 默认:[ [1,2] ] - 支持正整数和数组格式，如输入2，则识别第2页；如输入 [1,3,5] ，则识别第1,3,5页；如输入[1, [6,9] ,4]，则识别1,4页和第6到第9页。当识别全部页码设为"是"，则识别指定页码的输入失效。超出PDF页码总数的部分会报错，页码重叠部分仅识别1次  
- **sleepTime** (True) [number] 默认:10000 - 识别PDF每页的间隔时长（以毫秒为单位），默认10000毫秒(10秒)。识别页数较多，间隔较短可能会导致调用频率超限错误  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: jsonRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************PDF表格识别********************** 命令原型: jsonRet = Mage.PDFOCRTable({}, &#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"",false,[[1,2]],10000,30000) 入参: config--mage配置,需配置Pubkey和Secret.Type:Dict path--待识别图片的PDF文件路径.Type:String password--密码.无密码则不需要填写.Type:String all_pg_state--是否识别全部页.Type:Bool page_cfg--识别指定页码.Type:List sleepTime--间隔时间.默认单位:毫秒.Type:Int time--超时时间.默认单位:毫秒.Type:Int 出参: jsonRet:函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim path=&#x27;&#x27;&#x27;&#x27;&#x27;&#x27; // 待识别PDF的路径 Dim jsonRet="" // 输出结果 jsonRet = Mage.PDFOCRTable({"Pubkey":"SXX2ZbKqndP5QGhVyZM30eqh","Secret":"zidgGiVY5JzxoYMH2BB6o7YxBS97Xyv6","Url":"https://demo.laiye.com:8082"}, path,"",false,[[1,2]],10000,30000) TracePrint(jsonRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTable_图片/Mage_PDFOCRTable.png)  

---

## 屏幕表格识别

**说明**: 使用 Laiye Intelligent Document Processing 识别指定屏幕范围的多个表格，识别结果返回 JSON 格式  

**原型**: `jsonRet = Mage.ScreenOCRTable(target,rect,config,time,optionArgs)`  

**参数**:  
- **target** (True) [decorator] 默认:@ui"" - 通过鼠标选取或截取需要识别的目标屏幕范围。包含窗口、元素、范围等信息  
- **rect** (True) [dictionary] 默认:{ "x":0,"y":0,"width":0,"height":0 } - 需要查找的范围，程序会在控件这个范围内进行识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的识别  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: jsonRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************屏幕表格识别********************** 命令原型: jsonRet = Mage.ScreenOCRTable(@ui"",{"x":0,"y":0,"width":0,"height":0},{},30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) 入参: target--目标元素 rect--识别范围.Type:Dict config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参: jsonRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim jsonRet="" // 输出结果 jsonRet = Mage.ScreenOCRTable(@ui"窗口1",{"x":0,"y":0,"width":0,"height":0},{"Pubkey":"nOPmQGDCJR3XLCnkwykAEF9N","Secret":"dNu4FBCuy9J5rK6b0VxRhNvEBEMdAMCt","Url":"https://mage.uibot.com.cn"},30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) TracePrint(jsonRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTable_图片/Mage_ScreenOCRTable.png)  

---

## 获取模板识别结果

**说明**: 获取自定义模板识别结果中指定字段的结果  

**原型**: `arrayRet = Mage.ExtractOCRTemplateInfo(jsonRet,extractor,template_name,field_name,update_time)`  

**参数**:  
- **jsonRet** (True) [reference] 默认:jsonRet - 使用"屏幕自定义模板识别"、"图像自定义模板识别"命令输出到的变量。如是"PDF自定义模板识别"命令输出到的变量，则需使用遍历数组的值  
- **extractor** (True) [expression] 默认:{ } - 选择自定义模板的识别器  
- **template_name** (True) [string] 默认:"" - 选择自定义模板名称  
- **field_name** (True) [string] 默认:"" - 选择模板中的字段  
- **update_time** (True) [string] 默认:"" - 不可修改，选中命令时自动获取模板更新时间，如果与自定义模板识别的结果使用的版本不一致，则在运行时提示  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取模板的字段列表********************** 命令原型: arrayRet = Mage.ExtractOCRTemplateInfo(jsonRet,{},"","","") 入参: jsonRet--模板识别结果 extractor--识别器,选择使用的mage识别器 template_name--模板名称 field_name--字段名称 update_time--更新时间 出参: arrayRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL并配置mage模型后使用 ****************************************************/ Rem 测试数据 Dim jsonRet = {"ai_function" : "ocr_template","msgId" : "7bc9a3eb20922fd065b60eb4f934e573","page_number" : 1,"raw" : {"image_angle" : 0,"items" : [],"rotated_image_height" : 0,"rotated_image_width" : 0,"struct_content" : null,"tables" : []},"results" : [{"field_name" : "学号","results" : ["2021005"]},{"field_name" : "姓名","results" : ["孙七"]},{"field_name" : "性别","results" : ["男"]},{"field_name" : "考试日期","results" : ["2020.03.10"]},{"field_name" : "年级","results" : ["高一年级"]},{"field_name" : "语文","results" : ["85"]},{"field_name" : "数学","results" : ["79"]},{"field_name" : "英语","results" : ["85"]},{"field_name" : "历史","results" : ["75"]},{"field_name" : "化学","results" : ["79"]}],"template_hash" : "AAAAAAAAAAAAAAAAAAAAAC2Thko=00","template_name" : "成绩分析","update_time" : "2021-08-31 12:20:53"} // 测试数据 Dim arrayRet="" // 输出结果 arrayRet = Mage.ExtractOCRTemplateInfo(jsonRet,{"Pubkey":"wHCsSNCfWU2HVijhuU8WVf6s","Secret":"nmWAAYD2ax7Qb3TwpDaVu9DaRXGPmD3h","Url":"https://mage.uibot.com.cn"},"登机牌","航班号","2021-09-15 11:08:22") Traceprint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTemplate_图片/Mage_ExtractOCRTemplateInfo.png)  

---

## 获取自定义模板名称

**说明**: 获取自定义模板识别结果中的模板名称  

**原型**: `sRet = Mage.ExtractOCRTemplateName(jsonRet)`  

**参数**:  
- **jsonRet** (True) [expression] 默认:jsonRet - 使用"屏幕自定义模板识别"、"图像自定义模板识别"命令输出到的变量。如是"PDF自定义模板识别"命令输出到的变量，则需使用遍历数组的值  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取自定义模板名称********************** 命令原型: sRet = Mage.ExtractOCRTemplateName(jsonRet) 入参: jsonRet--模板识别结果 出参: sRet--函数调用的输出保存到的变量 ****************************************************/ Rem 测试数据 Dim jsonRet = {"ai_function" : "ocr_template","msgId" : "7bc9a3eb20922fd065b60eb4f934e573","page_number" : 1,"raw" : {"image_angle" : 0,"items" : [],"rotated_image_height" : 0,"rotated_image_width" : 0,"struct_content" : null,"tables" : []},"results" : [{"field_name" : "学号","results" : ["2021005"]},{"field_name" : "姓名","results" : ["孙七"]},{"field_name" : "性别","results" : ["男"]},{"field_name" : "考试日期","results" : ["2020.03.10"]},{"field_name" : "年级","results" : ["高一年级"]},{"field_name" : "语文","results" : ["85"]},{"field_name" : "数学","results" : ["79"]},{"field_name" : "英语","results" : ["85"]},{"field_name" : "历史","results" : ["75"]},{"field_name" : "化学","results" : ["79"]}],"template_hash" : "AAAAAAAAAAAAAAAAAAAAAC2Thko=00","template_name" : "成绩分析","update_time" : "2021-08-31 12:20:53"} // 测试数据 Dim sRet="" // 输出结果 sRet = Mage.ExtractOCRTemplateName(jsonRet) Traceprint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTemplate_图片/Mage_ExtractOCRTemplateName.png)  

---

## 获取模板的字段列表

**说明**: 从 Laiye Intelligent Document Processing 接口获取识别器中自定义模板的字段列表  

**原型**: `arrayRet = Mage.GetOCRTemplateFieldList(extractor,template_name,time)`  

**参数**:  
- **extractor** (True) [expression] 默认:{ } - 选择自定义模板的识别器  
- **template_name** (True) [string] 默认:"" - 选择自定义模板名称  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取模板的字段列表********************** 命令原型: arrayRet = Mage.GetOCRTemplateFieldList({},"",30000) 入参: extractor--识别器,选择使用的mage识别器 template_name--模板名称 time--超时时间.默认单位:毫秒.Type:Int 出参: arrayRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL并配置mage模型后使用 ****************************************************/ Dim arrayRet="" // 输出结果 arrayRet = Mage.GetOCRTemplateFieldList({"Pubkey":"sno8z32O9zrCoZxmd8x9rWNo","Secret":"HVOmsjPDoUgjytuFHEPSgsGbZORwdnyO","Url":"https://mage.uibot.com.cn"},"费用账单",30000) Traceprint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTemplate_图片/Mage_GetOCRTemplateFieldList.png)  

---

## 图像自定义模板识别

**说明**: 使用 Laiye Intelligent Document Processing 识别指定图像自定义模板内容，识别结果以 JSON 格式返回  

**原型**: `jsonRet = Mage.ImageOCRTemplate(path,config,time)`  

**参数**:  
- **path** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要识别图像的路径，支持 jpeg、jpg、png、bmp、tif、tiff 等格式，PDF格式仅支持首页识别  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: jsonRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************图像自定义模板识别********************** 命令原型: jsonRet = Mage.ImageOCRTemplate(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,{},30000) 入参: path--待识别图片路径 config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int 出参: jsonRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim path=&#x27;&#x27;&#x27;&#x27;&#x27;&#x27; // 待识别图片路径 Dim jsonRet="" // 输出结果 jsonRet = Mage.ImageOCRTemplate(path,{"Pubkey":"iUw55tA8jVn3UJF0oqqVBdHW","Secret":"QSFuxR4aaEsnm2LwX5NSK4frSAwhrvKX","Url":"https://demo.laiye.com:8082"},30000) Traceprint(jsonRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTemplate_图片/Mage_ImageOCRTemplate.png)  

---

## PDF自定义模板识别

**说明**: 将 PDF 指定的页码通过 Laiye Intelligent Document Processing 自定义模板识别，返回结果数组。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额  

**原型**: `arrayRet = Mage.PDFOCRTemplate(config, path,password,all_pg_state,page_cfg,sleepTime,time)`  

**参数**:  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **path** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - PDF文件路径  
- **password** (True) [string] 默认:"" - PDF文件密码，无密码不需要填写  
- **all_pg_state** (True) [boolean] 默认:False - 当全部页码设为"是"，则识别全部且指定页码输入无效。设为否时，可指定页码识别  
- **page_cfg** (True) [expression] 默认:[ [1,2] ] - 支持正整数和数组格式，如输入2，则识别第2页；如输入 [1,3,5] ，则识别第1,3,5页；如输入[1, [6,9] ,4]，则识别1,4页和第6到第9页。当识别全部页码设为"是"，则识别指定页码的输入失效。超出PDF页码总数的部分会报错，页码重叠部分仅识别1次  
- **sleepTime** (True) [number] 默认:10000 - 识别PDF每页的间隔时长（以毫秒为单位），默认10000毫秒(10秒)。识别页数较多，间隔较短可能会导致调用频率超限错误  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************PDF自定义模板识别********************** 命令原型: arrayRet = Mage.PDFOCRTemplate({}, &#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"",false,[[1,2]],10000,30000) 入参: config--mage配置,需配置Pubkey和Secret.Type:Dict path--待识别图片路径 password--密码 all_pg_state--识别全部页 page_cfg--识别指定页码 sleepTime--间隔时间.默认单位:毫秒.Type:Int time--超时时间.默认单位:毫秒.Type:Int 出参: arrayRet--函数调用的输出保存到的变量 ****************************************************/ Dim path=&#x27;&#x27;&#x27;&#x27;&#x27;&#x27; // 待识别图片路径 Dim arrayRet="" // 输出结果 arrayRet = Mage.PDFOCRTemplate({}, path,"",false,[[1,2]],10000,30000) Traceprint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTemplate_图片/Mage_PDFOCRTemplate.png)  

---

## 屏幕自定义模板识别

**说明**: 使用 Laiye Intelligent Document Processing 识别指定屏幕范围自定义模板内容，识别结果返回 JSON 格式  

**原型**: `jsonRet = Mage.ScreenOCRTemplate(target,rect,config,time,optionArgs)`  

**参数**:  
- **target** (True) [decorator] 默认:@ui"" - 通过鼠标选取或截取需要识别的目标屏幕范围。包含窗口、元素、范围等信息  
- **rect** (True) [dictionary] 默认:{ "x": 0, "y": 0, "width": 0, "height": 0 } - 需要查找的范围，程序会在控件这个范围内进行识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的识别  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: jsonRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************屏幕自定义模板识别********************** 命令原型: jsonRet = Mage.ScreenOCRTemplate(@ui"",{"x": 0, "y": 0, "width": 0, "height": 0},{},30000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200}) 入参: target--目标元素,该示例使用的是成绩单,需要根据实际情况在mage中先对模板进行标注等操作 rect--识别范围.Type:Dict config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参: jsonRet--函数调用的输出保存到的变量 ****************************************************/ Dim jsonRet="" // 输出结果 jsonRet = Mage.ScreenOCRTemplate(@ui"窗口",{"x": 0, "y": 0, "width": 0, "height": 0},{},30000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200}) Traceprint(jsonRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTemplate_图片/Mage_ScreenOCRTemplate.png)  

---

## 鼠标点击文本

**说明**: 使用 Laiye Intelligent Document Processing 对窗口范围内进行指定文字识别，如果识别到指定文字就点击它  

**原型**: `Mage.Click(objUiElement,objRect,config,sText,iRule,iOccurrence,iButton,iType,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 通过鼠标选取或截取需要识别的目标屏幕范围。包含窗口、元素、范围等信息  
- **objRect** (True) [dictionary] 默认:{ "x":0,"y":0,"width":0,"height":0 } - 需要进行OCR文字识别的范围，程序会在控件这个范围内进行文字识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的文字识别  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **sText** (True) [string] 默认:"" - 查找元素时使用的文本  
- **iRule** (True) [enum] 默认:"instr" - 查找文本时使用的规则  
- **iOccurrence** (True) [number] 默认:1 - 如果“文本”字段中的字符串在指示的界面元素中出现多次，请在此处指定要单击的出现次数。例如，如果字符串出现4次并且您要单击第一个匹配项，请在此字段中写入1  
- **iButton** (True) [enum] 默认:"left" - 鼠标按键 { left:左键, right:右键, middle:中键 }  
- **iType** (True) [enum] 默认:"click" - 点击类型 { click:单击, dbclick:双击, down:按下, up:弹起 }  
- **iTimeOut** (True) [number] 默认:30000 - 指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发错误。默认30000毫秒（30秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **sCursorPosition** (False) [enum] 默认:"Center" - 描述添加OffsetX和OffsetY属性的偏移量的光标起点。可以使用以下选项：TopLeft，TopRight，BottomLeft，BottomRight和Center。默认选项是Center  
- **iCursorOffsetX** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的水平位移  
- **iCursorOffsetY** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的垂直位移  
- **sKeyModifiers** (False) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **sSimulate** (False) [enum] 默认:"simulate" - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：模拟操作(simulate)  

**示例**:  
```
/**********************鼠标点击文本********************** 命令原型: Mage.Click(@ui"",{"x":0,"y":0,"width":0,"height":0},{},"","instr",1,"left","click",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate"}) 入参: objUiElement--目标,该示例中使用的是chrome打开百度首页的“百度一下”按钮元素 objRect--识别范围.默认值:{"x":0,"y":0,"width":0,"height":0} config--mage配置,需配置Pubkey和Secret.Type:Dict sText--查找文本 iRule--查找规则 iOccurrence--相似结果位置 iButton--鼠标点击 iType--点击类型 iTimeOut--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时/光标位置/横坐标偏移/纵坐标偏移/辅助按键/操作类型).Type:Dict 注意事项： 1. 需要获取mage对应的Key/Secret和URL 2. 元素需要重新录制 ****************************************************/ Mage.Click(@ui"输入控件<input>_百度一下",{"x":0,"y":0,"width":0,"height":0},{"Pubkey":"sCXf4tfmGpq8um0rY2MOvApD","Secret":"KLAhZgzHVqb945HAywi5sAhbxkakSHGx","Url":"https://demo.laiye.com:8082"},"百度一下","instr",1,"left","click",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate"})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageText_图片/Mage_Click.png)  

---

## 获取全部文本

**说明**: 获取通用文字识别结果中的全部文本  

**原型**: `sRet = Mage.ExtractAllText(jsonRet,include_enter)`  

**参数**:  
- **jsonRet** (True) [expression] 默认:jsonRet - 使用"屏幕文字识别"、"图像文字识别"、"PDF文字识别"命令输出到的变量  
- **include_enter** (True) [boolean] 默认:False - 全部文本中是否包含换行信息，为"是"则在每行后面添加\n。“否”则不添加  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取全部文本********************** 命令原型: sRet = Mage.ExtractAllText(jsonRet,false) 入参: jsonRet--使用"屏幕文字识别"、"图像文字识别"、"PDF文字识别"命令输出到的变量 include_enter--全部文本中是否包含换行信息，为"是"则在每行后面添加\n。“否”则不添加 出参: sRet:函数调用的输出保存到的变量 ****************************************************/ Rem 测试数据 Dim jsonRet = {"ai_function" : "ocr_text","items" : [{"char_positions" : [],"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格","handwrite_info" : null,"importance_info" : null,"page_number" : 1,"positions" : [{"x" : 71,"y" : 60},{"x" : 522,"y" : 61},{"x" : 522,"y" : 75},{"x" : 71,"y" : 74}],"probabilities" : []},{"char_positions" : [],"content" : "式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","handwrite_info" : null,"importance_info" : null,"page_number" : 1,"positions" : [{"x" : 72,"y" : 76},{"x" : 420,"y" : 76},{"x" : 420,"y" : 89},{"x" : 72,"y" : 89}],"probabilities" : []}],"struct_content" : {"page" : [{"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","page_id" : 0,"page_number" : 1}],"paragraph" : [{"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","page_number" : 1,"paragraph_id" : 0}],"row" : [{"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格","page_number" : 1,"row_id" : 0},{"content" : "式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","page_number" : 1,"row_id" : 1}]}} // 测试数据 Dim sRet="" // 输出结果 sRet = Mage.ExtractAllText(jsonRet,false) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageText_图片/Mage_ExtractAllText.png)  

---

## 获取每行文本

**说明**: 获取通用文字识别结果中按行划分的全部文本  

**原型**: `arrayRet = Mage.ExtractLineText(jsonRet)`  

**参数**:  
- **jsonRet** (True) [expression] 默认:jsonRet - 使用"屏幕文字识别"、"图像文字识别"、"PDF文字识别"命令输出到的变量  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取每行文本********************** 命令原型: arrayRet = Mage.ExtractLineText(jsonRet) 入参: jsonRet--使用"屏幕文字识别"、"图像文字识别"、"PDF文字识别"命令输出到的变量 出参: arrayRet:函数调用的输出保存到的变量 ****************************************************/ Rem 测试数据 Dim jsonRet = {"ai_function" : "ocr_text","items" : [{"char_positions" : [],"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格","handwrite_info" : null,"importance_info" : null,"page_number" : 1,"positions" : [{"x" : 71,"y" : 60},{"x" : 522,"y" : 61},{"x" : 522,"y" : 75},{"x" : 71,"y" : 74}],"probabilities" : []},{"char_positions" : [],"content" : "式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","handwrite_info" : null,"importance_info" : null,"page_number" : 1,"positions" : [{"x" : 72,"y" : 76},{"x" : 420,"y" : 76},{"x" : 420,"y" : 89},{"x" : 72,"y" : 89}],"probabilities" : []}],"struct_content" : {"page" : [{"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","page_id" : 0,"page_number" : 1}],"paragraph" : [{"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","page_number" : 1,"paragraph_id" : 0}],"row" : [{"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格","page_number" : 1,"row_id" : 0},{"content" : "式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","page_number" : 1,"row_id" : 1}]}} // 测试数据 Dim arrayRet="" // 输出结果 arrayRet = Mage.ExtractLineText(jsonRet) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageText_图片/Mage_ExtractLineText.png)  

---

## 获取段落文本

**说明**: 获取通用文字识别结果中按段落划分的全部文本  

**原型**: `arrayRet = Mage.ExtractParagraphText(jsonRet)`  

**参数**:  
- **jsonRet** (True) [expression] 默认:jsonRet - 使用"屏幕文字识别"、"图像文字识别"、"PDF文字识别"命令输出到的变量  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取段落文本********************** 命令原型: arrayRet = Mage.ExtractParagraphText(jsonRet) 入参: jsonRet--使用"屏幕文字识别"、"图像文字识别"、"PDF文字识别"命令输出到的变量 出参: arrayRet:函数调用的输出保存到的变量 ****************************************************/ Rem 测试数据 Dim jsonRet = {"ai_function" : "ocr_text","items" : [{"char_positions" : [],"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格","handwrite_info" : null,"importance_info" : null,"page_number" : 1,"positions" : [{"x" : 71,"y" : 60},{"x" : 522,"y" : 61},{"x" : 522,"y" : 75},{"x" : 71,"y" : 74}],"probabilities" : []},{"char_positions" : [],"content" : "式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","handwrite_info" : null,"importance_info" : null,"page_number" : 1,"positions" : [{"x" : 72,"y" : 76},{"x" : 420,"y" : 76},{"x" : 420,"y" : 89},{"x" : 72,"y" : 89}],"probabilities" : []}],"struct_content" : {"page" : [{"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","page_id" : 0,"page_number" : 1}],"paragraph" : [{"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","page_number" : 1,"paragraph_id" : 0}],"row" : [{"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格","page_number" : 1,"row_id" : 0},{"content" : "式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","page_number" : 1,"row_id" : 1}]}} // 测试数据 Dim arrayRet="" // 输出结果 arrayRet = Mage.ExtractParagraphText(jsonRet) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageText_图片/Mage_ExtractParagraphText.png)  

---

## 获取所有文本元素

**说明**: 获取通用文字识别结果中按文本元素划分的全部文本  

**原型**: `arrayRet = Mage.ExtractSentenceText(jsonRet)`  

**参数**:  
- **jsonRet** (True) [expression] 默认:jsonRet - 使用"屏幕文字识别"、"图像文字识别"、"PDF文字识别"命令输出到的变量  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取所有文本元素********************** 命令原型: arrayRet = Mage.ExtractSentenceText(jsonRet) 入参: jsonRet--使用"屏幕文字识别"、"图像文字识别"、"PDF文字识别"命令输出到的变量 出参: arrayRet:函数调用的输出保存到的变量 ****************************************************/ Dim jsonRet = {"ai_function" : "ocr_text","items" : [{"char_positions" : [],"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格","handwrite_info" : null,"importance_info" : null,"page_number" : 1,"positions" : [{"x" : 71,"y" : 60},{"x" : 522,"y" : 61},{"x" : 522,"y" : 75},{"x" : 71,"y" : 74}],"probabilities" : []},{"char_positions" : [],"content" : "式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","handwrite_info" : null,"importance_info" : null,"page_number" : 1,"positions" : [{"x" : 72,"y" : 76},{"x" : 420,"y" : 76},{"x" : 420,"y" : 89},{"x" : 72,"y" : 89}],"probabilities" : []}],"struct_content" : {"page" : [{"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","page_id" : 0,"page_number" : 1}],"paragraph" : [{"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","page_number" : 1,"paragraph_id" : 0}],"row" : [{"content" : "将PDF指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回JSON格","page_number" : 1,"row_id" : 0},{"content" : "式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额","page_number" : 1,"row_id" : 1}]}} // 测试数据 Dim arrayRet="" // 输出结果 arrayRet = Mage.ExtractSentenceText(jsonRet) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageText_图片/Mage_ExtractSentenceText.png)  

---

## 查找文本位置

**说明**: 使用 Laiye Intelligent Document Processing 查找文本位置，成功返回字典类型的文本位置，失败引发异常  

**原型**: `objPoint = Mage.Find(objElement, objRect,config,sText, sRule, iOccurrence, iTimeOut, optionArgs)`  

**参数**:  
- **objElement** (True) [decorator] 默认:@ui"" - 通过鼠标选取或截取需要识别的目标屏幕范围。包含窗口、元素、范围等信息  
- **objRect** (True) [dictionary] 默认:{ "x":0,"y":0,"width":0,"height":0 } - 需要进行OCR文字识别的范围，程序会在控件这个范围内进行文字识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的文字识别  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **sText** (True) [string] 默认:"" - 查找元素时使用的文本  
- **sRule** (True) [enum] 默认:"instr" - 查找文本时使用的规则  
- **iOccurrence** (True) [number] 默认:1 - 如果“文本”字段中的字符串在指示的界面元素中出现多次，请在此处指定要单击的出现次数。例如，如果字符串出现4次并且您要单击第一个匹配项，请在此字段中写入1  
- **iTimeOut** (True) [number] 默认:30000 - 指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发错误。默认30000毫秒（30秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: objPoint，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************查找文本位置********************** 命令原型: objPoint = Mage.Find(@ui"", {"x":0,"y":0,"width":0,"height":0},{},"", "instr", 1, 30000, {"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) 入参: objElement--目标,该示例中使用的是chrome打开百度首页的“百度一下”按钮元素 objRect--识别范围.默认值:{"x":0,"y":0,"width":0,"height":0} config--mage配置,需配置Pubkey和Secret.Type:Dict sText--查找文本 sRule--查找规则 iOccurrence--相似结果位置 iTimeOut--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时/光标位置/横坐标偏移/纵坐标偏移/辅助按键/操作类型).Type:Dict 出参: objPoint--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim objPoint="" // 输出结果 objPoint = Mage.Find(@ui"输入控件<input>_百度一下", {"x":0,"y":0,"width":0,"height":0},{"Pubkey":"sCXf4tfmGpq8um0rY9MOvApD","Secret":"KLAhZgzHVqb975HAywi5sAhbxkakSHGx","Url":"https://demo.laiye.com:8082"},"百度一下", "instr", 1, 30000, {"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) TracePrint objPoint
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageText_图片/Mage_Find.png)  

---

## 鼠标移动到文本上

**说明**: 使用 Laiye Intelligent Document Processing 对界面元素范围内进行指定文字识别，如果识别到指定文字将光标移动到文本所在的位置上  

**原型**: `Mage.Hover(objUiElement,objRect,config,sText,sRule,iOccurrence,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 通过鼠标选取或截取需要识别的目标屏幕范围。包含窗口、元素、范围等信息  
- **objRect** (True) [dictionary] 默认:{ "x":0,"y":0,"width":0,"height":0 } - 需要进行OCR文字识别的范围，程序会在控件这个范围内进行文字识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的文字识别  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **sText** (True) [string] 默认:"" - 查找元素时使用的文本  
- **sRule** (True) [enum] 默认:"instr" - 查找文本时使用的规则  
- **iOccurrence** (True) [number] 默认:1 - 如果“文本”字段中的字符串在指示的界面元素中出现多次，请在此处指定要单击的出现次数。例如，如果字符串出现4次并且您要单击第一个匹配项，请在此字段中写入1  
- **iTimeOut** (True) [number] 默认:30000 - 指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发错误。默认30000毫秒（30秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **sCursorPosition** (False) [enum] 默认:"Center" - 描述添加OffsetX和OffsetY属性的偏移量的光标起点。可以使用以下选项：TopLeft，TopRight，BottomLeft，BottomRight和Center。默认选项是Center  
- **iCursorOffsetX** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的水平位移  
- **iCursorOffsetY** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的垂直位移  
- **sKeyModifiers** (False) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **sSimulate** (False) [enum] 默认:"simulate" - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：模拟操作(simulate)  

**示例**:  
```
/**********************鼠标移动到文本上********************** 命令原型: Mage.Hover(@ui"",{"x":0,"y":0,"width":0,"height":0},{},"","instr",1,30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate"}) 入参: objUiElement--目标,该示例中使用的是chrome打开百度首页的“百度一下”按钮元素 objRect--识别范围.默认值:{"x":0,"y":0,"width":0,"height":0} config--mage配置,需配置Pubkey和Secret.Type:Dict sText--查找文本 sRule--查找规则 iOccurrence--相似结果位置 iTimeOut--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时/光标位置/横坐标偏移/纵坐标偏移/辅助按键/操作类型).Type:Dict 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Mage.Hover(@ui"输入控件<input>_百度一下",{"x":0,"y":0,"width":0,"height":0},{"Pubkey":"sCXf4tfmGpq8um0rY4MOvApD","Secret":"KLAhZgzHVqb475HAywi5sAhbxkakSHGx","Url":"https://demo.laiye.com:8082"},"百度一下","instr",1,30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate"})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageText_图片/Mage_Hover.png)  

---

## 图像文字识别

**说明**: 使用 Laiye Intelligent Document Processing 识别指定图像的文字，识别结果返回 JSON 格式  

**原型**: `jsonRet = Mage.ImageOCRText(path,config,time)`  

**参数**:  
- **path** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要识别图像的路径，支持 jpeg、jpg、png、bmp、tif、tiff 等格式，PDF格式仅支持首页识别  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: jsonRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************图像文字识别********************** 命令原型: jsonRet = Mage.ImageOCRText(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,{},30000) 入参: path--待识别图片的路径.Type:String config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int 出参: jsonRet:函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim path=&#x27;&#x27;&#x27;&#x27;&#x27;&#x27; // 待识别图片的路径 Dim config={"Pubkey":"","Secret":"","Url":""} // 从mage中获取 Dim jsonRet="" // 输出结果 jsonRet = Mage.ImageOCRText(path,{"Pubkey":"sCXf4tfmGpq8um0rY4MOvApD","Secret":"KLAhZgzHVqb945HAywi5sAhbxkakSHGx","Url":"https://demo.laiye.com:8082"},30000) TracePrint(jsonRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageText_图片/Mage_ImageOCRText.png)  

---

## PDF文字识别

**说明**: 将 PDF 指定的页码通过 Laiye Intelligent Document Processing 通用文字识别，识别结果返回 JSON 格式。在识别多页过程中如果其中一页失败则整个识别会返回错误，且会消耗配额  

**原型**: `jsonRet = Mage.PDFOCRText(config, path,password,all_pg_state,page_cfg,sleepTime,time)`  

**参数**:  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **path** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - PDF文件路径  
- **password** (True) [string] 默认:"" - PDF文件密码，无密码不需要填写  
- **all_pg_state** (True) [boolean] 默认:False - 当全部页码设为"是"，则识别全部且指定页码输入无效。设为否时，可指定页码识别  
- **page_cfg** (True) [expression] 默认:[ [1,2] ] - 支持正整数和数组格式，如输入2，则识别第2页；如输入 [1,3,5] ，则识别第1,3,5页；如输入[1, [6,9] ,4]，则识别1,4页和第6到第9页。当识别全部页码设为"是"，则识别指定页码的输入失效。超出PDF页码总数的部分会报错，页码重叠部分仅识别1次  
- **sleepTime** (True) [number] 默认:10000 - 识别PDF每页的间隔时长（以毫秒为单位），默认10000毫秒(10秒)。识别页数较多，间隔较短可能会导致调用频率超限错误  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: jsonRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************PDF文字识别********************** 命令原型: jsonRet = Mage.PDFOCRText({}, &#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"",false,[[1,2]],10000,30000) 入参: config--mage配置,需配置Pubkey和Secret.Type:Dict path--待识别图片的路径.Type:String password--密码.无密码则不需要填写.Type:String all_pg_state--是否识别全部页.Type:Bool page_cfg--识别指定页码.Type:List sleepTime--间隔时间.默认单位:毫秒.Type:Int time--超时时间.默认单位:毫秒.Type:Int 出参: jsonRet:函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim path=&#x27;&#x27;&#x27;&#x27;&#x27;&#x27; // 待识别PDF的路径 Dim jsonRet="" // 输出结果 jsonRet = Mage.PDFOCRText({"Pubkey":"sCXf4tfmGpq8um0rY9MOvApD","Secret":"KLAhZgzHVqb975HAywi5sAhbxkakSHGx","Url":"https://demo.laiye.com:8082"}, path,"",false,[[1,2]],10000,30000) TracePrint(jsonRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageText_图片/Mage_PDFOCRText.png)  

---

## 屏幕文字识别

**说明**: 使用 Laiye Intelligent Document Processing 识别指定屏幕范围的文字，识别结果返回 JSON 格式  

**原型**: `jsonRet = Mage.ScreenOCRText(target,rect,config,time,optionArgs)`  

**参数**:  
- **target** (True) [decorator] 默认:@ui"" - 通过鼠标选取或截取需要识别的目标屏幕范围。包含窗口、元素、范围等信息  
- **rect** (True) [dictionary] 默认:{ "height":0,"width":0,"x":0,"y":0 } - 需要查找的范围，程序会在控件这个范围内进行识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的识别  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: jsonRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************屏幕文字识别********************** 命令原型: jsonRet = Mage.ScreenOCRText(@ui"",{"height":0,"width":0,"x":0,"y":0},{},30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) 入参: target--目标元素,该示例中使用的为chrome打开百度后的“百度一下”按钮元素,需自行录制 rect--识别范围.Type:Dict config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参: jsonRet:函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim jsonRet="" // 输出结果 jsonRet = Mage.ScreenOCRText(@ui"输入控件<input>_百度一下",{"height":0,"width":0,"x":0,"y":0},{"Pubkey":"sCXf4tfmGpq8um0rY3MOvApD","Secret":"KLAhZgzHVqb955HAywi5sAhbxkakSHGx","Url":"https://demo.laiye.com:8082"},30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) TracePrint(jsonRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageText_图片/Mage_ScreenOCRText.png)  

---

## 获取排名结果

**说明**: 获取文本分类的排名结果  

**原型**: `arrayTopNRet = Mage.ExtractTextClassifyTopN(arrayRet, thrd, top_n)`  

**参数**:  
- **arrayRet** (True) [expression] 默认:arrayRet - 使用文本分类命令输出到的变量  
- **thrd** (True) [number] 默认:0.6 - 对分类结果的置信度范围设置阈值，支持输入0~1之间的小数，筛选大于等于阈值的结果  
- **top_n** (True) [number] 默认:1 - 支持输入从1开始的正整数，如输入5，置信度阈值为0.6，则返回排名前5且置信度大于等于0.6的结果  

**返回**: arrayTopNRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************获取排名结果********************** 命令原型: arrayTopNRet = Mage.ExtractTextClassifyTopN(arrayRet, 0.6, 1) 入参: arrayRet--文本分类结果.Type:List(文本分类的识别结果) thrd--置信度.Type:0~1之间的小数 top_n--获取前几名.Type:Int 出参: arrayTopNRet--函数调用的输出保存到的变量 注意事项： 需要搭配文本分类命令（NLPTextClassify）使用 ****************************************************/ Rem 测试数据 Dim arrayRet = [{"ai_function" : "nlp_text_classify","class_id" : 1473,"class_label" : "国际","debug_info" : [],"score" : 0.6527229},{"ai_function" : "nlp_text_classify","class_id" : 1469,"class_label" : "时政","debug_info" : [],"score" : 0.230379},{"ai_function" : "nlp_text_classify","class_id" : 1472,"class_label" : "科技","debug_info" : [],"score" : 0.05750105},{"ai_function" : "nlp_text_classify","class_id" : 1471,"class_label" : "体育","debug_info" : [],"score" : 0.037129827},{"ai_function" : "nlp_text_classify","class_id" : 1470,"class_label" : "财经","debug_info" : [],"score" : 0.022317171}] Dim arrayTopNRet="" // 输出结果 arrayTopNRet = Mage.ExtractTextClassifyTopN(arrayRet, 0.4, 1) TracePrint(arrayTopNRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTextClassify_图片/Mage_ExtractTextClassifyTopN.png)  

---

## 文本分类

**说明**: 对指定文本进行分类，需提前在 Laiye Intelligent Document Processing 后台训练分类模型  

**原型**: `arrayRet = Mage.NLPTextClassify(doc,config,time)`  

**参数**:  
- **doc** (True) [string] 默认:"" - 输入待分类的文本信息  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************文本分类********************** 命令原型: arrayRet = Mage.NLPTextClassify("",{},30000) 入参: doc--待分类文本.Type:String config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int 出参: arrayRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Rem 测试数据 Dim doc="马克龙宣布法国警务改革措施：增加警察人数 加强执法监督 中新社巴黎9月14日电 (记者 李洋)法国总统马克龙当地时间14日宣布法国警务改革的相关措施，其中包括增加警察人数、加强执法监督、提高警察待遇、简化调查程序等。 马克龙当天到访法国北部城市鲁贝，他在当地举行的安全事务圆桌会议闭幕式上宣布法国警务改革的“完整战略”，以改善法国的治安，保障民众的利益。" Dim arrayRet="" // 输出结果 arrayRet = Mage.NLPTextClassify(doc,{"Pubkey":"MLmCwXOgGKVxsFxUdhAoFr71","Secret":"ZwAiUnEW2qaOgMhhX8bwfL3bQDVuNnor","Url":"https://demo.laiye.com:8082"},3000) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTextClassify_图片/Mage_NLPTextClassify.png)  

---

## 获取字段结果

**说明**: 获取信息抽取结果中指定字段的结果  

**原型**: `sRet = Mage.ExtractTextExtractInfo(value,extractor,template_name,field_name,update_time,index,std_state)`  

**参数**:  
- **value** (True) [expression] 默认:value - 使用"信息抽取"命令输出到的变量并循环遍历的值  
- **extractor** (True) [expression] 默认:{ } - 选择信息抽取的抽取器  
- **template_name** (True) [string] 默认:"" - 选择信息抽取的模板  
- **field_name** (True) [string] 默认:"" - 选择模板中的输出字段  
- **update_time** (True) [string] 默认:"" - 不可修改，动态获取版本更新时间，如果与信息抽取的结果使用的版本不一致，则在运行时提示  
- **index** (True) [number] 默认:0 - 指定模板中相同字段的索引结果，从0开始，超出索引范围会中断流程且报错  
- **std_state** (True) [boolean] 默认:False - 为"否"时取原值，为"是"时取归一化值  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取字段名称列表*************************************** 命令原型： sRet = Mage.ExtractTextExtractInfo(value,{},"","","",0,false) 入参： value--信息抽取结果。注：使用"信息抽取"命令输出到的变量并循环遍历的值 extractor--抽取器。注：选择信息抽取的抽取器(配置包括Pubkey、Secret、Url，其中Pubkey和Secret需要在Mage指定模型里面进行获取) template_name--模板名称。注：选择信息抽取的模板 field_name--字段名称。注：选择模板中的输出字段 update_time--更新时间。注：不可修改，动态获取版本更新时间，如果与信息抽取的结果使用的版本不一致，则在运行时提示 index--指定结果索引。注：指定模板中相同字段的索引结果，从0开始，超出索引范围会中断流程且报错 std_state--取归一化值。注：为"否"时取原值，为"是"时取归一化值 出参： sRet--函数调用的输出保存到的变量。 注意事项： 1.要使用mage，保证能够连接到网络。 2.mage在使用时有次数限制，要注意mage次数余量。 3.可以结合文本信息抽取命令使用。 ********************************************************************************/ Dim arrayRet = "" Dim sRet = "" arrayRet = Mage.NLPTextExtract("公司去年经营性收入30万元，处于较好态势。另外，税后净利润为5万元，较之前同比增长12%。公司拥有博士 30多人，硕士60人。",{"Pubkey":"8ffHXjLv0u6nFiYM6WorRchN","Secret":"USMv2z9fABExxeeNRyWpvZwVaKcVbmLO","Url":"https://mage.uibot.com.cn"},30000) sRet = Mage.ExtractTextExtractInfo(arrayRet[0],{"Pubkey":Pubkey,"Secret":Secret,"Url":"https://mage.uibot.com.cn"},"净利润","税后净利润","2021-09-18 09:50:38",0,False) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTextExtract_图片/MageTextExtract_ExtractTextExtractInfo.png)  

---

## 获取模板名称

**说明**: 获取信息抽取结果中的模板名称  

**原型**: `sRet = Mage.ExtractTextExtractName(value)`  

**参数**:  
- **value** (True) [expression] 默认:value - 使用"信息抽取"命令输出到的变量并循环遍历的值  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取模板名称*************************************** 命令原型： sRet = Mage.ExtractTextExtractName(value) 入参： value--信息抽取命令抽取的识别内容。注：使用"信息抽取"命令输出到的变量并循环遍历的值 出参： sRet--函数调用的输出保存到的变量。 注意事项： 1.要结合文本或者文件信息抽取命令使用。 ********************************************************************************/ Dim arrayRet = "" Dim sRet = "" arrayRet = Mage.NLPTextFileExtract(@res"评测集-上市公司的披露公告.txt",{"Pubkey":Pubkey,"Secret":Secret,"Url":"https://mage.uibot.com.cn"},30000) For Each value In arrayRet sRet = Mage.ExtractTextExtractName(value) TracePrint(sRet) Next
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTextExtract_图片/MageTextExtract_ExtractTextExtractName.png)  

---

## 获取字段名称列表

**说明**: 从 Laiye Intelligent Document Processing 接口获取抽取器中信息抽取模板的字段列表  

**原型**: `arrayRet = Mage.GetTextExtractFieldList(extractor,template_name,time)`  

**参数**:  
- **extractor** (True) [expression] 默认:{ } - 选择信息抽取的抽取器  
- **template_name** (True) [string] 默认:"" - 选择信息抽取的模板  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取字段名称列表*************************************** 命令原型： arrayRet = Mage.GetTextExtractFieldList({},"",30000) 入参： extractor--抽取器。注：选择信息抽取的抽取器.(包括:Pubkey、Secret、Url，其中Pubkey和Secret需要在Mage指定模型里面进行获取).Type:Dict template_name--模板名称。注：选择信息抽取的模板 time--超时时间（毫秒）。注：指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒） 出参： arrayRet--函数调用的输出保存到的变量。 注意事项： 1.要使用mage，保证能够连接到网络。 2.要在mage中提前配置好模型以及对应配置信息。 ********************************************************************************/ Dim arrayRet = "" arrayRet = Mage.GetTextExtractFieldList({"Pubkey":Pubkey,"Secret":Secret,"Url":"https://mage.uibot.com.cn"},"同比减少",30000) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTextExtract_图片/MageTextExtract_GetTextExtractFieldList.png)  

---

## 文本信息抽取

**说明**: 对文本进行信息抽取，需在 Laiye Intelligent Document Processing 后台配置抽取模板  

**原型**: `arrayRet = Mage.NLPTextExtract(doc,config,time)`  

**参数**:  
- **doc** (True) [expression] 默认:"" - 输入待抽取的文本，如果在 Laiye Intelligent Document Processing 的抽取器设置中启用了“默认换行结束匹配”，则抽取前模型会先通过 \n 来切分文本，然后再进行匹配  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************文本信息抽取*************************************** 命令原型： arrayRet = Mage.NLPTextExtract("",{},30000) 入参： doc--待抽取文本。注：输入待抽取的文本，如果在 Laiye Intelligent Document Processing 的抽取器设置中启用了“默认换行结束匹配”，则抽取前模型会先通过 \n 来切分文本，然后再进行匹配 config--Mage配置。配置包括:Pubkey、Secret、Url，其中Pubkey和Secret需要在Mage指定模型里面进行获取.Type:Dict time--超时时间（毫秒）。注：指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒） 出参： arrayRet--函数调用的输出保存到的变量。 注意事项： 1.要使用mage，保证能够连接到网络。 2.mage在使用时有次数限制，要注意mage次数余量。 ********************************************************************************/ Dim arrayRet = "" arrayRet = Mage.NLPTextExtract("公告一 持有本公司股份876,000股（占本公司总股本0.0829%）的副总经理、董事会秘书张颜拟自本公告起十五个交易日后的六个月内，以集中竞价或大宗交易方式减持本公司股份不超过200,000股（占公司总股本的0.0189%）。 公告二 持有本公司股份885,000股（占本公司总股本0.0838%）的副总经理、财务总监陈圆圆拟自本公告起十五个交易日后的六个月内，以集中竞价或大宗交易方式减持本公司股份不超过206,000股（占公司总股本的0.0195%）。",{"Pubkey":Pubkey,"Secret":Secret,"Url":"https://mage.uibot.com.cn"},30000) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTextExtract_图片/MageTextExtract_NLPTextExtract.png)  

---

## 文件信息抽取

**说明**: 对文本文件进行信息抽取，需在 Laiye Intelligent Document Processing 后台配置抽取模板  

**原型**: `arrayRet = Mage.NLPTextFileExtract(file,config,time)`  

**参数**:  
- **file** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 选择待信息抽取的文本文件（UTF-8 编码）路径  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************文件信息抽取*************************************** 命令原型： arrayRet = Mage.NLPTextFileExtract(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,{},30000) 入参： file--待抽取文件路径。默认编码格式:UTF-8 config--Mage配置。配置包括：Pubkey、Secret、Url，其中Pubkey和Secret需要在Mage指定模型里面进行获取.Type:Dict time--超时时间（毫秒）。注：指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒） 出参： arrayRet--函数调用的输出保存到的变量。 注意事项： 1.要使用mage，保证能够连接到网络。 2.mage在使用时有次数限制，要注意mage次数余量。 3.要保证读取的文件在本地存在。 ********************************************************************************/ Dim arrayRet = "" arrayRet = Mage.NLPTextFileExtract(@res"评测集-上市公司的披露公告.txt",{"Pubkey":Pubkey,"Secret":Secret,"Url":"https://mage.uibot.com.cn"},30000) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageTextExtract_图片/MageTextExtract_NLPTextFileExtract.png)  

---

## 图像验证码识别

**说明**: 使用 Laiye Intelligent Document Processing 识别指定图像的验证码，返回识别结果  

**原型**: `sRet = Mage.ImageOCRVerifyCode(path,config,time)`  

**参数**:  
- **path** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要识别图像的路径，支持 jpeg、jpg、png、bmp、gif 等格式  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************图像验证码识别********************** 命令原型: sRet = Mage.ImageOCRVerifyCode(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,{},30000) 入参: path--待识别图片路径 config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int 出参: sRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim sRet="" // 输出结果 sRet = Mage.ImageOCRVerifyCode(&#x27;&#x27;&#x27;E:\Desktop\1.jpg&#x27;&#x27;&#x27;,{"Pubkey":"684wM60rdbDnciVS7iXgZmNL","Secret":"AYzWVO0srKrN5t9rhL0jfhS05uaJnj92","Url":"https://demo.laiye.com:8082"},30000) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageVerifyCode_图片/Mage_ImageOCRVerifyCode.png)  

---

## 屏幕验证码识别

**说明**: 使用 Laiye Intelligent Document Processing 识别指定屏幕范围的验证码，返回识别结果  

**原型**: `sRet = Mage.ScreenOCRVerifyCode(target,rect,config,time,optionArgs)`  

**参数**:  
- **target** (True) [decorator] 默认:@ui"" - 通过鼠标选取或截取需要识别的目标屏幕范围。包含窗口、元素、范围等信息  
- **rect** (True) [dictionary] 默认:{ "x":0,"y":0,"width":0,"height":0 } - 需要查找的范围，程序会在控件这个范围内进行识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的识别  
- **config** (True) [expression] 默认:{ } - Laiye Intelligent Document Processing 的调用配置  
- **time** (True) [number] 默认:30000 - 指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************屏幕验证码识别********************** 命令原型: sRet = Mage.ScreenOCRVerifyCode(@ui"",{"x":0,"y":0,"width":0,"height":0},{},30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) 入参: target--目标元素,该示例中使用的为chrome打开"北京电子税务局"后的图片验证码元素.Url:https://etax.beijing.chinatax.gov.cn/sso/login?service=http://etax.beijing.chinatax.gov.cn/xxmh/html/index_login.html?t=1643264339708 rect--识别范围 config--mage配置,需配置Pubkey和Secret.Type:Dict time--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参: sRet--函数调用的输出保存到的变量 注意事项： 需要获取mage对应的Key/Secret和URL ****************************************************/ Dim sRet="" // 输出结果 sRet = Mage.ScreenOCRVerifyCode(@ui"图像<img>1",{"x":0,"y":0,"width":0,"height":0},{"Pubkey":"684wM60rdbDnciVS4iXgZmNL","Secret":"AYzWVO0srKrN8t3rhL0jfhS05uaJnj92","Url":"https://demo.laiye.com:8082"},30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/MageVerifyCode_图片/Mage_ScreenOCRVerifyCode.png)  

---

## 连接邮箱

**说明**: 连接一个邮箱，并作为操控对象  

**原型**: `objMail = Mail.Connect(sServer, sUid, sPwd, sType, iPort, bSsl)`  

**参数**:  
- **sServer** (True) [string] 默认:"" - 服务器地址  
- **sUid** (True) [string] 默认:"" - 登录帐号  
- **sPwd** (True) [string] 默认:"" - 登录密码  
- **sType** (True) [enum] 默认:"POP3" - 使用协议  
- **iPort** (True) [number] 默认:110 - POP3服务器端口，默认为 110，一般不需要修改  
- **bSsl** (True) [boolean] 默认:False - 是否使用SSL协议加密，默认为false  

**返回**: objMail，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************连接邮箱*************************************** 命令原型： objMail = Mail.Connect("", "", "", "POP3", 110, false) 入参： sServer--服务器地址。注：服务器地址 sUid--登录账号。注：登录帐号 sPwd--登陆密码。注：登录密码 cType--使用协议。默认协议:POP3 iPort--服务器端口。注：POP3服务器端口，默认为 110，一般不需要修改 bSsl--SSL加密。注：是否使用SSL协议加密，默认为false 出参： objMail--函数调用的输出保存到的变量。 注意事项： 1.邮箱协议以及对应端口要对应，并且注意邮箱服务器存在SSL加密。 2.要保证目标邮箱能够被连接。 ********************************************************************************/ Dim objMail = "" objMail = Mail.Connect("127.0.0.1", "root", "test123", "POP3", 110, false) TracePrint(objMa)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Mail_图片/Mail_Connect.png)  

---

## 删除邮件

**说明**: 指定邮件对象删除对应邮件，在使用该命令删除邮件后，必须调用断开邮箱连接（Mail.Disconnect）命令，才能真正删除成功。如果邮件服务器设置了“禁止收信软件删除邮件”，则依然无法删除  

**原型**: `Mail.Delete(objMail,mailJSONObject)`  

**参数**:  
- **objMail** (True) [expression] 默认:objMail - 邮箱对象，使用连接邮箱（Mail.Connect）命令返回的邮箱对象  
- **mailJSONObject** (True) [expression] 默认:{ } - 使用获取邮件列表（Mail.GetMailList）命令返回邮件数组中的邮件对象  

**示例**:  
```
/*********************************删除邮件*************************************** 命令原型： Mail.Delete(objMail,{}) 入参： objMail--邮箱对象。注：邮箱对象，使用连接邮箱（Mail.Connect）命令返回的邮箱对象 mailJSONObject--邮件对象。 出参： objMail--函数调用的输出保存到的变量。 注意事项： 1.要结合连接邮箱命令以及获取邮件列表命令使用。 2.要确保邮件服务器没有设置“禁止收信软件删除邮件”，否则无法删除邮件。 ********************************************************************************/ Mail.Delete(objMail,arrayRet[0])
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Mail_图片/Mail_Delete.png)  

---

## 断开邮箱连接

**说明**: 断开通过连接命令（Connect）连接上的邮箱服务器的连接  

**原型**: `Mail.Disconnect(objMail)`  

**参数**:  
- **objMail** (True) [expression] 默认:objMail - 邮箱对象，使用连接邮箱（Mail.Connect）命令返回的邮箱对象  

**示例**:  
```
/*********************************断开邮箱连接*************************************** 命令原型： Mail.Disconnect(objMail) 入参： objMail--邮箱对象。注：邮箱对象，使用连接邮箱（Mail.Connect）命令返回的邮箱对象 注意事项： 1.要结合连接邮箱命令使用。 ********************************************************************************/ Mail.Disconnect(objMail)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Mail_图片/Mail_Disconnect.png)  

---

## 下载附件

**说明**: 下载邮件中的附件  

**原型**: `arrayRet = Mail.DownloadAttachment(objMail,mailJSONObject,sAttrPath)`  

**参数**:  
- **objMail** (True) [expression] 默认:objMail - 邮箱对象，使用连接邮箱（Mail.Connect）命令返回的邮箱对象  
- **mailJSONObject** (True) [expression] 默认:{ } - 使用获取邮件列表（Mail.GetMailList）命令返回邮件数组中的邮件对象  
- **sAttrPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 邮件附件保存的路径，可填写绝对路径也可使用@res"路径"形式表示当前流程res文件夹下的路径，路径分隔符需转义，如"C: \ Laiye RPA"或@res"Laiye RPA \ Laiye RPA"  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************下载附件*************************************** 命令原型： arrayRet = Mail.DownloadAttachment(objMail,{},&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： objMail--邮箱对象。注：邮箱对象，使用连接邮箱（Mail.Connect）命令返回的邮箱对象 mailJSONObject--邮件对象。注：使使用获取邮件列表（Mail.GetMailList）命令返回邮件数组中的邮件对象 sAttrPath--邮件附件保存的路径。 出参： objMail--函数调用的输出保存到的变量。 注意事项： 1.要结合连接邮箱命令使用。 2.目标邮件中要存在对应附件。 ********************************************************************************/ Dim arrayRet = "" arrayRet = Mail.DownloadAttachment(objMail,arrayRet[0],&#x27;&#x27;&#x27;C:\Users\来也科技\Desktop\123.jpeg&#x27;&#x27;&#x27;) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Mail_图片/Mail_DownloadAttachment.png)  

---

## 获取邮件列表

**说明**: 获取收件箱中的邮件列表，列表为一个数组，数组中的每一项为邮件对象  

**原型**: `arrayRet = Mail.GetMailList(objMail, iCount)`  

**参数**:  
- **objMail** (True) [expression] 默认:objMail - 邮箱对象，使用连接邮箱（Mail.Connect）命令返回的邮箱对象  
- **iCount** (True) [number] 默认:0 - 设置需要获取的邮件数量，设置0为获取收件箱中的所有邮件  

**返回**: arrayRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取邮件列表*************************************** 命令原型： arrayRet = Mail.GetMailList(objMail, 0) 入参： objMail--邮箱对象。注：邮箱对象，使用连接邮箱（Mail.Connect）命令返回的邮箱对象 iCount--邮件对象。注：使用获取邮件列表（Mail.GetMailList）命令返回邮件数组中的邮件对象 出参： arrayRet--函数调用的输出保存到的变量。 注意事项： 1.要结合连接邮箱命令使用。 ********************************************************************************/ Dim arrayRet = "" arrayRet = Mail.GetMailList(objMail, 0) TracePrint(arrayRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Mail_图片/Mail_GetMailList.png)  

---

## 发送邮件

**说明**: 发送邮件到指定邮箱  

**原型**: `bRet = Mail.SendEx(sServer,iPort,bSsl,sUid,sPwd,sSender,sTo,sCc,sTitle,sContent,sAttr)`  

**参数**:  
- **sServer** (True) [string] 默认:"" - SMTP服务器地址  
- **iPort** (True) [number] 默认:25 - SMTP服务器端口，常见为 25、465、587  
- **bSsl** (True) [boolean] 默认:False - 是否使用SSL协议加密，默认为否  
- **sUid** (True) [string] 默认:"" - 邮箱登录帐号，比如普通QQ邮箱的登录帐号与邮箱地址相同  
- **sPwd** (True) [string] 默认:"" - 登录密码  
- **sSender** (True) [string] 默认:"" - 发件人邮箱地址  
- **sTo** (True) [string] 默认:"" - 收件人邮箱地址，多个地址可用 ["abc@ui.bot","xyz@ui.bot"] 数组的形式填写  
- **sCc** (True) [string] 默认:"" - 抄送邮箱地址，多个地址可用 ["abc@ui.bot","xyz@ui.bot"] 数组的形式填写  
- **sTitle** (True) [string] 默认:"" - 邮件的标题  
- **sContent** (True) [string] 默认:"" - 邮件正文内容，支持HTML类型的正文内容  
- **sAttr** (True) [array] 默认:[&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;] - 邮件附件，多个附件可以用 ["附件一路径","附件二路径"] 数组的形式填写  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************发送邮件*************************************** 命令原型： bRet = Mail.SendEx("",25,false,"","","","","","","",[&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;]) 入参： sServer--服务器端口。注：SMTP服务器地址 iPort--服务器端口。注：SMTP服务器端口，默认25，常见为 25、465、587 bSsl--SSL加密。注：是否使用SSL协议加密，默认为false sUid--登录账号。注：邮箱登录帐号，比如普通QQ邮箱的登录帐号与邮箱地址相同 sPwd--登陆密码。注：登录密码 sSender--发件人。注：发件人邮箱地址 sTo--收件人。注：收件人邮箱地址，多个地址可用["abc@ui.bot","xyz@ui.bot"]数组的形式填写.Type:list sCc--抄送人。注：抄送邮箱地址，多个地址可用["abc@ui.bot","xyz@ui.bot"]数组的形式填写.Type:list sTitle--邮件标题。注：邮件的标题 sContent--邮箱内容。注：邮件正文内容，支持HTML类型的正文内容 sAttr--邮箱附件。注：邮件附件，多个附件可以用["附件一路径","附件二路径"]数组的形式填写.Type:list 出参： bRet--函数调用的输出保存到的变量。 注意事项： 1.邮箱协议以及对应端口要对应，并且注意邮箱服务器存在SSL加密。 2.要保证目标邮箱能够被连接。 ********************************************************************************/ Dim bRet = "" bRet = Mail.SendEx("smtp.feishu.cn",465,true,&#x27;12318asd@qq.com&#x27;, &#x27;test123&#x27;, &#x27;12318asd@qq.com&#x27;,[&#x27;asdasdasd@qq.com&#x27;,&#x27;asdasd23d@qq.com&#x27;],[&#x27;asdasdasd@qq.com&#x27;,&#x27;asdasd23d@qq.com&#x27;],"邮件标题","邮件内容",[@res"第一个附件.txt"]) TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Mail_图片/Mail_SendEx.png)  

---

## 取绝对值

**说明**: 获取数值的绝对值  

**原型**: `dRet = Math.Abs(dNum)`  

**参数**:  
- **dNum** (True) [number] 默认:0 - 要处理的数据  

**返回**: dRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取绝对值*************************************** 命令原型： Math.Abs(0) 入参： dNum -- 要处理的数据 出参： dRet -- 将命令运行后的结果赋值给此变量 注意事项： 入参dNum类型为number **********************************************************************************/ Dim dRet dRet = Math.Abs(-1.23) TracePrint(dRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Math_图片/Math_Abs.png)  

---

## 取反正切值

**说明**: 获取属性 dNum 的反正切值  

**原型**: `dRet = Math.Atn(dNum)`  

**参数**:  
- **dNum** (True) [number] 默认:0 - 要处理的数据  

**返回**: dRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取反正切值*************************************** 命令原型： Math.Atn(0) 入参： dNum -- 要处理的数据 出参： dRet -- 将命令运行后的结果赋值给此变量 注意事项： 入参dNum类型为number **********************************************************************************/ Dim dRet dRet = Math.Atn(90) TracePrint(dRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Math_图片/Math_Atn.png)  

---

## 取余弦值

**说明**: 获取属性 dNum 的余弦值  

**原型**: `dRet = Math.Cos(dNum)`  

**参数**:  
- **dNum** (True) [number] 默认:0 - 要处理的数据  

**返回**: dRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取余弦值*************************************** 命令原型： Math.Cos(0) 入参： dNum -- 要处理的数据 出参： dRet -- 将命令运行后的结果赋值给此变量 注意事项： 入参dNum类型为number **********************************************************************************/ Dim dRet dRet = Math.Cos(30) TracePrint(dRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Math_图片/Math_Cos.png)  

---

## 取自然对数e的N次幂

**说明**: 获取自然对数e的 dNum 次幂  

**原型**: `dRet = Math.Exp(dNum)`  

**参数**:  
- **dNum** (True) [number] 默认:0 - 要处理的数据  

**返回**: dRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取自然对数e的N次幂*************************************** 命令原型： Math.Exp(0) 入参： dNum -- 要处理的数据 出参： dRet -- 将命令运行后的结果赋值给此变量 注意事项： 入参dNum类型为number **********************************************************************************/ Dim dRet dRet = Math.Exp(2) TracePrint(dRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Math_图片/Math_Exp.png)  

---

## 取整数部分

**说明**: 获取属性 dNum 的整数部分，处理负数时，向下取整  

**原型**: `iRet = Math.Int(dNum)`  

**参数**:  
- **dNum** (True) [number] 默认:0 - 要处理的数据  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取整数部分*************************************** 命令原型： Math.Int(0) 入参： dNum -- 要处理的数据 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 入参dNum类型为number **********************************************************************************/ Dim iRet iRet = Math.Int(2.1) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Math_图片/Math_Int.png)  

---

## 取自然对数

**说明**: 获取属性 dNum 的自然对数  

**原型**: `dRet = Math.Ln(dNum)`  

**参数**:  
- **dNum** (True) [number] 默认:0 - 要处理的数据  

**返回**: dRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取自然对数*************************************** 命令原型： Math.Ln(0) 入参： dNum -- 要处理的数据 出参： dRet -- 将命令运行后的结果赋值给此变量 注意事项： 入参dNum类型为number **********************************************************************************/ Dim dRet dRet = Math.Ln(3) TracePrint(dRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Math_图片/Math_Ln.png)  

---

## 取四舍五入值

**说明**: 获取属性 dNum 的四舍五入值，可以指定保留几位小数  

**原型**: `iRet = Math.Round(dNum,dRetain)`  

**参数**:  
- **dNum** (True) [number] 默认:0 - 要处理的数据  
- **dRetain** (True) [number] 默认:2 - 此属性代表保留小数后几位  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取四舍五入值*************************************** 命令原型： Math.Round(0,2) 入参： dNum -- 要处理的数据 dRetain -- 此属性代表保留小数后几位 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 入参dNum、dRetain类型均为number **********************************************************************************/ Dim iRet iRet = Math.Round(3.12,1) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Math_图片/Math_Round.png)  

---

## 取正负符号

**说明**: 获取属性 dNum 的正负符号，数据为正数时返回 1，数据为负数时返回 -1  

**原型**: `iRet = Math.Sgn(dNum)`  

**参数**:  
- **dNum** (True) [number] 默认:0 - 要处理的数据  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取正负符号*************************************** 命令原型： Math.Sgn(0) 入参： dNum -- 要处理的数据 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 入参dNum类型为number **********************************************************************************/ Dim iRet iRet = Math.Sgn(-3) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Math_图片/Math_Sgn.png)  

---

## 取正弦值

**说明**: 获取属性 dNum 的正弦值  

**原型**: `dRet = Math.Sin(dNum)`  

**参数**:  
- **dNum** (True) [number] 默认:0 - 要处理的数据  

**返回**: dRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取正弦值*************************************** 命令原型： Math.Sin(0) 入参： dNum -- 要处理的数据 出参： dRet -- 将命令运行后的结果赋值给此变量 注意事项： 入参dNum类型为number **********************************************************************************/ Dim dRet dRet = Math.Sin(90) TracePrint(dRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Math_图片/Math_Sin.png)  

---

## 取平方根

**说明**: 获取属性 dNum 的平方根值  

**原型**: `dRet = Math.Sqr(dNum)`  

**参数**:  
- **dNum** (True) [number] 默认:0 - 要处理的数据  

**返回**: dRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取平方根*************************************** 命令原型： Math.Sqr(0) 入参： dNum -- 要处理的数据 出参： dRet -- 将命令运行后的结果赋值给此变量 注意事项： 入参dNum类型为number **********************************************************************************/ Dim dRet dRet = Math.Sqr(4) TracePrint(dRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Math_图片/Math_Sqr.png)  

---

## 取正切值

**说明**: 获取属性 dNum 的正切值  

**原型**: `dRet = Math.Tan(dNum)`  

**参数**:  
- **dNum** (True) [number] 默认:0 - 要处理的数据  

**返回**: dRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取正切值*************************************** 命令原型： Math.Tan(0) 入参： dNum -- 要处理的数据 出参： dRet -- 将命令运行后的结果赋值给此变量 注意事项： 入参dNum类型为number **********************************************************************************/ Dim dRet dRet = Math.Tan(30) TracePrint(dRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Math_图片/Math_Tan.png)  

---

## 点击目标

**说明**: 单击指定的界面元素  

**原型**: `Mouse.Action(objUiElement,iButton,iType,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **iButton** (True) [enum] 默认:"left" - 鼠标按键 {left:左键, right:右键, middle:中键}  
- **iType** (True) [enum] 默认:"click" - 点击类型 {click:单击, dbclick:双击, down:按下, up:弹起}  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sCursorPosition** (False) [enum] 默认:"Center" - 描述添加OffsetX和OffsetY属性的偏移量的光标起点。可以使用以下选项：TopLeft，TopRight，BottomLeft，BottomRight和Center。默认选项是Center  
- **iCursorOffsetX** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的水平位移  
- **iCursorOffsetY** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的垂直位移  
- **sKeyModifiers** (False) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **sSimulate** (False) [enum] 默认:simulate - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：模拟操作(simulate)  
- **bMoveSmoothly** (False) [boolean] 默认:False - 是否平滑移动鼠标  

**示例**:  
```
/************************点击鼠标************************ 命令原型: Mouse.Action(objUiElement,iButton,iType,iTimeOut,optionArgs) 入参: objUiElement--目标元素,点击"百度一下"按钮,使用者根据实际场景重新录制元素即可 iButton--鼠标点击(左键/右键/中键) iType--点击类型(单击/双击/按下/弹起) iTimeOut--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时/激活窗口/光标位置/横坐标偏移/纵坐标偏移/辅助按键/操作类型/平滑移动).Type:Dict 注意事项: 模拟操作：指通过调用系统api mouseevent等实现鼠标操作，会实际移动光标。 系统消息：指发送鼠标消息到目标元素，不移动光标。 后台操作：可以理解为调用了一次元素的鼠标响应回调函数。 *******************************************************/ Mouse.Action(@ui"输入控件<input>_百度一下6","left","click",10000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200, "bSetForeground": true, "sCursorPosition": "Center", "iCursorOffsetX": 0, "iCursorOffsetY": 0, "sKeyModifiers": [],"sSimulate": "simulate", "bMoveSmoothly": false})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Mouse_图片/Mouse_Action.png)  

---

## 模拟点击

**说明**: 模拟鼠标的点击动作  

**原型**: `Mouse.Click(iButton, iType, sKeyModifiers,optionArgs)`  

**参数**:  
- **iButton** (True) [enum] 默认:"left" - 鼠标按键 {left:左键, right:右键, middle:中键}  
- **iType** (True) [enum] 默认:"click" - 点击类型 {click:单击, dbclick:双击, down:按下, up:弹起}  
- **sKeyModifiers** (True) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/***************************模拟点击*************************** 命令原型: Mouse.Click(iButton, iType, sKeyModifiers,optionArgs) 入参: iButton--鼠标点击(左键/右键/中键) iType--点击类型(单击/双击/按下/弹起) sKeyModifiers--辅助按键 optionArgs--可选参数(包括:执行后延时/执行前延时).Type:Dict 出参: 无 注意事项: 必须选定目标 ************************************************************/ // 移动到指定目标上 Mouse.Hover(@ui"输入控件<input>_百度一下7",10000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200, "bSetForeground": true, "sCursorPosition": "Center", "iCursorOffsetX": 10, "iCursorOffsetY": 10, "sKeyModifiers": [],"sSimulate": "simulate", "bMoveSmoothly": false}) // 执行点击操作 Mouse.Click("left", "click", [],{"iDelayAfter": 300, "iDelayBefore": 300})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Mouse_图片/Mouse_Click.png)  

---

## 模拟拖动

**说明**: 将鼠标从某一位置拖动到另一位置  

**原型**: `Mouse.Drag(sx, sy, dx, dy, iButton, sKeyModifiers,optionArgs)`  

**参数**:  
- **sx** (True) [number] 默认:0 - 拖动鼠标的起始横坐标  
- **sy** (True) [number] 默认:0 - 拖动鼠标的起始纵坐标  
- **dx** (True) [number] 默认:0 - 拖动鼠标的结束横坐标  
- **dy** (True) [number] 默认:0 - 拖动鼠标的结束纵坐标  
- **iButton** (True) [enum] 默认:"left" - 鼠标按键 {left:左键, right:右键, middle:中键}  
- **sKeyModifiers** (True) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/***************************模拟拖动***************************** 命令原型: Mouse.Drag(sx, sy, dx, dy, iButton, sKeyModifiers,optionArgs) 入参: sx--起始横坐标 sy--起始纵坐标 dx--结束横坐标 dy--结束纵坐标 iButton--鼠标点击(左键/右键/中键) sKeyModifiers--辅助按键 optionArgs--可选参数(包括:执行后延时/执行前延时).Type:Dict 出参： 无 注意事项: 无 *************************************************************/ Mouse.Drag(0, 0, 100, 100, "left", [],{"iDelayAfter": 300, "iDelayBefore": 200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Mouse_图片/Mouse_Drag.png)  

---

## 获取鼠标位置

**说明**: 获取鼠标光标的位置  

**原型**: `objPoint=Mouse.GetPos()`  

**参数**:  

**返回**: objPoint，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*****************************获取鼠标位置******************************** 命令原型: objPoint=Mouse.GetPos() 入参： 无 出参: objPoint--将命令运行后的结果赋值给此变量 注意事项: 无 *********************************************************************/ objPoint=Mouse.GetPos() TracePrint(objPoint)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Mouse_图片/Mouse_GetPos.png)  

---

## 移动到目标上

**说明**: 光标移动到指定的界面元素上  

**原型**: `Mouse.Hover(objUiElement,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sCursorPosition** (False) [enum] 默认:"Center" - 描述添加OffsetX和OffsetY属性的偏移量的光标起点。可以使用以下选项：TopLeft，TopRight，BottomLeft，BottomRight和Center。默认选项是Center  
- **iCursorOffsetX** (False) [number] 默认:10 - 根据在“位置”字段中选择的选项，光标位置的水平位移  
- **iCursorOffsetY** (False) [number] 默认:10 - 根据在“位置”字段中选择的选项，光标位置的垂直位移  
- **sKeyModifiers** (False) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **sSimulate** (False) [enum] 默认:"simulate" - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：模拟操作(simulate)  
- **bMoveSmoothly** (False) [boolean] 默认:False - 是否平滑移动鼠标  

**示例**:  
```
/****************************移动到目标上********************************* 命令原型: Mouse.Hover(objUiElement,iTimeOut,optionArgs) 入参: objUiElement--目标元素,移动到"百度一下"按钮上,使用者根据实际场景重新录制元素即可 iTimeOut--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时/激活窗口/光标位置/横坐标偏移/纵坐标偏移/辅助按键/操作类型/平滑移动).Type:Dict 出参： 无 注意事项: 模拟操作：指通过调用系统api mouseevent等实现鼠标操作，会实际移动光标。 系统消息：指发送鼠标消息到目标元素，不移动光标。 后台操作：可以理解为调用了一次元素的鼠标响应回调函数。 *********************************************************************/ Mouse.Hover(@ui"输入控件<input>_百度一下7",10000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200, "bSetForeground": true, "sCursorPosition": "Center", "iCursorOffsetX": 10, "iCursorOffsetY": 10, "sKeyModifiers": [],"sSimulate": "simulate", "bMoveSmoothly": false})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Mouse_图片/Mouse_Hover.png)  

---

## 模拟移动

**说明**: 鼠标移动到指定坐标位置  

**原型**: `Mouse.Move(x, y, bStep,optionArgs)`  

**参数**:  
- **x** (True) [number] 默认:0 - 鼠标移动到指定位置的横坐标，以屏幕左上角为原点(0,0)  
- **y** (True) [number] 默认:0 - 鼠标移动到指定位置的纵坐标，以屏幕左上角为原点(0,0)  
- **bStep** (True) [boolean] 默认:False - 是否根据鼠标当前位置为原点进行坐标移动，默认为 false  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/******************************模拟移动******************************* 命令原型: Mouse.Move(x, y, bStep,optionArgs) 入参: x--横坐标 y--纵坐标 bStep--相对移动 optionArgs--可选参数(包括:执行后延时/执行前延时).Type:Dict 出参： 无 注意事项: 无 *********************************************************************/ Mouse.Move(1000, 500, false,{"iDelayAfter": 300, "iDelayBefore": 200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Mouse_图片/Mouse_Move.png)  

---

## 等待光标空闲

**说明**: 等待鼠标从繁忙状态切换到空闲状态  

**原型**: `Mouse.WaitCursorIdle(iTimeOut,optionArgs)`  

**参数**:  
- **iTimeOut** (True) [number] 默认:30000 - 循环检查鼠标的状态，若在指定时间内依旧不是空闲状态，则抛出超时异常。默认为30秒(30000毫秒)  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/****************************等待光标空闲********************************* 命令原型: Mouse.WaitCursorIdle(iTimeOut,optionArgs) 入参: iTimeOut--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参： 无 注意事项: 无 *********************************************************************/ Mouse.WaitCursorIdle(30000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Mouse_图片/Mouse_WaitCursorIdle.png)  

---

## 模拟滚轮

**说明**: 模拟鼠标的滚轮操作  

**原型**: `Mouse.Wheel(iCount,iDirection, sKeyModifiers,optionArgs)`  

**参数**:  
- **iCount** (True) [number] 默认:0 - 滚动的次数  
- **iDirection** (True) [enum] 默认:"down" - 滚轮滚动的方向  
- **sKeyModifiers** (True) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/******************************模拟滚轮******************************* 命令原型: Mouse.Wheel(iCount,iDirection, sKeyModifiers,optionArgs) 入参: iCount--滚动次数 iDirection--滚动方向(向下滚动/向上滚动) sKeyModifiers--辅助按键 optionArgs--可选参数(包括:执行后延时/执行前延时).Type:Dict 出参： 无 注意事项: 无 *********************************************************************/ Mouse.Wheel(2,"down", [],{"iDelayAfter":300, "iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Mouse_图片/Mouse_Wheel.png)  

---

## 实体抽取

**说明**: 将字序进行实体抽取，暂只支持系统（吾来平台预设）实体抽取。返回一个数组，包含实体在字序中的起始位置、实体标准值、实体名称、实体文本。调用时需要访问互联网，每台机器限制每分钟调用60次，超出限制将禁止调用10分钟  

**原型**: `arrEntity = NLP.Extract(strText)`  

**参数**:  
- **strText** (True) [string] 默认:"" - 需要抽取实体的字序，如“孙杨获得了奥运会金牌”  

**返回**: arrEntity，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************实体抽取*************************************** 命令原型： NLP.Extract("") 入参： strText -- 需要抽取实体的字序，如“孙杨获得了奥运会金牌” 出参： arrEntity -- 将命令运行后的结果赋值给此变量 注意事项： 调用时需要访问互联网，每台机器限制每分钟调用60次，超出限制将禁止调用10分钟 **********************************************************************************/ Dim arrEntity arrEntity = NLP.Extract("孙杨获得了奥运会金牌") TracePrint(arrEntity)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/NLP_图片/NLP_Extract.png)  

---

## 分词&词性标注

**说明**: 将字序分成单独的词，返回一个数组，包含每个词文本、词性、词在字序中的起始位置。调用时需要访问互联网，每台机器限制每分钟调用60次，超出限制将禁止调用10分钟  

**原型**: `arrParticiple = NLP.Tokenize(strText)`  

**参数**:  
- **strText** (True) [string] 默认:"" - 需要分词的字序，如“下雨天留客天天留我不留”  

**返回**: arrParticiple，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************分词&词性标注*************************************** 命令原型： NLP.Tokenize("") 入参： strText -- 需要分词的字序，如“下雨天留客天天留我不留” 出参： arrParticiple -- 将命令运行后的结果赋值给此变量 注意事项： 调用时需要访问互联网，每台机器限制每分钟调用60次，超出限制将禁止调用10分钟 **********************************************************************************/ Dim arrParticiple arrParticiple = NLP.Tokenize("下雨天留客天天留我不留") TracePrint(arrParticiple)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/NLP_图片/NLP_Tokenize.png)  

---

## 鼠标点击OCR文本

**说明**: 使用OCR对窗口范围内进行指定文字识别，如果识别到指定文字就点击它，调用时需要访问互联网  

**原型**: `OCR.Click(objUiElement,objRect,sEngine,sAcessKey,sSecretKey,sText,iRule,iOccurrence,iButton,iType,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **objRect** (True) [dictionary] 默认:{ "x":0,"y":0,"width":0,"height":0 } - 需要进行OCR文字识别的范围，程序会在控件这个范围内进行文字识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的文字识别  
- **sEngine** (True) [enum] 默认:"baidu2" - 使用的OCR引擎  
- **sAcessKey** (True) [string] 默认:"" - OCR服务的ApiKey  
- **sSecretKey** (True) [string] 默认:"" - OCR服务的SecretKey  
- **sText** (True) [string] 默认:"" - 查找元素时使用的文本  
- **iRule** (True) [enum] 默认:"instr" - 查找文本时使用的规则  
- **iOccurrence** (True) [number] 默认:1 - 如果“文本”字段中的字符串在指示的界面元素中出现多次，请在此处指定要单击的出现次数。例如，如果字符串出现4次并且您要单击第一个匹配项，请在此字段中写入1  
- **iButton** (True) [enum] 默认:"left" - 鼠标按键 { left:左键, right:右键, middle:中键 }  
- **iType** (True) [enum] 默认:"click" - 点击类型 { click:单击, dbclick:双击, down:按下, up:弹起 }  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sCursorPosition** (False) [enum] 默认:"Center" - 描述添加OffsetX和OffsetY属性的偏移量的光标起点。可以使用以下选项：TopLeft，TopRight，BottomLeft，BottomRight和Center。默认选项是Center  
- **iCursorOffsetX** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的水平位移  
- **iCursorOffsetY** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的垂直位移  
- **sKeyModifiers** (False) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **sSimulate** (False) [enum] 默认:"simulate" - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：模拟操作(simulate)  

**示例**:  
```
/*********************************鼠标点击OCR文本*************************************** 命令原型： OCR.Click(@ui"",{"x":0,"y":0,"width":0,"height":0},"baidu2","","","","instr",1,"left","click",10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate"}) 入参： objUiElement--识别目标。 objRect--目标识别范围。默认值:{"x":0,"y":0,"width":0,"height":0} sEngine--OCR引擎。注：使用的OCR引擎 sAcessKey--百度OCR引擎ApiKey。注：OCR服务的ApiKey（需要注册OCR账号，创建应用获取该参数） sSecretKey--百度OCR引擎SecretKey。注：OCR服务的SecretKey（需要注册OCR账号，创建应用获取该参数） sText--查找文本。注：查找元素时使用的文本 iRule--查找规则。注：查找文本时使用的规则 iOccurrence--相似结果位置。 iButton--鼠标点击。注：鼠标按键 {left:左键, right:右键, middle:中键} iType--点击类型。注：点击类型 {click:单击, dbclick:双击, down:按下, up:弹起} iTimeOut--超时时间。注：指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒） optionArgs--可选参数(包括:错误继续执行、执行后延时、执行前延时、激活窗口、光标位置、横坐标偏移、纵坐标偏移、辅助按键、操作类型).Type:Dict 注意事项： 1.该命令是调用外部百度OCR接口，需要提前创建好百度OCR账号以及对应应用，获取应用对应密钥（Access Key和Secret Key）。 2.要保证OCR操作的窗口存在，否则命令执行会报错。 3.运行该命令时需要能够连接外网。 ********************************************************************************/ OCR.Click(@ui"文本<span>_鼠标点击OCR文本",{"x":0,"y":0,"width":0,"height":0},"baidu2","P10qsjDcuoN5ZU7dnSHmXzGh","Fm8whYrNnXtKkSvqVPUYL573XfNbxlmO","文本","instr",1,"left","click",10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate"})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/OCR_图片/OCR_Click.png)  

---

## 删除邮件

**说明**: 删除指定邮件消息  

**原型**: `Outlook.DeleteMailMessage(message)`  

**参数**:  
- **message** (True) [expression] 默认:{ } - 邮件列表中的邮件对象  

**示例**:  
```
/*********************************删除邮件*************************************** 命令原型： Outlook.DeleteMailMessage({}) 入参： message--邮件列表中的邮件对象。 注意事项： 该命令需要配合“获取邮件列表”进行使用，根据返回对应进行操作。 **********************************************************************************/ Dim arrayRet = "" arrayRet = Outlook.GetMailMessages("lzz1712@outlook.com","收件箱","测试",False,True,0) Outlook.DeleteMailMessage(arrayRet[0]) TracePrint "获取收件箱列表中的第一封邮件"
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Outlook_图片/Outlook_DeleteMailMessage.png)  

---

## 获取所有图片

**说明**: 获取指定PDF文件中的所有图片，图片以“PDF文件名_序号”的命名方式保存  

**原型**: `PDF.GetAllPic(filePath,password,savePath,type)`  

**参数**:  
- **filePath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - PDF文件路径  
- **password** (True) [string] 默认:"" - PDF文件密码，无密码不需要填写  
- **savePath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 图片保存的路径，默认为PDF文件所在目录  
- **type** (True) [enum] 默认:"PNG" - 图片保存的格式，支持PNG、JPG、BMP格式  

**示例**:  
```
/**********************************获取所有图片*********************************** 命令原型： PDF.GetAllPic(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"",&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"PNG") 入参： filePath--PDF文件路径 password--PDF文件密码，无密码不需要填写 savePath--图片保存的路径，默认为PDF文件所在目录 imgType--图片保存的格式，支持PNG、JPG、BMP格式 出参： 无 ********************************************************************************/ Dim filePath=&#x27;&#x27;&#x27;C:\Users\Chance\Desktop\testPDF\来也科技RPA产品介绍.pdf&#x27;&#x27;&#x27; Dim password="" Dim savePath=&#x27;&#x27;&#x27;C:\Users\Chance\Desktop\testPDF&#x27;&#x27;&#x27; Dim imgType="PNG" PDF.GetAllPic(filePath,password,savePath,imgType)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/PDF_图片/PDF_GetAllPic.png)  

---

## 获取指定页图片

**说明**: 获取PDF文件中指定的页的图片。图片以“PDF文件名_序号”的命名方式保存  

**原型**: `PDF.GetPagePic(filePath,password,savePath,pageStart,pageEnd)`  

**参数**:  
- **filePath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - PDF文件路径  
- **password** (True) [string] 默认:"" - PDF文件密码，无密码不需要填写  
- **savePath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 图片保存的路径，默认为PDF文件所在目录  
- **pageStart** (True) [number] 默认:1 - 获取图片的开始页码，从1开始  
- **pageEnd** (True) [number] 默认:2 - 获取图片的结束页码  

**示例**:  
```
/*********************************获取指定页图片********************************** 命令原型： PDF.GetPagePic(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"",&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,1,2) 入参： filePath--PDF文件路径 password--PDF文件密码，无密码不需要填写 savePath--图片保存的路径，默认为PDF文件所在目录 pageStart--获取图片的开始页码，从1开始 pageEnd--获取图片的结束页码 出参： 无 ********************************************************************************/ Dim filePath=&#x27;&#x27;&#x27;C:\Users\Chance\Desktop\testPDF\来也科技RPA产品介绍.pdf&#x27;&#x27;&#x27; Dim savePath=&#x27;&#x27;&#x27;C:\Users\Chance\Desktop\testPDF&#x27;&#x27;&#x27; Dim password="" Dim pageStart=1 Dim pageEnd=2 PDF.GetPagePic(filePath,password,savePath, pageStart, pageEnd)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/PDF_图片/PDF_GetPagePic.png)  

---

## 获取指定页文本

**说明**: 获取PDF文件中指定的页的文本  

**原型**: `sRet = PDF.GetPageText(filePath,password,pageStart,pageEnd)`  

**参数**:  
- **filePath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - PDF文件路径  
- **password** (True) [string] 默认:"" - PDF文件密码，无密码不需要填写  
- **pageStart** (True) [number] 默认:1 - 获取文本的开始页码，从1开始  
- **pageEnd** (True) [number] 默认:2 - 获取文本的结束页码  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取指定页文本********************************** 命令原型： sRet = PDF.GetPageText(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"",1,2) 入参： filePath--PDF文件路径 password--PDF文件密码，无密码不需要填写 pageStart--获取文本的开始页码，从1开始 pageEnd--获取文本的结束页码 出参： sRet--命令运行后的结果 ********************************************************************************/ Dim filePath=&#x27;&#x27;&#x27;C:\Users\Chance\Desktop\testPDF\来也科技RPA产品介绍.pdf&#x27;&#x27;&#x27; Dim password="" Dim startPage=1 Dim endPage=2 sRet = PDF.GetPageText(filePath,password,startPage,endPage) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/PDF_图片/PDF_GetPageText.png)  

---

## 获取总页数

**说明**: 获取指定的PDF文件总页数  

**原型**: `iRet = PDF.PageCount(pdfName,password)`  

**参数**:  
- **pdfName** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - PDF文件路径  
- **password** (True) [string] 默认:"" - PDF文件密码，无密码不需要填写  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************************获取总页数************************************ 命令原型： iRet = PDF.PageCount(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"") 入参： filePath--PDF文件路径 password--PDF文件密码，无密码不需要填写 出参： iRet--命令运行后的结果 ********************************************************************************/ Dim filePath=&#x27;&#x27;&#x27;C:\Users\Chance\Desktop\testPDF\来也科技RPA产品介绍.pdf&#x27;&#x27;&#x27; Dim password="" iRet = PDF.PageCount(filePath,password) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/PDF_图片/PDF_PageCount.png)  

---

## 将指定页另存为图片

**说明**: 将PDF文件中指定的页另存为图片。图片以“PDF文件名_序号”的命名方式保存  

**原型**: `PDF.PageSaveToPic(filePath,password,savePath,pageStart,pageEnd)`  

**参数**:  
- **filePath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - PDF文件路径  
- **password** (True) [string] 默认:"" - PDF文件密码，无密码不需要填写  
- **savePath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 图片保存的路径，默认为PDF文件所在目录  
- **pageStart** (True) [number] 默认:1 - 需要另存为图片的开始页码，从1开始  
- **pageEnd** (True) [number] 默认:2 - 需要另存为图片的结束页码  

**示例**:  
```
/********************************将指定页另存为图片******************************* 命令原型： PDF.PageSaveToPic(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,"",&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;,1,2) 入参： filePath--PDF文件路径 password--PDF文件密码，无密码不需要填写 savePath--图片保存的路径，默认为PDF文件所在目录 pageStart--需要另存为图片的开始页码，从1开始 pageEnd--需要另存为图片的结束页码 出参： 无 ********************************************************************************/ Dim filePath=&#x27;&#x27;&#x27;C:\Users\Chance\Desktop\testPDF\来也科技RPA产品介绍.pdf&#x27;&#x27;&#x27; Dim password="" Dim savePath=&#x27;&#x27;&#x27;C:\Users\Chance\Desktop\testPDF&#x27;&#x27;&#x27; Dim startPage=1 Dim endPage=2 PDF.PageSaveToPic(filePath,password,savePath,startPage,endPage)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/PDF_图片/PDF_PageSaveToPic.png)  

---

## 合并PDF

**说明**: 将多个PDF文件合并成一个PDF文件  

**原型**: `PDF.PdfMerge(filePaths,savePath)`  

**参数**:  
- **filePaths** (True) [array] 默认:[&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;] - 需要合并的PDF文件，可以选择多个  
- **savePath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 合并后PDF文件保存的路径  

**示例**:  
```
/*************************************合并PDF************************************ 命令原型： PDF.PdfMerge([&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;],&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： filePaths--需要合并的PDF文件，可以选择多个 savePath--合并后PDF文件保存的路径 出参： 无 ********************************************************************************/ Dim filePaths=[&#x27;&#x27;&#x27;C:\Users\Chance\Desktop\testPDF\来也科技RPA产品介绍.pdf&#x27;&#x27;&#x27;,&#x27;&#x27;&#x27;C:\Users\Chance\Desktop\testPDF\来也科技Chatbot产品介绍.pdf&#x27;&#x27;&#x27;] Dim savePath=&#x27;&#x27;&#x27;C:\Users\Chance\Desktop\汇总.pdf&#x27;&#x27;&#x27; PDF.PdfMerge(filePaths,savePath)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/PDF_图片/PDF_PdfMerge.png)  

---

## 清除文字

**说明**: 清除显示在写屏窗口上的文字  

**原型**: `PrintToScreen.CleanText(objWindow)`  

**参数**:  
- **objWindow** (True) [expression] 默认:objWindow - 写屏窗口对象，“创建写屏对象”命令的输出  

**示例**:  
```
/*********************************清除文字*************************************** 命令原型： PrintToScreen.CleanText(objWindow) 入参： objWindow--写屏窗口对象。 注意事项： 1.要与创建写屏对象命令进行搭配使用。 ********************************************************************************/ PrintToScreen.CleanText(objWindow)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/PrintToScreen_图片/PrintToScreen_CleanText.png)  

---

## 关闭窗口

**说明**: 关闭正在显示的写屏窗口  

**原型**: `PrintToScreen.CloseWindow(objWindow)`  

**参数**:  
- **objWindow** (True) [expression] 默认:objWindow - 写屏窗口对象，“创建写屏对象”命令的输出  

**示例**:  
```
/*********************************关闭窗口*************************************** 命令原型： PrintToScreen.CloseWindow(objWindow) 入参： objWindow--写屏窗口对象。注：写屏窗口对象，创建写屏对象命令的输出 注意事项： 1.使用时要与创建写屏对象命令搭配使用。 ********************************************************************************/ PrintToScreen.CloseWindow(objWindow)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/PrintToScreen_图片/PrintToScreen_CloseWindow.png)  

---

## 创建写屏对象

**说明**: 创建一个写屏窗口对象，用于显示文字  

**原型**: `objWindow = PrintToScreen.CreateWindow(objRect,bResize)`  

**参数**:  
- **objRect** (True) [dictionary] 默认:{ "width": 0,"height": 0,"x": 0,"y": 0,"resolution": { "width": 0, "height": 0 } } - 创建一个写屏窗口对象，用于显示文字  
- **bResize** (True) [boolean] 默认:True - 是否自适应分辨率调整显示的位置和大小  

**返回**: objWindow，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************创建写屏对象*************************************** 命令原型： objWindow = PrintToScreen.CreateWindow({"width": 0,"height": 0,"x": 0,"y": 0,"resolution": {"width": 0, "height": 0}},true) 入参1： objRect--写屏区域。默认值:{"width": 0,"height": 0,"x": 0,"y": 0,"resolution": {"width": 0, "height": 0}} bResize--自适应。注：是否自适应分辨率调整显示的位置和大小 出参： objWindow--函数调用的输出保存到的变量。 注意事项： 1.使用时要注意写屏区域的大小以及电脑分辨率。 2.使用完成后使用关闭窗口关闭该写屏窗口。 ********************************************************************************/ Dim objWindow = "" objWindow = PrintToScreen.CreateWindow({"width": 1919,"height": 1079,"x": 0,"y": 0,"resolution": {"width":1920,"height":1080}},true) tracePrint(objWindow)
```  

---

## 绘制文字

**说明**: 绘制显示在写屏窗口上的文字  

**原型**: `PrintToScreen.DrawText(objWindow,strText,size,color)`  

**参数**:  
- **objWindow** (True) [expression] 默认:objWindow - 写屏窗口对象，创建写屏对象命令的输出  
- **strText** (True) [string] 默认:"" - 显示在写屏窗口上的文字，仅支持字符串和值为字符串的变量  
- **size** (True) [number] 默认:18 - 显示在写屏窗口上的文字大小，范围1~409数值越大文字越大  
- **color** (True) [expression] 默认:[255,0,0] - 显示在写屏窗口上的文字颜色，数组形式RGB颜色  

**示例**:  
```
/*********************************绘制文字*************************************** 命令原型： PrintToScreen.DrawText(objWindow,"",18,[255,0,0]) 入参： objWindow--写屏窗口对象。 strText--显示内容。注：显示在写屏窗口上的文字，仅支持字符串和值为字符串的变量 size--文字大小。注：显示在写屏窗口上的文字大小，范围1~409数值越大文字越大.默认值:18 color--文字颜色。注：显示在写屏窗口上的文字颜色，数组形式RGB颜色.默认值:[255,0,0] 注意事项： 1.写屏的文字长度不要超出写屏范围，否则不会显示。 2.要与创建写屏对象一同使用。 3.写屏时请勿随意中断流程，否则写屏文字会一直存在于桌面上，直到你彻底关闭UiBot。 ********************************************************************************/ PrintToScreen.DrawText(objWindow,"写屏文字",18,[255,0,0])
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/PrintToScreen_图片/PrintToScreen_DrawText.png)  

---

## 图像QR二维码识别

**说明**: 从指定图片中识别单个或多个QR二维码信息  

**原型**: `arrayText = QRCodeEx.ImageQRCode(sFileName)`  

**参数**:  
- **sFileName** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 要识别的图片路径，支持jpg、jpeg、bmp、png格式，图片大小不能超过 10M  

**返回**: arrayText，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************图像QR二维码识别*************************************** 命令原型： arrayText = QRCodeEx.ImageQRCode(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： sFileName--识别图片的路径。注：要识别的图片路径，支持jpg、jpeg、bmp、png格式，图片大小不能超过 10M 出参： arrayText--函数调用的输出保存到的变量。 注意事项： 1.保证本地二维码文件存在，以及文件格式和大小。 2.要保证有网络连接。 ********************************************************************************/ Dim arrayText = "" arrayText = QRCodeEx.ImageQRCode(@res"QR.png") TracePrint(arrayText)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/QRCodeEx_图片/QRCodeEx_ImageQRCode.png)  

---

## 屏幕QR二维码识别

**说明**: 从指定屏幕范围内识别单个或多个QR二维码信息  

**原型**: `arrayText = QRCodeEx.ScreenQRCode(objElement,objRect,iTimeOut)`  

**参数**:  
- **objElement** (True) [expression] 默认:{ } - 通过鼠标选取或截取需要识别的目标屏幕范围。包含窗口、元素、范围等信息  
- **objRect** (True) [dictionary] 默认:{ "x":0,"y":0,"width":0,"height":0 } - 需要查找的范围，程序会在控件这个范围内进行识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的识别  
- **iTimeOut** (True) [number] 默认:30000 - 指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒）  

**返回**: arrayText，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************屏幕QR二维码识别*************************************** 命令原型： arrayText = QRCodeEx.ScreenQRCode({},{"x":0,"y":0,"width":0,"height":0},30000) 入参： objElement--识别目标。 objRect--识别范围。默认值:{"x":0,"y":0,"width":0,"height":0} iTimeOut--超时时间（毫秒）。注：指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发异常。默认30000毫秒（30秒） 出参： arrayText--函数调用的输出保存到的变量。 注意事项： 要保证有网络连接。 ********************************************************************************/ Dim arrayText = "" #icon("@res:b5c39a70-7f27-11ec-ab08-1d7b5faf0fee.png") arrayText = QRCodeEx.ScreenQRCode({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"百度一下，你就知道 - Google Chrome","app":"chrome"}]},{"x":0,"y":0,"width":0,"height":0},30000) TracePrint(arrayText)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/QRCodeEx_图片/QRCodeEx_ScreenQRCode.png)  

---

## 屏幕锁屏

**说明**: 锁住系统屏幕，确保后续代码能在锁屏状态下正常运行  

**原型**: `bRet = RDP.LockScreen(username,password,optionArgs)`  

**参数**:  
- **username** (True) [string] 默认:"" - 登录系统的用户名或微软帐户  
- **password** (True) [string] 默认:"" - 登录系统的密码（微软帐户需要输入微软帐户的密码）  
- **width** (False) [number] 默认:0 - 设定锁屏分辨率宽度，参数为0则自动适配宽度  
- **height** (False) [number] 默认:0 - 设定锁屏分辨率高度，参数为0则自动适配高度  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************屏幕锁屏*************************************** 命令原型： bRet = RDP.LockScreen("","",{"width":0,"height":0}) 入参： username--用户或账户。注：登录系统的用户名或微软帐户 password--密码。注：登录系统的密码（微软帐户需要输入微软帐户的密码） optionArgs--可选参数(包括:分辨率宽度、分辨率高度).Type:Dict 出参： bRet--函数调用的输出保存到的变量。 注意事项： 1.该命令一定要放在代码执行的第一步，以此保证后续代码的执行。 2.解锁的windows系统需要是专业版，打开远程桌面选择允许远程连接到这台计算机 ********************************************************************************/ Dim bRet = "" bRet = RDP.LockScreen("Administrator","laiye666",{"width":0,"height":0}) TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/RDP_图片/RDP_LockScreen.png)  

---

## 屏幕解锁

**说明**: 解锁系统屏幕，进入系统桌面  

**原型**: `bRet = RDP.UnlockScreen(username,password)`  

**参数**:  
- **username** (True) [string] 默认:"" - 登录系统的用户名或微软帐户  
- **password** (True) [string] 默认:"" - 登录系统的密码（微软帐户需要输入微软帐户的密码）  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************屏幕解锁*************************************** 命令原型： bRet = RDP.UnlockScreen("","") 入参： username--用户或账户。注：登录系统的用户名或微软帐户 password--密码。注：登录系统的密码（微软帐户需要输入微软帐户的密码） 出参： bRet--函数调用的输出保存到的变量。 注意事项： 1.该命令可以单独执行，也可以搭配锁屏命令执行。 2.解锁的windows系统需要是专业版，打开远程桌面选择允许远程连接到这台计算机 ********************************************************************************/ Dim bRet = "" bRet = RDP.UnlockScreen("Administrator","laiye666") TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/RDP_图片/RDP_UnlockScreen.png)  

---

## 正则表达式查找

**说明**: 正则表达式查找字符串，返回找到的字符串数组  

**原型**: `arrRet = Regex.Find(sText,sPattern)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 进行操作的字符串  
- **sPattern** (True) [string] 默认:"" - 正则表达式  

**返回**: arrRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************正则表达式查找************************ 命令原型: arrRet=Regex.Find(sText,sPattern) 入参： sText--进行操作的字符串。 sPattern--正则表达式。 出参： arrRet--函数调用的输出保存到的变量。 注意事项: 正则表达式查找字符串，返回找到的字符串数组。 ***********************************************************/ arrRet=Regex.Find("UiBot2022","\\d{4}")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Regex_图片/Regex_Find.png)  

---

## 正则表达式查找全部

**说明**: 正则表达式查找全部字符串，返回查找到的字符串数组  

**原型**: `arrRet = Regex.FindAll(sText,sPattern)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 进行操作的字符串  
- **sPattern** (True) [string] 默认:"" - 正则表达式  

**返回**: arrRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************正则表达式查找全部************************ 命令原型: arrRet=Regex.FindAll(sText,sPattern) 入参： sText--进行操作的字符串。 sPattern--正则表达式。 出参： arrRet--函数调用的输出保存到的变量。 注意事项: 正则表达式查找全部字符串，返回找到的字符串数组。 ***********************************************************/ arrRet=Regex.FindAll("UiBot2022RPA123","[0-9]")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Regex_图片/Regex_FindAll.png)  

---

## 正则表达式查找子串

**说明**: 正则表达式查找字符串，返回找到的字符串子串  

**原型**: `sRet = Regex.FindStr(sText,sPattern,iGroup)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 进行操作的字符串  
- **sPattern** (True) [string] 默认:"" - 正则表达式  
- **iGroup** (True) [number] 默认:0 - 返回第几个子表达式的匹配结果，0表示返回匹配的整个字符串  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************正则表达式查找子串************************ 命令原型: sRet=Regex.FindStr(sText,sPattern,iGroup) 入参： sText--进行操作的字符串。 sPattern--正则表达式。 iGroup--返回第几个子表达式的匹配结果，0表示返回匹配的整个字符串。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 正则表达式查找字符串，返回找到的字符串子串。 ***********************************************************/ sRet=Regex.FindStr("UiBot123","[0-9]{3}",0)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Regex_图片/Regex_FindStr.png)  

---

## 正则表达式替换

**说明**: 正则表达式替换字符串，返回替换后的字符串结果  

**原型**: `sRet = Regex.Replace(sText,sPattern,sNew,nCount)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 需要被替换的字符串  
- **sPattern** (True) [string] 默认:"" - 正则表达式  
- **sNew** (True) [string] 默认:"" - 目标字符串中被正则表达式匹配的内容会被替换为此文本  
- **nCount** (True) [number] 默认:0 - 模式匹配后替换的最大次数，默认 0 表示替换所有的匹配  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************正则表达式替换************************ 命令原型: sRet=Regex.Replace(sText,sPattern,sNew,nCount) 入参： sText--进行操作的字符串。 sPattern--正则表达式。 sNew--目标字符串中被正则表达式匹配的内容会被替换为此文本。 nCount--模式匹配后替换的最大次数，默认 0 表示替换所有的匹配。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 正则表达式替换字符串，返回替换后的字符串结果 ***********************************************************/ sRet=Regex.Replace("UiBot123","\\d{3}","RPA",0)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Regex_图片/Regex_Replace.png)  

---

## 设置匹配换行

**说明**: 设置正则表达式.匹配换行符  

**原型**: `Regex.SetDotAll(bVal)`  

**参数**:  
- **bVal** (True) [boolean] 默认:True - 是否生效  

**示例**:  
```
/************************设置匹配换行************************ 命令原型: Regex.SetDotAll(bVal) 入参： bVal--是否生效。 出参： 无 注意事项: 设置正则表达式匹配换行符。 ***********************************************************/ Regex.SetDotAll(true)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Regex_图片/Regex_SetDotAll.png)  

---

## 设置忽略大小写

**说明**: 设置正则表达式忽略大小写  

**原型**: `Regex.SetIgnoreCase(bVal)`  

**参数**:  
- **bVal** (True) [boolean] 默认:True - 是否忽略  

**示例**:  
```
/************************设置忽略大小写************************ 命令原型: Regex.SetIgnoreCase(bVal) 入参： bVal--是否忽略。 出参： 无 注意事项: 设置正则表达式忽略大小写 ***********************************************************/ Regex.SetIgnoreCase(true)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Regex_图片/Regex_SetIgnoreCase.png)  

---

## 设置本地化识别

**说明**: 设置正则表达式本地化识别  

**原型**: `Regex.SetLocale(bVal)`  

**参数**:  
- **bVal** (True) [boolean] 默认:True - 是否生效  

**示例**:  
```
/************************设置本地化识别************************ 命令原型: Regex.SetLocale(bVal) 入参： bVal--是否生效。 出参： 无 注意事项: 设置正则表达式本地化识别 ***********************************************************/ Regex.SetLocale(true)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Regex_图片/Regex_SetLocale.png)  

---

## 设置多行匹配

**说明**: 设置正则表达式多行匹配  

**原型**: `Regex.SetMutexLine(bVal)`  

**参数**:  
- **bVal** (True) [boolean] 默认:True - 是否生效  

**示例**:  
```
/************************设置多行匹配************************ 命令原型: Regex.SetMutexLine(bVal) 入参： bVal--是否生效。 出参： 无 注意事项: 设置正则表达式本地化识别 ***********************************************************/ Regex.SetMutexLine(bVal)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Regex_图片/Regex_SetMutexLine.png)  

---

## 设置使用Unicode字符集

**说明**: 设置正则表达式使用Unicode字符集  

**原型**: `Regex.SetUnicode(bVal)`  

**参数**:  
- **bVal** (True) [boolean] 默认:True - 是否生效  

**示例**:  
```
/************************设置使用Unicode字符集************************ 命令原型: Regex.SetUnicode(bVal) 入参： bVal--是否生效。 出参： 无 注意事项: 设置正则表达式使用Unicode字符集。 ***********************************************************/ Regex.SetUnicode(bVal)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Regex_图片/Regex_SetUnicode.png)  

---

## 设置支持更灵活的格式

**说明**: 设置正则表达式支持更灵活的格式  

**原型**: `Regex.SetVerbose(bVal)`  

**参数**:  
- **bVal** (True) [boolean] 默认:True - 是否生效  

**示例**:  
```
/************************设置支持更灵活的格式************************ 命令原型: Regex.SetVerbose(bVal) 入参： bVal--是否生效。 出参： 无 注意事项: 设置正则表达式支持更灵活的格式。 ***********************************************************/ Regex.SetVerbose(bVal)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Regex_图片/Regex_SetVerbose.png)  

---

## 正则表达式查找测试

**说明**: 尝试使用正则表达式查找字符串，能够找到返回 true，找不到返回 false  

**原型**: `bRet = Regex.Test(sText,sPattern)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 进行操作的字符串  
- **sPattern** (True) [string] 默认:"" - 正则表达式  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************正则表达式查找测试************************ 命令原型: bRet=Regex.Test(sText,sPattern) 入参： sText--进行操作的字符串。 sPattern--正则表达式。 出参： bRet--函数调用的输出保存到的变量。 注意事项: 尝试使用正则表达式查找字符串，能够找到返回 true，找不到返回 false。 ***********************************************************/ bRet=Regex.Test("UiBot123","\\d{3}")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Regex_图片/Regex_Test.png)  

---

## 添加元素到集合

**说明**: 添加元素到集合  

**原型**: `Set.Add(ObjSet,varValue)`  

**参数**:  
- **ObjSet** (True) [expression] 默认:ObjSet - 要操作的集合对象  
- **varValue** (True) [expression] 默认:varValue - 要添加到集合中的元素，可以是任意类型数据  

**示例**:  
```
/*********************************添加元素到集合*************************************** 命令原型： Set.Add(ObjSet,varValue) 入参： ObjSet -- 要操作的集合对象 varValue -- 要添加到集合中的元素，可以是任意类型数据 注意事项： 该命令不能单独使用，需配合 "创建集合"命令(Set.Create())一起使用才能正常使用，单独使用则会报错 **********************************************************************************/ Dim ObjSet ObjSet=Set.Create() Set.Add(ObjSet,4) TracePrint(ObjSet)
```  

---

## 获取集合大小

**说明**: 获取集合内元素个数  

**原型**: `iRet = Set.Count(objSet)`  

**参数**:  
- **objSet** (True) [expression] 默认:objSet - 由“创建集合”命令返回的对象，是一个无序不重复元素集  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取集合大小*************************************** 命令原型： Set.Count(objSet) 入参： ObjSet -- 由“创建集合”命令返回的对象，是一个无序不重复元素集 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 该命令不能单独使用，需配合 "创建集合"命令(Set.Create())一起使用才能正常使用，单独使用则会报错 **********************************************************************************/ Dim iRet Dim ObjSet ObjSet = Set.Create() iRet = Set.Count(ObjSet) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Set_图片/Set_Count.png)  

---

## 创建集合

**说明**: 创建一个集合并保存到一个变量当中  

**原型**: `ObjSet=Set.Create()`  

**参数**:  

**返回**: ObjSet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************创建集合*************************************** 命令原型： Set.Create() 出参： ObjSet -- 将命令运行后的结果赋值给此变量 注意事项： 此命令无参数 **********************************************************************************/ Dim ObjSet ObjSet = Set.Create() TracePrint(ObjSet)
```  

---

## 取交集

**说明**: 取交集，将交集返回为新的集合  

**原型**: `objSetRet = Set.Intersection(ObjSet,ObjSet1)`  

**参数**:  
- **ObjSet** (True) [expression] 默认:ObjSet - 要操作的集合对象  
- **ObjSet1** (True) [expression] 默认:ObjSet1 - 要对比交集的集合  

**返回**: objSetRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取交集*************************************** 命令原型： Set.Intersection(ObjSet,ObjSet1) 入参： ObjSet -- 要操作的集合对象 ObjSet1 -- 要对比交集的集合 出参： objSetRet -- 将命令运行后的结果赋值给此变量 注意事项： 该命令不能单独使用，需配合 "创建集合"命令(Set.Create())、 "添加元素到集合"命令(Set.add())一起使用才能正常使用，单独使用则会报错 **********************************************************************************/ Dim ObjSet Dim ObjSet1 Dim objSetRet ObjSet = Set.Create() Set.Add(ObjSet,1) Set.Add(ObjSet,2) Set.Add(ObjSet,3) ObjSet1 = Set.Create() Set.Add(ObjSet1,2) Set.Add(ObjSet1,3) Set.Add(ObjSet1,4) objSetRet = Set.Intersection(ObjSet,ObjSet1) TracePrint(objSetRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Set_图片/Set_Intersection.png)  

---

## 判断是否有交集

**说明**: 判断是否有交集，有交集返回 true，无交集返回 false  

**原型**: `bRet = Set.IsDisjoint(ObjSet,ObjSet1)`  

**参数**:  
- **ObjSet** (True) [expression] 默认:ObjSet - 要操作的集合对象  
- **ObjSet1** (True) [expression] 默认:ObjSet1 - 要对比交集的集合  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************判断是否有交集*************************************** 命令原型： Set.IsDisjoint(ObjSet,ObjSet1) 入参： ObjSet -- 要操作的集合对象 ObjSet1 -- 要对比交集的集合 出参： bRet -- 将命令运行后的结果赋值给此变量 注意事项： 该命令不能单独使用，需配合 "创建集合"命令(Set.Create())、 "添加元素到集合"命令(Set.add())一起使用才能正常使用，单独使用则会报错 **********************************************************************************/ Dim ObjSet Dim ObjSet1 Dim bRet ObjSet = Set.Create() Set.Add(ObjSet,1) Set.Add(ObjSet,2) Set.Add(ObjSet,3) ObjSet1 = Set.Create() Set.Add(ObjSet1,2) Set.Add(ObjSet1,3) Set.Add(ObjSet1,4) bRet = Set.IsDisjoint(ObjSet,ObjSet1) TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Set_图片/Set_IsDisjoint.png)  

---

## 判断是否为子集

**说明**: 判断是否为子集，是子集返回 true，非子集返回 false  

**原型**: `bRet = Set.IsSubSet(ObjSet,ObjSet1)`  

**参数**:  
- **ObjSet** (True) [expression] 默认:ObjSet - 要操作的集合对象  
- **ObjSet1** (True) [expression] 默认:ObjSet1 - 要对比交集的集合  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************判断是否为子集*************************************** 命令原型： Set.IsSubSet(ObjSet,ObjSet1) 入参： ObjSet -- 要操作的集合对象 ObjSet1 -- 要对比交集的集合 出参： bRet -- 将命令运行后的结果赋值给此变量 注意事项： 该命令不能单独使用，需配合 "创建集合"命令(Set.Create())、 "添加元素到集合"命令(Set.add())一起使用才能正常使用，单独使用则会报错 **********************************************************************************/ Dim ObjSet Dim ObjSet1 Dim bRet ObjSet = Set.Create() Set.Add(ObjSet,1) Set.Add(ObjSet,2) ObjSet1 = Set.Create() Set.Add(ObjSet1,1) Set.Add(ObjSet1,2) Set.Add(ObjSet1,3) Set.Add(ObjSet1,4) bRet = Set.IsSubSet(ObjSet,ObjSet1) TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Set_图片/Set_IsSubSet.png)  

---

## 判断是否为父集

**说明**: 判断是否为父集，是父集返回 true，非父集返回 false  

**原型**: `bRet = Set.IsSuperSet(ObjSet,ObjSet1)`  

**参数**:  
- **ObjSet** (True) [expression] 默认:ObjSet - 要操作的集合对象  
- **ObjSet1** (True) [expression] 默认:ObjSet1 - 要对比交集的集合  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************判断是否为父集*************************************** 命令原型： Set.IsSuperSet(ObjSet,ObjSet1) 入参： ObjSet -- 要操作的集合对象 ObjSet1 -- 要对比交集的集合 出参： bRet -- 将命令运行后的结果赋值给此变量 注意事项： 该命令不能单独使用，需配合 "创建集合"命令(Set.Create())、 "添加元素到集合"命令(Set.add())一起使用才能正常使用，单独使用则会报错 **********************************************************************************/ Dim ObjSet Dim ObjSet1 Dim bRet ObjSet = Set.Create() Set.Add(ObjSet,1) Set.Add(ObjSet,2) ObjSet1 = Set.Create() Set.Add(ObjSet1,1) bRet = Set.IsSuperSet(ObjSet,ObjSet1) TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Set_图片/Set_IsSuperSet.png)  

---

## 删除元素

**说明**: 从集合中删除元素  

**原型**: `Set.Remove(ObjSet,varValue)`  

**参数**:  
- **ObjSet** (True) [expression] 默认:ObjSet - 要操作的集合对象  
- **varValue** (True) [expression] 默认:varValue - 要从集合中删除的元素，可以是任意类型数据  

**示例**:  
```
/*********************************删除元素*************************************** 命令原型： Set.Remove(ObjSet,varValue) 入参： ObjSet -- 要操作的集合对象 varValue -- 要从集合中删除的元素，可以是任意类型数据 注意事项： 该命令不能单独使用，需配合 "创建集合"命令(Set.Create())、 "添加元素到集合"命令(Set.add())一起使用才能正常使用，单独使用则会报错 **********************************************************************************/ Dim ObjSet ObjSet = Set.Create() Set.Add(ObjSet,1) Set.Add(ObjSet,2) Set.Add(ObjSet,3) Set.Remove(ObjSet,2) TracePrint(ObjSet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Set_图片/Set_Remove.png)  

---

## 取差集

**说明**: 合并集合中不同的元素，将合并结果返回为新的集合  

**原型**: `objSetRet = Set.Symmetric_Difference(ObjSet,ObjSet1)`  

**参数**:  
- **ObjSet** (True) [expression] 默认:ObjSet - 要操作的集合对象  
- **ObjSet1** (True) [expression] 默认:ObjSet1 - 要对比交集的集合  

**返回**: objSetRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取差集*************************************** 命令原型： Set.Symmetric_Difference(ObjSet,ObjSet1) 入参： ObjSet -- 要操作的集合对象 ObjSet1 -- 要对比交集的集合 出参： objSetRet -- 将命令运行后的结果赋值给此变量 注意事项： 该命令不能单独使用，需配合 "创建集合"命令(Set.Create())、 "添加元素到集合"命令(Set.add())一起使用才能正常使用，单独使用则会报错 **********************************************************************************/ Dim ObjSet Dim ObjSet1 Dim objSetRet ObjSet = Set.Create() Set.Add(ObjSet,1) Set.Add(ObjSet,2) Set.Add(ObjSet,3) ObjSet1 = Set.Create() Set.Add(ObjSet1,1) Set.Add(ObjSet1,2) Set.Add(ObjSet1,3) Set.Add(ObjSet1,4) objSetRet = Set.Symmetric_Difference(ObjSet,ObjSet1) TracePrint(objSetRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Set_图片/Set_Symmetric_Difference.png)  

---

## 转为数组

**说明**: 将集合转换为数组，一般用来遍历集合  

**原型**: `arrSet = Set.ToArray(objSet)`  

**参数**:  
- **objSet** (True) [expression] 默认:objSet - 需要转换的集合  

**返回**: arrSet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************转为数组*************************************** 命令原型： Set.ToArray(objSet) 入参： ObjSet -- 需要转换的集合 出参： arrSet -- 将命令运行后的结果赋值给此变量 注意事项： 该命令不能单独使用，需配合 "创建集合"命令(Set.Create())、 "添加元素到集合"命令(Set.add())一起使用才能正常使用，单独使用则会报错 **********************************************************************************/ Dim ObjSet Dim arrSet ObjSet = Set.Create() Set.Add(ObjSet,1) Set.Add(ObjSet,2) Set.Add(ObjSet,3) arrSet = Set.ToArray(objSet) TracePrint(arrSet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Set_图片/Set_ToArray.png)  

---

## 取并集

**说明**: 取并集，将并集返回为新的集合  

**原型**: `objSetRet = Set.Union(ObjSet,ObjSet1)`  

**参数**:  
- **ObjSet** (True) [expression] 默认:ObjSet - 要操作的集合对象  
- **ObjSet1** (True) [expression] 默认:ObjSet1 - 要对比交集的集合  

**返回**: objSetRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************取并集*************************************** 命令原型： Set.Union(ObjSet,ObjSet1) 入参： ObjSet -- 要操作的集合对象 ObjSet1 -- 要对比交集的集合 出参： arrSet -- 将命令运行后的结果赋值给此变量 注意事项： 该命令不能单独使用，需配合 "创建集合"命令(Set.Create())、 "添加元素到集合"命令(Set.add())一起使用才能正常使用，单独使用则会报错 **********************************************************************************/ Dim ObjSet Dim ObjSet1 Dim objSetRet ObjSet = Set.Create() Set.Add(ObjSet,1) Set.Add(ObjSet,3) ObjSet1 = Set.Create() Set.Add(ObjSet1,2) Set.Add(ObjSet1,4) objSetRet = Set.Union(ObjSet,ObjSet1) TracePrint(objSetRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Set_图片/Set_Union.png)  

---

## 创建字符串

**说明**: 创建一定数量的字符  

**原型**: `sRet = String(iCount,sChar)`  

**参数**:  
- **iCount** (True) [number] 默认:1 - 重复次数  
- **sChar** (True) [string] 默认:" " - 创建的字符  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************创建字符串***************** 命令原型: sRet=String(iCount,sChar) 入参： iCount--重复次数。 sChar--创建的字符。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=String(3,"RPA") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/String.png)  

---

## 取ASCII代码

**说明**: 获取指定字符的 ASCII 代码  

**原型**: `iRet = Asc(sChr)`  

**参数**:  
- **sChr** (True) [string] 默认:"A" - 要获取 ASCII 代码的字符  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************获取指定字符的ASCII代码***************** 命令原型: iRet=Asc(sChr) 入参： sChr--要获取指定字符的ASCII码。 出参： iRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ iRet=Asc("A") TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/Asc.png)  

---

## 取ASCII字符

**说明**: 获取指定 ASCII 代码对应的字符  

**原型**: `sRet = Chr(iAsc)`  

**参数**:  
- **iAsc** (True) [number] 默认:65 - 要转换为字符的 ASCII 代码  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************获取指定ASCII代码对应的字符***************** 命令原型: sRet=Chr(iAsc) 入参： iAsc--要获取指定ASCII代码的字符。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=Chr(65) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/Chr.png)  

---

## 抽取字符串中数字

**说明**: 抽取目标字符串中的所有数字  

**原型**: `sRet = DigitFromStr(sText)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 被抽取的源字符串  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************抽取目标字符串中的所有数字***************** 命令原型: sRet=DigitFromStr(sText) 入参： sText--被抽取的源字符串。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=DigitFromStr("ABC123") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/DigitFromStr.png)  

---

## 判断以指定后缀结尾

**说明**: 判断目标字符串是否以指定后缀结尾  

**原型**: `bRet = EndsWith(sText,sEndsStr)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 要判断的目标字符串  
- **sEndsStr** (True) [string] 默认:"" - 指定的后缀字符串  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************判断以指定后缀结尾***************** 命令原型: bRet=EndsWith(sText,sEndsStr) 入参： sText--要判断的目标字符串。 sEndsStr--指定的后缀字符串。 出参： bRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ bRet=EndsWith("UiBot","Bot") TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/EndsWith.png)  

---

## 格式化字符串

**说明**: 支持以占位符形式格式化字符串  

**原型**: `sRet = Format(sText,repText1,repText2)`  

**参数**:  
- **sText** (True) [string] 默认:"%d %s" - 包含占位符的字符串，遵循C标准库命令sprintf的规则：% [flags][width] [.precision][length] specifier，但对符号 *、h、 L、l、n、p 不支持；新增"%q"，对目标字符串两边加上双引号；"%%"代表转义自身，不具备占位符能力  
- **repText1** (True) [number] 默认:1 - 仅可填写一个替换值，且前后顺序直接对应格式字符串中占位符的前后顺序  
- **repText2** (True) [string] 默认:"Laiye RPA" - 仅可填写一个替换值，且前后顺序直接对应格式字符串中占位符的前后顺序  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************格式化字符串***************** 命令原型: sRet=Format(sText,repText1,repText2) 入参： sText--包含占位符的字符串。 repText1--对应格式字符串中占位符的前后顺序的一个替换值。 repText2--对应格式字符串中占位符的前后顺序的一个替换值。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet = Format("UiBot%d%s",1,"Laiye RPA") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/Format.png)  

---

## 查找字符串

**示例**:  
```
/************************查找字符串***************** 命令原型: iRet=InStr(sText,sSubText,iPos,bCompare) 入参： sText--进行操作的字符串。 sSubText--需要查找的子串。 iPos--从第几个字开始查找。 bCompare--对比字符串时是否区分大小写。 出参： iRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ iRet=InStr("Laiye RPA","RPA",1,True) TracePrint(iRet)
```  

---

## 倒序查找字符串

**说明**: 在字符串内查找指定的字符，返回查找到的字符的位置，如果没有找到，返回 0，倒序查找  

**原型**: `iRet = InStrRev(sText,sSubText,iPos,bCompare)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 进行操作的字符串  
- **sSubText** (True) [string] 默认:"" - 需要查找查找的子串  
- **iPos** (True) [number] 默认:1 - 从第几个字开始查找  
- **bCompare** (True) [boolean] 默认:False - 对比字符串时是否区分大小写  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************倒序查找字符串***************** 命令原型: iRet=InStrRev(sText,sSubText,iPos,bCompare) 入参： sText--进行操作的字符串。 sSubText--需要查找的子串。 iPos--从第几个字开始查找。 bCompare--对比字符串时是否区分大小写。 出参： iRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ iRet= InStrRev("UiBot","Bot",1,True) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/InStrRev.png)  

---

## 判断字符串

**说明**: 判断字符串是否全部为指定内容（选择英文字母、数字、大写、小写其中之一)  

**原型**: `bRet = IsSpecificStr(sText,sType)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 要判断的目标字符串  
- **sType** (True) [enum] 默认:"letter" - 判断类型，选择英文字母、数字、大写、小写其中之一  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************判断字符串***************** 命令原型: bRet=IsSpecificStr(sText,sType) 入参： sText--要判断的目标字符串。 sType--判断类型，选择英文字母、数字、大写、小写其中之一。 出参： bRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ bRet=IsSpecificStr("UiBot","letter") TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/IsSpecificStr.png)  

---

## 将字符串转换为小写

**说明**: 将字符串转换为小写  

**原型**: `sRet = LCase(sText)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 进行操作的字符串  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************将字符串转换为小写***************** 命令原型: sRet=LCase(sText) 入参： sText--进行操作的字符串。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=LCase("UiBot") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/LCase.png)  

---

## 获取左侧字符串

**说明**: 获取左侧字符串  

**原型**: `sRet = Left(sText,iSize)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 进行操作的字符串  
- **iSize** (True) [number] 默认:1 - 要获取多少个单位的字符串  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************获取左侧字符串***************** 命令原型: sRet=Left(sText,iSize) 入参： sText--进行操作的字符串。 iSize--要获取多少个单位的字符串。 出参： iRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=Left("UiBot",2) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/Left.png)  

---

## 获取字符串字节长度

**说明**: 获取字符串字节长度（实际占据空间的长度）  

**原型**: `iRet = LenB(sText)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 进行操作的字符串  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************获取字符串字节长度***************** 命令原型: iRet=LenB(sText) 入参： sText--进行操作的字符串。 出参： iRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ iRet=LenB("UiBot") TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/LenB.png)  

---

## 抽取字符串中字母

**说明**: 抽取目标字符串中的所有英文字母  

**原型**: `sRet = LetterFromStr(sText)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 被抽取的源字符串  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************抽取字符串中的字母***************** 命令原型: sRet = LetterFromStr(sText) 入参： sText--进行操作的字符串。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=LetterFromStr("Ab12C3") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/LetterFromStr.png)  

---

## 左侧裁剪

**说明**: 删除掉指定字符串左侧的特定字符  

**原型**: `sRet = LTrim(sData,sTrim)`  

**参数**:  
- **sData** (True) [string] 默认:"" - 被裁剪的字符串  
- **sTrim** (True) [string] 默认:"" - 输入单个字符或由多个字符组成的字符串，字符组成的所有组合都会被裁剪  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************左侧裁剪***************** 命令原型: sRet=LTrim(sData,sTrim) 入参： sData--被裁剪的字符串。 sTrim--要裁剪掉的字符串。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet = LTrim("UiBot","Ui") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/LTrim.png)  

---

## 获取中间字符串

**说明**: 获取中间字符串  

**原型**: `sRet = Mid(sText,iPos,iSize)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 进行操作的字符串  
- **iPos** (True) [number] 默认:1 - 从第N个字符开始截取，N >= 1 ，且为正整数，当 (N = 1) 时代表从第1个字符开始截取，若超出目标字符串的长度，则返回空字符串  
- **iSize** (True) [number] 默认:1 - 指定截取N个字符，N >= 1 ，且为正整数，若超出目标字符串的长度，则返回目标字符串本身  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************获取中间字符串***************** 命令原型: sRet=Mid(sText,iPos,iSize) 入参： sText--进行操作的字符串。 iPos--从第N个字符开始截取，N>=1，且为正整数，当 (N = 1) 时代表从第1个字符开始截取。 iSize--指定截取N个字符，N>=1，且为正整数。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=Mid("UiBot",3,1) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/Mid.png)  

---

## 替换字符串

**说明**: 对字符串执行查找替换操作，返回替换后的字符串  

**原型**: `sRet = Replace(sText,sSubText,sReplaceText,bCompare)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 进行操作的字符串  
- **sSubText** (True) [string] 默认:"" - 需要查找替换的子串  
- **sReplaceText** (True) [string] 默认:"" - 替换子串的字符串  
- **bCompare** (True) [boolean] 默认:False - 对比字符串时是否区分大小写  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************替换字符串***************** 命令原型: sRet=Mid(sText,iPos,iSize) 入参： sText--进行操作的字符串。 sSubText--需要查找替换的子串。 sReplaceText--替换子串的字符串。 bCompare--对比字符串时是否区分大小写。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=Replace("UsDot","sD","iB",true) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/Replace.png)  

---

## 获取右侧字符串

**说明**: 获取右侧字符串  

**原型**: `sRet = Right(sText,iSize)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 进行操作的字符串  
- **iSize** (True) [number] 默认:1 - 要获取多少个单位的字符串  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************获取右侧字符串***************** 命令原型: sRet=Right(sText,iSize) 入参： sText--进行操作的字符串。 iSize--要获取多少个单位的字符串。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=Right("UiBot",3) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/Right.png)  

---

## 右侧裁剪

**说明**: 删除掉指定字符串右侧的特定字符  

**原型**: `sRet = RTrim(sData,sTrim)`  

**参数**:  
- **sData** (True) [string] 默认:"" - 被裁剪的字符串  
- **sTrim** (True) [string] 默认:"" - 输入单个字符或由多个字符组成的字符串，字符组成的所有组合都会被裁剪  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************右侧裁剪***************** 命令原型: sRet=RTrim(sData,sTrim) 入参： sData--被裁剪的字符串。 sTrim--要裁剪掉的单个字符或由多个字符组成的字符串。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=RTrim("UiBot123a","123a") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/RTrim.png)  

---

## 创建空格

**说明**: 创建一定数量的空格字符  

**原型**: `sRet = Space(iCount)`  

**参数**:  
- **iCount** (True) [number] 默认:1 - 创建的空格字符数量  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************创建空格***************** 命令原型: sRet=Space(iCount) 入参： iCount--创建的空格字符数量。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=Space(3) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/Space.png)  

---

## 分割字符串

**说明**: 使用特定分隔符将字符串分割为数组  

**原型**: `arrRet = Split(sText,sSeparator)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 要处理的字符串  
- **sSeparator** (True) [string] 默认:" - "  

**返回**: arrRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************分割字符串***************** 命令原型: arrRet=Split(sText,sSeparator) 入参： sText--要处理的字符串。 sSeparator--用于分割字符串的分隔符。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ arrRet=Split("111,222,333",",") TracePrint(arrRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/Split.png)  

---

## 判断以指定前缀开头

**说明**: 判断目标字符串是否以指定前缀开头  

**原型**: `bRet = StartsWith(sText,sStartsStr)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 要判断的目标字符串  
- **sStartsStr** (True) [string] 默认:"" - 指定的前缀字符串  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************判断以指定前缀开头***************** 命令原型: bRet=StartsWith(sText,sStartsStr) 入参： sText--要判断的目标字符串。 sStartsStr--指定的前缀字符串。 出参： bRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ bRet=StartsWith("UiBot","Ui") TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/StartsWith.png)  

---

## 字符串比较

**说明**: 对两个字符串执行比较，字符串相同时返回 true，不相同返回 false  

**原型**: `bRet = StrComp(s1,s2,bCompare)`  

**参数**:  
- **s1** (True) [string] 默认:"" - 要比较的第一个字符串  
- **s2** (True) [string] 默认:"" - 要比较的第二个字符串  
- **bCompare** (True) [boolean] 默认:False - 对比字符串时是否区分大小写  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************字符串比较***************** 命令原型: bRet=StrComp(s1,s2,bCompare) 入参： s1--要比较的第一个字符串。 s2--要比较的第二个字符串。 bCompare--对比字符串时是否区分大小写。 出参： bRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ bRet=StrComp("UiBot","Ui",True) TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/StrComp.png)  

---

## 中间裁剪

**说明**: 从指定位置开始裁剪一定数量的字符  

**原型**: `sRet = StrCut(sData,iStart,iSize)`  

**参数**:  
- **sData** (True) [string] 默认:"" - 要处理的字符串  
- **iStart** (True) [number] 默认:1 - 裁剪字符串的开始位置  
- **iSize** (True) [number] 默认:1 - 裁剪字符串的长度  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************中间裁剪***************** 命令原型: sRet=StrCut(sData,iStart,iSize) 入参： sData--要处理的字符串。 iStart--裁剪字符串的开始位置。 iSize--裁剪字符串的长度。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=StrCut("UiBot",1,2) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/StrCut.png)  

---

## 获取字符

**说明**: 获取字符串指定位置的字符  

**原型**: `sRet = StrGetAt(sText,iPos)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 要处理的字符串  
- **iPos** (True) [number] 默认:1 - 要获取字符的位置  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************获取字符***************** 命令原型: sRet=StrGetAt(sText,iPos) 入参： sText--要处理的字符串。 iPos--要获取字符的位置。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=StrGetAt("UiBot",4) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/StrGetAt.png)  

---

## 获取MD5值

**说明**: 获取目标字符串的MD5值  

**原型**: `sRet = StrHash(sText)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 要获取MD5值的目标字符串  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************获取MD5值***************** 命令原型: sRet = StrHash(sText) 入参： sText--要获取MD5值的目标字符串。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet = StrHash("UiBot") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/StrHash.png)  

---

## 字符串指定长度比较

**说明**: 按指定的位数，从左开始对比2个字符串，字符串相同时返回true，否则返回false  

**原型**: `bRet = StrNComp(s1,s2,iLen,bCompare)`  

**参数**:  
- **s1** (True) [string] 默认:"" - 要比较的第一个字符串  
- **s2** (True) [string] 默认:"" - 要比较的第二个字符串  
- **iLen** (True) [number] 默认:1 - 要比较的字符串长度  
- **bCompare** (True) [boolean] 默认:False - 对比字符串时是否区分大小写  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************字符串指定长度比较***************** 命令原型: bRet=StrNComp(s1,s2,iLen,bCompare) 入参： s1--要比较的第一个字符串。 s2--要比较的第二个字符串。 iLen--要比较的字符串长度。 bCompare--对比字符串时是否区分大小写。 出参： bRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ bRet=StrNComp("Ui","UiBot",2,true) TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/StrNComp.png)  

---

## 颠倒文字

**说明**: 使字符串逆向排列  

**原型**: `sRet = StrReverse(sText)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 要处理的字符串  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************颠倒文字***************** 命令原型: sRet=StrReverse(sText) 入参： sText--要处理的字符串。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=StrReverse("UiBot") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/StrReverse.png)  

---

## 抽取指定长度字符

**说明**: 从字符串中抽取指定位置开始的指定数目的字符，位置从1开始  

**原型**: `sRet = SubStr(sData,iStart,iSize)`  

**参数**:  
- **sData** (True) [string] 默认:"" - 被抽取的源字符串  
- **iStart** (True) [number] 默认:1 - 抽取字符串的开始位置  
- **iSize** (True) [number] 默认:1 - 裁剪字符串的长度  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************抽取指定长度字符***************** 命令原型: sRet=SubStr(sData,iStart,iSize) 入参： sData--被抽取的源字符串。 iStart--抽取字符串的开始位置。 iSize--裁剪字符串的长度。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=SubStr("UiBot",1,2) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/SubStr.png)  

---

## 抽取指定位置字符

**说明**: 从字符串中抽取指定位置开始到指定位置结束的字符，位置从1开始  

**原型**: `sRet = SubString(sData,iStart,iEnd)`  

**参数**:  
- **sData** (True) [string] 默认:"" - 被抽取的源字符串  
- **iStart** (True) [number] 默认:1 - 抽取字符串的开始位置  
- **iEnd** (True) [number] 默认:1 - 抽取字符串的结束位置  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************抽取指定位置字符***************** 命令原型: sRet=SubString(sData,iStart,iEnd) 入参： sData--被抽取的源字符串。 iStart--抽取字符串的开始位置。 iEnd--抽取字符串的结束位置。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=SubString("UiBot",3,5) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/SubString.png)  

---

## 两侧裁剪

**说明**: 删除掉指定字符串两侧的特定字符  

**原型**: `sRet = Trim(sData,sTrim)`  

**参数**:  
- **sData** (True) [string] 默认:"" - 被裁剪的字符串  
- **sTrim** (True) [string] 默认:"" - 输入单个字符或由多个字符组成的字符串，字符组成的所有组合都会被裁剪  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************两侧裁剪***************** 命令原型: sRet=Trim(sData,sTrim) 入参： sData--被裁剪的字符串。 sTrim--单个字符或由多个字符组成的字符串。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=Trim("UiBot","Ui") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/LTrim.png)  

---

## 将字符串转换为大写

**说明**: 将字符串转换为大写  

**原型**: `sRet = UCase(sText)`  

**参数**:  
- **sText** (True) [string] 默认:"" - 进行操作的字符串  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/************************将字符串转换为大写***************** 命令原型: sRet=UCase(sText) 入参： sText--进行操作的字符串。 出参： sRet--函数调用的输出保存到的变量。 注意事项: 无 ***********************************************************/ sRet=UCase("uibot") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/String_图片/UCase.png)  

---

## 执行命令行

**说明**: 执行系统命令行，返回命令行执行过程中的控制台输出文本  

**原型**: `sRet = Sys.Command(sCommand)`  

**参数**:  
- **sCommand** (True) [string] 默认:"" - 要执行的命令行  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************************执行命令行************************************ 命令原型： sRet = Sys.Command("") 入参： sCommand--要执行的命令行 出参： sRet--命令运行后的结果 注意事项： 如果没有该命令，程序不会报错，只是返回空值 ********************************************************************************/ Dim sCommand="ipconfig" sRet = Sys.Command(sCommand) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Sys_图片/Sys_Command.png)  

---

## 读取环境变量

**说明**: 读取环境变量  

**原型**: `sRet = Sys.GetEnviron(sEnvironName)`  

**参数**:  
- **sEnvironName** (True) [string] 默认:"" - 要获取的环境变量名称  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************读取环境变量*********************************** 命令原型： sRet = Sys.GetEnviron("") 入参： sEnvironName--要获取的环境变量名称 出参： sRet--命令运行后的结果 注意事项： 建议先检查计算机系统是否存在该环境变量，若果不存在则会报错 ********************************************************************************/ Dim sEnvironName="OS" sRet = Sys.GetEnviron(sEnvironName) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Sys_图片/Sys_GetEnviron.png)  

---

## 获取用户文件夹路径

**说明**: 获取用户文件夹路径  

**原型**: `sRet = Sys.GetHomePath()`  

**参数**:  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*******************************获取用户文件夹路径********************************* 命令原型： sRet = Sys.GetHomePath() 入参： 无 出参： sRet--命令运行后的结果 注意事项： 无 ********************************************************************************/ sRet = Sys.GetHomePath() TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Sys_图片/Sys_GetHomePath.png)  

---

## 获取机器码

**说明**: 获取机器码，机器码可作为唯一标识使用  

**原型**: `sRet = Sys.GetMachineCode()`  

**参数**:  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/**********************************获取机器码************************************* 命令原型： sRet = Sys.GetMachineCode() 入参： 无 出参： sRet--命令运行后的结果 注意事项： 无 ********************************************************************************/ sRet = Sys.GetMachineCode() TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Sys_图片/Sys_GetMachineCode.png)  

---

## 获取系统文件夹路径

**说明**: 获取系统文件夹路径  

**原型**: `sRet = Sys.GetSystemPath(sFolderIndex)`  

**参数**:  
- **sFolderIndex** (True) [enum] 默认:"system" - 要获取哪一个文件夹，传递为 "system" 则获取系统文件夹路径，传递为 "windows" 则获取Windows文件夹路径，传递为 "desktop" 则获取桌面路径，传递为 "program" 则获取软件安装目录路径，传递为 "temp" 则获取临时目录路径，传递为 startmenu 则获取开始菜单文件夹路径  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*******************************获取系统文件夹路径********************************* 命令原型： sRet = Sys.GetSystemPath("system") 入参： sFolderIndex--要获取哪一个文件夹，默认“system” 出参： sRet--命令运行后的结果 注意事项： 默认获取系统目录，可以切换至源代码试图，在属性栏选择其他目录 ********************************************************************************/ sRet = Sys.GetSystemPath("system") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Sys_图片/Sys_GetSystemPath.png)  

---

## 获取临时文件夹路径

**说明**: 获取临时文件夹路径  

**原型**: `sRet = Sys.GetTempPath()`  

**参数**:  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*******************************获取临时文件夹路径********************************* 命令原型： sRet = Sys.GetTempPath() 入参： 无 出参： sRet--命令运行后的结果 注意事项： 无 ********************************************************************************/ sRet = Sys.GetTempPath() TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Sys_图片/Sys_GetTempPath.png)  

---

## 播放声音

**说明**: 播放一个wav格式的音频文件  

**原型**: `Sys.PlaySound(sPath)`  

**参数**:  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 存储音频文件的路径，仅支持wav格式  

**示例**:  
```
/***********************************播放声音************************************** 命令原型： Sys.PlaySound(&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27;) 入参： sPath--存储音频文件的路径，仅支持wav格式 出参： 无 注意事项： 仅支持wav格式，如果是其他后缀格式会报错 ********************************************************************************/ Dim sPath=&#x27;&#x27;&#x27;D:\app\music\Baby.wav&#x27;&#x27;&#x27; Sys.PlaySound(sPath)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Sys_图片/Sys_PlaySound.png)  

---

## 执行PowerShell

**说明**: 执行PowerShell，返回PowerShell执行过程中的控制台输出文本  

**原型**: `sRet = Sys.PowerShell(sCommand)`  

**参数**:  
- **sCommand** (True) [string] 默认:"" - 要执行的命令行  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*******************************执行PowerShell********************************** 命令原型： sRet = Sys.PowerShell("") 入参： sCommand--要执行的命令行 出参： sRet--命令运行后的结果 注意事项： 无 ********************************************************************************/ Dim sCommand="date" sRet = Sys.PowerShell(sCommand) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Sys_图片/Sys_PowerShell.png)  

---

## 设置环境变量

**说明**: 设置环境变量  

**原型**: `Sys.SetEnviron(sEnvironName,sValue)`  

**参数**:  
- **sEnvironName** (True) [string] 默认:"" - 要设置的环境变量名称  
- **sValue** (True) [string] 默认:"" - 环境变量的值  

**示例**:  
```
/*********************************设置环境变量*********************************** 命令原型： sRet = Sys.GetEnviron("") 入参： sEnvironName--设置的环境变量名称 sValue--环境变量的值 出参： 无 注意事项： 无 ********************************************************************************/ Dim sEnvironName="java_home" Dim sValue=&#x27;&#x27;&#x27;D:\app\Java\jdk1.7&#x27;&#x27;&#x27; Sys.SetEnviron(sEnvironName,sValue)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Sys_图片/Sys_SetEnviron.png)  

---

## 点击文本

**说明**: 按照规则搜索含有指定字符串的界面元素并点击这个界面元素，点击位置为查找到的文本位置  

**原型**: `Text.Click(objUiElement,sText,iRule,iOccurrence,iButton,iType,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 需要查找文本的父元素，程序会在这个元素内查找文本操作，当属性传递为 字符串 类型时，作为特征串查找界面元素后查找子元素，当属性传递为 UiElement 类型时，直接在这个 UiElement 元素中进行查找，如果传递为 null，则在所有窗口中查找  
- **sText** (True) [string] 默认:"" - 查找元素时使用的文本  
- **iRule** (True) [enum] 默认:"instr" - 查找文本时使用的规则  
- **iOccurrence** (True) [number] 默认:1 - 如果“文本”字段中的字符串在指示的界面元素中出现多次，请在此处指定要单击的出现次数。例如，如果字符串出现4次并且您要单击第一个匹配项，请在此字段中写入1  
- **iButton** (True) [enum] 默认:"left" - 鼠标按键 {left:左键, right:右键, middle:中键}  
- **iType** (True) [enum] 默认:"click" - 点击类型 {click:单击, dbclick:双击, down:按下, up:弹起}  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sCursorPosition** (False) [enum] 默认:"Center" - 描述添加OffsetX和OffsetY属性的偏移量的光标起点。可以使用以下选项：TopLeft，TopRight，BottomLeft，BottomRight和Center。默认选项是Center  
- **iCursorOffsetX** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的水平位移  
- **iCursorOffsetY** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的垂直位移  
- **sKeyModifiers** (False) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **sSimulate** (False) [enum] 默认:"simulate" - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：模拟操作(simulate)  

**示例**:  
```
/*********************************点击文本*************************************** 命令原型： Text.Click(@ui"","","instr",1,"left","click",10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate"}) 入参： objUiElement--识别目标。 sText--需要在目标中查找的文本。注：查找元素时使用的文本 iRule--查找规则。注：查找文本时使用的规则 iOccurrence--相似结果位置。 iButton--鼠标点击。注：鼠标按键 {left:左键, right:右键, middle:中键} iType--点击类型。注：点击类型 {click:单击, dbclick:双击, down:按下, up:弹起} iTimeOut--超时时间（毫秒）。注：指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发错误。默认30000毫秒（30秒） optionArgs--可选参数(包括:错误继续执行、执行后延时、执行前延时、激活窗口、光标位置、横坐标偏移、纵坐标偏移、辅助按键、操作类型).Type:Dict 注意事项： 1.在运行命令的同事要保证文本所在页面是打开的，否则命令会运行报错。 ********************************************************************************/ Text.Click(@ui"文本<span>_点击文本","点","instr",1,"left","click",10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate"})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Text_图片/Text_Click.png)  

---

## 判断文本是否存在

**说明**: 在指定元素中查找文本，文本存在返回 true，否则返回 false  

**原型**: `bRet = Text.Exists(objUiElement,sText,iRule,iOccurrence,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 需要查找文本的父元素，程序会在这个元素内查找文本操作，当属性传递为 字符串 类型时，作为特征串查找界面元素后查找子元素，当属性传递为 UiElement 类型时，直接在这个 UiElement 元素中进行查找，如果传递为 null，则在所有窗口中查找  
- **sText** (True) [string] 默认:"" - 查找元素时使用的文本  
- **iRule** (True) [enum] 默认:"instr" - 查找文本时使用的规则  
- **iOccurrence** (True) [number] 默认:1 - 如果“文本”字段中的字符串在指示的界面元素中出现多次，请在此处指定要单击的出现次数。例如，如果字符串出现4次并且您要单击第一个匹配项，请在此字段中写入1  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************判断文本是否存在*************************************** 命令原型： bRet = Text.Exists(@ui"","","instr",1,10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true}) 入参： objUiElement--识别目标。 sText--需要在目标中查找的文本。注：查找元素时使用的文本 iRule--查找规则。注：查找文本时使用的规则 iOccurrence--相似结果位置。 iTimeOut--超时时间（毫秒）。注：指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发错误。默认30000毫秒（30秒） optionArgs--可选参数(包括:错误继续执行、执行后延时、执行前延时、激活窗口、光标位置、横坐标偏移、纵坐标偏移、辅助按键、操作类型).Type:Dict 出参： bRet--函数调用的输出保存到的变量 注意事项： 1.在使用时要保证目标元素已经加载完成，否则很容易造成判断为false的情况。 ********************************************************************************/ Dim bRet = "" bRet = Text.Exists(@ui"块级元素<div>_Exists判断文本是否存在命令说明在指定元素中查找文本，文本存在返回true，","文本","instr",1,10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true}) TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Text_图片/Text_Exists.png)  

---

## 查找文本所在位置的界面元素

**说明**: 按照查找文本规则，查找出文本所在位置的界面元素  

**原型**: `arrElement = Text.FindElement(objUiElement,sText,iRule,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 通过鼠标选取的界面元素，包含窗口、元素等信息  
- **sText** (True) [string] 默认:"" - 查找时使用的文本  
- **iRule** (True) [enum] 默认:"instr" - 查找时使用的规则  
- **iTimeOut** (True) [number] 默认:30000 - 指定等待重试查找文本时间（以毫秒为单位），如果超出该时间，则引发错误。默认30000毫秒（30秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 当目标窗口为IE浏览器时，可设置操作前是否激活该窗口，默认为是  

**返回**: arrElement，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************查找文本所在位置的界面元素*************************************** 命令原型： arrElement = Text.FindElement(@ui"","","instr",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true}) 入参： objUiElement--识别目标。 sText--需要在目标中查找的文本。注：查找元素时使用的文本 iRule--查找规则。注：查找文本时使用的规则 iTimeOut--超时时间（毫秒）。注：指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发错误。默认30000毫秒（30秒） optionArgs--可选参数(包括:错误继续执行、执行后延时、执行前延时、激活窗口).Type:Dict 出参： arrElement--函数调用的输出保存到的变量 注意事项： 1.在使用时要注意查找规则的选择，区分字符串和正则表达式 ********************************************************************************/ Dim arrElement = "" arrElement = Text.FindElement(@ui"文本<span>_查找文本所在位置的界面元素","查找","instr",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true}) TracePrint(arrElement)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Text_图片/Text_FindElement.png)  

---

## 获取文本

**说明**: 获取指定界面元素的文本内容  

**原型**: `sRet = Text.Get(objUiElement,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行操作  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取文本*************************************** 命令原型： sRet = Text.Get(@ui"",10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true}) 入参： objUiElement--识别目标。 iTimeOut--超时时间（毫秒）。注：指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发错误。默认30000毫秒（30秒） optionArgs--可选参数(包括:错误继续执行、执行后延时、执行前延时、激活窗口).Type:Dict 出参： sRet--函数调用的输出保存到的变量 注意事项： 1.在使用该命令时要保证文本所在的目标网站时打开状态，否则命令会报错。 ********************************************************************************/ Dim sRet = "" sRet = Text.Get(@ui"块级元素<div>_Get获取文本命令说明获取指定界面元素的文本内容命令原型PlainText自动换",10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true}) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Text_图片/Text_Get.png)  

---

## 鼠标移动到文本上

**说明**: 按照规则搜索含有指定字符串的界面元素并将鼠标移动到这个界面元素上，鼠标停留位置为查找到的文本位置  

**原型**: `Text.Hover(objUiElement,sText,iRule,iOccurrence,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 需要查找文本的父元素，程序会在这个元素内查找文本操作，当属性传递为 字符串 类型时，作为特征串查找界面元素后查找子元素，当属性传递为 UiElement 类型时，直接在这个 UiElement 元素中进行查找，如果传递为 null，则在所有窗口中查找  
- **sText** (True) [string] 默认:"" - 查找元素时使用的文本  
- **iRule** (True) [enum] 默认:"instr" - 查找文本时使用的规则  
- **iOccurrence** (True) [number] 默认:1 - 如果“文本”字段中的字符串在指示的界面元素中出现多次，请在此处指定要单击的出现次数。例如，如果字符串出现4次并且您要单击第一个匹配项，请在此字段中写入1  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sCursorPosition** (False) [enum] 默认:"Center" - 描述添加OffsetX和OffsetY属性的偏移量的光标起点。可以使用以下选项：TopLeft，TopRight，BottomLeft，BottomRight和Center。默认选项是Center  
- **iCursorOffsetX** (False) [number] 默认:10 - 根据在“位置”字段中选择的选项，光标位置的水平位移  
- **iCursorOffsetY** (False) [number] 默认:10 - 根据在“位置”字段中选择的选项，光标位置的垂直位移  
- **sKeyModifiers** (False) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **sSimulate** (False) [enum] 默认:"simulate" - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：模拟操作(simulate)  

**示例**:  
```
/*********************************鼠标移动到文本上*************************************** 命令原型： Text.Hover(@ui"","","instr",1,10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true,"sCursorPosition":"Center","iCursorOffsetX":10,"iCursorOffsetY":10,"sKeyModifiers":[],"sSimulate":"simulate"}) 入参： objUiElement--识别目标。 sText--需要在目标中查找的文本。注：查找元素时使用的文本 iRule--查找规则。注：查找文本时使用的规则 iOccurrence--相似结果位置。 iTimeOut--超时时间（毫秒）。注：指定等待重试查找屏幕范围时间（以毫秒为单位），如果超出该时间，则引发错误。默认30000毫秒（30秒） optionArgs--可选参数(包括:错误继续执行、执行后延时、执行前延时、激活窗口、光标位置、横坐标偏移、纵坐标偏移、辅助按键、操作类型).Type:Dict 注意事项： 1.在该命令执行时要保证目标文本页面处于打开状态，并且不要触碰鼠标，显示器分辨率变更会影响该命令的执行。 ********************************************************************************/ Text.Hover(@ui"文本<span>_鼠标移动到文本上","文本","instr",1,10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true,"sCursorPosition":"Center","iCursorOffsetX":10,"iCursorOffsetY":10,"sKeyModifiers":[],"sSimulate":"simulate"})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Text_图片/Text_Hover.png)  

---

## 获取时间-无日期

**说明**: 获取浮点数形式表示的时间(不包含日期，以当前计算机设置的系统日期和时间为对象)，可通过格式化时间命令输出指定格式时间  

**原型**: `dTime = Time.Time()`  

**参数**:  

**返回**: dTime，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取时间-无日期*************************************** 命令原型： Time.Time() 出参： dTime -- 将命令运行后的结果赋值给此变量 注意事项： 此命令无参数，可搭配命令：格式化时间（Time.Format()）使用 **********************************************************************************/ Dim dTime dTime = Time.Time() TracePrint(dTime)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_CDate.png)  

---

## 字符串转换为时间

**说明**: 将一个字符串转换为时间数据  

**原型**: `dTime = Time.CDate(sText, sFormat)`  

**参数**:  
- **sText** (True) [string] 默认:"2020年1月1日 12:00:00" - 判断是否能够转换为时间数据的字符串  
- **sFormat** (True) [string] 默认:"yyyy.mm.dd.hh.mm.ss" - 时间格式，&#x27;.&#x27;代表任意非数字字符  

**返回**: dTime，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************字符串转换为时间*************************************** 命令原型： Time.CDate("2020年1月1日 12:00:00", "yyyy.mm.dd.hh.mm.ss") 入参： sText -- 判断是否能够转换为时间数据的字符串 sFormat -- 时间格式，&#x27;.&#x27;代表任意非数字字符 出参： dTime -- 将命令运行后的结果赋值给此变量 注意事项： 时间格式中"yyyy.mm.dd.hh.mm.ss"分别代表年、月、日、时、分、秒，且转换时请注意保持时间文本与时间文本格式保持一致。 例如："2020年1月1日 12:00:00"对应"yyyy.mm.dd.hh.mm.ss"、"2020年1月1日"对应"yyyy.mm.dd."、"1日1月2020年"对应"dd.mm.yyyy." **********************************************************************************/ Dim dTime dTime = Time.CDate("2022年2月9日 14:30:00", "yyyy.mm.dd.hh.mm.ss") TracePrint(dTime)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_CDate.png)  

---

## 获取时间-日期

**说明**: 获取从1900年1月1日开始到现在经过的天数（以当前计算机设置的系统日期和时间为对象），可通过格式化时间命令输出指定日期形式  

**原型**: `dTime = Time.Date()`  

**参数**:  

**返回**: dTime，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取时间-日期*************************************** 命令原型： Time.Date() 出参： dTime -- 将命令运行后的结果赋值给此变量 注意事项： 此命令无参数，可搭配命令：格式化时间（Time.Format()）使用 **********************************************************************************/ Dim dTime dTime = Time.Date() TracePrint(dTime)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_CDate.png)  

---

## 改变时间

**说明**: 改变时间中的一个度量单位  

**原型**: `iRet = Time.DateAdd(sUnit,iCount,dTime)`  

**参数**:  
- **sUnit** (True) [enum] 默认:"s" - 要改变的时间单位  
- **iCount** (True) [number] 默认:1 - 调整时间增加多少个时间单位  
- **dTime** (True) [expression] 默认:tRet - 需要改变的时间  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************改变时间*************************************** 命令原型： Time.DateAdd("s",1,tRet) 入参： sUnit -- 要改变的时间单位 yyyy：年份；q：季度；m：月份；ww：星期；d：天数；h：小时；n：分钟；s：秒 iCount -- 调整时间增加多少个时间单位 dTime -- 需要改变的时间 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 参数dTime需要输入时间,时间参数可通过命令：获取时间（Time.Now()）获取当前时间 **********************************************************************************/ Dim iRet iRet = Time.DateAdd("d",1,44601.0) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_DateAdd.png)  

---

## 计算时间差

**说明**: 对比两个时间，得出两个时间指定单位的时间差  

**原型**: `iRet = Time.DateDiff(sUnit,dTime1,dTime2)`  

**参数**:  
- **sUnit** (True) [enum] 默认:"s" - 要改变的时间单位  
- **dTime1** (True) [expression] 默认:dTime1 - 需要对比的时间  
- **dTime2** (True) [expression] 默认:dTime2 - 需要对比的时间  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************计算时间差*************************************** 命令原型： Time.DateDiff("s",dTime1,dTime2) 入参： sUnit -- 要改变的时间单位 yyyy：年份；q：季度；m：月份；ww：星期；w：完整的星期（7天周期）；d：天数；h：小时；n：分钟；s：秒 dTime1 -- 需要对比的时间 dTime2 -- 需要对比的时间 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 参数dTime1、dTime2需要输入时间,时间参数可通过命令：获取时间（Time.Now()）获取当前时间 **********************************************************************************/ Dim iRet iRet = Time.DateDiff("s",44601.0,44602.0) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_DateDiff.png)  

---

## 获取时间中的某个单位

**说明**: 获取一个时间数据指定单位的部分  

**原型**: `iRet = Time.DatePart(sUnit, dTime)`  

**参数**:  
- **sUnit** (True) [enum] 默认:"s" - 要获取的时间单位  
- **dTime** (True) [expression] 默认:tRet - 需要获取数据的时间  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取时间中的某个单位*************************************** 命令原型： Time.DatePart("s", tRet) 入参： sUnit -- 要改变的时间单位 yyyy：年份；q：季度；m：月份；w：本周第几天；ww：当年第几周；y：当年第几天；d：当月第几天；h：小时；n：分钟；s：秒 dTime -- 需要获取数据的时间 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 参数dTime需要输入时间,时间参数可通过命令：获取时间（Time.Now()）获取当前时间 **********************************************************************************/ Dim iRet iRet = Time.DatePart("s", 44601.610763889) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_DatePart.png)  

---

## 构造日期

**说明**: 根据 年、月、日 构造一个日期  

**原型**: `dTime = Time.DateSerial(iYear, iMonth, iDay)`  

**参数**:  
- **iYear** (True) [number] 默认:2020 - 要构造的时间在哪一年  
- **iMonth** (True) [number] 默认:1 - 要构造的时间在哪一月  
- **iDay** (True) [number] 默认:1 - 要构造的时间在哪一天  

**返回**: dTime，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************构造日期*************************************** 命令原型： Time.DateSerial(2020, 1, 1) 入参： iYear -- 要构造的时间在哪一年 iMonth -- 要构造的时间在哪一月 iDay -- 要构造的时间在哪一天 出参： dTime -- 将命令运行后的结果赋值给此变量 注意事项： 三个参数iYear、iMonth、iDay均为number类型 **********************************************************************************/ Dim dTime dTime = Time.DateSerial(2022, 2, 9) TracePrint(dTime)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_DateSerial.png)  

---

## 获取第几天

**说明**: 获取时间中的日期  

**原型**: `iRet = Time.Day(dTime)`  

**参数**:  
- **dTime** (True) [expression] 默认:tRet - 获取数据的时间  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取第几天*************************************** 命令原型： Time.Day(tRet) 入参： dTime -- 获取数据的时间 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 参数dTime需要输入时间,时间参数可通过命令：获取时间（Time.Now()）获取当前时间 **********************************************************************************/ Dim iRet iRet = Time.Day(44601.0) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_Day.png)  

---

## 格式化时间

**说明**: 将获取的时间格式化成指定的字符串形式（时间可通过获取时间命令获取）  

**原型**: `sRet = Time.Format(dTime,sFormat)`  

**参数**:  
- **dTime** (True) [expression] 默认:dTime - 获取数据的时间  
- **sFormat** (True) [string] 默认:"yyyy-mm-dd hh:mm:ss" - 格式化字符串  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************格式化时间*************************************** 命令原型： Time.Format(dTime,"yyyy-mm-dd hh:mm:ss") 入参： dTime -- 获取数据的时间 sFormat -- 格式化字符串 出参： sRet -- 将命令运行后的结果赋值给此变量 注意事项： 参数dTime需要输入时间,时间参数可通过命令：获取时间（Time.Now()）获取当前时间。时间格式中"yyyy.mm.dd.hh.mm.ss"分别代表年、月、日、时、分、秒。 可自由搭配转换成所需要的格式，例如："yyyy-mm-dd hh:mm:ss"、"yyyy-mm-dd"、"yyyy/mm/dd hh:mm:ss"、"dd-mm-yyyy" /************************************************************************/ Dim sRet sRet = Time.Format(44601.61681713,"yyyy-mm-dd hh:mm:ss") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_Format.png)  

---

## Unix时间戳转换为时间

**说明**: 将一个Unix时间戳转换为时间数据（从1900年1月1日开始到现在经过的时间）  

**原型**: `dRet = Time.FromUnixTime(dTime, bMS)`  

**参数**:  
- **dTime** (True) [expression] 默认:dTime - 要转换的Unix时间戳  
- **bMS** (True) [boolean] 默认:False - Unix时间戳是否为毫秒精度，缺省为 false  

**返回**: dRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************Unix时间戳转换为时间*************************************** 命令原型： Time.FromUnixTime(dTime, False) 入参： dTime -- 要转换的Unix时间戳 bMS -- Unix时间戳是否为毫秒精度，缺省为 false 出参： dRet -- 将命令运行后的结果赋值给此变量 注意事项： 参数dTime需要输入时间,时间参数可通过命令：获取时间（Time.Now()）获取当前时间 **********************************************************************************/ Dim dRet dRet = Time.FromUnixTime(44601.623032407, false) TracePrint(dRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_FromUnixTime.png)  

---

## 获取小时

**说明**: 获取时间中的小时  

**原型**: `iRet = Time.Hour(dTime)`  

**参数**:  
- **dTime** (True) [expression] 默认:tRet - 获取数据的时间  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取小时*************************************** 命令原型： Time.Hour(tRet) 入参： dTime -- 获取数据的时间 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 参数dTime需要输入时间,时间参数可通过命令：获取时间（Time.Now()）获取当前时间 **********************************************************************************/ Dim iRet iRet = Time.Hour(44601.623032407) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_Hour.png)  

---

## 判断是否能转换为时间

**说明**: 判断一个字符串是否能够转换为时间数据  

**原型**: `bRet = Time.IsDate(sText, sFormat)`  

**参数**:  
- **sText** (True) [string] 默认:"2020年1月1日 12:00:00" - 判断是否能够转换为时间数据的字符串  
- **sFormat** (True) [string] 默认:"yyyy.mm.dd.hh.mm.ss" - 时间文本格式，&#x27;.&#x27;代表任意非数字字符  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************判断是否能转换为时间*************************************** 命令原型： Time.IsDate("2020年1月1日 12:00:00", "yyyy.mm.dd.hh.mm.ss") 入参： sText -- 判断是否能够转换为时间数据的字符串 sFormat -- 时间文本格式，&#x27;.&#x27;代表任意非数字字符 出参： bRet -- 将命令运行后的结果赋值给此变量 注意事项： 时间格式中"yyyy.mm.dd.hh.mm.ss"分别代表年、月、日、时、分、秒，且转换时请注意保持时间文本与时间文本格式保持一致。 例如："2020年1月1日 12:00:00"对应"yyyy.mm.dd.hh.mm.ss"、"2020年1月1日"对应"yyyy.mm.dd."、"1日1月2020年"对应"dd.mm.yyyy." **********************************************************************************/ Dim bRet bRet = Time.IsDate("2020年1月1日 12:00:00", "yyyy.mm.dd.hh.mm.ss") TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_IsDate.png)  

---

## 获取分钟

**说明**: 获取时间中的分钟  

**原型**: `iRet = Time.Minute(dTime)`  

**参数**:  
- **dTime** (True) [expression] 默认:tRet - 获取数据的时间  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取分钟*************************************** 命令原型： Time.Minute(tRet) 入参： dTime -- 获取数据的时间 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 参数dTime需要输入时间,时间参数可通过命令：获取时间（Time.Now()）获取当前时间 **********************************************************************************/ Dim iRet iRet = Time.Minute(44601.623032407) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_Minute.png)  

---

## 获取月份

**说明**: 获取时间中的月份  

**原型**: `iRet = Time.Month(dTime)`  

**参数**:  
- **dTime** (True) [expression] 默认:tRet - 获取数据的时间  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取月份*************************************** 命令原型： Time.Month(tRet) 入参： dTime -- 获取数据的时间 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 参数dTime需要输入时间,时间参数可通过命令：获取时间（Time.Now()）获取当前时间 **********************************************************************************/ Dim iRet iRet = Time.Month(44601.623032407) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_Month.png)  

---

## 获取时间

**说明**: 获取从1900年1月1日开始到现在经过的时间（以当前计算机设置的系统日期和时间为对象），可通过格式化时间命令输出指定格式化时间  

**原型**: `dTime = Time.Now()`  

**参数**:  

**返回**: dTime，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取时间*************************************** 命令原型： Time.Now() 出参： dTime -- 将命令运行后的结果赋值给此变量 注意事项： 此命令无参数，可搭配命令：格式化时间（Time.Format()）使用 **********************************************************************************/ Dim dTime dTime = Time.Now() TracePrint(dTime)
```  

---

## 获取秒数

**说明**: 获取时间中的秒数  

**原型**: `iRet = Time.Second(dTime)`  

**参数**:  
- **dTime** (True) [expression] 默认:tRet - 获取数据的时间  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取秒数*************************************** 命令原型： Time.Second(tRet) 入参： dTime -- 获取数据的时间 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 参数dTime需要输入时间,时间参数可通过命令：获取时间（Time.Now()）获取当前时间 **********************************************************************************/ Dim iRet iRet = Time.Second(44601.631793981) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_Second.png)  

---

## 获取时间戳

**说明**: 返回当前时间的时间戳（开机后经过的浮点秒数）  

**原型**: `dTime = Time.Timer()`  

**参数**:  

**返回**: dTime，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取时间戳*************************************** 命令原型： Time.Timer() 出参： dTime -- 将命令运行后的结果赋值给此变量 注意事项： 返回的时间戳指的是开机后经过的浮点秒数 **********************************************************************************/ Dim dTime dTime = Time.Timer() TracePrint(dTime)
```  

---

## 构造时间-无日期

**说明**: 根据 时、分、秒 构造一个日期  

**原型**: `dTime = Time.TimeSerial(iHour, iMinute, iSecond)`  

**参数**:  
- **iHour** (True) [number] 默认:12 - 要构造的时间在几点  
- **iMinute** (True) [number] 默认:0 - 要构造的时间在几分  
- **iSecond** (True) [number] 默认:0 - 要构造的时间在几秒  

**返回**: dTime，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************构造时间-无日期*************************************** 命令原型： Time.TimeSerial(12, 0, 0) 入参： iHour -- 要构造的时间在几点 iMinute -- 要构造的时间在几分 iSecond -- 要构造的时间在几秒 出参： dTime -- 将命令运行后的结果赋值给此变量 注意事项： 三个参数iHour、iMinute、iSecond均为number类型 **********************************************************************************/ Dim dTime dTime = Time.TimeSerial(15, 20, 0) TracePrint(dTime)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_TimeSerial.png)  

---

## 时间转换为Unix时间戳

**说明**: 将一个时间数据（通过获取时间命令获取）转换为Unix时间戳  

**原型**: `iRet = Time.ToUnixTime(dTime, bMS)`  

**参数**:  
- **dTime** (True) [expression] 默认:dTime - 要转换的时间数据，缺省为当前时间  
- **bMS** (True) [boolean] 默认:False - 是否转换为毫秒精度的时间戳，缺省为 false  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************时间转换为Unix时间戳*************************************** 命令原型： Time.ToUnixTime(dTime, False) 入参： dTime -- 要转换的时间数据，缺省为当前时间 bMS -- 是否转换为毫秒精度的时间戳，缺省为 false 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 参数dTime需要输入时间,时间参数可通过命令：获取时间（Time.Now()）获取当前时间 **********************************************************************************/ Dim iRet iRet = Time.ToUnixTime(44601.642673611, false) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_ToUnixTime.png)  

---

## 获取本周第几天

**说明**: 获取时间相对于本周是第几天  

**原型**: `iRet = Time.WeekDay(dTime)`  

**参数**:  
- **dTime** (True) [expression] 默认:tRet - 获取数据的时间  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取本周第几天*************************************** 命令原型： Time.WeekDay(tRet) 入参： dTime -- 获取数据的时间 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 参数dTime需要输入时间,时间参数可通过命令：获取时间（Time.Now()）获取当前时间 **********************************************************************************/ Dim iRet iRet = Time.WeekDay(44601.642673611) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_WeekDay.png)  

---

## 获取年份

**说明**: 获取时间中的年份  

**原型**: `iRet = Time.Year(dTime)`  

**参数**:  
- **dTime** (True) [expression] 默认:tRet - 获取数据的时间  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取年份*************************************** 命令原型： Time.Year(tRet) 入参： dTime -- 获取数据的时间 出参： iRet -- 将命令运行后的结果赋值给此变量 注意事项： 参数dTime需要输入时间,时间参数可通过命令：获取时间（Time.Now()）获取当前时间 **********************************************************************************/ Dim iRet iRet = Time.Year(44601.642673611) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Time_图片/Time_Year.png)  

---

## 智能识别后点击

**说明**: 智能识别后点击元素  

**原型**: `UiDetection.Click(objUiElement,iButton,iType,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [expression] 默认:{ } - 选取的智能识别后的目标元素及锚点元素的特征值  
- **iButton** (True) [enum] 默认:"left" - 鼠标按键 { left:左键, right:右键, middle:中键 }  
- **iType** (True) [enum] 默认:"click" - 点击类型 { click:单击, dbclick:双击, down:按下, up:弹起 }  
- **iTimeOut** (True) [number] 默认:30000 - 查找目标引发异常之前等待命令重试运行的时间量，以毫秒为单位。默认30000(30秒)  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sCursorPosition** (False) [enum] 默认:"Center" - 描述添加OffsetX和OffsetY属性的偏移量的光标起点。可以使用以下选项：TopLeft，TopRight，BottomLeft，BottomRight和Center。默认选项是Center  
- **iCursorOffsetX** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的水平位移  
- **iCursorOffsetY** (False) [number] 默认:0 - 根据在“位置”字段中选择的选项，光标位置的垂直位移  
- **sKeyModifiers** (False) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **sSimulate** (False) [enum] 默认:"simulate" - 可选择操作类型为：后台操作(uia)、模拟操作(simulate)、系统消息(message)，默认选择：模拟操作(simulate)  

**示例**:  
```
/*********************************智能识别后点击*************************************** 命令原型： UiDetection.Click({},"left","click",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate"}) 入参： objUiElement--识别目标。 iButton--鼠标点击。注：鼠标按键 {left:左键, right:右键, middle:中键} iType--点击类型。注：点击类型 {click:单击, dbclick:双击, down:按下, up:弹起} iTimeOut--超时时间。注：指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒） optionArgs--可选参数(包括：错误继续执行、执行后延时、执行前延时、激活窗口、光标位置、横坐标偏移、纵坐标偏移、辅助按键、操作类型).Type:Dict 注意事项： 1.该命令只进行点击操作，需要结合智能识别屏幕范围命令一同使用。 2.执行命令前需要目标屏幕存在，否则会报错。 ********************************************************************************/ #icon("@res:38688e20-7f20-11ec-ab08-1d7b5faf0fee.png") UiDetection.Click({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"Click - 飞书云文档 - Google Chrome","app":"chrome"}],"cv_engine_version":1,"cv_region":{"x":0,"y":0,"width":0,"height":0},"cv_descriptor":{"anchors":[{"cls_type":100,"height":41,"text":"Click","width":89,"x":584,"y":225}],"confidence":0.800000011920929,"cv_handle":"\"2b3223b0-7f20-11ec-ab08-1d7b5faf0fee\"","match_version":1,"target":{"cls_type":100,"height":41,"text":"智能识别后点击","width":205,"x":584,"y":296}}},"left","click",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true,"sCursorPosition":"Center","iCursorOffsetX":0,"iCursorOffsetY":0,"sKeyModifiers":[],"sSimulate":"simulate"})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiDetection_图片/UiDetection_Click.png)  

---

## 智能识别后判断元素存在

**说明**: 智能识别后判断元素是否存在  

**原型**: `bRet = UiDetection.Exists(objUiElement,optionArgs)`  

**参数**:  
- **objUiElement** (True) [expression] 默认:{ } - 选取的智能识别后的目标元素及锚点元素的特征值  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:200 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:300 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************智能识别后判断元素存在*************************************** 命令原型： bRet = UiDetection.Exists({},{"bContinueOnError":false,"iDelayAfter":200,"iDelayBefore":300}) 入参： objUiElement--识别目标。 optionArgs--可选参数(包括:错误继续执行、执行后延时、执行前延时).Type:Dict 出参： bRet--函数调用的输出保存到的变量。 注意事项： 1.该命令只进行元素判断，需要结合智能识别屏幕范围命令一同使用。 2.执行命令前需要目标屏幕存在，否则会报错。 ********************************************************************************/ Dim bRet = "" #icon("@res:ad0d4140-7f24-11ec-ab08-1d7b5faf0fee.png") bRet = UiDetection.Exists({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"Click - 飞书云文档 - Google Chrome","app":"chrome"}],"cv_engine_version":1,"cv_region":{"x":0,"y":0,"width":0,"height":0},"cv_descriptor":{"anchors":[{"cls_type":100,"height":33,"text":"命令说明","width":99,"x":586,"y":357}],"confidence":0.800000011920929,"cv_handle":"\"2b3223b0-7f20-11ec-ab08-1d7b5faf0fee\"","match_version":1,"target":{"cls_type":100,"height":25,"text":"智能识别后点击元素","width":153,"x":589,"y":400}}},{"bContinueOnError":false,"iDelayAfter":200,"iDelayBefore":300}) TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiDetection_图片/UiDetection_Exists.png)  

---

## 智能识别后获取文本

**说明**: 获取智能识别的文本  

**原型**: `sRet = UiDetection.Get(objUiElement,sType,iTimeout,optionArgs)`  

**参数**:  
- **objUiElement** (True) [expression] 默认:{ } - 选取的智能识别后的目标元素及锚点元素的特征值  
- **sType** (True) [enum] 默认:"OCR" - OCR方式为识别框选范围；选择全部方式为模拟鼠标在目标中从左至右划取文字  
- **iTimeout** (True) [number] 默认:30000 - 查找目标引发异常之前等待命令重试运行的时间量，以毫秒为单位。默认30000(30秒)  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************智能识别后获取文本*************************************** 命令原型： sRet = UiDetection.Get({},"OCR",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) 入参： objUiElement--识别目标。 sType--点击类型。注：OCR方式为识别框选范围；选择全部方式为模拟鼠标在目标中从左至右划取文字 iTimeOut--超时时间。注：查找目标引发异常之前等待命令重试运行的时间量，以毫秒为单位。默认30000(30秒) optionArgs--可选参数(包括:错误继续执行、执行后延时、执行前延时、激活窗口、光标位置、横坐标偏移、纵坐标偏移、辅助按键、操作类型).Type:Dict 出参： sRet--函数调用的输出保存到的变量。 注意事项： 1.该命令只进行文本获取，需要结合智能识别屏幕范围命令一同使用。 2.执行命令前需要目标屏幕存在，否则会报错。 ********************************************************************************/ Dim sRet = "" #icon("@res:f334e0e0-7f21-11ec-ab08-1d7b5faf0fee.png") sRet = UiDetection.Get({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"Click - 飞书云文档 - Google Chrome","app":"chrome"}],"cv_engine_version":1,"cv_region":{"x":0,"y":0,"width":0,"height":0},"cv_descriptor":{"anchors":[{"cls_type":100,"height":41,"text":"智能识别后点击","width":205,"x":584,"y":296}],"confidence":0.800000011920929,"cv_handle":"\"2b3223b0-7f20-11ec-ab08-1d7b5faf0fee\"","match_version":1,"target":{"cls_type":100,"height":33,"text":"命令说明","width":99,"x":586,"y":357}}},"OCR",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiDetection_图片/UiDetection_Get.png)  

---

## 智能识别后鼠标悬停

**说明**: 鼠标悬停在元素上  

**原型**: `UiDetection.Hover(objUiElement,iTimeout,optionArgs)`  

**参数**:  
- **objUiElement** (True) [expression] 默认:{ } - 选取的智能识别后的目标元素及锚点元素的特征值  
- **iTimeout** (True) [number] 默认:30000 - 查找目标引发异常之前等待命令重试运行的时间量，以毫秒为单位。默认30000(30秒)  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  

**示例**:  
```
/*********************************智能识别后鼠标悬停*************************************** 命令原型： UiDetection.Hover({},30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true}) 入参： objUiElement--识别目标。 iTimeOut--超时时间（毫秒）。注：查找目标引发异常之前等待命令重试运行的时间量，以毫秒为单位。默认30000(30秒) optionArgs--可选参数(包括:错误继续执行、执行后延时、执行前延时、激活窗口、光标位置、横坐标偏移、纵坐标偏移、辅助按键、操作类型).Type:Dict 注意事项： 1.该命令只进行鼠标移动，需要结合智能识别屏幕范围命令一同使用。 2.执行命令前需要目标屏幕存在，否则会报错。 ********************************************************************************/ #icon("@res:5feef9d0-7f24-11ec-ab08-1d7b5faf0fee.png") UiDetection.Hover({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"Click - 飞书云文档 - Google Chrome","app":"chrome"}],"cv_engine_version":1,"cv_region":{"x":0,"y":0,"width":0,"height":0},"cv_descriptor":{"anchors":[{"cls_type":100,"height":41,"text":"Click","width":89,"x":584,"y":225}],"confidence":0.800000011920929,"cv_handle":"\"2b3223b0-7f20-11ec-ab08-1d7b5faf0fee\"","match_version":1,"target":{"cls_type":100,"height":41,"text":"智能识别后点击","width":205,"x":584,"y":296}}},30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"bSetForeground":true})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiDetection_图片/UiDetection_Hover.png)  

---

## 智能识别后输入文本

**说明**: 智能识别后输入文本  

**原型**: `UiDetection.InputText(objUiElement,sText,bEmptyField,iInterval,iTimeout,optionArgs)`  

**参数**:  
- **objUiElement** (True) [expression] 默认:{ } - 选取的智能识别后的目标元素及锚点元素的特征值  
- **sText** (True) [string] 默认:"" - 指定的目标元素中写入文本  
- **bEmptyField** (True) [boolean] 默认:True - 写入文本之前是否清空输入框  
- **iInterval** (True) [number] 默认:200 - 两次输入的间隔时间，间隔时间过小可能导致丢字，建议默认200毫秒  
- **iTimeout** (True) [number] 默认:30000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为30000毫秒（30秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **sKeyModifiers** (False) [set] 默认:[] - 触发鼠标动作时同时按下的键盘按键，可以使用以下选项：Alt，Ctrl，Shift，Win  
- **bSetForeground** (False) [boolean] 默认:True - 进行操作之前，是否先将目标窗口激活  
- **sSimulate** (False) [enum] 默认:"message" - 可选择操作类型为：模拟操作(simulate)、系统消息(message)，默认选择：系统消息(message)  

**示例**:  
```
/*********************************智能识别后输入文本*************************************** 命令原型： UiDetection.InputText({},"",true,200,30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sKeyModifiers":[],"bSetForeground":true,"sSimulate":"message"}) 入参： objUiElement--识别目标。注：选取的智能识别后的目标元素及锚点元素的特征值 sText--写入文本。注：指定的目标元素中写入文本 bEmptyField--清空原内容。注：写入文本之前是否清空输入框 iInterval--键入间隔（毫秒）。注：两次输入的间隔时间，间隔时间过小可能导致丢字，建议默认200毫秒 iTimeOut--超时时间（毫秒）。注：查找目标引发异常之前等待命令重试运行的时间量，以毫秒为单位。默认30000(30秒) optionArgs--可选参数(包括:错误继续执行、执行后延时、执行前延时、激活窗口、光标位置、横坐标偏移、纵坐标偏移、辅助按键、操作类型).Type:Dict 注意事项： 1.该命令只进行文本输入，需要结合智能识别屏幕范围命令一同使用。 2.执行命令前需要目标屏幕存在，否则会报错。 ********************************************************************************/ #icon("@res:95f63900-7f22-11ec-ab08-1d7b5faf0fee.png") UiDetection.InputText({"wnd":[{"cls":"Chrome_WidgetWin_1","title":"Click - 飞书云文档 - Google Chrome","app":"chrome"}],"cv_engine_version":1,"cv_region":{"x":0,"y":0,"width":0,"height":0},"cv_descriptor":{"confidence":0.800000011920929,"cv_handle":"\"2b3223b0-7f20-11ec-ab08-1d7b5faf0fee\"","match_version":1,"target":{"cls_type":100,"height":17,"text":"UiDetection.C1ick(objUiElement,iButton,iType,iTimeout,optionArgs","width":529,"x":643,"y":544}}},"测试",true,200,30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sKeyModifiers":[],"bSetForeground":true,"sSimulate":"message"})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiDetection_图片/UiDetection_InputText.png)  

---

## 智能识别屏幕范围

**说明**: 在窗口或者元素上启用智能识别的块级窗口，并可以执行屏幕范围图片抓取  

**原型**: `UiDetect Scope(id,objUiElement,objRect,iTimeOut,optionArgs) End UiDetect`  

**参数**:  
- **id** (True) [string] 默认:$uuid - 智能识别屏幕范围的唯一识别标识，当前流程块内重复可能会导致操作错误  
- **objUiElement** (True) [expression] 默认:{ } - 通过鼠标选取或截取需要智能识别的目标屏幕范围  
- **objRect** (True) [dictionary] 默认:{ "x":0,"y":0,"width":0,"height":0 } - 需要查找的范围，程序会在控件这个范围内进行文字识别，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则进行控件矩形区域范围内的文字识别  
- **iTimeOut** (True) [number] 默认:30000 - 查找目标引发异常之前等待命令重试运行的时间量，以毫秒为单位。默认30000(30秒)  
- **sDetectType** (False) [enum] 默认:"sLocalAI" - 智能识别的方式  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/*********************************智能识别屏幕范围*************************************** 命令原型： UiDetect Scope("73adc4b0-d66d-11ec-8a07-b50840ecfdf2",{},{"x":0,"y":0,"width":0,"height":0},30000,{"sDetectType":"sLocalAI","bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) End UiDetect 入参： id--唯一键。注：智能识别屏幕范围的唯一识别标识，当前流程块内重复可能会导致操作错误 objUiElement--识别目标。注：通过鼠标选取或截取需要智能识别的目标屏幕范围 objRect--识别范围。 iTimeOut--超时时间。注：查找目标引发异常之前等待命令重试运行的时间量，以毫秒为单位。默认30000(30秒) optionArgs--可选参数(包括:识别方式、错误继续执行、执行后延时、执行前延时).Type:Dict 注意事项： 1.该命令只抓取屏幕操作范围，需要结合其他命令一同使用。 ********************************************************************************/ #icon("@res:3310e8f0-7f20-11ec-ab08-1d7b5faf0fee.png") UiDetect Scope("2b3223b0-7f20-11ec-ab08-1d7b5faf0fee",{"wnd":[{"cls":"Chrome_WidgetWin_1","title":"Click - 飞书云文档 - Google Chrome","app":"chrome"}],"cv_engine_version":1},{"x":0,"y":0,"width":0,"height":0},30000,{"sDetectType":"sLocalAI","bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) End UiDetect
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiDetection_图片/UiDetection_Scope.png)  

---

## 判断元素是否存在

**说明**: 判断元素是否存在，如果元素存在，返回true，如果元素不存在，返回 false  

**原型**: `bRet = UiElement.Exists(objUiElement,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 元素特征字符串  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/******************************判断元素是否存在******************************* 命令原型: bRet = UiElement.Exists(objUiElement,optionArgs) 入参: objUiElement--目标元素 optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参: bRet--函数调用的输出保存到的变量 注意事项: 必须选定元素 *********************************************************************/ bRet = UiElement.Exists(@ui"输入控件<input>5",{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiElement_图片/UiElement_Exists.png)  

---

## 获取元素属性

**说明**: 获取元素的属性  

**原型**: `sRet = UiElement.GetAttribute(objUiElement,sAttribute,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 元素特征字符串  
- **sAttribute** (True) [string] 默认:"" - 要获取的属性名  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/******************************获取元素属性******************************* 命令原型: sRet = UiElement.GetAttribute(objUiElement,sAttribute,optionArgs) 入参: objUiElement--目标元素 sAttribute--包含元素 optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参: sRet--函数调用的输出保存到的变量 注意事项: 无 *********************************************************************/ sRet = UiElement.GetAttribute(@ui"输入控件<input>6","type",{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiElement_图片/UiElement_GetAttribute.png)  

---

## 获取元素勾选

**说明**: 获取元素的勾选（可以操作单选框、复选框 等类型的元素）  

**原型**: `bRet = UiElement.GetCheck(objUiElement,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 元素特征字符串  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*****************************获取元素勾选******************************** 命令原型: bRet = UiElement.GetCheck(objUiElement,optionArgs) 入参: objUiElement--目标元素 optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参: bRet--函数调用的输出保存到的变量 注意事项: 无 *********************************************************************/ bRet = UiElement.GetCheck(@ui"复选框<checkboxinput>_on",{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiElement_图片/UiElement_GetCheck.png)  

---

## 获取子元素

**说明**: 一个指定的目标元素可能由多个子元素聚合而成，而子元素自身可能又是被更内层的多个子元素聚合而成，依次类推而形成一个树结构，通过指定子元素层级，可获取至层级(含)范围内的所有元素，并以一维数组的形式返回，且数组内的元素为内存地址对象  

**原型**: `arrElement = UiElement.GetChildren(objUiElement,level, optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 指定被获取子元素的根节点元素  
- **level** (True) [number] 默认:1 - 默认子元素层级为1，即根节点元素下的第1级所有元素(子元素)。当子元素层级为2时，则代表返回包含第1级(子元素)和第2级(孙元素)的所有元素；当子元素层级为3时，则代表返回包含第1级(子元素)、第2级(孙元素)及第3级(曾孙元素)的所有元素；当子元素层级为4时，依次类推；当子元素层级超出实际层级范围时，则与至最末层级(即为0)的返回结果一样，即返回其包含所有层级的元素  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: arrElement，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*******************************获取子元素****************************** 命令原型: arrElement = UiElement.GetChildren(objUiElement,level, optionArgs) 入参: objUiElement--目标元素 level--子元素层级 optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参: arrElement--函数调用的输出保存到的变量 注意事项: 无 *********************************************************************/ arrElement = UiElement.GetChildren(@ui"块级元素<div>_百度首页设置登录新闻hao123地图贴吧视频图片网盘更多翻译学术文库百科知道健康",1, {"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) TracePrint(arrElement)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiElement_图片/UiElement_GetChildren.png)  

---

## 获取父元素

**说明**: 通常一个指定的目标元素有且为唯一的父元素，而父元素又有自己的父元素，直至最顶层的父元素(即桌面)，从而可获取指定父元素层级的且为唯一的父元素，返回结果为内存地址对象  

**原型**: `objElement = UiElement.GetParent(objUiElement,upLevels,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 指定被获取父元素的目标元素  
- **upLevels** (True) [number] 默认:1 - 默认父元素层级为1，即为直接父级元素。当父元素层级为2时，则获取指定目标元素的父级元素的父级元素(祖父元素)；当父元素层级为3时，则获取指定目标元素的父级元素的父级元素的父级元素(曾祖父元素)；当父元素层级为4时，依次类推；当父元素层级超出最顶层元素(桌面)或者父元素层级 <=0 时，则获取的父元素为桌面  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: objElement，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/******************************获取父元素******************************* 命令原型: objElement = UiElement.GetParent(objUiElement,upLevels,optionArgs) 入参: objUiElement--目标元素 upLevels--父元素层级 optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参: objElement--函数调用的输出保存到的变量 注意事项: 无 *********************************************************************/ objElement = UiElement.GetParent(@ui"输入控件<input>4",1,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) TracePrint(objElement)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiElement_图片/UiElement_GetParent.png)  

---

## 获取元素区域

**说明**: 获取元素的区域，返回包含元素所在位置的矩形对象  

**原型**: `objRect = UiElement.GetRect(objUiElement,sRelative,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 元素特征字符串  
- **sRelative** (True) [enum] 默认:"parent" - 返回元素位置是相对于哪一个坐标而言的  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: objRect，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*****************************获取元素区域******************************** 命令原型: objRect = UiElement.GetRect(objUiElement,sRelative,optionArgs) 入参: objUiElement--目标元素 sRelative--相对位置 optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参: objRect--函数调用的输出保存到的变量 注意事项: 无 *********************************************************************/ objRect = UiElement.GetRect(@ui"表格单元<td>_自我评价","parent",{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) TracePrint(objRect)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiElement_图片/UiElement_GetRect.png)  

---

## 获取元素选择

**说明**: 获取元素的选择（可以操作 列表框、下拉列表框 等类型的元素）  

**原型**: `arrSelItem = UiElement.GetSelect(objUiElement,sMode,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 元素特征字符串  
- **sMode** (True) [enum] 默认:"text" - 选择方式，传递为 index 时按照索引顺序选择（从0开始），传递为 text 时按照选项文本选择，传递为 value 时，按照选项的 value 属性选择  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: arrSelItem，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/******************************获取元素选择******************************* 命令原型: arrSelItem = UiElement.GetSelect(objUiElement,sMode,optionArgs) 入参: objUiElement--目标元素 sMode--选择方式 optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参: arrSelItem--函数调用的输出保存到的变量 注意事项: 无 *********************************************************************/ arrSelItem = UiElement.GetSelect(@ui"下拉列表<select>_省/市安徽澳门北京福建甘肃广东广西贵州海南河北河南黑龙江湖北湖南吉林江苏江","text",{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) TracePrint(arrSelItem)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiElement_图片/UiElement_GetSelect.png)  

---

## 获取元素文本

**说明**: 获取元素的文本内容（Value属性）  

**原型**: `sRet = UiElement.GetValue(objUiElement,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 元素特征字符串  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/******************************获取元素文本******************************* 命令原型: sRet = UiElement.GetValue(objUiElement,optionArgs) 入参: objUiElement--目标元素 optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参: sRet--函数调用的输出保存到的变量 注意事项: 无 *********************************************************************/ sRet = UiElement.GetValue(@ui"选项<option>_销售",{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiElement_图片/UiElement_GetValue.png)  

---

## 元素截图

**说明**: 对指定元素进行全区域或者局部区域截图  

**原型**: `UiElement.ScreenCapture(sFile,objUiElement,sRect,optionArgs)`  

**参数**:  
- **sFile** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 截图所存储的文件路径  
- **objUiElement** (True) [decorator] 默认:@ui"" - 通过鼠标选取的界面元素，包含窗口、元素等信息  
- **sRect** (True) [dictionary] 默认:{ "x": 0, "y": 0, "width": 0, "height": 0 } - 对指定界面元素截图的范围，如果范围传递为 { "x":0,"y":0,"width":0,"height":0 } ，则截取该元素的全区域，否则以该元素的左上角为坐标原点，根据高宽进行截图  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/*******************************元素截图****************************** 命令原型: UiElement.ScreenCapture(sFile,objUiElement,sRect,optionArgs) 入参: sFile--保存路径 objUiElement--目标元素 sRect--截图范围 optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参： 无 注意事项: 无 *********************************************************************/ UiElement.ScreenCapture(@res&#x27;&#x27;&#x27;1.png&#x27;&#x27;&#x27;,@ui"表格单元<td>_用户名",{"x":0,"y":0,"width":0,"height":0},{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiElement_图片/UiElement_ScreenCapture.png)  

---

## 设置元素属性

**说明**: 设置元素的属性  

**原型**: `UiElement.SetAttribute(objUiElement,sAttribute,sValue,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 元素特征字符串  
- **sAttribute** (True) [string] 默认:"" - 要修改的属性名  
- **sValue** (True) [string] 默认:"" - 要修改的属性值  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/******************************设置元素属性******************************* 命令原型: UiElement.SetAttribute(objUiElement,sAttribute,sValue,optionArgs) 入参: objUiElement--目标元素 sAttribute--属性名 sValue--属性值 optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参： 无 注意事项: 无 *********************************************************************/ UiElement.SetAttribute(@ui"输入控件<input>7","type","test",{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiElement_图片/UiElement_SetAttribute.png)  

---

## 设置元素勾选

**说明**: 勾选元素（可以操作单选框、复选框 等类型的元素）  

**原型**: `UiElement.SetCheck(objUiElement,bCheck,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 元素特征字符串  
- **bCheck** (True) [boolean] 默认:True - 勾选方式，传递为 true 时勾选元素，传递为 false 时取消勾选元素  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/****************************设置元素勾选********************************* 命令原型: UiElement.SetCheck(objUiElement,bCheck,optionArgs) 入参: objUiElement--目标元素 bCheck--是否勾选 optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参: bRet--函数调用的输出保存到的变量 注意事项: 无 *********************************************************************/ UiElement.SetCheck(@ui"复选框<checkboxinput>_on1",true,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiElement_图片/UiElement_SetCheck.png)  

---

## 设置元素选择

**说明**: 选择元素（可以操作 列表框、下拉列表框 等类型的元素）  

**原型**: `UiElement.SetSelect(objUiElement,arrItem,sMode,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 元素特征字符串  
- **arrItem** (True) [expression] 默认:[] - 要选择的元素数组（根据 sMode 属性决定选择的方式）  
- **sMode** (True) [enum] 默认:"text" - 选择方式，传递为 index 时按照索引顺序选择（从0开始），传递为 text 时按照选项文本选择，传递为 value 时，按照选项的 value 属性选择  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/******************************设置元素选择******************************* 命令原型: UiElement.SetSelect(objUiElement,arrItem,sMode,optionArgs) 入参: objUiElement--目标元素 arrItem--包含元素 sMode--选择方式 optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参： 无 注意事项: 无 *********************************************************************/ UiElement.SetSelect(@ui"下拉列表<select>_省/市安徽澳门北京福建甘肃广东广西贵州海南河北河南黑龙江湖北湖南吉林江苏江1",["安徽"],"text",{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiElement_图片/UiElement_SetSelect.png)  

---

## 设置元素文本

**说明**: 设置元素的文本内容  

**原型**: `UiElement.SetValue(objUiElement,sValue,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 元素特征字符串  
- **sValue** (True) [string] 默认:"" - 待写入UI界面元素的文本内容，文本格式为字符串类型  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/*******************************设置元素文本****************************** 命令原型: UiElement.SetValue(objUiElement,sValue,optionArgs) 入参: objUiElement--目标元素 sValue--写入文本 optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参： 无 注意事项: 无 *********************************************************************/ UiElement.SetValue(@ui"输入控件<input>8","test@126.com",{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiElement_图片/UiElement_SetValue.png)  

---

## 等待元素

**说明**: 等待元素显示或消失  

**原型**: `UiElement.Wait(objUiElement,iType,iTimeOut,optionArgs)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应需要操作的界面元素，当属性传递为 字符串 类型时，作为特征串查找界面元素，当属性传递为 UiElement 类型时，直接对 UiElement 对应的界面元素进行点击操作  
- **iType** (True) [enum] 默认:"show" - 等待方式，可以设置为等待元素显示后结束或等待元素消失后结束  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/*****************************等待元素******************************** 命令原型: UiElement.Wait(objUiElement,iType,iTimeOut,optionArgs) 入参: objUiElement--目标元素 iType--等待方式 iTimeOut--超时时间.默认单位:毫秒.Type:Int optionArgs--可选参数(包括:错误继续执行/执行后延时/执行前延时).Type:Dict 出参： 无 注意事项: 无 *********************************************************************/ UiElement.Wait(@ui"表格单元<td>_用户名1","show",10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/UiElement_图片/UiElement_Wait.png)  

---

## 绑定浏览器

**说明**: 绑定一个已经打开的浏览器，使 Laiye RPA 可以对这个浏览器进行操作，绑定的浏览器可以是 IE、Chrome、FireFox、360、Edge，命令运行成功会返回绑定的浏览器句柄字符串，失败返回 null  

**原型**: `hWeb = WebBrowser.BindBrowser(sType,iTimeOut,optionArgs)`  

**参数**:  
- **sType** (True) [enum] 默认:"ie" - 浏览器类型  
- **iTimeOut** (True) [number] 默认:10000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为10000毫秒（10秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: hWeb，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************绑定浏览器*************************************** 命令原型： WebBrowser.BindBrowser(sType,iTimeOut,optionArgs) 入参： sType--浏览器类型:IE,Chrome,FireFox,Edge,360Browser iTimeOut--异常等待时间 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: hWeb--浏览器句柄字符串 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************/ //绑定一个已经打开的浏览器，使 Laiye RPA 可以对这个浏览器进行操作，绑定的浏览器可以是 IE、Chrome、FireFox、360、Edge，命令运行成功会返回绑定的浏览器句柄字符串，失败返回 null Dim hWeb //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //在调试栏输出浏览器对象 TracePrint(hWeb) //绑定已经打开的IE浏览器 hWeb = WebBrowser.BindBrowser("ie",10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //在调试栏输出绑定的浏览器对象 TracePrint(hWeb)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_BindBrowser.png)  

---

## 关闭标签页

**说明**: 关闭标签页  

**原型**: `WebBrowser.Close(hWeb,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/*********************************关闭标签页*************************************** 命令原型： WebBrowser.Close(hWeb,optionArgs) 入参： hWeb--浏览器对象 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: 无 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,bRet //打开IE浏览器进入百度 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //在调试栏输出浏览器对象 TracePrint(hWeb) //绑定浏览器对象 hWeb = WebBrowser.BindBrowser("ie",10000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //在调试栏输出绑定的浏览器对象 TracePrint(hWeb) //关闭指定浏览器对象的当前标签页 WebBrowser.Close(hWeb,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_Close.png)  

---

## 启动新的浏览器

**说明**: 启动一个新的浏览器，使 Laiye RPA 可以对这个浏览器进行操作，启动的浏览器可以是 Internet Explorer、Chrome、FireFox、360、Edge、Laiye RPA浏览器（Laiye RPA浏览器仅支持启动一个浏览器窗口），命令运行成功会返回绑定的浏览器句柄字符串，失败返回 null  

**原型**: `hWeb = WebBrowser.Create(sType,sURL,iTimeOut,optionArgs)`  

**参数**:  
- **sType** (True) [enum] 默认:"ie" - 浏览器类型  
- **sURL** (True) [string] 默认:"about:blank" - 启动浏览器后打开的链接地址  
- **iTimeOut** (True) [number] 默认:30000 - 指定在SelectorNotFoundException引发异常之前等待活动运行的时间量（以毫秒为单位）。默认值为30000毫秒（30秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  
- **sBrowserPath** (False) [path] 默认:"" - 浏览器目录，默认为空字符串。当值为空字符串时，自动查找机器上安装的浏览器并尝试启动  
- **sStartArgs** (False) [string] 默认:"" - 浏览器启动参数  

**返回**: hWeb，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************启动新的浏览器*************************************** 命令原型： WebBrowser.Create(sType,sURL,iTimeOut,optionArgs) 入参： sType--浏览器类型:IE,Chrome,FireFox,Edge,360Browser,Laiye RPA浏览器 sURL--打开浏览器跳转的链接 iTimeOut--异常等待时间 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 sBrowserPath--浏览器安装路径 sStartArgs--浏览器启动参数 出参: hWeb--浏览器句柄字符串 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 使用该命令打开的浏览器如果在后台可以改用启动应用程序命令 *********************************************************************/ //启动一个新的浏览器，使 Laiye RPA 可以对这个浏览器进行操作，启动的浏览器可以是 Internet Explorer、Chrome、FireFox、360、Edge、Laiye RPA浏览器（Laiye RPA浏览器仅支持启动一个浏览器窗口），命令运行成功会返回绑定的浏览器句柄字符串，失败返回 null Dim hWeb //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //在调试栏输出浏览器对象 TracePrint(hWeb)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_Create.png)  

---

## 下载文件

**说明**: 利用浏览器下载指定链接的文件  

**原型**: `WebBrowser.Download(hWeb,sURL,sFile,bSync,iTimeOut,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **sURL** (True) [string] 默认:"" - 要下载的文件链接地址（URL）  
- **sFile** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 下载的文件在本机保存的路径  
- **bSync** (True) [boolean] 默认:True - 命令是否同步执行，传递为 true 则等待文件下载完成后才返回继续执行，传递为 false 则文件开始下载后立即返回  
- **iTimeOut** (True) [number] 默认:300000 - 等待文件下载的超时时间，超过这个时间则判定为文件下载失败，默认为 300000 毫秒（5分钟）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/*********************************下载文件*************************************** 命令原型： WebBrowser.Download(hWeb,sURL,sFile,bSync,iTimeOut,optionArgs) 入参： hWeb--浏览器对象 sURL--要下载的文件链接地址（URL） sFile--下载的文件在本机保存的路径 bSync--命令是否同步执行 iTimeOut--等待文件下载的超时时间 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: 无 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,bRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //在调试栏输出浏览器对象 TracePrint(hWeb) //打开配置的url iRet = WebBrowser.GoURL(hWeb,"https://laiye.com/download?source=product-process-creator-banner",false,{},30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) TracePrint(iRet) //下载url包含的资源保存到指定路径 WebBrowser.Download(hWeb,"https://down.uibot.com.cn/onekernel/6.0.0/UiBot_Community_Official_X86_V6.0.0_2021.12.15.2018.exe",&#x27;&#x27;&#x27;C:\Users\Downloads&#x27;&#x27;&#x27;,true,60000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_Download.png)  

---

## 读取网页Cookies

**说明**: 读取网页的 Cookies 数据  

**原型**: `dictRet = WebBrowser.GetCookies(hWeb,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: dictRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取Cookies*************************************** 命令原型： WebBrowser.GetCookies(hWeb,optionArgs) 入参： hWeb--浏览器对象 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: dictRet--命令运行的返回结果 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,bRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //在调试栏输出浏览器对象 TracePrint(hWeb) //获取网页的cookies dictRet = WebBrowser.GetCookies(hWeb,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //输出cookies到调试栏 TracePrint(dictRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_GetCookies.png)  

---

## 读取网页源码

**说明**: 读取当前页面的网页源代码（HTML），读取的代码和网页源文件有区别，如果网页是JS构建的，则读取的代码包含了渲染后的完整HTML结构树  

**原型**: `sRet = WebBrowser.GetHTML(hWeb,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************读取网页源码*************************************** 命令原型： WebBrowser.GetHTML(hWeb,optionArgs) 入参： hWeb--浏览器对象 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: sRet--将命令运行后的结果赋值给此变量 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,bRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //在调试栏输出浏览器对象 TracePrint(hWeb) //打开URL iRet = WebBrowser.GoURL(hWeb,"https://laiye.com/download?source=product-process-creator-banner",false,{},30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //获取网页源码 sRet = WebBrowser.GetHTML(hWeb,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //打印源码 TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_GetHTML.png)  

---

## 获取滚动条位置

**说明**: 获取当前页面滚动条的位置（像素）  

**原型**: `dictScrollPostion = WebBrowser.GetScroll(hWeb,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: dictScrollPostion，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取滚动条位置*************************************** 命令原型： WebBrowser.GetScroll(hWeb,optionArgs) 入参： hWeb--浏览器对象 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: dictScrollPostion--命令运行返回的结果 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,bRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //在调试栏输出浏览器对象 TracePrint(hWeb) //获取滚动条位置(包含横向和纵向滚动条) bRet = WebBrowser.GetScroll(hWeb,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //打印结果 TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_GetScroll.png)  

---

## 获取网页标题

**说明**: 获取当前页面的网页标题  

**原型**: `sRet = WebBrowser.GetTitle(hWeb,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取网页标题*************************************** 命令原型： WebBrowser.GetTitle(hWeb,optionArgs) 入参： hWeb--浏览器对象 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: sRet--命令运行后的返回值 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,sRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //获取当前标签页的标题 sRet = WebBrowser.GetTitle(hWeb,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //在调试栏输出结果 TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_GetTitle.png)  

---

## 获取网页URL

**说明**: 获取当前页面的链接地址（URL）  

**原型**: `sRet = WebBrowser.GetURL(hWeb,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取网页URL*************************************** 命令原型： WebBrowser.GetURL(hWeb,optionArgs) 入参： hWeb--浏览器对象 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: sRet--命令运行后的返回值 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,bRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //获取浏览器当前标签页的URL sRet = WebBrowser.GetURL(hWeb,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //在调试栏输出结果 TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_GetURL.png)  

---

## 后退

**说明**: 执行浏览器的后退操作（与工具栏的后退按钮功能相同）  

**原型**: `WebBrowser.GoBack(hWeb,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/*********************************后退*************************************** 命令原型： WebBrowser.GoBack(hWeb,optionArgs) 入参： hWeb--浏览器对象 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参： 无 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,bRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //执行回退 WebBrowser.GoBack(hWeb,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_GoBack.png)  

---

## 前进

**说明**: 执行浏览器的前进操作（与工具栏的前进按钮功能相同）  

**原型**: `WebBrowser.GoForward(hWeb,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/*********************************前进*************************************** 命令原型： GoForward(hWeb,optionArgs) 入参： hWeb--浏览器对象 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参： 无 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,bRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //执行回退 WebBrowser.GoBack(hWeb,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //执行前进 WebBrowser.GoForward(hWeb,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_GoForward.png)  

---

## 打开网页

**说明**: 控制浏览器加载指定链接（URL）  

**原型**: `iRet = WebBrowser.GoURL(hWeb,sURL,bWait,arrElement,iTimeOut,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **sURL** (True) [string] 默认:"" - 要加载的网页链接地址（URL）  
- **bWait** (True) [boolean] 默认:True - 是否等待网页加载完毕后命令才返回，传递为 true 则必须等页面加载完成或加载失败时才会继续操作，传递为 false 则开始加载页面后立刻返回  
- **arrElement** (True) [decorator] 默认:{ } - 当页面加载完后，判断是否存在指定的元素，不填写则不进行任何元素判断  
- **iTimeOut** (True) [number] 默认:30000 - 等待页面加载的超时时间，超过这个时间则判定为网页加载失败，默认为 30000 毫秒（30秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************打开网页*************************************** 命令原型： WebBrowser.GoURL(hWeb,sURL,bWait,arrElement,iTimeOut,optionArgs) 入参： hWeb--浏览器对象 sURL--要加载的网页链接地址 bWait--是否等待网页加载完毕后命令才返回 arrElement--判断是否存在指定的元素 iTimeOut--等待页面加载的超时时间 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: iRet--命令运行返回结果 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,iRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //加载指定的url并返回指定元素是否出现 iRet = WebBrowser.GoURL(hWeb,"weibo.com",true,{},30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //输出结果到调试栏 TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_GoURL.png)  

---

## 获取运行状态

**说明**: 获取浏览器的运行状态，浏览器还在运行时返回 true，浏览器已经退出时返回 false  

**原型**: `bRet = WebBrowser.IsRunning(hWeb,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************获取运行状态*************************************** 命令原型： WebBrowser.IsRunning(hWeb,optionArgs) 入参： hWeb--浏览器对象 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: bRet--命令运行后的结果 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,bRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //获取浏览器状态 bRet = WebBrowser.IsRunning(hWeb,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //输出运行结果 TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_IsRunning.png)  

---

## 刷新

**说明**: 刷新当前页面（与工具栏的刷新按钮功能相同）  

**原型**: `WebBrowser.Refresh(hWeb,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/*********************************刷新*************************************** 命令原型： WebBrowser.Refresh(hWeb,optionArgs) 入参： hWeb--浏览器对象 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参： 无 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,bRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //刷新浏览器 WebBrowser.Refresh(hWeb,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_Refresh.png)  

---

## 执行JS

**说明**: 执行JS，返回JS执行结果（字符串格式）  

**原型**: `sRet = WebBrowser.RunJS(hWeb,sScript,bSync,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **sScript** (True) [string] 默认:"function(){return 123}" - 要执行的JS脚本内容  
- **bSync** (True) [boolean] 默认:True - 是否同步执行，传递为 true 则等待JS运行完成后才返回继续执行，传递为 false 则JS开始执行立即返回  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************执行JS*************************************** 命令原型： WebBrowser.RunJS(hWeb,sScript,bSync,optionArgs) 入参： hWeb--浏览器对象 sScript--JS脚本内容 bSync--是否同步执行 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: sRet --命令运行返回的结果 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb, sRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) TracePrint(hWeb) //运行js代码,返回js代码的返回值 sRet = WebBrowser.RunJS(hWeb,"function(){return document.getElementsByClassName(&#x27;bg s_btn&#x27;)}",true,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //在调试栏打印结果 TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_RunJS.png)  

---

## 浏览器截图

**说明**: 浏览器截图  

**原型**: `WebBrowser.ScreenShot(hWeb,sPath,objRect,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **sPath** (True) [path] 默认:&#x27;&#x27;&#x27;C:\Users&#x27;&#x27;&#x27; - 截图保存到的文件路径  
- **objRect** (True) [dictionary] 默认:{ "x": 0, "y": 0, "width": 0, "height": 0 } - 截图的矩形范围，传递为 null 则截取整个浏览器的显示区域  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/*********************************浏览器截图*************************************** 命令原型： WebBrowser.ScreenShot(hWeb,sPath,objRect,optionArgs) 入参： hWeb--浏览器对象 sPath--截图保存到的文件路径 objRect--截图的矩形范围，传递为 null 则截取整个浏览器的显示区域 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: 无 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,bRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) TracePrint(hWeb) //浏览器截图保存到指定位置 WebBrowser.ScreenShot(hWeb,&#x27;&#x27;&#x27;C:\Users\DVA\OneDrive\Pictures\屏幕快照\dddd.png&#x27;&#x27;&#x27;,{"x": 500, "y": 500, "width": 0, "height": 0},{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_ScreenShot.png)  

---

## 设置网页Cookies

**说明**: 设置网页的 Cookies 数据  

**原型**: `WebBrowser.SetCookies(hWeb,dictCookies,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **dictCookies** (True) [expression] 默认:{ } - 一个或多个Cookies名称和Cookies值配对，组成的JSON对象  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/*********************************设置Cookies*************************************** 命令原型： WebBrowser.SetCookies(hWeb,dictCookies,optionArgs) 入参： hWeb--浏览器对象 dictCookies--一个或多个Cookies名称和Cookies值配对，组成的JSON对象 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: 无 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,iRet,vCookies //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //获取网站cookies vCookies = WebBrowser.GetCookies(hWeb,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //设置cookies到指定浏览器 WebBrowser.SetCookies(hWeb, vCookies ,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_SetCookies.png)  

---

## 设置滚动条位置

**说明**: 设置当前页面滚动条的位置（像素）  

**原型**: `WebBrowser.SetScroll(hWeb,dictScrollPostion,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **dictScrollPostion** (True) [expression] 默认:{ "ScrollLeft": 0,"ScrollTop": 0 } - 滚动条移动到的新位置，ScrollLeft元素表示横轴滚动条的位置，ScrollTop元素表示纵轴滚动条的位置  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/*********************************设置滚动条位置*************************************** 命令原型： SetScroll(hWeb,dictScrollPostion,optionArgs) 入参： hWeb--浏览器对象 dictScrollPostion--滚动条移动到的新位置 bContinueOnError--错误或是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: 无 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //点击滚动条 Mouse.Action(@ui"块级元素<div>_","left","click",10000,{"bContinueOnError": false, "iDelayAfter": 300, "iDelayBefore": 200, "bSetForeground": true, "sCursorPosition": "Center", "iCursorOffsetX": 0, "iCursorOffsetY": 0, "sKeyModifiers": [],"sSimulate": "simulate", "bMoveSmoothly": false}) //设置纵向滚动条到600像素位置 WebBrowser.SetScroll(hWeb,{"ScrollLeft": 0,"ScrollTop": 600},{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_SetScroll.png)  

---

## 停止加载页面

**说明**: 停止加载当前页面（与工具栏的停止按钮功能相同）  

**原型**: `WebBrowser.Stop(hWeb,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**示例**:  
```
/*********************************停止加载网页*************************************** 命令原型： WebBrowser.Stop(hWeb,optionArgs) 入参： hWeb--浏览器对象 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: 无 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,bRet,iRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //加载微博的url iRet = WebBrowser.GoURL(hWeb,"weibo.com",true,{},30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //停止加载 WebBrowser.Stop(hWeb,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200})
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_Stop.png)  

---

## 切换标签页

**说明**: 切换浏览器标签页，可通过地址栏、标题栏进行匹配，支持&#x27;*&#x27;通配符  

**原型**: `bRet = WebBrowser.SwitchTab(hWeb,sType,sContent,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **sType** (True) [enum] 默认:"title" - 匹配对象  
- **sContent** (True) [string] 默认:"" - 匹配内容  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: bRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************切换标签页*************************************** 命令原型： WebBrowser.SwitchTab(hWeb,sType,sContent,optionArgs) 入参： hWeb--浏览器对象 sType--匹配对象:标题栏,地址栏 sContent--匹配内容 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: bRet--命令运行后的结果 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用,需要全字匹配,不支持模糊匹配 *********************************************************************************/ Dim hWeb,bRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //加载微博的url iRet = WebBrowser.GoURL(hWeb,"weibo.com",true,{},30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //使用url匹配的方式切换到百度首页 bRet = WebBrowser.SwitchTab(hWeb,"url","www.baidu.com",{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //在调试栏输出返回结果 TracePrint(bRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_SwitchTab.png)  

---

## 等待网页加载

**说明**: 等待当前加载的页面加载完成  

**原型**: `iRet = WebBrowser.WaitPage(hWeb,arrElement,iTimeOut,optionArgs)`  

**参数**:  
- **hWeb** (True) [expression] 默认:hWeb - 使用 WebBrowser.Create 或 WebBrowser.Bind 命令返回的浏览器句柄字符串  
- **arrElement** (True) [decorator] 默认:{ } - 当页面加载完后，判断是否存在指定的元素，不填写则不进行任何元素判断并返回0；另支持传入元素数组来判断多个元素是否都存在，都存在则返回1，若任意一个元素不存在时则返回0  
- **iTimeOut** (True) [number] 默认:60000 - 等待页面加载的超时时间，超过这个时间则判定为网页加载失败，默认为 60000 毫秒（60秒）  
- **bContinueOnError** (False) [boolean] 默认:False - 指定即使活动引发错误，自动化是否仍应继续。该字段仅支持布尔值（True，False）。默认值为False  
- **iDelayAfter** (False) [number] 默认:300 - 执行活动后的延迟时间（以毫秒为单位）。默认时间为300毫秒  
- **iDelayBefore** (False) [number] 默认:200 - 活动开始执行任何操作之前的延迟时间（以毫秒为单位）。默认的时间量是200毫秒  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*********************************等待网页加载*************************************** 命令原型： WebBrowser.WaitPage(hWeb,arrElement,iTimeOut,optionArgs) 入参： hWeb--浏览器对象 bWait--是否等待网页加载完毕后命令才返回 arrElement--判断是否存在指定的元素 iTimeOut--等待页面加载的超时时间 bContinueOnError--错误后是否继续 iDelayAfter--执行后延时 iDelayBefore--执行前延时 出参: iRet--命令运行返回结果 注意事项： 使用谷歌内核的浏览器首次使用需要安装扩展并启用扩展 需要打开浏览器后才能使用 *********************************************************************************/ Dim hWeb,iRet //启动IE浏览器并打开百度首页 hWeb = WebBrowser.Create("ie","www.baidu.com",30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200,"sBrowserPath":"","sStartArgs":""}) //加载微博url iRet = WebBrowser.GoURL(hWeb,"weibo.com",true,{},30000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //等待页面加载,并返回指定的元素是否存在 iRet = WebBrowser.WaitPage(hWeb,{"wnd":[{"cls":"IEFrame","title":"*","app":"iexplore"},{"cls":"Internet Explorer_Server"}],"html":[{"aaname":"热门微博","tag":"SPAN"}]},60000,{"bContinueOnError":false,"iDelayAfter":300,"iDelayBefore":200}) //在调试栏输出返回值(0,1) TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/WebBrowser_图片/WebBrowser_WaitPage.png)  

---

## 关闭窗口

**说明**: 关闭一个窗口  

**原型**: `Window.Close(objUiElement)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应的窗口对象选择器，传递选择器时作为窗口特征使用，会查找符合的窗口进行操作  

**示例**:  
```
/*******************************关闭窗口****************************** 命令原型: Window.Close(objUiElement) 入参: objUiElement--目标元素 出参： 无 注意事项: 无 *********************************************************************/ Window.Close(@ui"窗口_www.vrbrothers.com/cn/wqm/demo/pages/Demo-Compl1")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Window_图片/Window_Close.png)  

---

## 判断窗口是否存在

**说明**: 判断窗口是否存在  

**原型**: `bRet = Window.Exists(objUiElement)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应的窗口对象，传递为字符串时作为窗口特征使用，会查找所有符合的窗口进行操作；传递为UiElement对象时，则对这个对象所属的窗口进行操作  
- **### 返回结果** () [] 默认: -   
- **bRet，将命令运行后的结果赋值给此变量。** () [] 默认: -   
- **### 运行实例** () [] 默认: -   

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Window_图片/Window_Exists.png)  

---

## 获取活动窗口

**说明**: 获取活动窗口（处于前台被激活的窗口），返回的值为该窗口句柄  

**原型**: `objUiElement = Window.GetActive()`  

**参数**:  

**返回**: objUiElement，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*******************************获取活动窗口****************************** 命令原型: objUiElement = Window.GetActive() 入参： 无 出参: objUiElement--函数调用的输出保存到的变量 注意事项: 无 *********************************************************************/ objUiElement = Window.GetActive()
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Window_图片/Window_GetActive.png)  

---

## 获取窗口类名

**说明**: 获取窗口类名  

**原型**: `sRet = Window.GetClass(objUiElement)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应的窗口对象，传递为字符串时作为窗口特征使用，会查找所有符合的窗口进行操作；传递为UiElement对象时，则对这个对象所属的窗口进行操作  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*******************************获取窗口类名****************************** 命令原型: sRet = Window.GetClass(objUiElement) 入参: objUiElement--目标元素 出参: sRet--函数调用的输出保存到的变量 注意事项: 无 *********************************************************************/ sRet = Window.GetClass(@ui"窗口_百度一下，你就知道-用户配置1-MicrosoftEdge8") TracePrint(sRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Window_图片/Window_GetClass.png)  

---

## 获取文件路径

**说明**: 获取窗口对应程序的可执行文件路径  

**原型**: `sRet = Window.GetPath(objUiElement)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应的窗口对象，传递为字符串时作为窗口特征使用，会查找所有符合的窗口进行操作；传递为UiElement对象时，则对这个对象所属的窗口进行操作  

**返回**: sRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*******************************获取文件路径****************************** 命令原型: sRet = Window.GetPath(objUiElement) 入参: objUiElement--目标元素 出参: sRet--函数调用的输出保存到的变量 注意事项: 无 *********************************************************************/ sRet = Window.GetPath(@ui"窗口_百度一下，你就知道-用户配置1-MicrosoftEdge9")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Window_图片/Window_GetPath.png)  

---

## 获取进程PID

**说明**: 获取窗口对应程序的运行PID  

**原型**: `iRet = Window.GetPID(objUiElement)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应的窗口对象，传递为字符串时作为窗口特征使用，会查找所有符合的窗口进行操作；传递为UiElement对象时，则对这个对象所属的窗口进行操作  

**返回**: iRet，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*******************************获取进程PID****************************** 命令原型: iRet = Window.GetPID(objUiElement) 入参: objUiElement--目标元素 出参: iRet--函数调用的输出保存到的变量 注意事项: 无 *********************************************************************/ iRet = Window.GetPID(@ui"窗口_百度一下，你就知道-用户配置1-MicrosoftEdge10") TracePrint(iRet)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Window_图片/Window_GetPID.png)  

---

## 获取窗口大小

**说明**: 获取窗口大小（像素）  

**原型**: `objRect = Window.GetSize(objUiElement)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应的窗口对象，传递为字符串时作为窗口特征使用，会查找所有符合的窗口进行操作；传递为UiElement对象时，则对这个对象所属的窗口进行操作  

**返回**: objRect，将命令运行后的结果赋值给此变量。  

**示例**:  
```
/*******************************获取窗口大小****************************** 命令原型: objRect = Window.GetSize(objUiElement) 入参: objUiElement--目标元素 出参: objRect--函数调用的输出保存到的变量 注意事项: 无 *********************************************************************/ objRect = Window.GetSize(@ui"窗口_百度一下，你就知道-用户配置1-MicrosoftEdge4") TracePrint(objRect)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Window_图片/Window_GetSize.png)  

---

## 移动窗口位置

**说明**: 将窗口移动到新的屏幕坐标位置（像素）  

**原型**: `Window.Move(objUiElement,x,y)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应的窗口对象，传递为字符串时作为窗口特征使用，会查找所有符合的窗口进行操作；传递为UiElement对象时，则对这个对象所属的窗口进行操作  
- **x** (True) [number] 默认:0 - 移动到新位置的横坐标  
- **y** (True) [number] 默认:0 - 移动到新位置的纵坐标  

**示例**:  
```
/*******************************移动窗口位置****************************** 命令原型: Window.Move(objUiElement,x,y) 入参: objUiElement--目标元素 x--横坐标 y--纵坐标 出参： 无 注意事项: 无 *********************************************************************/ Window.Move(@ui"窗口_百度一下，你就知道-用户配置1-MicrosoftEdge6",500,700)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Window_图片/Window_Move.png)  

---

## 设置活动窗口

**说明**: 将指定窗口设置为活动状态（处于前台被激活的窗口）  

**原型**: `Window.SetActive(objUiElement)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应的窗口对象，传递为字符串时作为窗口特征使用，会查找所有符合的窗口进行操作；传递为UiElement对象时，则对这个对象所属的窗口进行操作  

**示例**:  
```
/*******************************设置活动窗口****************************** 命令原型: Window.SetActive(objUiElement) 入参: objUiElement--目标元素 出参： 无 注意事项: 无 *********************************************************************/ Window.SetActive(@ui"窗口_百度一下，你就知道-用户配置1-MicrosoftEdge")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Window_图片/Window_SetActive.png)  

---

## 改变窗口大小

**说明**: 改变窗口大小（像素）  

**原型**: `Window.SetSize(objUiElement,w,h)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应的窗口对象，传递为字符串时作为窗口特征使用，会查找所有符合的窗口进行操作；传递为UiElement对象时，则对这个对象所属的窗口进行操作  
- **w** (True) [number] 默认:800 - 窗口宽度  
- **h** (True) [number] 默认:600 - 窗口高度  

**示例**:  
```
/*******************************改变窗口大小****************************** 命令原型: Window.SetSize(objUiElement,w,h) 入参: objUiElement--目标元素 w--宽度 h--高度 出参： 无 注意事项: 无 *********************************************************************/ Window.SetSize(@ui"窗口_百度一下，你就知道-用户配置1-MicrosoftEdge5",800,600)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Window_图片/Window_SetSize.png)  

---

## 更改窗口显示状态

**说明**: 更改窗口显示状态  

**原型**: `Window.Show(objUiElement,sShow)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应的窗口对象，传递为字符串时作为窗口特征使用，会查找所有符合的窗口进行操作；传递为UiElement对象时，则对这个对象所属的窗口进行操作  
- **sShow** (True) [enum] 默认:"show" - 窗口显示状态，&#x27;show&#x27; 为显示；&#x27;hide&#x27; 为隐藏；&#x27;min&#x27; 为最小化；&#x27;max&#x27; 为最大化；&#x27;restore&#x27; 为还原  

**示例**:  
```
/*******************************更改窗口显示状态****************************** 命令原型: Window.Show(objUiElement,sShow) 入参: objUiElement--目标元素 sShow--显示状态 出参： 无 注意事项: 无 *********************************************************************/ Window.Show(@ui"窗口_百度一下，你就知道-用户配置1-MicrosoftEdge2","show")
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Window_图片/Window_Show.png)  

---

## 窗口置顶

**说明**: 使窗口始终显示在最顶层或取消置顶操作  

**原型**: `Window.TopMost(objUiElement,bTopMost)`  

**参数**:  
- **objUiElement** (True) [decorator] 默认:@ui"" - 对应的窗口对象，传递为字符串时作为窗口特征使用，会查找所有符合的窗口进行操作；传递为UiElement对象时，则对这个对象所属的窗口进行操作  
- **bTopMost** (True) [boolean] 默认:True - 是否使窗口置顶，传递为 true 则使对应的窗口置顶，传递为 false 则使对应的窗口取消置顶  

**示例**:  
```
/*******************************窗口置顶****************************** 命令原型: Window.TopMost(objUiElement,bTopMost) 入参: objUiElement--目标元素 bTopMost--是否置顶 出参： 无 注意事项: 无 *********************************************************************/ Window.TopMost(@ui"窗口_百度一下，你就知道-用户配置1-MicrosoftEdge7",true)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Window_图片/Window_TopMost.png)  

---

## 退格键删除

**说明**: 对 Word 文档当前选中的内容执行退格键删除操作  

**原型**: `Word.Backspace(objWord)`  

**参数**:  
- **objWord** (True) [expression] 默认:objWord - Word 文档对象  

**示例**:  
```
/*******************************退格键删除********************************** 命令原型： Word.Backspace(objWord) 入参： objWord--Word文档对象 出参： 无 注意事项： 需要打开word后才能使用 ****************************************************************************/ Dim objWord //打开Word文件 objWord = Word.Open(&#x27;&#x27;&#x27;C:\Users\Administrator\Desktop\标准化执行注意事项.docx&#x27;&#x27;&#x27;,"xw4131221","",true) //设置光标位置并选中"一定录缺陷"文字内容 Word.SetTextPosition(objWord,"一定录缺陷",0) //在当前光标位置执行退格删除 Word.Backspace(objWord)
```  

**图示**: ![](https://docs-res.laiye.com/production/docs-res/rpa-command-manual/zh/v6.0.0/Word_图片/Word_Backspace.png)  

---
