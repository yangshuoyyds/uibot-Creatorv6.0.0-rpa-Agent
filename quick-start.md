# UIBot 快速入门指南

5分钟快速上手 UIBot RPA 开发，从零开始创建你的第一个自动化流程。

## 📋 目录

- [环境准备](#环境准备)
- [第一个流程](#第一个流程)
- [常用操作](#常用操作)
- [进阶学习](#进阶学习)

---

## 环境准备

### 1. 下载安装

**下载地址**: https://laiye.com/download

**推荐版本**: UIBot 6.0 社区版（永久免费）

**系统要求**:
- Windows 7/8/10/11
- .NET Framework 4.5+
- 2GB+ 内存

### 2. 注册登录

1. 启动 UIBot Creator（流程创造者）
2. 使用手机号或邮箱注册账号
3. 登录后即可使用

---

## 第一个流程

### 场景：自动打开记事本并输入文字

#### 步骤1：创建新流程

```
1. 打开 UIBot Creator
2. 点击"新建流程"
3. 输入流程名称：我的第一个流程
4. 点击"确定"
```

#### 步骤2：添加命令

**方法1：使用流程图模式（推荐新手）**

```
1. 在左侧命令面板找到"系统操作" → "启动应用程序"
2. 拖拽到流程图中
3. 双击命令，设置参数：
   - 应用程序路径：C:\Windows\notepad.exe
   - 等待方式：等待应用程序准备好
4. 点击"确定"
```

```
5. 继续添加"键盘输入" → "输入文本"
6. 设置参数：
   - 目标元素：点击"捕获"按钮，选择记事本窗口
   - 输入内容：Hello UIBot!
7. 点击"确定"
```

**方法2：使用代码模式**

```vb
' 启动记事本
App.Run("C:\Windows\notepad.exe", 1, 1)
Delay(2000)

' 输入文字
Keyboard.InputText(@ui"记事本编辑框", "Hello UIBot!", True, False)
```

#### 步骤3：运行流程

```
1. 点击工具栏的"运行"按钮（或按 F5）
2. 观察流程自动执行
3. 看到记事本打开并输入文字
```

**恭喜！你已经完成了第一个 RPA 流程！** 🎉

---

## 常用操作

### 1. 网页自动化

**场景**：自动打开百度并搜索

```vb
' 打开浏览器
Dim objBrowser = WebBrowser.Create("chrome", "https://www.baidu.com", 30)
Delay(2000)

' 输入搜索关键词
Keyboard.InputText(@ui"百度搜索框", "UIBot RPA", True, False)
Delay(500)

' 点击搜索按钮
Mouse.Click(@ui"百度一下按钮", "left", "single", 0, 0)
Delay(3000)

' 关闭浏览器
WebBrowser.Close(objBrowser)
```

**关键点**：
- 使用元素捕获工具定位搜索框和按钮
- 添加适当的延时等待页面加载

---

### 2. Excel 操作

**场景**：读取 Excel 数据并处理

```vb
' 打开 Excel 文件
Dim objExcel = Excel.Open("C:\data.xlsx", True, "")

' 读取单元格
Dim sValue = Excel.GetCell(objExcel, "A1")
TracePrint("读取到: " & sValue)

' 写入单元格
Excel.SetCell(objExcel, "B1", "已处理")

' 保存并关闭
Excel.Save(objExcel)
Excel.Close(objExcel)
```

**关键点**：
- 使用绝对路径
- 记得保存和关闭文件

---

### 3. 条件判断

**场景**：根据条件执行不同操作

```vb
Dim iAge = 25

If iAge >= 18 Then
    TracePrint("成年人")
Else
    TracePrint("未成年人")
End If
```

---

### 4. 循环处理

**场景**：批量处理数据

```vb
' For 循环
For i = 1 To 10
    TracePrint("第 " & i & " 次循环")
Next

' For Each 循环
Dim arrNames = ["张三", "李四", "王五"]
For Each sName In arrNames
    TracePrint("姓名: " & sName)
Next
```

---

### 5. 异常处理

**场景**：捕获和处理错误

```vb
Try
    ' 可能出错的代码
    Mouse.Click(@ui"按钮", "left", "single", 0, 0)
    TracePrint("点击成功")
Catch ex
    ' 错误处理
    TracePrint("点击失败: " & ex.Message)
End Try
```

---

## 进阶学习

### 学习路径

```
第1周：基础操作
├── 界面元素定位
├── 鼠标键盘操作
└── 简单流程编写

第2周：软件自动化
├── 浏览器自动化
├── Excel 自动化
└── 文件操作

第3周：逻辑控制
├── 条件判断
├── 循环处理
└── 异常处理

第4周：企业级开发
├── 配置管理
├── 日志记录
└── 流程设计模式
```

### 推荐文档

| 阶段 | 推荐文档 | 说明 |
|------|---------|------|
| 入门 | [developer-guide-index.md](developer-guide-index.md) | 官方开发者指南索引 |
| 速查 | [quick-index.md](quick-index.md) | 按功能快速查找命令 |
| 实战 | [examples.md](examples.md) | 15+ 个完整案例 |
| 模板 | [templates.md](templates.md) | 11 个即用代码模板 |
| 问题 | [faq.md](faq.md) | 23+ 个常见问题解答 |
| 企业 | [enterprise-best-practices.md](enterprise-best-practices.md) | 企业级开发指南 |

---

## 💡 实用技巧

### 1. 元素定位技巧

**问题**：元素定位不稳定

**解决**：
```vb
' 使用智能等待
If UiElement.Exists(@ui"按钮", 10) Then
    Mouse.Click(@ui"按钮")
Else
    TracePrint("元素未找到")
End If
```

---

### 2. 调试技巧

**使用 TracePrint 输出调试信息**：
```vb
Dim sValue = "测试"
TracePrint("当前值: " & sValue)
```

**使用断点**：
- 在代码行号处点击，设置断点
- 按 F5 运行，程序会在断点处暂停
- 按 F10 单步执行

---

### 3. 性能优化

**减少不必要的延时**：
```vb
' 不好的做法
Delay(5000)

' 好的做法
If UiElement.Exists(@ui"元素", 10) Then
    ' 继续操作
End If
```

---

## 🎯 实战练习

### 练习1：自动登录网站

**任务**：
1. 打开登录页面
2. 输入用户名和密码
3. 点击登录按钮
4. 验证登录成功

**提示**：参考 [examples.md](examples.md) 中的"示例1：自动登录网站"

---

### 练习2：Excel 数据处理

**任务**：
1. 读取 Excel 文件
2. 遍历每一行数据
3. 对数据进行处理
4. 将结果写回 Excel

**提示**：参考 [templates.md](templates.md) 中的"Excel处理模板"

---

### 练习3：网页数据采集

**任务**：
1. 打开目标网页
2. 提取页面数据
3. 保存到 Excel

**提示**：参考 [examples.md](examples.md) 中的"示例3：网页数据采集"

---

## 🔗 相关资源

### 官方资源
- **官方网站**: https://www.laiye.com
- **开发者指南**: https://documents.laiye.com/rpa-guide/docs/
- **命令手册**: https://documents.laiye.com/rpa-command-manual/docs/
- **社区论坛**: https://forum.laiye.com

### 本 Skill 文档
- [README.md](README.md) - Skill 总览
- [commands-reference.md](commands-reference.md) - 命令详细参考
- [quick-index.md](quick-index.md) - 快速索引
- [templates.md](templates.md) - 代码模板库
- [faq.md](faq.md) - 常见问题
- [enterprise-best-practices.md](enterprise-best-practices.md) - 企业级最佳实践
- [design-patterns.md](design-patterns.md) - 流程设计模式

---

## ❓ 常见问题

### Q1: 如何捕获界面元素？

**答**：
1. 点击命令中的"捕获"按钮
2. 按住 Ctrl 键，鼠标移动到目标元素
3. 元素会高亮显示
4. 点击鼠标左键完成捕获

---

### Q2: 为什么元素找不到？

**答**：
- 检查元素是否已加载完成
- 使用智能等待：`UiElement.Exists(@ui"元素", 10)`
- 尝试使用不同的定位方式（XPath、属性等）
- 参考 [faq.md](faq.md) 的 Q1

---

### Q3: 如何处理动态网页？

**答**：
- 使用 `WebBrowser.RunJS()` 执行 JavaScript
- 等待特定元素出现
- 使用 AJAX 完成事件监听
- 参考 [examples.md](examples.md) 的"示例4：BOSS直聘数据采集"

---

## 🎓 下一步

完成快速入门后，建议：

1. **深入学习**：阅读 [developer-guide-index.md](developer-guide-index.md)
2. **实战练习**：参考 [examples.md](examples.md) 完成更多案例
3. **企业级开发**：学习 [enterprise-best-practices.md](enterprise-best-practices.md)
4. **加入社区**：https://forum.laiye.com 与其他开发者交流

---

**文档版本**: v1.3.0  
**更新时间**: 2024-01-15  
**适用版本**: UIBot 6.0
