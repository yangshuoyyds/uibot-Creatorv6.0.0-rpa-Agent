# 剑气高级案例快速入门指南

> 5分钟快速上手剑气企业级RPA项目开发

## 📋 目录

- [快速开始](#快速开始)
- [使用代码生成器](#使用代码生成器)
- [学习案例代码](#学习案例代码)
- [常见场景实现](#常见场景实现)
- [最佳实践检查清单](#最佳实践检查清单)

---

## 快速开始

### 1. 查看案例分析文档

```bash
# 打开剑气高级案例分析文档
code jianqi-advanced-cases.md
```

**文档包含**：
- ✅ 完整的项目架构分析
- ✅ 7大核心技术要点
- ✅ 4种可复用代码模式
- ✅ 最佳实践总结
- ✅ 常见问题解决方案

### 2. 浏览真实案例代码

```bash
# 查看案例目录结构
ls 剑气高级案例/北京玖卓科技_张建琦_15801023818/

# 查看主流程文件
code 剑气高级案例/北京玖卓科技_张建琦_15801023818/2.编码/机票查询_人机交互/机票查询_人机交互.flow
```

**案例特点**：
- ✅ 完整的企业级项目结构
- ✅ 多数据源采集与对比
- ✅ 配置化设计
- ✅ 完善的异常处理
- ✅ 邮件通知功能

---

## 使用代码生成器

### 生成多数据源采集模块

```bash
# 交互模式（推荐新手）
python tools/jianqi-generator.py -i

# 命令行模式
python tools/jianqi-generator.py -t multi_source \
  -p '{"author":"张三","source_name":"南航","url_var":"csAddr","data_var":"csData"}' \
  -o 南航模块.task
```

**生成的代码包含**：
- ✅ 浏览器控制
- ✅ 元素定位与操作
- ✅ 数据抓取
- ✅ 数据格式化
- ✅ 完整的异常处理

### 生成数据对比模块

```bash
python tools/jianqi-generator.py -t data_compare \
  -p '{"author":"张三"}' \
  -o 数据对比模块.task
```

**功能特点**：
- ✅ 多数据源对比
- ✅ 价格智能选择
- ✅ 处理独有数据
- ✅ Excel结果输出

### 生成完整项目结构

```bash
python tools/jianqi-generator.py -t full_project \
  -p '{"project_name":"机票查询","author":"张三"}' \
  -o MyProject
```

**生成文件**：
- `MyProject_flow` - 主流程文件（JSON格式）
- `MyProject_config_ini` - 配置文件
- `MyProject_readme` - 项目说明文档

### 生成公共函数库

```bash
python tools/jianqi-generator.py -t public_block \
  -p '{"author":"张三"}' \
  -o PublicBlock.task
```

**包含函数**：
- `InitArgByLocal()` - 配置初始化
- `ErrCapture()` - 异常处理
- `GetDate()` - 日期处理
- `SendMailSMTP()` - 邮件发送
- `ExecuteWithRetry()` - 重试机制
- `ValidateData()` - 数据验证

---

## 学习案例代码

### 核心代码片段

#### 1. 智能价格对比

```vb
// 从案例中学习：数据对比模块.task
For Each csValue In csData
    For Each xcValue In xcData
        If xcValue[0] = '南方航空'
            If StartsWith(xcValue[1], csValue[2])
                // 选择更低价格
                If CInt(xcValue[2]) > CInt(csValue[5])
                    price = csValue[5]
                    source = '南航'
                Else
                    price = xcValue[2]
                    source = '携程'
                End If
            End If
        End If
    Next
Next
```

**学习要点**：
- 双层循环遍历数据
- 条件匹配（航班号）
- 价格对比逻辑
- 数据来源记录

#### 2. 配置化设计

```vb
// 从案例中学习：PublicBlock.task
Function InitArgByLocal()
    // 从INI文件读取配置
    logLevel = INI.Read(@res"config\\Config.ini", "参数值", "LogLevel", "2")
    g_dictGlobal["departure"] = INI.Read(@res"config\\Config.ini", 
                                         "城市", "出发地", "广州")
    g_dictGlobal["arrival"] = INI.Read(@res"config\\Config.ini", 
                                       "城市", "到达地", "北京")
End Function
```

**学习要点**：
- INI文件结构
- 配置参数读取
- 默认值设置
- 全局变量管理

#### 3. 异常处理机制

```vb
// 从案例中学习：三层异常处理
// 1. 全局异常标识
g_dictGlobal["isEx"] = False

// 2. 模块级Try-Catch
Try
    // 业务逻辑
Catch Ex
    PublicBlock.ErrCapture("操作失败:", Ex)
    Return
End Try

// 3. 流程级判断
If g_dictGlobal["isEx"] = True
    Log.Info("流程出错")
    Return
End If
```

**学习要点**：
- 全局异常标识
- Try-Catch捕获
- 统一异常处理函数
- 流程级异常判断

---

## 常见场景实现

### 场景1：多平台数据采集

**需求**：从3个网站采集数据并对比

**实现步骤**：

1. **生成数据源1采集模块**
```bash
python tools/jianqi-generator.py -t multi_source \
  -p '{"source_name":"平台A","url_var":"urlA","data_var":"dataA"}' \
  -o 平台A模块.task
```

2. **生成数据源2采集模块**
```bash
python tools/jianqi-generator.py -t multi_source \
  -p '{"source_name":"平台B","url_var":"urlB","data_var":"dataB"}' \
  -o 平台B模块.task
```

3. **生成数据对比模块**
```bash
python tools/jianqi-generator.py -t data_compare -o 数据对比模块.task
```

4. **配置Config.ini**
```ini
[地址]
urlA=https://platformA.com
urlB=https://platformB.com

[业务参数]
关键词=查询关键词
```

### 场景2：定时数据监控

**需求**：每天定时采集数据并发送邮件报告

**实现步骤**：

1. **生成完整项目**
```bash
python tools/jianqi-generator.py -t full_project \
  -p '{"project_name":"数据监控","author":"你的名字"}' \
  -o DataMonitor
```

2. **配置邮件参数**
```ini
[邮件]
服务器=smtp.qq.com
端口=25
发件人=your_email@qq.com
密码=授权码
收件人=receiver@example.com
```

3. **设置定时任务**（Windows）
```cmd
# 使用任务计划程序
# 每天早上9点运行
```

### 场景3：数据质量检查

**需求**：采集数据后进行验证和清洗

**实现步骤**：

1. **使用公共函数库的验证函数**
```vb
// 生成的PublicBlock.task中包含ValidateData函数
Dim rawData = g_dictGlobal["rawData"]
Dim validData = PublicBlock.ValidateData(rawData)
g_dictGlobal["cleanData"] = validData
```

2. **自定义验证规则**
```vb
Function ValidateData(data)
    Dim validData = []
    For Each item In data
        // 必填字段检查
        If item[0] <> "" And item[1] <> ""
            // 数据类型检查
            Try
                price = CInt(item[2])
                // 数据范围检查
                If price > 0 And price < 100000
                    validData = push(validData, item)
                End If
            Catch Ex
                Log.Warning("数据格式错误: " & CStr(item))
            End Try
        End If
    Next
    Return validData
End Function
```

---

## 最佳实践检查清单

### 项目结构 ✅

- [ ] 按标准结构组织文件（设计/编码/测试/上线）
- [ ] 使用res/config目录存放配置和资源
- [ ] 创建log目录存放日志
- [ ] 使用PublicBlock.task存放公共函数

### 代码规范 ✅

- [ ] 全局变量使用g_前缀
- [ ] 函数使用帕斯卡命名法
- [ ] 添加完整的注释说明
- [ ] 每个模块开始和结束使用TracePrint

### 配置管理 ✅

- [ ] 所有参数从配置文件读取
- [ ] 不在代码中硬编码敏感信息
- [ ] 提供默认值
- [ ] 配置文件结构清晰

### 异常处理 ✅

- [ ] 使用全局异常标识
- [ ] 关键操作使用Try-Catch
- [ ] 统一的异常处理函数
- [ ] 每个模块开始前检查异常状态

### 日志记录 ✅

- [ ] 设置合适的日志级别
- [ ] 记录关键步骤
- [ ] 错误信息详细清晰
- [ ] 使用TracePrint辅助调试

### 性能优化 ✅

- [ ] 及时关闭浏览器和Excel
- [ ] 使用智能等待而非固定延迟
- [ ] 避免重复操作
- [ ] 合理使用缓存

### 安全性 ✅

- [ ] 敏感信息从配置文件读取
- [ ] 验证用户输入
- [ ] 使用授权码而非明文密码
- [ ] 注意文件权限

---

## 快速参考

### 常用命令

```bash
# 查看所有可用模板
python tools/jianqi-generator.py -h

# 交互式生成代码
python tools/jianqi-generator.py -i

# 查看案例文档
code jianqi-advanced-cases.md

# 浏览案例代码
code 剑气高级案例/北京玖卓科技_张建琦_15801023818/2.编码/
```

### 学习路径

1. **第1天**：阅读案例分析文档，理解项目架构
2. **第2天**：浏览真实案例代码，学习核心技术
3. **第3天**：使用代码生成器创建简单模块
4. **第4天**：实现多数据源采集场景
5. **第5天**：完成完整项目并测试

### 获取帮助

- 📖 查看 [jianqi-advanced-cases.md](jianqi-advanced-cases.md) 了解详细技术分析
- 📖 查看 [README.md](README.md) 了解完整功能
- 📖 查看案例代码获取实战经验
- 🔧 使用代码生成器快速开发

---

**文档版本**：v1.0  
**更新时间**：2026-05-02  
**适用版本**：剑气 5.3.0+
