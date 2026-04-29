# UIBot 开发者指南索引

本文档基于来也科技官方开发者指南，提供完整的学习路径和知识索引。

## 📚 官方文档

**官方地址**: https://documents.laiye.com/rpa-guide/docs/

**适用版本**: UIBot 6.0 社区版

---

## 📖 开发者指南（第一辑）

### 1. RPA简介
**链接**: https://documents.laiye.com/rpa-guide/docs/

**核心内容**:
- RPA 基础知识
- RPA 平台介绍
- 来也科技发展历程
- RPA 和 AI 的结合

**适合人群**: RPA 初学者、了解 RPA 概念

---

### 2. 基本概念
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD1/Chapt2-BasicConcepts

**核心内容**:
- 流程创造者（Creator）基本概念
- 流程、任务、命令的关系
- 可视化编程 vs 代码编程
- 变量和数据类型
- 流程图的基本元素

**适合人群**: 开始使用 UIBot 的新手

**关键知识点**:
- ✅ 流程图模式：拖拽式可视化编程
- ✅ 代码模式：BotScript 脚本语言
- ✅ 混合模式：流程图 + 代码结合

---

### 3. 界面元素自动化
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD1/Chapt3-Target

**核心内容**:
- 元素识别和定位
- 元素选择器
- 鼠标和键盘操作
- 元素属性获取
- 等待元素出现

**适合人群**: 需要操作界面元素的开发者

**关键技术**:
- ✅ 元素捕获工具
- ✅ XPath 定位
- ✅ 属性定位
- ✅ 图像定位
- ✅ OCR 文字定位

**实战场景**:
- 网页表单填写
- 桌面应用操作
- 按钮点击
- 文本输入

---

### 4. 界面图像自动化
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD1/Chapt4-NoTarget

**核心内容**:
- 图像识别技术
- 屏幕截图和对比
- 图像点击
- OCR 文字识别
- 图像等待

**适合人群**: 处理无法用元素定位的场景

**关键技术**:
- ✅ 图像匹配算法
- ✅ 模板匹配
- ✅ 颜色识别
- ✅ OCR 引擎

**实战场景**:
- 游戏自动化
- 远程桌面操作
- 图片验证码识别
- 无障碍界面操作

---

### 5. 软件自动化
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD1/Chapt5-SoftwareAutomation

**核心内容**:
- 浏览器自动化（Chrome、IE、Edge）
- Excel 自动化
- Word 自动化
- Outlook 邮件自动化
- SAP 自动化
- 数据库操作

**适合人群**: 需要操作特定软件的开发者

**关键技术**:
- ✅ WebDriver 技术
- ✅ COM 组件调用
- ✅ ODBC 数据库连接
- ✅ SAP GUI Scripting

**实战场景**:
- 网页数据采集
- Excel 报表生成
- 邮件批量发送
- ERP 系统操作

---

### 6. 逻辑控制
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD1/Chapt6-LogicControl

**核心内容**:
- 条件判断（If-Else）
- 循环控制（For、While、Do-While）
- 异常处理（Try-Catch）
- 流程跳转
- 子流程调用

**适合人群**: 需要实现复杂业务逻辑的开发者

**关键技术**:
- ✅ 条件表达式
- ✅ 循环遍历
- ✅ 异常捕获
- ✅ 流程模块化

**实战场景**:
- 批量数据处理
- 条件分支处理
- 错误重试机制
- 流程模块化设计

---

### 7. 流程和任务管理
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD1/Chapt7-Worker

**核心内容**:
- 流程机器人（Worker）使用
- 机器人指挥官（Commander）使用
- 流程部署和发布
- 流程调度和监控
- 日志查看和分析

**适合人群**: 需要部署和管理流程的运维人员

**关键技术**:
- ✅ 流程打包
- ✅ 远程部署
- ✅ 定时调度
- ✅ 集中监控

**实战场景**:
- 企业级流程部署
- 多机器人协作
- 流程监控告警
- 运行日志分析

---

### 8. 结束语
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD1/Chapt9-End

**核心内容**:
- 学习总结
- 进阶方向
- 社区资源

---

### 9. 附录：编程基础知识
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD1/Chapt8-Prepare

**核心内容**:
- 变量和数据类型
- 运算符
- 条件判断
- 循环
- 函数

**适合人群**: 无编程基础的初学者

---

## 📖 开发者指南（第二辑）

### 1. 预备知识
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD2/Chapt0-Preliminary

**核心内容**:
- BotScript 语言基础
- 高级语法特性
- 调试技巧

---

### 2. 数据获取和处理
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD2/Chapt3-DataHandling

**核心内容**:
- 字符串处理
- 数组和字典操作
- 正则表达式
- JSON 和 XML 处理
- 文件读写
- Excel 高级操作

**关键技术**:
- ✅ 字符串分割、替换、查找
- ✅ 数组遍历、排序、过滤
- ✅ 正则匹配和提取
- ✅ JSON 解析和生成
- ✅ CSV 文件处理

**实战场景**:
- 数据清洗
- 文本提取
- 数据转换
- 报表生成

---

### 3. 网络和系统操作
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD2/Chapt4-NetworkAndSystem

**核心内容**:
- HTTP 请求（GET、POST）
- API 调用
- FTP 文件传输
- 系统命令执行
- 进程管理
- 剪贴板操作

**关键技术**:
- ✅ RESTful API 调用
- ✅ HTTP 头和参数设置
- ✅ 文件上传下载
- ✅ CMD 命令执行

**实战场景**:
- API 数据获取
- 文件批量上传
- 系统自动化运维
- 跨系统数据交互

---

### 4. 多流程协作
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD2/MultipleFlows

**核心内容**:
- 主流程和子流程
- 流程间参数传递
- 流程间通信
- 并行流程执行

**关键技术**:
- ✅ 子流程调用
- ✅ 输入输出参数
- ✅ 全局变量共享
- ✅ 流程同步机制

**实战场景**:
- 复杂流程拆分
- 流程模块化
- 并行任务处理
- 流程复用

---

### 5. 人工智能功能
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD2/Chapt5-AI

**核心内容**:
- OCR 文字识别
- 票据识别（发票、身份证等）
- 表格识别
- 文本分类
- 信息抽取
- NLP 自然语言处理

**关键技术**:
- ✅ 智能文档处理（IDP）
- ✅ 通用 OCR
- ✅ 票据结构化识别
- ✅ 文本分析

**实战场景**:
- 发票自动录入
- 身份证信息提取
- 合同信息抽取
- 邮件意图识别

---

### 6. BotScript 语言参考
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD2/Chapt1-LanguageRef

**核心内容**:
- BotScript 完整语法
- 内置函数参考
- 高级特性
- 最佳实践

**适合人群**: 需要深入使用代码模式的开发者

---

### 7. 高级开发功能
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD2/Chapt7-AdvancedFeatures

**核心内容**:
- 自定义命令
- 插件开发
- SDK 集成
- 性能优化
- 调试技巧

**关键技术**:
- ✅ Python 扩展
- ✅ .NET 扩展
- ✅ 命令封装
- ✅ 性能分析

**实战场景**:
- 自定义业务命令
- 第三方库集成
- 性能瓶颈优化
- 复杂调试场景

---

### 8. 扩展命令
**链接**: https://documents.laiye.com/rpa-guide/docs/DevGuideD2/Chapt8-CommandExtend

**核心内容**:
- 命令扩展机制
- 自定义命令开发
- 命令库管理
- 命令分享

---

## 🎯 学习路径推荐

### 初级路径（1-2周）
```
RPA简介 → 基本概念 → 界面元素自动化 → 软件自动化（浏览器/Excel）
```
**目标**: 能够完成简单的自动化任务

### 中级路径（2-4周）
```
逻辑控制 → 数据获取和处理 → 网络和系统操作 → 异常处理
```
**目标**: 能够开发复杂的业务流程

### 高级路径（1-2个月）
```
多流程协作 → 人工智能功能 → 高级开发功能 → 企业级最佳实践
```
**目标**: 能够设计企业级 RPA 解决方案

---

## 📝 配套资源

### 本 Skill 相关文档
- [commands-reference.md](commands-reference.md) - 命令详细参考
- [examples.md](examples.md) - 实战示例集
- [quick-index.md](quick-index.md) - 快速索引
- [templates.md](templates.md) - 代码模板库
- [faq.md](faq.md) - 常见问题解答
- [enterprise-best-practices.md](enterprise-best-practices.md) - 企业级最佳实践
- [design-patterns.md](design-patterns.md) - 流程设计模式

### 官方资源
- **官方网站**: https://www.laiye.com
- **下载地址**: https://laiye.com/download
- **社区论坛**: https://forum.laiye.com
- **视频教程**: https://www.laiye.com/video

---

## 💡 使用建议

### 1. 按需学习
- 根据实际项目需求选择章节
- 不必按顺序全部学习
- 边学边练，理论结合实践

### 2. 结合本 Skill
- 开发者指南：理论和概念
- 本 Skill 文档：实战和速查
- 两者结合，事半功倍

### 3. 实践为主
- 每学一个知识点，立即动手实践
- 参考 examples.md 中的案例
- 使用 templates.md 中的模板快速开发

### 4. 遇到问题
- 先查 faq.md 常见问题
- 再查官方文档详细说明
- 最后到社区论坛求助

---

## 🔗 快速跳转

| 需求 | 推荐文档 |
|------|---------|
| 了解 RPA 概念 | [RPA简介](https://documents.laiye.com/rpa-guide/docs/) |
| 学习基础操作 | [基本概念](https://documents.laiye.com/rpa-guide/docs/DevGuideD1/Chapt2-BasicConcepts) |
| 操作界面元素 | [界面元素自动化](https://documents.laiye.com/rpa-guide/docs/DevGuideD1/Chapt3-Target) |
| 操作浏览器/Excel | [软件自动化](https://documents.laiye.com/rpa-guide/docs/DevGuideD1/Chapt5-SoftwareAutomation) |
| 实现业务逻辑 | [逻辑控制](https://documents.laiye.com/rpa-guide/docs/DevGuideD1/Chapt6-LogicControl) |
| 处理数据 | [数据获取和处理](https://documents.laiye.com/rpa-guide/docs/DevGuideD2/Chapt3-DataHandling) |
| 调用 API | [网络和系统操作](https://documents.laiye.com/rpa-guide/docs/DevGuideD2/Chapt4-NetworkAndSystem) |
| 使用 AI 功能 | [人工智能功能](https://documents.laiye.com/rpa-guide/docs/DevGuideD2/Chapt5-AI) |
| 企业级开发 | [enterprise-best-practices.md](enterprise-best-practices.md) |
| 快速查命令 | [quick-index.md](quick-index.md) |
| 复制代码模板 | [templates.md](templates.md) |

---

**文档版本**: v1.3.0  
**更新时间**: 2024-01-15  
**官方文档版本**: v1.0.0
