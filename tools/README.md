# UIBot 辅助工具集

这是 UIBot Skill 的配套工具集，提供命令搜索、代码生成、文档查询和代码验证等功能。

## 📦 工具列表

### 1. search.py - 命令搜索工具
快速搜索 UIBot 命令，支持模糊匹配和功能描述搜索。

**功能特性**:
- ✅ 关键词搜索
- ✅ 功能描述搜索
- ✅ 分类筛选
- ✅ 正则表达式支持
- ✅ 交互式搜索模式

**使用方法**:

```bash
# 交互模式（推荐）
python tools/search.py -i

# 命令行搜索
python tools/search.py 点击
python tools/search.py -f "如何点击按钮"
python tools/search.py -c mouse 鼠标

# 正则搜索
python tools/search.py -r "Click|Mouse"
```

**交互模式命令**:
```
>>> search 点击        # 搜索包含"点击"的命令
>>> func 如何点击按钮  # 根据功能描述搜索
>>> list mouse         # 列出鼠标相关命令
>>> help               # 显示帮助
>>> quit               # 退出
```

---

### 2. generator.py - 代码生成器
根据模板快速生成 UIBot 代码。

**功能特性**:
- ✅ 6 种常用模板
- ✅ 参数化配置
- ✅ 自动生成时间戳
- ✅ 交互式生成
- ✅ 直接保存到文件

**可用模板**:
1. `web_automation` - 网页自动化基础模板
2. `excel_processing` - Excel 数据处理模板
3. `file_batch` - 文件批量处理模板
4. `data_collection` - 数据采集模板
5. `reframework` - REFramework 企业级模板
6. `error_handling` - 完整错误处理模板

**使用方法**:

```bash
# 交互模式（推荐）
python tools/generator.py -i

# 列出所有模板
python tools/generator.py -l

# 命令行生成
python tools/generator.py -t web_automation -p '{"url":"https://example.com","element_selector":"<button>","action":"点击"}' -o output.task

# 生成 REFramework 模板
python tools/generator.py -t reframework -p '{"process_name":"数据处理流程","config_file":"config.xlsx"}' -o main.task
```

**交互模式示例**:
```
请选择模板 (输入模板 key):
>>> web_automation

生成 网页自动化基础模板
请输入参数:
  url: https://www.baidu.com
  element_selector: <input name='wd' />
  action: 输入

是否保存到文件? (y/n): y
输出文件名: baidu_search.task
✓ 代码已保存到: baidu_search.task
```

---

### 3. query.py - 交互式查询工具
提供友好的文档查询界面。

**功能特性**:
- ✅ 多文档搜索
- ✅ 分类查询
- ✅ 全文搜索
- ✅ 快速索引
- ✅ 企业级指南

**使用方法**:

```bash
# 交互模式（推荐）
python tools/query.py -i

# 命令行搜索
python tools/query.py -s 点击 -t command
python tools/query.py -s Excel -t example
python tools/query.py -s 元素定位 -t faq
python tools/query.py -s 网页 -t template
```

**交互模式功能**:
```
功能菜单:
  1. 搜索命令
  2. 查看示例
  3. 查看常见问题
  4. 查看代码模板
  5. 快速索引
  6. 企业级最佳实践
  7. 全文搜索
  0. 退出
```

---

### 4. validator.py - 代码验证工具
检查 UIBot 代码质量和潜在问题。

**功能特性**:
- ✅ 语法检查
- ✅ 最佳实践检查
- ✅ 性能问题检测
- ✅ 错误处理检查
- ✅ 安全问题扫描

**检查项目**:
- 语法错误（未闭合引号、变量命名等）
- 最佳实践（硬编码路径、固定延迟等）
- 性能问题（循环中的重复操作等）
- 错误处理（缺少 Try-Catch 等）
- 安全问题（明文密码、SQL 注入等）

**使用方法**:

```bash
# 验证单个文件
python tools/validator.py main.task

# 验证目录下所有文件
python tools/validator.py -d ./project

# 递归验证子目录
python tools/validator.py -d ./project -r
```

**输出示例**:
```
UIBot 代码验证报告
================================================================================

文件: main.task

统计: 1 个错误, 3 个警告, 2 个建议

❌ 错误:
  行 15: [security] 可能存在明文密码
    建议: 建议使用配置文件或加密存储

⚠️  警告:
  行 8: [best_practice] 使用了硬编码路径
    建议: 建议使用配置文件或相对路径

💡 建议:
  行 20: [best_practice] 使用了固定延迟
    建议: 建议使用 UiElement.Wait() 等智能等待方法
```

---

## 🚀 快速开始

### 安装依赖

所有工具都是纯 Python 3 编写，无需额外依赖。

```bash
# 确保 Python 3.6+ 已安装
python --version

# 进入工具目录
cd tools
```

### 推荐工作流

#### 1. 查找命令
```bash
# 使用搜索工具找到需要的命令
python search.py -i
>>> func 如何点击网页按钮
```

#### 2. 生成代码
```bash
# 使用生成器创建代码框架
python generator.py -i
>>> web_automation
```

#### 3. 查询文档
```bash
# 查看详细文档和示例
python query.py -i
>>> 选择功能 2 (查看示例)
```

#### 4. 验证代码
```bash
# 验证代码质量
python validator.py your_code.task
```

---

## 📝 使用示例

### 示例 1: 创建网页自动化流程

```bash
# 1. 搜索相关命令
python search.py -f "打开浏览器"
# 找到: WebBrowser.Create, WebBrowser.Navigate

# 2. 生成代码模板
python generator.py -t web_automation -p '{"url":"https://www.baidu.com","element_selector":"<input name=\"wd\" />","action":"输入"}' -o baidu.task

# 3. 验证代码
python validator.py baidu.task
```

### 示例 2: Excel 数据处理

```bash
# 1. 查看 Excel 示例
python query.py -s Excel -t example

# 2. 生成 Excel 处理代码
python generator.py -t excel_processing -p '{"input_file":"data.xlsx","output_file":"result.xlsx","sheet_name":"Sheet1"}' -o excel_process.task

# 3. 验证代码
python validator.py excel_process.task
```

### 示例 3: 企业级流程开发

```bash
# 1. 查看企业级最佳实践
python query.py -i
>>> 选择功能 6 (企业级最佳实践)

# 2. 生成 REFramework 模板
python generator.py -t reframework -p '{"process_name":"订单处理流程","config_file":"config.xlsx"}' -o main.task

# 3. 验证代码
python validator.py main.task
```

---

## 🔧 高级用法

### 批量验证项目

```bash
# 验证整个项目
python validator.py -d ../企业级流程模板 -r
```

### 自定义模板

编辑 `generator.py`，在 `load_templates()` 方法中添加自定义模板：

```python
self.templates['my_template'] = {
    'name': '我的自定义模板',
    'description': '模板描述',
    'params': ['param1', 'param2'],
    'template': '''
    // 模板内容
    // 使用 {param1} 和 {param2} 作为占位符
    '''
}
```

### 扩展搜索功能

编辑 `search.py`，在 `search_by_function()` 方法中添加关键词映射：

```python
keywords_map = {
    '自定义功能': ['keyword1', 'keyword2'],
    # ... 其他映射
}
```

---

## 💡 使用技巧

### 1. 组合使用工具

```bash
# 搜索 → 生成 → 验证 一条龙
python search.py -f "网页自动化" && \
python generator.py -t web_automation -i && \
python validator.py output.task
```

### 2. 创建别名（Linux/Mac）

在 `~/.bashrc` 或 `~/.zshrc` 中添加：

```bash
alias uibot-search='python /path/to/tools/search.py -i'
alias uibot-gen='python /path/to/tools/generator.py -i'
alias uibot-query='python /path/to/tools/query.py -i'
alias uibot-check='python /path/to/tools/validator.py'
```

### 3. Windows 批处理脚本

创建 `uibot.bat`:

```batch
@echo off
if "%1"=="search" python tools\search.py -i
if "%1"=="gen" python tools\generator.py -i
if "%1"=="query" python tools\query.py -i
if "%1"=="check" python tools\validator.py %2
```

使用：
```cmd
uibot search
uibot gen
uibot check main.task
```

---

## 🐛 故障排除

### 问题 1: 找不到文档文件

**错误**: `⚠ 未找到 xxx.md`

**解决**: 确保工具在正确的目录运行，或使用绝对路径。

### 问题 2: 中文乱码

**解决**: 确保终端支持 UTF-8 编码。

Windows CMD:
```cmd
chcp 65001
```

### 问题 3: Python 版本问题

**解决**: 确保使用 Python 3.6+

```bash
python --version
# 或
python3 --version
```

---

## 📚 相关文档

- [README.md](../README.md) - 项目总览
- [quick-start.md](../quick-start.md) - 快速入门
- [commands-reference.md](../commands-reference.md) - 命令参考
- [templates.md](../templates.md) - 代码模板库
- [faq.md](../faq.md) - 常见问题

---

## 🔄 更新日志

### v1.0.0 (2024-01-15)
- ✅ 初始版本发布
- ✅ 命令搜索工具
- ✅ 代码生成器
- ✅ 交互式查询工具
- ✅ 代码验证工具

---

## 📞 反馈与建议

如果您在使用工具时遇到问题或有改进建议，欢迎反馈！

---

**制作时间**: 2024-01-15  
**适用版本**: UIBot 6.0+  
**Python 版本**: 3.6+
