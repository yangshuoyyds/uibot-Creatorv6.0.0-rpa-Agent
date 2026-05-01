# 🚀 UIBot 工具快速使用指南

## 一分钟上手

### Windows 用户

```cmd
cd tools
uibot-tools.bat
```

### Linux/Mac 用户

```bash
cd tools
./uibot-tools.sh
```

---

## 四大工具速查

### 1️⃣ 搜索命令

```bash
python search.py -i
>>> search 点击
>>> func 如何点击按钮
```

### 2️⃣ 生成代码

```bash
python generator.py -i
>>> web_automation
```

### 3️⃣ 查询文档

```bash
python query.py -i
>>> 选择功能 1-7
```

### 4️⃣ 验证代码

```bash
python validator.py your_code.task
```

---

## 常见场景

### 场景 1: 我想开发网页自动化

```bash
# 1. 搜索命令
python search.py -f "打开浏览器"

# 2. 生成代码
python generator.py -t web_automation -i

# 3. 验证代码
python validator.py output.task
```

### 场景 2: 我想处理 Excel

```bash
# 1. 查看示例
python query.py -s Excel -t example

# 2. 生成代码
python generator.py -t excel_processing -i

# 3. 验证代码
python validator.py output.task
```

### 场景 3: 我遇到了问题

```bash
# 1. 搜索 FAQ
python query.py -s "你的问题" -t faq

# 2. 查看示例
python query.py -s "你的问题" -t example
```

---

## 可用模板

| 模板 | 说明 | 使用场景 |
|------|------|---------|
| web_automation | 网页自动化 | 打开网页、点击、输入 |
| excel_processing | Excel 处理 | 读取、处理、保存 Excel |
| file_batch | 文件批处理 | 批量处理文件 |
| data_collection | 数据采集 | 循环采集网页数据 |
| reframework | 企业级框架 | 企业级流程开发 |
| error_handling | 错误处理 | 完整的错误处理框架 |

---

## 帮助命令

```bash
python search.py --help
python generator.py --help
python query.py --help
python validator.py --help
```

---

## 详细文档

- [工具文档](tools/README.md)
- [项目文档](README.md)
- [升级说明](UPGRADE-v1.4.0.md)

---

**需要帮助？查看 [tools/README.md](tools/README.md)**
