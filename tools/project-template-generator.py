#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
RPA项目模板生成器
基于UiBot实施方法指导V3.0的标准项目结构生成工具

功能：
1. 生成完整的RPA项目目录结构
2. 生成各阶段标准文档模板
3. 生成企业级框架代码
4. 生成配置文件和README

作者：UB-Skill Team
版本：1.0.0
日期：2026-05-02
"""

import os
import json
from datetime import datetime


class RPAProjectGenerator:
    """RPA项目模板生成器"""

    def __init__(self):
        self.templates = {
            'rpa_project': self.generate_rpa_project,
            'requirement_doc': self.generate_requirement_doc,
            'design_doc': self.generate_design_doc,
            'test_doc': self.generate_test_doc,
            'deployment_doc': self.generate_deployment_doc,
            'operation_doc': self.generate_operation_doc
        }

    def generate_rpa_project(self, params):
        """生成完整RPA项目结构"""
        project_name = params.get('project_name', 'MyRPAProject')
        author = params.get('author', 'RPA开发者')

        structure = {
            'structure': self._generate_project_structure(project_name),
            'readme': self._generate_project_readme(project_name, author),
            'config': self._generate_project_config(project_name),
            'requirement': self._generate_requirement_template(project_name),
            'design': self._generate_design_template(project_name),
            'code': self._generate_code_template(project_name, author),
            'test': self._generate_test_template(project_name),
            'deployment': self._generate_deployment_template(project_name),
            'operation': self._generate_operation_template(project_name)
        }

        return structure

    def _generate_project_structure(self, project_name):
        """生成项目目录结构"""
        structure = f'''
{project_name}/
├── 1、需求/
│   ├── {project_name}需求说明书.md
│   ├── {project_name}流程需求说明.md
│   └── {project_name}需求变更表.xlsx
├── 2、设计/
│   ├── {project_name}整体设计.md
│   ├── {project_name}流程设计.md
│   └── {project_name}流程图.vsdx
├── 3、编码/
│   ├── {project_name}.flow
│   ├── 初始化.task
│   ├── 业务处理.task
│   ├── 异常处理.task
│   ├── PublicBlock.task
│   ├── res/
│   │   ├── config/
│   │   │   ├── Config.ini
│   │   │   └── Config.xlsx
│   │   ├── data/
│   │   └── images/
│   └── log/
├── 4、测试/
│   ├── {project_name}单元测试.xlsx
│   ├── {project_name}系统集成测试.xlsx
│   └── {project_name}UAT测试.xlsx
├── 5、上线/
│   ├── {project_name}部署清单.xlsx
│   └── {project_name}用户手册.md
├── 6、运维/
│   ├── {project_name}运维手册.md
│   └── {project_name}问题处理记录.xlsx
└── README.md
'''
        return structure

    def _generate_project_readme(self, project_name, author):
        """生成项目README"""
        readme = f'''# {project_name}

## 项目概述

本项目是基于UiBot平台开发的RPA自动化流程。

## 项目信息

- **项目名称**：{project_name}
- **开发者**：{author}
- **创建时间**：{datetime.now().strftime('%Y-%m-%d')}
- **UiBot版本**：6.0+

## 项目结构

```
{project_name}/
├── 1、需求/          # 需求文档
├── 2、设计/          # 设计文档
├── 3、编码/          # 代码实现
├── 4、测试/          # 测试文档
├── 5、上线/          # 上线文档
└── 6、运维/          # 运维文档
```

## 快速开始

### 1. 环境准备

- 安装UiBot Creator 6.0+
- 配置目标应用系统
- 准备测试数据

### 2. 配置修改

修改 `3、编码/res/config/Config.ini` 文件：

```ini
[参数值]
LogLevel=2
maxRetryNum=3

[业务参数]
# 根据实际业务修改
```

### 3. 运行流程

1. 打开UiBot Creator
2. 打开 `3、编码/{project_name}.flow`
3. 点击运行按钮

## 功能说明

### 主要功能

1. **初始化模块**：环境初始化、配置加载
2. **业务处理模块**：核心业务逻辑
3. **异常处理模块**：异常捕获和处理
4. **公共函数库**：通用函数封装

### 配置说明

**Config.ini 配置项**：

- `LogLevel`：日志级别（0-错误，1-警告，2-信息）
- `maxRetryNum`：最大重试次数
- 其他业务参数根据实际需求配置

## 开发规范

### 命名规范

- 全局变量：`g_` 前缀
- 局部变量：驼峰命名
- 函数名：帕斯卡命名

### 注释规范

- 每个模块开头添加作者、时间、功能说明
- 关键步骤添加注释
- 复杂逻辑添加详细说明

## 测试说明

### 测试环境

- 操作系统：Windows 10+
- UiBot版本：6.0+
- 目标应用：[填写]

### 测试用例

参见 `4、测试/` 目录下的测试文档。

## 部署说明

### 部署前检查

- [ ] 环境检查
- [ ] 配置检查
- [ ] 权限检查
- [ ] 备份检查

### 部署步骤

参见 `5、上线/{project_name}部署清单.xlsx`

## 运维说明

### 监控指标

- 执行成功率
- 平均执行时长
- 异常次数
- 资源使用率

### 问题处理

参见 `6、运维/{project_name}运维手册.md`

## 版本历史

### v1.0.0 ({datetime.now().strftime('%Y-%m-%d')})

- 初始版本发布

## 联系方式

- 开发者：{author}
- 邮箱：[填写]
- 电话：[填写]

## 许可说明

本项目仅供内部使用。

---

**创建时间**：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
**生成工具**：RPA项目模板生成器 v1.0.0
'''
        return readme

    def _generate_project_config(self, project_name):
        """生成配置文件"""
        config = f'''[参数值]
LogLevel=2
maxRetryNum=3

[业务参数]
项目名称={project_name}
执行模式=自动
超时时间=30000

[邮件配置]
服务器=smtp.qq.com
端口=25
发件人=your_email@qq.com
密码=your_password
收件人=receiver@example.com

[数据库配置]
# 如需要数据库，请配置以下参数
服务器=localhost
端口=3306
数据库名=
用户名=
密码=

[系统配置]
日志路径=./log
数据路径=./res/data
配置路径=./res/config
'''
        return config

    def _generate_requirement_template(self, project_name):
        """生成需求文档模板"""
        template = f'''# {project_name} 需求说明书

## 1. 项目概述

### 1.1 项目背景

[描述项目的业务背景和目的]

### 1.2 项目目标

- 目标1：[描述]
- 目标2：[描述]
- 目标3：[描述]

### 1.3 项目范围

**包含范围**：
- [描述]

**不包含范围**：
- [描述]

## 2. 业务流程现状

### 2.1 流程描述

[详细描述当前的业务流程]

### 2.2 流程步骤

1. 步骤1：[描述]
2. 步骤2：[描述]
3. 步骤3：[描述]

### 2.3 痛点分析

| 痛点 | 影响 | 频率 |
|------|------|------|
| [痛点1] | [影响描述] | [高/中/低] |
| [痛点2] | [影响描述] | [高/中/低] |

## 3. 自动化需求

### 3.1 功能需求

#### FR001: [功能名称]

**需求描述**：[详细描述]

**输入**：
- [输入1]
- [输入2]

**处理**：
- [处理步骤1]
- [处理步骤2]

**输出**：
- [输出1]
- [输出2]

**优先级**：[高/中/低]

### 3.2 非功能需求

#### NFR001: 性能需求

- 单次执行时长：< [X]分钟
- 并发执行数：[X]个
- 成功率：> [X]%

#### NFR002: 可靠性需求

- 异常处理：完善的异常处理机制
- 重试机制：支持自动重试
- 日志记录：详细的日志记录

#### NFR003: 安全性需求

- 数据加密：敏感数据加密存储
- 权限控制：严格的权限控制
- 审计日志：完整的审计日志

## 4. 约束条件

### 4.1 技术约束

- 开发平台：UiBot 6.0+
- 操作系统：Windows 10+
- 目标应用：[应用名称及版本]

### 4.2 业务约束

- 执行时间：[时间段]
- 数据量：[数据量范围]
- 并发限制：[并发数]

## 5. 验收标准

### 5.1 功能验收

- [ ] 所有功能需求实现
- [ ] 异常处理完善
- [ ] 日志记录完整

### 5.2 性能验收

- [ ] 执行时长符合要求
- [ ] 成功率达标
- [ ] 资源使用合理

### 5.3 文档验收

- [ ] 用户手册完整
- [ ] 运维手册完整
- [ ] 代码注释清晰

## 6. 项目计划

### 6.1 里程碑

| 阶段 | 开始时间 | 结束时间 | 交付物 |
|------|---------|---------|--------|
| 需求 | [日期] | [日期] | 需求说明书 |
| 设计 | [日期] | [日期] | 设计文档 |
| 编码 | [日期] | [日期] | 流程代码 |
| 测试 | [日期] | [日期] | 测试报告 |
| 上线 | [日期] | [日期] | 上线文档 |

### 6.2 资源需求

- 开发人员：[X]人
- 测试人员：[X]人
- 业务人员：[X]人

## 7. 风险评估

| 风险 | 影响 | 概率 | 应对措施 |
|------|------|------|---------|
| [风险1] | [高/中/低] | [高/中/低] | [措施] |
| [风险2] | [高/中/低] | [高/中/低] | [措施] |

---

**文档版本**：v1.0
**创建时间**：{datetime.now().strftime('%Y-%m-%d')}
**创建人**：[姓名]
'''
        return template

    def _generate_design_template(self, project_name):
        """生成设计文档模板"""
        template = f'''# {project_name} 整体设计文档

## 1. 设计概述

### 1.1 设计目标

[描述设计目标]

### 1.2 设计原则

- 模块化设计
- 异常处理完善
- 配置化管理
- 可扩展性

## 2. 架构设计

### 2.1 整体架构

```
整体架构
├── 流程层
│   ├── 主流程
│   ├── 子流程
│   └── 公共模块
├── 数据层
│   ├── 输入数据
│   ├── 中间数据
│   └── 输出数据
└── 配置层
    ├── 业务配置
    └── 系统配置
```

### 2.2 流程架构

**主流程**：
```
开始 → 初始化 → 业务处理 → 结果输出 → 结束
         ↓          ↓          ↓
      异常处理   异常处理   异常处理
```

### 2.3 数据架构

**数据流转**：
```
输入数据 → 数据验证 → 数据处理 → 数据存储 → 输出数据
```

## 3. 模块设计

### 3.1 初始化模块

**功能**：环境初始化、配置加载

**输入**：无

**输出**：全局配置字典

**流程**：
1. 关闭相关进程
2. 读取配置文件
3. 初始化全局变量
4. 检查运行环境

### 3.2 业务处理模块

**功能**：核心业务逻辑处理

**输入**：业务数据

**输出**：处理结果

**流程**：
1. 获取业务数据
2. 数据验证
3. 业务处理
4. 结果保存

### 3.3 异常处理模块

**功能**：异常捕获和处理

**异常类型**：
- 业务异常
- 系统异常
- 网络异常

**处理策略**：
- 重试机制
- 降级处理
- 告警通知

## 4. 数据设计

### 4.1 全局变量

```vb
g_dictGlobal = {
    "isEx": False,           // 异常标识
    "maxRetryNum": 3,        // 最大重试次数
    "config": {},            // 配置信息
    "data": []               // 业务数据
}
```

### 4.2 配置文件

**Config.ini 结构**：
```ini
[参数值]
LogLevel=2
maxRetryNum=3

[业务参数]
# 业务相关配置
```

## 5. 异常设计

### 5.1 异常分类

| 异常类型 | 处理策略 | 是否重试 |
|---------|---------|---------|
| 业务异常 | 记录日志，继续执行 | 否 |
| 系统异常 | 重试，超过次数则告警 | 是 |
| 网络异常 | 重试，超过次数则告警 | 是 |

### 5.2 异常处理流程

```
异常发生 → 异常捕获 → 异常分类 → 处理策略 → 日志记录
                                    ↓
                              是否重试？
                              ↙      ↘
                            重试    告警通知
```

## 6. 日志设计

### 6.1 日志级别

- 0：错误（Error）
- 1：警告（Warning）
- 2：信息（Info）

### 6.2 日志内容

- 时间戳
- 日志级别
- 模块名称
- 日志内容

## 7. 性能设计

### 7.1 性能目标

- 单次执行时长：< [X]分钟
- 成功率：> 95%
- 资源使用率：< 80%

### 7.2 性能优化

- 减少不必要的等待
- 优化元素定位
- 合理使用缓存
- 及时释放资源

## 8. 安全设计

### 8.1 数据安全

- 敏感数据加密
- 配置文件权限控制
- 日志脱敏处理

### 8.2 访问控制

- 应用访问权限
- 数据库访问权限
- 文件访问权限

---

**文档版本**：v1.0
**创建时间**：{datetime.now().strftime('%Y-%m-%d')}
**设计人**：[姓名]
'''
        return template

    def _generate_code_template(self, project_name, author):
        """生成代码模板"""
        template = f'''/*
作者：{author}
创建时间：{datetime.now().strftime('%Y年%m月%d日')}
项目名称：{project_name}
功能说明：主流程入口
*/

TracePrint "——————{project_name}流程开始——————"

// 全局变量初始化
Dim g_dictGlobal = {{}}
Dim g_iRetryNum = 0

/*1.初始化*/
Try
    PublicBlock.InitConfig()
Catch Ex
    Log.Error("初始化失败: " & CStr(Ex))
    g_dictGlobal["isEx"] = True
End Try

/*2.检查初始化结果*/
If g_dictGlobal["isEx"] = True
    If g_iRetryNum < g_dictGlobal["maxRetryNum"]
        g_iRetryNum = g_iRetryNum + 1
        Log.Info("第" & g_iRetryNum & "次重试")
        // 跳转回初始化
    Else
        Log.Error("初始化失败，超过最大重试次数")
        // 发送告警邮件
        PublicBlock.SendAlertEmail("初始化失败")
    End If
Else
    /*3.业务处理*/
    Try
        // 调用业务处理模块
        // TODO: 实现业务逻辑

        Log.Info("业务处理完成")
    Catch Ex
        Log.Error("业务处理失败: " & CStr(Ex))
        g_dictGlobal["isEx"] = True
    End Try

    /*4.发送结果通知*/
    Try
        PublicBlock.SendResultEmail()
    Catch Ex
        Log.Error("发送邮件失败: " & CStr(Ex))
    End Try
End If

TracePrint "——————{project_name}流程结束——————"
'''
        return template

    def _generate_test_template(self, project_name):
        """生成测试文档模板"""
        return "测试文档模板"

    def _generate_deployment_template(self, project_name):
        """生成部署文档模板"""
        return "部署文档模板"

    def _generate_operation_template(self, project_name):
        """生成运维文档模板"""
        return "运维文档模板"

    def generate_requirement_doc(self, params):
        """生成需求文档"""
        return self._generate_requirement_template(params.get('project_name', 'MyProject'))

    def generate_design_doc(self, params):
        """生成设计文档"""
        return self._generate_design_template(params.get('project_name', 'MyProject'))

    def generate_test_doc(self, params):
        """生成测试文档"""
        return self._generate_test_template(params.get('project_name', 'MyProject'))

    def generate_deployment_doc(self, params):
        """生成部署文档"""
        return self._generate_deployment_template(params.get('project_name', 'MyProject'))

    def generate_operation_doc(self, params):
        """生成运维文档"""
        return self._generate_operation_template(params.get('project_name', 'MyProject'))

    def generate(self, template_type, params, output_dir=None):
        """生成模板"""
        if template_type not in self.templates:
            print(f"错误：不支持的模板类型 '{template_type}'")
            print(f"可用模板：{', '.join(self.templates.keys())}")
            return None

        result = self.templates[template_type](params)

        if output_dir:
            if isinstance(result, dict):
                # 完整项目结构
                os.makedirs(output_dir, exist_ok=True)

                for key, content in result.items():
                    if key == 'structure':
                        print(f"\n项目结构：\n{content}")
                        continue

                    filename = f"{key}.md" if key != 'config' else "Config.ini"
                    filepath = os.path.join(output_dir, filename)

                    with open(filepath, 'w', encoding='utf-8') as f:
                        f.write(content)
                    print(f"✓ 已生成: {filepath}")
            else:
                # 单个文档
                with open(output_dir, 'w', encoding='utf-8') as f:
                    f.write(result)
                print(f"✓ 已生成: {output_dir}")

        return result


def main():
    """主函数"""
    import argparse

    parser = argparse.ArgumentParser(description='RPA项目模板生成器')
    parser.add_argument('-t', '--template', required=True,
                       choices=['rpa_project', 'requirement_doc', 'design_doc',
                               'test_doc', 'deployment_doc', 'operation_doc'],
                       help='模板类型')
    parser.add_argument('-p', '--params', type=str, default='{}',
                       help='参数JSON字符串')
    parser.add_argument('-o', '--output', type=str,
                       help='输出路径')
    parser.add_argument('-i', '--interactive', action='store_true',
                       help='交互模式')

    args = parser.parse_args()

    generator = RPAProjectGenerator()

    if args.interactive:
        print("=== RPA项目模板生成器 - 交互模式 ===\n")
        print("可用模板：")
        for i, template in enumerate(generator.templates.keys(), 1):
            print(f"{i}. {template}")

        choice = input("\n请选择模板编号: ")
        template_list = list(generator.templates.keys())
        template_type = template_list[int(choice) - 1]

        project_name = input("项目名称 [MyRPAProject]: ") or "MyRPAProject"
        author = input("作者姓名 [RPA开发者]: ") or "RPA开发者"
        params = {"project_name": project_name, "author": author}

        output_path = input("输出路径 [./output]: ") or "./output"

        result = generator.generate(template_type, params, output_path)
    else:
        params = json.loads(args.params)
        result = generator.generate(args.template, params, args.output)

    if result and not args.output:
        print("\n生成的内容：\n")
        if isinstance(result, dict):
            for key in ['readme', 'config']:
                if key in result:
                    print(f"\n=== {key} ===\n")
                    print(result[key][:500] + "..." if len(result[key]) > 500 else result[key])
        else:
            print(result[:1000] + "..." if len(result) > 1000 else result)


if __name__ == '__main__':
    main()
