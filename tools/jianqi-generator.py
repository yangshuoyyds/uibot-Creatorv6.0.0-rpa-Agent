#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
剑气代码生成器
基于高级案例的标准化代码模板生成工具

功能：
1. 多数据源采集模板
2. 数据对比模板
3. 配置驱动流程模板
4. 完整项目结构生成

作者：UB-Skill Team
版本：1.0.0
日期：2026-05-02
"""

import json
import os
from datetime import datetime


class JianqiGenerator:
    """剑气代码生成器"""

    def __init__(self):
        self.templates = {
            'multi_source': self.generate_multi_source_template,
            'data_compare': self.generate_data_compare_template,
            'config_driven': self.generate_config_driven_template,
            'full_project': self.generate_full_project_template,
            'public_block': self.generate_public_block_template,
            'init_module': self.generate_init_module_template
        }

    def generate_multi_source_template(self, params):
        """生成多数据源采集模板"""
        source_name = params.get('source_name', '数据源1')
        url_var = params.get('url_var', 'source1Url')
        data_var = params.get('data_var', 'source1Data')

        template = f'''/*
作者：{params.get('author', 'RPA开发者')}
创建时间：{datetime.now().strftime('%Y年%m月%d日')}
本流程块用于实现{source_name}数据采集，对应设计步骤如下：
1. 打开{source_name}网站
2. 输入查询条件
3. 执行查询操作
4. 抓取数据
5. 处理数据格式
*/

TracePrint "——————进入{source_name}模块——————"
Dim hWeb = "" // 浏览器对象
Dim arrayData = "" // 数据抓取数组
Dim dataLen = "" // 数据数组长度
Dim eachDataArr = [] // 每一条数据
Dim dataArr = [] // 最终数据结果

/*1.打开{source_name}网站*/
Try 3
    App.Kill('chrome.exe')
    hWeb = WebBrowser.Create("chrome", g_dictGlobal["{url_var}"], 30000)

    // 窗口最大化
    Window.Show({{"wnd":[{{"cls":"Chrome_WidgetWin_1","title":"*","app":"chrome"}}]}}, "max")

    // 等待页面加载完成（通过点击某个元素判断）
    Mouse.Action({{"html":[{{"tag":"DIV","id":"main"}}]}}, "left", "click", 10000)

Catch Ex
    PublicBlock.ErrCapture("打开{source_name}网站失败:", Ex)
    Return
Else
    Log.Info('成功打开{source_name}网站')
End Try

/*2.输入查询条件*/
Try 3
    // 输入关键词
    Keyboard.InputText({{"html":[{{"tag":"INPUT","id":"keyword"}}]}},
                      g_dictGlobal["keyword"], True, 20, 10000)

    // 点击查询按钮
    Mouse.Action({{"html":[{{"tag":"BUTTON","id":"search"}}]}},
                "left", "click", 10000)

    // 等待结果加载
    Time.Sleep(2000)

Catch Ex
    PublicBlock.ErrCapture("查询操作失败:", Ex)
    Return
Else
    Log.Info('查询操作成功')
End Try

/*3.抓取数据*/
Try
    // 滚动页面确保数据加载
    Mouse.Wheel(30, "down", [], {{"iDelayAfter":300}})

    // 数据抓取
    arrayData = UiElement.DataScrap(
        {{"html":[{{"id":"result-list","tag":"DIV"}}]}},
        {{"Columns":[
            {{"props":["text"], "selecors":[...]}},
            {{"props":["text"], "selecors":[...]}}
        ], "ExtractTable":0}},
        {{"iMaxNumberOfPage":5, "iDelayBetweenMS":1000}}
    )

Catch Ex
    PublicBlock.ErrCapture("抓取数据失败:", Ex)
    Return
Else
    dataLen = Len(arrayData)
    If dataLen > 0
        Log.Info('抓取数据成功，共' & dataLen & '条')
    Else
        Log.Error('数据长度为0，抓取失败')
        g_dictGlobal["isEx"] = True
        Return
    End If
End Try

/*4.处理数据格式*/
Try
    For Each value In arrayData
        eachDataArr = []

        // 数据来源
        eachDataArr = push(eachDataArr, "{source_name}")

        // 其他字段处理
        eachDataArr = push(eachDataArr, value[0])
        eachDataArr = push(eachDataArr, value[1])

        dataArr = push(dataArr, eachDataArr)
    Next

    g_dictGlobal["{data_var}"] = dataArr

Catch Ex
    PublicBlock.ErrCapture("处理数据格式失败:", Ex)
    Return
Else
    Log.Info('处理数据格式成功')
End Try

TracePrint "——————退出{source_name}模块——————"
'''
        return template

    def generate_data_compare_template(self, params):
        """生成数据对比模板"""
        template = f'''/*
作者：{params.get('author', 'RPA开发者')}
创建时间：{datetime.now().strftime('%Y年%m月%d日')}
本流程块用于实现数据对比和写入，对应设计步骤如下：
1. 判断本流程块是否执行
2. 数据源1与数据源2对比
3. 处理独有数据
4. 结果写入Excel
5. 数据排序
*/

/*1.判断本流程块是否执行*/
If g_dictGlobal["isEx"] = True
    Log.Info("进入数据对比前流程出错")
    Return
End If

TracePrint "——————进入数据对比模块——————"
App.Kill('chrome.exe')

Dim source1Data = g_dictGlobal["source1Data"]
Dim source2Data = g_dictGlobal["source2Data"]
Dim objExcelWorkBook = ""
Dim price = ''
Dim source = ''
Dim eachArr = []
Dim resultArr = []

/*2.数据源1与数据源2对比*/
For Each s1Value In source1Data
    s1Value[2] = CInt(s1Value[2])  // 价格字段

    For Each s2Value In source2Data
        s2Value[2] = CInt(s2Value[2])

        // 匹配条件：相同的标识字段
        If s1Value[1] = s2Value[1]
            eachArr = []

            // 价格对比，选择更低价格
            If s1Value[2] > s2Value[2]
                price = s2Value[2]
                source = '数据源2'
            Else
                price = s1Value[2]
                source = '数据源1'
            End If

            eachArr = [source, s1Value[0], s1Value[1], price]
            resultArr = push(resultArr, eachArr)
        End If
    Next
Next

/*3.处理独有数据*/
For Each s1Value In source1Data
    Dim isFound = False

    For Each resValue In resultArr
        If s1Value[1] = resValue[2]
            isFound = True
            Break
        End If
    Next

    If isFound = False
        resultArr = push(resultArr, s1Value)
    End If
Next

/*4.结果写入Excel*/
Try
    objExcelWorkBook = Excel.OpenExcel(@res"config\\输出模板.xlsx", False)
    Excel.WriteRange(objExcelWorkBook, "Sheet1", "A2", resultArr, False)

    // 另存为
    Dim fileName = Time.Format(Time.Now(), "yyyymmdd") & "_结果.xlsx"
    Excel.SaveOtherFile(objExcelWorkBook, @res"config\\" & fileName)
    Excel.CloseExcel(objExcelWorkBook, False)

Catch Ex
    PublicBlock.ErrCapture("Excel写入失败:", Ex)
    Return
Else
    Log.Info('Excel写入成功')
End Try

TracePrint "——————退出数据对比模块——————"
'''
        return template

    def generate_config_driven_template(self, params):
        """生成配置驱动流程模板"""
        template = f'''/*
作者：{params.get('author', 'RPA开发者')}
创建时间：{datetime.now().strftime('%Y年%m月%d日')}
配置驱动流程 - 初始化模块
*/

TracePrint "——————进入初始化模块——————"

/*1.结束所有相关进程*/
App.Kill("chrome.exe")
App.Kill("excel.exe")

/*2.全局变量以及日志级别的初始化*/
PublicBlock.InitArgByLocal()

/*3.检查Excel是否满足流程运行环境*/
Try
    objExcelWorkBook = Excel.OpenExcel(@res"config\\模板.xlsx", True)
    Excel.CloseExcel(objExcelWorkBook, True)
Catch Ex
    PublicBlock.ErrCapture("初始化异常，Excel不满足流程运行条件:", Ex)
    Return
End Try

TracePrint "——————退出初始化模块——————"
'''
        return template

    def generate_public_block_template(self, params):
        """生成公共函数库模板"""
        template = f'''/*
作者：{params.get('author', 'RPA开发者')}
创建时间：{datetime.now().strftime('%Y年%m月%d日')}
公共模块，放置函数
*/

/*
功能：初始化全局变量
入参：无
出参：无
*/
Function InitArgByLocal()
    TracePrint "——————进入InitArgByLocal函数——————"

    Dim logLevel = ''
    Dim maxRetryNum = ''

    // 设置日志等级
    logLevel = INI.Read(@res"config\\Config.ini", "参数值", "LogLevel", "2")
    Log.SetLevel(logLevel)

    // 重试计数器
    g_iRetryNum = g_iRetryNum + 1
    Log.Info('第' & g_iRetryNum & '次开始流程')

    // 初始化全局参数
    g_dictGlobal = {{}}
    g_dictGlobal["isEx"] = False

    // 最大重试次数
    maxRetryNum = INI.Read(@res"config\\Config.ini", "参数值", "maxRetryNum", "3")
    g_dictGlobal["maxRetryNum"] = CInt(maxRetryNum)

    // 业务参数
    g_dictGlobal["url1"] = INI.Read(@res"config\\Config.ini", "地址", "URL1", "")
    g_dictGlobal["url2"] = INI.Read(@res"config\\Config.ini", "地址", "URL2", "")

    // 邮件配置
    g_dictGlobal["server"] = INI.Read(@res"config\\Config.ini", "邮件", "服务器", "smtp.qq.com")
    g_dictGlobal["port"] = CInt(INI.Read(@res"config\\Config.ini", "邮件", "端口", "25"))
    g_dictGlobal["passport"] = INI.Read(@res"config\\Config.ini", "邮件", "发件人", "")
    g_dictGlobal["password"] = INI.Read(@res"config\\Config.ini", "邮件", "密码", "")
    g_dictGlobal["sendAddr"] = INI.Read(@res"config\\Config.ini", "邮件", "收件人", "")

    TracePrint "——————退出InitArgByLocal函数——————"
End Function

/*其他公共函数省略...*/
'''
        return template

    def generate_init_module_template(self, params):
        """生成初始化模块模板"""
        template = f'''/*
作者：{params.get('author', 'RPA开发者')}
创建时间：{datetime.now().strftime('%Y年%m月%d日')}
本流程块用于实现流程的环境初始化
*/

TracePrint "——————进入初始化模块——————"

/*1.结束所有相关进程*/
App.Kill("chrome.exe")
App.Kill("excel.exe")

/*2.全局变量以及日志级别的初始化*/
PublicBlock.InitArgByLocal()

/*3.检查运行环境*/
Try
    objExcelWorkBook = Excel.OpenExcel(@res"config\\模板.xlsx", True)
    Excel.CloseExcel(objExcelWorkBook, True)
Catch Ex
    PublicBlock.ErrCapture("初始化异常，环境检查失败:", Ex)
    Return
End Try

TracePrint "——————退出初始化模块——————"
'''
        return template

    def generate_full_project_template(self, params):
        """生成完整项目结构"""
        project_name = params.get('project_name', 'MyRPAProject')

        structure = {
            'flow': self._generate_flow_file(params),
            'config_ini': self._generate_config_ini(params),
            'readme': self._generate_readme(params)
        }

        return structure

    def _generate_flow_file(self, params):
        """生成主流程文件"""
        flow_data = {
            "uuid": "auto-generated",
            "version": "5.5.0",
            "flow": [
                {
                    "rem": "begin",
                    "type": "begin",
                    "next": "init",
                    "x": 0,
                    "y": 10
                },
                {
                    "rem": "init",
                    "type": "task",
                    "desc": "初始化",
                    "file": "初始化.task",
                    "next": "check_init",
                    "x": 0,
                    "y": 90
                },
                {
                    "rem": "check_init",
                    "type": "if",
                    "desc": "初始化是否异常",
                    "expression": 'g_dictGlobal["isEx"] = True',
                    "yes": {"desc": "yes", "next": "retry"},
                    "no": {"desc": "no", "next": "process"},
                    "x": 0,
                    "y": 200
                },
                {
                    "rem": "process",
                    "type": "task",
                    "desc": "业务处理",
                    "file": "业务处理.task",
                    "next": "send_mail",
                    "x": 0,
                    "y": 350
                },
                {
                    "rem": "send_mail",
                    "type": "task",
                    "desc": "发送结果邮件",
                    "file": "发送结果邮件.task",
                    "next": "end",
                    "x": 0,
                    "y": 460
                },
                {
                    "rem": "retry",
                    "type": "if",
                    "desc": "是否重试",
                    "expression": 'g_iRetryNum < g_dictGlobal["maxRetryNum"]',
                    "yes": {"desc": "yes", "next": "init"},
                    "no": {"desc": "no", "next": "send_mail"},
                    "x": 200,
                    "y": 200
                },
                {
                    "rem": "end",
                    "type": "end",
                    "x": 0,
                    "y": 570
                }
            ],
            "dim": [
                {"var": "g_dictGlobal={}", "type": "none"},
                {"var": "g_iRetryNum=0", "type": "none"}
            ]
        }

        return json.dumps(flow_data, indent=2, ensure_ascii=False)

    def _generate_config_ini(self, params):
        """生成配置文件"""
        config = f'''[参数值]
LogLevel=2
maxRetryNum=3

[地址]
URL1=https://example.com
URL2=https://example2.com

[业务参数]
关键词=测试
出发地=广州
到达地=北京

[邮件]
服务器=smtp.qq.com
端口=25
发件人=your_email@qq.com
密码=your_password
收件人=receiver@example.com

[系统]
超时时间=30000
延迟时间=1000
'''
        return config

    def _generate_readme(self, params):
        """生成项目README"""
        project_name = params.get('project_name', 'RPA项目')

        readme = f'''# {project_name}

## 项目说明

本项目基于剑气RPA平台开发，实现自动化业务流程。

## 项目结构

```
{project_name}/
├── {project_name}.flow          # 主流程文件
├── 初始化.task                  # 初始化模块
├── 业务处理.task                # 业务处理模块
├── 发送结果邮件.task            # 邮件发送模块
├── PublicBlock.task             # 公共函数库
├── res/                         # 资源文件
│   └── config/
│       ├── Config.ini           # 配置文件
│       └── 模板.xlsx            # Excel模板
└── log/                         # 日志目录
```

## 配置说明

修改 `res/config/Config.ini` 文件中的参数：

- **日志级别**：0-错误，1-警告，2-信息
- **重试次数**：流程失败后的最大重试次数
- **业务参数**：根据实际业务需求配置
- **邮件配置**：配置邮件服务器和收发件人

## 运行说明

1. 确保已安装剑气RPA平台
2. 修改配置文件
3. 打开主流程文件运行

## 注意事项

- 首次运行前请检查Excel模板是否存在
- 邮件密码建议使用授权码
- 建议在测试环境充分测试后再用于生产

## 版本信息

- 创建时间：{datetime.now().strftime('%Y-%m-%d')}
- 剑气版本：5.5.0+
- 作者：{params.get('author', 'RPA开发者')}
'''
        return readme

    def generate(self, template_type, params, output_file=None):
        """生成代码"""
        if template_type not in self.templates:
            print(f"错误：不支持的模板类型 '{template_type}'")
            print(f"可用模板：{', '.join(self.templates.keys())}")
            return None

        result = self.templates[template_type](params)

        if output_file:
            with open(output_file, 'w', encoding='utf-8') as f:
                if isinstance(result, dict):
                    # 完整项目结构
                    for filename, content in result.items():
                        file_path = f"{output_file}_{filename}"
                        with open(file_path, 'w', encoding='utf-8') as sub_f:
                            sub_f.write(content)
                        print(f"✓ 已生成: {file_path}")
                else:
                    f.write(result)
                    print(f"✓ 已生成: {output_file}")

        return result


def main():
    """主函数"""
    import argparse

    parser = argparse.ArgumentParser(description='剑气代码生成器')
    parser.add_argument('-t', '--template', required=True,
                       choices=['multi_source', 'data_compare', 'config_driven',
                               'full_project', 'public_block', 'init_module'],
                       help='模板类型')
    parser.add_argument('-p', '--params', type=str, default='{}',
                       help='参数JSON字符串')
    parser.add_argument('-o', '--output', type=str,
                       help='输出文件路径')
    parser.add_argument('-i', '--interactive', action='store_true',
                       help='交互模式')

    args = parser.parse_args()

    generator = JianqiGenerator()

    if args.interactive:
        print("=== 剑气代码生成器 - 交互模式 ===\n")
        print("可用模板：")
        for i, template in enumerate(generator.templates.keys(), 1):
            print(f"{i}. {template}")

        choice = input("\n请选择模板编号: ")
        template_list = list(generator.templates.keys())
        template_type = template_list[int(choice) - 1]

        author = input("作者姓名 [RPA开发者]: ") or "RPA开发者"
        params = {"author": author}

        if template_type == 'multi_source':
            params['source_name'] = input("数据源名称 [数据源1]: ") or "数据源1"
        elif template_type == 'full_project':
            params['project_name'] = input("项目名称 [MyRPAProject]: ") or "MyRPAProject"

        output_file = input("输出文件路径 [output.task]: ") or "output.task"

        result = generator.generate(template_type, params, output_file)
    else:
        params = json.loads(args.params)
        result = generator.generate(args.template, params, args.output)

    if result and not args.output:
        print("\n生成的代码：\n")
        if isinstance(result, dict):
            for key, value in result.items():
                print(f"\n=== {key} ===\n")
                print(value[:500] + "..." if len(value) > 500 else value)
        else:
            print(result)


if __name__ == '__main__':
    main()

