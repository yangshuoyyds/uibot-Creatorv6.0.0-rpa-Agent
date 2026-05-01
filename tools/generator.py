#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
UIBot 代码生成器
根据模板和用户输入生成 UIBot 代码
"""

import os
import sys
import json
from pathlib import Path
from typing import Dict, List
from datetime import datetime

class UIBotCodeGenerator:
    def __init__(self, skill_dir: str = None):
        """初始化代码生成器"""
        if skill_dir is None:
            skill_dir = Path(__file__).parent.parent
        self.skill_dir = Path(skill_dir)
        self.templates = {}
        self.load_templates()

    def load_templates(self):
        """加载代码模板"""
        self.templates = {
            'web_automation': {
                'name': '网页自动化基础模板',
                'description': '打开网页、查找元素、执行操作',
                'params': ['url', 'element_selector', 'action'],
                'template': '''// 网页自动化流程
// 生成时间: {timestamp}
// URL: {url}

// 1. 打开浏览器
hWeb = WebBrowser.Create("cef", 30)
WebBrowser.Navigate(hWeb, "{url}")

// 2. 等待页面加载
Dim objUiElement
objUiElement = UiElement.Wait(hWeb, "<html app='chrome' title='*' />", 30)

// 3. 查找目标元素
Dim targetElement
targetElement = UiElement.Wait(hWeb, "{element_selector}", 10)

// 4. 执行操作
If targetElement <> Nothing Then
    {action_code}
    TracePrint("操作成功")
Else
    TracePrint("未找到元素")
End If

// 5. 关闭浏览器
// WebBrowser.Close(hWeb)
'''
            },
            'excel_processing': {
                'name': 'Excel 数据处理模板',
                'description': '读取、处理、写入 Excel 数据',
                'params': ['input_file', 'output_file', 'sheet_name'],
                'template': '''// Excel 数据处理流程
// 生成时间: {timestamp}
// 输入文件: {input_file}
// 输出文件: {output_file}

// 1. 打开 Excel 文件
Dim objExcel
objExcel = Excel.Open("{input_file}", True, "")

// 2. 选择工作表
Excel.SetSheet(objExcel, "{sheet_name}")

// 3. 读取数据
Dim iRow, iCol, sValue
Dim iMaxRow = Excel.GetUsedRowsCount(objExcel)
Dim iMaxCol = Excel.GetUsedColumnsCount(objExcel)

TracePrint("总行数: " & iMaxRow & ", 总列数: " & iMaxCol)

// 4. 处理数据
For iRow = 2 To iMaxRow
    // 读取每行数据
    sValue = Excel.Read(objExcel, iRow, 1)
    TracePrint("第" & iRow & "行: " & sValue)

    // TODO: 在这里添加数据处理逻辑

Next

// 5. 保存结果
Excel.SaveAs(objExcel, "{output_file}")
Excel.Close(objExcel)

TracePrint("处理完成")
'''
            },
            'file_batch': {
                'name': '文件批量处理模板',
                'description': '批量处理文件夹中的文件',
                'params': ['folder_path', 'file_pattern'],
                'template': '''// 文件批量处理流程
// 生成时间: {timestamp}
// 文件夹: {folder_path}
// 文件模式: {file_pattern}

// 1. 获取文件列表
Dim arrFiles
arrFiles = File.EnumFile("{folder_path}", "{file_pattern}", False)

// 2. 检查文件数量
If UBound(arrFiles) = -1 Then
    TracePrint("未找到匹配的文件")
    Exit Sub
End If

TracePrint("找到 " & (UBound(arrFiles) + 1) & " 个文件")

// 3. 批量处理
Dim sFile, sFileName, sFileExt
Dim iCount = 0

For Each sFile In arrFiles
    sFileName = File.GetName(sFile)
    sFileExt = File.GetExtensionName(sFile)

    TracePrint("处理文件: " & sFileName)

    // TODO: 在这里添加文件处理逻辑

    iCount = iCount + 1
Next

TracePrint("处理完成，共处理 " & iCount & " 个文件")
'''
            },
            'data_collection': {
                'name': '数据采集模板',
                'description': '循环采集网页数据并保存',
                'params': ['base_url', 'max_pages', 'output_file'],
                'template': '''// 数据采集流程
// 生成时间: {timestamp}
// 基础URL: {base_url}
// 最大页数: {max_pages}

// 1. 初始化
Dim hWeb = WebBrowser.Create("cef", 30)
Dim arrData = Array()
Dim iPage = 1

// 2. 循环采集
While iPage <= {max_pages}
    TracePrint("采集第 " & iPage & " 页")

    // 构建URL
    Dim sUrl = "{base_url}" & "?page=" & iPage
    WebBrowser.Navigate(hWeb, sUrl)

    // 等待页面加载
    Delay(2000)

    // TODO: 提取数据
    // Dim objElements = UiElement.GetAll(hWeb, "<html ... />")
    // For Each objElement In objElements
    //     Dim sData = UiElement.GetValue(objElement, "innertext")
    //     arrData.Push(sData)
    // Next

    iPage = iPage + 1
Wend

// 3. 保存数据
Dim objExcel = Excel.Create(True, "")
Dim iRow = 1
For Each sData In arrData
    Excel.Write(objExcel, iRow, 1, sData)
    iRow = iRow + 1
Next
Excel.SaveAs(objExcel, "{output_file}")
Excel.Close(objExcel)

// 4. 清理
WebBrowser.Close(hWeb)
TracePrint("采集完成，共 " & arrData.Length & " 条数据")
'''
            },
            'reframework': {
                'name': 'REFramework 企业级模板',
                'description': '企业级流程框架（初始化-处理-结束）',
                'params': ['process_name', 'config_file'],
                'template': '''// REFramework 企业级流程
// 生成时间: {timestamp}
// 流程名称: {process_name}
// 配置文件: {config_file}

'========== 全局变量 ==========
Dim g_dictConfig = Nothing      // 配置字典
Dim g_arrQueue = Array()        // 数据队列
Dim g_iRetryCount = 0           // 重试计数
Dim g_sLogFile = ""             // 日志文件

'========== 主流程 ==========
Sub Main()
    Try
        // 1. 初始化
        TracePrint("========== 流程开始 ==========")
        Call InitProcess()

        // 2. 获取数据
        Call GetTransactionData()

        // 3. 处理数据
        If g_arrQueue.Length > 0 Then
            Call ProcessTransactions()
        Else
            TracePrint("没有待处理的数据")
        End If

        // 4. 流程结束
        Call EndProcess()
        TracePrint("========== 流程结束 ==========")

    Catch ex
        TracePrint("流程异常: " & ex.Message)
        Call HandleException(ex)
    End Try
End Sub

'========== 初始化 ==========
Sub InitProcess()
    TracePrint(">>> 初始化流程")

    // 加载配置
    g_dictConfig = LoadConfig("{config_file}")

    // 初始化日志
    g_sLogFile = "log_" & Format(Now(), "yyyyMMdd_HHmmss") & ".txt"

    // 初始化应用程序
    // TODO: 打开目标应用、浏览器等

    TracePrint("初始化完成")
End Sub

'========== 获取数据 ==========
Sub GetTransactionData()
    TracePrint(">>> 获取待处理数据")

    // TODO: 从数据源获取数据
    // 示例：从 Excel 读取
    ' Dim objExcel = Excel.Open(g_dictConfig("DataFile"), True, "")
    ' Dim iRow = 2
    ' While Excel.Read(objExcel, iRow, 1) <> ""
    '     g_arrQueue.Push(Excel.Read(objExcel, iRow, 1))
    '     iRow = iRow + 1
    ' Wend
    ' Excel.Close(objExcel)

    TracePrint("获取到 " & g_arrQueue.Length & " 条数据")
End Sub

'========== 处理事务 ==========
Sub ProcessTransactions()
    TracePrint(">>> 开始处理事务")

    Dim iSuccess = 0
    Dim iFailed = 0

    For Each transactionItem In g_arrQueue
        Try
            TracePrint("处理: " & transactionItem)

            // TODO: 处理单个事务
            Call ProcessTransaction(transactionItem)

            iSuccess = iSuccess + 1
            g_iRetryCount = 0  // 重置重试计数

        Catch ex
            TracePrint("处理失败: " & ex.Message)
            iFailed = iFailed + 1

            // 重试机制
            If g_iRetryCount < CInt(g_dictConfig("MaxRetry")) Then
                g_iRetryCount = g_iRetryCount + 1
                TracePrint("重试 " & g_iRetryCount & " 次")
                // TODO: 重试逻辑
            Else
                TracePrint("超过最大重试次数，跳过")
                g_iRetryCount = 0
            End If
        End Try
    Next

    TracePrint("处理完成 - 成功: " & iSuccess & ", 失败: " & iFailed)
End Sub

'========== 处理单个事务 ==========
Sub ProcessTransaction(transactionItem)
    // TODO: 实现具体的业务逻辑
    TracePrint("  执行业务逻辑...")
    Delay(1000)
End Sub

'========== 异常处理 ==========
Sub HandleException(ex)
    TracePrint("!!! 异常处理 !!!")
    TracePrint("错误信息: " & ex.Message)

    // TODO: 记录日志、发送通知等

    // 截图保存
    ' Dim sScreenshot = "error_" & Format(Now(), "yyyyMMdd_HHmmss") & ".png"
    ' Image.CaptureScreen(sScreenshot)
End Sub

'========== 流程结束 ==========
Sub EndProcess()
    TracePrint(">>> 流程结束清理")

    // TODO: 关闭应用程序、保存日志等

    TracePrint("清理完成")
End Sub

'========== 加载配置 ==========
Function LoadConfig(sConfigFile)
    Dim dictConfig = CreateObject("Scripting.Dictionary")

    // TODO: 从配置文件加载
    dictConfig("MaxRetry") = "3"
    dictConfig("DataFile") = "data.xlsx"

    Return dictConfig
End Function

// 启动主流程
Call Main()
'''
            },
            'error_handling': {
                'name': '完整错误处理模板',
                'description': '包含日志、重试、通知的错误处理框架',
                'params': ['process_name'],
                'template': '''// 完整错误处理框架
// 生成时间: {timestamp}
// 流程名称: {process_name}

'========== 全局配置 ==========
Dim g_sLogFile = "log_" & Format(Now(), "yyyyMMdd_HHmmss") & ".txt"
Dim g_iMaxRetry = 3
Dim g_iRetryDelay = 2000

'========== 主流程 ==========
Sub Main()
    Try
        Call LogInfo("流程开始")

        // TODO: 主要业务逻辑
        Call YourBusinessLogic()

        Call LogInfo("流程成功完成")

    Catch ex
        Call LogError("流程异常: " & ex.Message)
        Call HandleCriticalError(ex)
    End Try
End Sub

'========== 业务逻辑（示例）==========
Sub YourBusinessLogic()
    // 使用重试机制执行操作
    Dim bSuccess = ExecuteWithRetry("ClickButton", Array())

    If Not bSuccess Then
        Throw "操作失败"
    End If
End Sub

'========== 带重试的执行 ==========
Function ExecuteWithRetry(sFunctionName, arrParams)
    Dim iRetry = 0

    While iRetry < g_iMaxRetry
        Try
            Call LogInfo("执行: " & sFunctionName & " (尝试 " & (iRetry + 1) & ")")

            // 根据函数名执行对应操作
            Select Case sFunctionName
                Case "ClickButton"
                    // TODO: 实现点击逻辑
                    TracePrint("执行点击操作")
                Case Else
                    TracePrint("未知操作: " & sFunctionName)
            End Select

            Call LogInfo("执行成功")
            Return True

        Catch ex
            iRetry = iRetry + 1
            Call LogWarning("执行失败 (尝试 " & iRetry & "): " & ex.Message)

            If iRetry < g_iMaxRetry Then
                Call LogInfo("等待 " & g_iRetryDelay & "ms 后重试")
                Delay(g_iRetryDelay)
            End If
        End Try
    Wend

    Call LogError("执行失败，已达最大重试次数")
    Return False
End Function

'========== 日志函数 ==========
Sub LogInfo(sMessage)
    Dim sLog = "[INFO] " & Format(Now(), "yyyy-MM-dd HH:mm:ss") & " - " & sMessage
    TracePrint(sLog)
    Call WriteLog(sLog)
End Sub

Sub LogWarning(sMessage)
    Dim sLog = "[WARN] " & Format(Now(), "yyyy-MM-dd HH:mm:ss") & " - " & sMessage
    TracePrint(sLog)
    Call WriteLog(sLog)
End Sub

Sub LogError(sMessage)
    Dim sLog = "[ERROR] " & Format(Now(), "yyyy-MM-dd HH:mm:ss") & " - " & sMessage
    TracePrint(sLog)
    Call WriteLog(sLog)
End Sub

Sub WriteLog(sMessage)
    Try
        File.AppendAllText(g_sLogFile, sMessage & vbCrLf)
    Catch ex
        TracePrint("写入日志失败: " & ex.Message)
    End Try
End Sub

'========== 严重错误处理 ==========
Sub HandleCriticalError(ex)
    // 1. 截图
    Try
        Dim sScreenshot = "error_" & Format(Now(), "yyyyMMdd_HHmmss") & ".png"
        Image.CaptureScreen(sScreenshot)
        Call LogInfo("已保存错误截图: " & sScreenshot)
    Catch
        Call LogError("截图失败")
    End Try

    // 2. 发送通知（可选）
    ' Call SendErrorNotification(ex.Message)

    // 3. 清理资源
    Call CleanupResources()
End Sub

'========== 资源清理 ==========
Sub CleanupResources()
    Try
        Call LogInfo("清理资源")
        // TODO: 关闭浏览器、Excel 等
    Catch ex
        Call LogError("清理资源失败: " & ex.Message)
    End Try
End Sub

// 启动主流程
Call Main()
'''
            }
        }

    def list_templates(self):
        """列出所有模板"""
        print("\n可用模板:")
        print("=" * 80)
        for key, template in self.templates.items():
            print(f"\n【{key}】 {template['name']}")
            print(f"说明: {template['description']}")
            print(f"参数: {', '.join(template['params'])}")
        print("=" * 80)

    def generate(self, template_key: str, params: Dict[str, str]) -> str:
        """生成代码"""
        if template_key not in self.templates:
            raise ValueError(f"模板不存在: {template_key}")

        template = self.templates[template_key]

        # 添加时间戳
        params['timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # 处理特殊参数
        if 'action' in params:
            action_map = {
                '点击': 'UiElement.Click(targetElement, 1, 1, 1)',
                '输入': 'UiElement.SetValue(targetElement, "输入内容", "value")',
                '获取文本': 'Dim sText = UiElement.GetValue(targetElement, "innertext")'
            }
            params['action_code'] = action_map.get(params['action'],
                                                   'UiElement.Click(targetElement, 1, 1, 1)')

        # 生成代码
        code = template['template'].format(**params)
        return code

    def save_code(self, code: str, output_file: str):
        """保存代码到文件"""
        output_path = Path(output_file)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(code)

        print(f"\n✓ 代码已保存到: {output_path}")

    def interactive_generate(self):
        """交互式生成"""
        print("\n" + "=" * 80)
        print("UIBot 代码生成器 - 交互模式")
        print("=" * 80)

        # 列出模板
        self.list_templates()

        # 选择模板
        print("\n请选择模板 (输入模板 key):")
        template_key = input(">>> ").strip()

        if template_key not in self.templates:
            print(f"❌ 模板不存在: {template_key}")
            return

        template = self.templates[template_key]

        # 输入参数
        print(f"\n生成 {template['name']}")
        print("请输入参数:")
        params = {}
        for param in template['params']:
            value = input(f"  {param}: ").strip()
            params[param] = value

        # 生成代码
        try:
            code = self.generate(template_key, params)
            print("\n" + "=" * 80)
            print("生成的代码:")
            print("=" * 80)
            print(code)
            print("=" * 80)

            # 保存
            save = input("\n是否保存到文件? (y/n): ").strip().lower()
            if save == 'y':
                output_file = input("输出文件名: ").strip()
                if not output_file.endswith('.task'):
                    output_file += '.task'
                self.save_code(code, output_file)

        except Exception as e:
            print(f"❌ 生成失败: {e}")

def main():
    """主函数"""
    import argparse

    parser = argparse.ArgumentParser(description='UIBot 代码生成器')
    parser.add_argument('-t', '--template', help='模板名称')
    parser.add_argument('-p', '--params', help='参数 JSON 字符串')
    parser.add_argument('-o', '--output', help='输出文件')
    parser.add_argument('-l', '--list', action='store_true', help='列出所有模板')
    parser.add_argument('-i', '--interactive', action='store_true', help='交互模式')

    args = parser.parse_args()

    # 初始化生成器
    generator = UIBotCodeGenerator()

    # 列出模板
    if args.list:
        generator.list_templates()
        return

    # 交互模式
    if args.interactive or not args.template:
        generator.interactive_generate()
        return

    # 命令行生成
    if args.template:
        try:
            params = json.loads(args.params) if args.params else {}
            code = generator.generate(args.template, params)

            if args.output:
                generator.save_code(code, args.output)
            else:
                print(code)

        except Exception as e:
            print(f"❌ 错误: {e}")
            sys.exit(1)

if __name__ == '__main__':
    main()
