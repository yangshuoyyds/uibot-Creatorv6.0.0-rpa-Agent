#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
UIBot 代码验证工具
检查 UIBot 代码的语法、最佳实践和潜在问题
"""

import re
import os
import sys
from pathlib import Path
from typing import List, Dict, Tuple

class UIBotValidator:
    def __init__(self):
        """初始化验证工具"""
        self.issues = []
        self.warnings = []
        self.suggestions = []

    def validate_file(self, file_path: str) -> Dict:
        """验证文件"""
        self.issues = []
        self.warnings = []
        self.suggestions = []

        if not os.path.exists(file_path):
            return {'error': f'文件不存在: {file_path}'}

        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 执行各项检查
        self.check_syntax(content)
        self.check_best_practices(content)
        self.check_performance(content)
        self.check_error_handling(content)
        self.check_security(content)

        return {
            'file': file_path,
            'issues': self.issues,
            'warnings': self.warnings,
            'suggestions': self.suggestions
        }

    def check_syntax(self, content: str):
        """检查语法问题"""
        lines = content.split('\n')

        for i, line in enumerate(lines, 1):
            # 检查未闭合的引号
            if line.count('"') % 2 != 0 and not line.strip().startswith('//'):
                self.issues.append({
                    'line': i,
                    'type': 'syntax',
                    'message': '可能存在未闭合的引号',
                    'code': line.strip()
                })

            # 检查 If 语句缺少 End If
            if re.search(r'\bIf\b.*\bThen\b', line, re.IGNORECASE):
                # 检查是否是单行 If
                if not re.search(r'\bThen\b.*\w+', line, re.IGNORECASE):
                    # 需要检查后续是否有 End If
                    pass

            # 检查变量命名规范
            var_match = re.search(r'\bDim\s+(\w+)', line, re.IGNORECASE)
            if var_match:
                var_name = var_match.group(1)
                if not self.is_valid_variable_name(var_name):
                    self.warnings.append({
                        'line': i,
                        'type': 'naming',
                        'message': f'变量命名不符合规范: {var_name}',
                        'suggestion': '建议使用匈牙利命名法 (如: sName, iCount, bFlag)'
                    })

    def is_valid_variable_name(self, name: str) -> bool:
        """检查变量名是否符合规范"""
        # UIBot 推荐的匈牙利命名法
        prefixes = ['s', 'i', 'b', 'f', 'd', 'arr', 'obj', 'dict', 'h', 'g_']
        for prefix in prefixes:
            if name.startswith(prefix) and len(name) > len(prefix):
                return True
        return False

    def check_best_practices(self, content: str):
        """检查最佳实践"""
        lines = content.split('\n')

        # 检查是否使用了硬编码路径
        for i, line in enumerate(lines, 1):
            if re.search(r'[C-Z]:\\', line) and not line.strip().startswith('//'):
                self.warnings.append({
                    'line': i,
                    'type': 'best_practice',
                    'message': '使用了硬编码路径',
                    'suggestion': '建议使用配置文件或相对路径'
                })

            # 检查是否使用了 Delay 而不是智能等待
            if re.search(r'\bDelay\s*\(\s*\d+\s*\)', line, re.IGNORECASE):
                self.suggestions.append({
                    'line': i,
                    'type': 'best_practice',
                    'message': '使用了固定延迟',
                    'suggestion': '建议使用 UiElement.Wait() 等智能等待方法'
                })

            # 检查是否缺少日志
            if re.search(r'\b(Try|Catch|If|For|While)\b', line, re.IGNORECASE):
                # 简单检查，实际应该更复杂
                pass

    def check_performance(self, content: str):
        """检查性能问题"""
        lines = content.split('\n')

        # 检查循环中的重复操作
        in_loop = False
        loop_start = 0

        for i, line in enumerate(lines, 1):
            if re.search(r'\b(For|While)\b', line, re.IGNORECASE):
                in_loop = True
                loop_start = i

            if in_loop and re.search(r'\b(Next|Wend)\b', line, re.IGNORECASE):
                in_loop = False

            # 检查循环中的文件操作
            if in_loop and re.search(r'\b(File\.|Excel\.Open|WebBrowser\.Create)\b', line):
                self.warnings.append({
                    'line': i,
                    'type': 'performance',
                    'message': '循环中存在重复的资源创建操作',
                    'suggestion': '建议将资源创建移到循环外'
                })

            # 检查是否使用了低效的字符串拼接
            if '+' in line and '"' in line and in_loop:
                self.suggestions.append({
                    'line': i,
                    'type': 'performance',
                    'message': '循环中使用字符串拼接',
                    'suggestion': '建议使用数组收集后再拼接'
                })

    def check_error_handling(self, content: str):
        """检查错误处理"""
        lines = content.split('\n')

        has_try = False
        has_catch = False

        for i, line in enumerate(lines, 1):
            if re.search(r'\bTry\b', line, re.IGNORECASE):
                has_try = True

            if re.search(r'\bCatch\b', line, re.IGNORECASE):
                has_catch = True

            # 检查关键操作是否有错误处理
            critical_ops = [
                r'Excel\.Open',
                r'WebBrowser\.Create',
                r'File\.Open',
                r'Database\.Connect'
            ]

            for op in critical_ops:
                if re.search(op, line, re.IGNORECASE):
                    # 简单检查前后是否有 Try-Catch
                    # 实际应该检查作用域
                    if not has_try:
                        self.warnings.append({
                            'line': i,
                            'type': 'error_handling',
                            'message': f'关键操作可能缺少错误处理: {op}',
                            'suggestion': '建议使用 Try-Catch 包裹'
                        })

        # 检查是否完全没有错误处理
        if not has_try and len(lines) > 20:
            self.warnings.append({
                'line': 0,
                'type': 'error_handling',
                'message': '代码中没有错误处理',
                'suggestion': '建议添加 Try-Catch 错误处理'
            })

    def check_security(self, content: str):
        """检查安全问题"""
        lines = content.split('\n')

        for i, line in enumerate(lines, 1):
            # 检查是否有明文密码
            if re.search(r'(password|pwd|passwd)\s*=\s*["\']', line, re.IGNORECASE):
                self.issues.append({
                    'line': i,
                    'type': 'security',
                    'message': '可能存在明文密码',
                    'suggestion': '建议使用配置文件或加密存储'
                })

            # 检查 SQL 注入风险
            if re.search(r'Database\.Execute.*\+.*', line, re.IGNORECASE):
                self.warnings.append({
                    'line': i,
                    'type': 'security',
                    'message': '可能存在 SQL 注入风险',
                    'suggestion': '建议使用参数化查询'
                })

            # 检查文件路径注入
            if re.search(r'File\.(Open|Delete|Move).*\+.*', line, re.IGNORECASE):
                self.warnings.append({
                    'line': i,
                    'type': 'security',
                    'message': '可能存在路径注入风险',
                    'suggestion': '建议验证和清理文件路径'
                })

    def print_report(self, result: Dict):
        """打印验证报告"""
        print("\n" + "=" * 80)
        print(f"UIBot 代码验证报告")
        print("=" * 80)
        print(f"\n文件: {result['file']}")

        # 统计
        issue_count = len(result['issues'])
        warning_count = len(result['warnings'])
        suggestion_count = len(result['suggestions'])

        print(f"\n统计: {issue_count} 个错误, {warning_count} 个警告, {suggestion_count} 个建议")

        # 错误
        if result['issues']:
            print("\n" + "-" * 80)
            print("❌ 错误:")
            for issue in result['issues']:
                print(f"\n  行 {issue['line']}: [{issue['type']}] {issue['message']}")
                if 'code' in issue:
                    print(f"    代码: {issue['code']}")
                if 'suggestion' in issue:
                    print(f"    建议: {issue['suggestion']}")

        # 警告
        if result['warnings']:
            print("\n" + "-" * 80)
            print("⚠️  警告:")
            for warning in result['warnings']:
                line_info = f"行 {warning['line']}" if warning['line'] > 0 else "全局"
                print(f"\n  {line_info}: [{warning['type']}] {warning['message']}")
                if 'suggestion' in warning:
                    print(f"    建议: {warning['suggestion']}")

        # 建议
        if result['suggestions']:
            print("\n" + "-" * 80)
            print("💡 建议:")
            for suggestion in result['suggestions']:
                print(f"\n  行 {suggestion['line']}: [{suggestion['type']}] {suggestion['message']}")
                if 'suggestion' in suggestion:
                    print(f"    建议: {suggestion['suggestion']}")

        # 总结
        print("\n" + "=" * 80)
        if issue_count == 0 and warning_count == 0:
            print("✓ 代码质量良好！")
        elif issue_count == 0:
            print("✓ 没有发现严重错误，但有一些改进建议")
        else:
            print("✗ 发现一些需要修复的问题")
        print("=" * 80)

def main():
    """主函数"""
    import argparse

    parser = argparse.ArgumentParser(description='UIBot 代码验证工具')
    parser.add_argument('file', nargs='?', help='要验证的文件路径')
    parser.add_argument('-d', '--directory', help='验证目录下所有文件')
    parser.add_argument('-r', '--recursive', action='store_true', help='递归验证子目录')

    args = parser.parse_args()

    validator = UIBotValidator()

    # 验证单个文件
    if args.file:
        result = validator.validate_file(args.file)
        if 'error' in result:
            print(f"❌ {result['error']}")
            sys.exit(1)
        validator.print_report(result)

    # 验证目录
    elif args.directory:
        directory = Path(args.directory)
        if not directory.exists():
            print(f"❌ 目录不存在: {directory}")
            sys.exit(1)

        pattern = '**/*.task' if args.recursive else '*.task'
        files = list(directory.glob(pattern))

        if not files:
            print(f"❌ 未找到 .task 文件")
            sys.exit(1)

        print(f"\n找到 {len(files)} 个文件")

        total_issues = 0
        total_warnings = 0

        for file in files:
            result = validator.validate_file(str(file))
            total_issues += len(result['issues'])
            total_warnings += len(result['warnings'])

            # 只显示有问题的文件
            if result['issues'] or result['warnings']:
                validator.print_report(result)

        print("\n" + "=" * 80)
        print(f"总计: {total_issues} 个错误, {total_warnings} 个警告")
        print("=" * 80)

    else:
        parser.print_help()

if __name__ == '__main__':
    main()
