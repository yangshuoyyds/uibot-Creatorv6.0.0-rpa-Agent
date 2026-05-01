#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
UIBot 命令搜索工具
支持模糊搜索、分类筛选、正则匹配
"""

import re
import os
import sys
import json
from pathlib import Path
from typing import List, Dict, Tuple

class UIBotCommandSearch:
    def __init__(self, skill_dir: str = None):
        """初始化搜索工具"""
        if skill_dir is None:
            skill_dir = Path(__file__).parent.parent
        self.skill_dir = Path(skill_dir)
        self.commands = []
        self.categories = {}
        self.load_commands()

    def load_commands(self):
        """从文档中加载命令"""
        uibot_file = self.skill_dir / "uibot.md"
        if not uibot_file.exists():
            print(f"错误: 找不到 uibot.md 文件: {uibot_file}")
            return

        with open(uibot_file, 'r', encoding='utf-8') as f:
            content = f.read()

        # 解析命令
        pattern = r'##\s+(.+?)\n\n\*\*说明\*\*:\s*(.+?)\s*\n\n\*\*原型\*\*:\s*`(.+?)`'
        matches = re.finditer(pattern, content, re.DOTALL)

        for match in matches:
            name = match.group(1).strip()
            description = match.group(2).strip()
            prototype = match.group(3).strip()

            self.commands.append({
                'name': name,
                'description': description,
                'prototype': prototype
            })

        print(f"✓ 已加载 {len(self.commands)} 个命令")

    def search(self, keyword: str, category: str = None, use_regex: bool = False) -> List[Dict]:
        """搜索命令"""
        results = []

        for cmd in self.commands:
            # 分类筛选
            if category and category.lower() not in cmd['name'].lower():
                continue

            # 关键词匹配
            if use_regex:
                pattern = re.compile(keyword, re.IGNORECASE)
                if pattern.search(cmd['name']) or pattern.search(cmd['description']):
                    results.append(cmd)
            else:
                keyword_lower = keyword.lower()
                if (keyword_lower in cmd['name'].lower() or
                    keyword_lower in cmd['description'].lower() or
                    keyword_lower in cmd['prototype'].lower()):
                    results.append(cmd)

        return results

    def search_by_function(self, function_desc: str) -> List[Dict]:
        """根据功能描述搜索命令"""
        keywords_map = {
            '点击': ['click', 'mouse', '鼠标'],
            '输入': ['input', 'type', 'keyboard', '键盘'],
            '浏览器': ['browser', 'web', 'chrome'],
            'excel': ['excel', 'xls', 'xlsx', '表格'],
            '文件': ['file', 'folder', '目录'],
            '邮件': ['mail', 'email', 'smtp'],
            '数据库': ['database', 'sql', 'db'],
            '图像': ['image', 'ocr', '识别'],
            '窗口': ['window', 'win', '界面'],
            '等待': ['wait', 'sleep', 'delay']
        }

        # 提取关键词
        keywords = []
        for key, values in keywords_map.items():
            if key in function_desc.lower():
                keywords.extend(values)

        if not keywords:
            keywords = [function_desc]

        # 搜索
        results = []
        for keyword in keywords:
            results.extend(self.search(keyword))

        # 去重
        seen = set()
        unique_results = []
        for cmd in results:
            if cmd['name'] not in seen:
                seen.add(cmd['name'])
                unique_results.append(cmd)

        return unique_results

    def print_results(self, results: List[Dict], max_results: int = 10):
        """打印搜索结果"""
        if not results:
            print("\n❌ 未找到匹配的命令")
            return

        print(f"\n✓ 找到 {len(results)} 个匹配命令")
        print("=" * 80)

        for i, cmd in enumerate(results[:max_results], 1):
            print(f"\n【{i}】 {cmd['name']}")
            print(f"说明: {cmd['description']}")
            print(f"原型: {cmd['prototype']}")
            print("-" * 80)

        if len(results) > max_results:
            print(f"\n... 还有 {len(results) - max_results} 个结果未显示")

    def interactive_search(self):
        """交互式搜索"""
        print("\n" + "=" * 80)
        print("UIBot 命令搜索工具 - 交互模式")
        print("=" * 80)
        print("\n命令:")
        print("  search <关键词>     - 搜索命令")
        print("  func <功能描述>     - 根据功能搜索")
        print("  list <分类>         - 列出分类下的命令")
        print("  help                - 显示帮助")
        print("  quit / exit         - 退出")
        print()

        while True:
            try:
                user_input = input(">>> ").strip()

                if not user_input:
                    continue

                if user_input.lower() in ['quit', 'exit', 'q']:
                    print("再见！")
                    break

                if user_input.lower() == 'help':
                    print("\n使用示例:")
                    print("  search 点击        - 搜索包含'点击'的命令")
                    print("  func 如何点击按钮  - 根据功能描述搜索")
                    print("  list mouse         - 列出鼠标相关命令")
                    continue

                parts = user_input.split(maxsplit=1)
                if len(parts) < 2:
                    print("❌ 请输入命令和参数，例如: search 点击")
                    continue

                command, arg = parts

                if command.lower() == 'search':
                    results = self.search(arg)
                    self.print_results(results)

                elif command.lower() == 'func':
                    results = self.search_by_function(arg)
                    self.print_results(results)

                elif command.lower() == 'list':
                    results = self.search(arg)
                    self.print_results(results, max_results=20)

                else:
                    print(f"❌ 未知命令: {command}")

            except KeyboardInterrupt:
                print("\n\n再见！")
                break
            except Exception as e:
                print(f"❌ 错误: {e}")

def main():
    """主函数"""
    import argparse

    parser = argparse.ArgumentParser(description='UIBot 命令搜索工具')
    parser.add_argument('keyword', nargs='?', help='搜索关键词')
    parser.add_argument('-f', '--function', help='根据功能描述搜索')
    parser.add_argument('-c', '--category', help='按分类筛选')
    parser.add_argument('-r', '--regex', action='store_true', help='使用正则表达式')
    parser.add_argument('-i', '--interactive', action='store_true', help='交互模式')
    parser.add_argument('-n', '--number', type=int, default=10, help='显示结果数量')

    args = parser.parse_args()

    # 初始化搜索工具
    searcher = UIBotCommandSearch()

    # 交互模式
    if args.interactive or (not args.keyword and not args.function):
        searcher.interactive_search()
        return

    # 命令行搜索
    if args.function:
        results = searcher.search_by_function(args.function)
    elif args.keyword:
        results = searcher.search(args.keyword, args.category, args.regex)
    else:
        parser.print_help()
        return

    searcher.print_results(results, args.number)

if __name__ == '__main__':
    main()
