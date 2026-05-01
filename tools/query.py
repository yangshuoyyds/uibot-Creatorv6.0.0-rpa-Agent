#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
UIBot 交互式查询工具
提供友好的命令行界面查询 UIBot 文档
"""

import os
import sys
import re
from pathlib import Path
from typing import List, Dict

class UIBotQuery:
    def __init__(self, skill_dir: str = None):
        """初始化查询工具"""
        if skill_dir is None:
            skill_dir = Path(__file__).parent.parent
        self.skill_dir = Path(skill_dir)
        self.docs = {}
        self.load_docs()

    def load_docs(self):
        """加载文档"""
        doc_files = {
            'commands': 'commands-reference.md',
            'examples': 'examples.md',
            'quick': 'quick-index.md',
            'templates': 'templates.md',
            'faq': 'faq.md',
            'enterprise': 'enterprise-best-practices.md',
            'patterns': 'design-patterns.md',
            'quickstart': 'quick-start.md'
        }

        for key, filename in doc_files.items():
            filepath = self.skill_dir / filename
            if filepath.exists():
                with open(filepath, 'r', encoding='utf-8') as f:
                    self.docs[key] = f.read()
                print(f"✓ 已加载 {filename}")
            else:
                print(f"⚠ 未找到 {filename}")

    def search_in_doc(self, doc_key: str, keyword: str) -> List[str]:
        """在文档中搜索"""
        if doc_key not in self.docs:
            return []

        content = self.docs[doc_key]
        results = []

        # 按段落分割
        sections = re.split(r'\n#{1,3}\s+', content)

        for section in sections:
            if keyword.lower() in section.lower():
                # 提取标题和内容
                lines = section.split('\n', 1)
                title = lines[0].strip() if lines else ''
                content_preview = lines[1][:200] if len(lines) > 1 else ''

                results.append({
                    'title': title,
                    'preview': content_preview.strip()
                })

        return results

    def get_command_detail(self, command_name: str) -> Dict:
        """获取命令详细信息"""
        if 'commands' not in self.docs:
            return None

        content = self.docs['commands']

        # 查找命令段落
        pattern = rf'###\s+{re.escape(command_name)}.*?\n(.*?)(?=\n###|\Z)'
        match = re.search(pattern, content, re.DOTALL | re.IGNORECASE)

        if match:
            detail_text = match.group(1)
            return {
                'name': command_name,
                'content': detail_text.strip()
            }

        return None

    def get_example(self, keyword: str) -> List[Dict]:
        """获取示例"""
        if 'examples' not in self.docs:
            return []

        content = self.docs['examples']
        results = []

        # 查找示例
        pattern = r'###\s+(.+?)\n(.*?)(?=\n###|\Z)'
        matches = re.finditer(pattern, content, re.DOTALL)

        for match in matches:
            title = match.group(1).strip()
            example_content = match.group(2).strip()

            if keyword.lower() in title.lower() or keyword.lower() in example_content.lower():
                results.append({
                    'title': title,
                    'content': example_content[:500]
                })

        return results

    def get_faq(self, keyword: str) -> List[Dict]:
        """获取常见问题"""
        if 'faq' not in self.docs:
            return []

        content = self.docs['faq']
        results = []

        # 查找问题
        pattern = r'###\s+(Q\d+[:.：].*?)\n(.*?)(?=\n###|\Z)'
        matches = re.finditer(pattern, content, re.DOTALL)

        for match in matches:
            question = match.group(1).strip()
            answer = match.group(2).strip()

            if keyword.lower() in question.lower() or keyword.lower() in answer.lower():
                results.append({
                    'question': question,
                    'answer': answer[:300]
                })

        return results

    def get_template(self, keyword: str) -> List[Dict]:
        """获取代码模板"""
        if 'templates' not in self.docs:
            return []

        content = self.docs['templates']
        results = []

        # 查找模板
        pattern = r'###\s+(.+?)\n(.*?)(?=\n###|\Z)'
        matches = re.finditer(pattern, content, re.DOTALL)

        for match in matches:
            title = match.group(1).strip()
            template_content = match.group(2).strip()

            if keyword.lower() in title.lower() or keyword.lower() in template_content.lower():
                results.append({
                    'title': title,
                    'content': template_content[:400]
                })

        return results

    def show_menu(self):
        """显示主菜单"""
        print("\n" + "=" * 80)
        print("UIBot 交互式查询工具")
        print("=" * 80)
        print("\n功能菜单:")
        print("  1. 搜索命令")
        print("  2. 查看示例")
        print("  3. 查看常见问题")
        print("  4. 查看代码模板")
        print("  5. 快速索引")
        print("  6. 企业级最佳实践")
        print("  7. 全文搜索")
        print("  0. 退出")
        print()

    def interactive(self):
        """交互式查询"""
        while True:
            self.show_menu()

            try:
                choice = input("请选择功能 (0-7): ").strip()

                if choice == '0':
                    print("再见！")
                    break

                elif choice == '1':
                    keyword = input("\n请输入命令关键词: ").strip()
                    results = self.search_in_doc('commands', keyword)
                    self.print_results("命令搜索结果", results)

                elif choice == '2':
                    keyword = input("\n请输入示例关键词: ").strip()
                    results = self.get_example(keyword)
                    self.print_results("示例搜索结果", results)

                elif choice == '3':
                    keyword = input("\n请输入问题关键词: ").strip()
                    results = self.get_faq(keyword)
                    self.print_faq_results(results)

                elif choice == '4':
                    keyword = input("\n请输入模板关键词: ").strip()
                    results = self.get_template(keyword)
                    self.print_results("模板搜索结果", results)

                elif choice == '5':
                    self.show_quick_index()

                elif choice == '6':
                    self.show_enterprise_guide()

                elif choice == '7':
                    keyword = input("\n请输入搜索关键词: ").strip()
                    self.full_text_search(keyword)

                else:
                    print("❌ 无效选择，请重新输入")

                input("\n按回车继续...")

            except KeyboardInterrupt:
                print("\n\n再见！")
                break
            except Exception as e:
                print(f"❌ 错误: {e}")

    def print_results(self, title: str, results: List[Dict]):
        """打印搜索结果"""
        print("\n" + "=" * 80)
        print(title)
        print("=" * 80)

        if not results:
            print("\n❌ 未找到匹配结果")
            return

        print(f"\n找到 {len(results)} 个结果:\n")

        for i, result in enumerate(results[:10], 1):
            print(f"【{i}】 {result.get('title', '无标题')}")
            if 'preview' in result:
                print(f"{result['preview'][:200]}...")
            elif 'content' in result:
                print(f"{result['content'][:200]}...")
            print("-" * 80)

        if len(results) > 10:
            print(f"\n... 还有 {len(results) - 10} 个结果未显示")

    def print_faq_results(self, results: List[Dict]):
        """打印 FAQ 结果"""
        print("\n" + "=" * 80)
        print("常见问题搜索结果")
        print("=" * 80)

        if not results:
            print("\n❌ 未找到匹配结果")
            return

        print(f"\n找到 {len(results)} 个问题:\n")

        for i, result in enumerate(results[:10], 1):
            print(f"【{i}】 {result['question']}")
            print(f"\n{result['answer'][:300]}...")
            print("-" * 80)

    def show_quick_index(self):
        """显示快速索引"""
        if 'quick' not in self.docs:
            print("❌ 快速索引文档未加载")
            return

        print("\n" + "=" * 80)
        print("快速索引")
        print("=" * 80)

        # 提取主要分类
        content = self.docs['quick']
        sections = re.findall(r'##\s+(.+?)\n', content)

        print("\n可用分类:")
        for i, section in enumerate(sections[:15], 1):
            print(f"  {i}. {section}")

        print("\n提示: 使用全文搜索功能查看详细内容")

    def show_enterprise_guide(self):
        """显示企业级指南"""
        if 'enterprise' not in self.docs:
            print("❌ 企业级文档未加载")
            return

        print("\n" + "=" * 80)
        print("企业级最佳实践")
        print("=" * 80)

        content = self.docs['enterprise']
        sections = re.findall(r'##\s+(.+?)\n', content)

        print("\n主要内容:")
        for i, section in enumerate(sections[:10], 1):
            print(f"  {i}. {section}")

        print("\n提示: 使用全文搜索功能查看详细内容")

    def full_text_search(self, keyword: str):
        """全文搜索"""
        print("\n" + "=" * 80)
        print(f"全文搜索: {keyword}")
        print("=" * 80)

        all_results = []

        for doc_key, doc_content in self.docs.items():
            results = self.search_in_doc(doc_key, keyword)
            if results:
                all_results.append({
                    'doc': doc_key,
                    'count': len(results),
                    'results': results[:3]
                })

        if not all_results:
            print("\n❌ 未找到匹配结果")
            return

        print(f"\n在 {len(all_results)} 个文档中找到结果:\n")

        for item in all_results:
            doc_names = {
                'commands': '命令参考',
                'examples': '实战示例',
                'quick': '快速索引',
                'templates': '代码模板',
                'faq': '常见问题',
                'enterprise': '企业级实践',
                'patterns': '设计模式',
                'quickstart': '快速入门'
            }

            print(f"【{doc_names.get(item['doc'], item['doc'])}】 - {item['count']} 个匹配")

            for result in item['results']:
                print(f"  • {result.get('title', '无标题')}")

            print()

def main():
    """主函数"""
    import argparse

    parser = argparse.ArgumentParser(description='UIBot 交互式查询工具')
    parser.add_argument('-s', '--search', help='搜索关键词')
    parser.add_argument('-t', '--type', choices=['command', 'example', 'faq', 'template'],
                       help='搜索类型')
    parser.add_argument('-i', '--interactive', action='store_true', help='交互模式')

    args = parser.parse_args()

    # 初始化查询工具
    query = UIBotQuery()

    # 交互模式
    if args.interactive or not args.search:
        query.interactive()
        return

    # 命令行搜索
    if args.search:
        if args.type == 'command':
            results = query.search_in_doc('commands', args.search)
            query.print_results("命令搜索结果", results)
        elif args.type == 'example':
            results = query.get_example(args.search)
            query.print_results("示例搜索结果", results)
        elif args.type == 'faq':
            results = query.get_faq(args.search)
            query.print_faq_results(results)
        elif args.type == 'template':
            results = query.get_template(args.search)
            query.print_results("模板搜索结果", results)
        else:
            query.full_text_search(args.search)

if __name__ == '__main__':
    main()
