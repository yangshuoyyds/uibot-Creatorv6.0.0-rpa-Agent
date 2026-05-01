#!/bin/bash
# UIBot 工具快速启动脚本

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 显示 Logo
show_logo() {
    echo -e "${BLUE}"
    echo "╔════════════════════════════════════════════════════════════╗"
    echo "║                                                            ║"
    echo "║              UIBot 辅助工具集 v1.0.0                       ║"
    echo "║                                                            ║"
    echo "║          快速搜索 | 代码生成 | 文档查询 | 代码验证          ║"
    echo "║                                                            ║"
    echo "╚════════════════════════════════════════════════════════════╝"
    echo -e "${NC}"
}

# 显示菜单
show_menu() {
    echo -e "\n${GREEN}请选择功能:${NC}"
    echo "  1. 命令搜索 (Search)"
    echo "  2. 代码生成 (Generator)"
    echo "  3. 文档查询 (Query)"
    echo "  4. 代码验证 (Validator)"
    echo "  5. 查看帮助"
    echo "  0. 退出"
    echo ""
}

# 命令搜索
run_search() {
    echo -e "\n${YELLOW}启动命令搜索工具...${NC}"
    python3 "$SCRIPT_DIR/search.py" -i
}

# 代码生成
run_generator() {
    echo -e "\n${YELLOW}启动代码生成器...${NC}"
    python3 "$SCRIPT_DIR/generator.py" -i
}

# 文档查询
run_query() {
    echo -e "\n${YELLOW}启动文档查询工具...${NC}"
    python3 "$SCRIPT_DIR/query.py" -i
}

# 代码验证
run_validator() {
    echo -e "\n${YELLOW}代码验证工具${NC}"
    echo "请输入要验证的文件路径 (或输入 'q' 返回):"
    read -r file_path

    if [ "$file_path" = "q" ]; then
        return
    fi

    if [ -f "$file_path" ]; then
        python3 "$SCRIPT_DIR/validator.py" "$file_path"
    else
        echo -e "${RED}错误: 文件不存在${NC}"
    fi

    echo -e "\n按回车继续..."
    read
}

# 显示帮助
show_help() {
    echo -e "\n${GREEN}UIBot 工具集使用帮助${NC}"
    echo ""
    echo "1. 命令搜索 (search.py)"
    echo "   - 快速搜索 UIBot 命令"
    echo "   - 支持关键词和功能描述搜索"
    echo "   - 命令行: python3 search.py -i"
    echo ""
    echo "2. 代码生成 (generator.py)"
    echo "   - 根据模板生成代码"
    echo "   - 支持 6 种常用模板"
    echo "   - 命令行: python3 generator.py -i"
    echo ""
    echo "3. 文档查询 (query.py)"
    echo "   - 交互式文档查询"
    echo "   - 支持多文档搜索"
    echo "   - 命令行: python3 query.py -i"
    echo ""
    echo "4. 代码验证 (validator.py)"
    echo "   - 检查代码质量"
    echo "   - 发现潜在问题"
    echo "   - 命令行: python3 validator.py <file>"
    echo ""
    echo "详细文档: tools/README.md"
    echo ""
    echo -e "按回车继续..."
    read
}

# 主函数
main() {
    show_logo

    while true; do
        show_menu
        read -p "请输入选项 (0-5): " choice

        case $choice in
            1)
                run_search
                ;;
            2)
                run_generator
                ;;
            3)
                run_query
                ;;
            4)
                run_validator
                ;;
            5)
                show_help
                ;;
            0)
                echo -e "\n${GREEN}再见！${NC}\n"
                exit 0
                ;;
            *)
                echo -e "${RED}无效选项，请重新选择${NC}"
                ;;
        esac
    done
}

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}错误: 未找到 Python 3${NC}"
    echo "请安装 Python 3.6 或更高版本"
    exit 1
fi

# 运行主程序
main
