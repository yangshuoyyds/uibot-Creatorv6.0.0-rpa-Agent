@echo off
REM UIBot 工具快速启动脚本 (Windows)
chcp 65001 >nul
setlocal enabledelayedexpansion

set SCRIPT_DIR=%~dp0

:LOGO
echo.
echo ╔════════════════════════════════════════════════════════════╗
echo ║                                                            ║
echo ║              UIBot 辅助工具集 v1.0.0                       ║
echo ║                                                            ║
echo ║          快速搜索 ^| 代码生成 ^| 文档查询 ^| 代码验证          ║
echo ║                                                            ║
echo ╚════════════════════════════════════════════════════════════╝
echo.

:MENU
echo.
echo 请选择功能:
echo   1. 命令搜索 (Search)
echo   2. 代码生成 (Generator)
echo   3. 文档查询 (Query)
echo   4. 代码验证 (Validator)
echo   5. 查看帮助
echo   0. 退出
echo.

set /p choice="请输入选项 (0-5): "

if "%choice%"=="1" goto SEARCH
if "%choice%"=="2" goto GENERATOR
if "%choice%"=="3" goto QUERY
if "%choice%"=="4" goto VALIDATOR
if "%choice%"=="5" goto HELP
if "%choice%"=="0" goto EXIT
echo 无效选项，请重新选择
goto MENU

:SEARCH
echo.
echo 启动命令搜索工具...
python "%SCRIPT_DIR%search.py" -i
goto MENU

:GENERATOR
echo.
echo 启动代码生成器...
python "%SCRIPT_DIR%generator.py" -i
goto MENU

:QUERY
echo.
echo 启动文档查询工具...
python "%SCRIPT_DIR%query.py" -i
goto MENU

:VALIDATOR
echo.
echo 代码验证工具
set /p file_path="请输入要验证的文件路径 (或输入 'q' 返回): "
if "%file_path%"=="q" goto MENU
if not exist "%file_path%" (
    echo 错误: 文件不存在
    pause
    goto MENU
)
python "%SCRIPT_DIR%validator.py" "%file_path%"
pause
goto MENU

:HELP
echo.
echo UIBot 工具集使用帮助
echo.
echo 1. 命令搜索 (search.py)
echo    - 快速搜索 UIBot 命令
echo    - 支持关键词和功能描述搜索
echo    - 命令行: python search.py -i
echo.
echo 2. 代码生成 (generator.py)
echo    - 根据模板生成代码
echo    - 支持 6 种常用模板
echo    - 命令行: python generator.py -i
echo.
echo 3. 文档查询 (query.py)
echo    - 交互式文档查询
echo    - 支持多文档搜索
echo    - 命令行: python query.py -i
echo.
echo 4. 代码验证 (validator.py)
echo    - 检查代码质量
echo    - 发现潜在问题
echo    - 命令行: python validator.py ^<file^>
echo.
echo 详细文档: tools\README.md
echo.
pause
goto MENU

:EXIT
echo.
echo 再见！
echo.
exit /b 0
