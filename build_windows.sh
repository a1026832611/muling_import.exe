#!/bin/bash
# ============================================================
# Windows 打包脚本
# 请在 Windows 的 Git Bash 中执行
# 前提：已安装 Python 3.8+ 和 pip
# ============================================================

set -euo pipefail

if ! command -v python >/dev/null 2>&1; then
    echo "未找到 python，请先安装 Python 3.8+ 并加入 PATH"
    exit 1
fi

if [ ! -d "web" ]; then
    echo "缺少 web 目录，无法打包"
    exit 1
fi

if [ ! -f "file/addPerson.xlsx" ]; then
    echo "缺少 file/addPerson.xlsx 模板文件，无法打包"
    exit 1
fi

echo "===== 1. 创建虚拟环境 ====="
if [ ! -d "venv" ]; then
    python -m venv venv
fi
source venv/Scripts/activate

echo "===== 2. 安装依赖 ====="
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo "===== 3. 清理旧的构建产物 ====="
rm -rf build dist *.spec

echo "===== 4. 使用 PyInstaller 打包 ====="
python -m PyInstaller \
    --noconfirm \
    --clean \
    --onefile \
    --windowed \
    --name "医护人员批量添加" \
    --add-data "web;web" \
    --add-data "file;file" \
    main.py

echo "===== 5. 打包完成 ====="
echo "生成的 exe 文件位于: dist/医护人员批量添加.exe"
echo "可以直接双击运行"
