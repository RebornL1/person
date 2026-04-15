#!/usr/bin/env python3
"""
跨平台启动脚本 - 自动检测虚拟环境并启动服务
支持 Windows、macOS、Linux
"""

import os
import sys
import subprocess
from pathlib import Path

# 项目根目录
BASE_DIR = Path(__file__).resolve().parent
VENV_DIR = BASE_DIR / ".venv"

def get_python_executable():
    """获取虚拟环境中的Python解释器路径"""
    if os.name == "nt":  # Windows
        return VENV_DIR / "Scripts" / "python.exe"
    else:  # macOS / Linux
        return VENV_DIR / "bin" / "python"

def get_uvicorn_executable():
    """获取uvicorn可执行文件路径"""
    if os.name == "nt":  # Windows
        return VENV_DIR / "Scripts" / "uvicorn.exe"
    else:  # macOS / Linux
        return VENV_DIR / "bin" / "uvicorn"

def check_venv_exists():
    """检查虚拟环境是否存在"""
    return VENV_DIR.exists() and get_python_executable().exists()

def create_venv():
    """创建虚拟环境"""
    print("正在创建虚拟环境...")
    subprocess.run([sys.executable, "-m", "venv", str(VENV_DIR)], check=True)
    print("虚拟环境创建完成！")

def install_dependencies():
    """安装依赖"""
    print("正在安装依赖...")
    python_exe = str(get_python_executable())
    subprocess.run([python_exe, "-m", "pip", "install", "--upgrade", "pip"], check=True)
    subprocess.run([python_exe, "-m", "pip", "install", "-r", str(BASE_DIR / "requirements.txt")], check=True)
    print("依赖安装完成！")

def start_server(host="127.0.0.1", port=8000, reload=True):
    """启动服务"""
    python_exe = str(get_python_executable())
    
    # 检查uvicorn是否存在
    uvicorn_exe = get_uvicorn_executable()
    if not uvicorn_exe.exists():
        print("uvicorn未安装，正在安装...")
        subprocess.run([python_exe, "-m", "pip", "install", "uvicorn"], check=True)
    
    print(f"\n启动服务: http://{host}:{port}/")
    print("按 Ctrl+C 停止服务\n")
    
    # 使用python -m uvicorn启动，更可靠
    cmd = [
        python_exe, "-m", "uvicorn",
        "app:app",
        "--host", host,
        "--port", str(port),
    ]
    if reload:
        cmd.append("--reload")
    
    # 设置工作目录
    os.chdir(BASE_DIR)
    
    try:
        subprocess.run(cmd)
    except KeyboardInterrupt:
        print("\n服务已停止")

def main():
    import argparse
    parser = argparse.ArgumentParser(description="J1 Excel分析服务启动脚本")
    parser.add_argument("--install", action="store_true", help="仅安装依赖，不启动服务")
    parser.add_argument("--host", default="127.0.0.1", help="服务主机地址 (默认: 127.0.0.1)")
    parser.add_argument("--port", type=int, default=8000, help="服务端口 (默认: 8000)")
    parser.add_argument("--no-reload", action="store_true", help="禁用自动重载（生产环境）")
    args = parser.parse_args()
    
    print("=" * 50)
    print("J1 - Excel 工作量分析 Web 应用")
    print("=" * 50)
    
    # 检查虚拟环境
    if not check_venv_exists():
        print("虚拟环境不存在，正在初始化...")
        create_venv()
        install_dependencies()
    else:
        # 检查依赖是否需要更新
        try:
            python_exe = str(get_python_executable())
            subprocess.run([python_exe, "-c", "import fastapi"], check=True, capture_output=True)
        except subprocess.CalledProcessError:
            print("依赖未安装，正在安装...")
            install_dependencies()
    
    if args.install:
        print("\n依赖安装完成！可以运行以下命令启动服务：")
        if os.name == "nt":
            print("  python start.py")
        else:
            print("  python3 start.py")
        return
    
    # 启动服务
    start_server(
        host=args.host,
        port=args.port,
        reload=not args.no_reload
    )

if __name__ == "__main__":
    main()