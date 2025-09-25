#!/usr/bin/env python3
"""
Windows EXE构建脚本
用于在本地或CI环境中构建Windows可执行文件
"""

import os
import sys
import subprocess
import platform

def main():
    """主函数"""
    print("🚀 开始构建Windows EXE...")
    
    # 检查当前目录
    if not os.path.exists('vat_validator_gui.py'):
        print("❌ 错误：找不到 vat_validator_gui.py 文件")
        print("请确保在项目根目录运行此脚本")
        sys.exit(1)
    
    # 检查spec文件
    spec_files = ["VAT验证工具-Windows.spec", "VAT验证工具.spec"]
    spec_file = None
    
    for file in spec_files:
        if os.path.exists(file):
            spec_file = file
            print(f"✅ 找到spec文件：{spec_file}")
            break
    
    if not spec_file:
        print("❌ 错误：找不到任何spec文件")
        print("请确保存在以下文件之一：")
        for file in spec_files:
            print(f"  - {file}")
        sys.exit(1)
    
    # 检查PyInstaller
    try:
        subprocess.run(['pyinstaller', '--version'], 
                      check=True, capture_output=True)
        print("✅ PyInstaller 已安装")
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("❌ 错误：PyInstaller 未安装")
        print("请运行：pip install pyinstaller")
        sys.exit(1)
    
    # 清理之前的构建
    print("🧹 清理之前的构建文件...")
    for dir_name in ['build', 'dist']:
        if os.path.exists(dir_name):
            import shutil
            shutil.rmtree(dir_name)
            print(f"   删除 {dir_name} 目录")
    
    # 构建EXE
    print("🔨 开始构建...")
    try:
        cmd = ['pyinstaller', spec_file]
        print(f"执行命令：{' '.join(cmd)}")
        
        result = subprocess.run(cmd, check=True, capture_output=False)
        
        print("✅ 构建成功！")
        
        # 检查输出文件
        exe_path = os.path.join('dist', 'VAT验证工具.exe')
        if os.path.exists(exe_path):
            size = os.path.getsize(exe_path)
            print(f"📦 生成的EXE文件：{exe_path}")
            print(f"📏 文件大小：{size / 1024 / 1024:.1f} MB")
        else:
            print("⚠️  警告：未找到生成的EXE文件")
            
    except subprocess.CalledProcessError as e:
        print(f"❌ 构建失败：{e}")
        sys.exit(1)
    
    print("\n🎉 构建完成！")
    print("生成的文件位于 dist/ 目录中")

if __name__ == '__main__':
    main()