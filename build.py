#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VAT验证工具打包脚本
使用PyInstaller将GUI程序打包成可执行文件
"""

import os
import sys
import subprocess
from pathlib import Path

def build_app():
    """打包应用程序"""
    print("开始打包VAT验证工具...")
    
    # 获取当前目录
    current_dir = Path(__file__).parent
    main_script = current_dir / "vat_validator_gui.py"
    
    if not main_script.exists():
        print(f"错误: 找不到主程序文件 {main_script}")
        return False
    
    # PyInstaller命令
    cmd = [
        "pyinstaller",
        "--onefile",  # 打包成单个文件
        "--windowed",  # 不显示控制台窗口
        "--name=VAT验证工具",  # 指定文件名
        "--clean",  # 清理临时文件
        str(main_script)
    ]
    
    try:
        # 执行打包命令
        print("执行PyInstaller命令...")
        result = subprocess.run(cmd, cwd=current_dir)
        
        if result.returncode == 0:
            print("✅ 打包成功!")
            print(f"可执行文件位置: {current_dir / 'dist'}")
            return True
        else:
            print("❌ 打包失败!")
            return False
            
    except Exception as e:
        print(f"打包过程中出现错误: {e}")
        return False

if __name__ == "__main__":
    print("VAT验证工具 - 打包脚本")
    print("=" * 30)
    
    if build_app():
        print("\n🎉 打包完成!")
    else:
        print("\n💥 打包失败，请检查错误信息。")
        sys.exit(1)