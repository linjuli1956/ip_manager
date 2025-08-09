#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
IP管理器打包脚本
用于将Python程序打包成EXE文件
"""

import os
import sys
import subprocess
import shutil

def install_pyinstaller():
    """安装PyInstaller"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("PyInstaller安装成功")
        return True
    except subprocess.CalledProcessError:
        print("PyInstaller安装失败")
        return False

def build_exe():
    """构建EXE文件"""
    try:
        # 检查PyInstaller是否已安装
        try:
            import PyInstaller
        except ImportError:
            print("PyInstaller未安装，正在安装...")
            if not install_pyinstaller():
                return False
        
        # 构建命令
        cmd = [
            "pyinstaller",
            "--onefile",  # 打包成单个文件
            "--windowed",  # 不显示控制台窗口
            "--name=IP管理器",  # 可执行文件名称
            "--icon=IP管理器.ico",  # 图标文件（如果存在）
            "--add-data=IP管理器.ico;.",  # 添加图标文件到打包中
            "--add-data=ip_manager_256x256.png;.",  # 用于窗口iconphoto
            "main.py"
        ]
        
        # 如果图标或PNG不存在，移除相关参数
        if not os.path.exists("IP管理器.ico"):
            cmd = [arg for arg in cmd if "IP管理器.ico" not in arg]
        if not os.path.exists("ip_manager_256x256.png"):
            cmd = [arg for arg in cmd if "ip_manager_256x256.png" not in arg]
        
        print("开始构建EXE文件...")
        print(f"执行命令: {' '.join(cmd)}")
        
        # 执行构建
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("构建成功！")
            print("EXE文件位置: dist/IP管理器.exe")
            
            # 检查生成的文件
            exe_path = "dist/IP管理器.exe"
            if os.path.exists(exe_path):
                file_size = os.path.getsize(exe_path) / (1024 * 1024)  # MB
                print(f"文件大小: {file_size:.2f} MB")
                
                # 复制到当前目录
                shutil.copy2(exe_path, "IP管理器.exe")
                print("已复制到当前目录: IP管理器.exe")
            
            return True
        else:
            print("构建失败！")
            print("错误信息:")
            print(result.stderr)
            return False
            
    except Exception as e:
        print(f"构建过程中出错: {str(e)}")
        return False

def clean_build():
    """清理构建文件"""
    try:
        dirs_to_remove = ["build", "dist", "__pycache__"]
        files_to_remove = ["IP管理器.spec"]
        
        for dir_name in dirs_to_remove:
            if os.path.exists(dir_name):
                shutil.rmtree(dir_name)
                print(f"已删除目录: {dir_name}")
        
        for file_name in files_to_remove:
            if os.path.exists(file_name):
                os.remove(file_name)
                print(f"已删除文件: {file_name}")
                
        print("清理完成")
        
    except Exception as e:
        print(f"清理过程中出错: {str(e)}")

def main():
    """主函数"""
    print("=" * 50)
    print("Windows IP地址管理器 - 打包工具")
    print("=" * 50)
    
    while True:
        print("\n请选择操作:")
        print("1. 构建EXE文件")
        print("2. 清理构建文件")
        print("3. 退出")
        
        choice = input("\n请输入选择 (1-3): ").strip()
        
        if choice == "1":
            print("\n开始构建...")
            if build_exe():
                print("\n构建完成！")
            else:
                print("\n构建失败！")
                
        elif choice == "2":
            print("\n开始清理...")
            clean_build()
            
        elif choice == "3":
            print("退出程序")
            break
            
        else:
            print("无效选择，请重新输入")

if __name__ == "__main__":
    main() 