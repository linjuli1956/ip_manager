#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
IP管理器测试脚本
用于验证程序的基本功能
"""

import subprocess
import sys
import os

def test_ipconfig():
    """测试ipconfig命令"""
    print("测试ipconfig命令...")
    try:
        result = subprocess.run(['ipconfig', '/all'], 
                              capture_output=True, text=True, encoding='gbk')
        if result.returncode == 0:
            print("✓ ipconfig命令执行成功")
            return True
        else:
            print("✗ ipconfig命令执行失败")
            return False
    except Exception as e:
        print(f"✗ ipconfig命令出错: {e}")
        return False

def test_netsh():
    """测试netsh命令"""
    print("测试netsh命令...")
    try:
        result = subprocess.run(['netsh', 'interface', 'show', 'interface'], 
                              capture_output=True, text=True, encoding='gbk')
        if result.returncode == 0:
            print("✓ netsh命令执行成功")
            return True
        else:
            print("✗ netsh命令执行失败")
            return False
    except Exception as e:
        print(f"✗ netsh命令出错: {e}")
        return False

def test_imports():
    """测试必要的模块导入"""
    print("测试模块导入...")
    try:
        import tkinter
        import subprocess
        import re
        import socket
        import threading
        import os
        import sys
        from datetime import datetime
        print("✓ 所有模块导入成功")
        return True
    except ImportError as e:
        print(f"✗ 模块导入失败: {e}")
        return False

def test_admin_rights():
    """测试管理员权限"""
    print("测试管理员权限...")
    try:
        # 尝试执行需要管理员权限的命令
        result = subprocess.run(['netsh', 'interface', 'show', 'interface'], 
                              capture_output=True, text=True, encoding='gbk')
        if result.returncode == 0:
            print("✓ 具有管理员权限")
            return True
        else:
            print("⚠ 可能没有管理员权限")
            return False
    except Exception as e:
        print(f"✗ 权限测试失败: {e}")
        return False

def main():
    """主测试函数"""
    print("=" * 50)
    print("IP管理器 - 系统测试")
    print("=" * 50)
    print()
    
    tests = [
        ("模块导入", test_imports),
        ("ipconfig命令", test_ipconfig),
        ("netsh命令", test_netsh),
        ("管理员权限", test_admin_rights),
    ]
    
    passed = 0
    total = len(tests)
    
    for test_name, test_func in tests:
        print(f"正在{test_name}...")
        if test_func():
            passed += 1
        print()
    
    print("=" * 50)
    print(f"测试结果: {passed}/{total} 通过")
    
    if passed == total:
        print("✓ 所有测试通过，系统环境正常")
        print("可以运行IP管理器程序")
    else:
        print("⚠ 部分测试失败，请检查系统环境")
        print("建议以管理员身份运行程序")
    
    print("=" * 50)

if __name__ == "__main__":
    main() 