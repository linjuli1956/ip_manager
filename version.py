#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
版本信息管理文件
"""

# 版本信息
VERSION = "v1.3.1"
BUILD_DATE = "2025-08-08"
BUILD_TIME = "22:00"

# 版本描述
VERSION_DESCRIPTION = "Windows IP地址管理器 - IP验证优化版本"

# 更新日志
CHANGELOG = """
v1.3.1 (2025-08-08)
- 完善IP地址验证功能
- 实时验证IP地址、子网掩码、网关、DNS
- 每个段严格限制在0-255范围内
- 提供详细的错误提示信息
- 优化用户体验和输入安全性

v1.3.0 (2025-08-08)
- 界面字体全面放大（9号→11号）
- 标题字体增大（16号→18号）
- 修复按钮悬停时颜色变化问题
- 优化整体视觉效果和可读性

v1.0.0 (2025-08-08)
- 初始版本发布
- 支持Windows IP地址管理
- 支持网络适配器控制
- 支持静态IP和DHCP配置
- 支持多IP地址配置
- 支持网络重置功能
- 添加版本号显示
"""

def get_version_info():
    """获取版本信息"""
    return {
        'version': VERSION,
        'build_date': BUILD_DATE,
        'build_time': BUILD_TIME,
        'description': VERSION_DESCRIPTION
    }

def get_version_string():
    """获取版本字符串"""
    return f"{VERSION} (更新日期: {BUILD_DATE})"

def get_changelog():
    """获取更新日志"""
    return CHANGELOG

if __name__ == "__main__":
    print(f"版本信息: {get_version_string()}")
    print(f"版本描述: {VERSION_DESCRIPTION}")
    print("\n更新日志:")
    print(CHANGELOG) 