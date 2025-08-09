#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
版本信息管理文件
"""

# 版本信息
VERSION = "v1.3.5"
BUILD_DATE = "2025-08-09"
BUILD_TIME = "23:00"

# 版本描述
VERSION_DESCRIPTION = "Windows IP地址管理器 - 诊断工具增强版本"

# 更新日志
CHANGELOG = """
v1.3.5 (2025-08-09)
- “网络诊断”新增“网络追踪（tracert）”，默认目标 www.baidu.com，并在新 CMD 窗口执行
- 与“内网测试（ping）/网络测试（ping）”统一布局与样式

v1.3.4 (2025-08-09)
- 启动时自动尝试管理员提权（UAC），避免权限相关功能失败
- 防火墙 ping 规则支持智能判断：优先启用、不存在则自动创建，并校验启用状态
- Win11 自动登录开关：设置后自动打开 netplwiz，便于一键完成配置

v1.3.3 (2025-08-09)
- 工具页重构为“网络诊断/系统工具”分组，布局更清晰
- 新增快速工具：刷新DNS、IP释放、IP续租、Winsock重置、打开防火墙ping
- “一键清理浏览器缓存”纳入系统工具分组，带进度条且不清除账号密码
- 统一子进程调用为静默执行，修复UnicodeDecodeError (gbk) 异常
- 网络测试区域按钮与输入框对齐优化

v1.3.2 (2025-08-09)
- 更换应用图标为可爱的二次元风格，并同步更新窗口内左上角图标
- 右侧布局进一步收窄，左侧保持不变
- 新增内网测试（ping）与外网测试（ping）工具：
  - 内网测试默认取所选网卡的网关地址
  - 点击按钮在新cmd窗口执行 ping <目标> -t

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