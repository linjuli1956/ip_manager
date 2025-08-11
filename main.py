import os
import sys

# 在导入tkinter之前设置环境变量
def setup_tcl_tk():
    """设置Tcl/Tk环境变量"""
    # 设置Tcl/Tk库路径
    python_dir = os.path.dirname(sys.executable)
    tcl_dir = os.path.join(python_dir, 'tcl')
    tcl86_dir = os.path.join(tcl_dir, 'tcl8.6')
    tk86_dir = os.path.join(tcl_dir, 'tk8.6')
    
    # 检查并设置TCL_LIBRARY
    if os.path.exists(tcl86_dir):
        os.environ['TCL_LIBRARY'] = tcl86_dir
        print(f"设置 TCL_LIBRARY: {tcl86_dir}")
    else:
        print(f"警告：Tcl8.6目录不存在: {tcl86_dir}")
    
    # 检查并设置TK_LIBRARY
    if os.path.exists(tk86_dir):
        os.environ['TK_LIBRARY'] = tk86_dir
        print(f"设置 TK_LIBRARY: {tk86_dir}")
    else:
        print(f"警告：Tk8.6目录不存在: {tk86_dir}")
    
    # 设置TCL_LIBRARY_PATH和TK_LIBRARY_PATH
    if os.path.exists(tcl_dir):
        os.environ['TCL_LIBRARY_PATH'] = tcl_dir
        os.environ['TK_LIBRARY_PATH'] = tcl_dir
        print(f"设置 TCL_LIBRARY_PATH: {tcl_dir}")
        print(f"设置 TK_LIBRARY_PATH: {tcl_dir}")

# 设置环境变量
setup_tcl_tk()

# 现在导入tkinter
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.font as tkfont
 
# Global theme constants (Light, modern)
PRIMARY_BG = '#F7FAFC'   # light background
SURFACE_BG = '#FFFFFF'   # surface background
SURFACE_CARD = '#FFFFFF' # card surface
PRIMARY_FG = '#0F172A'   # dark text
SUBTLE = '#6B7280'       # neutral gray
ACCENT = '#0A84FF'       # blue accent
ACCENT_HOVER = '#0063E1'
ACCENT_PRESSED = '#0052BF'
import subprocess
import winreg
import glob
import re
import socket
import threading
import time
from datetime import datetime
import wmi
import win32com.client
import math
import platform
from typing import Optional, List, Tuple

# 系统托盘相关
try:
    import pystray
    from PIL import Image, ImageDraw
    SYSTEM_TRAY_AVAILABLE = True
except ImportError:
    SYSTEM_TRAY_AVAILABLE = False
    print("警告：pystray或PIL未安装，系统托盘功能不可用")



# 导入版本信息
try:
    from version import VERSION, BUILD_DATE, get_version_string
except ImportError:
    # 如果version.py不存在，使用默认值
    VERSION = "v1.0.0"
    BUILD_DATE = "2025-08-08"
    def get_version_string():
        return f"{VERSION} (构建日期: {BUILD_DATE})"

def resource_path(relative_path: str) -> str:
    """获取打包后资源的真实路径（兼容PyInstaller）。"""
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class IPManager:
    
    def __init__(self, root):
        self.root = root
        self.root.title("Windows IP地址管理器")
        self.root.geometry("820x750")  # 默认窗口大小调整为 820x750
        self.root.resizable(True, True)
        
        # 系统托盘相关变量
        self.is_minimized_to_tray = False
        self.tray_icon = None
        self.first_close_asked = False  # 标记是否已经询问过第一次关闭
        
        # 动态缩放系统
        self.base_width = 900  # 基准宽度
        self.base_height = 680  # 基准高度
        self.scale_factor = 1.0  # 缩放因子
        
        # 存储需要缩放的UI元素
        self.scalable_widgets = []
        
        # 初始化WMI连接
        try:
            self.wmi = wmi.WMI()
        except Exception as e:
            print(f"WMI初始化失败: {e}")
            self.wmi = None
        
        # 绑定窗口大小变化事件
        self.root.bind('<Configure>', self._on_window_resize)
        
        # 初始化缩放因子
        self._update_scale_factor()
        
        # 设置UI
        self.setup_ui()
        
        # 初始化网络适配器
        self.refresh_network_adapters()
        
        # 初始化硬件监控
        self._init_lhm_bridge()
        
        # 绑定硬件信息页签切换事件
        self._bind_hw_tab_events()
        
        # 系统托盘相关
        self.tray_icon = None
        self.is_minimized_to_tray = False
        
        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # 确保任务栏图标显示
        self._ensure_taskbar_icon()
        
        # 初始化系统托盘
        self._init_system_tray()

        # 设置应用图标（EXE与窗口内一致）
        try:
            # 优先使用ico文件
            ico_path = resource_path("ip_manager.ico")
            if os.path.exists(ico_path):
                self.root.iconbitmap(ico_path)
                print(f"设置窗口图标: {ico_path}")
            else:
                # 备用方案：尝试其他可能的ico文件名
                for ico_name in ["IP管理器.ico", "icon.ico", "app.ico"]:
                    ico_path = resource_path(ico_name)
                    if os.path.exists(ico_path):
                        self.root.iconbitmap(ico_path)
                        print(f"设置窗口图标: {ico_path}")
                        break
        except Exception as e:
            print(f"设置ico图标失败: {e}")

        try:
            # 设置PhotoImage图标（用于任务栏）
            png_path = resource_path("ip_manager_256x256.png")
            if os.path.exists(png_path):
                # 保持引用避免被GC
                self._icon_img = tk.PhotoImage(file=png_path)
                self.root.iconphoto(True, self._icon_img)
                print(f"设置PhotoImage图标: {png_path}")
            else:
                # 备用方案：尝试其他尺寸的png文件
                for size in ["512x512", "128x128", "64x64", "32x32"]:
                    png_path = resource_path(f"ip_manager_{size}.png")
                    if os.path.exists(png_path):
                        self._icon_img = tk.PhotoImage(file=png_path)
                        self.root.iconphoto(True, self._icon_img)
                        print(f"设置PhotoImage图标: {png_path}")
                        break
        except Exception as e:
            print(f"设置PhotoImage图标失败: {e}")
        
        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')
        
        # 自定义样式 - 科技风配色
        style.configure('.', background=PRIMARY_BG, foreground=PRIMARY_FG)
        style.configure('TFrame', background=PRIMARY_BG)
        style.configure('TLabel', background=PRIMARY_BG, foreground=PRIMARY_FG, font=('Segoe UI', 11))
        style.configure('TLabelframe', background=PRIMARY_BG, borderwidth=1, relief='groove')
        style.configure('TLabelframe.Label', background=PRIMARY_BG, foreground=SUBTLE, font=('Segoe UI', 11, 'bold'))
        style.configure('TEntry', padding=6, fieldbackground='#FFFFFF')
        style.configure('TCombobox', padding=4, fieldbackground='#FFFFFF')
        style.configure('TButton', padding=(12, 7), relief='flat', background=ACCENT, foreground='#FFFFFF', font=('Segoe UI', 11, 'bold'))
        style.map('TButton', background=[('active', ACCENT_HOVER), ('pressed', ACCENT_PRESSED)])
        
        # 自定义按钮样式 - 增大字体并添加悬停效果
        # 刷新按钮 - 绿色，悬停时稍微变亮
        style.map('Refresh.TButton', 
                 background=[('active', '#66BB6A'), ('pressed', '#4CAF50')], 
                 foreground=[('active', 'white'), ('pressed', 'white')])
        style.configure('Refresh.TButton', font=('Arial', 11), background='#4CAF50', foreground='white')
        
        # 设置静态IP - 蓝色，悬停时稍微变亮
        style.map('StaticIP.TButton', 
                 background=[('active', '#42A5F5'), ('pressed', '#2196F3')], 
                 foreground=[('active', 'white'), ('pressed', 'white')])
        style.configure('StaticIP.TButton', font=('Arial', 11), background='#2196F3', foreground='white')
        
        # 设置DHCP - 橙色，悬停时稍微变亮
        style.map('DHCP.TButton', 
                 background=[('active', '#FFB74D'), ('pressed', '#FF9800')], 
                 foreground=[('active', 'white'), ('pressed', 'white')])
        style.configure('DHCP.TButton', font=('Arial', 11), background='#FF9800', foreground='white')
        
        # 刷新信息 - 紫色，悬停时稍微变亮
        style.map('RefreshInfo.TButton', 
                 background=[('active', '#BA68C8'), ('pressed', '#9C27B0')], 
                 foreground=[('active', 'white'), ('pressed', 'white')])
        style.configure('RefreshInfo.TButton', font=('Arial', 11), background='#9C27B0', foreground='white')
        
        # 导出配置 - 蓝灰色，悬停时稍微变亮
        style.map('Export.TButton', 
                 background=[('active', '#90A4AE'), ('pressed', '#607D8B')], 
                 foreground=[('active', 'white'), ('pressed', 'white')])
        style.configure('Export.TButton', font=('Arial', 11), background='#607D8B', foreground='white')
        
        # 添加IP - 绿色，悬停时稍微变亮
        style.map('AddIP.TButton', 
                 background=[('active', '#66BB6A'), ('pressed', '#4CAF50')], 
                 foreground=[('active', 'white'), ('pressed', 'white')])
        style.configure('AddIP.TButton', font=('Arial', 11), background='#4CAF50', foreground='white')
        
        # 清空/删除 - 红色，悬停时稍微变亮
        style.map('Clear.TButton', 
                 background=[('active', '#EF5350'), ('pressed', '#F44336')], 
                 foreground=[('active', 'white'), ('pressed', 'white')])
        style.configure('Clear.TButton', font=('Arial', 11), background='#F44336', foreground='white')
        
        # 禁用网卡 - 深橙色，悬停时稍微变亮
        style.map('Disable.TButton', 
                 background=[('active', '#FF8A65'), ('pressed', '#FF5722')], 
                 foreground=[('active', 'white'), ('pressed', 'white')])
        style.configure('Disable.TButton', font=('Arial', 11), background='#FF5722', foreground='white')
        
        # 启用网卡 - 绿色，悬停时稍微变亮
        style.map('Enable.TButton', 
                 background=[('active', '#66BB6A'), ('pressed', '#4CAF50')], 
                 foreground=[('active', 'white'), ('pressed', 'white')])
        style.configure('Enable.TButton', font=('Arial', 11), background='#4CAF50', foreground='white')
        
        # 重置网络 - 粉红色，悬停时稍微变亮
        style.map('Reset.TButton', 
                 background=[('active', '#F06292'), ('pressed', '#E91E63')], 
                 foreground=[('active', 'white'), ('pressed', 'white')])
        style.configure('Reset.TButton', font=('Arial', 11), background='#E91E63', foreground='white')

        # Ping按钮（统一macOS蓝）
        style.map('Ping.TButton',
                 background=[('active', '#0063E1'), ('pressed', '#0052BF')],
                 foreground=[('active', 'white'), ('pressed', 'white')])
        style.configure('Ping.TButton', font=('Arial', 11), background='#0A84FF', foreground='white')
        
        # 状态栏样式
        style.configure('Status.TLabel', font=('Segoe UI', 10), background='#E6F0FF', foreground='#0A84FF')
        
        # 初始化WMI
        self.wmi = None
        try:
            # 使用wmi.WMI()初始化WMI
            self.wmi = wmi.WMI()
        except Exception as e:
            messagebox.showerror("错误", f"WMI初始化失败: {str(e)}\n\n请确保以管理员身份运行程序。")
            # 继续创建UI，但禁用需要WMI的功能
        
        # OHM（可选，用于温度）
        self.ohm = None
        try:
            self._init_ohm_bridge()
        except Exception:
            self.ohm = None

        self._brand_images = {}

        self.setup_ui()
        
        # 只有在WMI初始化成功时才刷新网络适配器
        if self.wmi:
            self.refresh_network_adapters()
        else:
            self.status_var.set("WMI初始化失败，部分功能不可用")

    def run_ping(self, target: str) -> None:
        """在系统cmd中启动持续ping（-t）。"""
        if not target.strip():
            messagebox.showwarning("提示", "请输入要测试的目标地址")
            return
        try:
            # 在独立cmd窗口中持续ping，避免阻塞GUI
            # start 新开窗口；/k 执行后保持窗口
            cmd = f'start cmd /k ping {target} -t'
            subprocess.Popen(cmd, shell=True)
            self.status_var.set(f"正在测试网络: ping {target} -t（已在新窗口打开）")
        except Exception as e:
            messagebox.showerror("错误", f"启动ping失败: {e}")

    def run_tracert(self, target: str) -> None:
        """在系统cmd中启动 tracert 目标。"""
        if not target.strip():
            messagebox.showwarning("提示", "请输入要追踪的目标地址")
            return
        try:
            # start 新开窗口；/k 执行后保持窗口
            cmd = f'start cmd /k tracert {target}'
            subprocess.Popen(cmd, shell=True)
            self.status_var.set(f"正在进行网络追踪: tracert {target}（已在新窗口打开）")
        except Exception as e:
            messagebox.showerror("错误", f"启动网络追踪失败: {e}")

    def clear_browser_cache(self) -> None:
        """清理常见浏览器缓存（Edge/Chrome/360/IE），不删除账号密码等数据，带进度弹窗。"""
        def _remove_dir(path: str) -> bool:
            if not os.path.exists(path):
                return False
            try:
                import shutil
                shutil.rmtree(path, ignore_errors=True)
            except Exception:
                pass
            # 双保险（处理被占用残留）
            try:
                subprocess.run(f'cmd /c rmdir /s /q "{path}"', shell=True,
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception:
                pass
            return not os.path.exists(path)

        def _do_clear(update_progress=lambda v, t: None, close_dialog=lambda: None):
            try:
                self.status_var.set("正在清理浏览器缓存...")
                self.root.update()

                # 结束常见浏览器进程
                for exe in ("msedge.exe", "chrome.exe", "360chrome.exe", "360se.exe", "iexplore.exe"):
                    subprocess.run(f'taskkill /f /im {exe}', shell=True,
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

                userprofile = os.environ.get('USERPROFILE', '')

                # 需要清理的子目录（不会影响密码/登录状态的数据，如 Cookies/Login Data/Local Storage 均不删除）
                cache_subdirs = [
                    'Cache', 'Code Cache', os.path.join('Service Worker', 'CacheStorage'),
                    'GPUCache', 'Media Cache', 'ShaderCache'
                ]

                cleaned = 0
                # 统计总步数（Edge/Chrome/360 的 profiles × subdirs）
                edge_profiles = glob.glob(os.path.join(userprofile, 'AppData', 'Local', 'Microsoft', 'Edge', 'User Data', '*'))
                chrome_profiles = glob.glob(os.path.join(userprofile, 'AppData', 'Local', 'Google', 'Chrome', 'User Data', '*'))
                se360_profiles = glob.glob(os.path.join(userprofile, 'AppData', 'Local', '360Chrome', 'Chrome', 'User Data', '*'))
                total_steps = (
                    max(1, len(edge_profiles)) * len(cache_subdirs) +
                    max(1, len(chrome_profiles)) * len(cache_subdirs) +
                    max(1, len(se360_profiles)) * len(cache_subdirs) + 2  # + 结束进程 + IE
                )

                # Edge
                for prof in (edge_profiles or [os.path.join(userprofile, 'AppData', 'Local', 'Microsoft', 'Edge', 'User Data', 'Default')]):
                    for sub in cache_subdirs:
                        _remove_dir(os.path.join(prof, sub))
                        cleaned += 1
                        update_progress(cleaned, total_steps)

                # Chrome
                for prof in (chrome_profiles or [os.path.join(userprofile, 'AppData', 'Local', 'Google', 'Chrome', 'User Data', 'Default')]):
                    for sub in cache_subdirs:
                        _remove_dir(os.path.join(prof, sub))
                        cleaned += 1
                        update_progress(cleaned, total_steps)

                # 360极速
                for prof in (se360_profiles or [os.path.join(userprofile, 'AppData', 'Local', '360Chrome', 'Chrome', 'User Data', 'Default')]):
                    for sub in cache_subdirs:
                        _remove_dir(os.path.join(prof, sub))
                        cleaned += 1
                        update_progress(cleaned, total_steps)

                # IE/旧Edge 仅清理临时文件，不动Cookies/密码
                try:
                    subprocess.run('RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8', shell=True,
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                except Exception:
                    pass
                cleaned += 1
                update_progress(cleaned, total_steps)

                self.status_var.set("浏览器缓存清理完成")
                messagebox.showinfo("完成", "浏览器缓存清理完成（未清除登录状态与密码）。若浏览器已打开，请关闭并重新打开以生效。")
                close_dialog()
            except Exception as e:
                self.status_var.set(f"清理缓存出错: {e}")
                messagebox.showerror("错误", f"清理缓存时出错:\n{e}")
                close_dialog()

        # 后台执行，避免阻塞UI
        try:
            # 创建进度弹窗
            progress_win = tk.Toplevel(self.root)
            progress_win.title("正在清理缓存…")
            progress_win.resizable(False, False)
            ttk.Label(progress_win, text="正在清理浏览器缓存，请稍候…").grid(row=0, column=0, padx=12, pady=8)
            progress = ttk.Progressbar(progress_win, mode='determinate', length=260)
            progress.grid(row=1, column=0, padx=12, pady=(0, 12))

            def update_progress(cur, total):
                try:
                    progress['maximum'] = total
                    progress['value'] = cur
                    progress_win.update_idletasks()
                except Exception:
                    pass

            def close_dialog():
                try:
                    progress_win.destroy()
                except Exception:
                    pass

            t = threading.Thread(target=_do_clear, args=(update_progress, close_dialog), daemon=True)
            t.start()
        except Exception:
            _do_clear()

    def _run_cmd_silent(self, cmd: str) -> int:
        """静默执行命令，不读取输出，返回returncode。"""
        try:
            result = subprocess.run(
                cmd,
                shell=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
            )
            return int(result.returncode)
        except Exception:
            return 1

    def _run_cmd_text(self, cmd: str) -> tuple[int, str, str]:
        """执行命令并返回 (returncode, stdout, stderr)，使用gbk解码容错。"""
        try:
            result = subprocess.run(
                cmd,
                shell=True,
                capture_output=True,
                text=True,
                encoding='gbk',
                errors='ignore',
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
            )
            return int(result.returncode), result.stdout or "", result.stderr or ""
        except Exception as e:
            return 1, "", str(e)

    def _run_args_text(self, args: list) -> tuple[int, str, str]:
        """以参数列表方式执行命令，避免转义问题。返回 (returncode, stdout, stderr)。"""
        try:
            result = subprocess.run(
                args,
                shell=False,
                capture_output=True,
                text=True,
                encoding='gbk',
                errors='ignore',
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
            )
            return int(result.returncode), result.stdout or "", result.stderr or ""
        except Exception as e:
            return 1, "", str(e)

    def flush_dns_cache(self) -> None:
        """刷新DNS缓存。"""
        self.status_var.set("正在刷新DNS缓存…")
        self.root.update()
        rc = self._run_cmd_silent("ipconfig /flushdns")
        if rc == 0:
            self.status_var.set("已刷新DNS缓存")
            messagebox.showinfo("完成", "已刷新DNS缓存。")
        else:
            self.status_var.set("刷新DNS缓存失败")
            messagebox.showwarning("提示", "刷新DNS缓存失败，可能需要以管理员身份运行。")

    def ip_release(self) -> None:
        """释放当前适配器的IP地址。"""
        if not messagebox.askyesno("确认", "确定要释放当前IP地址吗？网络将短暂中断。"):
            return
        self.status_var.set("正在释放IP地址…")
        self.root.update()
        rc = self._run_cmd_silent("ipconfig /release")
        if rc == 0:
            self.status_var.set("已释放IP地址")
            messagebox.showinfo("完成", "已释放IP地址。")
        else:
            self.status_var.set("释放IP地址失败")
            messagebox.showwarning("提示", "释放IP地址失败，可能需要以管理员身份运行。")

    def ip_renew(self) -> None:
        """续租当前适配器的IP地址。"""
        self.status_var.set("正在续租IP地址…")
        self.root.update()
        rc = self._run_cmd_silent("ipconfig /renew")
        if rc == 0:
            self.status_var.set("已续租IP地址")
            messagebox.showinfo("完成", "已续租IP地址。")
        else:
            self.status_var.set("续租IP地址失败")
            messagebox.showwarning("提示", "续租IP地址失败，可能需要以管理员身份运行。")

    def winsock_reset_quick(self) -> None:
        """快速Winsock重置（可能需要重启）。"""
        if not messagebox.askyesno(
            "确认",
            "确定要重置Winsock吗？\n此操作可能需要重启后生效。",
        ):
            return
        self.status_var.set("正在执行 Winsock 重置…")
        self.root.update()
        rc = self._run_cmd_silent("netsh winsock reset")
        if rc == 0:
            self.status_var.set("Winsock 重置完成")
            if messagebox.askyesno("完成", "Winsock 重置完成，是否立即重启以生效？"):
                try:
                    self._run_cmd_silent("shutdown /r /t 0")
                except Exception:
                    messagebox.showwarning("提示", "无法自动重启，请手动重启计算机。")
        else:
            self.status_var.set("Winsock 重置失败")
            messagebox.showwarning("提示", "Winsock 重置失败，可能需要以管理员身份运行。")

    def enable_rdp_and_set_password(self) -> None:
        """一键开启远程桌面并可选设置当前本地账户密码。
        注意：修改本地账户密码具有风险，请确认后操作；域账户/微软账户可能不适用。
        """
        # 1) 开启远程桌面（注册表）
        try:
            with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as reg:
                key_path = r"SYSTEM\CurrentControlSet\Control\Terminal Server"
                with winreg.OpenKey(reg, key_path, 0, winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY) as k:
                    # 0 允许远程，1 禁止远程
                    winreg.SetValueEx(k, "fDenyTSConnections", 0, winreg.REG_DWORD, 0)
        except PermissionError:
            messagebox.showerror("错误", "需要管理员权限以开启远程桌面。")
            return
        except Exception as e:
            messagebox.showerror("错误", f"开启远程桌面失败：{e}")
            return

        # 2) 开启防火墙 远程桌面 规则（兼容中英文组名）
        self._run_cmd_silent('netsh advfirewall firewall set rule group="Remote Desktop" new enable=Yes')
        self._run_cmd_silent('netsh advfirewall firewall set rule group="远程桌面" new enable=Yes')

        # 3) 尝试启动远程桌面服务
        self._run_cmd_silent('sc start TermService')

        # 4) 可选设置本地用户密码
        password = (self.rdp_password.get() if hasattr(self, 'rdp_password') else "")
        if password:
            if not password.strip():
                # 密码为空：仅开启远程桌面，不改密码
                self.status_var.set("远程桌面已开启（未修改密码）。")
                messagebox.showinfo("提示", "密码为空：已仅开启远程桌面与防火墙放通，未修改本地账户密码。")
                return
            if len(password) < 4:
                messagebox.showwarning("提示", "远程密码长度至少为4位。")
                return
            if '"' in password:
                messagebox.showwarning("提示", '远程密码不得包含双引号(")字符。')
                return
            user_name = os.environ.get('USERNAME', '')
            if not user_name:
                messagebox.showwarning("提示", "无法获取当前用户名，已跳过密码设置。")
            else:
                if not messagebox.askyesno(
                    "确认",
                    f"将把本机账户 \"{user_name}\" 的登录密码改为你输入的内容。\n\n此操作具有风险，确定继续吗？"):
                    user_name = ""
                if user_name:
                    # 使用 net user 修改本地账户密码
                    # 注意：对于微软账户/域账户可能不适用
                    rc, out, err = self._run_args_text(['net', 'user', user_name, password])
                    if rc != 0:
                        messagebox.showwarning(
                            "提示",
                            "修改密码可能失败：\n"
                            "- 请确认使用的是本地账户（非微软/域账户）\n"
                            "- 密码需符合复杂性策略\n"
                            "- 请以管理员身份运行\n\n"
                            f"命令输出：\n{out or err}"
                        )
                    else:
                        messagebox.showinfo("完成", f"已尝试修改账户 \"{user_name}\" 的密码。")

        self.status_var.set("已开启远程桌面；若设置了密码，请使用该密码通过远程桌面连接。")
        messagebox.showinfo(
            "完成",
            "已开启远程桌面并放通防火墙。\n\n"
            "- 若密码为空：仅开启远程桌面，未修改本地账户密码\n"
            "- 若设置了新密码：请使用该密码进行远程登录（微软/域账户可能无效）"
        )

    def toggle_win11_autologon(self) -> None:
        """Win11 自动登录开关：
        - 设置 HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\PasswordLess\\Device\\DevicePasswordLessBuildVersion = 0 开启可见的 netplwiz 取消密码登录
        - 设置为 2 恢复默认(隐藏)
        显示当前状态并提供切换。
        """
        key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\PasswordLess\Device"
        value_name = "DevicePasswordLessBuildVersion"
        try:
            with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as reg:
                try:
                    with winreg.OpenKey(reg, key_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as k:
                        current, _ = winreg.QueryValueEx(k, value_name)
                except FileNotFoundError:
                    current = 2
        except Exception as e:
            messagebox.showerror("错误", f"读取注册表失败：{e}\n请以管理员身份运行。")
            return

        is_enabled = (int(current) == 0)
        action = "恢复默认(隐藏选项)" if is_enabled else "开启(显示选项)"

        if not messagebox.askyesno("确认切换", f"当前状态：{'已开启' if is_enabled else '未开启'}\n\n是否执行：{action}？"):
            return

        new_value = 2 if is_enabled else 0
        try:
            with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as reg:
                with winreg.CreateKeyEx(reg, key_path, 0, winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY) as k:
                    winreg.SetValueEx(k, value_name, 0, winreg.REG_DWORD, int(new_value))
            self.status_var.set("已更新自动登录选项")
            # 自动启动 netplwiz（新窗口，不阻塞）
            try:
                # 使用 cmd 的 start 启动，避免路径与位数问题
                self._run_cmd_silent('cmd /c start "" netplwiz')
            except Exception:
                pass
            messagebox.showinfo(
                "完成",
                "设置已应用，已为你打开 netplwiz。\n\n在弹出的窗口中取消勾选\"要使用本计算机，用户必须输入用户名和密码\"，并输入密码以启用开机自动登录。"
            )
        except PermissionError:
            messagebox.showerror("错误", "写入注册表失败：权限不足。请以管理员身份运行程序。")
        except Exception as e:
            messagebox.showerror("错误", f"写入注册表失败：{e}")

    def enable_firewall_ping(self) -> None:
        """允许 ICMPv4 ping：创建并启用允许规则，同时禁用阻止规则。"""
        try:
            self.status_var.set("正在配置防火墙允许 ping …")
            self.root.update()
            # 先禁用阻止规则
            self._run_cmd_silent('netsh advfirewall firewall set rule name="IPM_ICMPv4_Echo_Request_Block" new enable=no')
            self._run_cmd_silent('netsh advfirewall firewall set rule name="IPM_ICMPv4_Echo_Reply_Block" new enable=no')
            # 创建或启用允许规则
            self._run_cmd_silent('netsh advfirewall firewall set rule name="IPM_ICMPv4_Echo_Request_Allow" new enable=yes')
            self._run_cmd_silent('netsh advfirewall firewall set rule name="IPM_ICMPv4_Echo_Reply_Allow" new enable=yes')
            self._run_cmd_silent('netsh advfirewall firewall add rule name="IPM_ICMPv4_Echo_Request_Allow" dir=in action=allow protocol=icmpv4:8,any')
            self._run_cmd_silent('netsh advfirewall firewall add rule name="IPM_ICMPv4_Echo_Reply_Allow" dir=out action=allow protocol=icmpv4:0,any')

            if self._is_firewall_ping_enabled():
                self.status_var.set("已允许防火墙 ping（ICMPv4）")
                messagebox.showinfo("完成", "已配置 ICMPv4（ping）允许规则。若仍提示失败但能够ping通，这是正常的（状态判定为保守逻辑）。")
            else:
                self.status_var.set("配置ping规则失败，可能需要以管理员身份运行")
                messagebox.showwarning(
                    "提示",
                    "配置失败或被策略阻止。若实际可以 ping 通，请忽略此提示；也可检查是否有第三方安全策略覆盖。"
                )
        except Exception as e:
            self.status_var.set(f"配置失败: {e}")
            messagebox.showerror("错误", f"配置防火墙允许 ping 失败:\n{e}")

    def _is_firewall_ping_enabled(self) -> bool:
        """根据以下顺序判断是否允许 ping：
        1) 若任一自建 Block 规则启用 => 视为禁用
        2) 若任一自建 Allow 规则启用 => 视为允许
        3) 若防火墙总体关闭 => 视为允许
        4) 若存在任一启用的 ICMPv4 入站允许规则（任意名称/本地化）=> 视为允许
        5) 回退：通用名称检测
        """
        allow_req = 'IPM_ICMPv4_Echo_Request_Allow'
        allow_rep = 'IPM_ICMPv4_Echo_Reply_Allow'
        block_req = 'IPM_ICMPv4_Echo_Request_Block'
        block_rep = 'IPM_ICMPv4_Echo_Reply_Block'

        def _is_enabled(name: str) -> bool:
            rc, out, _ = self._run_cmd_text(f'netsh advfirewall firewall show rule name="{name}"')
            if rc != 0:
                return False
            t = out.lower()
            return ('enabled: yes' in t) or ('enable: yes' in t) or ('已启用' in out) or ('启用: 是' in out)

        # Block 优先
        if _is_enabled(block_req) or _is_enabled(block_rep):
            return False
        if _is_enabled(allow_req) or _is_enabled(allow_rep):
            return True

        # 3) 防火墙总体关闭 => 视为允许
        rc_fw, out_fw, _ = self._run_cmd_text('netsh advfirewall show allprofiles')
        if rc_fw == 0:
            txt = out_fw.lower()
            if ('state' in txt and 'off' in txt) or ('状态' in out_fw and '关闭' in out_fw):
                return True

        # 4) 任一启用的 ICMPv4 入站允许规则（名称/语言不确定，做内容匹配）
        rc_all, out_all, _ = self._run_cmd_text('netsh advfirewall firewall show rule name=all')
        if rc_all == 0:
            o_low = out_all.lower()
            if ('icmpv4' in o_low) and (('allow' in o_low) or ('允许' in out_all)) and (('dir' in o_low and 'in' in o_low) or ('方向' in out_all and '入站' in out_all)):
                return True

        # 回退到通用本地化名称（尽力）
        rc, out, _ = self._run_cmd_text('netsh advfirewall firewall show rule name="Allow ICMPv4 Echo Request"')
        if rc == 0:
            t = out.lower()
            if ('enabled: yes' in t) or ('enable: yes' in t) or ('已启用' in out) or ('启用: 是' in out):
                return True
        return False

    def disable_firewall_ping(self) -> None:
        """禁用 ICMPv4 ping：创建并启用阻止规则，同时禁用本程序创建的允许规则。"""
        try:
            self.status_var.set("正在禁用防火墙 ping …")
            self.root.update()
            # 禁用允许规则（本程序自建）
            self._run_cmd_silent('netsh advfirewall firewall set rule name="IPM_ICMPv4_Echo_Request_Allow" new enable=no')
            self._run_cmd_silent('netsh advfirewall firewall set rule name="IPM_ICMPv4_Echo_Reply_Allow" new enable=no')
            # 创建或启用阻止规则（阻止优先级高于允许）
            self._run_cmd_silent('netsh advfirewall firewall set rule name="IPM_ICMPv4_Echo_Request_Block" new enable=yes')
            self._run_cmd_silent('netsh advfirewall firewall set rule name="IPM_ICMPv4_Echo_Reply_Block" new enable=yes')
            self._run_cmd_silent('netsh advfirewall firewall add rule name="IPM_ICMPv4_Echo_Request_Block" dir=in action=block protocol=icmpv4:8,any')
            self._run_cmd_silent('netsh advfirewall firewall add rule name="IPM_ICMPv4_Echo_Reply_Block" dir=out action=block protocol=icmpv4:0,any')

            if not self._is_firewall_ping_enabled():
                self.status_var.set("已禁用防火墙 ping（ICMPv4）")
                messagebox.showinfo("完成", "已禁用 ICMPv4（ping）。如仍可ping通，可能是防火墙已关闭或存在更高优先级的允许策略。")
            else:
                self.status_var.set("禁用失败或被策略阻止")
                messagebox.showwarning("提示", "禁用失败，可能需要以管理员身份运行，或被安全策略禁止。若实际无法ping通，则可忽略。")
        except Exception as e:
            self.status_var.set(f"禁用失败: {e}")
            messagebox.showerror("错误", f"禁用防火墙 ping 失败:\n{e}")

    def toggle_firewall_ping(self) -> None:
        """开/关 防火墙 ping。根据当前状态提示并执行相反操作。"""
        enabled = self._is_firewall_ping_enabled()
        if enabled:
            if not messagebox.askyesno("确认", "检测到已允许 ping（ICMPv4）。是否现在禁用？"):
                return
            self.disable_firewall_ping()
        else:
            if not messagebox.askyesno("确认", "检测到当前未允许 ping（ICMPv4）。是否现在开启？"):
                return
            self.enable_firewall_ping()

    def open_network_control_panel(self) -> None:
        """一键打开Windows网络控制面板"""
        try:
            # 使用subprocess运行ncpa.cpl命令
            subprocess.Popen(['ncpa.cpl'], shell=True)
            self.status_var.set("已打开网络控制面板")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开网络控制面板：{str(e)}")
            self.status_var.set("打开网络控制面板失败")

    def open_devices_and_printers(self) -> None:
        """一键打开Windows设备和打印机"""
        try:
            # 使用explorer打开设备和打印机
            subprocess.Popen(['explorer', 'shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}'], shell=True)
            self.status_var.set("已打开设备和打印机")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开设备和打印机：{str(e)}")
            self.status_var.set("打开设备和打印机失败")
    
    def add_button_hover_effect(self, button):
        """为按钮添加鼠标悬停效果"""
        def on_enter(event):
            button.configure(cursor="hand2")
        
        def on_leave(event):
            button.configure(cursor="")
        
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)
        
    def setup_ui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="8")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # 标题（浅色简洁，文本置中）
        self.title_label = ttk.Label(main_frame, text='Windows IP地址管理器', font=self._get_scaled_font(18, 'bold'))
        self.title_label.grid(row=0, column=0, columnspan=3, pady=(0, 5), sticky=tk.N)
        self._register_scalable_widget(self.title_label, 'label', font=(18, 'bold'))
        
        # 版本号（小字体，灰色）
        version_text = f"版本: {get_version_string()}"
        self.version_label = ttk.Label(main_frame, text=version_text, font=self._get_scaled_font(10), foreground=SUBTLE)
        self.version_label.grid(row=1, column=0, columnspan=3, pady=(0, 10))
        self._register_scalable_widget(self.version_label, 'label', font=(10, 'normal'))

        # 顶部：网络测试（ping）行（移动到右侧容器中显示）
        
        # 右侧区域：先创建容器（后面把适配器区放进去）
        
        # 两列布局：左侧IP信息占一半，右侧动态缩放
        main_frame.columnconfigure(0, weight=1)  # 左侧占一半
        main_frame.columnconfigure(1, weight=1)  # 右侧占一半
        main_frame.rowconfigure(2, weight=1)
        
        # 左侧：当前IP信息（位于版本号下方，靠左整列）
        left_frame = ttk.LabelFrame(main_frame, text="当前IP信息", padding="8")
        left_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=8)
        
        # IP信息文本框
        self.ip_info_text = tk.Text(left_frame, height=16, width=45, state=tk.DISABLED, font=self._get_scaled_font(11, family='Consolas'), 
                                   bg=SURFACE_CARD, fg=PRIMARY_FG, insertbackground=PRIMARY_FG,
                                   selectbackground=ACCENT_HOVER, selectforeground='#0B1221', relief='flat')
        self.ip_info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self._register_scalable_widget(self.ip_info_text, 'text', width=45, height=16, font=(11, 'normal', 'Consolas'))
        
        # 滚动条
        ip_scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.ip_info_text.yview)
        ip_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.ip_info_text.configure(yscrollcommand=ip_scrollbar.set)
        
        # 配置左侧框架的网格权重
        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(0, weight=1)
        
        # 右侧：容器放适配器/网络测试/工具栏/内容
        right_frame = ttk.Frame(main_frame)
        right_frame.grid(row=2, column=1, sticky=(tk.N, tk.S, tk.W, tk.E), pady=8)
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(3, weight=1)
        right_frame.rowconfigure(4, weight=0)
        right_frame.rowconfigure(5, weight=0)

        # 适配器区：标题一行，选择框下一行
        adapter_section = ttk.Frame(right_frame)
        adapter_section.grid(row=0, column=0, sticky=(tk.E, tk.W), pady=(0,4))
        ttk.Label(adapter_section, text="网络适配器").grid(row=0, column=0, sticky=tk.W)
        # 在此处创建下拉框（避免未定义）
        self.adapter_var = tk.StringVar()
        self.adapter_combo = ttk.Combobox(
            adapter_section,
            textvariable=self.adapter_var,
            state="readonly",
            width=45,
        )
        self.adapter_combo.grid(row=1, column=0, sticky=(tk.E, tk.W), pady=(2, 0))
        self.adapter_combo.bind('<<ComboboxSelected>>', self.on_adapter_selected)
        self._register_scalable_widget(self.adapter_combo, 'combobox', width=45)
        # 右侧放置刷新按钮
        refresh_btn = ttk.Button(adapter_section, text="刷新", command=self.refresh_network_adapters, width=int(8 * self.scale_factor), style='Refresh.TButton')
        refresh_btn.grid(row=1, column=1, sticky=tk.E, padx=(6, 0))
        self.add_button_hover_effect(refresh_btn)
        self._register_scalable_widget(refresh_btn, 'button', width=8)

        # 自定义工具栏

        # 内网测试（ping 网关）行：放在网络测试上方
        lan_ping_frame = ttk.Frame(right_frame)
        # 放在Notebook（主IP配置区域）下方（将迁移到工具页）
        lan_ping_frame.grid_forget()
        self.lan_ping_target = tk.StringVar(value="")
        # 将在工具页中渲染

        # 清理浏览器缓存按钮（放在测试区下方，右对齐）
        cache_frame = ttk.Frame(right_frame)
        cache_frame.grid(row=6, column=0, sticky=tk.E, pady=(8, 0))
        clear_cache_btn = ttk.Button(
            cache_frame,
            text="一键清理浏览器缓存 - Edge/360极速/Chrome",
            command=self.clear_browser_cache
        )
        clear_cache_btn.grid(row=0, column=0)

        # Notebook 与自定义工具栏
        notebook = ttk.Notebook(right_frame)
        notebook.grid(row=3, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))
        
        # 设置notebook选项卡样式 - 确保文字居中
        try:
            style_notebook = ttk.Style()
            style_notebook.configure('TNotebook.Tab', 
                                   padding=(12, 6),
                                   font=('Segoe UI', 11),
                                   anchor='center',
                                   justify='center')
            style_notebook.map('TNotebook.Tab',
                             background=[('selected', '#E3F2FD'), ('active', '#F5F5F5')],
                             foreground=[('selected', '#1976D2'), ('active', '#424242')])
        except Exception:
            pass
        
        # 隐藏Notebook默认页签，避免与自定义工具栏重复
        try:
            style_hide = ttk.Style()
            style_hide.layout('TNotebook.Tab', [])
        except Exception:
            pass
        toolbar = ttk.Frame(right_frame)
        toolbar.grid(row=2, column=0, sticky=tk.E, pady=(0, 6))
        # 配置toolbar的列权重，确保按钮均匀分布
        for i in range(5):  # 5个按钮
            toolbar.columnconfigure(i, weight=1)
        # 科技风：分段按钮（Segmented）样式
        seg_style = ttk.Style()
        button_padding = (int(10 * self.scale_factor), int(6 * self.scale_factor))
        seg_style.configure('Segment.TButton', 
                           padding=button_padding, 
                           relief='flat',
                           background='#1F2937', 
                           foreground='#E5E7EB',
                           anchor='center',  # 设置文字居中对齐
                           justify='center')  # 设置文本居中对齐
        seg_style.map('Segment.TButton',
                      background=[('active', '#0EA5E9'), ('pressed', '#06B6D4')],
                      foreground=[('active', '#0B1221'), ('pressed', '#0B1221')])
        seg_style.configure('SegmentSelected.TButton', 
                           padding=button_padding, 
                           relief='flat',
                           background=ACCENT, 
                           foreground='#0B1221',
                           anchor='center',  # 设置文字居中对齐
                           justify='center')  # 设置文本居中对齐

        # 卡片与徽章样式（苹果风）
        ui_style = ttk.Style()
        ui_style.configure('Card.TFrame', background='#FFFFFF', relief='flat')
        ui_style.configure('CardTitle.TLabel', font=self._get_scaled_font(13, 'bold', 'Arial'), foreground='#0F172A', background='#FFFFFF')
        ui_style.configure('CardItemLeft.TLabel', font=self._get_scaled_font(10), foreground='#6B7280', background='#FFFFFF')
        ui_style.configure('CardItemRight.TLabel', font=self._get_scaled_font(10), foreground='#0F172A', background='#FFFFFF')
        ui_style.configure('CardValue.TLabel', font=self._get_scaled_font(11, family='Consolas'), foreground='#111827', background='#FFFFFF')
        ui_style.configure('Badge.TLabel', background='#E6F0FF', foreground='#0A84FF', padding=(6,2))
        ui_style.configure('Subtle.TLabel', foreground=SUBTLE, background='#FFFFFF')
        # 温度徽章样式
        ui_style.configure('TempGreen.TLabel', background='#DEF7EC', foreground='#03543F', padding=(6,2))
        ui_style.configure('TempYellow.TLabel', background='#FEF3C7', foreground='#92400E', padding=(6,2))
        ui_style.configure('TempOrange.TLabel', background='#FFEDD5', foreground='#92400E', padding=(6,2))
        ui_style.configure('TempRed.TLabel', background='#FEE2E2', foreground='#991B1B', padding=(6,2))
        # 风扇转速徽章样式
        ui_style.configure('FanGreen.TLabel', background='#E0F2FE', foreground='#0C4A6E', padding=(6,2))
        ui_style.configure('FanYellow.TLabel', background='#FEF9C3', foreground='#A16207', padding=(6,2))
        ui_style.configure('FanOrange.TLabel', background='#FFEDD5', foreground='#C2410C', padding=(6,2))
        ui_style.configure('FanRed.TLabel', background='#FEE2E2', foreground='#B91C1C', padding=(6,2))

        self._seg_btns = []
        def set_segment_selection(active_idx: int) -> None:
            for idx, btn in enumerate(self._seg_btns):
                btn.configure(style='SegmentSelected.TButton' if idx == active_idx else 'Segment.TButton')

        self._current_tab_index = 0
        def switch_tab(i:int):
            # 离开硬件页处理
            try:
                if hasattr(self, '_current_tab_index') and self._current_tab_index == 4:
                    self._on_hw_tab_leave()
            except Exception:
                pass
            notebook.select(i)
            set_segment_selection(i)
            self._current_tab_index = i
            # 进入硬件页处理
            try:
                if i == 4:
                    self._on_hw_tab_enter()
            except Exception:
                pass

        # 与实际添加的选项卡保持一致
        tb_texts = ["IP配置", "额外IP地址", "网卡控制", "工具", "硬件信息"]
        for i, t in enumerate(tb_texts):
            b = ttk.Button(toolbar, text=t, style='Segment.TButton', command=lambda idx=i: switch_tab(idx))
            b.grid(row=0, column=i, padx=(0 if i==0 else 2, 0), sticky=tk.EW)
            self._seg_btns.append(b)
        set_segment_selection(0)

        # Tab1: IP配置
        tab_ipcfg = ttk.Frame(notebook)
        notebook.add(tab_ipcfg, text="IP配置")
        main_ip_frame = ttk.LabelFrame(tab_ipcfg, text="主IP配置", padding="5")
        main_ip_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        main_ip_frame.columnconfigure(1, weight=1)

        ttk.Label(main_ip_frame, text="IP地址:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.ip_var = tk.StringVar()
        self.ip_entry = ttk.Entry(main_ip_frame, textvariable=self.ip_var, width=int(22 * self.scale_factor), validate='key')
        self.ip_entry['validatecommand'] = (self.ip_entry.register(self.validate_ipv4_entry), '%P')
        self.ip_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        self._register_scalable_widget(self.ip_entry, 'entry', width=22)

        ttk.Label(main_ip_frame, text="子网掩码:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.mask_var = tk.StringVar()
        self.mask_entry = ttk.Entry(main_ip_frame, textvariable=self.mask_var, width=int(22 * self.scale_factor), validate='key')
        self.mask_entry['validatecommand'] = (self.mask_entry.register(self.validate_ipv4_entry), '%P')
        self.mask_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        self._register_scalable_widget(self.mask_entry, 'entry', width=22)

        ttk.Label(main_ip_frame, text="默认网关:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.gateway_var = tk.StringVar()
        self.gateway_entry = ttk.Entry(main_ip_frame, textvariable=self.gateway_var, width=int(22 * self.scale_factor), validate='key')
        self.gateway_entry['validatecommand'] = (self.gateway_entry.register(self.validate_ipv4_entry), '%P')
        self.gateway_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        self._register_scalable_widget(self.gateway_entry, 'entry', width=22)

        ttk.Label(main_ip_frame, text="DNS服务器:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.dns_var = tk.StringVar()
        self.dns_entry = ttk.Entry(main_ip_frame, textvariable=self.dns_var, width=int(22 * self.scale_factor), validate='key')
        self.dns_entry['validatecommand'] = (self.dns_entry.register(self.validate_ipv4_entry), '%P')
        self.dns_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        self._register_scalable_widget(self.dns_entry, 'entry', width=22)

        # 构建公用操作按钮组（设置静态IP/设置DHCP/刷新/导出配置）
        def build_ip_actions(parent: tk.Widget) -> ttk.Frame:
            actions = ttk.Frame(parent)
            actions.columnconfigure(0, weight=1)
            actions.columnconfigure(1, weight=1)

            btn1 = ttk.Button(actions, text="设置静态IP", command=self.set_static_ip, width=int(12 * self.scale_factor), style='StaticIP.TButton')
            btn1.grid(row=0, column=0, padx=3, pady=3, sticky=tk.E)
            self.add_button_hover_effect(btn1)
            self._register_scalable_widget(btn1, 'button', width=12)

            btn2 = ttk.Button(actions, text="设置DHCP", command=self.set_dhcp, width=int(12 * self.scale_factor), style='DHCP.TButton')
            btn2.grid(row=0, column=1, padx=3, pady=3, sticky=tk.W)
            self.add_button_hover_effect(btn2)
            self._register_scalable_widget(btn2, 'button', width=12)

            btn3 = ttk.Button(actions, text="刷新", command=self.refresh_ip_info, width=int(12 * self.scale_factor), style='RefreshInfo.TButton')
            btn3.grid(row=1, column=0, padx=3, pady=3, sticky=tk.E)
            self.add_button_hover_effect(btn3)
            self._register_scalable_widget(btn3, 'button', width=12)

            btn4 = ttk.Button(actions, text="导出配置", command=self.export_config, width=int(12 * self.scale_factor), style='Export.TButton')
            btn4.grid(row=1, column=1, padx=3, pady=3, sticky=tk.W)
            self.add_button_hover_effect(btn4)
            self._register_scalable_widget(btn4, 'button', width=12)

            return actions

        # 在"IP配置"页底部加入操作区
        actions1 = build_ip_actions(tab_ipcfg)
        actions1.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # Tab2: 额外IP地址
        tab_extra = ttk.Frame(notebook)
        notebook.add(tab_extra, text="额外IP地址")
        multi_ip_frame = ttk.LabelFrame(tab_extra, text="额外IP地址", padding="5")
        multi_ip_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        multi_ip_frame.columnconfigure(1, weight=1)

        self.extra_ips = []
        self.extra_ip_frame = ttk.Frame(multi_ip_frame)
        self.extra_ip_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)

        btn_frame = ttk.Frame(multi_ip_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)

        add_ip_btn = ttk.Button(btn_frame, text="添加IP", command=self.add_extra_ip, width=int(10 * self.scale_factor), style='AddIP.TButton')
        add_ip_btn.grid(row=0, column=0, padx=2, pady=2, sticky=tk.E)
        self.add_button_hover_effect(add_ip_btn)
        self._register_scalable_widget(add_ip_btn, 'button', width=10)

        clear_ip_btn = ttk.Button(btn_frame, text="清空", command=self.clear_extra_ips, width=int(10 * self.scale_factor), style='Clear.TButton')
        clear_ip_btn.grid(row=0, column=1, padx=2, pady=2, sticky=tk.W)
        self.add_button_hover_effect(clear_ip_btn)
        self._register_scalable_widget(clear_ip_btn, 'button', width=10)

        # 在"额外IP地址"页底部加入同样的操作区
        actions2 = build_ip_actions(tab_extra)
        actions2.grid(row=2, column=0, sticky=(tk.W, tk.E))

        # Tab3: 网卡控制
        tab_adapter = ttk.Frame(notebook)
        notebook.add(tab_adapter, text="网卡控制")
        # Tab4: 工具（分组展示，更整洁）
        tab_tools = ttk.Frame(notebook)
        notebook.add(tab_tools, text="工具")
        tab_tools.columnconfigure(0, weight=1)
        tab_tools.rowconfigure(0, weight=1)

        # 为工具区域创建可滚动容器
        tools_canvas = tk.Canvas(tab_tools, highlightthickness=0, bg=PRIMARY_BG)
        tools_canvas.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))
        tools_scrollbar = ttk.Scrollbar(tab_tools, orient=tk.VERTICAL, command=tools_canvas.yview)
        tools_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        tools_canvas.configure(yscrollcommand=tools_scrollbar.set)
        
        # 绑定鼠标滚轮事件
        def _on_tools_mousewheel(event):
            try:
                delta = -1 * int(event.delta/120) if event.delta else (1 if event.num==5 else -1)
                tools_canvas.yview_scroll(delta, 'units')
            except Exception:
                pass
            return 'break'
        tools_canvas.bind('<MouseWheel>', _on_tools_mousewheel)
        tools_canvas.bind('<Button-4>', _on_tools_mousewheel)
        tools_canvas.bind('<Button-5>', _on_tools_mousewheel)
        tools_canvas.configure(xscrollincrement=0, yscrollincrement=20)

        # 创建工具内容框架
        tools_content_frame = ttk.Frame(tools_canvas, style='TFrame')
        tools_window = tools_canvas.create_window((0,0), window=tools_content_frame, anchor='nw')
        
        # 配置工具内容框架的列权重
        tools_content_frame.columnconfigure(0, weight=1)
        tools_content_frame.columnconfigure(1, weight=1)
        
        def _on_tools_frame_configure(event):
            tools_canvas.configure(scrollregion=tools_canvas.bbox("all"))
        tools_content_frame.bind('<Configure>', _on_tools_frame_configure)

        # Tab5: 硬件信息
        tab_hw = ttk.Frame(notebook)
        notebook.add(tab_hw, text="硬件信息")
        tab_hw.columnconfigure(0, weight=1)
        tab_hw.rowconfigure(0, weight=0)
        tab_hw.rowconfigure(1, weight=1)

        # 顶部控制条：实时刷新 + 刷新间隔 + 手动刷新
        ctrl_bar = ttk.Frame(tab_hw)
        ctrl_bar.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0,6))
        self.hw_live_enabled_var = tk.BooleanVar(value=False)
        self.hw_refresh_interval_var = tk.IntVar(value=2)
        self._hw_live_job_id = None
        ttk.Checkbutton(ctrl_bar, text="实时刷新", variable=self.hw_live_enabled_var,
                        command=lambda: self._toggle_hw_live()).pack(side=tk.LEFT)
        ttk.Label(ctrl_bar, text="间隔").pack(side=tk.LEFT, padx=(12,4))
        self.hw_interval_combo = ttk.Combobox(ctrl_bar, state='readonly', width=6,
                                              values=["1s", "2s", "5s", "10s"])
        self.hw_interval_combo.set("2s")
        def _on_interval_change(event=None):
            text = self.hw_interval_combo.get().replace('s','')
            try:
                self.hw_refresh_interval_var.set(int(text))
            except Exception:
                self.hw_refresh_interval_var.set(2)
        self.hw_interval_combo.bind('<<ComboboxSelected>>', _on_interval_change)
        self.hw_interval_combo.pack(side=tk.LEFT)
        self._register_scalable_widget(self.hw_interval_combo, 'combobox', width=6)
        copy_btn = ttk.Button(ctrl_bar, text="复制信息", command=self.copy_all_hardware_info)
        copy_btn.pack(side=tk.RIGHT, padx=(8,0))
        self._register_scalable_widget(copy_btn, 'button')
        
        refresh_btn = ttk.Button(ctrl_bar, text="立即刷新", command=self.refresh_hardware_info, style='RefreshInfo.TButton')
        refresh_btn.pack(side=tk.RIGHT)
        self._register_scalable_widget(refresh_btn, 'button')

        # 可滚动容器（Canvas + Frame）
        self.hw_canvas = tk.Canvas(tab_hw, highlightthickness=0, bg=PRIMARY_BG)
        self.hw_canvas.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.W, tk.E))
        hw_scrollbar = ttk.Scrollbar(tab_hw, orient=tk.VERTICAL, command=self.hw_canvas.yview)
        hw_scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.hw_canvas.configure(yscrollcommand=hw_scrollbar.set)
        # 绑定鼠标滚轮与拖动事件
        def _on_mousewheel(event):
            try:
                delta = -1 * int(event.delta/120) if event.delta else (1 if event.num==5 else -1)
                self.hw_canvas.yview_scroll(delta, 'units')
            except Exception:
                pass
            return 'break'
        self.hw_canvas.bind_all('<MouseWheel>', _on_mousewheel)
        self.hw_canvas.bind_all('<Button-4>', _on_mousewheel)
        self.hw_canvas.bind_all('<Button-5>', _on_mousewheel)
        self.hw_canvas.configure(xscrollincrement=0, yscrollincrement=20)

        self.hw_cards_frame = ttk.Frame(self.hw_canvas, style='TFrame')
        self.hw_window = self.hw_canvas.create_window((0,0), window=self.hw_cards_frame, anchor='nw')

        def _on_frame_configure(event):
            self._update_hw_scrollregion()
        self.hw_cards_frame.bind('<Configure>', _on_frame_configure)

        # 底部按钮条移除（已合并至顶部控制条）

        # 分组1：网络诊断
        diag_group = ttk.Labelframe(tools_content_frame, text="网络诊断", padding=8)
        diag_group.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=4, pady=(8, 4))
        diag_group.columnconfigure(1, weight=1)

        ttk.Label(diag_group, text="内网测试（网关）").grid(row=0, column=0, sticky=tk.E, padx=(0,8), pady=4)
        lan_entry = ttk.Entry(diag_group, textvariable=self.lan_ping_target, width=28)
        lan_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=4)
        self._register_scalable_widget(lan_entry, 'entry', width=28)
        lan_btn = ttk.Button(diag_group, text="内网测试 (ping)", style='Ping.TButton',
                             command=lambda: self.run_ping(self.lan_ping_target.get()))
        lan_btn.grid(row=0, column=2, sticky=tk.W, padx=(8,0), pady=4)
        self._register_scalable_widget(lan_btn, 'button')

        ttk.Label(diag_group, text="网络测试（外网）").grid(row=1, column=0, sticky=tk.E, padx=(0,8), pady=4)
        self.ping_target = tk.StringVar(value="www.baidu.com")
        ping_entry = ttk.Entry(diag_group, textvariable=self.ping_target, width=28)
        ping_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=4)
        self._register_scalable_widget(ping_entry, 'entry', width=28)
        ping_btn = ttk.Button(diag_group, text="网络测试 (ping)", style='Ping.TButton',
                              command=lambda: self.run_ping(self.ping_target.get()))
        ping_btn.grid(row=1, column=2, sticky=tk.W, padx=(8,0), pady=4)
        self._register_scalable_widget(ping_btn, 'button')

        # 网络追踪（tracert）
        ttk.Label(diag_group, text="网络追踪（tracert）").grid(row=2, column=0, sticky=tk.E, padx=(0,8), pady=4)
        self.tracert_target = tk.StringVar(value="www.baidu.com")
        tracert_entry = ttk.Entry(diag_group, textvariable=self.tracert_target, width=28)
        tracert_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=4)
        self._register_scalable_widget(tracert_entry, 'entry', width=28)
        tracert_btn = ttk.Button(diag_group, text="网络追踪 (tracert)", style='Ping.TButton',
                                 command=lambda: self.run_tracert(self.tracert_target.get()))
        tracert_btn.grid(row=2, column=2, sticky=tk.W, padx=(8,0), pady=4)
        self._register_scalable_widget(tracert_btn, 'button')

        # 分组2：系统工具
        sys_group = ttk.Labelframe(tools_content_frame, text="系统工具", padding=8)
        sys_group.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=4, pady=(4, 8))
        for i in range(3):
            sys_group.columnconfigure(i, weight=1)

        ttk.Button(sys_group, text="刷新DNS", command=self.flush_dns_cache).grid(row=0, column=0, sticky=tk.EW, padx=2, pady=4)
        ttk.Button(sys_group, text="IP释放", command=self.ip_release).grid(row=0, column=1, sticky=tk.EW, padx=2, pady=4)
        ttk.Button(sys_group, text="IP续租", command=self.ip_renew).grid(row=0, column=2, sticky=tk.EW, padx=2, pady=4)

        ttk.Button(sys_group, text="Winsock重置", style='Reset.TButton', command=self.winsock_reset_quick).grid(row=1, column=0, sticky=tk.EW, padx=2, pady=4)
        ttk.Button(sys_group, text="打开防火墙ping", command=self.enable_firewall_ping).grid(row=1, column=1, sticky=tk.EW, padx=2, pady=4)
        ttk.Button(sys_group, text="关闭防火墙ping", command=self.disable_firewall_ping).grid(row=1, column=2, sticky=tk.EW, padx=2, pady=4)

        # Win11 自动登录开关
        ttk.Button(sys_group, text="Win11自动登录开关", command=self.toggle_win11_autologon).grid(row=2, column=0, sticky=tk.EW, padx=2, pady=4)
        ttk.Button(sys_group, text="一键清理浏览器缓存 - Edge/360极速/Chrome", command=self.clear_browser_cache).grid(row=2, column=1, columnspan=2, sticky=tk.EW, padx=2, pady=4)

        # 分组3：远程桌面
        rdp_group = ttk.Labelframe(tools_content_frame, text="远程桌面", padding=8)
        rdp_group.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=4, pady=(0, 8))
        rdp_group.columnconfigure(1, weight=1)
        ttk.Label(rdp_group, text="远程密码").grid(row=0, column=0, sticky=tk.E, padx=(0,8))
        self.rdp_password = tk.StringVar(value="")
        rdp_entry = ttk.Entry(rdp_group, textvariable=self.rdp_password, width=28)
        rdp_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        self._register_scalable_widget(rdp_entry, 'entry', width=28)
        rdp_btn = ttk.Button(rdp_group, text="一键打开远程桌面", command=self.enable_rdp_and_set_password)
        rdp_btn.grid(row=0, column=2, sticky=tk.W, padx=(8,0))
        self._register_scalable_widget(rdp_btn, 'button')
        adapter_control_frame = ttk.LabelFrame(tab_adapter, text="网卡控制", padding="5")
        adapter_control_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        # 配置3列布局，每列权重相等
        adapter_control_frame.columnconfigure(0, weight=1)
        adapter_control_frame.columnconfigure(1, weight=1)
        adapter_control_frame.columnconfigure(2, weight=1)

        # 第一行：禁用网卡、启用网卡、重置网络
        self.disable_btn = ttk.Button(adapter_control_frame, text="禁用网卡", 
                                     command=self.disable_adapter, width=int(12 * self.scale_factor), style='Disable.TButton')
        self.disable_btn.grid(row=0, column=0, padx=3, pady=3, sticky=tk.EW)
        self.add_button_hover_effect(self.disable_btn)
        self._register_scalable_widget(self.disable_btn, 'button', width=12)

        self.enable_btn = ttk.Button(adapter_control_frame, text="启用网卡", 
                                    command=self.enable_adapter, width=int(12 * self.scale_factor), style='Enable.TButton')
        self.enable_btn.grid(row=0, column=1, padx=3, pady=3, sticky=tk.EW)
        self.add_button_hover_effect(self.enable_btn)
        self._register_scalable_widget(self.enable_btn, 'button', width=12)

        self.reset_network_btn = ttk.Button(adapter_control_frame, text="重置网络", 
                                           command=self.reset_network, width=int(12 * self.scale_factor), style='Reset.TButton')
        self.reset_network_btn.grid(row=0, column=2, padx=3, pady=3, sticky=tk.EW)
        self.add_button_hover_effect(self.reset_network_btn)
        self._register_scalable_widget(self.reset_network_btn, 'button', width=12)

        # 第二行：网络控制面板、设备和打印机
        self.open_ncpa_btn = ttk.Button(adapter_control_frame, text="网络控制面板", 
                                       command=self.open_network_control_panel, width=int(15 * self.scale_factor), style='Enable.TButton')
        self.open_ncpa_btn.grid(row=1, column=0, columnspan=2, padx=3, pady=3, sticky=tk.EW)
        self.add_button_hover_effect(self.open_ncpa_btn)
        self._register_scalable_widget(self.open_ncpa_btn, 'button', width=15)

        self.open_devices_btn = ttk.Button(adapter_control_frame, text="设备和打印机", 
                                          command=self.open_devices_and_printers, width=int(15 * self.scale_factor), style='Enable.TButton')
        self.open_devices_btn.grid(row=1, column=2, padx=3, pady=3, sticky=tk.EW)
        self.add_button_hover_effect(self.open_devices_btn)
        self._register_scalable_widget(self.open_devices_btn, 'button', width=15)
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W, style='Status.TLabel')
        status_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 适配器映射
        self.adapter_map = {}
        self.wmi_adapters = {}
        
        # 初次加载硬件信息
        try:
            self.refresh_hardware_info()
        except Exception:
            pass

    # ===== 硬件信息 =====
    def _bytes_to_gb(self, size_bytes: int) -> str:
        try:
            if size_bytes is None:
                return "未知"
            gb = float(size_bytes) / (1024 ** 3)
            # Windows磁盘显示通常以GiB近似为GB
            return f"{gb:.0f} GB"
        except Exception:
            return "未知"

    def _safe_str(self, v) -> str:
        try:
            if v is None:
                return "未知"
            s = str(v).strip()
            return s if s else "未知"
        except Exception:
            return "未知"

    def _get_display_info(self):
        """获取主显示分辨率、刷新率以及物理尺寸(英寸，若可用)。"""
        width = height = refresh = None
        diag_inch = None
        try:
            vcs = self.wmi.Win32_VideoController()
            if vcs:
                vc = vcs[0]
                width = getattr(vc, 'CurrentHorizontalResolution', None)
                height = getattr(vc, 'CurrentVerticalResolution', None)
                refresh = getattr(vc, 'CurrentRefreshRate', None)
        except Exception:
            pass
        # 读取EDID尺寸
        try:
            wmi_wmi = wmi.WMI(namespace='root\\wmi')
            params = wmi_wmi.WmiMonitorBasicDisplayParams()
            if params:
                p = params[0]
                w_cm = getattr(p, 'MaxHorizontalImageSize', None)
                h_cm = getattr(p, 'MaxVerticalImageSize', None)
                if w_cm and h_cm and w_cm > 0 and h_cm > 0:
                    diag_inch = round(math.sqrt(w_cm ** 2 + h_cm ** 2) / 2.54, 1)
        except Exception:
            pass
        return width, height, refresh, diag_inch

    def _decode_u16_array(self, arr) -> str:
        try:
            if not arr:
                return ""
            return ''.join(chr(x) for x in arr if isinstance(x, int) and x > 0)
        except Exception:
            return ""

    def _get_monitors_detailed_info(self):
        """返回监视器详细信息列表：[{name, manufacturer, serial, width_cm, height_cm, diag_inch}]"""
        mons = []
        try:
            w = wmi.WMI(namespace='root\\wmi')
            ids = {m.InstanceName: m for m in getattr(w, 'WmiMonitorID')()}
            params = {p.InstanceName: p for p in getattr(w, 'WmiMonitorBasicDisplayParams')()}
            for inst, mid in ids.items():
                name = self._decode_u16_array(getattr(mid, 'UserFriendlyName', None)) or 'Unknown'
                manu = self._decode_u16_array(getattr(mid, 'ManufacturerName', None)) or 'Unknown'
                serial = self._decode_u16_array(getattr(mid, 'SerialNumberID', None)) or 'Unknown'
                p = params.get(inst)
                wcm = getattr(p, 'MaxHorizontalImageSize', None) if p else None
                hcm = getattr(p, 'MaxVerticalImageSize', None) if p else None
                diag = None
                try:
                    if wcm and hcm and wcm > 0 and hcm > 0:
                        import math as _m
                        diag = round((_m.sqrt(wcm**2 + hcm**2))/2.54, 1)
                except Exception:
                    diag = None
                mons.append({
                    'name': name,
                    'manufacturer': manu,
                    'serial': serial,
                    'width_cm': wcm,
                    'height_cm': hcm,
                    'diag_inch': diag,
                })
        except Exception:
            pass
        return mons

    def _clear_hw_cards(self):
        try:
            for child in self.hw_cards_frame.winfo_children():
                child.destroy()
        except Exception:
            pass

    def _add_card(self, parent: ttk.Frame, title: str, rows: List[Tuple[str, str]],
                  title_icon: Optional[tk.PhotoImage] = None,
                  badges: Optional[List[Tuple[str, str]]] = None,
                  accent: Optional[str] = None):
        card = ttk.Frame(parent, style='Card.TFrame')
        card.pack(fill=tk.X, padx=6, pady=8)
        # 左侧色条（科技风）
        bar = tk.Frame(card, width=4, height=1, bg=accent or '#0EA5E9')
        bar.pack(side=tk.LEFT, fill=tk.Y)
        inner = ttk.Frame(card, style='Card.TFrame', padding=12)
        inner.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        title_row = ttk.Frame(inner, style='Card.TFrame')
        title_row.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E))
        # 渐变标题（使用 Canvas 绘制）
        gradient = tk.Canvas(title_row, height=22, highlightthickness=0, bg='#FFFFFF')
        gradient.pack(side=tk.LEFT, fill=tk.X, expand=True)
        # 生成渐变条
        try:
            g1 = accent or '#0EA5E9'
            g2 = '#60A5FA'
            width = 240
            steps = 60
            for i in range(steps):
                r = i / (steps-1)
                # 简单线性插值（hex -> rgb）
                def _hex_to_rgb(h):
                    h = h.lstrip('#')
                    return tuple(int(h[j:j+2], 16) for j in (0,2,4))
                def _rgb_to_hex(c):
                    return '#%02x%02x%02x' % c
                c1 = _hex_to_rgb(g1)
                c2 = _hex_to_rgb(g2)
                mix = tuple(int(c1[k]*(1-r)+c2[k]*r) for k in range(3))
                x0 = int(i*width/steps)
                x1 = int((i+1)*width/steps)
                gradient.create_rectangle(x0, 0, x1, 22, outline='', fill=_rgb_to_hex(mix))
        except Exception:
            pass
        # 标题文本覆盖在渐变上
        t_holder = ttk.Frame(title_row, style='Card.TFrame')
        t_holder.place(in_=gradient, relx=0.0, rely=0.0, relwidth=1.0, relheight=1.0)
        t_inner = ttk.Frame(t_holder, style='Card.TFrame')
        t_inner.pack(fill=tk.BOTH, expand=True)
        if title_icon is not None:
            ttk.Label(t_inner, image=title_icon, style='Card.TFrame').pack(side=tk.LEFT, padx=(8,6))
        ttk.Label(t_inner, text=title, style='CardTitle.TLabel').pack(side=tk.LEFT, padx=(0,8))
        if badges:
            for text, style_name in badges:
                ttk.Label(t_inner, text=text, style=style_name).pack(side=tk.LEFT, padx=(0, 8))
        # 将行渲染为可选择的 Text（方便鼠标选择复制）
        # 计算左列最大宽度便于对齐
        max_left = 0
        for l, _ in rows:
            max_left = max(max_left, len(str(l)))
        lines = []
        for l, r in rows:
            ltxt = str(l).ljust(max_left)
            lines.append(f"{ltxt}  {r}")
        text_widget = tk.Text(inner, height=min(12, len(lines)+1), bg=SURFACE_CARD, fg=PRIMARY_FG,
                               font=self._get_scaled_font(11, family='Consolas'), relief='flat', borderwidth=0, highlightthickness=0,
                               wrap='word', cursor='xterm', insertbackground=PRIMARY_FG)
        text_widget.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))
        text_widget.insert('1.0', "\n".join(lines))
        text_widget.configure(state=tk.DISABLED)
        # 注册文本控件用于缩放
        self._register_scalable_widget(text_widget, 'text', height=min(12, len(lines)+1), font=(11, 'normal', 'Consolas'))
        inner.columnconfigure(0, weight=1)

    def _update_hw_scrollregion(self):
        try:
            if hasattr(self, 'hw_canvas'):
                self.hw_canvas.update_idletasks()
                self.hw_canvas.configure(scrollregion=self.hw_canvas.bbox('all'))
                if hasattr(self, 'hw_window'):
                    self.hw_canvas.itemconfig(self.hw_window, width=self.hw_canvas.winfo_width())
                # 不强制跳到顶部，保持用户拖动位置；仅在首次渲染后无滚动时设为0
                if self.hw_canvas.yview() == (0.0, 1.0):
                    self.hw_canvas.yview_moveto(0)
        except Exception:
            pass

    # ---- 品牌图标加载（可选） ----
    def _get_brand_icon(self, kind: str, name: str) -> Optional[tk.PhotoImage]:
        try:
            key = (kind, name)
            if key in self._brand_images:
                return self._brand_images[key]
            brand = ""
            n = (name or "").lower()
            if kind == 'cpu':
                if 'intel' in n:
                    brand = 'intel'
                elif 'amd' in n or 'ryzen' in n:
                    brand = 'amd'
            elif kind == 'gpu':
                if 'nvidia' in n or 'geforce' in n or 'rtx' in n or 'gtx' in n:
                    brand = 'nvidia'
                elif 'amd' in n or 'radeon' in n:
                    brand = 'amd'
                elif 'intel' in n or 'uhd' in n or 'arc' in n:
                    brand = 'intel'
            elif kind == 'board':
                if 'asus' in n:
                    brand = 'asus'
                elif 'gigabyte' in n or 'aorus' in n:
                    brand = 'gigabyte'
                elif 'msi' in n:
                    brand = 'msi'
                elif 'asrock' in n:
                    brand = 'asrock'
            if not brand:
                return None
            # 支持的查找路径
            candidates = [
                resource_path(f"Resources/brands/{kind}_{brand}.png"),
                resource_path(f"Resources/brands/{brand}.png"),
                os.path.join('Resources', 'brands', f"{kind}_{brand}.png"),
                os.path.join('Resources', 'brands', f"{brand}.png"),
            ]
            for p in candidates:
                if os.path.exists(p):
                    img = tk.PhotoImage(file=p)
                    self._brand_images[key] = img
                    return img
        except Exception:
            return None
        return None

    # ---- 温度读取（LibreHardwareMonitor，可选） ----
    def _init_lhm_bridge(self):
        try:
            import clr  # type: ignore
        except Exception:
            self.lhm = None
            return
        # 搜索 LibreHardwareMonitor DLL
        dll_candidates = [
            resource_path('LibreHardwareMonitor/LibreHardwareMonitorLib.dll'),
            resource_path('LibreHardwareMonitorLib.dll'),
            os.path.join('LibreHardwareMonitor', 'LibreHardwareMonitorLib.dll'),
        ]
        dll_path = next((p for p in dll_candidates if os.path.exists(p)), None)
        if not dll_path:
            self.lhm = None
            return
        try:
            import clr  # type: ignore
            sys.path.append(os.path.dirname(dll_path))
            clr.AddReference(os.path.basename(dll_path).replace('.dll',''))
            from LibreHardwareMonitor.Hardware import Computer, SensorType  # type: ignore
            comp = Computer()
            comp.IsMotherboardEnabled = True
            comp.IsCpuEnabled = True
            comp.IsMemoryEnabled = True
            comp.IsGpuEnabled = True
            comp.IsStorageEnabled = True
            comp.Open()
            self.lhm = {'lib': Computer, 'computer': comp}
        except Exception as e:
            self.lhm = None

    def _read_temperatures(self) -> dict:
        temps: dict = {}
        try:
            if not self.lhm:
                # 仅回退到 WMI(LibreHardwareMonitor)（不主动启动 EXE）
                return self._read_temperatures_via_wmi_lhm()
            Hardware = self.lhm['lib']
            comp = self.lhm['computer']
            def update_hw(hw):
                try:
                    hw.Update()
                    # 子硬件
                    for sh in hw.SubHardware:
                        update_hw(sh)
                    for s in hw.Sensors:
                        if s.SensorType == Hardware.SensorType.Temperature:
                            name = f"{hw.HardwareType}_{hw.Name}_{s.Name}"
                            temps[name] = float(s.Value) if s.Value is not None else None
                except Exception:
                    pass
            for hw in comp.Hardware:
                update_hw(hw)
        except Exception:
            pass
        return temps

    def _read_lhm_metrics(self) -> dict:
        """读取 LibreHardwareMonitor 提供的关键指标：温度与风扇。
        返回 {'temps': {...}, 'fans': {...}}
        """
        temps: dict = {}
        fans: dict = {}
        try:
            # 首选 DLL 方式
            if self.lhm:
                Computer = self.lhm['lib']
                comp = self.lhm['computer']
                
                # 导入SensorType
                from LibreHardwareMonitor.Hardware import SensorType
                
                # 简单的硬件更新方法
                def update_hardware(hardware):
                    try:
                        hardware.Update()
                        for subhardware in hardware.SubHardware:
                            update_hardware(subhardware)
                    except Exception as e:
                        print(f"更新硬件失败: {e}")
                
                # 更新所有硬件数据
                for hardware in comp.Hardware:
                    update_hardware(hardware)
                
                # 遍历所有硬件和传感器
                def collect_sensors(hardware):
                    try:
                        for sensor in hardware.Sensors:
                            if sensor.SensorType == SensorType.Temperature:
                                temps[f"{hardware.HardwareType}_{hardware.Name}_{sensor.Name}"] = float(sensor.Value) if sensor.Value is not None else None
                            elif sensor.SensorType == SensorType.Fan:
                                fans[f"{hardware.HardwareType}_{hardware.Name}_{sensor.Name}"] = float(sensor.Value) if sensor.Value is not None else None
                        
                        # 递归处理子硬件
                        for subhardware in hardware.SubHardware:
                            collect_sensors(subhardware)
                    except Exception as e:
                        print(f"收集传感器数据失败: {e}")
                
                for hardware in comp.Hardware:
                    collect_sensors(hardware)
                
                return {'temps': temps, 'fans': fans}
            
            # 回退到 WMI 提供器
            lhm_ns = wmi.WMI(namespace='root\\LibreHardwareMonitorLib')
            for sensor in lhm_ns.Sensor():
                try:
                    st = getattr(sensor, 'SensorType', '').lower()
                    name = f"{getattr(sensor, 'HardwareType', '')}_{getattr(sensor, 'Hardware', '')}_{getattr(sensor, 'Name', '')}"
                    val = getattr(sensor, 'Value', None)
                    if val is None:
                        continue
                    if st == 'temperature':
                        temps[name] = float(val)
                    elif st == 'fan':
                        fans[name] = float(val)
                except Exception:
                    pass
        except Exception as e:
            print(f"读取LibreHardwareMonitor指标失败: {e}")
        return {'temps': temps, 'fans': fans}

    def _read_temperatures_via_wmi_lhm(self) -> dict:
        """通过 LibreHardwareMonitor 的 WMI 提供器读取温度（无需 pythonnet）。"""
        result = {}
        try:
            lhm_ns = wmi.WMI(namespace='root\\LibreHardwareMonitorLib')
            for sensor in lhm_ns.Sensor():
                try:
                    if getattr(sensor, 'SensorType', '').lower() == 'temperature':
                        name = f"{getattr(sensor, 'HardwareType', '')}_{getattr(sensor, 'Hardware', '')}_{getattr(sensor, 'Name', '')}"
                        val = getattr(sensor, 'Value', None)
                        if val is not None:
                            result[name] = float(val)
                except Exception:
                    pass
        except Exception:
            return {}
        return result

    def _start_lhm_background(self) -> None:
        """尝试无窗口启动 LibreHardwareMonitor.exe（若存在）。"""
        try:
            candidates = [
                resource_path('LibreHardwareMonitor/LibreHardwareMonitor.exe'),
                os.path.join('LibreHardwareMonitor', 'LibreHardwareMonitor.exe'),
            ]
            exe = next((p for p in candidates if os.path.exists(p)), None)
            if not exe:
                return
            # 已有进程则略过
            try:
                # Windows: tasklist 过滤
                rc = subprocess.run('tasklist | findstr /I "LibreHardwareMonitor.exe"', shell=True,
                                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                if rc.returncode == 0:
                    return
            except Exception:
                pass
            subprocess.Popen(f'"{exe}" /minimized', shell=True,
                             creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        except Exception:
            pass

    def _pick_temp(self, temps: dict, prefer_keys: List[str]) -> Optional[float]:
        for k in prefer_keys:
            # 模糊匹配
            for name, v in temps.items():
                if k.lower() in name.lower() and v is not None:
                    return float(v)
        return None

    def _pick_fan(self, fans: dict, prefer_keys: List[str]) -> Optional[float]:
        for k in prefer_keys:
            for name, v in fans.items():
                if k.lower() in name.lower() and v is not None:
                    return float(v)
        return None

    def _temp_badge(self, value: Optional[float]) -> Optional[Tuple[str, str]]:
        if value is None:
            return None
        t = float(value)
        if t < 60:
            style = 'TempGreen.TLabel'
        elif t < 75:
            style = 'TempYellow.TLabel'
        elif t < 85:
            style = 'TempOrange.TLabel'
        else:
            style = 'TempRed.TLabel'
        return (f"{t:.0f}℃", style)

    def _fan_badge(self, value: Optional[float]) -> Optional[Tuple[str, str]]:
        if value is None:
            return None
        f = float(value)
        if f < 800:
            style = 'FanGreen.TLabel'
        elif f < 1500:
            style = 'FanYellow.TLabel'
        elif f < 2500:
            style = 'FanOrange.TLabel'
        else:
            style = 'FanRed.TLabel'
        return (f"{f:.0f} RPM", style)

    def refresh_hardware_info(self) -> None:
        """刷新并展示硬件信息（卡片式UI）。"""
        try:
            if not self.wmi:
                self.wmi = wmi.WMI()
        except Exception:
            pass
        # 清空旧卡片
        self._clear_hw_cards()

        # 采集温度/风扇（若可用）
        temps = {}
        fans = {}
        try:
            metrics = self._read_lhm_metrics()
            temps = metrics.get('temps', {})
            fans = metrics.get('fans', {})
        except Exception:
            temps, fans = {}, {}

        # CPU
        try:
            cpus = self.wmi.Win32_Processor()
            if cpus:
                cpu = cpus[0]
                name = self._safe_str(cpu.Name)
                cores = getattr(cpu, 'NumberOfCores', None)
                threads = getattr(cpu, 'NumberOfLogicalProcessors', None)
                mhz = getattr(cpu, 'MaxClockSpeed', None)
                ghz_str = f"{(mhz or 0)/1000:.1f}GHz" if mhz else "未知频率"
                rows = [("型号", name), ("频率", ghz_str), ("核心/线程", f"{cores or '未知'} / {threads or '未知'}")]
                cpu_temp = self._pick_temp(temps, ["CPU Package", "CPU Core", "CPU", "Tdie", "CCD"])
                if cpu_temp is not None:
                    rows.append(("温度", f"{cpu_temp:.0f}℃"))
                cpu_fan = self._pick_fan(fans, ["CPU", "Processor", "PUMP", "Fan #1", "Fan #2", "Fan #3", "Fan #4", "Fan #5", "Fan #6"])  # 可能读到水泵转速
                if cpu_fan is not None:
                    rows.append(("风扇", f"{cpu_fan:.0f} RPM"))
                temp_badge = self._temp_badge(cpu_temp)
                fan_badge = self._fan_badge(cpu_fan)
                badges = []
                if temp_badge:
                    badges.append(temp_badge)
                if fan_badge:
                    badges.append(fan_badge)
                self._add_card(self.hw_cards_frame, "处理器", rows,
                               title_icon=self._get_brand_icon('cpu', name),
                               badges=badges if badges else None,
                               accent='#3B82F6')
        except Exception:
            pass

        # GPU
        try:
            gpus = self.wmi.Win32_VideoController()
            if gpus:
                rows = []
                for idx, gpu in enumerate(gpus, 1):
                    gname = self._safe_str(gpu.Name)
                    vram = getattr(gpu, 'AdapterRAM', None)
                    vram_str = self._bytes_to_gb(vram) if vram else ""
                    label = f"显卡{idx}" if len(gpus) > 1 else "显卡"
                    gtemp = self._pick_temp(temps, ["GPU Hot Spot", "GPU Core", "GPU", gname])
                    info = f"{gname}{('（显存 '+vram_str+'）') if vram_str else ''}"
                    if gtemp is not None:
                        info += f"  温度 {gtemp:.0f}℃"
                    gfan = self._pick_fan(fans, ["GPU", gname])
                    if gfan is not None:
                        info += f"  风扇 {gfan:.0f} RPM"
                    rows.append((label, info))
                # 显卡温度和风扇徽章（展示第一张卡的数据为主）
                gpu_temp = self._pick_temp(temps, ["GPU Core", "GPU", "GPU Hot Spot"])
                gpu_fan = self._pick_fan(fans, ["GPU", "GPU Fan", "GPU Core Fan"])
                temp_badge = self._temp_badge(gpu_temp)
                fan_badge = self._fan_badge(gpu_fan)
                badges = []
                if temp_badge:
                    badges.append(temp_badge)
                if fan_badge:
                    badges.append(fan_badge)
                icon = self._get_brand_icon('gpu', gpus[0].Name if gpus else '')
                self._add_card(self.hw_cards_frame, "显卡", rows,
                               title_icon=icon,
                               badges=badges if badges else None,
                               accent='#10B981')
        except Exception:
            pass

        # 内存
        try:
            mems = self.wmi.Win32_PhysicalMemory()
            if mems:
                total = 0
                speeds = set()
                rows = []
                for m in mems:
                    cap = int(getattr(m, 'Capacity', 0) or 0)
                    total += cap
                    speed = getattr(m, 'ConfiguredClockSpeed', None) or getattr(m, 'Speed', None)
                    if speed:
                        speeds.add(int(speed))
                    man = self._safe_str(getattr(m, 'Manufacturer', None))
                    pn = self._safe_str(getattr(m, 'PartNumber', None))
                    rows.append((f"{man} {pn}", f"{self._bytes_to_gb(cap)}  @ {speed or '未知'} MHz"))
                total_str = self._bytes_to_gb(total)
                count = len(mems)
                speed_str = (f"{max(speeds)}" if len(speeds)==1 else ("/".join(str(s) for s in sorted(speeds)) if speeds else "未知"))
                rows.insert(0, ("总容量/条数", f"{total_str} / {count}"))
                rows.insert(1, ("频率", f"{speed_str} MHz"))
                self._add_card(self.hw_cards_frame, "内存", rows, accent='#06B6D4')
        except Exception:
            pass

        # 硬盘
        try:
            disks = self.wmi.Win32_DiskDrive()
            if disks:
                rows = []
                for d in disks:
                    model = self._safe_str(getattr(d, 'Model', None))
                    size = getattr(d, 'Size', None)
                    size_str = self._bytes_to_gb(int(size)) if size else "未知"
                    is_ssd = None
                    try:
                        media_type = getattr(d, 'MediaType', None)
                        if media_type and isinstance(media_type, str):
                            if 'ssd' in media_type.lower():
                                is_ssd = True
                            elif 'hard' in media_type.lower():
                                is_ssd = False
                    except Exception:
                        pass
                    if is_ssd is None:
                        is_ssd = 'SSD' in model.upper()
                    typ = 'SSD' if is_ssd else 'HDD'
                    dtemp = self._pick_temp(temps, [model, 'HDD', 'SSD', 'Drive'])
                    tail = f"容量 {size_str}  类型 {typ}"
                    if dtemp is not None:
                        tail += f"  温度 {dtemp:.0f}℃"
                    dfan = self._pick_fan(fans, [model, 'HDD', 'Drive'])
                    if dfan is not None:
                        tail += f"  风扇 {dfan:.0f} RPM"
                    rows.append((model, tail))
                self._add_card(self.hw_cards_frame, "硬盘", rows, accent='#F59E0B')
        except Exception:
            pass

        # 主板
        try:
            bbs = self.wmi.Win32_BaseBoard()
            if bbs:
                bb = bbs[0]
                manu = self._safe_str(getattr(bb, 'Manufacturer', None))
                prod = self._safe_str(getattr(bb, 'Product', None))
                board_badge = self._temp_badge(self._pick_temp(temps, ["Mainboard", "Motherboard"]))
                self._add_card(self.hw_cards_frame, "主板", [("制造商", manu), ("型号", prod)],
                               title_icon=self._get_brand_icon('board', manu + ' ' + prod),
                               badges=[board_badge] if board_badge else None,
                               accent='#6366F1')
        except Exception:
            pass

        # 显示器（放在主板后）
        try:
            width, height, refresh, diag_in = self._get_display_info()
            rows = []
            if width and height:
                rows.append(("分辨率", f"{width} x {height}"))
                rows.append(("刷新率", f"{refresh} Hz" if refresh else "未知"))
                rows.append(("屏幕尺寸", f"{diag_in} 英寸" if diag_in else "未知"))
            # 多显示器补充（名称/厂商/序列号/物理尺寸）
            mons = self._get_monitors_detailed_info()
            for idx, m in enumerate(mons, 1):
                rows.append((f"显示器{idx}", m['name']))
                rows.append(("厂商", m['manufacturer']))
                if m.get('serial') and m['serial'] != 'Unknown':
                    rows.append(("序列号", m['serial']))
                if m.get('width_cm') and m.get('height_cm'):
                    rows.append(("物理尺寸", f"{m['width_cm']} x {m['height_cm']} cm"))
                if m.get('diag_inch'):
                    rows.append(("对角线", f"{m['diag_inch']} 英寸"))
            if rows:
                self._add_card(self.hw_cards_frame, "显示器", rows, accent='#A855F7')
        except Exception:
            pass

        # 系统信息
        try:
            os_name = platform.system()
            release = platform.release()
            version = platform.version()
            build = platform.win32_ver()[1] if hasattr(platform, 'win32_ver') else ''
            # 安装时间（WMI）
            install_time = '未知'
            try:
                os_wmi = self.wmi.Win32_OperatingSystem()[0]
                # Convert WMI datetime format yyyymmddHHMMSS.mmmmmmsUUU
                raw = getattr(os_wmi, 'InstallDate', None)
                if raw:
                    from datetime import datetime
                    install_time = datetime.strptime(raw.split('.')[0], '%Y%m%d%H%M%S').strftime('%Y-%m-%d %H:%M:%S')
            except Exception:
                pass
            rows = [
                ("系统版本", f"{os_name} {release} (Build {build})".strip()),
                ("版本号", version),
                ("安装时间", install_time),
            ]
            self._add_card(self.hw_cards_frame, "系统信息", rows, accent='#0EA5E9')
        except Exception:
            pass

        self.status_var.set("硬件信息已刷新")
        self._update_hw_scrollregion()

    # ---- 硬件页实时刷新控制 ----
    def _toggle_hw_live(self):
        if self.hw_live_enabled_var.get():
            self._schedule_hw_tick()
        else:
            self._cancel_hw_tick()

    def _schedule_hw_tick(self):
        self._cancel_hw_tick()
        try:
            interval_ms = max(1, int(self.hw_refresh_interval_var.get())) * 1000
        except Exception:
            interval_ms = 2000
        self._hw_live_job_id = self.root.after(interval_ms, self._hw_tick)

    def _cancel_hw_tick(self):
        try:
            if self._hw_live_job_id:
                self.root.after_cancel(self._hw_live_job_id)
        except Exception:
            pass
        self._hw_live_job_id = None

    def _hw_tick(self):
        try:
            self.refresh_hardware_info()
        finally:
            if self.hw_live_enabled_var.get():
                self._schedule_hw_tick()

    def copy_all_hardware_info(self) -> None:
        try:
            # 遍历硬件卡片中的 Text 控件，拼接文本
            texts = []
            for child in self.hw_cards_frame.winfo_children():
                try:
                    inner = child.winfo_children()[1]  # [bar, inner]
                    # 找到标题区域后的 Text
                    for w in inner.winfo_children():
                        if isinstance(w, tk.Text):
                            w.configure(state=tk.NORMAL)
                            texts.append(w.get('1.0', tk.END).rstrip())
                            w.configure(state=tk.DISABLED)
                            break
                except Exception:
                    pass
            clip = "\n\n".join(texts)
            self.root.clipboard_clear()
            self.root.clipboard_append(clip)
            self.status_var.set("已复制全部硬件信息到剪贴板")
        except Exception as e:
            self.status_var.set(f"复制失败: {e}")

    def _on_hw_tab_enter(self):
        if self.hw_live_enabled_var.get():
            self._schedule_hw_tick()

    def _on_hw_tab_leave(self):
        self._cancel_hw_tick()
        
    def add_extra_ip(self):
        """添加额外IP地址"""
        if len(self.extra_ips) >= 5:  # 限制最多5个额外IP
            messagebox.showwarning("警告", "最多只能添加5个额外IP地址")
            return
        
        # 添加标签（只在第一次添加时创建）
        if len(self.extra_ips) == 0:
            if not hasattr(self, 'extra_ip_label_frame') or not self.extra_ip_label_frame.winfo_exists():
                self.extra_ip_label_frame = ttk.Frame(self.extra_ip_frame)
                self.extra_ip_label_frame.pack(fill=tk.X, pady=2)
                ttk.Label(self.extra_ip_label_frame, text="IP地址").pack(side=tk.LEFT, padx=(0, 5))
                ttk.Label(self.extra_ip_label_frame, text="子网掩码").pack(side=tk.LEFT, padx=(0, 5))
                ttk.Label(self.extra_ip_label_frame, text="操作").pack(side=tk.LEFT)
            
        ip_frame = ttk.Frame(self.extra_ip_frame)
        ip_frame.pack(fill=tk.X, pady=2)
        
        # IP地址输入
        ip_var = tk.StringVar()
        ip_entry = ttk.Entry(ip_frame, textvariable=ip_var, width=15, validate='key')
        ip_entry['validatecommand'] = (ip_entry.register(self.validate_ipv4_entry), '%P')
        ip_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        # 子网掩码输入
        mask_var = tk.StringVar()
        mask_entry = ttk.Entry(ip_frame, textvariable=mask_var, width=15, validate='key')
        mask_entry['validatecommand'] = (mask_entry.register(self.validate_ipv4_entry), '%P')
        mask_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        # 删除按钮
        del_btn = ttk.Button(ip_frame, text="删除", 
                           command=lambda: self.remove_extra_ip(ip_frame, ip_var, mask_var), style='Clear.TButton')
        del_btn.pack(side=tk.LEFT)
        self.add_button_hover_effect(del_btn)
        
        # 保存引用
        self.extra_ips.append({
            'frame': ip_frame,
            'ip_var': ip_var,
            'mask_var': mask_var
        })
        
    def remove_extra_ip(self, frame, ip_var, mask_var):
        """删除额外IP地址"""
        frame.destroy()
        self.extra_ips = [ip for ip in self.extra_ips if ip['ip_var'] != ip_var]
        
        # 如果删除后没有额外IP了，也删除标签框架
        if len(self.extra_ips) == 0:
            if hasattr(self, 'extra_ip_label_frame') and self.extra_ip_label_frame.winfo_exists():
                self.extra_ip_label_frame.destroy()
        
    def clear_extra_ips(self):
        """清空所有额外IP地址"""
        for ip_info in self.extra_ips:
            ip_info['frame'].destroy()
        self.extra_ips.clear()
        
        # 删除标签框架
        if hasattr(self, 'extra_ip_label_frame') and self.extra_ip_label_frame.winfo_exists():
            self.extra_ip_label_frame.destroy()
        
    def get_extra_ips(self):
        """获取所有额外IP地址"""
        extra_ips = []
        for ip_info in self.extra_ips:
            ip = ip_info['ip_var'].get().strip()
            mask = ip_info['mask_var'].get().strip()
            # 使用新的验证函数
            if ip and mask:
                is_valid_ip, _ = self.validate_ip_address(ip, "额外IP地址")
                is_valid_mask, _ = self.validate_ip_address(mask, "额外子网掩码")
                if is_valid_ip and is_valid_mask:
                    extra_ips.append((ip, mask))
        return extra_ips
        
    def refresh_network_adapters(self):
        """使用WMI获取网络适配器列表（包括已禁用的）"""
        if not self.wmi:
            messagebox.showerror("错误", "WMI未初始化，无法获取网络适配器信息")
            return
            
        try:
            self.status_var.set("正在获取网络适配器...")
            # 使用after方法延迟更新，避免阻塞UI
            self.root.after(10, self._do_refresh_network_adapters)
            
        except Exception as e:
            self.status_var.set(f"错误: {str(e)}")
            messagebox.showerror("错误", f"获取网络适配器时出错:\n{str(e)}")
    
    def _do_refresh_network_adapters(self):
        """实际执行网络适配器刷新（在后台线程中）"""
        try:
            adapters = []
            self.adapter_map = {}
            self.wmi_adapters = {}
            self.network_adapters = {}
            
            # 重新初始化WMI连接，确保获取最新状态
            self.wmi = wmi.WMI()
            
            # 获取所有网络适配器（包括已禁用的）
            # 使用PhysicalAdapter=True筛选物理网卡，不再使用IPEnabled=True筛选
            network_adapters = self.wmi.Win32_NetworkAdapter(PhysicalAdapter=True)
            
            # 获取所有网卡配置信息
            nic_configs = self.wmi.Win32_NetworkAdapterConfiguration()
            
            # 创建配置映射
            config_map = {}
            for config in nic_configs:
                if config.SettingID:
                    config_map[config.SettingID] = config
            
            for adapter in network_adapters:
                # 获取适配器名称
                adapter_name = adapter.Name
                if adapter_name:
                    # 将适配器添加到列表中（无论是否启用）
                    adapters.append(adapter_name)
                    
                    # 保存网卡对象引用
                    self.network_adapters[adapter_name] = adapter
                    
                    # 获取对应的配置（如果有）
                    config = config_map.get(adapter.GUID)
                    if config:
                        self.wmi_adapters[adapter_name] = config
                    
                    # 同时保存到映射中，用于兼容性
                    self.adapter_map[adapter_name] = adapter_name
            
            # 在主线程中更新UI
            self.root.after(0, lambda: self._update_adapter_ui(adapters))
            
        except Exception as e:
            self.root.after(0, lambda: self._handle_adapter_error(str(e)))
    
    def _update_adapter_ui(self, adapters):
        """在主线程中更新适配器UI"""
        try:
            # 更新下拉列表
            self.adapter_combo['values'] = adapters
            
            # 如果有适配器，选择第一个并显示信息
            if adapters:
                # 保持当前选择（如果存在于新列表中）
                current_selection = self.adapter_var.get()
                if current_selection in adapters:
                    self.adapter_combo.set(current_selection)
                else:
                    self.adapter_combo.set(adapters[0])
                self.on_adapter_selected()
            
            self.status_var.set(f"找到 {len(adapters)} 个网络适配器")
            
        except Exception as e:
            self.status_var.set(f"更新UI失败: {str(e)}")
    
    def _handle_adapter_error(self, error_msg):
        """处理适配器错误"""
        self.status_var.set(f"错误: {error_msg}")
        messagebox.showerror("错误", f"获取网络适配器时出错:\n{error_msg}")
    
    def on_adapter_selected(self, event=None):
        """当选择网络适配器时"""
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            return
            
        try:
            self.status_var.set("正在获取IP信息...")
            # 使用after方法延迟更新，避免阻塞UI
            self.root.after(10, lambda: self._do_get_adapter_info(adapter_name))
            
        except Exception as e:
            self.status_var.set(f"错误: {str(e)}")
    
    def _do_get_adapter_info(self, adapter_name):
        """实际获取适配器信息（在后台线程中）"""
        try:
            # 重新获取网卡对象，确保状态是最新的
            network_adapters = self.wmi.Win32_NetworkAdapter(Name=adapter_name)
            if len(network_adapters) > 0:
                # 更新网卡对象引用
                adapter = network_adapters[0]
                self.network_adapters[adapter_name] = adapter
            else:
                adapter = self.network_adapters.get(adapter_name)
                if not adapter:
                    self.root.after(0, lambda: self.status_var.set("未找到适配器信息"))
                    return
                
            # 重新获取网卡配置信息
            nic_configs = self.wmi.Win32_NetworkAdapterConfiguration(Index=adapter.Index)
            if len(nic_configs) > 0:
                # 更新网卡配置引用
                nic = nic_configs[0]
                self.wmi_adapters[adapter_name] = nic
            else:
                nic = self.wmi_adapters.get(adapter_name)
            
            # 在主线程中更新UI
            self.root.after(0, lambda: self._update_adapter_info(adapter, nic))
            
        except Exception as e:
            self.root.after(0, lambda: self._handle_adapter_info_error(str(e)))
    
    def _update_adapter_info(self, adapter, nic):
        """在主线程中更新适配器信息UI"""
        try:
            if nic:
                self.display_wmi_ip_info(nic, adapter)
                self.status_var.set("IP信息已更新")
                # 更新内网测试目标（默认取第一个网关）
                try:
                    if hasattr(self, 'lan_ping_target') and hasattr(nic, 'DefaultIPGateway') and nic.DefaultIPGateway:
                        self.lan_ping_target.set(str(nic.DefaultIPGateway[0]))
                except Exception:
                    pass
            else:
                # 即使没有配置信息，也显示基本的网卡信息
                self.display_adapter_info(adapter)
                self.status_var.set("仅显示基本网卡信息")
                
        except Exception as e:
            self.status_var.set(f"错误: {str(e)}")
            messagebox.showerror("错误", f"获取IP信息时出错:\n{str(e)}")
    
    def _handle_adapter_info_error(self, error_msg):
        """处理适配器信息错误"""
        self.status_var.set(f"获取适配器信息失败: {error_msg}")
    
    def _get_adapter_status(self, adapter, nic=None):
        """获取适配器的详细状态"""
        try:
            # 检查适配器的多个状态属性
            net_enabled = getattr(adapter, 'NetEnabled', None)
            net_connection_status = getattr(adapter, 'NetConnectionStatus', None)
            adapter_enabled = getattr(adapter, 'AdapterEnabled', None)
            
            # 调试信息（可选）
            # print(f"调试 - {adapter.Name}: NetEnabled={net_enabled}, NetConnectionStatus={net_connection_status}, AdapterEnabled={adapter_enabled}")
            
            # 首先检查是否被禁用（但需要特殊处理NetConnectionStatus=7的情况）
            if adapter_enabled is False:
                return "已禁用"
            
            # 对于NetEnabled=False的情况，需要结合NetConnectionStatus判断
            if net_enabled is False:
                # 如果NetConnectionStatus=7（媒体断开连接），则显示未连接而不是禁用
                if net_connection_status == 7:
                    return "未连接"
                # 如果NetConnectionStatus=4或5（硬件不存在或被禁用），则显示禁用
                elif net_connection_status in [4, 5]:
                    return "已禁用"
                # 其他情况显示禁用
                else:
                    return "已禁用"
            
            # 如果NetEnabled为None，可能是虚拟适配器，检查其他属性
            if net_enabled is None:
                if adapter_enabled is False:
                    return "已禁用"
                elif adapter_enabled is True:
                    # 虚拟适配器，检查是否有IP地址
                    if nic and hasattr(nic, 'IPAddress') and nic.IPAddress and len(nic.IPAddress) > 0:
                        return "已连接"
                    else:
                        return "未连接"
                else:
                    return "未知状态"
            
            # 检查连接状态
            if net_connection_status is not None:
                # NetConnectionStatus 值含义：
                # 0 = 断开连接
                # 1 = 连接中
                # 2 = 已连接
                # 3 = 断开连接中
                # 4 = 硬件不存在
                # 5 = 硬件被禁用
                # 6 = 硬件故障
                # 7 = 媒体断开连接
                # 8 = 正在验证
                # 9 = 验证失败
                # 10 = 连接失败
                # 11 = 断开连接中
                # 12 = 已断开连接
                
                # 优先根据NetConnectionStatus判断状态
                if net_connection_status == 7:
                    return "未连接"  # 媒体断开连接
                elif net_connection_status in [0, 3, 11, 12]:
                    return "未连接"  # 其他断开状态
                elif net_connection_status in [1, 8]:
                    return "连接中"
                elif net_connection_status == 2:
                    return "已连接"
                elif net_connection_status in [4, 5]:
                    return "已禁用"
                elif net_connection_status in [6, 9, 10]:
                    return "连接失败"
                else:
                    return f"状态{net_connection_status}"
            
            # 如果NetConnectionStatus不可用，回退到IP地址检查
            if nic and hasattr(nic, 'IPAddress') and nic.IPAddress and len(nic.IPAddress) > 0:
                # 检查是否有有效的IPv4地址
                for ip in nic.IPAddress:
                    if re.match(r'^(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)){3}$', ip):
                        return "已连接"
                return "已启用(仅IPv6)"
            else:
                return "未连接"
                
        except Exception as e:
            # 如果出现异常，回退到基本状态
            print(f"获取适配器状态时出错: {e}")
            return "已启用" if getattr(adapter, 'NetEnabled', False) else "已禁用"
    
    def display_adapter_info(self, adapter):
        """显示网卡基本信息（用于禁用状态）"""
        self.ip_info_text.config(state=tk.NORMAL)
        self.ip_info_text.delete(1.0, tk.END)
        
        # 显示适配器基本信息
        info_lines = []
        info_lines.append(f"适配器名称: {adapter.Name}")
        info_lines.append(f"MAC地址: {adapter.MACAddress if adapter.MACAddress else '未知'}")
        
        # 尝试获取配置信息以确定准确状态
        nic = None
        try:
            nic_configs = self.wmi.Win32_NetworkAdapterConfiguration(Index=adapter.Index)
            if len(nic_configs) > 0:
                nic = nic_configs[0]
        except Exception:
            pass
        
        # 使用新的状态判断方法
        status = self._get_adapter_status(adapter, nic)
        info_lines.append(f"适配器状态: {status}")
        
        # 尝试获取静态IP配置（即使网卡未连接）
        try:
            # 尝试通过Index获取配置信息
            nic_configs = self.wmi.Win32_NetworkAdapterConfiguration(Index=adapter.Index)
            if len(nic_configs) > 0:
                nic = nic_configs[0]
                
                # 保存配置信息
                self.wmi_adapters[adapter.Name] = nic
                
                # 显示IP地址信息
                if hasattr(nic, 'IPAddress') and nic.IPAddress:
                    info_lines.append("\nIP地址:")
                    for i, ip in enumerate(nic.IPAddress):
                        if i < len(nic.IPSubnet):
                            info_lines.append(f"  {ip} / {nic.IPSubnet[i]}")
                        else:
                            info_lines.append(f"  {ip}")
                else:
                    info_lines.append("\nIP地址: 未配置")
                
                # DHCP状态
                info_lines.append(f"\nDHCP启用: {'是' if nic.DHCPEnabled else '否'}")
                
                # 提取配置信息到输入框
                self.extract_wmi_config(nic)
            else:
                info_lines.append("\nIP地址: 未配置")
                info_lines.append("\nDHCP启用: 未知")
                
                # 清空IP配置输入框
                self.ip_var.set("")
                self.mask_var.set("")
                self.gateway_var.set("")
                self.dns_var.set("")
        except Exception as e:
            info_lines.append(f"\n获取IP配置失败: {str(e)}")
            
            # 清空IP配置输入框
            self.ip_var.set("")
            self.mask_var.set("")
            self.gateway_var.set("")
            self.dns_var.set("")
        
        # 显示到文本框
        for line in info_lines:
            self.ip_info_text.insert(tk.END, line + '\n')
        
        self.ip_info_text.config(state=tk.DISABLED)
    
    def refresh_ip_info(self):
        """刷新当前选中适配器的IP信息"""
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            return
            
        try:
            self.status_var.set("正在刷新IP信息...")
            # 使用after方法延迟更新，避免阻塞UI
            self.root.after(10, lambda: self._do_refresh_ip_info(adapter_name))
            
        except Exception as e:
            self.status_var.set(f"错误: {str(e)}")
    
    def _do_refresh_ip_info(self, adapter_name):
        """实际执行IP信息刷新（在后台线程中）"""
        try:
            adapter = self.network_adapters.get(adapter_name)
            if not adapter:
                self.root.after(0, lambda: self.status_var.set("未找到适配器信息"))
                return
                
            nic = self.wmi_adapters.get(adapter_name)
            
            # 在主线程中更新UI
            self.root.after(0, lambda: self._update_ip_info_ui(adapter, nic))
            
        except Exception as e:
            self.root.after(0, lambda: self._handle_refresh_ip_error(str(e)))
    
    def _update_ip_info_ui(self, adapter, nic):
        """在主线程中更新IP信息UI"""
        try:
            if nic:
                self.display_wmi_ip_info(nic, adapter)
                self.status_var.set("IP信息已刷新")
            else:
                # 即使没有配置信息，也显示基本的网卡信息
                self.display_adapter_info(adapter)
                self.status_var.set("仅显示基本网卡信息")
                
        except Exception as e:
            self.status_var.set(f"更新IP信息失败: {str(e)}")
    
    def _handle_refresh_ip_error(self, error_msg):
        """处理刷新IP信息错误"""
        self.status_var.set(f"刷新IP信息失败: {error_msg}")
        messagebox.showerror("错误", f"刷新IP信息时出错:\n{error_msg}")
    
    def display_wmi_ip_info(self, nic, adapter):
        """显示WMI获取的IP信息"""
        self.ip_info_text.config(state=tk.NORMAL)
        self.ip_info_text.delete(1.0, tk.END)
        
        # 提取配置信息到输入框
        self.extract_wmi_config(nic)
        
        # 显示详细信息
        info_lines = []
        info_lines.append(f"适配器名称: {adapter.Name}")
        info_lines.append(f"MAC地址: {adapter.MACAddress if adapter.MACAddress else '未知'}")
        
        # 使用新的状态判断方法
        status = self._get_adapter_status(adapter, nic)
        info_lines.append(f"适配器状态: {status}")
        
        # IP地址信息
        if hasattr(nic, 'IPAddress') and nic.IPAddress and len(nic.IPAddress) > 0:
            info_lines.append("\nIP地址:")
            
            # 分离IPv4和IPv6地址
            ipv4_addresses = []
            ipv6_addresses = []
            
            for i, ip in enumerate(nic.IPAddress):
                mask = nic.IPSubnet[i] if i < len(nic.IPSubnet) else ""
                
                # 检查是IPv4还是IPv6
                if re.match(r'^(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)){3}$', ip):
                    ipv4_addresses.append((ip, mask))
                else:
                    try:
                        socket.inet_pton(socket.AF_INET6, ip)
                        ipv6_addresses.append((ip, mask))
                    except (socket.error, AttributeError):
                        pass  # 忽略无效的IP地址
            
            # 显示IPv4地址
            if ipv4_addresses:
                info_lines.append("  IPv4地址:")
                
                # 根据网关确定主IP
                main_ip_found = False
                main_ip_info = None
                extra_ips = []
                
                if hasattr(nic, 'DefaultIPGateway') and nic.DefaultIPGateway:
                    # 获取第一个IPv4网关
                    gateway = None
                    for gw in nic.DefaultIPGateway:
                        if re.match(r'^(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)){3}$', gw):
                            gateway = gw
                            break
                    
                    if gateway:
                        # 检查哪个IP地址与网关在同一网段
                        for ip, mask in ipv4_addresses:
                            if self.is_same_network(ip, gateway, mask):
                                main_ip_info = (ip, mask)
                                main_ip_found = True
                            else:
                                extra_ips.append((ip, mask))
                
                # 如果没有找到与网关匹配的主IP，使用第一个IPv4地址
                if not main_ip_found:
                    main_ip_info = ipv4_addresses[0]
                    extra_ips = ipv4_addresses[1:]
                
                # 显示主IP
                if main_ip_info:
                    main_ip, main_mask = main_ip_info
                    info_lines.append(f"    主IP: {main_ip} / {main_mask}")
                
                # 显示额外IP
                if extra_ips:
                    info_lines.append("    额外IP:")
                    for i, (ip, mask) in enumerate(extra_ips, 1):
                        info_lines.append(f"      {i}. {ip} / {mask}")
            
            # 显示IPv6地址
            if ipv6_addresses:
                info_lines.append("  IPv6地址:")
                for i, (ip, mask) in enumerate(ipv6_addresses, 1):
                    info_lines.append(f"    {i}. {ip} / {mask}")
        else:
            info_lines.append("\nIP地址: 未配置")
        
        # DHCP状态
        info_lines.append(f"\nDHCP启用: {'是' if nic.DHCPEnabled else '否'}")
        if nic.DHCPEnabled and nic.DHCPServer:
            info_lines.append(f"DHCP服务器: {nic.DHCPServer}")
        
        # 网关信息
        if nic.DefaultIPGateway:
            info_lines.append("\n默认网关:")
            for gateway in nic.DefaultIPGateway:
                info_lines.append(f"  {gateway}")
        else:
            info_lines.append("\n默认网关: 未配置")
        
        # DNS信息
        if nic.DNSServerSearchOrder:
            info_lines.append("\nDNS服务器:")
            for dns in nic.DNSServerSearchOrder:
                info_lines.append(f"  {dns}")
        else:
            info_lines.append("\nDNS服务器: 未配置")
        
        # 显示到文本框
        for line in info_lines:
            self.ip_info_text.insert(tk.END, line + '\n')
        
        # 添加点击事件，用于选择IP地址
        self.ip_info_text.tag_configure("ip_clickable", foreground="blue", underline=1)
        
        # 为每个IP地址添加点击事件
        if hasattr(nic, 'IPAddress') and nic.IPAddress and len(nic.IPAddress) > 0:
            # 主IP地址行号
            main_ip_line = 5  # 适配器名称、MAC地址、状态行、空行、IP地址标题行后的主IP行
            
            # 为主IP添加点击事件
            if self.is_valid_ip(nic.IPAddress[0]):
                main_ip = nic.IPAddress[0]
                start_pos = f"{main_ip_line}.8"  # "  主IP: "后面的位置
                end_pos = f"{main_ip_line}.{8 + len(main_ip)}"
                
                self.ip_info_text.tag_add("ip_clickable", start_pos, end_pos)
                mask = nic.IPSubnet[0] if len(nic.IPSubnet) > 0 else ""
                self.ip_info_text.tag_bind("ip_clickable", "<Button-1>", 
                                         lambda e, ip=main_ip, mask=mask: self.select_ip_address(ip, mask))
            
            # 为额外IP添加点击事件
            if len(nic.IPAddress) > 1:
                extra_ip_start_line = main_ip_line + 2  # 主IP行、额外IP标题行后的第一个额外IP行
                
                for i in range(1, len(nic.IPAddress)):
                    if self.is_valid_ip(nic.IPAddress[i]):
                        extra_ip = nic.IPAddress[i]
                        line_num = extra_ip_start_line + (i-1)
                        start_pos = f"{line_num}.6"  # "    数字. "后面的位置
                        end_pos = f"{line_num}.{6 + len(extra_ip)}"
                        
                        self.ip_info_text.tag_add("ip_clickable", start_pos, end_pos)
                        mask = nic.IPSubnet[i] if i < len(nic.IPSubnet) else ""
                        self.ip_info_text.tag_bind("ip_clickable", "<Button-1>", 
                                                 lambda e, ip=extra_ip, mask=mask: self.select_ip_address(ip, mask))
        
        self.ip_info_text.config(state=tk.DISABLED)
    
    def select_ip_address(self, ip, mask):
        """选择点击的IP地址，并更新输入框"""
        if not ip:
            return
        
        # 获取当前适配器
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            return
            
        nic = self.wmi_adapters.get(adapter_name)
        if not nic:
            return
        
        # 检查IP类型（IPv4或IPv6）
        is_ipv4 = re.match(r'^(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)){3}$', ip) is not None
        
        # 分离IPv4和IPv6地址
        ipv4_addresses = []
        ipv6_addresses = []
        
        if hasattr(nic, 'IPAddress') and nic.IPAddress:
            for i, addr in enumerate(nic.IPAddress):
                addr_mask = nic.IPSubnet[i] if i < len(nic.IPSubnet) else ""
                if re.match(r'^(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)){3}$', addr):
                    ipv4_addresses.append((addr, addr_mask))
                else:
                    try:
                        socket.inet_pton(socket.AF_INET6, addr)
                        ipv6_addresses.append((addr, addr_mask))
                    except (socket.error, AttributeError):
                        pass
        
        # 检查选择的IP是主IP还是额外IP
        is_main_ip = False
        
        # 对于IPv4，检查是否是第一个IPv4地址
        if is_ipv4 and ipv4_addresses and ipv4_addresses[0][0] == ip:
            is_main_ip = True
        # 对于IPv6，如果没有IPv4地址，检查是否是第一个IPv6地址
        elif not is_ipv4 and not ipv4_addresses and ipv6_addresses and ipv6_addresses[0][0] == ip:
            is_main_ip = True
        
        if is_main_ip:
            # 如果是主IP，更新主IP输入框
            self.ip_var.set(ip)
            if mask:
                self.mask_var.set(mask)
                
            # 更新网关和DNS（使用与选中IP相关的配置）
            if nic.DefaultIPGateway:
                # 尝试匹配IP类型的网关
                for gateway in nic.DefaultIPGateway:
                    gateway_is_ipv4 = re.match(r'^(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)){3}$', gateway) is not None
                    if gateway_is_ipv4 == is_ipv4:
                        self.gateway_var.set(gateway)
                        break
                else:
                    # 如果没有匹配的，使用第一个
                    self.gateway_var.set(nic.DefaultIPGateway[0])
                
            if nic.DNSServerSearchOrder:
                # 尝试匹配IP类型的DNS
                for dns in nic.DNSServerSearchOrder:
                    dns_is_ipv4 = re.match(r'^(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)){3}$', dns) is not None
                    if dns_is_ipv4 == is_ipv4:
                        self.dns_var.set(dns)
                        break
                else:
                    # 如果没有匹配的，使用第一个
                    self.dns_var.set(nic.DNSServerSearchOrder[0])
        else:
            # 如果是额外IP，添加到额外IP列表
            # 先检查是否已经存在于额外IP列表中
            ip_exists = False
            for ip_info in self.extra_ips:
                if ip_info['ip_var'].get() == ip:
                    ip_exists = True
                    break
            
            # 如果不存在，添加到额外IP列表
            if not ip_exists:
                self.add_extra_ip()
                if len(self.extra_ips) > 0:
                    self.extra_ips[-1]['ip_var'].set(ip)
                    if mask:
                        self.extra_ips[-1]['mask_var'].set(mask)
    
    def extract_wmi_config(self, nic):
        """从WMI配置中提取当前设置"""
        # 清空当前设置
        self.ip_var.set("")
        self.mask_var.set("")
        self.gateway_var.set("")
        self.dns_var.set("")
        
        # 清空额外IP
        self.clear_extra_ips()
        
        # 提取IP地址
        if hasattr(nic, 'IPAddress') and nic.IPAddress and len(nic.IPAddress) > 0:
            # 分离IPv4和IPv6地址
            ipv4_addresses = []
            ipv6_addresses = []
            
            for i, ip in enumerate(nic.IPAddress):
                mask = nic.IPSubnet[i] if i < len(nic.IPSubnet) else ""
                
                # 检查是IPv4还是IPv6
                if re.match(r'^(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)){3}$', ip):
                    ipv4_addresses.append((ip, mask))
                else:
                    try:
                        socket.inet_pton(socket.AF_INET6, ip)
                        ipv6_addresses.append((ip, mask))
                    except (socket.error, AttributeError):
                        pass  # 忽略无效的IP地址
            
            # 根据网关确定主IP
            main_ip_found = False
            if ipv4_addresses and hasattr(nic, 'DefaultIPGateway') and nic.DefaultIPGateway:
                # 获取第一个IPv4网关
                gateway = None
                for gw in nic.DefaultIPGateway:
                    if re.match(r'^(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)){3}$', gw):
                        gateway = gw
                        break
                
                if gateway:
                    # 检查哪个IP地址与网关在同一网段
                    for ip, mask in ipv4_addresses:
                        if self.is_same_network(ip, gateway, mask):
                            # 找到与网关同网段的IP作为主IP
                            self.ip_var.set(ip)
                            self.mask_var.set(mask)
                            main_ip_found = True
                            
                            # 添加其他IPv4地址为额外IP
                            for other_ip, other_mask in ipv4_addresses:
                                if other_ip != ip:
                                    self.add_extra_ip()
                                    if len(self.extra_ips) > 0:
                                        self.extra_ips[-1]['ip_var'].set(other_ip)
                                        self.extra_ips[-1]['mask_var'].set(other_mask)
                            break
            
            # 如果没有找到与网关匹配的主IP，使用第一个IPv4地址
            if not main_ip_found and ipv4_addresses:
                main_ip, main_mask = ipv4_addresses[0]
                self.ip_var.set(main_ip)
                self.mask_var.set(main_mask)
                
                # 添加额外IPv4地址
                for i, (ip, mask) in enumerate(ipv4_addresses[1:]):
                    self.add_extra_ip()
                    if i < len(self.extra_ips):
                        self.extra_ips[i]['ip_var'].set(ip)
                        self.extra_ips[i]['mask_var'].set(mask)
            
            # 如果没有IPv4地址但有IPv6地址，使用第一个IPv6地址作为主IP
            elif not main_ip_found and ipv6_addresses:
                main_ip, main_mask = ipv6_addresses[0]
                self.ip_var.set(main_ip)
                self.mask_var.set(main_mask)
                
                # 添加额外IPv6地址
                for i, (ip, mask) in enumerate(ipv6_addresses[1:]):
                    self.add_extra_ip()
                    if i < len(self.extra_ips):
                        self.extra_ips[i]['ip_var'].set(ip)
                        self.extra_ips[i]['mask_var'].set(mask)
        
        # 如果没有IP地址但有静态IP配置，尝试从注册表或其他配置获取
        elif not nic.DHCPEnabled:
            try:
                # 尝试从WMI配置中获取静态IP配置
                static_ips = getattr(nic, 'IPAddress', None)
                static_masks = getattr(nic, 'IPSubnet', None)
                
                if static_ips and len(static_ips) > 0:
                    # 分离IPv4和IPv6静态地址
                    ipv4_static = []
                    ipv6_static = []
                    
                    for i, ip in enumerate(static_ips):
                        mask = static_masks[i] if i < len(static_masks) else ""
                        
                        # 检查是IPv4还是IPv6
                        if re.match(r'^(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)){3}$', ip):
                            ipv4_static.append((ip, mask))
                        else:
                            try:
                                socket.inet_pton(socket.AF_INET6, ip)
                                ipv6_static.append((ip, mask))
                            except (socket.error, AttributeError):
                                pass  # 忽略无效的IP地址
                    
                    # 优先使用IPv4地址
                    if ipv4_static:
                        main_ip, main_mask = ipv4_static[0]
                        self.ip_var.set(main_ip)
                        self.mask_var.set(main_mask)
                        
                        # 添加额外IPv4静态地址
                        for i, (ip, mask) in enumerate(ipv4_static[1:]):
                            self.add_extra_ip()
                            if i < len(self.extra_ips):
                                self.extra_ips[i]['ip_var'].set(ip)
                                self.extra_ips[i]['mask_var'].set(mask)
                    
                    # 如果没有IPv4静态地址但有IPv6静态地址
                    elif ipv6_static:
                        main_ip, main_mask = ipv6_static[0]
                        self.ip_var.set(main_ip)
                        self.mask_var.set(main_mask)
                        
                        # 添加额外IPv6静态地址
                        for i, (ip, mask) in enumerate(ipv6_static[1:]):
                            self.add_extra_ip()
                            if i < len(self.extra_ips):
                                self.extra_ips[i]['ip_var'].set(ip)
                                self.extra_ips[i]['mask_var'].set(mask)
            except Exception:
                # 如果获取失败，不做任何处理
                pass
        
        # 提取网关
        if hasattr(nic, 'DefaultIPGateway') and nic.DefaultIPGateway:
            for gateway in nic.DefaultIPGateway:
                if self.is_valid_ip(gateway):
                    self.gateway_var.set(gateway)
                    break
        
        # 提取DNS
        if hasattr(nic, 'DNSServerSearchOrder') and nic.DNSServerSearchOrder:
            for dns in nic.DNSServerSearchOrder:
                if self.is_valid_ip(dns):
                    self.dns_var.set(dns)
                    break
    
    def set_static_ip(self):
        """使用WMI设置静态IP"""
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            messagebox.showwarning("警告", "请先选择网络适配器")
            return
        
        ip = self.ip_var.get().strip()
        mask = self.mask_var.get().strip()
        gateway = self.gateway_var.get().strip()
        dns = self.dns_var.get().strip()
        
        if not ip or not mask:
            messagebox.showwarning("警告", "请填写IP地址和子网掩码")
            return
        
        # 详细验证IP地址
        is_valid, error_msg = self.validate_ip_address(ip, "IP地址")
        if not is_valid:
            messagebox.showerror("错误", error_msg)
            return
        
        # 详细验证子网掩码
        is_valid, error_msg = self.validate_ip_address(mask, "子网掩码")
        if not is_valid:
            messagebox.showerror("错误", error_msg)
            return
        
        # 验证网关（如果填写了）
        if gateway:
            is_valid, error_msg = self.validate_ip_address(gateway, "默认网关")
            if not is_valid:
                messagebox.showerror("错误", error_msg)
                return
        
        # 验证DNS（如果填写了）
        if dns:
            is_valid, error_msg = self.validate_ip_address(dns, "DNS服务器")
            if not is_valid:
                messagebox.showerror("错误", error_msg)
                return
        
        try:
            self.status_var.set("正在设置静态IP...")
            self.root.update()
            
            nic = self.wmi_adapters.get(adapter_name)
            if not nic:
                messagebox.showerror("错误", "未找到适配器配置")
                return
            
            # 准备所有IP地址（主IP + 额外IP）
            all_ips = [ip]
            all_masks = [mask]
            
            # 添加额外IP地址
            extra_ips = self.get_extra_ips()
            for extra_ip, extra_mask in extra_ips:
                all_ips.append(extra_ip)
                all_masks.append(extra_mask)
            
            # 设置静态IP（包括额外IP）
            result = nic.EnableStatic(all_ips, all_masks)
            if result[0] == 0:
                # 设置网关
                if gateway and self.is_valid_ip(gateway):
                    nic.SetGateways(DefaultIPGateway=[gateway])
                
                # 设置DNS
                if dns and self.is_valid_ip(dns):
                    nic.SetDNSServerSearchOrder(DNSServerSearchOrder=[dns])
                
                messagebox.showinfo("成功", f"静态IP设置成功\n主IP: {ip}\n额外IP: {len(extra_ips)}个")
                self.status_var.set("静态IP设置完成")
                self.refresh_ip_info()
            else:
                error_msg = f"设置失败，错误代码: {result[0]}"
                messagebox.showerror("错误", f"设置静态IP失败:\n{error_msg}")
                self.status_var.set("设置静态IP失败")
                
        except Exception as e:
            messagebox.showerror("错误", f"设置静态IP时出错:\n{str(e)}")
            self.status_var.set(f"错误: {str(e)}")
    
    def set_dhcp(self):
        """使用WMI设置DHCP"""
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            messagebox.showwarning("警告", "请先选择网络适配器")
            return
        
        try:
            self.status_var.set("正在设置DHCP...")
            self.root.update()
            
            nic = self.wmi_adapters.get(adapter_name)
            if not nic:
                messagebox.showerror("错误", "未找到适配器配置")
                return
            
            # 启用DHCP
            result = nic.EnableDHCP()
            if result[0] == 0:
                # 设置DNS为DHCP
                nic.SetDNSServerSearchOrder(DNSServerSearchOrder=[])
                
                messagebox.showinfo("成功", "DHCP设置成功")
                self.status_var.set("DHCP设置完成")
                self.refresh_ip_info()
            else:
                error_msg = f"设置失败，错误代码: {result[0]}"
                messagebox.showerror("错误", f"设置DHCP失败:\n{error_msg}")
                self.status_var.set("设置DHCP失败")
                
        except Exception as e:
            messagebox.showerror("错误", f"设置DHCP时出错:\n{str(e)}")
            self.status_var.set(f"错误: {str(e)}")
    
    def export_config(self):
        """导出配置"""
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            messagebox.showwarning("警告", "请先选择网络适配器")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            title="保存IP配置"
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(f"网络适配器: {adapter_name}\n")
                    f.write(f"主IP地址: {self.ip_var.get()}\n")
                    f.write(f"子网掩码: {self.mask_var.get()}\n")
                    f.write(f"默认网关: {self.gateway_var.get()}\n")
                    f.write(f"DNS服务器: {self.dns_var.get()}\n")
                    
                    # 导出额外IP
                    extra_ips = self.get_extra_ips()
                    if extra_ips:
                        f.write(f"\n额外IP地址:\n")
                        for i, (ip, mask) in enumerate(extra_ips, 1):
                            f.write(f"  {i}. {ip}/{mask}\n")
                    
                    f.write(f"\n导出时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                
                messagebox.showinfo("成功", f"配置已导出到:\n{filename}")
                self.status_var.set("配置导出完成")
                
            except Exception as e:
                messagebox.showerror("错误", f"导出配置时出错:\n{str(e)}")
    
    def validate_ipv4_entry(self, value):
        """验证IPv4输入格式和数值范围"""
        if len(value) > 15:
            return False
        if value == "":
            return True
        
        # 检查格式
        pattern = r'^\d{0,3}(\.\d{0,3}){0,3}$'
        if not re.match(pattern, value):
            return False
        
        # 检查每个段的值是否超过255
        parts = value.split('.')
        for part in parts:
            if part and int(part) > 255:
                return False
        
        return True

    def is_valid_ip(self, ip):
        """校验IP地址（支持IPv4和IPv6）"""
        if not ip or not ip.strip():
            return False
            
        ip = ip.strip()
        
        # 检查是否为IPv4
        ipv4_pattern = r'^(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)){3}$'
        if re.match(ipv4_pattern, ip):
            return True
            
        # 检查是否为IPv6
        try:
            socket.inet_pton(socket.AF_INET6, ip)
            return True
        except (socket.error, AttributeError):
            return False
    
    def validate_ip_address(self, ip, field_name="IP地址"):
        """详细验证IP地址并返回错误信息"""
        if not ip or not ip.strip():
            return False, f"{field_name}不能为空"
        
        ip = ip.strip()
        
        # 检查基本格式
        if not re.match(r'^\d{1,3}(\.\d{1,3}){3}$', ip):
            return False, f"{field_name}格式错误，应为x.x.x.x格式"
        
        # 检查每个段的值
        parts = ip.split('.')
        for i, part in enumerate(parts):
            try:
                value = int(part)
                if value < 0 or value > 255:
                    return False, f"{field_name}第{i+1}段值({value})超出范围(0-255)"
            except ValueError:
                return False, f"{field_name}第{i+1}段({part})不是有效数字"
        
        return True, ""
    
    def is_same_network(self, ip1, ip2, subnet_mask):
        """检查两个IP地址是否在同一网段"""
        try:
            # 将IP地址和子网掩码转换为整数
            def ip_to_int(ip):
                parts = ip.split('.')
                return (int(parts[0]) << 24) + (int(parts[1]) << 16) + (int(parts[2]) << 8) + int(parts[3])
            
            def mask_to_int(mask):
                if '.' in mask:
                    # 点分十进制格式的子网掩码
                    parts = mask.split('.')
                    return (int(parts[0]) << 24) + (int(parts[1]) << 16) + (int(parts[2]) << 8) + int(parts[3])
                else:
                    # CIDR格式的子网掩码
                    prefix_len = int(mask)
                    return (0xFFFFFFFF << (32 - prefix_len)) & 0xFFFFFFFF
            
            ip1_int = ip_to_int(ip1)
            ip2_int = ip_to_int(ip2)
            mask_int = mask_to_int(subnet_mask)
            
            # 计算网络地址
            network1 = ip1_int & mask_int
            network2 = ip2_int & mask_int
            
            return network1 == network2
        except (ValueError, IndexError):
            return False

    def disable_adapter(self):
        """禁用选中的网卡"""
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            messagebox.showwarning("警告", "请先选择网络适配器")
            return
        
        # 确认对话框
        result = messagebox.askyesno("确认", f"确定要禁用网卡 '{adapter_name}' 吗？\n\n禁用后网络连接将中断。")
        if not result:
            return
        
        try:
            self.status_var.set("正在禁用网卡...")
            self.root.update()
            
            # 直接使用保存的网卡对象引用
            adapter = self.network_adapters.get(adapter_name)
            if not adapter:
                messagebox.showerror("错误", "未找到网卡对象")
                return
            
            # 检查网卡是否已经禁用（使用新的状态判断逻辑）
            current_status = self._get_adapter_status(adapter)
            if current_status == "已禁用":
                messagebox.showinfo("提示", f"网卡 '{adapter_name}' 已经处于禁用状态")
                self.status_var.set("网卡已禁用")
                return
            
            # 禁用网卡
            result = adapter.Disable()
            if result[0] == 0:
                messagebox.showinfo("成功", f"网卡 '{adapter_name}' 已禁用")
                self.status_var.set("网卡禁用完成")
                self.refresh_network_adapters()  # 刷新适配器列表
            else:
                error_msg = f"禁用失败，错误代码: {result[0]}"
                messagebox.showerror("错误", f"禁用网卡失败:\n{error_msg}")
                self.status_var.set("禁用网卡失败")
                
        except Exception as e:
            messagebox.showerror("错误", f"禁用网卡时出错:\n{str(e)}")
            self.status_var.set(f"错误: {str(e)}")

    def enable_adapter(self):
        """启用选中的网卡"""
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            messagebox.showwarning("警告", "请先选择网络适配器")
            return
        
        # 确认对话框
        result = messagebox.askyesno("确认", f"确定要启用网卡 '{adapter_name}' 吗？")
        if not result:
            return
        
        try:
            self.status_var.set("正在启用网卡...")
            self.root.update()
            
            # 直接使用保存的网卡对象引用
            adapter = self.network_adapters.get(adapter_name)
            if not adapter:
                messagebox.showerror("错误", "未找到网卡对象")
                return
            
            # 检查网卡是否已经启用（使用新的状态判断逻辑）
            current_status = self._get_adapter_status(adapter)
            if current_status in ["已连接", "未连接", "连接中", "已启用(仅IPv6)"]:
                messagebox.showinfo("提示", f"网卡 '{adapter_name}' 已经处于启用状态")
                self.status_var.set("网卡已启用")
                return
            
            # 启用网卡
            result = adapter.Enable()
            if result[0] == 0:
                messagebox.showinfo("成功", f"网卡 '{adapter_name}' 已启用")
                self.status_var.set("网卡启用完成")
                
                # 等待一段时间，让系统完成网卡启用
                self.root.after(1000, self._refresh_after_enable, adapter_name)
            else:
                error_msg = f"启用失败，错误代码: {result[0]}"
                messagebox.showerror("错误", f"启用网卡失败:\n{error_msg}")
                self.status_var.set("启用网卡失败")
                
        except Exception as e:
            messagebox.showerror("错误", f"启用网卡时出错:\n{str(e)}")
            self.status_var.set(f"错误: {str(e)}")
    
    def _refresh_after_enable(self, adapter_name):
        """在启用网卡后刷新网卡状态"""
        try:
            # 重新获取网卡对象，确保状态是最新的
            network_adapters = self.wmi.Win32_NetworkAdapter(Name=adapter_name)
            if len(network_adapters) > 0:
                # 更新网卡对象引用
                self.network_adapters[adapter_name] = network_adapters[0]
                
                # 刷新网卡列表和信息
                self.refresh_network_adapters()
        except Exception as e:
            self.status_var.set(f"刷新网卡状态时出错: {str(e)}")
            # 即使出错也尝试刷新网卡列表
            self.refresh_network_adapters()

    def reset_network(self):
        """重置网络设置"""
        # 确认对话框
        result = messagebox.askyesno("确认", 
                                   "确定要重置网络设置吗？\n\n"
                                   "这将执行以下操作：\n"
                                   "1. netsh int ip reset\n"
                                   "2. netsh winsock reset\n\n"
                                   "重置后需要重启计算机才能生效。\n"
                                   "此操作将清除所有网络配置！")
        if not result:
            return
        
        try:
            self.status_var.set("正在重置网络设置...")
            self.root.update()
            
            # 执行网络重置命令
            commands = [
                "netsh int ip reset",
                "netsh winsock reset"
            ]
            
            success_count = 0
            for cmd in commands:
                try:
                    self.status_var.set(f"正在执行: {cmd}")
                    self.root.update()
                    
                    # 方法1: 直接执行命令
                    try:
                        result = subprocess.run(cmd, shell=True, capture_output=True, 
                                              text=True, encoding='gbk', errors='ignore',
                                              creationflags=subprocess.CREATE_NO_WINDOW)
                        
                        # 检查命令是否成功执行
                        # 对于 netsh int ip reset，即使返回码是1，如果输出包含"完成"字样，也认为是成功的
                        if result.returncode == 0 or ("完成" in result.stdout and "重置" in cmd):
                            success_count += 1
                            self.status_var.set(f"命令执行成功: {cmd}")
                        else:
                            # 方法2: 尝试以管理员权限运行
                            try:
                                # 使用runas命令
                                admin_cmd = f'runas /user:Administrator "{cmd}"'
                                result = subprocess.run(admin_cmd, shell=True, capture_output=True,
                                                      text=True, encoding='gbk', errors='ignore')
                                
                                if result.returncode == 0:
                                    success_count += 1
                                    self.status_var.set(f"命令执行成功: {cmd}")
                                else:
                                    messagebox.showwarning("警告", 
                                                         f"命令执行可能有问题:\n{cmd}\n\n"
                                                         f"返回码: {result.returncode}\n"
                                                         f"错误输出: {result.stderr}\n\n"
                                                         f"请确保以管理员身份运行程序。")
                            except Exception as e2:
                                messagebox.showwarning("警告", 
                                                     f"执行命令时出错:\n{cmd}\n\n"
                                                     f"错误: {str(e2)}\n\n"
                                                     f"请确保以管理员身份运行程序。")
                                
                    except Exception as e1:
                        messagebox.showwarning("警告", 
                                             f"执行命令时出错:\n{cmd}\n\n"
                                             f"错误: {str(e1)}\n\n"
                                             f"请确保以管理员身份运行程序。")
                            
                except Exception as e:
                    messagebox.showwarning("警告", 
                                         f"执行命令时出错:\n{cmd}\n\n"
                                         f"错误: {str(e)}\n\n"
                                         f"请确保以管理员身份运行程序。")
            
            if success_count > 0:
                # 询问是否重启
                restart_result = messagebox.askyesno("重启确认", 
                                                   "网络重置完成！\n\n"
                                                   "需要重启计算机才能生效。\n\n"
                                                   "是否立即重启计算机？\n"
                                                   "点击'是'立即重启\n"
                                                   "点击'否'稍后手动重启")
                
                if restart_result:
                    # 立即重启
                    try:
                        subprocess.run("shutdown /r /t 0", shell=True, check=True)
                    except Exception as e:
                        messagebox.showerror("错误", f"重启失败:\n{str(e)}\n\n请手动重启计算机。")
                else:
                    messagebox.showinfo("提示", "请稍后手动重启计算机以使网络重置生效。")
            else:
                messagebox.showerror("错误", 
                                   "所有网络重置命令都执行失败。\n\n"
                                   "可能的原因：\n"
                                   "1. 没有管理员权限\n"
                                   "2. 命令被系统阻止\n"
                                   "3. 网络服务正在运行\n\n"
                                   "建议：\n"
                                   "1. 以管理员身份运行程序\n"
                                   "2. 关闭所有网络相关程序\n"
                                   "3. 手动在命令提示符中执行这些命令")
            
            self.status_var.set("网络重置完成")
            
        except Exception as e:
            messagebox.showerror("错误", f"重置网络时出错:\n{str(e)}")
            self.status_var.set(f"错误: {str(e)}")

    def _on_window_resize(self, event):
        # 更新缩放因子
        self._update_scale_factor()

    def _update_scale_factor(self):
        # 获取当前窗口大小
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        
        # 计算缩放因子
        new_scale_factor = min(width / self.base_width, height / self.base_height)
        
        # 只有当缩放因子真正改变时才更新
        if abs(new_scale_factor - self.scale_factor) > 0.01:
            self.scale_factor = new_scale_factor
            # 应用缩放到所有UI元素
            self._apply_scaling()
    
    def _get_scaled_font(self, base_size, weight='normal', family='Segoe UI'):
        """获取缩放后的字体"""
        scaled_size = max(int(base_size * self.scale_factor), 8)  # 最小字体大小为8
        return (family, scaled_size, weight)
    
    def _get_scaled_width(self, base_width):
        """获取缩放后的宽度"""
        return max(int(base_width * self.scale_factor), 1)
    
    def _get_scaled_height(self, base_height):
        """获取缩放后的高度"""
        return max(int(base_height * self.scale_factor), 1)
    
    def _get_scaled_padding(self, base_padding):
        """获取缩放后的内边距"""
        if isinstance(base_padding, (tuple, list)):
            return tuple(max(int(p * self.scale_factor), 0) for p in base_padding)
        else:
            return max(int(base_padding * self.scale_factor), 0)
    
    def _register_scalable_widget(self, widget, widget_type, **kwargs):
        """注册需要缩放的UI元素"""
        self.scalable_widgets.append({
            'widget': widget,
            'type': widget_type,
            'params': kwargs
        })
    
    def _apply_scaling(self):
        """应用缩放到所有UI元素"""
        try:
            # 更新所有注册的UI元素
            for item in self.scalable_widgets:
                widget = item['widget']
                widget_type = item['type']
                params = item['params']
                
                if widget_type == 'label':
                    if 'font' in params:
                        widget.configure(font=self._get_scaled_font(*params['font']))
                elif widget_type == 'entry':
                    if 'width' in params:
                        widget.configure(width=self._get_scaled_width(params['width']))
                    if 'font' in params:
                        widget.configure(font=self._get_scaled_font(*params['font']))
                elif widget_type == 'button':
                    if 'width' in params:
                        widget.configure(width=self._get_scaled_width(params['width']))
                    if 'font' in params:
                        widget.configure(font=self._get_scaled_font(*params['font']))
                elif widget_type == 'text':
                    if 'width' in params:
                        widget.configure(width=self._get_scaled_width(params['width']))
                    if 'height' in params:
                        widget.configure(height=self._get_scaled_height(params['height']))
                    if 'font' in params:
                        widget.configure(font=self._get_scaled_font(*params['font']))
                elif widget_type == 'combobox':
                    if 'width' in params:
                        widget.configure(width=self._get_scaled_width(params['width']))
                    if 'font' in params:
                        widget.configure(font=self._get_scaled_font(*params['font']))
            
            # 更新样式配置
            self._update_styles()
            
        except Exception as e:
            print(f"应用缩放时出错: {e}")
    
    def _update_styles(self):
        """更新所有样式配置"""
        try:
            style = ttk.Style()
            
            # 更新选项卡标题样式 - 确保文字居中
            style.configure('TNotebook.Tab', 
                          padding=self._get_scaled_padding((12, 6)),
                          font=self._get_scaled_font(11),
                          anchor='center',  # 设置文字居中对齐
                          justify='center')  # 设置文本居中对齐
            
            # 针对Windows系统的额外样式设置
            style.map('TNotebook.Tab',
                     background=[('selected', '#E3F2FD'), ('active', '#F5F5F5')],
                     foreground=[('selected', '#1976D2'), ('active', '#424242')])
            
            # 更新卡片样式
            style.configure('CardTitle.TLabel', 
                          font=self._get_scaled_font(13, 'bold'), 
                          foreground='#0F172A', 
                          background='#FFFFFF')
            style.configure('CardItemLeft.TLabel', 
                          font=self._get_scaled_font(10),
                          foreground='#6B7280', 
                          background='#FFFFFF')
            style.configure('CardItemRight.TLabel', 
                          font=self._get_scaled_font(10),
                          foreground='#0F172A', 
                          background='#FFFFFF')
            style.configure('CardValue.TLabel', 
                          font=self._get_scaled_font(11, family='Consolas'), 
                          foreground='#111827', 
                          background='#FFFFFF')
            
            # 更新按钮样式
            button_padding = self._get_scaled_padding((10, 6))
            style.configure('Segment.TButton', 
                          padding=button_padding, 
                          relief='flat',
                          background='#1F2937', 
                          foreground='#E5E7EB')
            style.configure('SegmentSelected.TButton', 
                          padding=button_padding, 
                          relief='flat',
                          background=ACCENT, 
                          foreground='#0B1221')
            
        except Exception as e:
            print(f"更新样式时出错: {e}")
    
    def _bind_hw_tab_events(self):
        """绑定硬件信息页签切换事件"""
        try:
            # 获取notebook和硬件信息页签
            notebook = None
            for widget in self.root.winfo_children():
                if isinstance(widget, ttk.Frame):
                    for child in widget.winfo_children():
                        if isinstance(child, ttk.Notebook):
                            notebook = child
                            break
                    if notebook:
                        break
            
            if notebook:
                # 绑定页签切换事件
                notebook.bind('<<NotebookTabChanged>>', self._on_notebook_tab_changed)
        except Exception as e:
            print(f"绑定硬件信息页签事件失败: {e}")
    
    def _on_notebook_tab_changed(self, event):
        """处理页签切换事件"""
        try:
            notebook = event.widget
            current_tab = notebook.select()
            tab_id = notebook.index(current_tab)
            
            # 检查是否是硬件信息页签（通常是第5个页签，索引为4）
            if tab_id == 4:  # 硬件信息页签
                self._on_hw_tab_enter()
            else:
                self._on_hw_tab_leave()
        except Exception as e:
            print(f"处理页签切换事件失败: {e}")
    
    def _ensure_taskbar_icon(self):
        """确保任务栏图标正确显示"""
        try:
            # 强制更新窗口图标
            self.root.update_idletasks()
            
            # 确保窗口有正确的图标
            if hasattr(self, '_icon_img') and self._icon_img:
                self.root.iconphoto(True, self._icon_img)
                print("任务栏图标已确保设置")
            
            # 设置窗口类图标（关键修复）
            try:
                import ctypes
                hwnd = self.root.winfo_id()
                ico_path = resource_path("ip_manager.ico")
                
                if os.path.exists(ico_path):
                    # 加载图标
                    icon_handle = ctypes.windll.user32.LoadImageW(
                        None, ico_path, 1, 0, 0, 0x00000010  # IMAGE_ICON, LR_LOADFROMFILE
                    )
                    
                    if icon_handle:
                        # 设置窗口图标
                        ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 0, icon_handle)  # WM_SETICON, ICON_SMALL
                        ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 1, icon_handle)  # WM_SETICON, ICON_BIG
                        
                        # 设置类图标
                        ctypes.windll.user32.SetClassLongW(hwnd, -14, icon_handle)  # GCL_HICONSM
                        ctypes.windll.user32.SetClassLongW(hwnd, -15, icon_handle)  # GCL_HICON
                        
                        # 强制重绘窗口
                        ctypes.windll.user32.InvalidateRect(hwnd, None, True)
                        ctypes.windll.user32.UpdateWindow(hwnd)
                        
                        print("✓ 窗口类图标设置成功")
                        
                        # 延迟后再次设置，确保生效
                        self.root.after(1000, self._force_set_icon_again)
            except Exception as e:
                print(f"设置窗口类图标失败: {e}")
            
            # 在Windows上，有时需要强制刷新任务栏
            try:
                import ctypes
                ctypes.windll.user32.SetWindowPos(
                    self.root.winfo_id(), 0, 0, 0, 0, 0,
                    0x0001 | 0x0002 | 0x0004  # SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER
                )
            except Exception:
                pass
                
        except Exception as e:
            print(f"确保任务栏图标失败: {e}")
    
    def _force_set_icon_again(self):
        """延迟后再次强制设置图标"""
        try:
            import ctypes
            hwnd = self.root.winfo_id()
            ico_path = resource_path("ip_manager.ico")
            
            if os.path.exists(ico_path):
                # 加载图标
                icon_handle = ctypes.windll.user32.LoadImageW(
                    None, ico_path, 1, 0, 0, 0x00000010  # IMAGE_ICON, LR_LOADFROMFILE
                )
                
                if icon_handle:
                    # 再次设置窗口图标
                    ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 0, icon_handle)  # WM_SETICON, ICON_SMALL
                    ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 1, icon_handle)  # WM_SETICON, ICON_BIG
                    
                    # 再次设置类图标
                    ctypes.windll.user32.SetClassLongW(hwnd, -14, icon_handle)  # GCL_HICONSM
                    ctypes.windll.user32.SetClassLongW(hwnd, -15, icon_handle)  # GCL_HICON
                    
                    # 强制重绘
                    ctypes.windll.user32.InvalidateRect(hwnd, None, True)
                    ctypes.windll.user32.UpdateWindow(hwnd)
                    
                    print("✓ 延迟图标设置成功")
        except Exception as e:
            print(f"延迟图标设置失败: {e}")

    def _init_system_tray(self):
        """初始化系统托盘"""
        if not SYSTEM_TRAY_AVAILABLE:
            return
            
        try:
            # 创建托盘图标
            icon_image = self._create_tray_icon()
            
            # 创建托盘菜单
            menu = pystray.Menu(
                pystray.MenuItem("显示主窗口", self._show_window),
                pystray.MenuItem("退出", self._quit_application)
            )
            
            # 创建托盘图标，设置单击事件
            self.tray_icon = pystray.Icon(
                "ip_manager",
                icon_image,
                "Windows IP地址管理器",
                menu
            )
            
            # 设置单击托盘图标显示窗口
            self.tray_icon.on_click = self._on_tray_click
            
        except Exception as e:
            print(f"初始化系统托盘失败: {e}")
    
    def _create_tray_icon(self):
        """创建托盘图标"""
        try:
            # 优先使用ico文件
            ico_path = resource_path("ip_manager.ico")
            if os.path.exists(ico_path):
                return Image.open(ico_path)
        except Exception:
            pass
            
        try:
            # 备用ico文件名
            for ico_name in ["IP管理器.ico", "icon.ico", "app.ico"]:
                ico_path = resource_path(ico_name)
                if os.path.exists(ico_path):
                    return Image.open(ico_path)
        except Exception:
            pass
            
        try:
            # 使用PNG文件
            png_path = resource_path("ip_manager_256x256.png")
            if os.path.exists(png_path):
                return Image.open(png_path)
        except Exception:
            pass
        
        # 如果图标文件不存在，创建一个简单的图标
        return self._create_default_tray_icon()
    
    def _create_default_tray_icon(self):
        """创建默认托盘图标"""
        # 创建一个16x16的图标
        image = Image.new('RGBA', (16, 16), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        
        # 绘制一个简单的图标
        draw.rectangle([2, 2, 13, 13], fill=(0, 132, 255, 255), outline=(255, 255, 255, 255))
        draw.text((4, 4), "IP", fill=(255, 255, 255, 255))
        
        return image
    
    def _on_closing(self):
        """处理窗口关闭事件"""
        if self.is_minimized_to_tray:
            # 如果已经最小化到托盘，则真正退出
            self._quit_application()
        else:
            # 第一次关闭时询问用户
            if not self.first_close_asked:
                self.first_close_asked = True
                result = messagebox.askyesnocancel(
                    "关闭确认", 
                    "您想要关闭程序还是最小化到系统托盘？\n\n选择：\n• 是(Y) - 最小化到系统托盘\n• 否(N) - 关闭程序\n• 取消 - 返回程序",
                    icon=messagebox.QUESTION
                )
                
                if result is True:  # 用户选择"是" - 最小化到托盘
                    self._minimize_to_tray()
                elif result is False:  # 用户选择"否" - 关闭程序
                    self._quit_application()
                else:  # 用户选择"取消" - 返回程序
                    return
            else:
                # 不是第一次关闭，直接最小化到托盘
                self._minimize_to_tray()
    
    def _on_tray_click(self, icon, event):
        """托盘图标单击事件"""
        try:
            if self.is_minimized_to_tray:
                self._show_window()
        except Exception as e:
            print(f"托盘单击事件处理失败: {e}")
    
    def _minimize_to_tray(self):
        """最小化到系统托盘"""
        if not SYSTEM_TRAY_AVAILABLE or not self.tray_icon:
            return
            
        try:
            # 隐藏主窗口
            self.root.withdraw()
            self.is_minimized_to_tray = True
            
            # 启动托盘图标（如果还没有启动）
            if not self.tray_icon.visible:
                threading.Thread(target=self._run_tray_icon, daemon=True).start()
                
        except Exception as e:
            print(f"最小化到托盘失败: {e}")
    
    def _run_tray_icon(self):
        """在后台线程中运行托盘图标"""
        try:
            if self.tray_icon and not self.tray_icon.visible:
                self.tray_icon.run()
        except Exception as e:
            print(f"运行托盘图标失败: {e}")
    
    def _show_window(self, icon=None, item=None):
        """显示主窗口"""
        try:
            # 显示主窗口
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
            self.is_minimized_to_tray = False
            
            # 确保任务栏图标重新显示
            self._ensure_taskbar_icon()
            
            # 注意：不要停止托盘图标，只是隐藏窗口
            # 托盘图标会继续在后台运行，等待下次最小化
                
        except Exception as e:
            print(f"显示窗口失败: {e}")
    
    def _quit_application(self, icon=None, item=None):
        """退出应用程序"""
        try:
            # 停止托盘图标
            if self.tray_icon and self.tray_icon.visible:
                self.tray_icon.stop()
            
            # 取消硬件信息定时器
            self._cancel_hw_tick()
            
            # 销毁主窗口
            self.root.quit()
            self.root.destroy()
            
        except Exception as e:
            print(f"退出应用程序失败: {e}")

def is_admin():
    """检查是否具有管理员权限"""
    try:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def main():
    # 检查操作系统
    if not os.name == 'nt':
        messagebox.showerror("错误", "此程序仅支持Windows系统")
        return
    
    # 启动时尝试申请管理员权限（UAC），避免后续操作失败
    try:
        # 标记位，避免重复提权循环
        elevated_flag = "--elevated"
        if elevated_flag in sys.argv:
            # 清理掉标记，避免影响后续参数处理
            sys.argv = [sys.argv[0]] + [a for a in sys.argv[1:] if a != elevated_flag]
        elif not is_admin():
            import ctypes
            # 组装提权启动参数
            if getattr(sys, 'frozen', False):
                # 打包后的exe，直接以自身提权
                exe_path = sys.executable
                params = " ".join(sys.argv[1:] + [elevated_flag])
            else:
                # 脚本运行，使用python解释器提权执行当前脚本
                exe_path = sys.executable
                script_path = os.path.abspath(__file__)
                params = f'"{script_path}" ' + " ".join(sys.argv[1:] + [elevated_flag])
            rc = ctypes.windll.shell32.ShellExecuteW(None, "runas", exe_path, params, None, 1)
            if rc > 32:
                # 成功启动了提权进程，退出当前普通权限进程
                return
    except Exception:
        # 忽略提权异常，继续以普通权限运行（关键操作前仍会有提示）
        pass
    
    root = tk.Tk()
    # 全局苹果风样式：更圆的控件、浅灰背景、细分隔
    try:
        style = ttk.Style()
        style.theme_use('clam')
        # 主题主色：浅绿背景 + macOS 蓝色按钮
        PRIMARY_BG = '#ECFDF5'   # 浅绿（柔和）
        PRIMARY_FG = '#222'
        ACCENT = '#0A84FF'       # macOS 蓝
        ACCENT_HOVER = '#0063E1'
        ACCENT_PRESSED = '#0052BF'
        SUBTLE = '#6B7280'

        # 全局字体（苹果风：SF Pro 替补为 Segoe UI / Arial）
        try:
            base_font = tkfont.Font(family='SF Pro Text', size=11)
            title_font = tkfont.Font(family='SF Pro Display', size=18, weight='bold')
        except Exception:
            base_font = tkfont.Font(family='Segoe UI', size=11)
            title_font = tkfont.Font(family='Segoe UI', size=18, weight='bold')

        root.option_add('*Font', base_font)

        # 窗体背景与分隔色
        root.configure(bg=PRIMARY_BG)
        style.configure('.', background=PRIMARY_BG)
        style.configure('TFrame', background=PRIMARY_BG)
        style.configure('TLabelframe', background=PRIMARY_BG, borderwidth=1, relief='groove')
        style.configure('TLabelframe.Label', background=PRIMARY_BG, foreground=SUBTLE, padding=2)
        style.configure('TLabel', background=PRIMARY_BG, foreground=PRIMARY_FG)
        # 圆角模拟：给分组框和文本框添加更柔的外观（Tk本身不支持真正圆角，采用留白与内边距权衡）
        # 输入类控件圆角感（通过内边距和浅边框模拟）
        style.configure('TEntry', padding=6, fieldbackground="#FFFFFF")
        style.configure('TCombobox', padding=4, fieldbackground="#FFFFFF")
        # 按钮扁平浅色与悬停
        # 基础按钮统一为macOS蓝
        style.configure('TButton', padding=(10, 6), relief='flat', background=ACCENT, foreground='#FFFFFF', borderwidth=0)
        style.map('TButton',
                  background=[('active', ACCENT_HOVER), ('pressed', ACCENT_PRESSED)],
                  relief=[('pressed', 'sunken')])
    except Exception:
        pass
    app = IPManager(root)
    
    # 图标已在IPManager.__init__()中设置，这里不需要重复设置
    
    root.mainloop()

if __name__ == "__main__":
    main()