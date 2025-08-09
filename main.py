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
 
# Global theme constants (Apple-like)
PRIMARY_BG = '#ECFDF5'  # light green background
PRIMARY_FG = '#222'
SUBTLE = '#6B7280'
ACCENT = '#0A84FF'       # macOS blue
ACCENT_HOVER = '#0063E1'
ACCENT_PRESSED = '#0052BF'
import subprocess
import winreg
import glob
import re
import socket
import threading
from datetime import datetime
import wmi
import win32com.client



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
        self.root.geometry("1024x680")  # 默认窗口大小 1024x680
        self.root.resizable(True, True)

        # 设置应用图标（EXE与窗口内一致）
        try:
            ico_path = resource_path("IP管理器.ico")
            if os.path.exists(ico_path):
                self.root.iconbitmap(ico_path)
        except Exception:
            pass

        try:
            png_path = resource_path("ip_manager_256x256.png")
            if os.path.exists(png_path):
                # 保持引用避免被GC
                self._icon_img = tk.PhotoImage(file=png_path)
                self.root.iconphoto(True, self._icon_img)
        except Exception:
            pass
        
        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')
        
        # 自定义样式 - 增大字体
        style.configure('TButton', font=('Arial', 11))
        style.configure('TLabel', font=('Arial', 11))
        style.configure('TEntry', font=('Arial', 11))
        style.configure('TCombobox', font=('Arial', 11))
        style.configure('TLabelframe', font=('Arial', 11, 'bold'))
        style.configure('TLabelframe.Label', font=('Arial', 11, 'bold'))
        
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
        style.configure('Status.TLabel', font=('Arial', 11), background='#E3F2FD', foreground='#1976D2')
        
        # 初始化WMI
        self.wmi = None
        try:
            # 使用wmi.WMI()初始化WMI
            self.wmi = wmi.WMI()
        except Exception as e:
            messagebox.showerror("错误", f"WMI初始化失败: {str(e)}\n\n请确保以管理员身份运行程序。")
            # 继续创建UI，但禁用需要WMI的功能
        
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

    def toggle_win11_autologon(self) -> None:
        """Win11 自动登录开关：
        - 设置 HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\PasswordLess\Device\DevicePasswordLessBuildVersion = 0 开启可见的 netplwiz 取消密码登录
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
                "设置已应用，已为你打开 netplwiz。\n\n在弹出的窗口中取消勾选“要使用本计算机，用户必须输入用户名和密码”，并输入密码以启用开机自动登录。"
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
        
        # 配置网格权重（左侧更宽，右侧更窄）
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=4)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="Windows IP地址管理器", 
                               font=("Arial", 18, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 5))
        
        # 版本号（小字体，灰色）
        version_text = f"版本: {get_version_string()}"
        version_label = ttk.Label(main_frame, text=version_text, 
                                 font=("Arial", 11), foreground="gray")
        version_label.grid(row=1, column=0, columnspan=3, pady=(0, 10))

        # 顶部：网络测试（ping）行（移动到右侧容器中显示）
        
        # 右侧区域：先创建容器（后面把适配器区放进去）
        
        # 两列布局：左宽右窄
        main_frame.columnconfigure(0, weight=4)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # 左侧：当前IP信息（位于版本号下方，靠左整列）
        left_frame = ttk.LabelFrame(main_frame, text="当前IP信息", padding="8")
        left_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=8)
        
        # IP信息文本框
        self.ip_info_text = tk.Text(left_frame, height=16, width=45, state=tk.DISABLED, font=("Consolas", 11), 
                                   bg='#F5F5F5', fg='#333333', selectbackground='#2196F3', selectforeground='white')
        self.ip_info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
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
        # 右侧放置刷新按钮
        refresh_btn = ttk.Button(adapter_section, text="刷新", command=self.refresh_network_adapters, width=8, style='Refresh.TButton')
        refresh_btn.grid(row=1, column=1, sticky=tk.E, padx=(6, 0))
        self.add_button_hover_effect(refresh_btn)

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
        # 隐藏Notebook默认页签，避免与自定义工具栏重复
        try:
            style_hide = ttk.Style()
            style_hide.layout('TNotebook.Tab', [])
        except Exception:
            pass
        toolbar = ttk.Frame(right_frame)
        toolbar.grid(row=2, column=0, sticky=tk.E, pady=(0, 6))
        # 苹果风：分段按钮（Segmented）样式
        seg_style = ttk.Style()
        seg_style.configure('Segment.TButton', padding=(10, 6), relief='flat',
                            background='#D1FAE5', foreground='#222')
        seg_style.map('Segment.TButton',
                      background=[('active', '#A7F3D0'), ('pressed', '#6EE7B7')])
        seg_style.configure('SegmentSelected.TButton', padding=(10, 6), relief='flat',
                            background=ACCENT, foreground='#FFFFFF')

        self._seg_btns = []
        def set_segment_selection(active_idx: int) -> None:
            for idx, btn in enumerate(self._seg_btns):
                btn.configure(style='SegmentSelected.TButton' if idx == active_idx else 'Segment.TButton')

        def switch_tab(i:int):
            notebook.select(i)
            set_segment_selection(i)

        # 与实际添加的四个选项卡保持一致
        tb_texts = ["IP配置", "额外IP地址", "网卡控制", "工具"]
        for i, t in enumerate(tb_texts):
            b = ttk.Button(toolbar, text=t, style='Segment.TButton', command=lambda idx=i: switch_tab(idx))
            b.grid(row=0, column=i, padx=(0 if i==0 else 2, 0))
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
        self.ip_entry = ttk.Entry(main_ip_frame, textvariable=self.ip_var, width=22, validate='key')
        self.ip_entry['validatecommand'] = (self.ip_entry.register(self.validate_ipv4_entry), '%P')
        self.ip_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))

        ttk.Label(main_ip_frame, text="子网掩码:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.mask_var = tk.StringVar()
        self.mask_entry = ttk.Entry(main_ip_frame, textvariable=self.mask_var, width=22, validate='key')
        self.mask_entry['validatecommand'] = (self.mask_entry.register(self.validate_ipv4_entry), '%P')
        self.mask_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))

        ttk.Label(main_ip_frame, text="默认网关:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.gateway_var = tk.StringVar()
        self.gateway_entry = ttk.Entry(main_ip_frame, textvariable=self.gateway_var, width=22, validate='key')
        self.gateway_entry['validatecommand'] = (self.gateway_entry.register(self.validate_ipv4_entry), '%P')
        self.gateway_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))

        ttk.Label(main_ip_frame, text="DNS服务器:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.dns_var = tk.StringVar()
        self.dns_entry = ttk.Entry(main_ip_frame, textvariable=self.dns_var, width=22, validate='key')
        self.dns_entry['validatecommand'] = (self.dns_entry.register(self.validate_ipv4_entry), '%P')
        self.dns_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))

        # 构建公用操作按钮组（设置静态IP/设置DHCP/刷新/导出配置）
        def build_ip_actions(parent: tk.Widget) -> ttk.Frame:
            actions = ttk.Frame(parent)
            actions.columnconfigure(0, weight=1)
            actions.columnconfigure(1, weight=1)

            btn1 = ttk.Button(actions, text="设置静态IP", command=self.set_static_ip, width=12, style='StaticIP.TButton')
            btn1.grid(row=0, column=0, padx=3, pady=3, sticky=tk.E)
            self.add_button_hover_effect(btn1)

            btn2 = ttk.Button(actions, text="设置DHCP", command=self.set_dhcp, width=12, style='DHCP.TButton')
            btn2.grid(row=0, column=1, padx=3, pady=3, sticky=tk.W)
            self.add_button_hover_effect(btn2)

            btn3 = ttk.Button(actions, text="刷新", command=self.refresh_ip_info, width=12, style='RefreshInfo.TButton')
            btn3.grid(row=1, column=0, padx=3, pady=3, sticky=tk.E)
            self.add_button_hover_effect(btn3)

            btn4 = ttk.Button(actions, text="导出配置", command=self.export_config, width=12, style='Export.TButton')
            btn4.grid(row=1, column=1, padx=3, pady=3, sticky=tk.W)
            self.add_button_hover_effect(btn4)

            return actions

        # 在“IP配置”页底部加入操作区
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

        add_ip_btn = ttk.Button(btn_frame, text="添加IP", command=self.add_extra_ip, width=10, style='AddIP.TButton')
        add_ip_btn.grid(row=0, column=0, padx=2, pady=2, sticky=tk.E)
        self.add_button_hover_effect(add_ip_btn)

        clear_ip_btn = ttk.Button(btn_frame, text="清空", command=self.clear_extra_ips, width=10, style='Clear.TButton')
        clear_ip_btn.grid(row=0, column=1, padx=2, pady=2, sticky=tk.W)
        self.add_button_hover_effect(clear_ip_btn)

        # 在“额外IP地址”页底部加入同样的操作区
        actions2 = build_ip_actions(tab_extra)
        actions2.grid(row=2, column=0, sticky=(tk.W, tk.E))

        # Tab3: 网卡控制
        tab_adapter = ttk.Frame(notebook)
        notebook.add(tab_adapter, text="网卡控制")
        # Tab4: 工具（分组展示，更整洁）
        tab_tools = ttk.Frame(notebook)
        notebook.add(tab_tools, text="工具")
        tab_tools.columnconfigure(0, weight=1)
        tab_tools.columnconfigure(1, weight=1)

        # 分组1：网络诊断
        diag_group = ttk.Labelframe(tab_tools, text="网络诊断", padding=8)
        diag_group.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=4, pady=(8, 4))
        diag_group.columnconfigure(1, weight=1)

        ttk.Label(diag_group, text="内网测试（网关）").grid(row=0, column=0, sticky=tk.E, padx=(0,8), pady=4)
        lan_entry = ttk.Entry(diag_group, textvariable=self.lan_ping_target, width=28)
        lan_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=4)
        lan_btn = ttk.Button(diag_group, text="内网测试 (ping)", style='Ping.TButton',
                             command=lambda: self.run_ping(self.lan_ping_target.get()))
        lan_btn.grid(row=0, column=2, sticky=tk.W, padx=(8,0), pady=4)

        ttk.Label(diag_group, text="网络测试（外网）").grid(row=1, column=0, sticky=tk.E, padx=(0,8), pady=4)
        self.ping_target = tk.StringVar(value="www.baidu.com")
        ping_entry = ttk.Entry(diag_group, textvariable=self.ping_target, width=28)
        ping_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=4)
        ping_btn = ttk.Button(diag_group, text="网络测试 (ping)", style='Ping.TButton',
                              command=lambda: self.run_ping(self.ping_target.get()))
        ping_btn.grid(row=1, column=2, sticky=tk.W, padx=(8,0), pady=4)

        # 网络追踪（tracert）
        ttk.Label(diag_group, text="网络追踪（tracert）").grid(row=2, column=0, sticky=tk.E, padx=(0,8), pady=4)
        self.tracert_target = tk.StringVar(value="www.baidu.com")
        tracert_entry = ttk.Entry(diag_group, textvariable=self.tracert_target, width=28)
        tracert_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=4)
        tracert_btn = ttk.Button(diag_group, text="网络追踪 (tracert)", style='Ping.TButton',
                                 command=lambda: self.run_tracert(self.tracert_target.get()))
        tracert_btn.grid(row=2, column=2, sticky=tk.W, padx=(8,0), pady=4)

        # 分组2：系统工具
        sys_group = ttk.Labelframe(tab_tools, text="系统工具", padding=8)
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
        adapter_control_frame = ttk.LabelFrame(tab_adapter, text="网卡控制", padding="5")
        adapter_control_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        adapter_control_frame.columnconfigure(0, weight=1)
        adapter_control_frame.columnconfigure(1, weight=1)

        self.disable_btn = ttk.Button(adapter_control_frame, text="禁用网卡", 
                                     command=self.disable_adapter, width=12, style='Disable.TButton')
        self.disable_btn.grid(row=0, column=0, padx=3, pady=3, sticky=tk.E)
        self.add_button_hover_effect(self.disable_btn)

        self.enable_btn = ttk.Button(adapter_control_frame, text="启用网卡", 
                                    command=self.enable_adapter, width=12, style='Enable.TButton')
        self.enable_btn.grid(row=0, column=1, padx=3, pady=3, sticky=tk.W)
        self.add_button_hover_effect(self.enable_btn)

        self.reset_network_btn = ttk.Button(adapter_control_frame, text="重置网络", 
                                           command=self.reset_network, width=12, style='Reset.TButton')
        self.reset_network_btn.grid(row=1, column=0, padx=3, pady=3, sticky=tk.E)
        self.add_button_hover_effect(self.reset_network_btn)
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W, style='Status.TLabel')
        status_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 适配器映射
        self.adapter_map = {}
        self.wmi_adapters = {}
        
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
            self.root.update()
            
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
            self.status_var.set(f"错误: {str(e)}")
            messagebox.showerror("错误", f"获取网络适配器时出错:\n{str(e)}")
    
    def on_adapter_selected(self, event=None):
        """当选择网络适配器时"""
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            return
            
        try:
            self.status_var.set("正在获取IP信息...")
            self.root.update()
            
            # 重新获取网卡对象，确保状态是最新的
            network_adapters = self.wmi.Win32_NetworkAdapter(Name=adapter_name)
            if len(network_adapters) > 0:
                # 更新网卡对象引用
                adapter = network_adapters[0]
                self.network_adapters[adapter_name] = adapter
            else:
                adapter = self.network_adapters.get(adapter_name)
                if not adapter:
                    self.status_var.set("未找到适配器信息")
                    return
                
            # 重新获取网卡配置信息
            nic_configs = self.wmi.Win32_NetworkAdapterConfiguration(Index=adapter.Index)
            if len(nic_configs) > 0:
                # 更新网卡配置引用
                nic = nic_configs[0]
                self.wmi_adapters[adapter_name] = nic
            else:
                nic = self.wmi_adapters.get(adapter_name)
            
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
    
    def display_adapter_info(self, adapter):
        """显示网卡基本信息（用于禁用状态）"""
        self.ip_info_text.config(state=tk.NORMAL)
        self.ip_info_text.delete(1.0, tk.END)
        
        # 显示适配器基本信息
        info_lines = []
        info_lines.append(f"适配器名称: {adapter.Name}")
        info_lines.append(f"MAC地址: {adapter.MACAddress if adapter.MACAddress else '未知'}")
        info_lines.append(f"适配器状态: {'已启用' if adapter.NetEnabled else '已禁用'}")
        
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
            self.root.update()
            
            adapter = self.network_adapters.get(adapter_name)
            if not adapter:
                self.status_var.set("未找到适配器信息")
                return
                
            nic = self.wmi_adapters.get(adapter_name)
            if nic:
                self.display_wmi_ip_info(nic, adapter)
                self.status_var.set("IP信息已刷新")
            else:
                # 即使没有配置信息，也显示基本的网卡信息
                self.display_adapter_info(adapter)
                self.status_var.set("仅显示基本网卡信息")
                
        except Exception as e:
            self.status_var.set(f"错误: {str(e)}")
            messagebox.showerror("错误", f"刷新IP信息时出错:\n{str(e)}")
    
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
        info_lines.append(f"适配器状态: {'已启用' if adapter.NetEnabled else '已禁用'}")
        
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
            
            # 检查网卡是否已经禁用
            if not adapter.NetEnabled:
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
            
            # 检查网卡是否已经启用
            if adapter.NetEnabled:
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
    
    # 设置窗口图标（如果有的话）
    try:
        root.iconbitmap('icon.ico')
    except:
        pass
    
    root.mainloop()

if __name__ == "__main__":
    main()