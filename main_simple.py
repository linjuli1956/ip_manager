import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import re
import os
import sys
from datetime import datetime

class IPManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows IP地址管理器")
        self.root.geometry("1000x700")
        self.root.resizable(True, True)
        
        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')
        
        # 初始化变量
        self.extra_ips = []
        
        # 创建UI
        self.setup_ui()
        
        # 获取网络适配器
        self.refresh_network_adapters()
        
    def setup_ui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="Windows IP地址管理器", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 15))
        
        # 网络适配器选择
        ttk.Label(main_frame, text="网络适配器:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.adapter_var = tk.StringVar()
        self.adapter_combo = ttk.Combobox(main_frame, textvariable=self.adapter_var, 
                                         state="readonly", width=40)
        self.adapter_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        self.adapter_combo.bind('<<ComboboxSelected>>', self.on_adapter_selected)
        
        refresh_btn = ttk.Button(main_frame, text="刷新", command=self.refresh_network_adapters)
        refresh_btn.grid(row=1, column=2, padx=(10, 0), pady=5)
        
        # 创建左右分栏
        paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned_window.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # 左侧：当前IP信息
        left_frame = ttk.LabelFrame(paned_window, text="当前IP信息", padding="10")
        paned_window.add(left_frame, weight=1)
        
        # IP信息文本框
        self.ip_info_text = tk.Text(left_frame, height=20, width=45, state=tk.DISABLED)
        self.ip_info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 滚动条
        ip_scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.ip_info_text.yview)
        ip_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.ip_info_text.configure(yscrollcommand=ip_scrollbar.set)
        
        # 配置左侧框架的网格权重
        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(0, weight=1)
        
        # 右侧：IP配置
        right_frame = ttk.LabelFrame(paned_window, text="IP配置", padding="10")
        paned_window.add(right_frame, weight=1)
        
        # 主IP配置框架
        main_ip_frame = ttk.LabelFrame(right_frame, text="主IP配置", padding="8")
        main_ip_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        main_ip_frame.columnconfigure(1, weight=1)
        
        # IP地址
        ttk.Label(main_ip_frame, text="IP地址:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.ip_var = tk.StringVar()
        self.ip_entry = ttk.Entry(main_ip_frame, textvariable=self.ip_var, width=25, validate='key')
        self.ip_entry['validatecommand'] = (self.ip_entry.register(self.validate_ipv4_entry), '%P')
        self.ip_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(10, 0))
        
        # 子网掩码
        ttk.Label(main_ip_frame, text="子网掩码:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.mask_var = tk.StringVar()
        self.mask_entry = ttk.Entry(main_ip_frame, textvariable=self.mask_var, width=25, validate='key')
        self.mask_entry['validatecommand'] = (self.mask_entry.register(self.validate_ipv4_entry), '%P')
        self.mask_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2, padx=(10, 0))
        
        # 默认网关
        ttk.Label(main_ip_frame, text="默认网关:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.gateway_var = tk.StringVar()
        self.gateway_entry = ttk.Entry(main_ip_frame, textvariable=self.gateway_var, width=25, validate='key')
        self.gateway_entry['validatecommand'] = (self.gateway_entry.register(self.validate_ipv4_entry), '%P')
        self.gateway_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=2, padx=(10, 0))
        
        # DNS服务器
        ttk.Label(main_ip_frame, text="DNS服务器:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.dns_var = tk.StringVar()
        self.dns_entry = ttk.Entry(main_ip_frame, textvariable=self.dns_var, width=25, validate='key')
        self.dns_entry['validatecommand'] = (self.dns_entry.register(self.validate_ipv4_entry), '%P')
        self.dns_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=2, padx=(10, 0))
        
        # 多IP配置框架
        multi_ip_frame = ttk.LabelFrame(right_frame, text="额外IP地址", padding="8")
        multi_ip_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        multi_ip_frame.columnconfigure(1, weight=1)
        
        # 额外IP列表
        self.extra_ip_frame = ttk.Frame(multi_ip_frame)
        self.extra_ip_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # 添加额外IP按钮
        add_ip_btn = ttk.Button(multi_ip_frame, text="添加IP", command=self.add_extra_ip)
        add_ip_btn.grid(row=1, column=0, pady=5)
        
        clear_ip_btn = ttk.Button(multi_ip_frame, text="清空", command=self.clear_extra_ips)
        clear_ip_btn.grid(row=1, column=1, pady=5)
        
        # 操作按钮框架
        button_frame = ttk.Frame(right_frame)
        button_frame.grid(row=2, column=0, pady=10)
        
        # 操作按钮
        self.set_static_btn = ttk.Button(button_frame, text="设置静态IP", 
                                        command=self.set_static_ip)
        self.set_static_btn.pack(side=tk.LEFT, padx=5)
        
        self.set_dhcp_btn = ttk.Button(button_frame, text="设置DHCP", 
                                      command=self.set_dhcp)
        self.set_dhcp_btn.pack(side=tk.LEFT, padx=5)
        
        self.refresh_ip_btn = ttk.Button(button_frame, text="刷新", 
                                        command=self.refresh_ip_info)
        self.refresh_ip_btn.pack(side=tk.LEFT, padx=5)
        
        self.export_btn = ttk.Button(button_frame, text="导出配置", 
                                    command=self.export_config)
        self.export_btn.pack(side=tk.LEFT, padx=5)
        
        # 网卡控制按钮框架
        adapter_control_frame = ttk.Frame(right_frame)
        adapter_control_frame.grid(row=3, column=0, pady=10)
        
        # 网卡控制按钮
        self.disable_btn = ttk.Button(adapter_control_frame, text="禁用网卡", 
                                     command=self.disable_adapter)
        self.disable_btn.pack(side=tk.LEFT, padx=5)
        
        self.enable_btn = ttk.Button(adapter_control_frame, text="启用网卡", 
                                    command=self.enable_adapter)
        self.enable_btn.pack(side=tk.LEFT, padx=5)
        
        self.reset_network_btn = ttk.Button(adapter_control_frame, text="重置网络", 
                                           command=self.reset_network)
        self.reset_network_btn.pack(side=tk.LEFT, padx=5)
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
    def refresh_network_adapters(self):
        """获取网络适配器列表"""
        try:
            self.status_var.set("正在获取网络适配器...")
            self.root.update()
            
            # 使用netsh命令获取网络适配器 - 使用utf-8编码
            result = subprocess.run("netsh interface show interface", 
                                  shell=True, capture_output=True, text=True, 
                                  encoding='utf-8', errors='ignore')
            
            adapters = []
            if result.returncode == 0:
                lines = result.stdout.split('\n')
                for line in lines:
                    # 查找包含"已启用"的行（表示启用的网络适配器）
                    if '已启用' in line:
                        # 提取适配器名称 - 适配器名称在行的最后部分
                        parts = line.split()
                        if len(parts) >= 4:
                            # 从第4个部分开始到最后都是适配器名称
                            adapter_name = ' '.join(parts[3:])
                            # 过滤掉一些虚拟适配器（可选）
                            if not any(skip in adapter_name.lower() for skip in ['loopback', 'isatap', 'teredo']):
                                adapters.append(adapter_name)
            
            self.adapter_combo['values'] = adapters
            
            if adapters:
                self.adapter_combo.set(adapters[0])
                self.on_adapter_selected()
            
            self.status_var.set(f"找到 {len(adapters)} 个网络适配器")
            
        except Exception as e:
            self.status_var.set(f"错误: {str(e)}")
            messagebox.showerror("错误", f"获取网络适配器时出错:\n{str(e)}")
    
    def on_adapter_selected(self, event=None):
        """当选择网络适配器时"""
        self.refresh_ip_info()
    
    def refresh_ip_info(self):
        """刷新当前IP信息"""
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            return
            
        try:
            self.status_var.set("正在获取IP信息...")
            self.root.update()
            
            # 使用ipconfig命令获取IP信息 - 使用cp936编码（Windows中文系统标准编码）
            result = subprocess.run("ipconfig /all", 
                                  shell=True, capture_output=True, text=True, 
                                  encoding='cp936', errors='ignore')
            
            self.ip_info_text.config(state=tk.NORMAL)
            self.ip_info_text.delete(1.0, tk.END)
            
            if result.returncode == 0 and result.stdout.strip():
                # 解析并显示简化的网络信息
                info = self.parse_network_info(result.stdout, adapter_name)
                self.ip_info_text.insert(tk.END, info)
            else:
                self.ip_info_text.insert(tk.END, "无法获取IP信息")
            
            self.ip_info_text.config(state=tk.DISABLED)
            self.status_var.set("IP信息已更新")
            
        except Exception as e:
            self.status_var.set(f"错误: {str(e)}")
            messagebox.showerror("错误", f"获取IP信息时出错:\n{str(e)}")
    
    def parse_network_info(self, ipconfig_output, target_adapter):
        """解析ipconfig输出，提取指定适配器的关键信息"""
        lines = ipconfig_output.split('\n')
        info_lines = []
        
        # 查找目标适配器
        adapter_found = False
        current_adapter = ""
        
        for line in lines:
            line = line.strip()
            
            # 检查是否是适配器标题行
            if line.endswith(':') and not line.startswith('   '):
                current_adapter = line[:-1]  # 移除冒号
                if current_adapter == target_adapter:
                    adapter_found = True
                    info_lines.append(f"适配器名称: {current_adapter}")
                else:
                    adapter_found = False
                continue
            
            if not adapter_found:
                continue
            
            # 提取MAC地址
            if '物理地址' in line or 'MAC' in line:
                mac_match = re.search(r'([0-9A-Fa-f]{2}[:-]){5}[0-9A-Fa-f]{2}', line)
                if mac_match:
                    info_lines.append(f"MAC地址: {mac_match.group()}")
            
            # 提取DHCP状态
            elif 'DHCP 已启用' in line:
                if '是' in line:
                    info_lines.append("DHCP启用: 是")
                elif '否' in line:
                    info_lines.append("DHCP启用: 否")
            
            # 提取IPv4地址
            elif 'IPv4 地址' in line and '(首选)' in line:
                ip_match = re.search(r'(\d+\.\d+\.\d+\.\d+)', line)
                if ip_match:
                    ip = ip_match.group(1)
                    # 查找对应的子网掩码
                    mask = self.find_subnet_mask(lines, lines.index(line))
                    if mask:
                        info_lines.append(f"IP地址配置: {ip}/{mask}")
                    else:
                        info_lines.append(f"IP地址: {ip}")
            
            # 提取IPv6地址
            elif 'IPv6 地址' in line and '(首选)' in line:
                ipv6_match = re.search(r'([0-9a-fA-F:]+::[0-9a-fA-F:]+)', line)
                if ipv6_match:
                    ipv6 = ipv6_match.group(1)
                    # 查找对应的前缀长度
                    prefix = self.find_ipv6_prefix(lines, lines.index(line))
                    if prefix:
                        info_lines.append(f"IPv6地址: {ipv6}/{prefix}")
                    else:
                        info_lines.append(f"IPv6地址: {ipv6}")
            
            # 提取默认网关
            elif '默认网关' in line:
                gateway_match = re.search(r'(\d+\.\d+\.\d+\.\d+)', line)
                if gateway_match:
                    info_lines.append(f"默认网关: {gateway_match.group(1)}")
            
            # 提取DNS服务器
            elif 'DNS 服务器' in line and not 'DHCPv6' in line:
                dns_match = re.search(r'(\d+\.\d+\.\d+\.\d+)', line)
                if dns_match:
                    info_lines.append(f"DNS服务器: {dns_match.group(1)}")
        
        if not info_lines:
            return f"未找到适配器 '{target_adapter}' 的详细信息"
        
        return '\n'.join(info_lines)
    
    def find_subnet_mask(self, lines, current_index):
        """查找子网掩码"""
        for i in range(current_index + 1, min(current_index + 5, len(lines))):
            line = lines[i].strip()
            if '子网掩码' in line:
                mask_match = re.search(r'(\d+\.\d+\.\d+\.\d+)', line)
                if mask_match:
                    return mask_match.group(1)
        return None
    
    def find_ipv6_prefix(self, lines, current_index):
        """查找IPv6前缀长度"""
        for i in range(current_index + 1, min(current_index + 5, len(lines))):
            line = lines[i].strip()
            if '临时 IPv6 地址' in line or 'IPv6 地址' in line:
                prefix_match = re.search(r'%(\d+)', line)
                if prefix_match:
                    return prefix_match.group(1)
        return "64"  # 默认前缀长度
    
    def add_extra_ip(self):
        """添加额外IP地址"""
        if len(self.extra_ips) >= 5:  # 限制最多5个额外IP
            messagebox.showwarning("警告", "最多只能添加5个额外IP地址")
            return
            
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
                           command=lambda: self.remove_extra_ip(ip_frame, ip_var, mask_var))
        del_btn.pack(side=tk.LEFT)
        
        # 保存引用
        self.extra_ips.append({
            'frame': ip_frame,
            'ip_var': ip_var,
            'mask_var': mask_var
        })
        
        # 添加标签
        if len(self.extra_ips) == 1:
            label_frame = ttk.Frame(self.extra_ip_frame)
            label_frame.pack(fill=tk.X, pady=2)
            ttk.Label(label_frame, text="IP地址").pack(side=tk.LEFT, padx=(0, 5))
            ttk.Label(label_frame, text="子网掩码").pack(side=tk.LEFT, padx=(0, 5))
            ttk.Label(label_frame, text="操作").pack(side=tk.LEFT)
        
    def remove_extra_ip(self, frame, ip_var, mask_var):
        """删除额外IP地址"""
        frame.destroy()
        self.extra_ips = [ip for ip in self.extra_ips if ip['ip_var'] != ip_var]
        
    def clear_extra_ips(self):
        """清空所有额外IP地址"""
        for ip_info in self.extra_ips:
            ip_info['frame'].destroy()
        self.extra_ips.clear()
        
    def get_extra_ips(self):
        """获取所有额外IP地址"""
        extra_ips = []
        for ip_info in self.extra_ips:
            ip = ip_info['ip_var'].get().strip()
            mask = ip_info['mask_var'].get().strip()
            if ip and mask and self.is_valid_ip(ip) and self.is_valid_ip(mask):
                extra_ips.append((ip, mask))
        return extra_ips
    
    def set_static_ip(self):
        """设置静态IP"""
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
        
        if not self.is_valid_ip(ip):
            messagebox.showerror("错误", "IP地址格式不正确")
            return
        
        try:
            self.status_var.set("正在设置静态IP...")
            self.root.update()
            
            # 构建netsh命令
            cmd = f'netsh interface ip set address "{adapter_name}" static {ip} {mask}'
            if gateway:
                cmd += f' {gateway}'
            
            result = subprocess.run(cmd, shell=True, capture_output=True, 
                                  text=True, encoding='utf-8', errors='ignore')
            
            if result.returncode == 0:
                # 设置DNS
                if dns:
                    dns_cmd = f'netsh interface ip set dns "{adapter_name}" static {dns}'
                    subprocess.run(dns_cmd, shell=True, capture_output=True, 
                                 text=True, encoding='utf-8', errors='ignore')
                
                messagebox.showinfo("成功", f"静态IP设置成功\nIP: {ip}")
                self.status_var.set("静态IP设置完成")
                self.refresh_ip_info()
            else:
                messagebox.showerror("错误", f"设置静态IP失败:\n{result.stderr}")
                self.status_var.set("设置静态IP失败")
                
        except Exception as e:
            messagebox.showerror("错误", f"设置静态IP时出错:\n{str(e)}")
            self.status_var.set(f"错误: {str(e)}")
    
    def set_dhcp(self):
        """设置DHCP"""
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            messagebox.showwarning("警告", "请先选择网络适配器")
            return
        
        try:
            self.status_var.set("正在设置DHCP...")
            self.root.update()
            
            # 设置DHCP
            cmd = f'netsh interface ip set address "{adapter_name}" dhcp'
            result = subprocess.run(cmd, shell=True, capture_output=True, 
                                  text=True, encoding='utf-8', errors='ignore')
            
            if result.returncode == 0:
                # 设置DNS为DHCP
                dns_cmd = f'netsh interface ip set dns "{adapter_name}" dhcp'
                subprocess.run(dns_cmd, shell=True, capture_output=True, 
                             text=True, encoding='utf-8', errors='ignore')
                
                messagebox.showinfo("成功", "DHCP设置成功")
                self.status_var.set("DHCP设置完成")
                self.refresh_ip_info()
            else:
                messagebox.showerror("错误", f"设置DHCP失败:\n{result.stderr}")
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
        """验证IPv4输入格式"""
        if len(value) > 15:
            return False
        if value == "":
            return True
        pattern = r'^\d{0,3}(\.\d{0,3}){0,3}$'
        return re.match(pattern, value) is not None

    def is_valid_ip(self, ip):
        """严格校验IPv4"""
        pattern = r'^(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)){3}$'
        return re.match(pattern, ip) is not None

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
            
            cmd = f'netsh interface set interface "{adapter_name}" admin=disable'
            result = subprocess.run(cmd, shell=True, capture_output=True, 
                                  text=True, encoding='utf-8', errors='ignore')
            
            if result.returncode == 0:
                messagebox.showinfo("成功", f"网卡 '{adapter_name}' 已禁用")
                self.status_var.set("网卡禁用完成")
                self.refresh_network_adapters()
            else:
                messagebox.showerror("错误", f"禁用网卡失败:\n{result.stderr}")
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
            
            cmd = f'netsh interface set interface "{adapter_name}" admin=enable'
            result = subprocess.run(cmd, shell=True, capture_output=True, 
                                  text=True, encoding='utf-8', errors='ignore')
            
            if result.returncode == 0:
                messagebox.showinfo("成功", f"网卡 '{adapter_name}' 已启用")
                self.status_var.set("网卡启用完成")
                self.refresh_network_adapters()
            else:
                messagebox.showerror("错误", f"启用网卡失败:\n{result.stderr}")
                self.status_var.set("启用网卡失败")
                
        except Exception as e:
            messagebox.showerror("错误", f"启用网卡时出错:\n{str(e)}")
            self.status_var.set(f"错误: {str(e)}")

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
                    
                    result = subprocess.run(cmd, shell=True, capture_output=True, 
                                          text=True, encoding='utf-8', errors='ignore',
                                          creationflags=subprocess.CREATE_NO_WINDOW)
                    
                    # 检查命令是否成功执行
                    if result.returncode == 0 or ("完成" in result.stdout and "重置" in cmd):
                        success_count += 1
                        self.status_var.set(f"命令执行成功: {cmd}")
                    else:
                        messagebox.showwarning("警告", 
                                             f"命令执行可能有问题:\n{cmd}\n\n"
                                             f"返回码: {result.returncode}\n"
                                             f"错误输出: {result.stderr}\n\n"
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
    
    # 检查管理员权限
    if not is_admin():
        messagebox.showwarning("权限警告", 
                             "程序需要管理员权限才能执行某些操作。\n\n"
                             "建议以管理员身份运行此程序，以确保所有功能正常工作。\n\n"
                             "某些功能（如网络重置、网卡控制）可能需要管理员权限。")
    
    root = tk.Tk()
    app = IPManager(root)
    
    # 设置窗口图标（如果有的话）
    try:
        root.iconbitmap('icon.ico')
    except:
        pass
    
    root.mainloop()

if __name__ == "__main__":
    main() 