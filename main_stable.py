import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import re
import socket
import threading
import os
import sys
from datetime import datetime
import time

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
        self.wmi = None
        self.adapter_map = {}
        self.wmi_adapters = {}
        
        # 创建UI
        self.setup_ui()
        
        # 在后台线程中初始化WMI
        self.init_wmi_async()
        
    def init_wmi_async(self):
        """在后台线程中初始化WMI"""
        def init_wmi():
            try:
                import wmi
                self.wmi = wmi.WMI()
                # 在主线程中更新状态
                self.root.after(0, lambda: self.status_var.set("WMI初始化成功"))
                self.root.after(0, self.refresh_network_adapters)
            except Exception as e:
                # 在主线程中显示错误
                self.root.after(0, lambda: self.status_var.set("WMI初始化失败，部分功能不可用"))
                self.root.after(0, lambda: messagebox.showerror("警告", 
                    f"WMI初始化失败: {str(e)}\n\n部分功能可能不可用，但程序仍可运行。"))
        
        # 启动后台线程
        thread = threading.Thread(target=init_wmi, daemon=True)
        thread.start()
        
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
        self.extra_ips = []
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
        self.status_var.set("正在初始化...")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
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
        
    def refresh_network_adapters(self):
        """使用WMI获取网络适配器列表"""
        if not self.wmi:
            messagebox.showerror("错误", "WMI未初始化，无法获取网络适配器信息")
            return
            
        # 在后台线程中执行
        def get_adapters():
            try:
                self.root.after(0, lambda: self.status_var.set("正在获取网络适配器..."))
                
                adapters = []
                self.adapter_map = {}
                self.wmi_adapters = {}
                
                # 获取启用的网络适配器配置
                nic_configs = self.wmi.Win32_NetworkAdapterConfiguration(IPEnabled=True)
                
                for nic in nic_configs:
                    if nic.IPAddress and len(nic.IPAddress) > 0:
                        # 获取适配器名称
                        adapter_name = nic.Description
                        if adapter_name:
                            adapters.append(adapter_name)
                            self.wmi_adapters[adapter_name] = nic
                            # 同时保存到映射中，用于兼容性
                            self.adapter_map[adapter_name] = adapter_name
                
                # 在主线程中更新UI
                self.root.after(0, lambda: self.update_adapter_list(adapters))
                
            except Exception as e:
                self.root.after(0, lambda: self.status_var.set(f"错误: {str(e)}"))
                self.root.after(0, lambda: messagebox.showerror("错误", f"获取网络适配器时出错:\n{str(e)}"))
        
        thread = threading.Thread(target=get_adapters, daemon=True)
        thread.start()
    
    def update_adapter_list(self, adapters):
        """更新适配器列表（在主线程中调用）"""
        self.adapter_combo['values'] = adapters
        
        if adapters:
            self.adapter_combo.set(adapters[0])
            self.on_adapter_selected()
        
        self.status_var.set(f"找到 {len(adapters)} 个网络适配器")
    
    def on_adapter_selected(self, event=None):
        """当选择网络适配器时"""
        self.refresh_ip_info()
    
    def refresh_ip_info(self):
        """使用WMI刷新当前IP信息"""
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            return
            
        # 在后台线程中执行
        def get_ip_info():
            try:
                self.root.after(0, lambda: self.status_var.set("正在获取IP信息..."))
                
                nic = self.wmi_adapters.get(adapter_name)
                if nic:
                    self.root.after(0, lambda: self.display_wmi_ip_info(nic))
                    self.root.after(0, lambda: self.status_var.set("IP信息已更新"))
                else:
                    self.root.after(0, lambda: self.status_var.set("未找到适配器信息"))
                    
            except Exception as e:
                self.root.after(0, lambda: self.status_var.set(f"错误: {str(e)}"))
                self.root.after(0, lambda: messagebox.showerror("错误", f"获取IP信息时出错:\n{str(e)}"))
        
        thread = threading.Thread(target=get_ip_info, daemon=True)
        thread.start()
    
    def display_wmi_ip_info(self, nic):
        """显示WMI获取的IP信息"""
        self.ip_info_text.config(state=tk.NORMAL)
        self.ip_info_text.delete(1.0, tk.END)
        
        # 显示适配器基本信息
        info_lines = []
        info_lines.append(f"适配器名称: {nic.Description}")
        info_lines.append(f"MAC地址: {nic.MACAddress}")
        info_lines.append(f"DHCP启用: {'是' if nic.DHCPEnabled else '否'}")
        info_lines.append("")
        
        # 显示IP地址信息
        if nic.IPAddress:
            info_lines.append("IP地址配置:")
            for i, ip in enumerate(nic.IPAddress):
                if i < len(nic.IPSubnet):
                    info_lines.append(f"  {ip}/{nic.IPSubnet[i]}")
                else:
                    info_lines.append(f"  {ip}")
        
        # 显示网关信息
        if nic.DefaultIPGateway:
            info_lines.append("")
            info_lines.append("默认网关:")
            for gateway in nic.DefaultIPGateway:
                info_lines.append(f"  {gateway}")
        
        # 显示DNS信息
        if nic.DNSServerSearchOrder:
            info_lines.append("")
            info_lines.append("DNS服务器:")
            for dns in nic.DNSServerSearchOrder:
                info_lines.append(f"  {dns}")
        
        # 显示到文本框
        for line in info_lines:
            self.ip_info_text.insert(tk.END, line + '\n')
        
        # 提取当前配置到输入框
        self.extract_wmi_config(nic)
        
        self.ip_info_text.config(state=tk.DISABLED)
    
    def extract_wmi_config(self, nic):
        """从WMI配置中提取当前设置"""
        # 提取第一个IPv4地址
        if nic.IPAddress:
            for i, ip in enumerate(nic.IPAddress):
                if self.is_valid_ip(ip):
                    self.ip_var.set(ip)
                    if i < len(nic.IPSubnet):
                        self.mask_var.set(nic.IPSubnet[i])
                    break
        
        # 提取网关
        if nic.DefaultIPGateway:
            for gateway in nic.DefaultIPGateway:
                if self.is_valid_ip(gateway):
                    self.gateway_var.set(gateway)
                    break
        
        # 提取DNS
        if nic.DNSServerSearchOrder:
            for dns in nic.DNSServerSearchOrder:
                if self.is_valid_ip(dns):
                    self.dns_var.set(dns)
                    break
    
    def set_static_ip(self):
        """使用WMI设置静态IP"""
        if not self.wmi:
            messagebox.showerror("错误", "WMI未初始化，无法设置IP")
            return
            
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
        
        # 在后台线程中执行
        def set_ip():
            try:
                self.root.after(0, lambda: self.status_var.set("正在设置静态IP..."))
                
                nic = self.wmi_adapters.get(adapter_name)
                if not nic:
                    self.root.after(0, lambda: messagebox.showerror("错误", "未找到适配器配置"))
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
                    
                    self.root.after(0, lambda: messagebox.showinfo("成功", f"静态IP设置成功\n主IP: {ip}\n额外IP: {len(extra_ips)}个"))
                    self.root.after(0, lambda: self.status_var.set("静态IP设置完成"))
                    self.root.after(0, self.refresh_ip_info)
                else:
                    error_msg = f"设置失败，错误代码: {result[0]}"
                    self.root.after(0, lambda: messagebox.showerror("错误", f"设置静态IP失败:\n{error_msg}"))
                    self.root.after(0, lambda: self.status_var.set("设置静态IP失败"))
                    
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("错误", f"设置静态IP时出错:\n{str(e)}"))
                self.root.after(0, lambda: self.status_var.set(f"错误: {str(e)}"))
        
        thread = threading.Thread(target=set_ip, daemon=True)
        thread.start()
    
    def set_dhcp(self):
        """使用WMI设置DHCP"""
        if not self.wmi:
            messagebox.showerror("错误", "WMI未初始化，无法设置DHCP")
            return
            
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            messagebox.showwarning("警告", "请先选择网络适配器")
            return
        
        # 在后台线程中执行
        def set_dhcp():
            try:
                self.root.after(0, lambda: self.status_var.set("正在设置DHCP..."))
                
                nic = self.wmi_adapters.get(adapter_name)
                if not nic:
                    self.root.after(0, lambda: messagebox.showerror("错误", "未找到适配器配置"))
                    return
                
                # 启用DHCP
                result = nic.EnableDHCP()
                if result[0] == 0:
                    # 设置DNS为DHCP
                    nic.SetDNSServerSearchOrder(DNSServerSearchOrder=[])
                    
                    self.root.after(0, lambda: messagebox.showinfo("成功", "DHCP设置成功"))
                    self.root.after(0, lambda: self.status_var.set("DHCP设置完成"))
                    self.root.after(0, self.refresh_ip_info)
                else:
                    error_msg = f"设置失败，错误代码: {result[0]}"
                    self.root.after(0, lambda: messagebox.showerror("错误", f"设置DHCP失败:\n{error_msg}"))
                    self.root.after(0, lambda: self.status_var.set("设置DHCP失败"))
                    
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("错误", f"设置DHCP时出错:\n{str(e)}"))
                self.root.after(0, lambda: self.status_var.set(f"错误: {str(e)}"))
        
        thread = threading.Thread(target=set_dhcp, daemon=True)
        thread.start()
    
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
        if not self.wmi:
            messagebox.showerror("错误", "WMI未初始化，无法禁用网卡")
            return
            
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            messagebox.showwarning("警告", "请先选择网络适配器")
            return
        
        # 确认对话框
        result = messagebox.askyesno("确认", f"确定要禁用网卡 '{adapter_name}' 吗？\n\n禁用后网络连接将中断。")
        if not result:
            return
        
        # 在后台线程中执行
        def disable():
            try:
                self.root.after(0, lambda: self.status_var.set("正在禁用网卡..."))
                
                # 使用WMI禁用网卡 - 需要找到对应的网络适配器对象
                nic_config = self.wmi_adapters.get(adapter_name)
                if not nic_config:
                    self.root.after(0, lambda: messagebox.showerror("错误", "未找到适配器配置"))
                    return
                
                # 通过MAC地址找到对应的网络适配器对象
                mac_address = nic_config.MACAddress
                if not mac_address:
                    self.root.after(0, lambda: messagebox.showerror("错误", "无法获取网卡MAC地址"))
                    return
                
                # 查找对应的网络适配器
                adapters = self.wmi.Win32_NetworkAdapter(MACAddress=mac_address)
                if not adapters:
                    self.root.after(0, lambda: messagebox.showerror("错误", "未找到对应的网络适配器"))
                    return
                
                adapter = adapters[0]
                
                # 禁用网卡
                result = adapter.Disable()
                if result[0] == 0:
                    self.root.after(0, lambda: messagebox.showinfo("成功", f"网卡 '{adapter_name}' 已禁用"))
                    self.root.after(0, lambda: self.status_var.set("网卡禁用完成"))
                    self.root.after(0, self.refresh_network_adapters)  # 刷新适配器列表
                else:
                    error_msg = f"禁用失败，错误代码: {result[0]}"
                    self.root.after(0, lambda: messagebox.showerror("错误", f"禁用网卡失败:\n{error_msg}"))
                    self.root.after(0, lambda: self.status_var.set("禁用网卡失败"))
                    
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("错误", f"禁用网卡时出错:\n{str(e)}"))
                self.root.after(0, lambda: self.status_var.set(f"错误: {str(e)}"))
        
        thread = threading.Thread(target=disable, daemon=True)
        thread.start()

    def enable_adapter(self):
        """启用选中的网卡"""
        if not self.wmi:
            messagebox.showerror("错误", "WMI未初始化，无法启用网卡")
            return
            
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            messagebox.showwarning("警告", "请先选择网络适配器")
            return
        
        # 确认对话框
        result = messagebox.askyesno("确认", f"确定要启用网卡 '{adapter_name}' 吗？")
        if not result:
            return
        
        # 在后台线程中执行
        def enable():
            try:
                self.root.after(0, lambda: self.status_var.set("正在启用网卡..."))
                
                # 使用WMI启用网卡 - 需要找到对应的网络适配器对象
                nic_config = self.wmi_adapters.get(adapter_name)
                if not nic_config:
                    self.root.after(0, lambda: messagebox.showerror("错误", "未找到适配器配置"))
                    return
                
                # 通过MAC地址找到对应的网络适配器对象
                mac_address = nic_config.MACAddress
                if not mac_address:
                    self.root.after(0, lambda: messagebox.showerror("错误", "无法获取网卡MAC地址"))
                    return
                
                # 查找对应的网络适配器
                adapters = self.wmi.Win32_NetworkAdapter(MACAddress=mac_address)
                if not adapters:
                    self.root.after(0, lambda: messagebox.showerror("错误", "未找到对应的网络适配器"))
                    return
                
                adapter = adapters[0]
                
                # 启用网卡
                result = adapter.Enable()
                if result[0] == 0:
                    self.root.after(0, lambda: messagebox.showinfo("成功", f"网卡 '{adapter_name}' 已启用"))
                    self.root.after(0, lambda: self.status_var.set("网卡启用完成"))
                    self.root.after(0, self.refresh_network_adapters)  # 刷新适配器列表
                else:
                    error_msg = f"启用失败，错误代码: {result[0]}"
                    self.root.after(0, lambda: messagebox.showerror("错误", f"启用网卡失败:\n{error_msg}"))
                    self.root.after(0, lambda: self.status_var.set("启用网卡失败"))
                    
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("错误", f"启用网卡时出错:\n{str(e)}"))
                self.root.after(0, lambda: self.status_var.set(f"错误: {str(e)}"))
        
        thread = threading.Thread(target=enable, daemon=True)
        thread.start()

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
        
        # 在后台线程中执行
        def reset():
            try:
                self.root.after(0, lambda: self.status_var.set("正在重置网络设置..."))
                
                # 执行网络重置命令
                commands = [
                    "netsh int ip reset",
                    "netsh winsock reset"
                ]
                
                success_count = 0
                for cmd in commands:
                    try:
                        self.root.after(0, lambda cmd=cmd: self.status_var.set(f"正在执行: {cmd}"))
                        
                        # 方法1: 直接执行命令
                        try:
                            result = subprocess.run(cmd, shell=True, capture_output=True, 
                                                  text=True, encoding='gbk', errors='ignore',
                                                  creationflags=subprocess.CREATE_NO_WINDOW)
                            
                            # 检查命令是否成功执行
                            # 对于 netsh int ip reset，即使返回码是1，如果输出包含"完成"字样，也认为是成功的
                            if result.returncode == 0 or ("完成" in result.stdout and "重置" in cmd):
                                success_count += 1
                                self.root.after(0, lambda cmd=cmd: self.status_var.set(f"命令执行成功: {cmd}"))
                            else:
                                # 方法2: 尝试以管理员权限运行
                                try:
                                    # 使用runas命令
                                    admin_cmd = f'runas /user:Administrator "{cmd}"'
                                    result = subprocess.run(admin_cmd, shell=True, capture_output=True,
                                                          text=True, encoding='gbk', errors='ignore')
                                    
                                    if result.returncode == 0:
                                        success_count += 1
                                        self.root.after(0, lambda cmd=cmd: self.status_var.set(f"命令执行成功: {cmd}"))
                                    else:
                                        self.root.after(0, lambda cmd=cmd, result=result: messagebox.showwarning("警告", 
                                                             f"命令执行可能有问题:\n{cmd}\n\n"
                                                             f"返回码: {result.returncode}\n"
                                                             f"错误输出: {result.stderr}\n\n"
                                                             f"请确保以管理员身份运行程序。"))
                                except Exception as e2:
                                    self.root.after(0, lambda cmd=cmd, e2=e2: messagebox.showwarning("警告", 
                                                 f"执行命令时出错:\n{cmd}\n\n"
                                                 f"错误: {str(e2)}\n\n"
                                                 f"请确保以管理员身份运行程序。"))
                                    
                        except Exception as e1:
                            self.root.after(0, lambda cmd=cmd, e1=e1: messagebox.showwarning("警告", 
                                         f"执行命令时出错:\n{cmd}\n\n"
                                         f"错误: {str(e1)}\n\n"
                                         f"请确保以管理员身份运行程序。"))
                            
                    except Exception as e:
                        self.root.after(0, lambda cmd=cmd, e=e: messagebox.showwarning("警告", 
                                     f"执行命令时出错:\n{cmd}\n\n"
                                     f"错误: {str(e)}\n\n"
                                     f"请确保以管理员身份运行程序。"))
                
                if success_count > 0:
                    # 询问是否重启
                    def ask_restart():
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
                    
                    self.root.after(0, ask_restart)
                else:
                    self.root.after(0, lambda: messagebox.showerror("错误", 
                                    "所有网络重置命令都执行失败。\n\n"
                                    "可能的原因：\n"
                                    "1. 没有管理员权限\n"
                                    "2. 命令被系统阻止\n"
                                    "3. 网络服务正在运行\n\n"
                                    "建议：\n"
                                    "1. 以管理员身份运行程序\n"
                                    "2. 关闭所有网络相关程序\n"
                                    "3. 手动在命令提示符中执行这些命令"))
                
                self.root.after(0, lambda: self.status_var.set("网络重置完成"))
                
            except Exception as e:
                self.root.after(0, lambda e=e: messagebox.showerror("错误", f"重置网络时出错:\n{str(e)}"))
                self.root.after(0, lambda e=e: self.status_var.set(f"错误: {str(e)}"))
        
        thread = threading.Thread(target=reset, daemon=True)
        thread.start()

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