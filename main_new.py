import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import re
import socket
import threading
import os
import sys
from datetime import datetime
import wmi

class IPManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows IP地址管理器")
        self.root.geometry("1000x700")
        self.root.resizable(True, True)
        
        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')
        
        # 初始化WMI
        try:
            self.wmi = wmi.WMI()
        except Exception as e:
            messagebox.showerror("错误", f"WMI初始化失败: {str(e)}")
            return
        
        self.setup_ui()
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
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 适配器映射
        self.adapter_map = {}
        self.wmi_adapters = {}
        
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
        try:
            self.status_var.set("正在获取网络适配器...")
            self.root.update()
            
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
        """使用WMI刷新当前IP信息"""
        adapter_name = self.adapter_var.get()
        if not adapter_name:
            return
            
        try:
            self.status_var.set("正在获取IP信息...")
            self.root.update()
            
            nic = self.wmi_adapters.get(adapter_name)
            if nic:
                self.display_wmi_ip_info(nic)
                self.status_var.set("IP信息已更新")
            else:
                self.status_var.set("未找到适配器信息")
                
        except Exception as e:
            self.status_var.set(f"错误: {str(e)}")
            messagebox.showerror("错误", f"获取IP信息时出错:\n{str(e)}")
    
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

def main():
    # 检查管理员权限
    if not os.name == 'nt':
        messagebox.showerror("错误", "此程序仅支持Windows系统")
        return
    
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