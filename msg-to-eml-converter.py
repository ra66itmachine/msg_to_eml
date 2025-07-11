#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header, decode_header
from email.utils import formatdate, parsedate_to_datetime, formataddr
import threading
import mimetypes
import datetime
import re
import base64
import quopri
import chardet
import codecs
import uuid
import subprocess
import platform

# 安装命令: pip install extract-msg chardet
try:
    import extract_msg
    EXTRACT_MSG_AVAILABLE = True
except ImportError:
    EXTRACT_MSG_AVAILABLE = False

class EnhancedMSGToEMLConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("MSG转EML转换器")
        # 设置窗口大小并居中
        window_width = 1100
        window_height = 800
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.update_idletasks()
        self.root.resizable(True, True)
        
        # 设置应用图标（如果有）
        try:
            self.root.iconbitmap('converter.ico')
        except:
            pass
        
        # 存储选择的文件和转换结果
        self.file_items = {}  # 存储文件信息和tree item id的映射
        self.conversion_results = {}
        self.output_dir = None
        
        # 转换选项
        self.include_attachments = tk.BooleanVar(value=True)
        self.preserve_headers = tk.BooleanVar(value=True)
        self.auto_decode = tk.BooleanVar(value=True)
        self.detect_encoding = tk.BooleanVar(value=True)
        self.preserve_transport_headers = tk.BooleanVar(value=True)
        self.show_ip_info = tk.BooleanVar(value=True)
        
        self.setup_ui()
        
        # 检查依赖
        if not EXTRACT_MSG_AVAILABLE:
            messagebox.showwarning(
                "需要安装依赖库",
                "请先安装 extract-msg 和 chardet 库：\n\n"
                "打开命令行运行：\n"
                "pip install extract-msg chardet\n\n"
                "安装完成后重新运行程序。"
            )
    
    def create_tooltip(self, widget, text):
        """为控件创建工具提示"""
        def on_enter(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            
            label = ttk.Label(tooltip, text=text, 
                            background="#FFFFDD", 
                            relief=tk.SOLID, 
                            borderwidth=1,
                            font=("Arial", 9))
            label.pack()
            
            widget.tooltip = tooltip
        
        def on_leave(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                del widget.tooltip
        
        widget.bind('<Enter>', on_enter)
        widget.bind('<Leave>', on_leave)
    
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="MSG转EML转换器", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # 顶部控制区域
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        control_frame.columnconfigure(1, weight=1)
        
        # 第一行：按钮和输出目录
        button_row = ttk.Frame(control_frame)
        button_row.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 按钮框架
        button_frame = ttk.Frame(button_row)
        button_frame.pack(side=tk.LEFT)
        
        # 选择文件按钮
        self.select_btn = ttk.Button(button_frame, text="添加MSG文件", 
                                    command=self.select_files)
        self.select_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 清空列表按钮
        self.clear_btn = ttk.Button(button_frame, text="清空列表", 
                                   command=self.clear_files, state=tk.DISABLED)
        self.clear_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 删除选中按钮
        self.remove_btn = ttk.Button(button_frame, text="删除选中", 
                                    command=self.remove_selected, state=tk.DISABLED)
        self.remove_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 转换按钮
        self.convert_btn = ttk.Button(button_frame, text="开始转换", 
                                     command=self.start_conversion, 
                                     state=tk.DISABLED)
        self.convert_btn.pack(side=tk.LEFT, padx=(20, 0))
        
        # 输出目录框架
        output_frame = ttk.Frame(button_row)
        output_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(20, 0))
        
        ttk.Label(output_frame, text="输出目录:").pack(side=tk.LEFT, padx=(0, 5))
        self.output_dir_var = tk.StringVar(value="与源文件相同目录")
        self.output_dir_label = ttk.Label(output_frame, textvariable=self.output_dir_var, 
                                         relief=tk.SUNKEN, padding="5", width=40)
        self.output_dir_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        # 选择输出目录按钮
        self.select_output_btn = ttk.Button(output_frame, text="选择目录", 
                                           command=self.select_output_dir)
        self.select_output_btn.pack(side=tk.LEFT)
        
        # 转换选项区域（重新排列）
        options_frame = ttk.LabelFrame(control_frame, text="转换选项", padding="10")
        options_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 使用网格布局，分成两行三列
        # 第一行：核心功能选项
        core_options_frame = ttk.Frame(options_frame)
        core_options_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(core_options_frame, text="核心功能：", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=(0, 10))
        
        # 保留传输头复选框
        self.preserve_transport_cb = ttk.Checkbutton(
            core_options_frame, 
            text="保留完整传输路径",
            variable=self.preserve_transport_headers
        )
        self.preserve_transport_cb.pack(side=tk.LEFT, padx=(0, 15))
        self.create_tooltip(self.preserve_transport_cb, 
                          "保留原始邮件头信息，包括：\n"
                          "• Received头（包含服务器IP地址）\n"
                          "• X-Mailer（邮件客户端信息）\n"
                          "• Authentication-Results等安全信息")
        
        # 保留MSG属性复选框
        self.preserve_headers_cb = ttk.Checkbutton(
            core_options_frame, 
            text="保留MSG扩展属性",
            variable=self.preserve_headers
        )
        self.preserve_headers_cb.pack(side=tk.LEFT, padx=(0, 15))
        self.create_tooltip(self.preserve_headers_cb,
                          "保留MSG文件特有的扩展属性：\n"
                          "• Thread-Topic（会话主题）\n"
                          "• X-Message-Class（消息类别）\n"
                          "• 读取回执请求等Outlook特有信息")
        
        # 显示IP信息复选框
        self.show_ip_cb = ttk.Checkbutton(
            core_options_frame, 
            text="增强IP信息显示",
            variable=self.show_ip_info
        )
        self.show_ip_cb.pack(side=tk.LEFT, padx=(0, 15))
        self.create_tooltip(self.show_ip_cb,
                          "增强显示网络传输信息：\n"
                          "• 发送和接收的SMTP地址\n"
                          "• 详细的时间戳信息\n"
                          "• 有助于追踪邮件传输路径")
        
        # 第二行：辅助功能选项
        aux_options_frame = ttk.Frame(options_frame)
        aux_options_frame.pack(fill=tk.X)
        
        ttk.Label(aux_options_frame, text="辅助功能：", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=(0, 10))
        
        # 包含附件内容复选框
        self.include_attachments_cb = ttk.Checkbutton(
            aux_options_frame, 
            text="包含附件内容",
            variable=self.include_attachments
        )
        self.include_attachments_cb.pack(side=tk.LEFT, padx=(0, 15))
        self.create_tooltip(self.include_attachments_cb,
                          "控制附件处理方式：\n"
                          "• 勾选：完整提取附件内容到EML文件\n"
                          "• 不勾选：只创建附件占位符，减小文件大小")
        
        # 智能编码检测复选框
        self.detect_encoding_cb = ttk.Checkbutton(
            aux_options_frame, 
            text="智能编码检测",
            variable=self.detect_encoding
        )
        self.detect_encoding_cb.pack(side=tk.LEFT, padx=(0, 15))
        self.create_tooltip(self.detect_encoding_cb,
                          "使用chardet库智能检测文本编码：\n"
                          "• 自动识别UTF-8、GBK、GB2312等编码\n"
                          "• 避免中文和其他语言出现乱码")
        
        # 自动解码复选框
        self.auto_decode_cb = ttk.Checkbutton(
            aux_options_frame, 
            text="自动解码编码内容",
            variable=self.auto_decode
        )
        self.auto_decode_cb.pack(side=tk.LEFT)
        self.create_tooltip(self.auto_decode_cb,
                          "自动解码邮件中的编码内容：\n"
                          "• Base64编码（如：5Lit6K+t → 中文）\n"
                          "• Quoted-Printable编码\n"
                          "• RFC 2047编码的邮件头")
        
        # 文件列表区域（使用Treeview）
        list_frame = ttk.LabelFrame(main_frame, text="文件列表", padding="10")
        list_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # 创建Treeview
        columns = ('status', 'result')
        self.file_tree = ttk.Treeview(list_frame, columns=columns, show='tree headings', height=15)
        self.file_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 设置列标题和宽度
        self.file_tree.heading('#0', text='文件名称')
        self.file_tree.heading('status', text='转换情况')
        self.file_tree.heading('result', text='转换结果')
        
        self.file_tree.column('#0', width=400)
        self.file_tree.column('status', width=100, anchor='center')
        self.file_tree.column('result', width=400)
        
        # 添加滚动条
        tree_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        tree_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.file_tree.configure(yscrollcommand=tree_scrollbar.set)
        
        # 绑定右键菜单
        self.file_tree.bind('<Button-3>', self.show_context_menu)
        
        # 创建右键菜单
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="打开文件", command=self.open_file)
        self.context_menu.add_command(label="打开文件所在文件夹", command=self.open_file_location)
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 底部操作按钮
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.grid(row=4, column=0, pady=(0, 10))
        
        # 查看邮件头按钮
        self.view_headers_btn = ttk.Button(bottom_frame, text="查看邮件头详情", 
                                          command=self.view_email_headers)
        self.view_headers_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 查看MSG属性按钮
        self.view_msg_attrs_btn = ttk.Button(bottom_frame, text="调试：查看MSG属性", 
                                           command=self.view_msg_attributes)
        self.view_msg_attrs_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 测试选项按钮
        self.test_options_btn = ttk.Button(bottom_frame, text="测试：比较选项效果", 
                                          command=self.test_option_effects)
        self.test_options_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 状态栏
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=5, column=0, sticky=(tk.W, tk.E))
        status_frame.columnconfigure(0, weight=1)
        
        self.status_label = ttk.Label(status_frame, text="准备就绪", relief=tk.SUNKEN)
        self.status_label.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # 文件计数标签
        self.file_count_label = ttk.Label(status_frame, text="未选择文件", font=("Arial", 9))
        self.file_count_label.grid(row=0, column=1, padx=(10, 0))
        
        # 绑定选项变化事件，用于调试
        for var in [self.preserve_headers, self.preserve_transport_headers, self.show_ip_info]:
            var.trace('w', self.on_option_changed)
    
    def on_option_changed(self, *args):
        """选项变化时的回调（用于调试）"""
        print(f"选项状态 - 保留MSG属性: {self.preserve_headers.get()}, "
              f"保留传输头: {self.preserve_transport_headers.get()}, "
              f"显示IP信息: {self.show_ip_info.get()}")
    
    def show_context_menu(self, event):
        """显示右键菜单"""
        # 获取点击的项目
        item = self.file_tree.identify('item', event.x, event.y)
        if item:
            self.file_tree.selection_set(item)
            # 检查是否有转换结果
            if item in self.file_items:
                file_info = self.file_items[item]
                if file_info.get('output_file') and os.path.exists(file_info['output_file']):
                    self.context_menu.post(event.x_root, event.y_root)
    
    def open_file(self):
        """打开转换后的文件"""
        selection = self.file_tree.selection()
        if selection:
            item = selection[0]
            if item in self.file_items:
                file_info = self.file_items[item]
                output_file = file_info.get('output_file')
                if output_file and os.path.exists(output_file):
                    try:
                        if platform.system() == 'Windows':
                            os.startfile(output_file)
                        elif platform.system() == 'Darwin':  # macOS
                            subprocess.run(['open', output_file])
                        else:  # Linux
                            subprocess.run(['xdg-open', output_file])
                    except Exception as e:
                        messagebox.showerror("错误", f"无法打开文件: {str(e)}")
    
    def open_file_location(self):
        """打开文件所在文件夹"""
        selection = self.file_tree.selection()
        if selection:
            item = selection[0]
            if item in self.file_items:
                file_info = self.file_items[item]
                output_file = file_info.get('output_file')
                if output_file and os.path.exists(output_file):
                    folder = os.path.dirname(output_file)
                    try:
                        if platform.system() == 'Windows':
                            os.startfile(folder)
                        elif platform.system() == 'Darwin':  # macOS
                            subprocess.run(['open', folder])
                        else:  # Linux
                            subprocess.run(['xdg-open', folder])
                    except Exception as e:
                        messagebox.showerror("错误", f"无法打开文件夹: {str(e)}")
    
    def select_files(self):
        """选择MSG文件"""
        files = filedialog.askopenfilenames(
            title="选择MSG文件",
            filetypes=[("MSG files", "*.msg"), ("All files", "*.*")]
        )
        
        if files:
            new_files_count = 0
            for file_path in files:
                # 检查是否已经添加
                already_exists = False
                for item_id, file_info in self.file_items.items():
                    if file_info['path'] == file_path:
                        already_exists = True
                        break
                
                if not already_exists:
                    # 添加到树形视图
                    filename = os.path.basename(file_path)
                    item = self.file_tree.insert('', 'end', text=filename, values=('待转换', ''))
                    
                    # 保存文件信息
                    self.file_items[item] = {
                        'path': file_path,
                        'filename': filename,
                        'status': 'pending',
                        'output_file': None
                    }
                    new_files_count += 1
            
            self.update_file_count()
            
            if new_files_count > 0:
                self.status_label.config(text=f"添加了 {new_files_count} 个新文件")
    
    def clear_files(self):
        """清空文件列表"""
        if messagebox.askyesno("确认", "确定要清空所有已选择的文件吗？"):
            self.file_tree.delete(*self.file_tree.get_children())
            self.file_items.clear()
            self.conversion_results.clear()
            self.update_file_count()
            self.status_label.config(text="已清空文件列表")
    
    def remove_selected(self):
        """删除选中的文件"""
        selection = self.file_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要删除的文件")
            return
        
        for item in selection:
            if item in self.file_items:
                del self.file_items[item]
            self.file_tree.delete(item)
        
        self.update_file_count()
        self.status_label.config(text=f"已删除 {len(selection)} 个文件")
    
    def update_file_count(self):
        """更新文件计数和按钮状态"""
        count = len(self.file_items)
        if count == 0:
            self.file_count_label.config(text="未选择文件")
            self.convert_btn.config(state=tk.DISABLED)
            self.clear_btn.config(state=tk.DISABLED)
            self.remove_btn.config(state=tk.DISABLED)
        else:
            self.file_count_label.config(text=f"已选择 {count} 个文件")
            if EXTRACT_MSG_AVAILABLE:
                self.convert_btn.config(state=tk.NORMAL)
            self.clear_btn.config(state=tk.NORMAL)
            self.remove_btn.config(state=tk.NORMAL)
    
    def select_output_dir(self):
        """选择输出目录"""
        directory = filedialog.askdirectory(title="选择EML文件输出目录")
        if directory:
            self.output_dir = directory
            self.output_dir_var.set(directory)
    
    def start_conversion(self):
        """开始转换文件"""
        if not self.file_items:
            messagebox.showwarning("警告", "请先选择MSG文件")
            return
            
        if not EXTRACT_MSG_AVAILABLE:
            messagebox.showerror("错误", "请先安装 extract-msg 库")
            return
        
        # 重置所有文件状态
        for item_id, file_info in self.file_items.items():
            self.file_tree.set(item_id, 'status', '待转换')
            self.file_tree.set(item_id, 'result', '')
            file_info['status'] = 'pending'
        
        self.conversion_results.clear()
        
        self.convert_btn.config(state=tk.DISABLED)
        self.select_btn.config(state=tk.DISABLED)
        self.clear_btn.config(state=tk.DISABLED)
        self.remove_btn.config(state=tk.DISABLED)
        
        thread = threading.Thread(target=self.convert_files)
        thread.daemon = True
        thread.start()
    
    def convert_files(self):
        """转换MSG文件到EML格式"""
        total_files = len(self.file_items)
        success_count = 0
        failed_count = 0
        
        self.progress.config(maximum=total_files)
        
        for index, (item_id, file_info) in enumerate(self.file_items.items()):
            msg_file = file_info['path']
            filename = file_info['filename']
            
            try:
                # 更新状态为转换中
                self.root.after(0, lambda i=item_id: self.file_tree.set(i, 'status', '转换中...'))
                self.root.after(0, lambda f=filename: self.status_label.config(
                    text=f"正在转换: {f}"))
                
                # 打开MSG文件
                msg = extract_msg.openMsg(msg_file)
                
                # 创建EML内容
                eml_content = self.create_eml_content(msg)
                
                # 确定输出目录
                if self.output_dir:
                    output_dir = self.output_dir
                else:
                    output_dir = os.path.dirname(msg_file)
                
                os.makedirs(output_dir, exist_ok=True)
                
                # 生成输出文件名
                name_without_ext = os.path.splitext(filename)[0]
                name_without_ext = re.sub(r'[<>:"|?*]', '_', name_without_ext)
                eml_filename = f"{name_without_ext}.eml"
                eml_path = os.path.join(output_dir, eml_filename)
                
                # 处理同名文件
                counter = 1
                while os.path.exists(eml_path):
                    eml_filename = f"{name_without_ext}_{counter}.eml"
                    eml_path = os.path.join(output_dir, eml_filename)
                    counter += 1
                
                # 保存EML文件
                with open(eml_path, 'wb') as f:
                    f.write(eml_content.encode('utf-8'))
                
                # 更新文件信息
                file_info['status'] = 'success'
                file_info['output_file'] = eml_path
                
                self.conversion_results[msg_file] = {
                    'status': 'success',
                    'output_file': eml_path,
                    'options': {
                        'include_attachments': self.include_attachments.get(),
                        'preserve_headers': self.preserve_headers.get(),
                        'auto_decode': self.auto_decode.get(),
                        'detect_encoding': self.detect_encoding.get(),
                        'preserve_transport_headers': self.preserve_transport_headers.get(),
                        'show_ip_info': self.show_ip_info.get()
                    }
                }
                
                # 更新UI
                self.root.after(0, lambda i=item_id, f=eml_filename: (
                    self.file_tree.set(i, 'status', '已完成'),
                    self.file_tree.set(i, 'result', f)
                ))
                
                success_count += 1
                msg.close()
                
            except Exception as e:
                error_msg = str(e)
                file_info['status'] = 'failed'
                
                self.conversion_results[msg_file] = {
                    'status': 'failed',
                    'error': error_msg
                }
                
                # 更新UI显示错误
                self.root.after(0, lambda i=item_id, e=error_msg: (
                    self.file_tree.set(i, 'status', '转换失败'),
                    self.file_tree.set(i, 'result', f'错误: {e[:50]}...')
                ))
                
                failed_count += 1
            
            # 更新进度条
            self.root.after(0, lambda v=index+1: self.progress.config(value=v))
        
        # 转换完成
        summary = f"\n转换完成！成功: {success_count} 个，失败: {failed_count} 个\n"
        self.root.after(0, lambda: self.status_label.config(text=summary.strip()))
        
        if success_count > 0:
            self.root.after(0, lambda: messagebox.showinfo("转换完成", summary.strip()))
        
        # 恢复按钮状态
        self.root.after(0, lambda: (
            self.convert_btn.config(state=tk.NORMAL),
            self.select_btn.config(state=tk.NORMAL),
            self.clear_btn.config(state=tk.NORMAL),
            self.remove_btn.config(state=tk.NORMAL),
            self.progress.config(value=0)
        ))
    
    # 以下是所有的辅助方法
    
    def create_eml_content(self, msg):
        """创建EML格式内容（增强版，包含完整传输信息）"""
        try:
            # 获取邮件正文内容
            body_text = self.safe_get_str(msg, 'body')
            html_text = self.safe_get_str(msg, 'htmlBody')
            
            # 如果HTML内容为空但有RTF内容，尝试获取RTF
            if not html_text and hasattr(msg, 'rtfBody'):
                rtf_text = self.safe_get_str(msg, 'rtfBody')
                if rtf_text and not body_text:
                    body_text = rtf_text
            
            # 检查是否有附件
            has_attachments = hasattr(msg, 'attachments') and len(msg.attachments) > 0
            
            # 创建根邮件对象
            if has_attachments and self.include_attachments.get():
                email_msg = MIMEMultipart('mixed')
                
                if body_text and html_text:
                    msg_body = MIMEMultipart('alternative')
                    msg_body.attach(MIMEText(body_text, 'plain', 'utf-8'))
                    msg_body.attach(MIMEText(html_text, 'html', 'utf-8'))
                    email_msg.attach(msg_body)
                elif html_text:
                    email_msg.attach(MIMEText(html_text, 'html', 'utf-8'))
                elif body_text:
                    email_msg.attach(MIMEText(body_text, 'plain', 'utf-8'))
                else:
                    email_msg.attach(MIMEText("", 'plain', 'utf-8'))
                    
            elif body_text and html_text:
                email_msg = MIMEMultipart('alternative')
                email_msg.attach(MIMEText(body_text, 'plain', 'utf-8'))
                email_msg.attach(MIMEText(html_text, 'html', 'utf-8'))
            elif html_text:
                email_msg = MIMEText(html_text, 'html', 'utf-8')
            else:
                email_msg = MIMEText(body_text or "", 'plain', 'utf-8')
            
            # 添加原始邮件头（如果启用了保留传输头选项）
            if self.preserve_transport_headers.get():
                original_headers = self.extract_original_headers(msg)
                for header_name, header_value in original_headers:
                    if header_value and header_name.lower() not in ['content-type', 'content-transfer-encoding', 'mime-version']:
                        email_msg[header_name] = header_value
            
            # 设置基本邮件头（检查是否已存在）
            existing_headers = {key.lower() for key in email_msg.keys()}
            
            if 'subject' not in existing_headers:
                subject = self.safe_get_str(msg, 'subject')
                if subject:
                    email_msg['Subject'] = self.encode_header(subject)
            
            if 'from' not in existing_headers:
                sender = self.safe_get_str(msg, 'sender')
                if sender:
                    email_msg['From'] = self.encode_header(sender)
            
            if 'to' not in existing_headers:
                to_recipients = self.safe_get_str(msg, 'to')
                if to_recipients:
                    email_msg['To'] = self.encode_header(to_recipients)
            
            if 'cc' not in existing_headers:
                cc_recipients = self.safe_get_str(msg, 'cc')
                if cc_recipients:
                    email_msg['Cc'] = self.encode_header(cc_recipients)
            
            if 'bcc' not in existing_headers:
                bcc_recipients = self.safe_get_str(msg, 'bcc')
                if bcc_recipients:
                    email_msg['Bcc'] = self.encode_header(bcc_recipients)
            
            if 'date' not in existing_headers:
                date_obj = None
                if hasattr(msg, 'date'):
                    date_obj = msg.date
                elif hasattr(msg, 'sentOn'):
                    date_obj = msg.sentOn
                email_msg['Date'] = self.format_email_date(date_obj)
            
            if 'message-id' not in existing_headers:
                message_id = self.safe_get_str(msg, 'messageId')
                if message_id:
                    email_msg['Message-ID'] = message_id
                else:
                    email_msg['Message-ID'] = f"<{uuid.uuid4()}@msg-to-eml-converter>"
            
            if 'reply-to' not in existing_headers:
                reply_to = self.safe_get_str(msg, 'replyTo')
                if reply_to:
                    email_msg['Reply-To'] = self.encode_header(reply_to)
            
            # 设置MIME版本
            if 'mime-version' not in existing_headers:
                email_msg['MIME-Version'] = '1.0'
            
            # 添加MSG扩展属性（如果启用了保留MSG属性选项）
            if self.preserve_headers.get():
                self.add_extended_headers(email_msg, msg)
            
            # 添加额外的传输信息（如果启用了显示IP信息选项）
            if self.show_ip_info.get():
                self.add_ip_related_headers(email_msg, msg)
            
            # 添加转换器信息
            email_msg['X-Converted-From'] = 'MSG'
            email_msg['X-Converter'] = 'Enhanced-MSG-to-EML-Converter-v2'
            email_msg['X-Conversion-Date'] = formatdate(localtime=True)
            
            # 处理附件
            if has_attachments and self.include_attachments.get():
                for i, attachment in enumerate(msg.attachments):
                    try:
                        filename = self.get_attachment_filename(attachment, i)
                        mime_part = self.create_attachment_mime(attachment, filename)
                        email_msg.attach(mime_part)
                    except Exception as e:
                        print(f"处理附件 {i+1} 时出错: {e}")
            
            return email_msg.as_string()
            
        except Exception as e:
            print(f"创建EML内容时出错: {e}")
            error_msg = MIMEText(f"MSG文件转换错误:\n{str(e)}", 'plain', 'utf-8')
            error_msg['Subject'] = "MSG转换错误"
            error_msg['From'] = "enhanced-msg-to-eml-converter@localhost"
            error_msg['Date'] = formatdate(localtime=True)
            return error_msg.as_string()
    
    def extract_original_headers(self, msg):
        """提取MSG文件中的原始邮件头"""
        original_headers = []
        
        try:
            # 尝试多种方式获取原始邮件头
            # 方法1：transportMessageHeaders（通常包含最完整的头信息）
            if hasattr(msg, 'transportMessageHeaders'):
                transport_headers = None
                try:
                    transport_headers = msg.transportMessageHeaders
                    if isinstance(transport_headers, bytes):
                        transport_headers = transport_headers.decode('utf-8', errors='replace')
                    elif transport_headers is not None:
                        transport_headers = str(transport_headers)
                except:
                    transport_headers = self.safe_get_str(msg, 'transportMessageHeaders')
                
                if transport_headers:
                    headers = self.parse_header_string(transport_headers)
                    if headers:
                        original_headers.extend(headers)
            
            # 方法2：header属性
            if hasattr(msg, 'header'):
                header_data = self.safe_get_str(msg, 'header')
                if header_data:
                    headers = self.parse_header_string(header_data)
                    original_headers.extend(headers)
            
            # 方法3：internetHeaders属性
            if hasattr(msg, 'internetHeaders'):
                internet_headers = self.safe_get_str(msg, 'internetHeaders')
                if internet_headers:
                    headers = self.parse_header_string(internet_headers)
                    original_headers.extend(headers)
            
            # 去重
            seen = set()
            unique_headers = []
            for header_name, header_value in original_headers:
                key = header_name.lower()
                if key not in seen or key == 'received':  # Received头可以有多个
                    seen.add(key)
                    unique_headers.append((header_name, header_value))
            
            return unique_headers
            
        except Exception as e:
            print(f"提取原始邮件头时出错: {e}")
            return []
    
    def parse_header_string(self, header_string):
        """解析邮件头字符串"""
        headers = []
        
        if not header_string:
            return headers
        
        try:
            # 标准化行结束符
            header_string = header_string.replace('\r\n', '\n').replace('\r', '\n')
            
            # 分割成行
            lines = header_string.split('\n')
            current_header = None
            current_value = []
            
            for line in lines:
                # 空行表示邮件头结束
                if not line.strip():
                    if current_header:
                        headers.append((current_header, ' '.join(current_value)))
                        current_header = None
                        current_value = []
                    continue
                
                # 以空格或制表符开头的行是上一个头的延续
                if line and line[0] in ' \t':
                    if current_header:
                        current_value.append(line.strip())
                # 包含冒号的行是新的头
                elif ':' in line:
                    # 保存之前的头
                    if current_header:
                        headers.append((current_header, ' '.join(current_value)))
                    
                    # 解析新的头
                    colon_pos = line.find(':')
                    header_name = line[:colon_pos].strip()
                    header_value = line[colon_pos + 1:].strip()
                    
                    if header_name:
                        current_header = header_name
                        current_value = [header_value] if header_value else []
            
            # 保存最后一个头
            if current_header:
                headers.append((current_header, ' '.join(current_value)))
            
        except Exception as e:
            print(f"解析邮件头字符串时出错: {e}")
        
        return headers
    
    def add_extended_headers(self, email_msg, msg):
        """添加MSG扩展属性"""
        try:
            # 会话相关
            conv_topic = self.safe_get_str(msg, 'conversationTopic')
            if conv_topic:
                email_msg['Thread-Topic'] = self.encode_header(conv_topic)
            
            conv_index = self.safe_get_str(msg, 'conversationIndex')
            if conv_index:
                email_msg['Thread-Index'] = conv_index
            
            # MSG特有属性
            if hasattr(msg, 'messageClass'):
                msg_class = self.safe_get_str(msg, 'messageClass')
                if msg_class:
                    email_msg['X-Message-Class'] = msg_class
            
            # 其他扩展属性
            extended_attrs = [
                ('sensitivity', 'X-Sensitivity'),
                ('flag', 'X-Flag-Status'),
                ('categories', 'X-Categories'),
                ('companies', 'X-Companies'),
                ('readReceiptRequested', 'X-Read-Receipt-Requested'),
                ('deliveryReceiptRequested', 'X-Delivery-Receipt-Requested'),
            ]
            
            for attr, header_name in extended_attrs:
                value = self.safe_get_str(msg, attr)
                if value:
                    email_msg[header_name] = self.encode_header(value)
                    
        except Exception as e:
            print(f"添加扩展头时出错: {e}")
    
    def add_ip_related_headers(self, email_msg, msg):
        """添加IP相关的额外信息"""
        try:
            # 添加发送和接收的SMTP地址
            sender_smtp = self.safe_get_str(msg, 'senderSmtpAddress')
            if sender_smtp:
                email_msg['X-Sender-SMTP-Address'] = sender_smtp
                
            received_smtp = self.safe_get_str(msg, 'receivedBySmtpAddress')
            if received_smtp:
                email_msg['X-Received-By-SMTP-Address'] = received_smtp
            
            # 添加时间戳信息（有助于追踪传输路径）
            if hasattr(msg, 'clientSubmitTime'):
                submit_time = msg.clientSubmitTime
                if submit_time:
                    email_msg['X-Client-Submit-Time'] = self.format_email_date(submit_time)
            
            if hasattr(msg, 'messageDeliveryTime'):
                delivery_time = msg.messageDeliveryTime
                if delivery_time:
                    email_msg['X-Message-Delivery-Time'] = self.format_email_date(delivery_time)
            
            # 添加提示信息
            email_msg['X-IP-Info-Note'] = 'IP addresses preserved from original MSG headers'
            
        except Exception as e:
            print(f"添加IP相关头时出错: {e}")
    
    def safe_get_str(self, obj, attr, default=""):
        """安全获取字符串属性"""
        try:
            if hasattr(obj, attr):
                value = getattr(obj, attr)
                if value is None:
                    return default
                
                # 处理字节串
                if isinstance(value, bytes):
                    if self.detect_encoding.get():
                        value, _ = self.detect_text_encoding(value)
                    else:
                        value = value.decode('utf-8', errors='replace')
                # 处理字符串
                elif isinstance(value, str):
                    if self.auto_decode.get():
                        value = self.auto_decode_content(value)
                else:
                    value = str(value)
                
                return value.strip() if value else default
            return default
        except Exception as e:
            print(f"获取属性 {attr} 时出错: {e}")
            return default
    
    def detect_text_encoding(self, text_data):
        """智能检测文本编码"""
        if not text_data:
            return text_data, 'utf-8'
        
        if isinstance(text_data, str):
            return text_data, 'utf-8'
        
        if isinstance(text_data, bytes):
            try:
                detected = chardet.detect(text_data)
                encoding = detected.get('encoding', 'utf-8')
                confidence = detected.get('confidence', 0)
                
                if confidence < 0.7:
                    for enc in ['utf-8', 'gbk', 'gb2312', 'big5', 'utf-16']:
                        try:
                            decoded_text = text_data.decode(enc)
                            return decoded_text, enc
                        except UnicodeDecodeError:
                            continue
                
                try:
                    decoded_text = text_data.decode(encoding)
                    return decoded_text, encoding
                except UnicodeDecodeError:
                    decoded_text = text_data.decode('utf-8', errors='replace')
                    return decoded_text, 'utf-8'
                    
            except Exception:
                decoded_text = text_data.decode('utf-8', errors='replace')
                return decoded_text, 'utf-8'
        
        return str(text_data), 'utf-8'
    
    def auto_decode_content(self, content):
        """自动解码Base64或Quoted-Printable编码的内容"""
        if not content or not isinstance(content, str):
            return content
        
        # Base64解码
        if self.is_base64_encoded(content):
            try:
                decoded_bytes = base64.b64decode(content)
                decoded_text, _ = self.detect_text_encoding(decoded_bytes)
                return decoded_text
            except:
                pass
        
        # Quoted-Printable解码
        if self.is_quoted_printable_encoded(content):
            try:
                decoded_bytes = quopri.decodestring(content.encode('ascii'))
                decoded_text, _ = self.detect_text_encoding(decoded_bytes)
                return decoded_text
            except:
                pass
        
        # RFC 2047解码
        if '=?' in content and '?=' in content:
            try:
                decoded_parts = decode_header(content)
                decoded_text = ""
                for part, encoding in decoded_parts:
                    if isinstance(part, bytes):
                        if encoding:
                            decoded_text += part.decode(encoding)
                        else:
                            decoded_text += part.decode('utf-8', errors='replace')
                    else:
                        decoded_text += str(part)
                return decoded_text
            except:
                pass
        
        return content
    
    def is_base64_encoded(self, text):
        """检查文本是否是Base64编码"""
        if not text or len(text) < 4:
            return False
        
        import string
        base64_chars = string.ascii_letters + string.digits + '+/='
        
        cleaned = text.replace('\n', '').replace('\r', '').replace(' ', '')
        
        if not all(c in base64_chars for c in cleaned):
            return False
        
        if len(cleaned) % 4 != 0:
            return False
        
        try:
            base64.b64decode(cleaned, validate=True)
            return len(cleaned) > 20
        except:
            return False
    
    def is_quoted_printable_encoded(self, text):
        """检查文本是否是Quoted-Printable编码"""
        if not text:
            return False
        
        qp_pattern = re.compile(r'=([0-9A-Fa-f]{2})')
        return bool(qp_pattern.search(text))
    
    def encode_header(self, text):
        """编码邮件头"""
        if not text:
            return ""
        try:
            text.encode('ascii')
            return text
        except UnicodeEncodeError:
            return str(Header(text, 'utf-8'))
    
    def format_email_date(self, date_obj):
        """格式化日期"""
        if not date_obj:
            return formatdate(localtime=True)
        
        try:
            if isinstance(date_obj, str):
                return date_obj
            elif hasattr(date_obj, 'strftime'):
                return formatdate(date_obj.timestamp(), localtime=True)
            else:
                return formatdate(localtime=True)
        except:
            return formatdate(localtime=True)
    
    def get_attachment_filename(self, attachment, index):
        """获取附件文件名"""
        filename = None
        
        filename_attrs = ['longFilename', 'shortFilename', 'FileName', 'displayName']
        
        for attr in filename_attrs:
            if hasattr(attachment, attr):
                filename = self.safe_get_str(attachment, attr)
                if filename:
                    break
        
        if not filename:
            filename = f"attachment_{index + 1}"
        
        filename = re.sub(r'[<>:"|?*]', '_', filename)
        
        return filename
    
    def create_attachment_mime(self, attachment, filename):
        """创建附件MIME部分"""
        try:
            attachment_data = None
            if hasattr(attachment, 'data'):
                attachment_data = attachment.data
            
            if attachment_data:
                mime_type, _ = mimetypes.guess_type(filename)
                
                if mime_type:
                    maintype, subtype = mime_type.split('/', 1)
                    part = MIMEBase(maintype, subtype)
                else:
                    part = MIMEBase('application', 'octet-stream')
                
                if isinstance(attachment_data, bytes):
                    part.set_payload(attachment_data)
                else:
                    part.set_payload(str(attachment_data).encode('utf-8'))
                
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                
                return part
            else:
                return self.create_attachment_placeholder(filename)
                
        except Exception as e:
            print(f"创建附件MIME时出错: {e}")
            return self.create_attachment_placeholder(filename, f"错误: {str(e)}")
    
    def create_attachment_placeholder(self, filename, error_msg=None):
        """创建附件占位符"""
        if error_msg:
            placeholder_text = f"[附件内容不可用: {error_msg}]"
        else:
            placeholder_text = "[附件内容未包含]"
        
        part = MIMEBase('text', 'plain')
        part.set_payload(placeholder_text.encode('utf-8'))
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        
        return part
    
    def view_email_headers(self):
        """查看邮件头详情（修复版，可点击颜色过滤）"""
        # 获取选中的项目
        selection = self.file_tree.selection()
        if not selection:
            messagebox.showinfo("提示", "请先选择一个已转换的文件")
            return
        
        item = selection[0]
        if item not in self.file_items:
            return
        
        file_info = self.file_items[item]
        if file_info['status'] != 'success' or not file_info.get('output_file'):
            messagebox.showinfo("提示", "请选择一个已成功转换的文件")
            return
        
        eml_file = file_info['output_file']
        if not os.path.exists(eml_file):
            messagebox.showinfo("提示", "转换后的文件不存在")
            return
        
        # 创建新窗口
        headers_window = tk.Toplevel(self.root)
        headers_window.title(f"邮件头详情 - {os.path.basename(eml_file)}")
        headers_window.geometry("1000x750")
        
        # 创建主框架
        main_frame = ttk.Frame(headers_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建控制面板框架
        control_frame = ttk.LabelFrame(main_frame, text="过滤控制", padding="10")
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 定义颜色和样式
        color_config = {
            'transport': {'color': '#FFFF00', 'name': '传输相关头'},
            'extended': {'color': '#90EE90', 'name': 'MSG扩展属性'},
            'ip': {'color': '#FFB6C1', 'name': '包含IP地址'},
            'converter': {'color': '#E0E0E0', 'name': '转换器信息'},
            'basic': {'color': '#FFFFFF', 'name': '基本邮件头'}
        }
        
        # 创建过滤状态变量
        filter_vars = {}
        for key in color_config.keys():
            filter_vars[key] = tk.BooleanVar(value=True)
        
        # 创建文本框框架
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建文本框和滚动条
        headers_text = tk.Text(text_frame, wrap=tk.WORD, font=("Consolas", 9))
        headers_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=headers_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        headers_text.configure(yscrollcommand=scrollbar.set)
        
        # 存储所有邮件头数据
        headers_data = {
            'transport': [],
            'extended': [],
            'ip': [],
            'converter': [],
            'basic': []
        }
        
        try:
            # 读取EML文件
            with open(eml_file, 'r', encoding='utf-8', errors='replace') as f:
                eml_content = f.read()
            
            # 解析邮件
            msg = email.message_from_string(eml_content)
            
            # IP地址匹配模式
            ip_pattern = re.compile(r'\b(?:\d{1,3}\.){3}\d{1,3}\b')
            
            # 分类邮件头
            for key, value in msg.items():
                header_line = f"{key}: {value}\n"
                
                # 根据头类型分类
                if key.lower() in ['x-converted-from', 'x-converter', 'x-conversion-date']:
                    headers_data['converter'].append((key, value, header_line))
                elif ip_pattern.search(value):
                    headers_data['ip'].append((key, value, header_line))
                elif key.lower() in ['received', 'x-mailer', 'x-originating-ip', 'x-sender-ip', 
                                   'authentication-results', 'received-spf', 'dkim-signature',
                                   'x-sender-smtp-address', 'x-received-by-smtp-address']:
                    headers_data['transport'].append((key, value, header_line))
                elif key.lower().startswith('thread-') or key.lower() in ['x-message-class', 
                                                                         'x-sensitivity', 
                                                                         'x-flag-status',
                                                                         'x-categories',
                                                                         'x-companies']:
                    headers_data['extended'].append((key, value, header_line))
                else:
                    headers_data['basic'].append((key, value, header_line))
            
        except Exception as e:
            headers_text.insert(tk.END, f"读取邮件头时出错: {str(e)}")
            return
        
        def update_display():
            """更新显示内容"""
            headers_text.config(state=tk.NORMAL)
            headers_text.delete(1.0, tk.END)
            
            # 添加标题
            headers_text.insert(tk.END, "=== 邮件头信息 ===\n\n", "section_header")
            
            # 统计信息
            total_count = 0
            shown_count = 0
            category_counts = {}
            
            # 按顺序显示各类头部
            display_order = ['basic', 'transport', 'extended', 'ip', 'converter']
            
            for category in display_order:
                if not filter_vars[category].get():
                    continue
                    
                headers_list = headers_data[category]
                if not headers_list:
                    continue
                
                # 添加分类标题
                if category == 'basic':
                    section_title = "\n--- 基本邮件头 ---\n"
                elif category == 'transport':
                    section_title = "\n--- 传输相关头 ---\n"
                elif category == 'extended':
                    section_title = "\n--- MSG扩展属性 ---\n"
                elif category == 'ip':
                    section_title = "\n--- 包含IP地址的头 ---\n"
                elif category == 'converter':
                    section_title = "\n--- 转换器信息 ---\n"
                
                headers_text.insert(tk.END, section_title, "category_header")
                
                # 添加头部内容
                for key, value, header_line in headers_list:
                    headers_text.insert(tk.END, header_line, f"{category}_header")
                    shown_count += 1
                
                category_counts[category] = len(headers_list)
            
            # 计算总数
            for headers_list in headers_data.values():
                total_count += len(headers_list)
            
            # 添加统计信息
            stats_text = f"\n=== 统计信息 ===\n"
            stats_text += f"显示邮件头: {shown_count} / {total_count}\n"
            
            for category, count in category_counts.items():
                if count > 0:
                    category_name = color_config.get(category, {}).get('name', category)
                    stats_text += f"{category_name}: {count}\n"
            
            headers_text.insert(tk.END, stats_text, "stats")
            
            headers_text.config(state=tk.DISABLED)
        
        def create_filter_callback(filter_type):
            """创建过滤回调函数"""
            def callback():
                if filter_type == 'all':
                    # 检查当前是否全部显示
                    all_showing = all(filter_vars[key].get() for key in color_config.keys())
                    # 如果全部显示，则隐藏全部；否则显示全部
                    new_state = not all_showing
                    
                    for key in color_config.keys():
                        filter_vars[key].set(new_state)
                else:
                    # 切换单个过滤器
                    current = filter_vars[filter_type].get()
                    filter_vars[filter_type].set(not current)
                
                # 更新按钮状态
                update_button_states()
                # 更新显示
                update_display()
            
            return callback
        
        def update_button_states():
            """更新按钮状态"""
            # 检查是否全部显示
            all_showing = all(filter_vars[key].get() for key in color_config.keys())
            
            # 更新全选按钮
            if all_showing:
                all_btn.config(relief=tk.SUNKEN, bg="#D0D0D0", text="● 隐藏全部")
            else:
                all_btn.config(relief=tk.RAISED, bg="#F0F0F0", text="● 显示全部")
            
            # 更新颜色按钮
            for filter_type, btn in color_buttons.items():
                if filter_vars[filter_type].get():
                    btn.config(relief=tk.SUNKEN, 
                              bg=self.darken_color(color_config[filter_type]['color']))
                else:
                    btn.config(relief=tk.RAISED, 
                              bg=color_config[filter_type]['color'])
        
        # 添加显示/隐藏全部按钮
        all_btn = tk.Button(control_frame, 
                           text="● 显示全部", 
                           font=("Arial", 10, "bold"),
                           bg="#F0F0F0",
                           relief=tk.RAISED,
                           borderwidth=2,
                           padx=15,
                           pady=5,
                           command=create_filter_callback('all'))
        all_btn.pack(side=tk.LEFT, padx=(0, 15))
        
        # 创建颜色按钮
        color_buttons = {}
        for filter_type, config in color_config.items():
            btn = tk.Button(control_frame, 
                           text=f"● {config['name']}", 
                           bg=config['color'],
                           font=("Arial", 9, "bold"),
                           relief=tk.RAISED,
                           borderwidth=2,
                           padx=10,
                           pady=3,
                           command=create_filter_callback(filter_type))
            btn.pack(side=tk.LEFT, padx=(0, 10))
            color_buttons[filter_type] = btn
        
        # 配置文本标签样式
        headers_text.tag_configure("section_header", font=("Arial", 14, "bold"), foreground="blue")
        headers_text.tag_configure("category_header", font=("Arial", 11, "bold"), foreground="darkgreen")
        headers_text.tag_configure("stats", font=("Arial", 10, "bold"), foreground="purple")
        
        # 配置各类头部样式
        for category, config in color_config.items():
            headers_text.tag_configure(f"{category}_header", background=config['color'])
        
        # 初始化显示
        update_button_states()
        update_display()
        
        # 关闭按钮
        close_frame = ttk.Frame(main_frame)
        close_frame.pack(pady=(10, 0))
        
        close_btn = ttk.Button(close_frame, text="关闭", command=headers_window.destroy)
        close_btn.pack()
        
        # 添加说明
        info_label = ttk.Label(close_frame, 
                              text="提示：点击上方颜色按钮可过滤显示对应类型的邮件头", 
                              font=("Arial", 9), 
                              foreground="gray")
        info_label.pack(pady=(5, 0))
    
    def darken_color(self, color):
        """将颜色变暗（用于按下状态）"""
        if color == '#FFFFFF':
            return '#E0E0E0'
        elif color == '#FFFF00':
            return '#E6E600'
        elif color == '#90EE90':
            return '#7DD67D'
        elif color == '#FFB6C1':
            return '#E6A3AE'
        elif color == '#E0E0E0':
            return '#CCCCCC'
        else:
            return color
    
    def view_msg_attributes(self):
        """查看MSG文件的所有属性"""
        selection = self.file_tree.selection()
        if not selection:
            messagebox.showinfo("提示", "请先选择一个MSG文件")
            return
        
        item = selection[0]
        if item not in self.file_items:
            return
        
        msg_file = self.file_items[item]['path']
        
        # 创建新窗口
        attrs_window = tk.Toplevel(self.root)
        attrs_window.title(f"MSG属性 - {os.path.basename(msg_file)}")
        attrs_window.geometry("900x700")
        
        # 创建文本框
        text_frame = ttk.Frame(attrs_window, padding="10")
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        attrs_text = tk.Text(text_frame, wrap=tk.WORD)
        attrs_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=attrs_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        attrs_text.configure(yscrollcommand=scrollbar.set)
        
        try:
            msg = extract_msg.openMsg(msg_file)
            
            attrs_text.insert(tk.END, "=== MSG文件属性列表 ===\n\n", "section_header")
            
            # 列出所有属性
            all_attrs = dir(msg)
            for attr in sorted(all_attrs):
                if not attr.startswith('_'):
                    try:
                        value = getattr(msg, attr)
                        if not callable(value):
                            value_str = str(value)[:200]
                            if len(str(value)) > 200:
                                value_str += "..."
                            attrs_text.insert(tk.END, f"{attr}: {value_str}\n")
                    except:
                        pass
            
            # 显示原始邮件头
            attrs_text.insert(tk.END, "\n\n=== 原始邮件头 ===\n", "section_header")
            original_headers = self.extract_original_headers(msg)
            if original_headers:
                for name, value in original_headers:
                    attrs_text.insert(tk.END, f"{name}: {value}\n")
            else:
                attrs_text.insert(tk.END, "未找到原始邮件头\n")
            
            attrs_text.tag_configure("section_header", font=("Arial", 12, "bold"), foreground="blue")
            
            msg.close()
            
        except Exception as e:
            attrs_text.insert(tk.END, f"读取MSG文件时出错: {str(e)}")
        
        attrs_text.configure(state=tk.DISABLED)
        
        # 关闭按钮
        close_btn = ttk.Button(attrs_window, text="关闭", command=attrs_window.destroy)
        close_btn.pack(pady=10)
    
    def test_option_effects(self):
        """测试不同选项的效果"""
        selection = self.file_tree.selection()
        if not selection:
            messagebox.showinfo("提示", "请先选择一个MSG文件")
            return
        
        item = selection[0]
        if item not in self.file_items:
            return
        
        msg_file = self.file_items[item]['path']
        
        # 创建测试窗口
        test_window = tk.Toplevel(self.root)
        test_window.title(f"选项效果对比 - {os.path.basename(msg_file)}")
        test_window.geometry("1000x700")
        
        # 创建主框架
        main_frame = ttk.Frame(test_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建选项卡
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        try:
            msg = extract_msg.openMsg(msg_file)
            
            # 保存当前选项
            original_options = {
                'preserve_transport_headers': self.preserve_transport_headers.get(),
                'preserve_headers': self.preserve_headers.get(),
                'show_ip_info': self.show_ip_info.get()
            }
            
            # 测试用例
            test_cases = [
                {
                    'name': '基本转换',
                    'preserve_transport_headers': False,
                    'preserve_headers': False,
                    'show_ip_info': False
                },
                {
                    'name': '仅传输头',
                    'preserve_transport_headers': True,
                    'preserve_headers': False,
                    'show_ip_info': False
                },
                {
                    'name': '仅MSG属性',
                    'preserve_transport_headers': False,
                    'preserve_headers': True,
                    'show_ip_info': False
                },
                {
                    'name': '全部启用',
                    'preserve_transport_headers': True,
                    'preserve_headers': True,
                    'show_ip_info': True
                }
            ]
            
            for test_case in test_cases:
                # 设置选项
                self.preserve_transport_headers.set(test_case['preserve_transport_headers'])
                self.preserve_headers.set(test_case['preserve_headers'])
                self.show_ip_info.set(test_case['show_ip_info'])
                
                # 创建EML内容
                eml_content = self.create_eml_content(msg)
                eml_msg = email.message_from_string(eml_content)
                
                # 创建选项卡页面
                tab_frame = ttk.Frame(notebook)
                notebook.add(tab_frame, text=test_case['name'])
                
                # 创建文本框
                text_widget = tk.Text(tab_frame, wrap=tk.WORD)
                text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                
                scrollbar = ttk.Scrollbar(tab_frame, orient=tk.VERTICAL, command=text_widget.yview)
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                text_widget.configure(yscrollcommand=scrollbar.set)
                
                # 显示选项和结果
                text_widget.insert(tk.END, f"=== 选项设置 ===\n")
                text_widget.insert(tk.END, f"保留传输头: {'是' if test_case['preserve_transport_headers'] else '否'}\n")
                text_widget.insert(tk.END, f"保留MSG属性: {'是' if test_case['preserve_headers'] else '否'}\n")
                text_widget.insert(tk.END, f"显示IP信息: {'是' if test_case['show_ip_info'] else '否'}\n\n")
                
                text_widget.insert(tk.END, "=== 邮件头 ===\n")
                
                header_count = 0
                for key, value in eml_msg.items():
                    header_count += 1
                    text_widget.insert(tk.END, f"{key}: {value}\n")
                
                text_widget.insert(tk.END, f"\n总邮件头数: {header_count}\n")
                
                text_widget.configure(state=tk.DISABLED)
            
            # 恢复原始选项
            for option, value in original_options.items():
                getattr(self, option).set(value)
            
            msg.close()
            
        except Exception as e:
            messagebox.showerror("错误", f"测试时出错: {str(e)}")
        
        # 关闭按钮
        close_btn = ttk.Button(main_frame, text="关闭", command=test_window.destroy)
        close_btn.pack(pady=(10, 0))

def main():
    """主函数"""
    root = tk.Tk()
    app = EnhancedMSGToEMLConverter(root)
    
    # 设置最小窗口大小
    root.minsize(800, 600)
    
    # 居中显示窗口
    
    
    root.mainloop()

if __name__ == "__main__":
    main()