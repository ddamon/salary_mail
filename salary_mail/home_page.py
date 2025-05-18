import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk
import tkinter.messagebox
from datetime import datetime
import openpyxl
import re
import threading
import base64
from email.mime.text import MIMEText
from email.utils import formataddr
from smtplib import SMTP_SSL, SMTP
import webbrowser
import tempfile
import os
import sys

# 修改导入语句
from salary_mail.EmployeeManageWin import EmployeeManageWin
from salary_mail.SalaryManageWin import SalaryManageWin
from salary_mail.db_instance import set_db
from salary_mail.setting_box import (
    EmailSettingWin,  # 只保留需要的类
    TemplateSettingWin,
    InfoManageWin
)
from salary_mail.login_window import LoginWindow

class HomePage(ThemedTk):
    def __init__(self):
        # 先登录
        db = set_db()
        login_window = LoginWindow(db)
        login_window.mainloop()
        
        # 如果登录窗口被关闭但没有登录成功，退出程序
        if not hasattr(login_window, 'current_user'):
            sys.exit()
        
        # 保存当前用户
        self.current_user = login_window.current_user
        self.db = db
        
        super().__init__()
        self.set_theme("arc")  # 使用一个现代主题
        self.title('工资条管理系统')
        
        # 加载图标
        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'assets', 'icon.png')
            self.iconphoto(True, tk.PhotoImage(file=icon_path))
        except:
            pass  # 如果图标加载失败，使用默认图标
        
        # 设置窗口位置在屏幕中心
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - 800) // 2
        y = (screen_height - 600) // 2
        self.geometry(f'800x600+{x}+{y}')
        
        self.configure(bg='#f5f5f5')  # 设置背景色
        self.resizable(True, True)  # 允许调整大小
        self.minsize(800, 600)  # 设置最小窗口大小
        
        # 配置自定义样式
        self.style = ttk.Style()
        self.style.configure('TFrame', background='white')
        self.style.configure('TLabel', background='white')
        self.style.configure('TButton', padding=8, font=('Microsoft YaHei UI', 10))
        self.style.map('TButton',
                       background=[('active', '#e3f2fd')],
                       relief=[('pressed', 'sunken')])
        
        # 卡片样式
        self.style.configure('Card.TFrame', 
                           background='white',
                           relief='solid',
                           borderwidth=1)
        
        # 阴影效果样式
        self.style.configure('Shadow.TFrame',
                           background='white',
                           relief='solid',
                           borderwidth=1)
        
        # 添加卡片悬停样式
        self.style.configure('CardHover.TFrame',
                            background='#f5f5f5',
                            relief='solid',
                            borderwidth=1)
        
        # 添加卡片按压效果样式
        self.style.configure('CardPressed.TFrame',
                            background='#e3f2fd',
                            relief='solid',
                            borderwidth=1)
        
        # 添加按钮动画效果
        self.style.configure('TButton.Active',
                            background='#e3f2fd',
                            relief='solid')
        
        self.setupUI()
        
        # 修改窗口状态追踪
        self.open_windows = {
            'email': None,
            'info': None,
            'salary': None,
            'employee': None
        }
        
    def get_center(self):
        """获取窗口中心坐标，用于子窗口定位"""
        px = self.winfo_x()
        py = self.winfo_y()
        pw = self.winfo_width()
        ph = self.winfo_height()
        return (int(px + pw/2), int(py + ph/2))
        
    def setupUI(self):
        """搭建主界面"""
        # 创建主框架
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill='both', expand=True)

        # 创建内容容器，用于居中显示
        content_container = ttk.Frame(main_frame)
        content_container.place(relx=0.5, rely=0.5, anchor='center')

        # 顶部区域
        header_frame = ttk.Frame(content_container)
        header_frame.pack(fill='x', pady=(0, 40))  # 增加与卡片的间距

        # Logo和标题区域 - 居中显示
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(expand=True)  # 使用expand确保居中

        # 标题和副标题
        text_container = ttk.Frame(title_frame)
        text_container.pack(side=tk.LEFT)

        title_label = ttk.Label(
            text_container,
            text="工资条管理系统",
            font=('Microsoft YaHei UI', 32, 'bold'),  # 增大标题字号
            foreground='#1976d2'
        )
        title_label.pack(anchor='w')

        subtitle_label = ttk.Label(
            text_container,
            text="简单高效的工资条发送管理工具",
            font=('Microsoft YaHei UI', 14),  # 增大副标题字号
            foreground='#757575'
        )
        subtitle_label.pack(anchor='w', pady=(5, 0))  # 增加与标题的间距

        # 功能卡片区域
        cards_container = ttk.Frame(content_container)
        cards_container.pack(pady=10)

        # 配置网格布局
        cards_container.grid_columnconfigure(0, weight=1, pad=15)  # 增加列间距
        cards_container.grid_columnconfigure(1, weight=1, pad=15)
        for i in range(2):
            cards_container.grid_rowconfigure(i, weight=1, pad=20)  # 增加行间距

        # 工资管理卡片
        self.create_card(
            cards_container, 0, 0,
            "工资管理",
            "导入和发送工资条",
            [("工资管理", self.show_salary_manage, "#4caf50")],
            "salary",
            padx=15, pady=12  # 增加卡片内边距
        )

        # 员工管理卡片
        self.create_card(
            cards_container, 0, 1,
            "员工管理",
            "管理员工信息和邮箱",
            [("员工管理", self.show_employee_manage, "#2196f3")],
            "employee",
            padx=15, pady=12
        )

        # 系统设置卡片 - 跨两列显示
        self.create_card(
            cards_container, 1, 0,
            "系统设置",
            "配置系统参数和邮件相关设置",
            [
                ("邮箱设置", self.show_email_setting, "#ff9800"),
                ("模板设置", self.show_template_setting_box, "#9c27b0"),
                ("信息管理", self.show_info_manage, "#795548")
            ],
            "settings",
            padx=15, pady=12,
            columnspan=2,  # 跨两列
            button_width=12
        )

        # 底部状态栏
        status_frame = ttk.Frame(main_frame, style='Shadow.TFrame')
        status_frame.pack(side='bottom', fill='x')

        # 版权信息
        ttk.Label(
            status_frame,
            text="© 2025 工资条管理系统",
            font=('Microsoft YaHei UI', 9),
            foreground='#757575'
        ).pack(side=tk.LEFT, padx=20, pady=12)  # 增加内边距

        # 状态信息
        self.status_label = ttk.Label(
            status_frame,
            text="就绪",
            font=('Microsoft YaHei UI', 9),
            foreground='#4caf50'
        )
        self.status_label.pack(side=tk.RIGHT, padx=20, pady=12)

    def create_card(self, parent, row, col, title, description, buttons, icon_name, **kwargs):
        """创建功能卡片"""
        # 提取自定义参数
        button_width = kwargs.pop('button_width', 15)
        
        # 卡片容器
        card = ttk.Frame(parent, style='Card.TFrame')
        card.grid(row=row, column=col, sticky='nsew', **kwargs)

        # 卡片内容框架
        content = ttk.Frame(card, padding=15)
        content.pack(fill='both', expand=True)

        # 标题区域
        header = ttk.Frame(content)
        header.pack(fill='x', pady=(0, 10))

        # 加载图标
        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'assets', f'{icon_name}.png')
            icon = tk.PhotoImage(file=icon_path)
            icon = icon.subsample(3, 3)  # 缩小到原来的1/3
            icon_label = ttk.Label(header, image=icon)
            icon_label.image = icon
            icon_label.pack(side=tk.LEFT, padx=(0, 10))
        except Exception as e:
            print(f"图标加载失败: {e}")

        # 标题
        ttk.Label(
            header,
            text=title,
            font=('Microsoft YaHei UI', 16, 'bold'),
            foreground='#1976d2'
        ).pack(side=tk.LEFT)

        # 描述
        ttk.Label(
            content,
            text=description,
            font=('Microsoft YaHei UI', 10),
            foreground='#757575',
            wraplength=250
        ).pack(fill='x', pady=(0, 15))

        # 按钮区域
        btn_frame = ttk.Frame(content)
        btn_frame.pack(fill='x')

        for text, command, color in buttons:
            btn = ttk.Button(
                btn_frame,
                text=text,
                command=command,
                style='TButton',
                width=button_width
            )
            btn.pack(side=tk.LEFT, padx=2)

        # 添加卡片悬停效果
        def on_enter(e):
            card.configure(style='CardHover.TFrame')

        def on_leave(e):
            card.configure(style='Card.TFrame')

        def on_click(e):
            card.configure(style='CardPressed.TFrame')
            self.after(100, lambda: card.configure(style='Card.TFrame'))

        card.bind('<Enter>', on_enter)
        card.bind('<Leave>', on_leave)
        card.bind('<Button-1>', on_click)

    def show_salary_manage(self):
        """显示工资管理窗口"""
        try:
            if self.open_windows.get('salary') and self.open_windows['salary'].winfo_exists():
                self.open_windows['salary'].focus_force()
                return
            dialog = SalaryManageWin(parent=self)
            self.open_windows['salary'] = dialog
            dialog.protocol("WM_DELETE_WINDOW", lambda: self._on_dialog_close('salary'))
        except Exception as e:
            print(f"打开工资管理窗口失败: {e}")
            self.open_windows['salary'] = None

    def show_email_setting(self):
        """显示邮箱设置窗口"""
        try:
            if self.open_windows.get('email') and self.open_windows['email'].winfo_exists():
                self.open_windows['email'].focus_force()
                return
            dialog = EmailSettingWin(parent=self)
            self.open_windows['email'] = dialog
            dialog.protocol("WM_DELETE_WINDOW", lambda: self._on_dialog_close('email'))
        except Exception as e:
            print(f"打开邮箱设置窗口失败: {e}")
            self.open_windows['email'] = None

    def show_template_setting_box(self):
        """显示邮件模板设置窗口"""
        dialog = TemplateSettingWin(parent=self)
        self.wait_window(dialog)

    def show_employee_manage(self):
        """显示员工管理窗口"""
        try:
            if self.open_windows.get('employee') and self.open_windows['employee'].winfo_exists():
                self.open_windows['employee'].focus_force()
                return
            dialog = EmployeeManageWin(parent=self)
            self.open_windows['employee'] = dialog
            dialog.protocol("WM_DELETE_WINDOW", lambda: self._on_dialog_close('employee'))
        except Exception as e:
            print(f"打开员工管理窗口失败: {e}")
            self.open_windows['employee'] = None

    def show_info_manage(self):
        """显示信息管理窗口"""
        try:
            if self.open_windows.get('info') and self.open_windows['info'].winfo_exists():
                self.open_windows['info'].focus_force()
                return
            dialog = InfoManageWin(parent=self)
            self.open_windows['info'] = dialog
            dialog.protocol("WM_DELETE_WINDOW", lambda: self._on_dialog_close('info'))
        except Exception as e:
            print(f"打开信息管理窗口失败: {e}")
            self.open_windows['info'] = None

    def update_status(self, text, color='#4caf50'):
        """更新状态栏文本，带渐变动画效果"""
        def fade_in():
            alpha = 0.0
            while alpha < 1.0:
                self.status_label.configure(foreground=color + f"{int(alpha * 255):02x}")
                alpha += 0.1
                self.after(50)
        
        self.status_label.configure(text=text)
        self.after(0, fade_in) 

    def _on_dialog_close(self, window_key):
        """处理窗口关闭"""
        try:
            if self.open_windows.get(window_key):
                self.open_windows[window_key].destroy()
            self.open_windows[window_key] = None
        except Exception as e:
            print(f"关闭窗口失败: {e}")
            self.open_windows[window_key] = None 