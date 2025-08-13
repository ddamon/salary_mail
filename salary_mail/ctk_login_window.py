# coding:utf-8
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import hashlib
from datetime import datetime
from salary_mail.db_instance import User
import os
from PIL import Image

class CTKLoginWindow(ctk.CTk):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.title('登录 - 工资条管理系统')

        # 初始化延迟任务列表，用于窗口关闭时取消
        self.pending_tasks = []

        # 设置窗口样式
        self.geometry("420x480")
        self.minsize(380, 420)
        self.resizable(True, True)

        # 居中显示
        self.center_window()

        # 设置UI
        self.setup_ui()

        # 设置焦点
        self.schedule_task(100, self.safe_focus_username)

    def schedule_task(self, delay, callback):
        """安全地调度延迟任务"""
        try:
            task_id = self.after(delay, callback)
            self.pending_tasks.append(task_id)
            return task_id
        except:
            return None

    def cancel_all_tasks(self):
        """取消所有待执行的任务"""
        for task_id in self.pending_tasks:
            try:
                self.after_cancel(task_id)
            except:
                pass
        self.pending_tasks.clear()

    def safe_focus_username(self):
        """安全地设置用户名输入框焦点"""
        try:
            if self.winfo_exists() and hasattr(self, 'username_entry'):
                self.username_entry.focus()
        except:
            pass

    def safe_focus_password(self):
        """安全地设置密码输入框焦点"""
        try:
            if self.winfo_exists() and hasattr(self, 'password_entry'):
                self.password_entry.focus()
        except:
            pass

    def center_window(self):
        """窗口居中显示"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def setup_ui(self):
        """设置用户界面"""
        # 主容器
        main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=30, pady=30)

        # 标题区域
        self.create_header(main_frame)

        # 登录表单区域
        self.create_login_form(main_frame)

        # 底部信息
        self.create_footer(main_frame)

        # 检查首次使用
        self.check_first_time_use()

    def create_header(self, parent):
        """创建标题区域"""
        header_frame = ctk.CTkFrame(parent, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 40))

        # 主标题
        title_label = ctk.CTkLabel(
            header_frame,
            text="工资条管理系统",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=28, weight="bold"),
            text_color=("#1976d2", "#64b5f6")
        )
        title_label.pack(pady=(0, 10))

        # 副标题
        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="请登录以继续使用",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14),
            text_color=("#666666", "#aaaaaa")
        )
        subtitle_label.pack()

    def create_login_form(self, parent):
        """创建登录表单"""
        # 表单容器
        form_frame = ctk.CTkFrame(parent, corner_radius=15)
        form_frame.pack(fill="x", pady=(0, 30))

        # 表单内容
        form_content = ctk.CTkFrame(form_frame, fg_color="transparent")
        form_content.pack(fill="both", expand=True, padx=30, pady=30)

        # 用户名输入
        username_label = ctk.CTkLabel(
            form_content,
            text="用户名",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            anchor="w"
        )
        username_label.pack(fill="x", pady=(0, 5))

        self.username_entry = ctk.CTkEntry(
            form_content,
            placeholder_text="请输入用户名",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14),
            height=40,
            corner_radius=10
        )
        self.username_entry.pack(fill="x", pady=(0, 20))

        # 密码输入
        password_label = ctk.CTkLabel(
            form_content,
            text="密码",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            anchor="w"
        )
        password_label.pack(fill="x", pady=(0, 5))

        self.password_entry = ctk.CTkEntry(
            form_content,
            placeholder_text="请输入密码",
            show="●",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14),
            height=40,
            corner_radius=10
        )
        self.password_entry.pack(fill="x", pady=(0, 30))

        # 登录按钮
        self.login_btn = ctk.CTkButton(
            form_content,
            text="登 录",
            command=self.login,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=16, weight="bold"),
            height=45,
            corner_radius=10,
            fg_color=("#1976d2", "#1976d2"),
            hover_color=("#1565c0", "#1565c0")
        )
        self.login_btn.pack(fill="x")

        # 绑定回车键
        self.username_entry.bind('<Return>', self._on_username_return)
        self.password_entry.bind('<Return>', self._on_password_return)

    def _on_return_key(self, event):
        """处理回车键事件"""
        try:
            self.login()
        except:
            pass

    def _on_username_return(self, event):
        """处理用户名输入框回车键事件"""
        try:
            if hasattr(self, 'password_entry'):
                self.password_entry.focus()
        except:
            pass

    def _on_password_return(self, event):
        """处理密码输入框回车键事件"""
        try:
            self.login()
        except:
            pass

    def create_footer(self, parent):
        """创建底部信息"""
        footer_frame = ctk.CTkFrame(parent, fg_color="transparent")
        footer_frame.pack(fill="x")

        # 版权信息
        copyright_label = ctk.CTkLabel(
            footer_frame,
            text="© 2025 工资条管理系统",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color=("#999999", "#666666")
        )
        copyright_label.pack()

    def check_first_time_use(self):
        """检查首次使用"""
        user_count = self.db.query(User).count()
        if user_count == 1:
            first_user = self.db.query(User).first()
            if first_user.username == 'admin' and first_user.last_login is None:
                self.schedule_task(500, self.show_first_time_dialog)
                self.username_entry.insert(0, 'admin')
                self.schedule_task(600, self.safe_focus_password)

    def show_first_time_dialog(self):
        """显示首次使用对话框"""
        dialog = CTKMessageDialog(
            self,
            title="欢迎使用",
            message="首次使用系统\n默认用户名：admin\n默认密码：admin123\n请登录后及时修改密码！",
            icon="info"
        )
        dialog.show()

    def login(self):
        """登录验证"""
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()

        if not username or not password:
            self.show_error_dialog('用户名和密码不能为空！')
            return

        # 密码加密
        pwd_hash = hashlib.sha256(password.encode()).hexdigest()

        # 验证用户
        user = self.db.query(User).filter_by(username=username).first()
        if not user or user.password != pwd_hash:
            self.show_error_dialog('用户名或密码错误！')
            return

        # 更新登录时间
        user.last_login = datetime.now()
        self.db.commit()

        # 保存登录用户信息
        self.current_user = user

        # 取消所有待执行的任务
        self.cancel_all_tasks()

        self.destroy()

    def show_error_dialog(self, message):
        """显示错误对话框"""
        dialog = CTKMessageDialog(
            self,
            title="错误",
            message=message,
            icon="error"
        )
        dialog.show()


class CTKMessageDialog(ctk.CTkToplevel):
    """自定义消息对话框"""

    def __init__(self, parent, title="提示", message="", icon="info"):
        super().__init__(parent)

        # 初始化延迟任务列表
        self.pending_tasks = []

        self.title(title)
        self.geometry("350x200")
        self.resizable(False, False)

        # 设置为模态窗口
        self.transient(parent)
        self.grab_set()

        # 居中显示
        self.center_on_parent(parent)

        # 创建UI
        self.create_ui(title, message, icon)

    def center_on_parent(self, parent):
        """在父窗口中心显示"""
        self.update_idletasks()

        # 获取父窗口位置和大小
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()

        # 计算居中位置
        x = parent_x + (parent_width - self.winfo_width()) // 2
        y = parent_y + (parent_height - self.winfo_height()) // 2

        self.geometry(f"+{x}+{y}")

    def schedule_task(self, delay, callback):
        """安全地调度延迟任务"""
        try:
            task_id = self.after(delay, callback)
            self.pending_tasks.append(task_id)
            return task_id
        except:
            return None

    def cancel_all_tasks(self):
        """取消所有待执行的任务"""
        for task_id in self.pending_tasks:
            try:
                self.after_cancel(task_id)
            except:
                pass
        self.pending_tasks.clear()

    def safe_focus_button(self, button):
        """安全地设置按钮焦点"""
        try:
            # 检查对话框和按钮是否都存在
            if self.winfo_exists() and button and button.winfo_exists():
                button.focus()
        except:
            pass

    def create_ui(self, title, message, icon):
        """创建对话框UI"""
        # 主容器
        main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # 图标和消息区域
        content_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, pady=(0, 20))

        # 消息文本
        message_label = ctk.CTkLabel(
            content_frame,
            text=message,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14),
            justify="center"
        )
        message_label.pack(expand=True)

        # 确定按钮
        self.ok_button = ctk.CTkButton(
            main_frame,
            text="确定",
            command=self.safe_destroy,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            width=100,
            height=35
        )
        self.ok_button.pack()

        # 绑定回车键
        self.bind('<Return>', self._on_return_key)

        # 设置焦点
        self.schedule_task(100, lambda: self.safe_focus_button(self.ok_button))

    def _on_return_key(self, event):
        """处理回车键事件"""
        try:
            self.safe_destroy()
        except:
            pass

    def safe_destroy(self):
        """安全地销毁对话框"""
        try:
            # 取消所有待执行的任务
            self.cancel_all_tasks()
            # 确保窗口仍然存在再销毁
            if self.winfo_exists():
                self.destroy()
        except:
            pass

    def show(self):
        """显示对话框"""
        self.wait_window()
