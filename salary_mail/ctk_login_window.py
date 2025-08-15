# coding:utf-8
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import hashlib
from datetime import datetime
from salary_mail.db_instance import User
import os
from PIL import Image
from .theme_config import theme_manager as responsive_theme_manager

class CTKLoginWindow(ctk.CTk):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.title('登录 - 工资条管理系统')

        # 初始化延迟任务列表，用于窗口关闭时取消
        self.pending_tasks = []

        # 获取响应式窗口配置但不自动居中
        self.responsive_config = responsive_theme_manager.get_responsive_config()
        window_config = self.responsive_config.get('login_window', self.responsive_config['main_window'])

        # 设置窗口尺寸但不设置位置
        width = window_config['width']
        height = window_config['height']
        self.geometry(f'{width}x{height}')

        if 'min_width' in window_config and 'min_height' in window_config:
            self.minsize(window_config['min_width'], window_config['min_height'])

        self.resizable(True, True)

        # 设置UI
        self.setup_ui()

        # UI加载完成后再居中显示
        self.schedule_task(200, self.ensure_centered)

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

    def ensure_centered(self):
        """确保窗口在屏幕中心显示"""
        try:
            if self.winfo_exists():
                # 强制更新窗口信息，确保获取准确的尺寸
                self.update_idletasks()

                # 获取当前窗口尺寸
                current_width = self.winfo_width()
                current_height = self.winfo_height()

                # 使用主题管理器的居中方法
                responsive_theme_manager.center_window_on_screen(self)

                # 验证居中效果
                self.update_idletasks()
                final_x = self.winfo_x()
                final_y = self.winfo_y()

                # 如果您觉得位置不对，可以手动微调
                if final_x < 100 or final_y < 100:
                    # 手动计算居中位置
                    screen_w = self.winfo_screenwidth()
                    screen_h = self.winfo_screenheight()

                    # 您可以调整这些值来微调位置
                    manual_x = (screen_w - current_width) // 2
                    manual_y = (screen_h - current_height) // 2 - 50  # 向上调整50像素

                    self.geometry(f"+{manual_x}+{manual_y}")
                    self.update()

        except Exception as e:
            print(f"登录窗口居中失败: {e}")



    def setup_ui(self):
        """设置响应式用户界面"""
        # 获取响应式配置
        padding = responsive_theme_manager.get_size('padding_large')

        # 主容器
        main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=padding, pady=padding)

        # 标题区域
        self.create_header(main_frame)

        # 登录表单区域
        self.create_login_form(main_frame)

        # 底部信息
        self.create_footer(main_frame)

        # 检查首次使用
        self.check_first_time_use()

    def create_header(self, parent):
        """创建响应式标题区域"""
        # 获取响应式配置
        section_spacing = responsive_theme_manager.get_size('margin_xl')
        title_font_size = responsive_theme_manager.get_font('title')[1]
        subtitle_font_size = responsive_theme_manager.get_font('subtitle')[1]
        show_subtitle = responsive_theme_manager.get_responsive_config().get('show_subtitle', True)

        header_frame = ctk.CTkFrame(parent, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, section_spacing))

        # 主标题
        title_label = ctk.CTkLabel(
            header_frame,
            text="工资条管理系统",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=title_font_size, weight="bold"),
            text_color=("#1976d2", "#64b5f6")
        )
        title_label.pack(pady=(0, 10))

        # 副标题（根据屏幕尺寸决定是否显示）
        if show_subtitle:
            subtitle_label = ctk.CTkLabel(
                header_frame,
                text="请登录以继续使用",
                font=ctk.CTkFont(family="Microsoft YaHei UI", size=subtitle_font_size),
                text_color=("#666666", "#aaaaaa")
            )
            subtitle_label.pack()

    def create_login_form(self, parent):
        """创建响应式登录表单"""
        # 获取响应式配置
        form_padding = responsive_theme_manager.get_size('padding_large')
        section_spacing = responsive_theme_manager.get_size('margin_large')
        show_card_effects = responsive_theme_manager.get_responsive_config().get('card_effects', True)

        # 表单容器
        corner_radius = 15 if show_card_effects else 8
        form_frame = ctk.CTkFrame(parent, corner_radius=corner_radius)
        form_frame.pack(fill="x", pady=(0, section_spacing))

        # 表单内容
        form_content = ctk.CTkFrame(form_frame, fg_color="transparent")
        form_content.pack(fill="both", expand=True, padx=form_padding, pady=form_padding)

        # 获取响应式字体和尺寸配置
        label_font_size = responsive_theme_manager.get_font('heading')[1]
        input_font_size = responsive_theme_manager.get_font('input')[1]
        input_height = responsive_theme_manager.get_size('input_height')
        input_spacing = responsive_theme_manager.get_responsive_config().get('input_spacing', 15)

        # 用户名输入
        username_label = ctk.CTkLabel(
            form_content,
            text="用户名",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=label_font_size, weight="bold"),
            anchor="w"
        )
        username_label.pack(fill="x", pady=(0, 5))

        self.username_entry = ctk.CTkEntry(
            form_content,
            placeholder_text="请输入用户名",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=input_font_size),
            height=input_height,
            corner_radius=10
        )
        self.username_entry.pack(fill="x", pady=(0, input_spacing))

        # 密码输入
        password_label = ctk.CTkLabel(
            form_content,
            text="密码",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=label_font_size, weight="bold"),
            anchor="w"
        )
        password_label.pack(fill="x", pady=(0, 5))

        self.password_entry = ctk.CTkEntry(
            form_content,
            placeholder_text="请输入密码",
            show="●",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=input_font_size),
            height=input_height,
            corner_radius=10
        )
        self.password_entry.pack(fill="x", pady=(0, input_spacing * 2))

        # 登录按钮
        button_font_size = responsive_theme_manager.get_font('button')[1]
        button_height = responsive_theme_manager.get_size('button_height')

        self.login_btn = ctk.CTkButton(
            form_content,
            text="登 录",
            command=self.login,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=button_font_size, weight="bold"),
            height=button_height,
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
        """创建响应式底部信息"""
        # 根据屏幕尺寸决定是否显示页脚
        show_footer = responsive_theme_manager.get_responsive_config().get('show_footer', True)

        if show_footer:
            footer_frame = ctk.CTkFrame(parent, fg_color="transparent")
            footer_frame.pack(fill="x")

            # 版权信息
            footer_font_size = responsive_theme_manager.get_responsive_config().get('footer_font', 12)
            copyright_label = ctk.CTkLabel(
                footer_frame,
                text="© 2025 工资条管理系统",
                font=ctk.CTkFont(family="Microsoft YaHei UI", size=footer_font_size),
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
