# coding:utf-8
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import os
import sys
from PIL import Image

# CustomTkinter主题将由主题管理器设置

from salary_mail.db_instance import set_db
from salary_mail.ctk_login_window import CTKLoginWindow
from salary_mail.ctk_theme_manager import theme_manager as ctk_theme_manager, CTKThemeSelector
from salary_mail.ctk_components import CTKWindowSizeManager
from salary_mail.theme_config import theme_manager as responsive_theme_manager


class CTKSettingsMenu(ctk.CTkToplevel):
    """系统设置菜单"""
    
    def __init__(self, parent):
        super().__init__(parent)
        
        self.title('系统设置')
        
        # 设置响应式窗口配置（但不立即居中）
        self.responsive_config = responsive_theme_manager.get_responsive_config()
        window_config = self.responsive_config.get('dialog_window', self.responsive_config['main_window'])
        
        # 设置窗口尺寸但不设置位置
        width = window_config['width']
        height = window_config['height']
        self.geometry(f'{width}x{height}')
        
        if 'min_width' in window_config and 'min_height' in window_config:
            self.minsize(window_config['min_width'], window_config['min_height'])
            
        self.resizable(True, True)
        
        # 设置窗口属性
        self.transient(parent)
        self.grab_set()
        
        self.parent = parent
        
        self.setup_ui()
        
        # 延迟居中显示在父窗口中心
        self.after(100, lambda: responsive_theme_manager.center_window_on_parent(self, parent))
        
        # 设置焦点
        self.focus_force()
    

    
    def setup_ui(self):
        """设置响应式UI"""
        # 获取响应式配置
        padding = responsive_theme_manager.get_size('padding_large')
        title_font_size = responsive_theme_manager.get_font('title')[1]
        section_spacing = responsive_theme_manager.get_size('margin_large')
        
        # 主容器
        main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=padding, pady=padding)
        
        # 标题
        title_label = ctk.CTkLabel(
            main_frame,
            text="系统设置",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=title_font_size, weight="bold")
        )
        title_label.pack(pady=(0, section_spacing))
        
        # 设置选项
        show_card_effects = responsive_theme_manager.get_responsive_config().get('card_effects', True)
        corner_radius = 10 if show_card_effects else 5
        options_frame = ctk.CTkFrame(main_frame, corner_radius=corner_radius)
        options_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        # 选项内容
        options_content = ctk.CTkFrame(options_frame, fg_color="transparent")
        options_content.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 获取响应式按钮配置
        button_height = responsive_theme_manager.get_size('button_height')
        button_font_size = responsive_theme_manager.get_font('button')[1]
        button_spacing = responsive_theme_manager.get_responsive_config().get('button_spacing', 8)
        
        # 邮箱设置按钮
        email_btn = ctk.CTkButton(
            options_content,
            text="📧 邮箱设置",
            command=self.open_email_setting,
            height=button_height,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=button_font_size, weight="bold"),
            anchor="w"
        )
        email_btn.pack(fill="x", pady=(0, button_spacing))
        
        # 模板设置按钮
        template_btn = ctk.CTkButton(
            options_content,
            text="📄 模板设置",
            command=self.open_template_setting,
            height=button_height,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=button_font_size, weight="bold"),
            fg_color="#9C27B0",
            hover_color="#7B1FA2",
            anchor="w"
        )
        template_btn.pack(fill="x", pady=(0, button_spacing))
        
        # 信息管理按钮
        info_btn = ctk.CTkButton(
            options_content,
            text="🏢 信息管理",
            command=self.open_info_manage,
            height=button_height,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=button_font_size, weight="bold"),
            fg_color="#795548",
            hover_color="#5D4037",
            anchor="w"
        )
        info_btn.pack(fill="x", pady=(0, button_spacing))
        
        # 主题设置按钮
        theme_btn = ctk.CTkButton(
            options_content,
            text="🎨 主题设置",
            command=self.open_theme_setting,
            height=button_height,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=button_font_size, weight="bold"),
            fg_color="#E91E63",
            hover_color="#C2185B",
            anchor="w"
        )
        theme_btn.pack(fill="x")
        
        # 关闭按钮
        close_btn = ctk.CTkButton(
            main_frame,
            text="关闭",
            command=self.destroy,
            width=100,
            height=35,
            fg_color="gray",
            hover_color="darkgray"
        )
        close_btn.pack(pady=(10, 0))
    
    def open_email_setting(self):
        """打开邮箱设置"""
        self.destroy()
        self.parent.show_email_setting()
    
    def open_template_setting(self):
        """打开模板设置"""
        self.destroy()
        self.parent.show_template_setting()
    
    def open_info_manage(self):
        """打开信息管理"""
        self.destroy()
        self.parent.show_info_manage()
    
    def open_theme_setting(self):
        """打开主题设置"""
        self.destroy()
        theme_selector = CTKThemeSelector(self.parent, ctk_theme_manager)
        theme_selector.mainloop()
    
    def show(self):
        """显示菜单"""
        self.wait_window()

class CTKHomePage(ctk.CTk):
    def __init__(self):
        # 先登录
        db = set_db()
        login_window = CTKLoginWindow(db)
        login_window.mainloop()
        
        # 如果登录窗口被关闭但没有登录成功，退出程序
        if not hasattr(login_window, 'current_user'):
            sys.exit()
        
        # 保存当前用户
        self.current_user = login_window.current_user
        self.db = db
        
        super().__init__()
        self.title('工资条管理系统')
        
        # 初始化延迟任务列表
        self.pending_tasks = []
        
        # 设置响应式窗口配置
        self.responsive_config = responsive_theme_manager.setup_responsive_window(self, 'main_window')
        
        # 在设置UI之前立即设置全屏，避免先显示小窗口
        self.set_fullscreen_immediately()
        
        # 加载图标
        self.load_icons()
        
        # 设置UI
        self.setup_ui()
        
        # 绑定快捷键
        self.bind('<F11>', self.toggle_fullscreen)
        self.bind('<Escape>', self.exit_fullscreen)
        
        # 延迟确认全屏状态，确保UI加载完成后仍然是全屏
        self.schedule_task(100, self.verify_fullscreen_after_ui_load)
        
        # 修改窗口状态追踪
        self.open_windows = {
            'email': None,
            'info': None,
            'salary': None,
            'employee': None
        }
        
        # 全屏状态追踪
        self.is_fullscreen = False
    
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
    
    def set_fullscreen_immediately(self):
        """立即设置全屏，避免先显示小窗口"""
        try:
            # 首先隐藏窗口，避免用户看到调整过程
            self.withdraw()
            
            # 强制更新窗口状态
            self.update_idletasks()
            
            # 方法1：直接使用Windows的zoomed状态（最可靠）
            self.state('zoomed')
            self.is_fullscreen = True
            
            # 显示窗口
            self.deiconify()
            
        except Exception as e:
            print(f"zoomed状态失败: {e}")
            try:
                # 方法2：手动设置全屏几何
                # 先确保窗口已初始化
                self.update_idletasks()
                
                screen_width = self.winfo_screenwidth()
                screen_height = self.winfo_screenheight()
                
                print(f"屏幕尺寸: {screen_width}x{screen_height}")
                
                # 设置窗口几何为全屏尺寸，确保位置在左上角(0,0)
                self.geometry(f"{screen_width}x{screen_height}+0+0")
                self.is_fullscreen = True
                
                print(f"手动设置全屏: {screen_width}x{screen_height}+0+0")
                
                # 显示窗口
                self.deiconify()
                
            except Exception as e2:
                print(f"手动设置全屏也失败: {e2}")
                # 最后备用方案
                self.geometry("1400x900+0+0")
                self.is_fullscreen = False
                self.deiconify()
    
    def verify_fullscreen_after_ui_load(self):
        """在UI加载完成后验证全屏状态"""
        try:
            # 检查窗口是否仍然存在
            if not self.winfo_exists():
                return
                
            # 如果当前不是全屏状态，重新设置
            if not self.is_fullscreen or self.state() != 'zoomed':
                self.state('zoomed')
                self.is_fullscreen = True
                
                # 再次延迟检查
                self.schedule_task(200, self.final_fullscreen_check)
                
        except Exception as e:
            print(f"验证全屏状态失败: {e}")
    
    def final_fullscreen_check(self):
        """最终的全屏状态检查"""
        try:
            # 检查窗口是否仍然存在
            if not self.winfo_exists():
                return
                
            current_width = self.winfo_width()
            current_height = self.winfo_height()
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            
            # 如果窗口大小明显小于屏幕大小，使用备用方案
            if current_width < screen_width - 50 or current_height < screen_height - 50:
                print(f"最终检查：窗口大小 {current_width}x{current_height}，屏幕大小 {screen_width}x{screen_height}")
                print("使用几何设置强制全屏")
                self.geometry(f"{screen_width}x{screen_height}+0+0")
                
        except Exception as e:
            print(f"最终全屏检查失败: {e}")
    
    def set_fullscreen(self):
        """设置窗口为全屏"""
        try:
            # 确保窗口已经显示
            self.update_idletasks()
            
            # 获取屏幕尺寸
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            
            # 方法1：使用zoomed状态（Windows推荐）
            self.state('zoomed')
            self.is_fullscreen = True
            
            # 延迟确认全屏状态
            self.schedule_task(50, self.confirm_fullscreen)
            
        except Exception as e:
            print(f"设置全屏失败: {e}")
            # 备用方案：手动设置窗口大小和位置
            self.fallback_fullscreen()
    
    def confirm_fullscreen(self):
        """确认全屏状态，如果不是全屏则使用备用方案"""
        try:
            # 检查窗口是否真的全屏了
            current_width = self.winfo_width()
            current_height = self.winfo_height()
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            
            # 如果窗口大小明显小于屏幕大小，说明全屏失败
            if current_width < screen_width - 100 or current_height < screen_height - 100:
                print("检测到全屏未生效，使用备用方案")
                self.fallback_fullscreen()
                
        except Exception as e:
            print(f"确认全屏状态失败: {e}")
            self.fallback_fullscreen()
    
    def fallback_fullscreen(self):
        """备用全屏方案"""
        try:
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            
            # 设置窗口几何为全屏尺寸
            self.geometry(f"{screen_width}x{screen_height}+0+0")
            
            # 尝试移除窗口装饰（标题栏等）
            self.overrideredirect(True)
            self.overrideredirect(False)  # 立即恢复，只是为了刷新状态
            
            # 再次尝试zoomed状态
            self.state('zoomed')
            
        except Exception as e:
            print(f"备用全屏方案失败: {e}")
            # 最后的备用方案：设置为接近全屏的大窗口
            try:
                screen_width = self.winfo_screenwidth()
                screen_height = self.winfo_screenheight()
                self.geometry(f"{screen_width-50}x{screen_height-50}+0+0")
            except:
                self.geometry("1400x900+0+0")
    
    def toggle_fullscreen(self, event=None):
        """切换全屏模式（F11快捷键）"""
        try:
            if self.is_fullscreen:
                # 当前是全屏，切换到窗口模式
                self.exit_fullscreen()
            else:
                # 当前是窗口模式，切换到全屏
                self.set_fullscreen()
        except Exception as e:
            print(f"切换全屏模式失败: {e}")
    
    def exit_fullscreen(self, event=None):
        """退出全屏模式（ESC快捷键）"""
        try:
            self.state('normal')
            self.geometry("1200x800+100+50")
            self.is_fullscreen = False
        except Exception as e:
            print(f"退出全屏模式失败: {e}")
    
    def center_window(self):
        """窗口居中显示"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def load_icons(self):
        """加载图标"""
        self.icons = {}
        icon_names = ['employee', 'salary', 'settings', 'template']
        
        for name in icon_names:
            try:
                icon_path = os.path.join(os.path.dirname(__file__), 'assets', f'{name}.png')
                if os.path.exists(icon_path):
                    # CustomTkinter使用PIL Image
                    image = Image.open(icon_path)
                    # 调整图标大小
                    image = image.resize((32, 32), Image.Resampling.LANCZOS)
                    
                    # 为深色模式创建适配的图标
                    if name == 'employee':
                        # 员工管理图标保持原样（已经是灰色）
                        self.icons[name] = ctk.CTkImage(light_image=image, dark_image=image, size=(32, 32))
                    else:
                        # 其他图标为深色模式创建白色版本
                        dark_image = self.create_dark_mode_icon(image)
                        self.icons[name] = ctk.CTkImage(light_image=image, dark_image=dark_image, size=(32, 32))
                else:
                    self.icons[name] = None
            except Exception as e:
                print(f"加载图标失败 {name}: {e}")
                self.icons[name] = None
    
    def create_dark_mode_icon(self, original_image):
        """为深色模式创建白色图标"""
        try:
            # 转换为RGBA模式
            if original_image.mode != 'RGBA':
                original_image = original_image.convert('RGBA')
            
            # 创建白色图标
            white_icon = Image.new('RGBA', original_image.size, (255, 255, 255, 0))
            
            # 获取原图的数据
            data = original_image.getdata()
            new_data = []
            
            for item in data:
                # 如果像素不是完全透明的，则设为白色
                if item[3] > 0:  # alpha > 0
                    new_data.append((255, 255, 255, item[3]))
                else:
                    new_data.append(item)
            
            white_icon.putdata(new_data)
            return white_icon
            
        except Exception as e:
            print(f"创建深色模式图标失败: {e}")
            return original_image
    
    def setup_ui(self):
        """设置响应式用户界面"""
        # 获取响应式配置
        container_padding = responsive_theme_manager.get_size('padding_xl')
        
        # 主容器框架 - 响应式内边距
        main_container = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=container_padding, pady=container_padding)
        
        # 顶部标题区域
        self.create_header(main_container)
        
        # 功能卡片区域
        self.create_cards_section(main_container)
        
        # 底部状态栏
        self.create_status_bar(main_container)
    
    def create_header(self, parent):
        """创建响应式顶部标题区域"""
        # 获取响应式配置
        title_font_size = responsive_theme_manager.get_font('title')[1] +8  # 主页面标题更大
        subtitle_font_size = responsive_theme_manager.get_font('subtitle')[1]
        section_spacing = responsive_theme_manager.get_size('margin_xl')
        show_subtitle = responsive_theme_manager.get_responsive_config().get('show_subtitle', True)
        show_card_effects = responsive_theme_manager.get_responsive_config().get('card_effects', True)
        
        # 动态调整header高度
        header_height = 120 if show_subtitle else 80
        corner_radius = 20 if show_card_effects else 10
        
        header_frame = ctk.CTkFrame(parent, height=header_height, corner_radius=corner_radius)
        header_frame.pack(fill="x", pady=(0, section_spacing))
        header_frame.pack_propagate(False)
        
        # 标题
        title_label = ctk.CTkLabel(
            header_frame,
            text="工资条管理系统",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=title_font_size, weight="bold"),
            text_color=("#1976d2", "#64b5f6")  # 浅色模式和深色模式的颜色
        )
        title_label.pack(pady=(20, 8))
        
        # 副标题（根据屏幕尺寸决定是否显示）
        if show_subtitle:
            subtitle_label = ctk.CTkLabel(
                header_frame,
                text="现代化的工资条发送管理工具",
                font=ctk.CTkFont(family="Microsoft YaHei UI", size=subtitle_font_size),
                text_color=("#666666", "#aaaaaa")
            )
            subtitle_label.pack(pady=(0, 20))
    
    def create_cards_section(self, parent):
        """创建功能卡片区域"""
        cards_frame = ctk.CTkFrame(parent, fg_color="transparent")
        cards_frame.pack(fill="both", expand=True, pady=(0, 30))
        
        # 配置网格 - 全屏模式下增加权重
        cards_frame.grid_columnconfigure((0, 1), weight=1)
        cards_frame.grid_rowconfigure((0, 1), weight=1)
        
        # 员工管理卡片
        self.create_feature_card(
            cards_frame, 0, 0,
            "员工管理",
            "管理员工信息和邮箱\n支持批量导入和实时搜索",
            self.icons.get('employee'),
            self.show_employee_manage,
            "#4CAF50"
        )
        
        # 工资管理卡片
        self.create_feature_card(
            cards_frame, 0, 1,
            "工资管理",
            "导入工资数据并发送工资条\n支持批量发送和状态跟踪",
            self.icons.get('salary'),
            self.show_salary_manage,
            "#2196F3"
        )
        
        # 邮件设置卡片
        self.create_feature_card(
            cards_frame, 1, 0,
            "邮件设置",
            "配置SMTP邮箱参数\n设置发送邮箱和服务器",
            self.icons.get('settings'),
            self.show_email_setting,
            "#FF9800"
        )
        
        # 系统设置卡片
        self.create_feature_card(
            cards_frame, 1, 1,
            "系统设置",
            "模板设置和信息管理\n密码修改和公司信息",
            self.icons.get('template'),
            self.show_system_settings,
            "#9C27B0"
        )
    
    def create_feature_card(self, parent, row, col, title, description, icon, command, color):
        """创建功能卡片"""
        # 卡片框架 - 全屏模式下增加大小
        card = ctk.CTkFrame(
            parent,
            corner_radius=25,
            fg_color=("white", "#2b2b2b"),
            border_width=2,
            border_color=("#e0e0e0", "#404040")
        )
        card.grid(row=row, column=col, padx=30, pady=30, sticky="nsew")
        
        # 卡片内容容器 - 增加内边距
        content_frame = ctk.CTkFrame(card, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=30, pady=30)
        
        # 图标和标题行
        header_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 20))
        
        # 图标 - 全屏模式下增大图标
        if icon:
            icon_label = ctk.CTkLabel(header_frame, image=icon, text="")
            icon_label.pack(side="left", padx=(0, 20))
        
        # 标题 - 增大字体
        title_label = ctk.CTkLabel(
            header_frame,
            text=title,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=24, weight="bold"),
            text_color=color
        )
        title_label.pack(side="left", anchor="w")
        
        # 描述文本 - 增大字体
        desc_label = ctk.CTkLabel(
            content_frame,
            text=description,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=16),
            text_color=("#666666", "#aaaaaa"),
            justify="left"
        )
        desc_label.pack(fill="x", pady=(0, 25))
        
        # 操作按钮 - 增大按钮
        action_btn = ctk.CTkButton(
            content_frame,
            text=f"打开{title}",
            command=command,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=16, weight="bold"),
            fg_color=color,
            corner_radius=15,
            height=50
        )
        action_btn.pack(fill="x")
        
        # 添加悬停效果
        def on_enter(event):
            card.configure(border_width=2, border_color=color)
        
        def on_leave(event):
            card.configure(border_width=0)
        
        card.bind("<Enter>", on_enter)
        card.bind("<Leave>", on_leave)
        
        # 为所有子组件绑定事件
        for widget in [content_frame, header_frame, title_label, desc_label]:
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)
    
    def create_status_bar(self, parent):
        """创建底部状态栏"""
        status_frame = ctk.CTkFrame(parent, height=60, corner_radius=15)
        status_frame.pack(fill="x")
        status_frame.pack_propagate(False)
        
        # 版权信息
        copyright_label = ctk.CTkLabel(
            status_frame,
            text="© 2025 工资条管理系统",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color=("#666666", "#aaaaaa")
        )
        copyright_label.pack(side="left", padx=20, pady=15)
        
        # 状态信息
        self.status_label = ctk.CTkLabel(
            status_frame,
            text="系统就绪",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color=("#4CAF50", "#81C784")
        )
        self.status_label.pack(side="right", padx=20, pady=15)
    
    def darken_color(self, color):
        """使颜色变暗用于悬停效果"""
        color_map = {
            "#4CAF50": "#45a049",
            "#2196F3": "#1976D2",
            "#FF9800": "#F57C00",
            "#9C27B0": "#7B1FA2"
        }
        return color_map.get(color, color)
    
    def get_center(self):
        """获取窗口中心坐标，用于子窗口定位"""
        px = self.winfo_x()
        py = self.winfo_y()
        pw = self.winfo_width()
        ph = self.winfo_height()
        return (int(px + pw/2), int(py + ph/2))
    
    def update_status(self, text, color='#4CAF50'):
        """更新状态栏文本"""
        self.status_label.configure(text=text, text_color=color)
    
    # 功能窗口显示方法
    def show_employee_manage(self):
        """显示员工管理窗口"""
        try:
            if self.open_windows.get('employee') and self.open_windows['employee'].winfo_exists():
                self.open_windows['employee'].focus_force()
                return
            from salary_mail.ctk_employee_manage import CTKEmployeeManageWin
            dialog = CTKEmployeeManageWin(parent=self)
            self.open_windows['employee'] = dialog
            dialog.protocol("WM_DELETE_WINDOW", lambda: self._on_dialog_close('employee'))
        except Exception as e:
            print(f"打开员工管理窗口失败: {e}")
            self.open_windows['employee'] = None
    
    def show_salary_manage(self):
        """显示工资管理窗口"""
        try:
            if self.open_windows.get('salary') and self.open_windows['salary'].winfo_exists():
                self.open_windows['salary'].focus_force()
                return
            from salary_mail.ctk_salary_manage import CTKSalaryManageWin
            dialog = CTKSalaryManageWin(parent=self)
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
            from salary_mail.ctk_settings import CTKEmailSettingWin
            dialog = CTKEmailSettingWin(parent=self)
            self.open_windows['email'] = dialog
            dialog.protocol("WM_DELETE_WINDOW", lambda: self._on_dialog_close('email'))
        except Exception as e:
            print(f"打开邮箱设置窗口失败: {e}")
            self.open_windows['email'] = None
    
    def show_system_settings(self):
        """显示系统设置菜单"""
        # 创建设置菜单
        settings_menu = CTKSettingsMenu(self)
        settings_menu.show()
    
    def show_template_setting(self):
        """显示模板设置窗口"""
        try:
            from salary_mail.ctk_settings import CTKTemplateSettingWin
            dialog = CTKTemplateSettingWin(parent=self)
            self.wait_window(dialog)
        except Exception as e:
            print(f"打开模板设置窗口失败: {e}")
    
    def show_info_manage(self):
        """显示信息管理窗口"""
        try:
            if self.open_windows.get('info') and self.open_windows['info'].winfo_exists():
                self.open_windows['info'].focus_force()
                return
            from salary_mail.ctk_settings import CTKInfoManageWin
            dialog = CTKInfoManageWin(parent=self)
            self.open_windows['info'] = dialog
            dialog.protocol("WM_DELETE_WINDOW", lambda: self._on_dialog_close('info'))
        except Exception as e:
            print(f"打开信息管理窗口失败: {e}")
            self.open_windows['info'] = None
    
    def _on_dialog_close(self, window_key):
        """处理窗口关闭"""
        try:
            if self.open_windows.get(window_key):
                self.open_windows[window_key].destroy()
            self.open_windows[window_key] = None
        except Exception as e:
            print(f"关闭窗口失败: {e}")
            self.open_windows[window_key] = None


if __name__ == '__main__':
    app = CTKHomePage()
    app.mainloop()
