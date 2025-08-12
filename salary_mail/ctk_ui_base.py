# coding:utf-8
"""
CustomTkinter UI基础组件
提供现代化的UI样式和通用功能
"""
import customtkinter as ctk
import tkinter as tk
from typing import Optional, Tuple, Dict, Any


class ModernColors:
    """现代化配色方案"""
    
    # 主色调
    PRIMARY = "#1976D2"
    PRIMARY_HOVER = "#1565C0"
    PRIMARY_LIGHT = "#BBDEFB"
    
    # 辅助色
    SUCCESS = "#4CAF50"
    SUCCESS_HOVER = "#45A049"
    WARNING = "#FF9800"
    WARNING_HOVER = "#F57C00"
    ERROR = "#F44336"
    ERROR_HOVER = "#D32F2F"
    INFO = "#2196F3"
    INFO_HOVER = "#1976D2"
    
    # 中性色
    GRAY_50 = "#FAFAFA"
    GRAY_100 = "#F5F5F5"
    GRAY_200 = "#EEEEEE"
    GRAY_300 = "#E0E0E0"
    GRAY_400 = "#BDBDBD"
    GRAY_500 = "#9E9E9E"
    GRAY_600 = "#757575"
    GRAY_700 = "#616161"
    GRAY_800 = "#424242"
    GRAY_900 = "#212121"
    
    # 文本色
    TEXT_PRIMARY = "#212121"
    TEXT_SECONDARY = "#757575"
    TEXT_DISABLED = "#BDBDBD"


class ModernFonts:
    """现代化字体配置"""
    
    FAMILY = "Microsoft YaHei UI"
    
    # 字体大小
    SIZE_SMALL = 11
    SIZE_NORMAL = 12
    SIZE_MEDIUM = 14
    SIZE_LARGE = 16
    SIZE_XL = 18
    SIZE_XXL = 20
    SIZE_TITLE = 24
    SIZE_HERO = 28
    
    @classmethod
    def get(cls, size: int = SIZE_NORMAL, weight: str = "normal") -> ctk.CTkFont:
        """获取字体"""
        return ctk.CTkFont(family=cls.FAMILY, size=size, weight=weight)


class ModernSpacing:
    """现代化间距系统"""
    
    XS = 4
    SM = 8
    MD = 12
    LG = 16
    XL = 20
    XXL = 24
    XXXL = 32


class ModernWindow:
    """现代化窗口基类"""
    
    @staticmethod
    def center_on_screen(window, width: int, height: int):
        """在屏幕中心显示窗口"""
        window.update_idletasks()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")
    
    @staticmethod
    def center_on_parent(window, parent):
        """在父窗口中心显示"""
        window.update_idletasks()
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        window_width = window.winfo_width()
        window_height = window.winfo_height()
        
        x = parent_x + (parent_width - window_width) // 2
        y = parent_y + (parent_height - window_height) // 2
        
        window.geometry(f"+{x}+{y}")
    
    @staticmethod
    def setup_window_properties(window, parent=None, modal=True):
        """设置窗口属性"""
        if parent:
            window.transient(parent)
        if modal:
            window.grab_set()
        window.focus_force()


class ModernCard(ctk.CTkFrame):
    """现代化卡片组件"""
    
    def __init__(self, parent, title: str = None, **kwargs):
        # 默认样式
        default_kwargs = {
            'corner_radius': 12,
            'border_width': 1,
            'border_color': ModernColors.GRAY_200
        }
        default_kwargs.update(kwargs)
        
        super().__init__(parent, **default_kwargs)
        
        if title:
            self.create_title(title)
    
    def create_title(self, title: str):
        """创建卡片标题"""
        title_frame = ctk.CTkFrame(self, fg_color="transparent")
        title_frame.pack(fill="x", padx=ModernSpacing.LG, pady=(ModernSpacing.LG, ModernSpacing.SM))
        
        title_label = ctk.CTkLabel(
            title_frame,
            text=title,
            font=ModernFonts.get(ModernFonts.SIZE_MEDIUM, "bold"),
            anchor="w"
        )
        title_label.pack(fill="x")
        
        return title_frame


class ModernButton(ctk.CTkButton):
    """现代化按钮组件"""
    
    def __init__(self, parent, variant: str = "primary", size: str = "medium", **kwargs):
        # 按钮变体样式
        variants = {
            "primary": {
                "fg_color": ModernColors.PRIMARY,
                "hover_color": ModernColors.PRIMARY_HOVER,
                "text_color": "white"
            },
            "secondary": {
                "fg_color": "transparent",
                "hover_color": ModernColors.GRAY_100,
                "border_width": 1,
                "border_color": ModernColors.GRAY_300,
                "text_color": ModernColors.TEXT_PRIMARY
            },
            "success": {
                "fg_color": ModernColors.SUCCESS,
                "hover_color": ModernColors.SUCCESS_HOVER,
                "text_color": "white"
            },
            "warning": {
                "fg_color": ModernColors.WARNING,
                "hover_color": ModernColors.WARNING_HOVER,
                "text_color": "white"
            },
            "error": {
                "fg_color": ModernColors.ERROR,
                "hover_color": ModernColors.ERROR_HOVER,
                "text_color": "white"
            }
        }
        
        # 按钮尺寸
        sizes = {
            "small": {"height": 32, "font": ModernFonts.get(ModernFonts.SIZE_SMALL)},
            "medium": {"height": 40, "font": ModernFonts.get(ModernFonts.SIZE_NORMAL)},
            "large": {"height": 48, "font": ModernFonts.get(ModernFonts.SIZE_MEDIUM)}
        }
        
        # 合并样式
        style = variants.get(variant, variants["primary"])
        size_style = sizes.get(size, sizes["medium"])
        
        default_kwargs = {
            "corner_radius": 8,
            **style,
            **size_style
        }
        default_kwargs.update(kwargs)
        
        super().__init__(parent, **default_kwargs)


class ModernInput(ctk.CTkEntry):
    """现代化输入框组件"""
    
    def __init__(self, parent, label: str = None, **kwargs):
        # 默认样式
        default_kwargs = {
            "height": 40,
            "corner_radius": 8,
            "border_width": 1,
            "font": ModernFonts.get()
        }
        default_kwargs.update(kwargs)
        
        # 如果有标签，创建容器
        if label:
            container = ctk.CTkFrame(parent, fg_color="transparent")
            container.pack(fill="x", pady=ModernSpacing.SM)
            
            # 标签
            label_widget = ctk.CTkLabel(
                container,
                text=label,
                font=ModernFonts.get(ModernFonts.SIZE_SMALL, "bold"),
                anchor="w"
            )
            label_widget.pack(fill="x", pady=(0, ModernSpacing.XS))
            
            super().__init__(container, **default_kwargs)
        else:
            super().__init__(parent, **default_kwargs)


class ModernScrollableFrame(ctk.CTkScrollableFrame):
    """现代化滚动框架"""
    
    def __init__(self, parent, **kwargs):
        default_kwargs = {
            "corner_radius": 12,
            "scrollbar_button_color": ModernColors.GRAY_300,
            "scrollbar_button_hover_color": ModernColors.GRAY_400
        }
        default_kwargs.update(kwargs)
        
        super().__init__(parent, **default_kwargs)


class ModernDialog(ctk.CTkToplevel):
    """现代化对话框基类"""
    
    def __init__(self, parent, title: str, width: int = 400, height: int = 300):
        super().__init__(parent)
        
        self.title(title)
        self.geometry(f"{width}x{height}")
        self.resizable(False, False)
        
        # 设置窗口属性
        ModernWindow.setup_window_properties(self, parent)
        ModernWindow.center_on_parent(self, parent)
        
        self.result = None
        
        # 主容器
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=ModernSpacing.XL, pady=ModernSpacing.XL)
        
        self.setup_ui()
    
    def setup_ui(self):
        """子类重写此方法"""
        pass
    
    def show(self):
        """显示对话框并返回结果"""
        self.wait_window()
        return self.result


class ModernProgressDialog(ModernDialog):
    """现代化进度对话框"""
    
    def __init__(self, parent, title: str = "处理中...", message: str = "请稍候..."):
        self.message = message
        super().__init__(parent, title, 350, 150)
    
    def setup_ui(self):
        """设置UI"""
        # 消息
        message_label = ctk.CTkLabel(
            self.main_frame,
            text=self.message,
            font=ModernFonts.get(ModernFonts.SIZE_MEDIUM)
        )
        message_label.pack(pady=(0, ModernSpacing.LG))
        
        # 进度条
        self.progress_bar = ctk.CTkProgressBar(
            self.main_frame,
            height=8,
            corner_radius=4
        )
        self.progress_bar.pack(fill="x", pady=ModernSpacing.SM)
        self.progress_bar.set(0)
        
        # 状态标签
        self.status_label = ctk.CTkLabel(
            self.main_frame,
            text="准备中...",
            font=ModernFonts.get(ModernFonts.SIZE_SMALL),
            text_color=ModernColors.TEXT_SECONDARY
        )
        self.status_label.pack()
    
    def update_progress(self, value: float, status: str = None):
        """更新进度"""
        self.progress_bar.set(value)
        if status:
            self.status_label.configure(text=status)
        self.update()
