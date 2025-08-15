# coding:utf-8
"""
CustomTkinter 主题管理器
提供主题切换和自定义主题功能
"""
import customtkinter as ctk
import tkinter as tk
import json
import os
from typing import Dict, Tuple

class CTKThemeManager:
    """CustomTkinter主题管理器"""
    
    def __init__(self):
        self.config_file = "theme_config.json"
        
        # 主题配置 - 先定义themes
        self.themes = {
            "light": {
                "appearance_mode": "light",
                "color_theme": "blue",
                "colors": {
                    "primary": "#1976d2",
                    "primary_hover": "#1565c0",
                    "success": "#4CAF50",
                    "success_hover": "#45a049",
                    "warning": "#FF9800",
                    "warning_hover": "#F57C00",
                    "error": "#F44336",
                    "error_hover": "#D32F2F",
                    "info": "#2196F3",
                    "info_hover": "#1976D2"
                }
            },
            "dark": {
                "appearance_mode": "dark",
                "color_theme": "dark-blue",
                "colors": {
                    "primary": "#64b5f6",
                    "primary_hover": "#42a5f5",
                    "success": "#66BB6A",
                    "success_hover": "#4CAF50",
                    "warning": "#FFB74D",
                    "warning_hover": "#FF9800",
                    "error": "#EF5350",
                    "error_hover": "#F44336",
                    "info": "#64B5F6",
                    "info_hover": "#2196F3"
                }
            },
            "green": {
                "appearance_mode": "light",
                "color_theme": "green",
                "colors": {
                    "primary": "#4CAF50",
                    "primary_hover": "#45a049",
                    "success": "#4CAF50",
                    "success_hover": "#45a049",
                    "warning": "#FF9800",
                    "warning_hover": "#F57C00",
                    "error": "#F44336",
                    "error_hover": "#D32F2F",
                    "info": "#2196F3",
                    "info_hover": "#1976D2"
                }
            }
        }
        
        # 加载并应用主题
        self.current_theme = self.load_theme_config()
        self.apply_theme()
    
    def load_theme_config(self) -> str:
        """加载主题配置"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    return config.get('theme', 'light')
        except Exception as e:
            print(f"加载主题配置失败: {e}")
        return 'light'
    
    def save_theme_config(self, theme: str):
        """保存主题配置"""
        try:
            config = {'theme': theme}
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存主题配置失败: {e}")
    
    def apply_theme(self, theme: str = None):
        """应用主题"""
        if theme is None:
            theme = self.current_theme
        
        if theme in self.themes:
            theme_config = self.themes[theme]
            ctk.set_appearance_mode(theme_config["appearance_mode"])
            ctk.set_default_color_theme(theme_config["color_theme"])
            self.current_theme = theme
            self.save_theme_config(theme)
    
    def get_theme_colors(self, theme: str = None) -> Dict[str, str]:
        """获取主题颜色"""
        if theme is None:
            theme = self.current_theme
        
        if theme in self.themes:
            return self.themes[theme]["colors"]
        return self.themes["light"]["colors"]
    
    def get_current_theme(self) -> str:
        """获取当前主题"""
        return self.current_theme
    
    def get_available_themes(self) -> list:
        """获取可用主题列表"""
        return list(self.themes.keys())
    
    def switch_theme(self, theme: str):
        """切换主题"""
        if theme in self.themes:
            self.apply_theme(theme)
            return True
        return False


class CTKThemeSelector(ctk.CTkToplevel):
    """主题选择器窗口"""
    
    def __init__(self, parent, theme_manager: CTKThemeManager):
        super().__init__(parent)
        
        self.title('主题设置')
        # 导入窗口大小管理器
        from salary_mail.ctk_components import CTKWindowSizeManager
        CTKWindowSizeManager.adjust_window_size(self, 450, 480, 400, 400)
        self.resizable(True, True)
        
        # 设置窗口属性
        self.transient(parent)
        self.grab_set()
        
        # 确保窗口在父窗口中心显示
        self.after(100, lambda: self._center_window_on_parent_with_theme_manager(parent))
        
        self.parent = parent
        self.theme_manager = theme_manager
        self.selected_theme = tk.StringVar(value=theme_manager.get_current_theme())
        
        self.setup_ui()
        
        # 设置焦点
        self.focus_force()
    
    def _center_window_on_parent_with_theme_manager(self, parent):
        """使用主题管理器在父窗口中心显示 - 支持高DPI缩放"""
        try:
            # 导入主题管理器（避免循环导入）
            from salary_mail.theme_config import theme_manager as responsive_theme_manager
            responsive_theme_manager.center_window_on_parent(self, parent)
        except Exception as e:
            print(f"主题选择器窗口居中失败: {e}")
    
    def setup_ui(self):
        """设置UI"""
        # 主容器
        main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 标题
        title_label = ctk.CTkLabel(
            main_frame,
            text="🎨 主题设置",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=18, weight="bold")
        )
        title_label.pack(pady=(0, 15))
        
        # 主题选择区域
        theme_frame = ctk.CTkFrame(main_frame, corner_radius=10)
        theme_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        theme_content = ctk.CTkFrame(theme_frame, fg_color="transparent")
        theme_content.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 主题选项
        themes = [
            ("浅色主题", "light", "☀️ 清新明亮的浅色界面"),
            ("深色主题", "dark", "🌙 护眼舒适的深色界面"),
            ("绿色主题", "green", "🌿 自然清新的绿色界面")
        ]
        
        for name, value, description in themes:
            self.create_theme_option(theme_content, name, value, description)
        
        # 预览区域
        preview_frame = ctk.CTkFrame(theme_content, corner_radius=8)
        preview_frame.pack(fill="x", pady=(8, 0))
        
        preview_label = ctk.CTkLabel(
            preview_frame,
            text="💡 提示：选择主题后点击应用即可预览效果",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            text_color="gray"
        )
        preview_label.pack(pady=12)
        
        # 按钮区域
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x", pady=(10, 0))
        
        # 取消按钮
        cancel_btn = ctk.CTkButton(
            button_frame,
            text="取消",
            command=self.cancel,
            width=100,
            height=35,
            fg_color="gray",
            hover_color="darkgray"
        )
        cancel_btn.pack(side="right", padx=5)
        
        # 应用按钮
        apply_btn = ctk.CTkButton(
            button_frame,
            text="应用",
            command=self.apply_theme,
            width=100,
            height=35
        )
        apply_btn.pack(side="right", padx=5)
    
    def create_theme_option(self, parent, name, value, description):
        """创建主题选项"""
        option_frame = ctk.CTkFrame(parent, corner_radius=8)
        option_frame.pack(fill="x", pady=5)
        
        # 单选按钮和标题
        header_frame = ctk.CTkFrame(option_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=15, pady=(10, 5))
        
        radio_btn = ctk.CTkRadioButton(
            header_frame,
            text=name,
            variable=self.selected_theme,
            value=value,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold")
        )
        radio_btn.pack(side="left")
        
        # 描述文本
        desc_label = ctk.CTkLabel(
            option_frame,
            text=description,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=11),
            text_color="gray"
        )
        desc_label.pack(anchor="w", padx=15, pady=(0, 10))
    
    def apply_theme(self):
        """应用选中的主题"""
        selected = self.selected_theme.get()
        if self.theme_manager.switch_theme(selected):
            # 主题切换成功，重新启动应用以完全应用新主题
            from salary_mail.ctk_components import CTKMessageBox
            result = CTKMessageBox.ask_yes_no(
                self, 
                '主题切换', 
                '主题已更改！\n为了完全应用新主题，建议重启应用程序。\n\n是否现在重启？'
            )
            if result == "是":
                self.restart_application()
            else:
                self.destroy()
        else:
            from salary_mail.ctk_components import CTKMessageBox
            CTKMessageBox.show_error(self, '错误', '主题切换失败！')
    
    def cancel(self):
        """取消设置"""
        self.destroy()
    
    def restart_application(self):
        """重启应用程序"""
        import sys
        import subprocess
        
        try:
            # 关闭当前应用
            self.destroy()
            if hasattr(self.parent, 'destroy'):
                self.parent.destroy()
            
            # 重新启动应用
            subprocess.Popen([sys.executable] + sys.argv)
            sys.exit(0)
        except Exception as e:
            print(f"重启应用失败: {e}")


# 全局主题管理器实例
theme_manager = CTKThemeManager()
