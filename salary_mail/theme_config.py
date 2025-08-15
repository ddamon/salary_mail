# -*- coding: utf-8 -*-
"""
全局主题配置系统
提供统一的颜色、字体和样式配置
集成响应式设计支持
"""

import tkinter as tk
from tkinter import ttk
from .responsive_utils import responsive

class ModernTheme:
    """现代化主题配置类"""

    # 颜色配置
    COLORS = {
        # 主色调
        'primary': '#1976d2',
        'primary_dark': '#1565c0',
        'primary_darker': '#0d47a1',

        # 辅助色
        'secondary': '#757575',
        'secondary_dark': '#616161',

        # 状态色
        'success': '#4caf50',
        'success_dark': '#388e3c',
        'warning': '#ff9800',
        'warning_dark': '#f57c00',
        'error': '#f44336',
        'error_dark': '#c62828',
        'info': '#2196f3',
        'info_dark': '#1565c0',

        # 中性色
        'white': '#ffffff',
        'light_gray': '#f8f9fa',
        'gray': '#e0e0e0',
        'dark_gray': '#757575',
        'black': '#333333',

        # 背景色
        'bg_primary': '#f8f9fa',
        'bg_card': '#ffffff',
        'bg_hover': '#fafafa',
        'bg_pressed': '#f5f5f5',
        'bg_selected': '#e3f2fd',

        # 文本色
        'text_primary': '#333333',
        'text_secondary': '#757575',
        'text_disabled': '#999999',
    }

    # 字体配置
    FONTS = {
        'default': ('Microsoft YaHei UI', 9),
        'small': ('Microsoft YaHei UI', 8),
        'medium': ('Microsoft YaHei UI', 10),
        'large': ('Microsoft YaHei UI', 12),
        'title': ('Microsoft YaHei UI', 16, 'bold'),
        'subtitle': ('Microsoft YaHei UI', 14),
        'heading': ('Microsoft YaHei UI', 12, 'bold'),
        'button': ('Microsoft YaHei UI', 9),
        'input': ('Microsoft YaHei UI', 10),
    }

    # 尺寸配置
    SIZES = {
        'padding_small': 5,
        'padding_medium': 10,
        'padding_large': 15,
        'padding_xl': 20,

        'margin_small': 5,
        'margin_medium': 10,
        'margin_large': 15,
        'margin_xl': 20,

        'border_radius': 4,
        'border_width': 1,

        'button_height': 32,
        'input_height': 36,
        'row_height': 28,
    }

    @classmethod
    def apply_modern_styles(cls, root):
        """应用现代化样式到根窗口"""
        style = ttk.Style()
        style.theme_use('clam')

        # 配置基础样式
        cls._configure_frame_styles(style)
        cls._configure_button_styles(style)
        cls._configure_entry_styles(style)
        cls._configure_treeview_styles(style)
        cls._configure_combobox_styles(style)

        return style

    @classmethod
    def _configure_frame_styles(cls, style):
        """配置框架样式"""
        style.configure('Modern.TFrame',
                       background=cls.COLORS['bg_card'],
                       borderwidth=0,
                       relief='flat')

        style.configure('Card.TFrame',
                       background=cls.COLORS['bg_card'],
                       relief='solid',
                       borderwidth=cls.SIZES['border_width'])

        style.configure('Toolbar.TFrame',
                       background=cls.COLORS['bg_card'],
                       relief='flat')

    @classmethod
    def _configure_button_styles(cls, style):
        """配置按钮样式"""
        style.configure('Modern.TButton',
                       font=cls.FONTS['button'],
                       padding=(cls.SIZES['padding_medium'], 6),
                       background=cls.COLORS['primary'],
                       foreground=cls.COLORS['white'],
                       borderwidth=0,
                       focuscolor='none')

        style.map('Modern.TButton',
                 background=[('active', cls.COLORS['primary_dark']),
                           ('pressed', cls.COLORS['primary_darker'])],
                 foreground=[('active', cls.COLORS['white']),
                           ('pressed', cls.COLORS['white'])])

    @classmethod
    def _configure_entry_styles(cls, style):
        """配置输入框样式"""
        style.configure('Modern.TEntry',
                       font=cls.FONTS['input'],
                       padding=cls.SIZES['padding_small'],
                       fieldbackground=cls.COLORS['white'],
                       borderwidth=cls.SIZES['border_width'],
                       relief='solid')

        style.map('Modern.TEntry',
                 focuscolor=[('focus', cls.COLORS['primary'])])

    @classmethod
    def _configure_treeview_styles(cls, style):
        """配置树视图样式"""
        style.configure('Modern.Treeview',
                       font=cls.FONTS['default'],
                       rowheight=cls.SIZES['row_height'],
                       background=cls.COLORS['white'],
                       fieldbackground=cls.COLORS['white'],
                       foreground=cls.COLORS['text_primary'],
                       borderwidth=0,
                       relief='flat')

        style.configure('Modern.Treeview.Heading',
                       font=cls.FONTS['heading'],
                       background=cls.COLORS['light_gray'],
                       foreground=cls.COLORS['text_primary'],
                       relief='flat',
                       borderwidth=cls.SIZES['border_width'])

        style.map('Modern.Treeview',
                 background=[('selected', cls.COLORS['bg_selected'])],
                 foreground=[('selected', cls.COLORS['primary'])])

        style.map('Modern.Treeview.Heading',
                 background=[('active', cls.COLORS['bg_selected'])])

    @classmethod
    def _configure_combobox_styles(cls, style):
        """配置下拉框样式"""
        style.configure('Modern.TCombobox',
                       font=cls.FONTS['input'],
                       fieldbackground=cls.COLORS['white'],
                       borderwidth=cls.SIZES['border_width'],
                       relief='solid')

class ThemeManager:
    """响应式主题管理器"""

    def __init__(self):
        self.current_theme = ModernTheme()
        self._responsive_config = None
        self._screen_category = None
        self._update_responsive_config()

    def _update_responsive_config(self):
        """更新响应式配置"""
        self._screen_category = responsive.get_screen_size_category()
        self._responsive_config = responsive.get_responsive_config(self._screen_category)

    def apply_theme(self, root):
        """应用主题到根窗口"""
        return self.current_theme.apply_modern_styles(root)

    def setup_responsive_window(self, window, window_type='main_window'):
        """为窗口设置响应式配置"""
        config = responsive.setup_responsive_window(window, window_type)
        self._responsive_config = config
        return config

    def get_color(self, color_name):
        """获取颜色值"""
        return self.current_theme.COLORS.get(color_name, '#000000')

    def get_font(self, font_name):
        """获取响应式字体配置"""
        base_font = self.current_theme.FONTS.get(font_name, ('Microsoft YaHei UI', 9))

        if self._responsive_config and font_name in ['default', 'small', 'medium', 'large', 'title', 'subtitle', 'heading', 'button', 'input']:
            # 映射字体名称到响应式配置
            font_mapping = {
                'title': 'title_font',
                'subtitle': 'subtitle_font',
                'heading': 'label_font',
                'default': 'label_font',
                'small': 'label_font',
                'medium': 'input_font',
                'large': 'title_font',
                'button': 'button_font',
                'input': 'input_font'
            }

            responsive_key = font_mapping.get(font_name, 'label_font')
            responsive_size = self._responsive_config.get(responsive_key, base_font[1])

            # 返回响应式字体
            return (base_font[0], responsive_size) + base_font[2:] if len(base_font) > 2 else (base_font[0], responsive_size)

        return base_font

    def get_size(self, size_name):
        """获取响应式尺寸配置"""
        base_size = self.current_theme.SIZES.get(size_name, 10)

        if self._responsive_config:
            # 映射尺寸名称到响应式配置
            size_mapping = {
                'padding_small': 'padding',
                'padding_medium': 'padding',
                'padding_large': 'padding',
                'padding_xl': 'padding',
                'margin_small': 'margin',
                'margin_medium': 'margin',
                'margin_large': 'margin',
                'margin_xl': 'margin',
                'button_height': 'button_height',
                'input_height': 'input_height'
            }

            responsive_key = size_mapping.get(size_name)
            if responsive_key:
                return self._responsive_config.get(responsive_key, base_size)

        return base_size

    def get_responsive_config(self):
        """获取当前响应式配置"""
        if not self._responsive_config:
            self._update_responsive_config()
        return self._responsive_config

    def get_screen_category(self):
        """获取屏幕分类"""
        if not self._screen_category:
            self._update_responsive_config()
        return self._screen_category

    def center_window_on_parent(self, window, parent):
        """在父窗口中心显示子窗口 - 支持高DPI缩放"""
        try:
            window.update_idletasks()

            # 获取父窗口信息
            parent_x = parent.winfo_x()
            parent_y = parent.winfo_y()
            parent_width = parent.winfo_width()
            parent_height = parent.winfo_height()

            # 获取子窗口信息
            window_width = window.winfo_width()
            window_height = window.winfo_height()

            # 计算居中位置
            x = parent_x + (parent_width - window_width) // 2
            y = parent_y + (parent_height - window_height) // 2

            # 获取真实屏幕尺寸来检查边界
            screen_width, screen_height = self._get_real_screen_size(window)

            # 确保窗口不会超出屏幕边界
            x = max(0, min(x, screen_width - window_width))
            y = max(0, min(y, screen_height - window_height))

            window.geometry(f"+{x}+{y}")
        except Exception as e:
            print(f"窗口居中失败: {e}")

    def center_window_on_screen(self, window):
        """在屏幕中心显示窗口 - 支持高DPI缩放"""
        try:
            # 强制更新窗口信息
            window.update_idletasks()

            # 获取真实屏幕尺寸和DPI信息
            tkinter_screen_width = window.winfo_screenwidth()
            tkinter_screen_height = window.winfo_screenheight()

            # 获取窗口尺寸
            window_width = window.winfo_width()
            window_height = window.winfo_height()

            # 尝试获取真实屏幕尺寸
            try:
                import ctypes
                user32 = ctypes.windll.user32
                user32.SetProcessDPIAware()

                real_screen_width = user32.GetSystemMetrics(0)
                real_screen_height = user32.GetSystemMetrics(1)

                # 计算DPI缩放比例
                scale_x = real_screen_width / tkinter_screen_width
                scale_y = real_screen_height / tkinter_screen_height

                # 在高DPI环境下，使用tkinter坐标系计算
                if scale_x > 1.5:  # 高DPI环境
                    # 使用tkinter坐标系
                    screen_width = tkinter_screen_width
                    screen_height = tkinter_screen_height
                    print("使用tkinter坐标系进行居中计算")
                else:
                    # 标准DPI环境
                    screen_width = real_screen_width
                    screen_height = real_screen_height
                    print("使用真实坐标系进行居中计算")
            except:
                # 无法获取真实屏幕尺寸，使用tkinter报告的尺寸
                screen_width = tkinter_screen_width
                screen_height = tkinter_screen_height
                print("使用tkinter坐标系（备用方案）")

            # 计算居中位置 - 更精确的居中
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2

            # 微调：向上移动一点，考虑视觉居中
            y = max(0, y - 30)

            print(f"计算的居中位置: +{x}+{y}")

            # 设置窗口位置
            geometry_string = f"+{x}+{y}"
            window.geometry(geometry_string)

            # 多次设置确保生效
            window.update()
            window.geometry(geometry_string)

            # 验证最终位置
            window.update_idletasks()
            final_x = window.winfo_x()
            final_y = window.winfo_y()

            print(f"最终位置: +{final_x}+{final_y}")

            # 如果位置明显不对，尝试直接设置到屏幕中心
            if final_x < 100 or final_y < 50:
                print("位置异常，尝试强制居中...")
                center_x = screen_width // 2 - window_width // 2
                center_y = screen_height // 2 - window_height // 2
                window.geometry(f"+{center_x}+{center_y}")
                window.update()
                print(f"强制居中位置: +{center_x}+{center_y}")

            # 额外验证：如果仍然在左上角，尝试使用Windows API
            window.update_idletasks()
            check_x = window.winfo_x()
            check_y = window.winfo_y()

            if check_x < 500 or check_y < 150:
                print("检测到窗口可能在左上角，尝试Windows API强制居中...")
                try:
                    import ctypes
                    from ctypes import wintypes

                    # 获取窗口句柄
                    hwnd_str = window.wm_frame()
                    if hwnd_str.startswith('0x'):
                        hwnd = int(hwnd_str, 16)
                    else:
                        hwnd = int(hwnd_str)

                    # 计算真实屏幕的居中位置
                    real_center_x = (real_screen_width - window_width * int(scale_x)) // 2
                    real_center_y = (real_screen_height - window_height * int(scale_y)) // 2

                    # 使用Windows API设置窗口位置
                    user32 = ctypes.windll.user32
                    user32.SetWindowPos(hwnd, 0, real_center_x, real_center_y, 0, 0, 0x0001)

                    print(f"Windows API居中: {real_center_x}, {real_center_y}")
                except Exception as api_e:
                    print(f"Windows API居中失败: {api_e}")

        except Exception as e:
            print(f"屏幕居中失败: {e}")

    def create_modern_button(self, parent, text, command=None, style='primary', **kwargs):
        """创建响应式现代化按钮"""
        color_map = {
            'primary': self.get_color('primary'),
            'success': self.get_color('success'),
            'warning': self.get_color('warning'),
            'error': self.get_color('error'),
            'secondary': self.get_color('secondary')
        }

        bg_color = color_map.get(style, self.get_color('primary'))

        # 使用响应式配置
        button_height = self.get_size('button_height')
        button_font = self.get_font('button')
        padding_x = self.get_size('padding_medium')

        button = tk.Button(
            parent,
            text=text,
            command=command,
            font=button_font,
            bg=bg_color,
            fg=self.get_color('white'),
            activebackground=self._darken_color(bg_color),
            activeforeground=self.get_color('white'),
            relief='flat',
            cursor='hand2',
            bd=0,
            height=1,
            **kwargs
        )

        self._add_button_hover_effect(button, bg_color)
        return button

    def create_modern_entry(self, parent, **kwargs):
        """创建响应式现代化输入框"""
        container = tk.Frame(parent, bg=self.get_color('white'), relief='solid', bd=1)

        # 使用响应式配置
        input_font = self.get_font('input')
        padding_x = self.get_size('padding_medium')
        padding_y = self.get_size('padding_small')

        entry = tk.Entry(
            container,
            font=input_font,
            bg=self.get_color('white'),
            fg=self.get_color('text_primary'),
            relief='flat',
            bd=0,
            **kwargs
        )
        entry.pack(fill='both', padx=padding_x, pady=padding_y)

        self._add_input_focus_effects(container, entry)
        return container, entry

    def create_card_frame(self, parent, **kwargs):
        """创建响应式卡片框架"""
        # 根据屏幕尺寸决定是否显示边框效果
        screen_category = self.get_screen_category()

        if screen_category in ['xs', 'sm']:
            # 小屏幕下简化边框
            return tk.Frame(
                parent,
                bg=self.get_color('bg_card'),
                relief='flat',
                bd=0,
                **kwargs
            )
        else:
            # 大屏幕下显示完整边框
            return tk.Frame(
                parent,
                bg=self.get_color('bg_card'),
                relief='solid',
                bd=self.get_size('border_width'),
                **kwargs
            )

    def _darken_color(self, color):
        """加深颜色"""
        color_map = {
            self.get_color('primary'): self.get_color('primary_dark'),
            self.get_color('success'): self.get_color('success_dark'),
            self.get_color('warning'): self.get_color('warning_dark'),
            self.get_color('error'): self.get_color('error_dark'),
            self.get_color('secondary'): self.get_color('secondary_dark')
        }
        return color_map.get(color, color)

    def _add_button_hover_effect(self, button, original_color):
        """为按钮添加悬停效果"""
        dark_color = self._darken_color(original_color)

        def on_enter(e):
            button['bg'] = dark_color

        def on_leave(e):
            button['bg'] = original_color

        button.bind('<Enter>', on_enter)
        button.bind('<Leave>', on_leave)

    def _add_input_focus_effects(self, container, entry):
        """为输入框添加焦点效果"""
        def on_focus_in(e):
            container.configure(bg=self.get_color('white'), relief='solid', bd=2)

        def on_focus_out(e):
            container.configure(bg=self.get_color('white'), relief='solid', bd=1)

        entry.bind('<FocusIn>', on_focus_in)
        entry.bind('<FocusOut>', on_focus_out)

    def _get_dpi_scale_factor(self, window):
        """获取DPI缩放因子"""
        try:
            # 尝试使用Windows API获取缩放因子
            import ctypes
            try:
                user32 = ctypes.windll.user32
                user32.SetProcessDPIAware()

                # 获取真实屏幕尺寸和tkinter报告的尺寸
                real_width = user32.GetSystemMetrics(0)
                tkinter_width = window.winfo_screenwidth()

                if real_width > 0 and tkinter_width > 0:
                    scale_factor = real_width / tkinter_width
                    return scale_factor
            except:
                pass

            # 备用方案：通过DPI计算
            try:
                screen_width = window.winfo_screenwidth()
                screen_width_mm = window.winfo_screenmmwidth()

                if screen_width_mm > 0:
                    dpi_x = screen_width * 25.4 / screen_width_mm
                    if dpi_x < 90:  # 明显低于标准DPI，可能有缩放
                        return 96 / dpi_x
            except:
                pass

            return 1.0  # 默认无缩放

        except Exception as e:
            print(f"获取DPI缩放因子失败: {e}")
            return 1.0

    def _get_real_screen_size(self, window):
        """获取真实屏幕尺寸，处理DPI缩放"""
        try:
            # 首先尝试使用Windows API获取真实屏幕尺寸
            import ctypes
            try:
                user32 = ctypes.windll.user32
                # 设置DPI感知
                user32.SetProcessDPIAware()

                # 获取真实屏幕尺寸
                real_width = user32.GetSystemMetrics(0)  # SM_CXSCREEN
                real_height = user32.GetSystemMetrics(1)  # SM_CYSCREEN

                if real_width > 0 and real_height > 0:
                    return real_width, real_height
            except:
                pass

            # 备用方案：使用tkinter报告的尺寸
            screen_width = window.winfo_screenwidth()
            screen_height = window.winfo_screenheight()

            # 尝试检测DPI缩放并补偿
            try:
                # 获取DPI信息
                screen_width_mm = window.winfo_screenmmwidth()
                screen_height_mm = window.winfo_screenmmheight()

                # 计算DPI
                dpi_x = screen_width * 25.4 / screen_width_mm
                dpi_y = screen_height * 25.4 / screen_height_mm

                # 如果DPI明显低于96，可能存在缩放
                if dpi_x < 90 or dpi_y < 90:
                    # 估算缩放比例并补偿
                    scale_factor = 96 / min(dpi_x, dpi_y)
                    screen_width = int(screen_width * scale_factor)
                    screen_height = int(screen_height * scale_factor)
            except:
                pass

            return screen_width, screen_height

        except Exception as e:
            print(f"获取真实屏幕尺寸失败: {e}")
            # 最后的备用方案
            return window.winfo_screenwidth(), window.winfo_screenheight()

# 全局主题管理器实例
theme_manager = ThemeManager()
