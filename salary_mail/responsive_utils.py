# -*- coding: utf-8 -*-
"""
响应式设计工具类
提供统一的分辨率适配和响应式布局支持
"""

import tkinter as tk

class ResponsiveDesign:
    """响应式设计工具类"""
    
    # 分辨率断点
    BREAKPOINTS = {
        'xs': (0, 800),        # 极小屏幕
        'sm': (801, 1024),     # 小屏幕
        'md': (1025, 1366),    # 中等屏幕
        'lg': (1367, 1920),    # 大屏幕
        'xl': (1921, 9999)     # 超大屏幕
    }
    
    @classmethod
    def get_screen_size_category(cls, width=None, height=None):
        """获取屏幕尺寸分类"""
        if width is None:
            root = tk.Tk()
            width = root.winfo_screenwidth()
            height = root.winfo_screenheight()
            root.destroy()
        
        for category, (min_w, max_w) in cls.BREAKPOINTS.items():
            if min_w <= width <= max_w:
                return category
        return 'xl'
    
    @classmethod
    def get_responsive_config(cls, screen_category=None):
        """获取响应式配置"""
        if screen_category is None:
            screen_category = cls.get_screen_size_category()
        
        # 获取屏幕尺寸来计算对话框高度（屏幕高度的2/3）
        try:
            root = tk.Tk()
            screen_height = root.winfo_screenheight()
            root.destroy()
            dialog_height_2_3 = int(screen_height * 2 / 3)
        except:
            dialog_height_2_3 = 600  # 备用值
        
        configs = {
            'xs': {
                # 窗口配置
                'login_window': {'width': 380, 'height': 380},
                'main_window': {'width': 800, 'height': 600, 'min_width': 750, 'min_height': 550},
                'dialog_window': {'width': 400, 'height': max(300, dialog_height_2_3)},
                
                # 字体配置
                'title_font': 16,
                'subtitle_font': 8,
                'label_font': 8,
                'input_font': 9,
                'button_font': 9,
                'footer_font': 7,
                
                # 间距配置
                'padding': 10,
                'margin': 8,
                'section_spacing': 10,
                'input_spacing': 10,
                'button_spacing': 10,
                
                # 元素大小
                'button_height': 30,
                'input_height': 28,
                'logo_size': 40,
                'icon_size': 11,
                
                # 显示控制
                'show_logo': True,
                'show_subtitle': False,
                'show_footer': True,
                'show_shadows': False,
                'card_effects': False
            },
            
            'sm': {
                # 窗口配置
                'login_window': {'width': 420, 'height': 420},
                'main_window': {'width': 900, 'height': 650, 'min_width': 850, 'min_height': 600},
                'dialog_window': {'width': 450, 'height': max(350, dialog_height_2_3)},
                
                # 字体配置
                'title_font': 18,
                'subtitle_font': 9,
                'label_font': 9,
                'input_font': 10,
                'button_font': 10,
                'footer_font': 7,
                
                # 间距配置
                'padding': 15,
                'margin': 10,
                'section_spacing': 12,
                'input_spacing': 12,
                'button_spacing': 12,
                
                # 元素大小
                'button_height': 35,
                'input_height': 32,
                'logo_size': 45,
                'icon_size': 12,
                
                # 显示控制
                'show_logo': True,
                'show_subtitle': False,
                'show_footer': True,
                'show_shadows': True,
                'card_effects': True
            },
            
            'md': {
                # 窗口配置
                'login_window': {'width': 480, 'height': 480},
                'main_window': {'width': 1000, 'height': 700, 'min_width': 900, 'min_height': 650},
                'dialog_window': {'width': 500, 'height': max(400, dialog_height_2_3)},
                
                # 字体配置
                'title_font': 22,
                'subtitle_font': 10,
                'label_font': 10,
                'input_font': 11,
                'button_font': 11,
                'footer_font': 8,
                
                # 间距配置
                'padding': 25,
                'margin': 15,
                'section_spacing': 18,
                'input_spacing': 15,
                'button_spacing': 15,
                
                # 元素大小
                'button_height': 40,
                'input_height': 36,
                'logo_size': 55,
                'icon_size': 13,
                
                # 显示控制
                'show_logo': True,
                'show_subtitle': True,
                'show_footer': True,
                'show_shadows': True,
                'card_effects': True
            },
            
            'lg': {
                # 窗口配置
                'login_window': {'width': 550, 'height': 520},
                'main_window': {'width': 1200, 'height': 800, 'min_width': 1000, 'min_height': 700},
                'dialog_window': {'width': 600, 'height': max(450, dialog_height_2_3)},
                
                # 字体配置
                'title_font': 26,
                'subtitle_font': 11,
                'label_font': 10,
                'input_font': 11,
                'button_font': 12,
                'footer_font': 8,
                
                # 间距配置
                'padding': 30,
                'margin': 20,
                'section_spacing': 25,
                'input_spacing': 18,
                'button_spacing': 18,
                
                # 元素大小
                'button_height': 45,
                'input_height': 40,
                'logo_size': 60,
                'icon_size': 14,
                
                # 显示控制
                'show_logo': True,
                'show_subtitle': True,
                'show_footer': True,
                'show_shadows': True,
                'card_effects': True
            },
            
            'xl': {
                # 窗口配置
                'login_window': {'width': 600, 'height': 550},
                'main_window': {'width': 1400, 'height': 900, 'min_width': 1200, 'min_height': 800},
                'dialog_window': {'width': 700, 'height': max(500, dialog_height_2_3)},
                
                # 字体配置
                'title_font': 30,
                'subtitle_font': 12,
                'label_font': 11,
                'input_font': 12,
                'button_font': 13,
                'footer_font': 9,
                
                # 间距配置
                'padding': 40,
                'margin': 25,
                'section_spacing': 30,
                'input_spacing': 22,
                'button_spacing': 22,
                
                # 元素大小
                'button_height': 50,
                'input_height': 44,
                'logo_size': 70,
                'icon_size': 16,
                
                # 显示控制
                'show_logo': True,
                'show_subtitle': True,
                'show_footer': True,
                'show_shadows': True,
                'card_effects': True
            }
        }
        
        return configs.get(screen_category, configs['md'])
    
    @classmethod
    def setup_responsive_window(cls, window, window_type='main_window', custom_config=None):
        """为窗口设置响应式配置 - 支持高DPI缩放"""
        screen_category = cls.get_screen_size_category()
        config = cls.get_responsive_config(screen_category)
        
        if custom_config:
            config.update(custom_config)
        
        # 获取窗口配置
        window_config = config.get(window_type, config['main_window'])
        
        # 设置窗口尺寸
        width = window_config['width']
        height = window_config['height']
        
        # 获取真实屏幕尺寸来计算居中位置
        screen_width, screen_height = cls._get_real_screen_size(window)
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        window.geometry(f'{width}x{height}+{x}+{y}')
        
        # 设置最小尺寸（如果有）
        if 'min_width' in window_config and 'min_height' in window_config:
            window.minsize(window_config['min_width'], window_config['min_height'])
        
        # 将配置保存到窗口对象
        window._responsive_config = config
        window._screen_category = screen_category
        
        return config
    
    @classmethod
    def get_font_config(cls, window, font_type='label_font'):
        """获取字体配置"""
        config = getattr(window, '_responsive_config', cls.get_responsive_config())
        return config.get(font_type, 10)
    
    @classmethod
    def get_spacing_config(cls, window, spacing_type='padding'):
        """获取间距配置"""
        config = getattr(window, '_responsive_config', cls.get_responsive_config())
        return config.get(spacing_type, 15)
    
    @classmethod
    def get_size_config(cls, window, size_type='button_height'):
        """获取尺寸配置"""
        config = getattr(window, '_responsive_config', cls.get_responsive_config())
        return config.get(size_type, 35)
    
    @classmethod
    def should_show_element(cls, window, element_type='show_logo'):
        """判断是否应该显示某个元素"""
        config = getattr(window, '_responsive_config', cls.get_responsive_config())
        return config.get(element_type, True)
    
    @classmethod
    def create_responsive_frame(cls, parent, bg='white', **kwargs):
        """创建响应式框架"""
        frame = tk.Frame(parent, bg=bg, **kwargs)
        
        # 根据屏幕尺寸调整框架属性
        screen_category = cls.get_screen_size_category()
        if screen_category in ['xs', 'sm']:
            # 小屏幕下简化边框
            frame.configure(relief='flat', bd=0)
        
        return frame
    
    @classmethod
    def create_responsive_button(cls, parent, text, command=None, style_type='primary', **kwargs):
        """创建响应式按钮"""
        config = cls.get_responsive_config()
        
        # 颜色配置
        colors = {
            'primary': '#1976d2',
            'success': '#4caf50',
            'warning': '#ff9800',
            'error': '#f44336',
            'secondary': '#757575'
        }
        
        bg_color = colors.get(style_type, colors['primary'])
        
        button = tk.Button(
            parent,
            text=text,
            command=command,
            font=('Microsoft YaHei UI', config['button_font']),
            bg=bg_color,
            fg='white',
            relief='flat',
            cursor='hand2',
            bd=0,
            height=1,
            **kwargs
        )
        
        return button
    
    @classmethod
    def _get_real_screen_size(cls, window):
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

# 全局响应式设计实例
responsive = ResponsiveDesign()
