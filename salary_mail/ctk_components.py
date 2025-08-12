# coding:utf-8
"""
CustomTkinter 通用组件库
提供项目中常用的自定义组件
"""
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image
import os

class CTKTreeview(ctk.CTkFrame):
    """CustomTkinter风格的表格组件（基于ttk.Treeview）"""

    def __init__(self, parent, columns, show_checkboxes=False, **kwargs):
        super().__init__(parent, **kwargs)

        self.columns = columns
        self.show_checkboxes = show_checkboxes
        self.checked_items = set()  # 存储选中的项目
        self.checkbox_callback = None  # 复选框状态改变的回调

        # 如果显示复选框，在第一列前添加复选框列
        if self.show_checkboxes:
            self.columns = [('☐', 30)] + list(self.columns)

        self.setup_treeview()

    def setup_treeview(self):
        """设置表格视图"""
        # 创建滚动条
        self.v_scrollbar = ttk.Scrollbar(self, orient="vertical")
        self.h_scrollbar = ttk.Scrollbar(self, orient="horizontal")

        # 创建Treeview
        self.tree = ttk.Treeview(
            self,
            columns=[col[0] for col in self.columns],
            show="headings",
            selectmode="browse",  # 单选模式，确保选中稳定性
            yscrollcommand=self.v_scrollbar.set,
            xscrollcommand=self.h_scrollbar.set
        )

        # 配置列
        for col_name, width in self.columns:
            self.tree.heading(col_name, text=col_name)
            self.tree.column(col_name, width=width, anchor='center')

        # 配置滚动条
        self.v_scrollbar.config(command=self.tree.yview)
        self.h_scrollbar.config(command=self.tree.xview)

        # 布局
        self.tree.grid(row=0, column=0, sticky='nsew')
        self.v_scrollbar.grid(row=0, column=1, sticky='ns')
        self.h_scrollbar.grid(row=1, column=0, sticky='ew')

        # 配置权重
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 应用CustomTkinter风格
        self.apply_ctk_style()
        
        # 设置事件绑定
        self.setup_checkbox_events()

    def apply_ctk_style(self):
        """应用CustomTkinter风格"""
        style = ttk.Style()

        # 获取当前主题模式
        appearance_mode = ctk.get_appearance_mode()

        if appearance_mode == "Dark":
            # 深色主题 - 优化对比度，更专业的配色
            bg_color = "#1A1A1A"          # 主背景色，深灰但不刺眼
            fg_color = "#D4D4D4"          # 文字颜色，柔和的白色
            select_bg = "#264F78"         # 选中背景色，深蓝灰
            select_fg = "#FFFFFF"         # 选中文字保持白色
            field_bg = "#252525"          # 字段背景色，稍微亮一点
            border_color = "#3C3C3C"      # 边框颜色，柔和的灰色
            alternate_bg = "#202020"      # 交替行背景色
            hover_bg = "#2A2A2A"          # 悬停背景色
            heading_bg = "#2D2D2D"        # 表头背景色，比字段背景稍亮
            heading_fg = "#FFFFFF"        # 表头文字色，纯白色确保清晰
        else:
            # 浅色主题
            bg_color = "#FFFFFF"
            fg_color = "#000000"
            select_bg = "#0078D4"
            select_fg = "#FFFFFF"
            field_bg = "#F0F0F0"
            border_color = "#CCCCCC"

        # 配置Treeview样式
        style.configure("CTK.Treeview",
                       rowheight=35,
                       background=bg_color,
                       fieldbackground=field_bg,
                       foreground=fg_color,
                       borderwidth=1,
                       relief="solid",
                       font=('Microsoft YaHei UI', 11))

        style.configure("CTK.Treeview.Heading",
                       font=('Microsoft YaHei UI', 12, 'bold'),
                       background=heading_bg if appearance_mode == "Dark" else field_bg,
                       foreground=heading_fg if appearance_mode == "Dark" else fg_color,
                       borderwidth=1,
                       relief="flat")

        style.map("CTK.Treeview",
                 background=[('selected', select_bg)],
                 foreground=[('selected', select_fg)])

        # 深色模式下添加交替行颜色和悬停效果
        if appearance_mode == "Dark":
            style.map("CTK.Treeview",
                     background=[('selected', select_bg),
                                ('alternate', alternate_bg),
                                ('hover', hover_bg)],
                     foreground=[('selected', select_fg)])
        else:
            # 浅色模式下的悬停效果
            style.map("CTK.Treeview",
                     background=[('selected', select_bg),
                                ('hover', "#F5F5F5")],
                     foreground=[('selected', select_fg)])

        # 表头悬停效果
        if appearance_mode == "Dark":
            style.map("CTK.Treeview.Heading",
                     background=[('active', "#404040")])  # 深色模式下的悬停色
        else:
            style.map("CTK.Treeview.Heading",
                     background=[('active', border_color)])  # 浅色模式下的悬停色

        # 应用样式
        self.tree.configure(style="CTK.Treeview")

        # 绑定主题变化事件
        self.bind_theme_change()

    def bind_theme_change(self):
        """绑定主题变化事件"""
        def on_theme_change():
            # 延迟应用新样式，确保主题已完全切换
            self.after(100, self.apply_ctk_style)

        # 监听主题变化
        try:
            # 使用after方法定期检查主题是否变化
            self.check_theme_change()
        except:
            pass

    def check_theme_change(self):
        """检查主题变化"""
        current_mode = ctk.get_appearance_mode()
        if not hasattr(self, '_last_appearance_mode'):
            self._last_appearance_mode = current_mode
        elif self._last_appearance_mode != current_mode:
            self._last_appearance_mode = current_mode
            self.apply_ctk_style()

        # 每500ms检查一次主题变化
        self.after(500, self.check_theme_change)

    def insert(self, parent, index, **kwargs):
        """插入数据"""
        return self.tree.insert(parent, index, **kwargs)

    def delete(self, *items):
        """删除数据"""
        return self.tree.delete(*items)

    def get_children(self, item=""):
        """获取子项"""
        return self.tree.get_children(item)

    def selection(self):
        """获取选中项"""
        return self.tree.selection()

    def item(self, item, option=None, **kwargs):
        """获取或设置项目属性"""
        return self.tree.item(item, option, **kwargs)

    def bind(self, sequence, func, add=None):
        """绑定事件"""
        return self.tree.bind(sequence, func, add)

    def setup_checkbox_events(self):
        """设置事件绑定"""
        if self.show_checkboxes:
            # 绑定复选框点击事件
            self.tree.bind('<Button-1>', self.on_checkbox_click)
            # 绑定选择事件
            self.tree.bind('<<TreeviewSelect>>', self.on_selection_change)
            # 绑定焦点事件，确保选择状态稳定
            self.tree.bind('<FocusIn>', self.on_focus_in)
            self.tree.bind('<FocusOut>', self.on_focus_out)
        else:
            # 没有复选框时，不绑定任何事件，让外部直接绑定
            # 这样工资管理和员工管理页面可以直接绑定自己的选择事件
            pass

    def on_checkbox_click(self, event):
        """处理复选框点击"""
        if not self.show_checkboxes:
            return

        # 获取点击的位置
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)

        # 如果点击的是第一列（复选框列）且点击在复选框区域内
        if item and column == '#1':
            # 简化判断：如果点击在第一列的前30像素内，认为是复选框点击
            if event.x <= 30:
                # 切换复选框状态
                self.toggle_checkbox(item)
                return "break"  # 阻止默认选择行为

        # 其他区域允许正常的行选择
        return None

    def on_selection_change(self, event):
        """处理选择变化事件"""
        # 不做额外处理，保持选择状态稳定
        pass

    def on_focus_in(self, event):
        """焦点进入事件"""
        pass

    def on_focus_out(self, event):
        """焦点离开事件"""
        # 保持选择状态，不清除
        pass

    def toggle_checkbox(self, item):
        """切换复选框状态"""
        if item in self.checked_items:
            self.checked_items.remove(item)
            # 更新显示为未选中
            values = list(self.tree.item(item)['values'])
            values[0] = '☐'
            self.tree.item(item, values=values)
        else:
            self.checked_items.add(item)
            # 更新显示为选中
            values = list(self.tree.item(item)['values'])
            values[0] = '☑'
            self.tree.item(item, values=values)

        # 延迟触发回调，避免在事件处理过程中产生副作用
        if self.checkbox_callback:
            self.after(10, self.checkbox_callback)

    def set_checkbox_callback(self, callback):
        """设置复选框状态改变的回调函数"""
        self.checkbox_callback = callback

    def get_checked_items(self):
        """获取所有选中的项目"""
        return list(self.checked_items)

    def check_all(self):
        """全选"""
        if not self.show_checkboxes:
            return

        for item in self.tree.get_children():
            if item not in self.checked_items:
                self.toggle_checkbox(item)

    def uncheck_all(self):
        """取消全选"""
        if not self.show_checkboxes:
            return

        for item in list(self.checked_items):
            self.toggle_checkbox(item)


class CTKProgressDialog(ctk.CTkToplevel):
    """进度条对话框"""

    def __init__(self, parent, title="处理中", message="请稍候..."):
        super().__init__(parent)

        self.title(title)
        self.geometry("400x200")
        self.resizable(False, False)

        # 设置为模态窗口
        self.transient(parent)
        self.grab_set()

        # 居中显示
        self.center_on_parent(parent)

        # 创建UI
        self.setup_ui(message)

        # 初始化变量
        self.cancelled = False

    def center_on_parent(self, parent):
        """在父窗口中心显示"""
        self.update_idletasks()

        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()

        x = parent_x + (parent_width - self.winfo_width()) // 2
        y = parent_y + (parent_height - self.winfo_height()) // 2

        self.geometry(f"+{x}+{y}")

    def setup_ui(self, message):
        """设置UI"""
        main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # 消息标签
        self.message_label = ctk.CTkLabel(
            main_frame,
            text=message,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold")
        )
        self.message_label.pack(pady=(0, 20))

        # 进度条
        self.progress_bar = ctk.CTkProgressBar(
            main_frame,
            width=350,
            height=20
        )
        self.progress_bar.pack(pady=10)
        self.progress_bar.set(0)

        # 百分比标签
        self.percent_label = ctk.CTkLabel(
            main_frame,
            text="0%",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12)
        )
        self.percent_label.pack(pady=5)

        # 详细信息标签
        self.detail_label = ctk.CTkLabel(
            main_frame,
            text="准备中...",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=10)
        )
        self.detail_label.pack(pady=5)

        # 取消按钮
        self.cancel_btn = ctk.CTkButton(
            main_frame,
            text="取消",
            command=self.cancel_operation,
            width=100,
            height=32
        )
        self.cancel_btn.pack(pady=10)

    def update_progress(self, value, message=None, detail=None):
        """更新进度"""
        self.progress_bar.set(value / 100.0)
        self.percent_label.configure(text=f"{value:.1f}%")

        if message:
            self.message_label.configure(text=message)

        if detail:
            self.detail_label.configure(text=detail)

        self.update()

    def cancel_operation(self):
        """取消操作"""
        self.cancelled = True
        self.destroy()


class CTKMessageBox:
    """CustomTkinter风格的消息框"""

    @staticmethod
    def show_info(parent, title, message):
        """显示信息对话框"""
        return CTKMessageBox._show_dialog(parent, title, message, "info")

    @staticmethod
    def show_warning(parent, title, message):
        """显示警告对话框"""
        return CTKMessageBox._show_dialog(parent, title, message, "warning")

    @staticmethod
    def show_error(parent, title, message):
        """显示错误对话框"""
        return CTKMessageBox._show_dialog(parent, title, message, "error")

    @staticmethod
    def ask_yes_no(parent, title, message):
        """显示确认对话框"""
        return CTKMessageBox._show_dialog(parent, title, message, "question", buttons=["是", "否"])

    @staticmethod
    def _show_dialog(parent, title, message, icon_type, buttons=None):
        """显示对话框的内部实现"""
        if buttons is None:
            buttons = ["确定"]

        dialog = CTKDialog(parent, title, message, icon_type, buttons)
        return dialog.show()


class CTKDialog(ctk.CTkToplevel):
    """通用对话框"""

    def __init__(self, parent, title, message, icon_type, buttons):
        super().__init__(parent)

        self.result = None
        self.buttons = buttons

        self.title(title)
        self.geometry("400x200")
        self.resizable(False, False)

        # 设置为模态窗口
        self.transient(parent)
        self.grab_set()

        # 居中显示
        self.center_on_parent(parent)

        # 创建UI
        self.setup_ui(message, icon_type)

    def center_on_parent(self, parent):
        """在父窗口中心显示"""
        self.update_idletasks()

        if parent:
            parent_x = parent.winfo_x()
            parent_y = parent.winfo_y()
            parent_width = parent.winfo_width()
            parent_height = parent.winfo_height()

            x = parent_x + (parent_width - self.winfo_width()) // 2
            y = parent_y + (parent_height - self.winfo_height()) // 2

            self.geometry(f"+{x}+{y}")

    def setup_ui(self, message, icon_type):
        """设置UI"""
        main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # 消息区域
        message_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        message_frame.pack(fill="both", expand=True, pady=(0, 20))

        # 消息文本
        message_label = ctk.CTkLabel(
            message_frame,
            text=message,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14),
            justify="center",
            wraplength=350
        )
        message_label.pack(expand=True)

        # 按钮区域
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x")

        # 创建按钮
        for i, button_text in enumerate(self.buttons):
            btn = ctk.CTkButton(
                button_frame,
                text=button_text,
                command=lambda result=button_text: self.on_button_click(result),
                width=80,
                height=35
            )
            btn.pack(side="right", padx=5)

        # 绑定回车键到第一个按钮
        self.bind('<Return>', lambda e: self.on_button_click(self.buttons[0]))

        # 绑定ESC键到最后一个按钮
        if len(self.buttons) > 1:
            self.bind('<Escape>', lambda e: self.on_button_click(self.buttons[-1]))

    def on_button_click(self, result):
        """按钮点击处理"""
        self.result = result
        self.destroy()

    def show(self):
        """显示对话框并返回结果"""
        self.wait_window()
        return self.result


class CTKFileDialog:
    """CustomTkinter风格的文件对话框封装"""

    @staticmethod
    def open_file(parent, title="选择文件", filetypes=None):
        """打开文件对话框"""
        if filetypes is None:
            filetypes = [("所有文件", "*.*")]

        # 临时取消窗口置顶
        if hasattr(parent, 'attributes'):
            parent.attributes("-topmost", 0)

        file_path = filedialog.askopenfilename(
            title=title,
            filetypes=filetypes,
            parent=parent
        )

        # 恢复窗口焦点
        if parent:
            parent.focus_force()

        return file_path

    @staticmethod
    def save_file(parent, title="保存文件", filetypes=None, defaultextension=None):
        """保存文件对话框"""
        if filetypes is None:
            filetypes = [("所有文件", "*.*")]

        # 临时取消窗口置顶
        if hasattr(parent, 'attributes'):
            parent.attributes("-topmost", 0)

        file_path = filedialog.asksaveasfilename(
            title=title,
            filetypes=filetypes,
            defaultextension=defaultextension,
            parent=parent
        )

        # 恢复窗口焦点
        if parent:
            parent.focus_force()

        return file_path


class CTKWindowSizeManager:
    """窗口大小管理器 - 处理高分辨率缩放适配"""
    
    @staticmethod
    def adjust_window_size(window, base_width, base_height, min_width=None, min_height=None, max_screen_ratio=(0.8, 0.8)):
        """
        根据屏幕DPI和分辨率自动调整窗口大小
        
        Args:
            window: 窗口对象
            base_width: 基础宽度
            base_height: 基础高度
            min_width: 最小宽度（默认为base_width的0.9倍）
            min_height: 最小高度（默认为base_height的0.9倍）
            max_screen_ratio: 最大屏幕占比 (width_ratio, height_ratio)
        """
        try:
            # 获取屏幕信息
            screen_width = window.winfo_screenwidth()
            screen_height = window.winfo_screenheight()
            
            # 默认最小尺寸
            if min_width is None:
                min_width = int(base_width * 0.9)
            if min_height is None:
                min_height = int(base_height * 0.9)
            
            # 根据屏幕尺寸调整
            if screen_width >= 3000:  # 高分辨率显示器
                width = int(base_width * 1.2)
                height = int(base_height * 1.2)
                min_width = int(min_width * 1.2)
                min_height = int(min_height * 1.2)
            elif screen_width >= 2000:  # 中等分辨率
                width = int(base_width * 1.1)
                height = int(base_height * 1.1)
                min_width = int(min_width * 1.1)
                min_height = int(min_height * 1.1)
            else:  # 标准分辨率
                width = base_width
                height = base_height
            
            # 确保窗口不会超过屏幕限制
            max_width = int(screen_width * max_screen_ratio[0])
            max_height = int(screen_height * max_screen_ratio[1])
            
            width = min(width, max_width)
            height = min(height, max_height)
            
            window.geometry(f'{width}x{height}')
            window.minsize(min_width, min_height)
            
            return width, height
            
        except Exception as e:
            print(f"调整窗口大小失败: {e}")
            # 使用默认尺寸
            window.geometry(f'{base_width}x{base_height}')
            if min_width and min_height:
                window.minsize(min_width, min_height)
            return base_width, base_height

    @staticmethod
    def get_scroll_frame_size(base_width, base_height):
        """
        根据屏幕分辨率获取滚动框架的合适尺寸
        
        Args:
            base_width: 基础宽度
            base_height: 基础高度
            
        Returns:
            tuple: (scroll_width, scroll_height)
        """
        try:
            import tkinter as tk
            root = tk.Tk()
            screen_width = root.winfo_screenwidth()
            root.destroy()
            
            if screen_width >= 3000:
                return int(base_width * 1.2), int(base_height * 1.2)
            elif screen_width >= 2000:
                return int(base_width * 1.1), int(base_height * 1.1)
            else:
                return base_width, base_height
                
        except Exception:
            return base_width, base_height


class CTKSearchEntry(ctk.CTkFrame):
    """带搜索功能的输入框"""

    def __init__(self, parent, placeholder="搜索...", **kwargs):
        super().__init__(parent, **kwargs)

        self.search_var = tk.StringVar()
        self.search_callback = None
        self.setup_ui(placeholder)

    def setup_ui(self, placeholder):
        """设置UI"""
        # 搜索输入框
        self.search_entry = ctk.CTkEntry(
            self,
            textvariable=self.search_var,
            placeholder_text=placeholder,
            height=35
        )
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        # 搜索按钮
        self.search_btn = ctk.CTkButton(
            self,
            text="🔍",
            width=35,
            height=35,
            command=self.perform_search
        )
        self.search_btn.pack(side="right")

        # 绑定实时搜索
        self.search_var.trace('w', self.on_search_changed)

        # 绑定回车键
        self.search_entry.bind('<Return>', lambda e: self.perform_search())

    def on_search_changed(self, *args):
        """搜索内容变化时的处理"""
        if self.search_callback:
            # 延迟搜索，避免频繁触发
            if hasattr(self, 'search_timer'):
                self.after_cancel(self.search_timer)

            self.search_timer = self.after(300, self.perform_search)

    def perform_search(self):
        """执行搜索"""
        if self.search_callback:
            search_text = self.search_var.get().strip()
            self.search_callback(search_text)

    def set_search_callback(self, callback):
        """设置搜索回调函数"""
        self.search_callback = callback

    def get_text(self):
        """获取搜索文本"""
        return self.search_var.get().strip()

    def clear(self):
        """清空搜索框"""
        self.search_var.set("")
