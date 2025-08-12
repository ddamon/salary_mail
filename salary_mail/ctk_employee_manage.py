# coding:utf-8
"""
CustomTkinter版本的员工管理窗口
"""
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
import openpyxl
import re
from salary_mail.db_instance import Employee
from salary_mail.ctk_components import CTKTreeview, CTKMessageBox, CTKFileDialog, CTKSearchEntry, CTKWindowSizeManager

class CTKEmployeeDialog(ctk.CTkToplevel):
    """员工信息编辑对话框"""

    def __init__(self, parent, employee=None):
        super().__init__(parent)

        self.title('编辑员工' if employee else '添加员工')
        
        # 根据屏幕DPI自动调整窗口大小
        CTKWindowSizeManager.adjust_window_size(self, 600, 650, 550, 700)
        self.resizable(True, True)

        # 设置为模态窗口
        self.transient(parent)
        self.grab_set()

        # 居中显示
        self.center_on_parent(parent)

        self.parent = parent
        self.db = parent.db
        self.employee = employee

        self.setup_ui()



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

    def setup_ui(self):
        """设置UI"""
        # 主容器
        main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=25, pady=25)

        # 标题
        title_label = ctk.CTkLabel(
            main_frame,
            text="编辑员工信息" if self.employee else "添加员工信息",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=20, weight="bold")
        )
        title_label.pack(pady=(0, 25))

        # 表单区域 - 使用滚动容器
        form_frame = ctk.CTkFrame(main_frame, corner_radius=15)
        form_frame.pack(fill="both", expand=True, pady=(0, 25))

        # 创建滚动框架 - 适配高分辨率缩放
        scroll_width, scroll_height = CTKWindowSizeManager.get_scroll_frame_size(520, 450)
            
        self.scroll_frame = ctk.CTkScrollableFrame(
            form_frame,
            width=scroll_width,
            height=scroll_height,
            corner_radius=10,
            scrollbar_button_color="#606060",
            scrollbar_button_hover_color="#404040"
        )
        self.scroll_frame.pack(fill="both", expand=True, padx=15, pady=15)

        # 表单内容（放在滚动框架内）
        form_content = self.scroll_frame

        # 创建输入字段
        self.create_form_fields(form_content)

        # 按钮区域
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x", pady=(10, 0))

        # 取消按钮
        cancel_btn = ctk.CTkButton(
            button_frame,
            text="取消",
            command=self.destroy,
            width=120,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13, weight="bold"),
            fg_color="gray",
            hover_color="darkgray"
        )
        cancel_btn.pack(side="right", padx=(10, 0))

        # 保存按钮
        save_btn = ctk.CTkButton(
            button_frame,
            text="保存",
            command=self.save,
            width=120,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13, weight="bold"),
            fg_color="#4CAF50",
            hover_color="#45a049"
        )
        save_btn.pack(side="right", padx=(0, 10))

        # 绑定回车键
        self.bind('<Return>', lambda e: self.save())

        # 设置初始焦点
        self.set_initial_focus()

    def create_form_fields(self, parent):
        """创建表单字段"""
        fields = [
            ('职工编号', 'employee_id', True, '请输入职工编号'),
            ('姓名', 'name', True, '请输入员工姓名'),
            ('邮箱', 'email', True, '请输入邮箱地址'),
            ('手机号码', 'phone', False, '请输入手机号码（可选）')
        ]

        self.entries = {}

        for i, (label, field, required, placeholder) in enumerate(fields):
            # 标签
            label_text = f"*{label}" if required else label
            field_label = ctk.CTkLabel(
                parent,
                text=label_text,
                font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
                anchor="w"
            )
            field_label.pack(fill="x", pady=(15 if i > 0 else 0, 8))

            # 输入框
            entry = ctk.CTkEntry(
                parent,
                placeholder_text=placeholder,
                height=45,
                font=ctk.CTkFont(family="Microsoft YaHei UI", size=13)
            )
            entry.pack(fill="x", pady=(0, 10))

            # 保存引用
            self.entries[field] = entry

            # 如果是编辑模式，填充现有数据
            if self.employee:
                value = getattr(self.employee, field) or ''
                entry.insert(0, value)

        # 状态选择
        status_label = ctk.CTkLabel(
            parent,
            text="员工状态",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            anchor="w"
        )
        status_label.pack(fill="x", pady=(20, 8))

        # 状态选择框
        status_frame = ctk.CTkFrame(parent, corner_radius=10)
        status_frame.pack(fill="x", pady=(0, 15))

        # 状态选择内容
        status_content = ctk.CTkFrame(status_frame, fg_color="transparent")
        status_content.pack(fill="x", padx=15, pady=15)

        self.status_var = tk.IntVar(value=self.employee.status if self.employee else 1)

        # 在职选项
        active_radio = ctk.CTkRadioButton(
            status_content,
            text="在职",
            variable=self.status_var,
            value=1,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13)
        )
        active_radio.pack(side="left", padx=(0, 30))

        # 离职选项
        inactive_radio = ctk.CTkRadioButton(
            status_content,
            text="离职",
            variable=self.status_var,
            value=0,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=13)
        )
        inactive_radio.pack(side="left")

        # 说明文本
        note_label = ctk.CTkLabel(
            parent,
            text="注：* 表示必填字段",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=11),
            text_color="gray"
        )
        note_label.pack(pady=(15, 5))

    def set_initial_focus(self):
        """设置初始焦点"""
        if hasattr(self, 'entries') and 'employee_id' in self.entries:
            self.after(100, lambda: self.entries['employee_id'].focus())

    def save(self):
        """保存员工信息"""
        try:
            # 验证必填字段
            required_fields = ['employee_id', 'name', 'email']
            for field in required_fields:
                if not self.entries[field].get().strip():
                    field_names = {'employee_id': '职工编号', 'name': '姓名', 'email': '邮箱'}
                    CTKMessageBox.show_warning(self, '提示', f'{field_names[field]}不能为空！')
                    self.entries[field].focus()
                    return

            # 验证邮箱格式
            email = self.entries['email'].get().strip()
            if not re.match(r'^[\w\.-]+@[\w\.-]+\.\w+$', email):
                CTKMessageBox.show_warning(self, '提示', '请输入有效的邮箱地址！')
                self.entries['email'].focus()
                return

            # 检查职工编号是否已存在
            employee_id = self.entries['employee_id'].get().strip()
            existing = self.db.query(Employee).filter_by(employee_id=employee_id).first()
            if existing and (not self.employee or existing.id != self.employee.id):
                CTKMessageBox.show_warning(self, '提示', '该职工编号已存在！')
                self.entries['employee_id'].focus()
                return

            # 检查邮箱是否已存在
            existing = self.db.query(Employee).filter_by(email=email).first()
            if existing and (not self.employee or existing.id != self.employee.id):
                CTKMessageBox.show_warning(self, '提示', '该邮箱已存在！')
                self.entries['email'].focus()
                return

            # 保存数据
            if not self.employee:
                self.employee = Employee()

            for field, entry in self.entries.items():
                setattr(self.employee, field, entry.get().strip())
            self.employee.status = self.status_var.get()

            self.db.add(self.employee)
            self.db.commit()

            CTKMessageBox.show_info(self, '成功', '保存成功！')
            self.destroy()

        except Exception as e:
            CTKMessageBox.show_error(self, '错误', f'保存失败：\n{str(e)}')
            self.db.rollback()


class CTKEmployeeManageWin(ctk.CTkToplevel):
    """员工管理窗口"""

    def __init__(self, parent):
        super().__init__(parent)

        self.title('员工管理')
        # 设置为全屏显示
        self.set_fullscreen_window()

        # 设置窗口属性
        self.transient(parent)

        self.parent = parent
        self.db = parent.db

        # 定义列配置（不包含复选框列，会自动添加）
        self.columns = [
            ('序号', 60),
            ('职工编号', 100),
            ('姓名', 80),
            ('邮箱', 200),
            ('手机号码', 120),
            ('状态', 80)
        ]

        self.setup_ui()

        # 设置焦点
        self.focus_force()

    def set_fullscreen_window(self):
        """设置窗口为全屏显示"""
        try:
            # 获取屏幕尺寸
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            
            # 设置窗口几何为全屏尺寸
            self.geometry(f"{screen_width}x{screen_height}+0+0")
            
            # 设置为最大化状态
            self.state('zoomed')
            
        except Exception as e:
            print(f"设置员工管理窗口全屏失败: {e}")
            # 备用方案：使用窗口大小管理器
            CTKWindowSizeManager.adjust_window_size(self, 1200, 750, 1000, 650, (0.9, 0.8))

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

    def setup_ui(self):
        """设置UI"""
        # 主容器
        main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)

        # 顶部工具栏
        self.create_toolbar(main_frame)

        # 员工列表
        self.create_employee_list(main_frame)

        # 加载数据
        self.load_employees()

    def create_toolbar(self, parent):
        """创建工具栏"""
        toolbar_frame = ctk.CTkFrame(parent, height=60, corner_radius=10)
        toolbar_frame.pack(fill="x", pady=(0, 15))
        toolbar_frame.pack_propagate(False)

        # 工具栏内容
        toolbar_content = ctk.CTkFrame(toolbar_frame, fg_color="transparent")
        toolbar_content.pack(fill="both", expand=True, padx=15, pady=10)

        # 左侧按钮组
        left_buttons = ctk.CTkFrame(toolbar_content, fg_color="transparent")
        left_buttons.pack(side="left", fill="y")

        # 导入员工按钮
        import_btn = ctk.CTkButton(
            left_buttons,
            text="📋 导入员工",
            command=self.import_employees,
            width=120,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold")
        )
        import_btn.pack(side="left", padx=(0, 10))

        # 添加员工按钮
        add_btn = ctk.CTkButton(
            left_buttons,
            text="➕ 添加员工",
            command=self.add_employee,
            width=120,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold"),
            fg_color="#4CAF50",
            hover_color="#45a049"
        )
        add_btn.pack(side="left", padx=(0, 10))

        # 修改按钮
        self.edit_btn = ctk.CTkButton(
            left_buttons,
            text="✏️ 修改",
            command=self.edit_employee,
            width=80,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold"),
            fg_color="#FF9800",
            hover_color="#F57C00",

        )
        self.edit_btn.pack(side="left", padx=(0, 10))

        # 删除按钮
        self.delete_btn = ctk.CTkButton(
            left_buttons,
            text="🗑️ 删除",
            command=self.delete_employee,
            width=80,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold"),
            fg_color="#F44336",
            hover_color="#D32F2F",

        )
        self.delete_btn.pack(side="left", padx=(0, 20))

        # 右侧搜索区域
        search_frame = ctk.CTkFrame(toolbar_content, fg_color="transparent")
        search_frame.pack(side="right", fill="y")

        # 搜索标签
        search_label = ctk.CTkLabel(
            search_frame,
            text="搜索：",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold")
        )
        search_label.pack(side="left", padx=(0, 5))

        # 搜索框
        self.search_entry = CTKSearchEntry(
            search_frame,
            placeholder="输入关键字搜索员工...",
            width=250,
            height=40
        )
        self.search_entry.pack(side="left")
        self.search_entry.set_search_callback(self.perform_search)

        # 搜索结果计数
        self.search_count_label = ctk.CTkLabel(
            search_frame,
            text="",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=10),
            text_color="gray"
        )
        self.search_count_label.pack(side="left", padx=(10, 0))

    def create_employee_list(self, parent):
        """创建员工列表"""
        # 列表容器 - 深色模式下使用深色背景
        list_frame = ctk.CTkFrame(parent, corner_radius=10)
        list_frame.pack(fill="both", expand=True)

        # 列表内容 - 深色模式下使用深色背景
        list_content = ctk.CTkFrame(list_frame, fg_color="transparent")
        list_content.pack(fill="both", expand=True, padx=15, pady=15)

        # 恢复使用CTKTreeview，与工资管理页面保持一致
        self.employee_tree = CTKTreeview(
            list_content,
            columns=self.columns,
            corner_radius=5
        )
        self.employee_tree.pack(fill="both", expand=True)

        # 使用与工资管理页面相同的简单选择事件
        self.employee_tree.bind('<<TreeviewSelect>>', self.on_select)

    def load_employees(self):
        """加载员工数据"""
        # 清空现有数据
        for item in self.employee_tree.get_children():
            self.employee_tree.delete(item)

        # 从数据库加载数据
        employees = self.db.query(Employee).all()
        for idx, emp in enumerate(employees, 1):
            status = "🟢 在职" if emp.status == 1 else "🔴 离职"

            self.employee_tree.insert('', 'end', values=(
                str(idx),
                emp.employee_id,
                emp.name,
                emp.email,
                emp.phone or '',
                status
            ))

        # 重置搜索
        self.search_entry.clear()
        self.search_count_label.configure(text="")
        


    def perform_search(self, search_text):
        """执行搜索"""
        if not search_text:
            self.load_employees()
            return

        # 清空现有数据
        for item in self.employee_tree.get_children():
            self.employee_tree.delete(item)

        # 搜索数据库
        employees = self.db.query(Employee).filter(
            (Employee.name.contains(search_text)) |
            (Employee.employee_id.contains(search_text)) |
            (Employee.email.contains(search_text)) |
            (Employee.phone.contains(search_text))
        ).all()

        # 显示搜索结果
        for idx, emp in enumerate(employees, 1):
            status = "🟢 在职" if emp.status == 1 else "🔴 离职"

            self.employee_tree.insert('', 'end', values=(
                str(idx),
                emp.employee_id,
                emp.name,
                emp.email,
                emp.phone or '',
                status
            ))

        # 更新搜索结果计数
        total_count = self.db.query(Employee).count()
        found_count = len(employees)

        if found_count == total_count:
            self.search_count_label.configure(text="")
        else:
            self.search_count_label.configure(text=f"找到 {found_count}/{total_count} 个员工")

    def import_employees(self):
        """导入员工"""
        file_path = CTKFileDialog.open_file(
            self,
            title='选择Excel文件',
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )

        if not file_path:
            return

        try:
            wb = openpyxl.load_workbook(file_path)
            sheet = wb.active

            success_count = 0
            error_count = 0

            # 从第二行开始读取数据
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row[0]:  # 跳过空行
                    continue

                try:
                    employee = Employee(
                        employee_id=str(row[0]),
                        name=row[1],
                        email=row[2],
                        phone=row[3] if len(row) > 3 else None,
                        status=1
                    )
                    self.db.add(employee)
                    success_count += 1
                except Exception:
                    error_count += 1
                    continue

            self.db.commit()
            self.load_employees()

            CTKMessageBox.show_info(
                self,
                '导入完成',
                f'导入完成！\n成功：{success_count} 条\n失败：{error_count} 条'
            )

        except Exception as e:
            CTKMessageBox.show_error(self, '错误', f'导入失败：\n{str(e)}')
            self.db.rollback()

    def on_select(self, event):
        """处理选择事件"""
        selection = self.employee_tree.selection()
        if selection:
            # 可以在这里添加选中行的相关操作
            pass
    


    def add_employee(self):
        """添加员工"""
        dialog = CTKEmployeeDialog(self)
        self.wait_window(dialog)
        self.load_employees()

    def edit_employee(self):
        """编辑员工"""
        selection = self.employee_tree.selection()
        if not selection:
            CTKMessageBox.show_warning(self, '提示', '请先选择要编辑的员工！')
            return

        try:
            # 获取选中的员工数据
            item = self.employee_tree.item(selection[0])
            employee_id = item['values'][1]  # employee_id在索引1位置（序号后面）

            employee = self.db.query(Employee).filter_by(employee_id=employee_id).first()
            if employee:
                dialog = CTKEmployeeDialog(self, employee)
                self.wait_window(dialog)
                self.load_employees()
            else:
                CTKMessageBox.show_warning(self, '警告', '未找到该员工！')
        except Exception as e:
            CTKMessageBox.show_error(self, '错误', f'编辑员工失败：\n{str(e)}')

    def delete_employee(self):
        """删除员工"""
        selection = self.employee_tree.selection()
        if not selection:
            CTKMessageBox.show_warning(self, '提示', '请先选择要删除的员工！')
            return

        try:
            # 获取选中的员工数据
            item = self.employee_tree.item(selection[0])
            employee_id = item['values'][1]  # employee_id在索引1位置（序号后面）
            employee_name = item['values'][2]  # 员工姓名

            if CTKMessageBox.ask_yes_no(self, '确认', f'确定要删除员工 "{employee_name}" 吗？') != "是":
                return

            employee = self.db.query(Employee).filter_by(employee_id=employee_id).first()
            if employee:
                self.db.delete(employee)
                self.db.commit()
                # 重新加载数据
                self.load_employees()
                CTKMessageBox.show_info(self, '成功', f'成功删除员工 "{employee_name}"！')
            else:
                CTKMessageBox.show_warning(self, '警告', '未找到该员工！')

        except Exception as e:
            CTKMessageBox.show_error(self, '错误', f'删除失败：\n{str(e)}')
            self.db.rollback()


