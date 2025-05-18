from salary_mail.db_instance import Employee


import openpyxl


import tkinter as tk
from tkinter import ttk


class EmployeeDialog(tk.Toplevel):
    """员工信息编辑对话框"""

    def __init__(self, parent, employee=None):
        super(EmployeeDialog, self).__init__()
        self.title('编辑员工' if employee else '添加员工')
        self.geometry('400x350')
        self.resizable(False, False)

        self.parent = parent
        self.db = parent.db
        self.employee = employee

        self.setupUI()

    def setupUI(self):
        main_frame = tk.Frame(self, padx=20, pady=10)
        main_frame.pack(fill='both', expand=True)

        # 表单字段
        fields = [
            ('职工编号：', 'employee_id', True),
            ('姓名：', 'name', True),
            ('邮箱：', 'email', True),
            ('手机号码：', 'phone', False)
        ]

        self.entries = {}

        for label, field, required in fields:
            row = tk.Frame(main_frame)
            row.pack(fill='x', pady=5)
            tk.Label(row, text=('*' if required else '') + label, width=12).pack(side=tk.LEFT)
            entry = tk.Entry(row, width=30)
            entry.pack(side=tk.LEFT, padx=5)
            self.entries[field] = entry

            # 如果是编辑模式，填充现有数据
            if self.employee:
                entry.insert(0, getattr(self.employee, field) or '')

        # 状态选择
        status_frame = tk.Frame(main_frame)
        status_frame.pack(fill='x', pady=5)
        tk.Label(status_frame, text='状态：', width=12).pack(side=tk.LEFT)
        self.status_var = tk.IntVar(value=self.employee.status if self.employee else 1)
        tk.Radiobutton(status_frame, text='在职', variable=self.status_var, value=1).pack(side=tk.LEFT)
        tk.Radiobutton(status_frame, text='离职', variable=self.status_var, value=0).pack(side=tk.LEFT)

        # 按钮
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=10)
        tk.Button(btn_frame, text='取消', width=10, command=self.destroy).pack(side=tk.RIGHT, padx=5)
        tk.Button(btn_frame, text='保存', width=10, command=self.save).pack(side=tk.RIGHT, padx=5)

        # 添加说明文本
        note = tk.Label(main_frame, text='注：* 表示必填字段', fg='gray')
        note.pack(pady=10)

    def save(self):
        """保存员工信息"""
        try:
            # 验证必填字段
            for field in ['employee_id', 'name', 'email']:
                if not self.entries[field].get().strip():
                    tk.messagebox.showwarning('提示', f'{field}不能为空！', parent=self)
                    return

            # 验证邮箱格式
            email = self.entries['email'].get().strip()
            if not re.match(r'^[\w\.-]+@[\w\.-]+\.\w+$', email):
                tk.messagebox.showwarning('提示', '请输入有效的邮箱地址！', parent=self)
                return

            # 检查职工编号是否已存在
            employee_id = self.entries['employee_id'].get().strip()
            existing = self.db.query(Employee).filter_by(employee_id=employee_id).first()
            if existing and (not self.employee or existing.id != self.employee.id):
                tk.messagebox.showwarning('提示', '该职工编号已存在！', parent=self)
                return

            # 检查邮箱是否已存在
            existing = self.db.query(Employee).filter_by(email=email).first()
            if existing and (not self.employee or existing.id != self.employee.id):
                tk.messagebox.showwarning('提示', '该邮箱已存在！', parent=self)
                return

            # 保存数据
            if not self.employee:
                self.employee = Employee()

            for field, entry in self.entries.items():
                setattr(self.employee, field, entry.get().strip())
            self.employee.status = self.status_var.get()

            self.db.add(self.employee)
            self.db.commit()

            tk.messagebox.showinfo('成功', '保存成功！', parent=self)
            self.destroy()

        except Exception as e:
            tk.messagebox.showerror('错误', f'保存失败：\n{str(e)}', parent=self)
            self.db.rollback()


class EmployeeManageWin(tk.Toplevel):
    '''员工管理窗口'''

    def __init__(self, parent):
        super(EmployeeManageWin, self).__init__()
        self.title('员工管理')
        cx, cy = parent.get_center()
        self.geometry('900x600+{}+{}'.format(cx - 450, cy - 300))
        self.minsize(800, 500)  # 设置最小窗口大小
        self.attributes("-topmost", 1)
        self.resizable(width=True, height=True)
        self.parent = parent
        self.db = parent.db

        # 定义列配置
        self.columns = [
            ('序号', 60),
            ('职工编号', 100),
            ('姓名', 80),
            ('邮箱', 200),
            ('手机号码', 120),
            ('状态', 80)
        ]

        # 添加搜索相关变量
        self.search_after_id = None  # 用于延迟搜索
        self.search_count_var = tk.StringVar()  # 用于显示搜索结果数量

        self.setupUI()

        # 设置焦点并等待窗口出现
        self.focus_force()
        self.wait_visibility()
        # 取消置顶
        self.attributes("-topmost", 0)

    def setupUI(self):
        # 创建主框架并添加内边距
        main_frame = ttk.Frame(self, padding="10 5 10 10")
        main_frame.pack(fill='both', expand=True)

        # 工具栏优化
        toolbar = ttk.Frame(main_frame)
        toolbar.pack(fill='x', pady=(0, 10))

        # 左侧按钮组
        left_tools = ttk.Frame(toolbar)
        left_tools.pack(side=tk.LEFT)

        # 创建带图标的按钮
        ttk.Button(left_tools, text="导入员工",
                  command=self.import_employees, width=12).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(left_tools, text="添加员工",
                  command=self.add_employee, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(left_tools, text="修改",
                  command=self.edit_employee, width=8).pack(side=tk.LEFT, padx=5)
        ttk.Button(left_tools, text="删除",
                  command=self.delete_employee, width=8).pack(side=tk.LEFT, padx=5)

        # 搜索框（右侧）优化
        search_frame = ttk.Frame(toolbar)
        search_frame.pack(side=tk.RIGHT)

        # 添加搜索结果计数标签
        self.search_count_label = ttk.Label(search_frame, textvariable=self.search_count_var)
        self.search_count_label.pack(side=tk.RIGHT, padx=5)

        # 搜索输入框
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.on_search_changed)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=20)
        search_entry.pack(side=tk.RIGHT, padx=(5, 0))

        # 添加搜索提示
        search_entry.insert(0, "输入关键字搜索...")
        search_entry.config(foreground='gray')

        def on_focus_in(event):
            if search_entry.get() == "输入关键字搜索...":
                search_entry.delete(0, tk.END)
                search_entry.config(foreground='black')

        def on_focus_out(event):
            if not search_entry.get():
                search_entry.insert(0, "输入关键字搜索...")
                search_entry.config(foreground='gray')

        search_entry.bind('<FocusIn>', on_focus_in)
        search_entry.bind('<FocusOut>', on_focus_out)

        ttk.Label(search_frame, text="搜索:").pack(side=tk.RIGHT)

        # 员工列表区域
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill='both', expand=True)

        # 创建带样式的表格
        style = ttk.Style()
        style.configure("Treeview",
                       rowheight=25,
                       background="#ffffff",
                       fieldbackground="#ffffff",
                       foreground="#333333")

        style.configure("Treeview.Heading",
                       font=('微软雅黑', 9, 'bold'),
                       background="#f0f0f0",
                       relief="flat")

        style.map("Treeview",
                 background=[('selected', '#e3f2fd')],
                 foreground=[('selected', '#000000')])

        # 创建滚动条
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill='y')

        # 创建表格
        self.employee_list = ttk.Treeview(
            list_frame,
            columns=[col for col, _ in self.columns],
            show="headings",
            yscrollcommand=scrollbar.set,
            style="Treeview"
        )

        # 设置列
        for col, width in self.columns:
            self.employee_list.column(col, width=width, anchor='center')
            self.employee_list.heading(col, text=col)

        self.employee_list.pack(side=tk.LEFT, fill='both', expand=True)
        scrollbar.config(command=self.employee_list.yview)

        # 添加工具提示
        self.tooltip = None
        self.employee_list.bind('<Motion>', self.show_tooltip)
        self.employee_list.bind('<Leave>', self.hide_tooltip)

        # 加载员工数据
        self.load_employees()

    def load_employees(self):
        """加载员工数据"""
        # 清空现有数据
        for item in self.employee_list.get_children():
            self.employee_list.delete(item)

        # 从数据库加载数据
        employees = self.db.query(Employee).all()
        for idx, emp in enumerate(employees, 1):
            status = "✓ 在职" if emp.status == 1 else "✗ 离职"

            item_id = self.employee_list.insert('', 'end', values=(
                str(idx),
                emp.employee_id,
                emp.name,
                emp.email,
                emp.phone or '',
                status
            ))

            # 设置交替行颜色
            if idx % 2 == 0:
                self.employee_list.tag_configure('evenrow', background='#f8f9fa')
                self.employee_list.item(item_id, tags=('evenrow',))

            # 设置状态颜色
            if emp.status == 1:
                self.employee_list.tag_configure('active', foreground='#2e7d32')
                self.employee_list.item(item_id, tags=('active', 'evenrow' if idx % 2 == 0 else ''))
            else:
                self.employee_list.tag_configure('inactive', foreground='#c62828')
                self.employee_list.item(item_id, tags=('inactive', 'evenrow' if idx % 2 == 0 else ''))

        # 重置搜索
        self.search_var.set("")
        self.search_count_var.set("")

    def on_search_changed(self, *args):
        """处理搜索文本变化"""
        # 取消之前的延迟搜索
        if self.search_after_id:
            self.after_cancel(self.search_after_id)

        # 设置300ms的延迟后执行搜索
        self.search_after_id = self.after(300, self.perform_search)

    def perform_search(self):
        """执行搜索"""
        search_text = self.search_var.get().lower()
        if search_text == "输入关键字搜索...":
            search_text = ""

        visible_count = 0
        all_items = self.employee_list.get_children()

        if not search_text:
            # 如果搜索框为空，显示所有项目
            for item in all_items:
                self.employee_list.reattach(item, '', 'end')
            visible_count = len(all_items)
        else:
            # 执行搜索
            for item in all_items:
                values = self.employee_list.item(item)['values']
                # 改进搜索逻辑，支持多个关键字
                keywords = search_text.split()
                match = True
                for keyword in keywords:
                    if not any(str(value).lower().find(keyword) >= 0 for value in values):
                        match = False
                        break

                if match:
                    self.employee_list.reattach(item, '', 'end')
                    visible_count += 1
                else:
                    self.employee_list.detach(item)

        # 更新搜索结果计数
        if visible_count == len(all_items):
            self.search_count_var.set("")
        else:
            total = len(all_items)
            self.search_count_var.set(f"显示 {visible_count}/{total}")

        # 如果没有匹配项，显示提示
        if visible_count == 0 and search_text:
            no_result_id = self.employee_list.insert('', 'end', values=('', '', '没有找到匹配的员工', '', '', ''))
            self.employee_list.tag_configure('no_result', foreground='gray')
            self.employee_list.item(no_result_id, tags=('no_result',))

    def import_employees(self):
        """导入员工"""
        # 临时取消窗口置顶
        self.attributes("-topmost", 0)

        file_path = tk.filedialog.askopenfilename(
            title='选择Excel文件',
            filetypes=[("Excel File", "*.xlsx *.xls")],
            parent=self  # 设置父窗口
        )

        # 恢复窗口焦点
        self.focus_force()

        if not file_path:
            return

        try:
            wb = openpyxl.load_workbook(file_path)
            sheet = wb.active

            # 跳过标题行
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row[0]:  # 跳过空行
                    continue

                employee = Employee(
                    employee_id=str(row[0]),
                    name=row[1],
                    email=row[2],
                    phone=row[3] if len(row) > 3 else None,
                    status=1
                )
                self.db.add(employee)

            self.db.commit()
            self.load_employees()
            tk.messagebox.showinfo('成功', '员工数据导入成功！', parent=self)

        except Exception as e:
            tk.messagebox.showerror('错误', f'导入失败：\n{str(e)}', parent=self)
            self.db.rollback()

    def add_employee(self):
        """添加员工"""
        dialog = EmployeeDialog(self)
        self.wait_window(dialog)
        self.load_employees()

    def edit_employee(self):
        """编辑员工"""
        selection = self.employee_list.selection()
        if not selection:
            tk.messagebox.showwarning('提示', '请先选择要编辑的员工！', parent=self)
            return

        # 获取选中的员工数据
        item = self.employee_list.item(selection[0])
        employee_id = item['values'][1]

        employee = self.db.query(Employee).filter_by(employee_id=employee_id).first()
        if employee:
            dialog = EmployeeDialog(self, employee)
            self.wait_window(dialog)
            self.load_employees()

    def delete_employee(self):
        """删除员工"""
        selection = self.employee_list.selection()
        if not selection:
            tk.messagebox.showwarning('提示', '请先选择要删除的员工！', parent=self)
            return

        if not tk.messagebox.askyesno('确认', '确定要删除选中的员工吗？', parent=self):
            return

        try:
            for item in selection:
                email = self.employee_list.item(item)['values'][2]
                employee = self.db.query(Employee).filter_by(email=email).first()
                if employee:
                    self.db.delete(employee)

            self.db.commit()
            self.load_employees()
            tk.messagebox.showinfo('成功', '删除成功！', parent=self)

        except Exception as e:
            tk.messagebox.showerror('错误', f'删除失败：\n{str(e)}', parent=self)
            self.db.rollback()

    def show_tooltip(self, event):
        """显示工具提示"""
        item = self.employee_list.identify_row(event.y)
        if item:
            # 获取鼠标所在列
            column = self.employee_list.identify_column(event.x)
            if not column:
                return

            # 获取列标题
            col_num = int(column[1]) - 1
            col_name = self.employee_list.heading(column)['text']

            # 获取单元格值
            cell_value = self.employee_list.item(item)['values'][col_num]

            # 创建工具提示
            if self.tooltip:
                self.tooltip.destroy()

            x, y, _, _ = self.employee_list.bbox(item, column)
            x += self.employee_list.winfo_rootx() + 25
            y += self.employee_list.winfo_rooty() + 20

            self.tooltip = tk.Toplevel(self)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")

            # 根据列类型设置不同的提示内容
            if col_name == '邮箱':
                tooltip_text = f"邮箱: {cell_value}"
            elif col_name == '手机号码':
                tooltip_text = f"手机: {cell_value}"
            else:
                tooltip_text = str(cell_value)

            label = ttk.Label(self.tooltip, text=tooltip_text,
                            background="#ffffe0", relief="solid", borderwidth=1)
            label.pack()

    def hide_tooltip(self, event):
        """隐藏工具提示"""
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None