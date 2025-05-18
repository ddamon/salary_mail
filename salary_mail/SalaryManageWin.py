from salary_mail.db_instance import Employee, SalaryEmail, SalaryRecord


import openpyxl


import base64
import tkinter as tk
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from smtplib import SMTP, SMTP_SSL
from tkinter import ttk


class SalaryManageWin(tk.Toplevel):
    '''工资管理窗口'''

    def __init__(self, parent):
        super(SalaryManageWin, self).__init__()
        self.title('工资管理')
        cx, cy = parent.get_center()
        self.geometry('1200x600+{}+{}'.format(cx - 600, cy - 300))
        self.minsize(900, 600)  # 设置最小窗口大小
        self.attributes("-topmost", 1)
        self.resizable(width=True, height=True)  # 允许调整大小
        self.parent = parent
        self.db = parent.db

        # 修改列配置，优化宽度和显示格式
        self.columns = [
            ('序号', 50),
            ('员工编号', 80),
            ('姓名', 70),
            ('岗位工资', 85),
            ('薪级工资', 85),
            ('绩效工资', 85),
            ('餐补', 70),
            ('交补', 70),
            ('防暑降温费', 90),
            ('补发', 70),
            ('事假扣款', 85),
            ('病假扣款', 85),
            ('其他扣款', 85),
            ('税前工资', 85),
            ('社会保险', 85),
            ('公积金', 80),
            ('个人所得税', 85),
            ('代缴工会会费', 90),
            ('实发工资', 85),
            ('发送状态', 70)
        ]

        self.setupUI()

        # 设置焦点并等待窗口出现
        self.focus_force()
        self.wait_visibility()

        # 取消置顶
        self.attributes("-topmost", 0)

    def setupUI(self):
        """搭建界面"""
        # 创建主框架并添加内边距
        main_frame = ttk.Frame(self, padding="10 5 10 10")
        main_frame.pack(fill='both', expand=True)

        # 工具栏优化
        toolbar = ttk.Frame(main_frame)
        toolbar.pack(fill='x', pady=(0, 10))

        # 左侧工具按钮
        left_tools = ttk.Frame(toolbar)
        left_tools.pack(side=tk.LEFT)

        ttk.Button(left_tools, text="导入工资数据",
                  command=self.import_salary_data, width=15).pack(side=tk.LEFT, padx=(0, 10))

        # 月份选择框优化
        month_frame = ttk.Frame(toolbar)
        month_frame.pack(side=tk.LEFT, padx=20)

        ttk.Label(month_frame, text="选择月份:").pack(side=tk.LEFT, padx=(0, 5))
        self.month_var = tk.StringVar()
        self.month_combo = ttk.Combobox(month_frame, textvariable=self.month_var,
                                      width=8, state='readonly')
        self.month_combo.pack(side=tk.LEFT)

        # 发送按钮（右侧）优化
        send_frame = ttk.Frame(toolbar)
        send_frame.pack(side=tk.RIGHT)

        ttk.Button(send_frame, text="发送选中",
                  command=self.send_selected_salary, width=12).pack(side=tk.RIGHT, padx=5)
        ttk.Button(send_frame, text="全部发送",
                  command=self.send_all_salary, width=12).pack(side=tk.RIGHT, padx=5)

        # 工资列表区域优化
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill='both', expand=True)

        # 创建带样式的表格
        style = ttk.Style()
        style.configure("Treeview",
                       rowheight=25,  # 增加行高
                       background="#ffffff",
                       fieldbackground="#ffffff",
                       foreground="#333333")

        # 设置表头样式
        style.configure("Treeview.Heading",
                       font=('微软雅黑', 9, 'bold'),
                       background="#f0f0f0",
                       relief="flat")

        # 设置选中行样式
        style.map("Treeview",
                 background=[('selected', '#e3f2fd')],
                 foreground=[('selected', '#000000')])

        # 创建水平和垂直滚动条
        v_scrollbar = ttk.Scrollbar(list_frame, orient="vertical")
        h_scrollbar = ttk.Scrollbar(list_frame, orient="horizontal")

        # 配置列表和滚动条
        self.salary_list = ttk.Treeview(
            list_frame,
            columns=[col for col, _ in self.columns],
            show='headings',
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            style="Treeview"
        )

        # 设置列宽和标题
        for col, width in self.columns:
            self.salary_list.heading(col, text=col)
            # 根据列类型设置不同的对齐方式
            if col in ['序号', '员工编号', '姓名', '发送状态']:
                anchor = 'center'
            else:
                anchor = 'e'
            self.salary_list.column(col, width=width, anchor=anchor, minwidth=width)

        # 配置滚动条
        v_scrollbar.config(command=self.salary_list.yview)
        h_scrollbar.config(command=self.salary_list.xview)

        # 使用grid布局优化表格和滚动条的位置
        self.salary_list.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')

        # 配置grid权重
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        # 加载月份列表和数据
        self.load_salary_months()

        # 绑定选择事件
        self.salary_list.bind('<<TreeviewSelect>>', self.on_select)

        # 添加工具提示
        self.tooltip = None
        self.salary_list.bind('<Motion>', self.show_tooltip)
        self.salary_list.bind('<Leave>', self.hide_tooltip)

    def load_salary_months(self):
        """加载工资月份列表"""
        months = self.db.query(SalaryRecord.salary_month)\
            .distinct().order_by(SalaryRecord.salary_month.desc()).all()
        month_list = [m[0] for m in months]
        self.month_combo['values'] = month_list

        if month_list:
            self.month_var.set(month_list[0])
            self.load_salary_data(month_list[0])

    def load_salary_data(self, month):
        """加载指定月份的工资数据"""
        # 清空现有数据
        for item in self.salary_list.get_children():
            self.salary_list.delete(item)

        # 查询数据
        records = self.db.query(SalaryRecord, Employee)\
            .join(Employee, SalaryRecord.employee_id == Employee.employee_id)\
            .filter(SalaryRecord.salary_month == month).all()

        # 优化数据显示格式
        for idx, (record, emp) in enumerate(records, 1):
            status_map = {
                0: '⏳ 待发送',
                1: '✓ 已发送',
                2: '✗ 失败'
            }

            # 格式化金额显示
            def format_money(value):
                try:
                    return f"{float(value):,.2f}" if value else "0.00"
                except:
                    return "0.00"

            # 插入数据并保存返回的item ID
            item_id = self.salary_list.insert('', 'end', values=(
                str(idx),
                emp.employee_id,
                emp.name,
                format_money(record.post_salary),
                format_money(record.level_salary),
                format_money(record.performance),
                format_money(record.meal_allowance),
                format_money(record.traffic_allowance),
                format_money(record.cooling_allowance),
                format_money(record.additional_payment),
                format_money(record.leave_deduct),
                format_money(record.sick_deduct),
                format_money(record.other_deduct),
                format_money(record.pre_tax_salary),
                format_money(record.insurance),
                format_money(record.house_fund),
                format_money(record.tax),
                format_money(record.union_fee),
                format_money(record.actual_salary),
                status_map.get(record.send_status, '未知')
            ))

            # 设置交替行颜色
            if idx % 2 == 0:
                self.salary_list.tag_configure('evenrow', background='#f8f9fa')
                self.salary_list.item(item_id, tags=('evenrow',))

            # 根据发送状态设置行的标签颜色
            if record.send_status == 1:  # 发送成功
                self.salary_list.tag_configure('success', foreground='#2e7d32')
                self.salary_list.item(item_id, tags=('success', 'evenrow' if idx % 2 == 0 else ''))
            elif record.send_status == 2:  # 发送失败
                self.salary_list.tag_configure('failed', foreground='#c62828')
                self.salary_list.item(item_id, tags=('failed', 'evenrow' if idx % 2 == 0 else ''))

    def on_month_selected(self, event):
        """月份选择改变时的处理"""
        month = self.month_var.get()
        if month:
            self.load_salary_data(month)

    def import_salary_data(self):
        """导入工资数据"""
        file_path = tk.filedialog.askopenfilename(
            title='选择工资数据文件',
            filetypes=[("Excel File", "*.xlsx *.xls")],
            parent=self
        )

        if not file_path:
            return

        try:
            wb = openpyxl.load_workbook(file_path)
            sheet = wb.active

            # 获取表头
            headers = [str(cell.value).strip() if cell.value else '' for cell in sheet[1]]

            # 验证必要的列是否存在
            required_fields = [
                '员工编号', '姓名', '岗位工资', '薪级工资', '绩效工资',
                '餐补', '交补', '防暑降温费', '补发', '事假扣款', '病假扣款',
                '其他扣款', '税前工资', '社会保险', '公积金', '个人所得税',
                '代缴工会会费', '实发工资', '备注', '月份'
            ]

            # 验证所有必要的列是否存在
            missing_fields = [f for f in required_fields if f not in headers]
            if missing_fields:
                raise ValueError(f"缺少必要的列: {', '.join(missing_fields)}")

            # 获取列索引
            field_indexes = {field: headers.index(field) for field in required_fields}

            # 字段映射
            field_mapping = {
                '员工编号': 'employee_id',
                '岗位工资': 'post_salary',
                '薪级工资': 'level_salary',
                '绩效工资': 'performance',
                '餐补': 'meal_allowance',
                '交补': 'traffic_allowance',
                '防暑降温费': 'cooling_allowance',
                '补发': 'additional_payment',
                '事假扣款': 'leave_deduct',
                '病假扣款': 'sick_deduct',
                '其他扣款': 'other_deduct',
                '税前工资': 'pre_tax_salary',
                '社会保险': 'insurance',
                '公积金': 'house_fund',
                '个人所得税': 'tax',
                '代缴工会会费': 'union_fee',
                '实发工资': 'actual_salary',
                '备注': 'remark',
                '月份': 'salary_month'
            }

            # 需要验证格式的字段
            numeric_fields = [
                '岗位工资', '薪级工资', '绩效工资', '餐补', '交补',
                '防暑降温费', '补发', '事假扣款', '病假扣款', '其他扣款',
                '税前工资', '社会保险', '公积金', '个人所得税',
                '代缴工会会费', '实发工资'
            ]

            # 开始导入数据
            success_count = 0
            error_count = 0
            error_msgs = []

            # 从第二行开始读取数据
            for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
                if not any(row):  # 跳过空行
                    continue

                try:
                    # 获取月份
                    salary_month = str(row[field_indexes['月份']]).strip()
                    if not salary_month:
                        raise ValueError("月份不能为空")

                    # 处理月份格式
                    salary_month = salary_month.replace('-', '')[:6]  # 确保格式为YYYYMM

                    employee_id = str(row[field_indexes['员工编号']]).strip()
                    if not employee_id:
                        raise ValueError("员工编号不能为空")

                    # 检查员工是否存在
                    employee = self.db.query(Employee)\
                        .filter_by(employee_id=employee_id, status=1).first()
                    if not employee:
                        # 如果员工不存在，自动创建
                        employee = Employee(
                            employee_id=employee_id,
                            name=str(row[field_indexes['姓名']]).strip(),
                            email='',  # 需要后续补充
                            status=1
                        )
                        self.db.add(employee)

                    # 检查是否已存在该月工资记录
                    existing = self.db.query(SalaryRecord)\
                        .filter_by(employee_id=employee_id, salary_month=salary_month).first()
                    if existing:
                        record = existing
                    else:
                        record = SalaryRecord(
                            employee_id=employee_id,
                            salary_month=salary_month
                        )

                    # 设置工资数据
                    for field, index in field_indexes.items():
                        if field in field_mapping:
                            value = row[index]
                            if field in numeric_fields:
                                value = self._validate_number(value, field)
                            else:
                                value = str(value).strip() if value is not None else ''
                            setattr(record, field_mapping[field], value)

                    record.send_status = 0  # 重置发送状态
                    self.db.add(record)
                    success_count += 1

                except Exception as e:
                    error_count += 1
                    error_msgs.append(f"第 {row_idx} 行: {str(e)}")
                    continue

            self.db.commit()
            self.load_salary_months()

            # 显示导入结果
            result_msg = f"导入完成!\n成功: {success_count} 条\n失败: {error_count} 条"
            if error_msgs:
                result_msg += "\n\n失败详情:\n" + "\n".join(error_msgs[:10])
                if len(error_msgs) > 10:
                    result_msg += f"\n... 等共 {len(error_msgs)} 条错误"

            tk.messagebox.showinfo("导入结果", result_msg, parent=self)

        except Exception as e:
            tk.messagebox.showerror("导入失败", f"导入过程出错：\n{str(e)}", parent=self)
            self.db.rollback()

    def _validate_number(self, value, field_name):
        """验证数值格式"""
        if not value:
            return '0.00'
        try:
            value = str(value).strip()
            float_val = float(value)
            return f"{float_val:.2f}"
        except ValueError:
            raise ValueError(f"{field_name}格式不正确: {value}")

    def send_selected_salary(self):
        """发送选中的工资条"""
        selection = self.salary_list.selection()
        if not selection:
            tk.messagebox.showwarning('提示', '请先选择要发送的记录！', parent=self)
            return

        self._send_salary_records(selection)

    def send_all_salary(self):
        """发送所有工资条"""
        selection = self.salary_list.get_children()
        if not selection:
            tk.messagebox.showwarning('提示', '当前月份没有工资记录！', parent=self)
            return

        if not tk.messagebox.askyesno('确认', '确定要发送所有工资条吗？', parent=self):
            return

        self._send_salary_records(selection)

    def _send_salary_records(self, items):
        """发送工资条的具体实现"""
        try:
            # 计算窗口位置
            window_width = 400
            window_height = 180
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2
            
            # 创建进度条窗口并直接设置位置
            progress_window = tk.Toplevel(self)
            progress_window.title("发送进度")
            progress_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
            progress_window.transient(self)  # 设置为主窗口的临时窗口
            progress_window.grab_set()  # 模态窗口
            
            # 主框架
            main_frame = ttk.Frame(progress_window, padding="20 10")
            main_frame.pack(fill='both', expand=True)
            
            # 进度标签 - 使用大字体显示进度
            progress_label = ttk.Label(
                main_frame, 
                text="正在发送...", 
                font=('Microsoft YaHei UI', 12, 'bold')
            )
            progress_label.pack(pady=(0, 10))
            
            # 进度条框架
            progress_frame = ttk.Frame(main_frame)
            progress_frame.pack(fill='x', pady=5)
            
            # 进度百分比标签
            percent_label = ttk.Label(
                progress_frame, 
                text="0%",
                font=('Microsoft YaHei UI', 9)
            )
            percent_label.pack(side='right', padx=(5, 0))
            
            # 进度条
            progress_bar = ttk.Progressbar(
                progress_frame, 
                mode='determinate',
                length=300
            )
            progress_bar.pack(side='left', fill='x', expand=True)
            
            # 详细信息标签
            detail_frame = ttk.Frame(main_frame)
            detail_frame.pack(fill='x', pady=10)
            
            detail_label = ttk.Label(
                detail_frame, 
                text="准备发送...",
                font=('Microsoft YaHei UI', 9)
            )
            detail_label.pack(side='left')
            
            # 成功/失败计数标签
            count_label = ttk.Label(
                detail_frame,
                text="成功: 0  失败: 0",
                font=('Microsoft YaHei UI', 9)
            )
            count_label.pack(side='right')

            # 获取邮件配置
            sender = self.db.query(SalaryEmail).filter(SalaryEmail.field_name=='sender').first()
            password = self.db.query(SalaryEmail).filter(SalaryEmail.field_name=='password').first()
            sender_name = self.db.query(SalaryEmail).filter(SalaryEmail.field_name=='sender_name').first()
            smtp_server = self.db.query(SalaryEmail).filter(SalaryEmail.field_name=='smtp_server').first()
            port = self.db.query(SalaryEmail).filter(SalaryEmail.field_name=='port').first()
            template = self.db.query(SalaryEmail).filter(SalaryEmail.field_name=='email_template').first()

            if not all([sender, password, sender_name, smtp_server, port, template]):
                tk.messagebox.showerror('错误', '请先完成邮件配置！', parent=self)
                progress_window.destroy()
                return

            # 解密密码
            try:
                if isinstance(password.field_value, bytes):
                    decoded_password = base64.decodebytes(password.field_value).decode()
                else:
                    decoded_password = base64.decodebytes(password.field_value.encode()).decode()
            except Exception as e:
                raise ValueError(f"密码解密失败: {str(e)}")

            # 连接SMTP服务器
            if int(port.field_value) == 25:
                smtp = SMTP(host=smtp_server.field_value, port=int(port.field_value))
            else:
                smtp = SMTP_SSL(host=smtp_server.field_value, port=int(port.field_value))

            smtp.login(sender.field_value, decoded_password)

            # 获取当前月份
            month = self.month_var.get()

            success_count = 0
            error_count = 0
            total_count = len(items)
            
            # 设置进度条最大值
            progress_bar['maximum'] = total_count

            # 开始发送
            for index, item in enumerate(items, 1):
                try:
                    values = self.salary_list.item(item)['values']
                    employee_id = values[1]
                    employee_name = values[2]

                    # 更新进度显示
                    progress = (index / total_count) * 100
                    progress_label.config(text=f"正在发送 ({index}/{total_count})")
                    percent_label.config(text=f"{progress:.1f}%")
                    detail_label.config(text=f"正在处理: {employee_name}")
                    count_label.config(text=f"成功: {success_count}  失败: {error_count}")
                    progress_bar['value'] = index
                    progress_window.update()

                    # 获取工资记录和员工信息
                    record = self.db.query(SalaryRecord)\
                        .filter_by(employee_id=employee_id, salary_month=month).first()
                    employee = self.db.query(Employee)\
                        .filter_by(employee_id=employee_id).first()

                    if not record or not employee:
                        continue

                    # 生成邮件内容
                    msg = self._make_salary_email(
                        template.field_value, record, employee,
                        sender.field_value, sender_name.field_value
                    )

                    # 发送邮件
                    smtp.sendmail(
                        from_addr=sender.field_value,
                        to_addrs=[employee.email],
                        msg=msg.as_string()
                    )

                    # 更新发送状态
                    record.send_status = 1
                    record.send_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.db.add(record)
                    success_count += 1
                    count_label.config(text=f"成功: {success_count}  失败: {error_count}")

                except Exception as e:
                    print(f"发送失败: {str(e)}")
                    error_count += 1
                    count_label.config(text=f"成功: {success_count}  失败: {error_count}")

                # 每发送一封邮件就提交一次
                self.db.commit()

            smtp.quit()

            # 刷新显示
            self.load_salary_data(month)

            # 关闭进度窗口
            progress_window.destroy()

            # 显示结果
            tk.messagebox.showinfo(
                '发送完成',
                f'发送完成！\n成功：{success_count} 条\n失败：{error_count} 条',
                parent=self
            )

        except Exception as e:
            if 'progress_window' in locals():
                progress_window.destroy()
            tk.messagebox.showerror('错误', f'发送过程出错：\n{str(e)}', parent=self)
            self.db.rollback()

    def _make_salary_email(self, template, record, employee, sender, sender_name):
        """生成工资邮件内容"""
        year = record.salary_month[:4]
        month = str(int(record.salary_month[4:]))  # 去掉月份前导零
        # 确保模板中的大括号被正确转义
        template = template.replace('{', '{{').replace('}', '}}')
        # 恢复需要替换的变量占位符
        template = template.replace('{{name}}', '{name}')\
                            .replace('{{year}}', '{year}')\
                            .replace('{{month}}', '{month}')\
                            .replace('{{salary_items}}', '{salary_items}')

        # 构建工资项目表格
        salary_items = []

        # 基本工资项
        salary_items.append("<tr><td>岗位工资</td><td>{}</td></tr>".format(record.post_salary))
        salary_items.append("<tr><td>薪级工资</td><td>{}</td></tr>".format(record.level_salary))
        salary_items.append("<tr><td>绩效工资</td><td>{}</td></tr>".format(record.performance))

        # 补贴项
        salary_items.append("<tr><td>餐补</td><td>{}</td></tr>".format(record.meal_allowance))
        salary_items.append("<tr><td>交补</td><td>{}</td></tr>".format(record.traffic_allowance))
        salary_items.append("<tr><td>防暑降温费</td><td>{}</td></tr>".format(record.cooling_allowance))
        salary_items.append("<tr><td>补发</td><td>{}</td></tr>".format(record.additional_payment))

        # 扣款项
        salary_items.append("<tr><td>事假扣款</td><td>{}</td></tr>".format(record.leave_deduct))
        salary_items.append("<tr><td>病假扣款</td><td>{}</td></tr>".format(record.sick_deduct))
        salary_items.append("<tr><td>其他扣款</td><td>{}</td></tr>".format(record.other_deduct))

        # 税前工资
        salary_items.append("<tr><td>税前工资</td><td>{}</td></tr>".format(record.pre_tax_salary))

        # 各类扣缴
        salary_items.append("<tr><td>社会保险</td><td>{}</td></tr>".format(record.insurance))
        salary_items.append("<tr><td>公积金</td><td>{}</td></tr>".format(record.house_fund))
        salary_items.append("<tr><td>个人所得税</td><td>{}</td></tr>".format(record.tax))
        salary_items.append("<tr><td>代缴工会会费</td><td>{}</td></tr>".format(record.union_fee))

        # 实发工资（高亮显示）
        salary_items.append("<tr style=\"background-color:#e8f5e9;font-weight:bold\"><td>实发工资</td><td>{}</td></tr>".format(record.actual_salary))

        # 如果有备注，添加备注行
        if record.remark:
            salary_items.append(
                "<tr><td colspan=\"2\" style=\"background-color:#fff3e0;padding:15px\">"
                "<div style=\"font-weight:bold;margin-bottom:10px\">备注</div>{}</td></tr>".format(record.remark)
            )

        # 生成邮件内容
        try:
            email_content = template.format(
                name=employee.name,
                year=year,
                month=month,
                salary_items='\n'.join(salary_items)
            )
        except Exception as e:
            raise ValueError(f"邮件内容生成失败: {str(e)}\n模板内容: {template}")
        # 创建邮件对象
        msg = MIMEMultipart()
        msg['From'] = formataddr([sender_name, sender])
        msg['To'] = formataddr([employee.name, employee.email])
        msg['Subject'] = f"{year}年{month}月工资明细"

        # 添加HTML内容
        msg.attach(MIMEText(email_content, 'html'))

        return msg

    def cancel(self):
        """取消编辑"""
        res = tk.messagebox.askyesno(title='是否取消设置？', message="设置的内容未保存，是否退出？", parent=self)
        if res:
            self.parent.open_windows['salary'] = None
            self.destroy()

    def on_select(self, event):
        """处理选择事件"""
        selection = self.salary_list.selection()
        if selection:
            # 可以在这里添加选中行的相关操作
            pass

    def show_tooltip(self, event):
        """显示工具提示"""
        item = self.salary_list.identify_row(event.y)
        if item:
            # 获取鼠标所在列
            column = self.salary_list.identify_column(event.x)
            if not column:
                return

            # 获取列标题
            col_num = int(column[1]) - 1
            col_name = self.salary_list.heading(column)['text']

            # 获取单元格值
            cell_value = self.salary_list.item(item)['values'][col_num]

            # 创建工具提示
            if self.tooltip:
                self.tooltip.destroy()

            x, y, _, _ = self.salary_list.bbox(item, column)
            x += self.salary_list.winfo_rootx() + 25
            y += self.salary_list.winfo_rooty() + 20

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