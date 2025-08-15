# coding:utf-8
"""
CustomTkinter版本的工资管理窗口
"""
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import openpyxl
import base64
import threading
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from smtplib import SMTP, SMTP_SSL

from salary_mail.db_instance import Employee, SalaryEmail, SalaryRecord
from salary_mail.ctk_components import CTKTreeview, CTKMessageBox, CTKFileDialog, CTKProgressDialog, CTKWindowSizeManager
from salary_mail.theme_config import theme_manager as responsive_theme_manager

class CTKSalaryManageWin(ctk.CTkToplevel):
    """工资管理窗口"""
    
    def __init__(self, parent):
        super().__init__(parent)
        
        self.title('工资管理')
        
        # 设置响应式窗口配置
        self.responsive_config = responsive_theme_manager.setup_responsive_window(self, 'main_window')
        
        # 设置为全屏显示
        self.set_fullscreen_window()
        
        # 设置窗口属性
        self.transient(parent)
        
        self.parent = parent
        self.db = parent.db
        
        # 定义列配置
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
            print(f"设置工资管理窗口全屏失败: {e}")
            # 备用方案：使用响应式配置的主窗口尺寸
            config = self.responsive_config
            if config and 'main_window' in config:
                window_config = config['main_window']
                width = window_config.get('width', 1200)
                height = window_config.get('height', 800)
                self.geometry(f"{width}x{height}")
                self.center_window()
    
    def center_window(self):
        """窗口居中显示 - 支持高DPI缩放"""
        responsive_theme_manager.center_window_on_screen(self)
    
    def setup_ui(self):
        """设置响应式UI"""
        # 获取响应式配置
        padding = responsive_theme_manager.get_size('padding_medium')
        
        # 主容器
        main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=padding, pady=padding)
        
        # 顶部工具栏
        self.create_toolbar(main_frame)
        
        # 工资列表
        self.create_salary_list(main_frame)
        
        # 加载月份列表
        self.load_salary_months()
    
    def create_toolbar(self, parent):
        """创建工具栏"""
        toolbar_frame = ctk.CTkFrame(parent, height=70, corner_radius=10)
        toolbar_frame.pack(fill="x", pady=(0, 15))
        toolbar_frame.pack_propagate(False)
        
        # 工具栏内容
        toolbar_content = ctk.CTkFrame(toolbar_frame, fg_color="transparent")
        toolbar_content.pack(fill="both", expand=True, padx=15, pady=10)
        
        # 左侧按钮
        left_frame = ctk.CTkFrame(toolbar_content, fg_color="transparent")
        left_frame.pack(side="left", fill="y")
        
        # 导入工资数据按钮
        import_btn = ctk.CTkButton(
            left_frame,
            text="📊 导入工资数据",
            command=self.import_salary_data,
            width=150,
            height=45,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold")
        )
        import_btn.pack(side="left", padx=(0, 15))
        
        # 中间月份选择
        month_frame = ctk.CTkFrame(toolbar_content, fg_color="transparent")
        month_frame.pack(side="left", fill="y", padx=20)
        
        month_label = ctk.CTkLabel(
            month_frame,
            text="选择月份：",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold")
        )
        month_label.pack(side="left", padx=(0, 10))
        
        self.month_var = tk.StringVar()
        self.month_combo = ctk.CTkComboBox(
            month_frame,
            variable=self.month_var,
            command=self.on_month_selected,
            width=120,
            height=35,
            state="readonly",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14)
        )
        self.month_combo.pack(side="left")
        
        # 右侧发送按钮
        send_frame = ctk.CTkFrame(toolbar_content, fg_color="transparent")
        send_frame.pack(side="right", fill="y")
        
        # 发送选中按钮
        send_selected_btn = ctk.CTkButton(
            send_frame,
            text="📧 发送选中",
            command=self.send_selected_salary,
            width=120,
            height=45,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            fg_color="#4CAF50",
            hover_color="#45a049"
        )
        send_selected_btn.pack(side="right", padx=5)
        
        # 全部发送按钮
        send_all_btn = ctk.CTkButton(
            send_frame,
            text="📬 全部发送",
            command=self.send_all_salary,
            width=120,
            height=45,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            fg_color="#2196F3",
            hover_color="#1976D2"
        )
        send_all_btn.pack(side="right", padx=5)
    
    def create_salary_list(self, parent):
        """创建工资列表"""
        # 获取当前主题模式
        appearance_mode = ctk.get_appearance_mode()
        
        # 列表容器 - 根据主题设置背景色
        if appearance_mode == "Dark":
            list_frame = ctk.CTkFrame(parent, corner_radius=10, fg_color="#212121")
        else:
            list_frame = ctk.CTkFrame(parent, corner_radius=10, fg_color="#F8F9FA")
        list_frame.pack(fill="both", expand=True)
        
        # 列表内容 - 使用透明背景让容器背景透出
        list_content = ctk.CTkFrame(list_frame, fg_color="transparent")
        list_content.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 创建表格
        self.salary_tree = CTKTreeview(
            list_content,
            columns=self.columns,
            corner_radius=5
        )
        self.salary_tree.pack(fill="both", expand=True)
        
        # 保存容器引用以便主题切换时更新
        self.list_frame = list_frame
        
        # 启动主题监听
        self._theme_monitor_active = True
        self.monitor_theme_changes()
        
        # 绑定选择事件
        self.salary_tree.bind('<<TreeviewSelect>>', self.on_select)
    
    def monitor_theme_changes(self):
        """监听主题变化并更新容器背景色"""
        try:
            # 检查窗口是否仍然存在
            if not self.winfo_exists():
                return
                
            current_mode = ctk.get_appearance_mode()
            if not hasattr(self, '_last_theme_mode'):
                self._last_theme_mode = current_mode
            elif self._last_theme_mode != current_mode:
                self._last_theme_mode = current_mode
                self.update_container_colors()
            
            # 每500ms检查一次
            if hasattr(self, '_theme_monitor_active') and self._theme_monitor_active:
                self.after(500, self.monitor_theme_changes)
        except:
            # 如果出现错误，停止监听
            pass
    
    def update_container_colors(self):
        """更新容器背景色"""
        appearance_mode = ctk.get_appearance_mode()
        if hasattr(self, 'list_frame'):
            if appearance_mode == "Dark":
                self.list_frame.configure(fg_color="#212121")
            else:
                self.list_frame.configure(fg_color="#F8F9FA")
    
    def stop_theme_monitoring(self):
        """停止主题监听"""
        self._theme_monitor_active = False
    
    def destroy(self):
        """重写destroy方法，确保停止主题监听"""
        if hasattr(self, '_theme_monitor_active'):
            self.stop_theme_monitoring()
        super().destroy()
    
    def load_salary_months(self):
        """加载工资月份列表"""
        months = self.db.query(SalaryRecord.salary_month)\
            .distinct().order_by(SalaryRecord.salary_month.desc()).all()
        month_list = [m[0] for m in months]
        
        if month_list:
            self.month_combo.configure(values=month_list)
            self.month_var.set(month_list[0])
            self.load_salary_data(month_list[0])
        else:
            self.month_combo.configure(values=[])
            self.month_var.set("")
    
    def on_month_selected(self, selected_month):
        """月份选择改变时的处理"""
        if selected_month:
            self.load_salary_data(selected_month)
    
    def load_salary_data(self, month):
        """加载指定月份的工资数据"""
        # 清空现有数据
        for item in self.salary_tree.get_children():
            self.salary_tree.delete(item)
        
        # 查询数据
        records = self.db.query(SalaryRecord, Employee)\
            .join(Employee, SalaryRecord.employee_id == Employee.employee_id)\
            .filter(SalaryRecord.salary_month == month).all()
        
        # 显示数据
        for idx, (record, emp) in enumerate(records, 1):
            status_map = {
                0: '⏳ 待发送',
                1: '✅ 已发送',
                2: '❌ 失败'
            }
            
            # 格式化金额显示
            def format_money(value):
                try:
                    return f"{float(value):,.2f}" if value else "0.00"
                except:
                    return "0.00"
            
            self.salary_tree.insert('', 'end', values=(
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
    
    def import_salary_data(self):
        """导入工资数据"""
        file_path = CTKFileDialog.open_file(
            self,
            title='选择工资数据文件',
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
        
        # 创建进度对话框
        progress_dialog = CTKProgressDialog(self, "导入工资数据", "正在处理Excel文件...")
        
        def import_worker():
            # 在工作线程中创建新的数据库会话
            from salary_mail.db_instance import set_db
            worker_db = set_db()
            
            try:
                wb = openpyxl.load_workbook(file_path)
                sheet = wb.active
                
                # 获取表头
                headers = [str(cell.value).strip() if cell.value else '' for cell in sheet[1]]
                
                # 验证必要的列
                required_fields = [
                    '员工编号', '姓名', '岗位工资', '薪级工资', '绩效工资',
                    '餐补', '交补', '防暑降温费', '补发', '事假扣款', '病假扣款',
                    '其他扣款', '税前工资', '社会保险', '公积金', '个人所得税',
                    '代缴工会会费', '实发工资', '备注', '月份'
                ]
                
                missing_fields = [f for f in required_fields if f not in headers]
                if missing_fields:
                    self.after(0, lambda: progress_dialog.destroy())
                    self.after(100, lambda: CTKMessageBox.show_error(
                        self, '错误', f"缺少必要的列: {', '.join(missing_fields)}"
                    ))
                    return
                
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
                
                # 数值字段
                numeric_fields = [
                    '岗位工资', '薪级工资', '绩效工资', '餐补', '交补',
                    '防暑降温费', '补发', '事假扣款', '病假扣款', '其他扣款',
                    '税前工资', '社会保险', '公积金', '个人所得税',
                    '代缴工会会费', '实发工资'
                ]
                
                # 开始导入数据
                success_count = 0
                error_count = 0
                total_rows = sheet.max_row - 1  # 减去标题行
                
                for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
                    if not any(row):  # 跳过空行
                        continue
                    
                    # 更新进度
                    progress = ((row_idx - 2) / total_rows) * 100 if total_rows > 0 else 0
                    self.after(0, lambda p=progress, r=row_idx-1, t=total_rows: 
                              progress_dialog.update_progress(p, f"正在处理第 {r}/{t} 行", 
                                                            f"成功: {success_count}, 失败: {error_count}"))
                    
                    if progress_dialog.cancelled:
                        break
                    
                    try:
                        # 获取月份
                        salary_month = str(row[field_indexes['月份']]).strip()
                        if not salary_month:
                            raise ValueError("月份不能为空")
                        
                        salary_month = salary_month.replace('-', '')[:6]
                        
                        employee_id = str(row[field_indexes['员工编号']]).strip()
                        if not employee_id:
                            raise ValueError("员工编号不能为空")
                        
                        # 检查员工是否存在
                        employee = worker_db.query(Employee)\
                            .filter_by(employee_id=employee_id, status=1).first()
                        if not employee:
                            # 如果员工不存在，自动创建
                            employee = Employee(
                                employee_id=employee_id,
                                name=str(row[field_indexes['姓名']]).strip(),
                                email='',
                                status=1
                            )
                            worker_db.add(employee)
                        
                        # 检查是否已存在该月工资记录
                        existing = worker_db.query(SalaryRecord)\
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
                        worker_db.add(record)
                        success_count += 1
                        
                    except Exception as e:
                        error_count += 1
                        continue
                
                worker_db.commit()
                
                # 关闭进度对话框
                self.after(0, lambda: progress_dialog.destroy())
                
                # 刷新界面
                self.after(100, lambda: self.load_salary_months())
                
                # 显示结果
                result_msg = f"导入完成!\n成功: {success_count} 条\n失败: {error_count} 条"
                self.after(200, lambda: CTKMessageBox.show_info(self, "导入结果", result_msg))
                
            except Exception as e:
                error_msg = f"导入过程出错：\n{str(e)}"
                self.after(0, lambda: progress_dialog.destroy())
                self.after(100, lambda: CTKMessageBox.show_error(self, "导入失败", error_msg))
                worker_db.rollback()
            finally:
                # 确保关闭工作线程的数据库连接
                worker_db.close()
        
        # 在后台线程中执行导入
        import_thread = threading.Thread(target=import_worker)
        import_thread.daemon = True
        import_thread.start()
    
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
        selection = self.salary_tree.selection()
        if not selection:
            CTKMessageBox.show_warning(self, '提示', '请先选择要发送的记录！')
            return
        
        self._send_salary_records(selection)
    
    def send_all_salary(self):
        """发送所有工资条"""
        selection = self.salary_tree.get_children()
        if not selection:
            CTKMessageBox.show_warning(self, '提示', '当前月份没有工资记录！')
            return
        
        if CTKMessageBox.ask_yes_no(self, '确认', '确定要发送所有工资条吗？') != "是":
            return
        
        self._send_salary_records(selection)
    
    def _send_salary_records(self, items):
        """发送工资条的具体实现"""
        # 创建进度对话框
        progress_dialog = CTKProgressDialog(self, "发送工资条", "正在准备发送...")
        
        def send_worker():
            # 在工作线程中创建新的数据库会话
            from salary_mail.db_instance import set_db
            worker_db = set_db()
            
            try:
                # 获取邮件配置
                sender = worker_db.query(SalaryEmail).filter(SalaryEmail.field_name=='sender').first()
                password = worker_db.query(SalaryEmail).filter(SalaryEmail.field_name=='password').first()
                sender_name = worker_db.query(SalaryEmail).filter(SalaryEmail.field_name=='sender_name').first()
                smtp_server = worker_db.query(SalaryEmail).filter(SalaryEmail.field_name=='smtp_server').first()
                port = worker_db.query(SalaryEmail).filter(SalaryEmail.field_name=='port').first()
                template = worker_db.query(SalaryEmail).filter(SalaryEmail.field_name=='email_template').first()
                
                if not all([sender, password, sender_name, smtp_server, port, template]):
                    self.after(0, lambda: progress_dialog.destroy())
                    self.after(100, lambda: CTKMessageBox.show_error(self, '错误', '请先完成邮件配置！'))
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
                
                # 开始发送
                for index, item in enumerate(items, 1):
                    if progress_dialog.cancelled:
                        break
                    
                    # 更新进度
                    progress = (index / total_count) * 100
                    values = self.salary_tree.item(item)['values']
                    employee_name = values[2]
                    
                    self.after(0, lambda p=progress, i=index, t=total_count, n=employee_name, s=success_count, e=error_count:
                              progress_dialog.update_progress(p, f"正在发送 ({i}/{t})", 
                                                            f"处理: {n} | 成功: {s}, 失败: {e}"))
                    
                    try:
                        employee_id = values[1]
                        
                        # 获取工资记录和员工信息
                        record = worker_db.query(SalaryRecord)\
                            .filter_by(employee_id=employee_id, salary_month=month).first()
                        employee = worker_db.query(Employee)\
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
                        worker_db.add(record)
                        success_count += 1
                        
                    except Exception as e:
                        print(f"发送失败: {str(e)}")
                        error_count += 1
                    
                    # 每发送一封邮件就提交一次
                    worker_db.commit()
                
                smtp.quit()
                
                # 关闭进度对话框
                self.after(0, lambda: progress_dialog.destroy())
                
                # 刷新显示
                self.after(100, lambda: self.load_salary_data(month))
                
                # 显示结果
                result_msg = f'发送完成！\n成功：{success_count} 条\n失败：{error_count} 条'
                self.after(200, lambda: CTKMessageBox.show_info(self, '发送完成', result_msg))
                
            except Exception as e:
                error_msg = f'发送过程出错：\n{str(e)}'
                self.after(0, lambda: progress_dialog.destroy())
                self.after(100, lambda: CTKMessageBox.show_error(self, '错误', error_msg))
                worker_db.rollback()
            finally:
                # 确保关闭工作线程的数据库连接
                worker_db.close()
        
        # 在后台线程中执行发送
        send_thread = threading.Thread(target=send_worker)
        send_thread.daemon = True
        send_thread.start()
    
    def _make_salary_email(self, template, record, employee, sender, sender_name):
        """生成工资邮件内容"""
        year = record.salary_month[:4]
        month = str(int(record.salary_month[4:]))
        
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
            raise ValueError(f"邮件内容生成失败: {str(e)}")
        
        # 创建邮件对象
        msg = MIMEMultipart()
        msg['From'] = formataddr([sender_name, sender])
        msg['To'] = formataddr([employee.name, employee.email])
        msg['Subject'] = f"{year}年{month}月工资明细"
        
        # 添加HTML内容
        msg.attach(MIMEText(email_content, 'html'))
        
        return msg
    
    def on_select(self, event):
        """处理选择事件"""
        selection = self.salary_tree.selection()
        if selection:
            # 可以在这里添加选中行的相关操作
            pass
