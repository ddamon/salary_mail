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
import json
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from smtplib import SMTP, SMTP_SSL

from salary_mail.db_instance import Employee, SalaryEmail, SalaryRecord, SalaryFieldConfig
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
        
        # 动态列配置
        self.columns = []
        self.salary_fields = []
        
        # 检查工资项配置
        if not self.check_salary_config():
            self.show_config_required_dialog()
            return
        
        self.setup_ui()
        
        # 设置焦点
        self.focus_force()
    
    def check_salary_config(self):
        """检查工资项配置是否存在"""
        try:
            field_count = self.db.query(SalaryFieldConfig).count()
            return field_count > 0
        except Exception as e:
            print(f"检查工资项配置失败: {e}")
            return False
    
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
    
    def show_config_required_dialog(self):
        """显示需要配置工资项的对话框"""
        result = CTKMessageBox.ask_yes_no(
            self, 
            '配置提醒', 
            '检测到工资项配置为空，需要先配置工资项模板。\n\n是否现在前往配置？'
        )
        
        if result == "是":
            # 打开工资项配置窗口
            from salary_mail.ctk_salary_config import CTKSalaryConfigWin
            config_window = CTKSalaryConfigWin(self.parent)
            config_window.wait_window()
            
            # 重新检查配置
            if self.check_salary_config():
                self.setup_ui()
                self.focus_force()
            else:
                self.destroy()
        else:
            self.destroy()
    
    def load_salary_fields(self):
        """加载工资项配置"""
        try:
            # 查询工资项配置
            fields = self.db.query(SalaryFieldConfig).order_by(SalaryFieldConfig.display_order).all()
            self.salary_fields = fields
            
            # 生成列配置
            self.columns = []
            for field in fields:
                if field.is_fixed:
                    # 固定字段
                    if field.field_key == 'serial_number':
                        self.columns.append(('序号', field.display_width))
                    elif field.field_key == 'employee_id':
                        self.columns.append(('员工编号', field.display_width))
                    elif field.field_key == 'name':
                        self.columns.append(('姓名', field.display_width))
                    elif field.field_key == 'month':
                        self.columns.append(('月份', field.display_width))
                else:
                    # 动态工资项
                    self.columns.append((field.field_name, field.display_width))
            
            # 添加发送状态列
            self.columns.append(('发送状态', 70))
            
        except Exception as e:
            CTKMessageBox.show_error(self, '错误', f'加载工资项配置失败：{str(e)}')
    
    def setup_ui(self):
        """设置UI"""
        # 加载工资项配置
        self.load_salary_fields()
        
        # 主框架
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 按钮框架 - 使用工具栏样式
        button_frame = ctk.CTkFrame(main_frame, height=60, corner_radius=10)
        button_frame.pack(fill='x', pady=(0, 15))
        button_frame.pack_propagate(False)
        
        # 按钮内容框架
        button_content = ctk.CTkFrame(button_frame, fg_color="transparent")
        button_content.pack(fill='both', expand=True, padx=15, pady=10)
        
        # 左侧按钮组
        left_buttons = ctk.CTkFrame(button_content, fg_color="transparent")
        left_buttons.pack(side="left", fill="y")
        
        # 导入工资数据按钮
        import_btn = ctk.CTkButton(
            left_buttons,
            text="📊 导入工资数据",
            command=self.import_salary_data,
            width=130,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold")
        )
        import_btn.pack(side='left', padx=(0, 10))
        
        # 发送选中按钮
        send_selected_btn = ctk.CTkButton(
            left_buttons,
            text="📤 发送选中",
            command=self.send_selected_salary,
            width=120,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold"),
            fg_color="#4CAF50",
            hover_color="#45a049"
        )
        send_selected_btn.pack(side='left', padx=(0, 10))
        
        # 全部发送按钮
        send_all_btn = ctk.CTkButton(
            left_buttons,
            text="📨 全部发送",
            command=self.send_all_salary,
            width=120,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold"),
            fg_color="#2196F3",
            hover_color="#1976D2"
        )
        send_all_btn.pack(side='left', padx=(0, 10))
        
        # 刷新按钮
        refresh_btn = ctk.CTkButton(
            left_buttons,
            text="🔄 刷新",
            command=self.refresh_data,
            width=100,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold"),
            fg_color="#FF9800",
            hover_color="#F57C00"
        )
        refresh_btn.pack(side='left', padx=(0, 10))
        
        # 右侧月份选择区域
        month_frame = ctk.CTkFrame(button_content, fg_color="transparent")
        month_frame.pack(side='right', fill='y')
        
        month_label = ctk.CTkLabel(
            month_frame, 
            text="选择月份:",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold")
        )
        month_label.pack(side='left', padx=(0, 5), pady=10)
        
        self.month_var = ctk.StringVar(value=self.get_current_month())
        self.month_combo = ctk.CTkComboBox(
            month_frame, 
            values=self.get_available_months(),
            variable=self.month_var,
            command=self.on_month_changed,
            width=100,
            height=35,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12)
        )
        self.month_combo.pack(side='left', padx=(0, 10), pady=10)
        
        # 工资数据表格
        self.setup_salary_table(main_frame)
        
        # 状态栏
        status_frame = ctk.CTkFrame(main_frame)
        status_frame.pack(fill='x', pady=(10, 0))
        
        self.status_label = ctk.CTkLabel(status_frame, text="就绪")
        self.status_label.pack(side='left', padx=10, pady=5)
        
        # 加载初始数据
        if self.month_var.get():
            self.load_salary_data(self.month_var.get())
    
    def setup_salary_table(self, parent):
        """设置工资数据表格"""
        # 表格框架
        table_frame = ctk.CTkFrame(parent)
        table_frame.pack(fill='both', expand=True)
        
        # 创建表格
        self.salary_tree = CTKTreeview(table_frame, columns=self.columns)
        
        # 设置列标题
        for col, width in self.columns:
            self.salary_tree.tree.heading(col, text=col)
            self.salary_tree.tree.column(col, width=width, minwidth=width)
        
        # 添加滚动条
        scrollbar = ctk.CTkScrollbar(table_frame, command=self.salary_tree.tree.yview)
        self.salary_tree.tree.configure(yscrollcommand=scrollbar.set)
        
        # 布局
        self.salary_tree.pack(side='left', fill='both', expand=True, padx=(10, 0), pady=(10, 0))
        scrollbar.pack(side='right', fill='y', pady=(10, 0))
    
    def get_available_months(self):
        """获取可用的月份列表"""
        try:
            months = self.db.query(SalaryRecord.salary_month).distinct().all()
            month_list = [month[0] for month in months if month[0]]
            month_list.sort(reverse=True)
            
            # 如果没有历史数据，添加当前月份
            current_month = datetime.now().strftime('%Y%m')
            if current_month not in month_list:
                month_list.insert(0, current_month)
            
            return month_list if month_list else [current_month]
        except Exception as e:
            print(f"获取月份列表失败: {e}")
            return [datetime.now().strftime('%Y%m')]
    
    def get_current_month(self):
        """获取当前月份"""
        return datetime.now().strftime('%Y%m')
    
    def on_month_changed(self, month):
        """月份选择改变事件"""
        if month:
            self.load_salary_data(month)
    
    def load_salary_data(self, month):
        """加载指定月份的工资数据"""
        try:
            # 清空表格
            for item in self.salary_tree.tree.get_children():
                self.salary_tree.tree.delete(item)
            
            # 构建查询对象
            query = self.db.query(
                SalaryRecord.id,
                SalaryRecord.employee_id,
                SalaryRecord.salary_month,
                SalaryRecord.salary_data,
                SalaryRecord.remark,
                SalaryRecord.send_status,
                SalaryRecord.send_time,
                Employee.id.label('emp_id'),
                Employee.employee_id.label('emp_employee_id'),
                Employee.name.label('emp_name'),
                Employee.email.label('emp_email'),
                Employee.phone.label('emp_phone')
            ).join(
                Employee, SalaryRecord.employee_id == Employee.employee_id
            ).filter(SalaryRecord.salary_month == month)
            
            # 打印SQL语句用于调试
            print("=== SQL查询语句 ===")
            print(str(query))
            print("==================")
            
            # 执行查询
            records = query.all()
            
            # 显示数据
            for idx, record in enumerate(records, 1):
                status_map = {
                    0: '⏳ 待发送',
                    1: '✅ 已发送',
                    2: '❌ 失败'
                }
                
                # 获取工资数据
                salary_data = self.parse_salary_data(record.salary_data)
                
                # 构建行数据
                row_data = []
                for field in self.salary_fields:
                    if field.is_fixed:
                        if field.field_key == 'serial_number':
                            row_data.append(str(idx))
                        elif field.field_key == 'employee_id':
                            row_data.append(record.emp_employee_id)
                        elif field.field_key == 'name':
                            row_data.append(record.emp_name)
                        elif field.field_key == 'month':
                            row_data.append(record.salary_month)
                    else:
                        # 动态工资项
                        value = salary_data.get(field.field_key, '0.00')
                        row_data.append(self.format_money(value))
                
                # 添加发送状态
                row_data.append(status_map.get(record.send_status, '未知'))
                
                # 插入行
                self.salary_tree.tree.insert('', 'end', values=row_data)
            
            self.status_label.configure(text=f"已加载 {len(records)} 条工资记录")
            
        except Exception as e:
            print(f"=== 错误详情 ===")
            print(f"错误类型: {type(e).__name__}")
            print(f"错误信息: {str(e)}")
            print(f"==================")
            # 不弹窗，只在控制台显示错误
            self.status_label.configure(text=f"加载失败: {str(e)}")
    
    def parse_salary_data(self, salary_data_json):
        """解析工资数据JSON字符串"""
        if salary_data_json:
            try:
                return json.loads(salary_data_json)
            except json.JSONDecodeError:
                return {}
        return {}
    
    def format_money(self, value):
        """格式化金额显示"""
        try:
            return f"{float(value):,.2f}" if value else "0.00"
        except:
            return "0.00"
    
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
                required_fields = ['员工编号', '姓名', '月份']
                missing_fields = [f for f in required_fields if f not in headers]
                if missing_fields:
                    self.after(0, lambda: progress_dialog.destroy())
                    self.after(100, lambda: CTKMessageBox.show_error(
                        self, '错误', f"缺少必要的列: {', '.join(missing_fields)}"
                    ))
                    return
                
                # 获取列索引
                field_indexes = {field: headers.index(field) for field in headers}
                
                # 获取工资项配置
                salary_fields = worker_db.query(SalaryFieldConfig).filter(
                    SalaryFieldConfig.is_fixed == 0
                ).order_by(SalaryFieldConfig.display_order).all()
                
                success_count = 0
                error_count = 0
                
                # 处理数据行
                for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
                    if not row[0]:  # 跳过空行
                        continue
                    
                    try:
                        # 构建工资数据字典
                        salary_data = {}
                        for field in salary_fields:
                            if field.field_name in headers:
                                col_idx = field_indexes[field.field_name]
                                if col_idx < len(row):
                                    salary_data[field.field_key] = str(row[col_idx]) if row[col_idx] is not None else '0.00'
                        
                        # 创建工资记录
                        salary_record = SalaryRecord(
                            employee_id=str(row[field_indexes['员工编号']]),
                            salary_month=str(row[field_indexes['月份']]),
                            remark=row[field_indexes.get('备注', -1)] if '备注' in headers and field_indexes['备注'] < len(row) else '',
                            send_status=0
                        )
                        salary_record.set_salary_data(salary_data)
                        
                        worker_db.add(salary_record)
                        success_count += 1
                        
                    except Exception as e:
                        print(f"处理第{row_idx}行数据失败: {e}")
                        error_count += 1
                        continue
                
                worker_db.commit()
                
                # 更新UI
                self.after(0, lambda: progress_dialog.destroy())
                self.after(100, lambda: self.show_import_result(success_count, error_count))
                
            except Exception as e:
                self.after(0, lambda: progress_dialog.destroy())
                self.after(100, lambda: CTKMessageBox.show_error(
                    self, '错误', f"导入失败：{str(e)}"
                ))
            finally:
                worker_db.close()
        
        # 启动工作线程
        thread = threading.Thread(target=import_worker)
        thread.daemon = True
        thread.start()
    
    def show_import_result(self, success_count, error_count):
        """显示导入结果"""
        CTKMessageBox.show_info(
            self, 
            '导入完成', 
            f'导入完成！\n成功：{success_count} 条\n失败：{error_count} 条'
        )
        
        # 刷新数据
        self.refresh_data()
    
    def send_selected_salary(self):
        """发送选中的工资条"""
        selection = self.salary_tree.tree.selection()
        if not selection:
            CTKMessageBox.show_warning(self, '提示', '请先选择要发送的工资条')
            return
        
        # 确认发送
        result = CTKMessageBox.ask_yes_no(
            self, 
            '确认发送', 
            f'确定要发送选中的 {len(selection)} 条工资条吗？'
        )
        
        if result == "是":
            self._send_salary_records(selection)
    
    def send_all_salary(self):
        """发送所有工资条"""
        all_items = self.salary_tree.tree.get_children()
        if not all_items:
            CTKMessageBox.show_warning(self, '提示', '没有可发送的工资条')
            return
        
        # 确认发送
        result = CTKMessageBox.ask_yes_no(
            self, 
            '确认发送', 
            f'确定要发送所有 {len(all_items)} 条工资条吗？'
        )
        
        if result == "是":
            self._send_salary_records(all_items)
    
    def _send_salary_records(self, items):
        """发送工资记录"""
        # 这里实现发送逻辑，暂时显示提示
        CTKMessageBox.show_info(self, '提示', '发送功能待实现')
    
    def refresh_data(self):
        """刷新数据"""
        # 重新获取月份列表
        months = self.get_available_months()
        self.month_combo.configure(values=months)
        
        # 重新加载当前月份数据
        if self.month_var.get():
            self.load_salary_data(self.month_var.get())
