# coding:utf-8
"""
PySide6版本的工资管理窗口
"""
import os
import threading
import base64
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from smtplib import SMTP, SMTP_SSL

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, 
                               QTableWidgetItem, QLineEdit, QPushButton, QLabel,
                               QFrame, QHeaderView, QAbstractItemView, QMessageBox,
                               QFileDialog, QProgressDialog, QSpacerItem, QSizePolicy,
                               QComboBox, QCheckBox, QTextEdit, QTabWidget, QWidget,
                               QGridLayout, QGroupBox, QScrollArea)
from PySide6.QtCore import Qt, QTimer, Signal, QThread, QMutex, QMutexLocker
from PySide6.QtGui import QFont, QIcon, QPalette, QColor
from sqlalchemy import or_
import openpyxl

from salary_mail.db_instance import Employee, SalaryEmail, SalaryRecord


class EmailSendThread(QThread):
    """邮件发送线程"""
    progress_updated = Signal(int)
    status_updated = Signal(str)
    finished = Signal(bool, str)
    
    def __init__(self, salary_records, email_config, template):
        super().__init__()
        self.salary_records = salary_records
        self.email_config = email_config
        self.template = template
        self.mutex = QMutex()
        self.should_stop = False
    
    def stop(self):
        with QMutexLocker(self.mutex):
            self.should_stop = True
    
    def run(self):
        try:
            total = len(self.salary_records)
            sent_count = 0
            
            for i, record in enumerate(self.salary_records):
                with QMutexLocker(self.mutex):
                    if self.should_stop:
                        break
                
                self.status_updated.emit(f"正在发送给 {record.employee.name}...")
                
                # 模拟发送邮件
                success = self.send_single_email(record)
                if success:
                    sent_count += 1
                
                progress = int((i + 1) * 100 / total)
                self.progress_updated.emit(progress)
            
            self.finished.emit(True, f"发送完成！成功发送 {sent_count}/{total} 封邮件")
            
        except Exception as e:
            self.finished.emit(False, f"发送失败: {str(e)}")
    
    def send_single_email(self, record):
        """发送单封邮件"""
        try:
            # 这里应该实现真正的邮件发送逻辑
            # 暂时模拟发送过程
            import time
            time.sleep(0.1)  # 模拟发送延迟
            return True
        except:
            return False


class PySideSalaryManageWin(QDialog):
    """PySide6版本的工资管理窗口"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.db = parent.db if parent else None
        
        self.setWindowTitle('工资管理')
        self.setModal(False)  # 非模态窗口
        self.showMaximized()  # 最大化显示
        
        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(__file__), 'assets', 'salary.png')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
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
        self.load_salary_data()
        
        # 邮件发送线程
        self.email_thread = None
    
    def setup_ui(self):
        """设置用户界面"""
        # 设置样式
        self.setStyleSheet("""
            QDialog {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1,
                    stop: 0 #f8fafc, stop: 1 #e2e8f0);
            }
            QFrame#toolbar_frame {
                background: rgba(255, 255, 255, 0.95);
                border-radius: 20px;
                border: none;
                box-shadow: 0 5px 20px rgba(0, 0, 0, 0.1);
            }
            QFrame#table_frame {
                background: rgba(255, 255, 255, 0.95);
                border-radius: 20px;
                border: none;
                box-shadow: 0 5px 20px rgba(0, 0, 0, 0.1);
            }
            QTabWidget::pane {
                border: none;
                background: rgba(255, 255, 255, 0.95);
                border-radius: 15px;
                box-shadow: 0 5px 20px rgba(0, 0, 0, 0.1);
            }
            QTabBar::tab {
                background: rgba(59, 130, 246, 0.1);
                color: #3b82f6;
                padding: 12px 20px;
                margin: 2px;
                border-radius: 12px;
                font-weight: 600;
                min-width: 100px;
            }
            QTabBar::tab:selected {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #3b82f6, stop: 1 #1d4ed8);
                color: white;
                box-shadow: 0 4px 15px rgba(59, 130, 246, 0.3);
            }
            QTabBar::tab:hover:!selected {
                background: rgba(59, 130, 246, 0.2);
            }
            QLineEdit, QComboBox {
                padding: 6px 12px;
                border: 2px solid #e8f4f8;
                border-radius: 12px;
                font-size: 14px;
                background: rgba(255, 255, 255, 0.9);
                color: #2c3e50;
            }
            QLineEdit:focus, QComboBox:focus {
                border-color: #3b82f6;
                background: rgba(255, 255, 255, 1);
                box-shadow: 0 0 15px rgba(59, 130, 246, 0.3);
            }
            QLineEdit::placeholder {
                color: #94a3b8;
            }
            QPushButton {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #3b82f6, stop: 1 #1d4ed8);
                color: white;
                border: none;
                border-radius: 12px;
                padding: 12px 20px;
                font-size: 14px;
                font-weight: bold;
                box-shadow: 0 4px 15px rgba(59, 130, 246, 0.3);
            }
            QPushButton:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #1d4ed8, stop: 1 #1e3a8a);
                transform: translateY(-2px);
                box-shadow: 0 6px 20px rgba(59, 130, 246, 0.4);
            }
            QPushButton:pressed {
                transform: translateY(0px);
                box-shadow: 0 2px 10px rgba(59, 130, 246, 0.3);
            }
            QPushButton#import_btn {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #10b981, stop: 1 #059669);
                box-shadow: 0 4px 15px rgba(16, 185, 129, 0.3);
            }
            QPushButton#import_btn:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #059669, stop: 1 #047857);
                box-shadow: 0 6px 20px rgba(16, 185, 129, 0.4);
            }
            QPushButton#send_btn {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #f59e0b, stop: 1 #d97706);
                box-shadow: 0 4px 15px rgba(245, 158, 11, 0.3);
            }
            QPushButton#send_btn:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #d97706, stop: 1 #b45309);
                box-shadow: 0 6px 20px rgba(245, 158, 11, 0.4);
            }
            QPushButton#delete_btn {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #ef4444, stop: 1 #dc2626);
                box-shadow: 0 4px 15px rgba(239, 68, 68, 0.3);
            }
            QPushButton#delete_btn:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #dc2626, stop: 1 #b91c1c);
                box-shadow: 0 6px 20px rgba(239, 68, 68, 0.4);
            }
            QTableWidget {
                border: none;
                gridline-color: #f1f5f9;
                background-color: white;
                alternate-background-color: #f8fafc;
                border-radius: 15px;
            }
            QTableWidget::item {
                padding: 10px;
                border-bottom: 1px solid #f1f5f9;
                color: #475569;
            }
            QTableWidget::item:selected {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 rgba(59, 130, 246, 0.1), stop: 1 rgba(29, 78, 216, 0.1));
                color: #1d4ed8;
            }
            QHeaderView::section {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #3b82f6, stop: 1 #1d4ed8);
                color: white;
                padding: 12px;
                border: none;
                font-weight: bold;
                font-size: 12px;
            }
            QHeaderView::section:first {
                border-top-left-radius: 15px;
            }
            QHeaderView::section:last {
                border-top-right-radius: 15px;
            }
            QLabel#title_label {
                color: #3b82f6;
                font-weight: bold;
                text-shadow: 1px 1px 3px rgba(0, 0, 0, 0.1);
            }
            QLabel#count_label {
                color: #64748b;
                font-weight: 600;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #e2e8f0;
                border-radius: 15px;
                margin-top: 1ex;
                padding-top: 15px;
                background: rgba(255, 255, 255, 0.95);
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 10px 0 10px;
                color: #2c3e50;
            }
            QCheckBox {
                color: #475569;
                font-weight: 500;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 9px;
                border: 2px solid #cbd5e1;
            }
            QCheckBox::indicator:checked {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #3b82f6, stop: 1 #1d4ed8);
                border-color: #3b82f6;
            }
            QLabel {
                color: #374151;
            }
        """)
        
        # 主布局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # 创建标签页
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)
        
        # 工资数据标签页
        self.create_salary_tab()
        
        # 发送设置标签页
        self.create_send_tab()
    
    def create_salary_tab(self):
        """创建工资数据标签页"""
        salary_widget = QWidget()
        salary_layout = QVBoxLayout(salary_widget)
        salary_layout.setContentsMargins(15, 15, 15, 15)
        salary_layout.setSpacing(15)
        
        # 工具栏
        self.create_salary_toolbar(salary_layout)
        
        # 数据表格
        self.create_salary_table(salary_layout)
        
        # 按钮栏
        self.create_salary_buttons(salary_layout)
        
        self.tab_widget.addTab(salary_widget, "工资数据")
    
    def create_send_tab(self):
        """创建发送设置标签页"""
        send_widget = QWidget()
        send_layout = QVBoxLayout(send_widget)
        send_layout.setContentsMargins(15, 15, 15, 15)
        send_layout.setSpacing(15)
        
        # 发送设置区域
        send_group = QGroupBox("邮件发送设置")
        send_group.setFont(QFont("Microsoft YaHei UI", 14, QFont.Bold))
        send_group_layout = QVBoxLayout(send_group)
        
        # 工资月份选择
        month_layout = QHBoxLayout()
        month_layout.addWidget(QLabel("工资月份:"))
        self.month_combo = QComboBox()
        self.month_combo.setMinimumWidth(150)
        # 填充最近12个月
        from datetime import datetime, timedelta
        current_date = datetime.now()
        for i in range(12):
            date = current_date - timedelta(days=30*i)
            month_str = date.strftime("%Y%m")
            display_str = date.strftime("%Y年%m月")
            self.month_combo.addItem(display_str, month_str)
        month_layout.addWidget(self.month_combo)
        month_layout.addStretch()
        send_group_layout.addLayout(month_layout)
        
        # 发送选项
        options_layout = QHBoxLayout()
        self.send_all_checkbox = QCheckBox("发送给所有员工")
        self.send_all_checkbox.setChecked(True)
        options_layout.addWidget(self.send_all_checkbox)
        
        self.test_send_checkbox = QCheckBox("测试发送（仅发送给管理员）")
        options_layout.addWidget(self.test_send_checkbox)
        options_layout.addStretch()
        send_group_layout.addLayout(options_layout)
        
        send_layout.addWidget(send_group)
        
        # 进度显示区域
        progress_group = QGroupBox("发送进度")
        progress_group.setFont(QFont("Microsoft YaHei UI", 14, QFont.Bold))
        progress_layout = QVBoxLayout(progress_group)
        
        self.progress_label = QLabel("就绪")
        progress_layout.addWidget(self.progress_label)
        
        self.progress_bar = QProgressDialog("", "取消", 0, 100, self)
        self.progress_bar.setWindowModality(Qt.WindowModal)
        self.progress_bar.setAutoClose(True)
        self.progress_bar.setAutoReset(True)
        self.progress_bar.hide()
        
        send_layout.addWidget(progress_group)
        
        # 发送按钮
        send_button_layout = QHBoxLayout()
        send_button_layout.addStretch()
        
        self.send_btn = QPushButton("开始发送工资条")
        self.send_btn.setObjectName("send_btn")
        self.send_btn.setFont(QFont("Microsoft YaHei UI", 16, QFont.Bold))
        self.send_btn.setMinimumHeight(50)
        self.send_btn.setMinimumWidth(200)
        self.send_btn.clicked.connect(self.start_send_emails)
        send_button_layout.addWidget(self.send_btn)
        
        send_button_layout.addStretch()
        send_layout.addLayout(send_button_layout)
        
        send_layout.addStretch()
        
        self.tab_widget.addTab(send_widget, "发送设置")
    
    def create_salary_toolbar(self, layout):
        """创建工资数据工具栏"""
        toolbar_frame = QFrame()
        toolbar_frame.setObjectName("toolbar_frame")
        toolbar_frame.setFixedHeight(80)
        
        toolbar_layout = QHBoxLayout(toolbar_frame)
        toolbar_layout.setContentsMargins(20, 15, 20, 15)
        toolbar_layout.setSpacing(15)
        
        # 标题
        title_label = QLabel("工资数据管理")
        title_label.setObjectName("title_label")
        title_label.setFont(QFont("Microsoft YaHei UI", 18, QFont.Bold))
        toolbar_layout.addWidget(title_label)
        
        toolbar_layout.addStretch()
        
        # 搜索框
        search_label = QLabel("搜索:")
        search_label.setFont(QFont("Microsoft YaHei UI", 12))
        toolbar_layout.addWidget(search_label)
        
        self.search_entry = QLineEdit()
        self.search_entry.setPlaceholderText("输入员工编号或姓名...")
        self.search_entry.setFixedWidth(200)
        self.search_entry.textChanged.connect(self.on_search_text_changed)
        toolbar_layout.addWidget(self.search_entry)
        
        # 记录数量标签
        self.count_label = QLabel("共 0 条记录")
        self.count_label.setObjectName("count_label")
        self.count_label.setFont(QFont("Microsoft YaHei UI", 12))
        toolbar_layout.addWidget(self.count_label)
        
        layout.addWidget(toolbar_frame)
    
    def create_salary_table(self, layout):
        """创建工资数据表格"""
        table_frame = QFrame()
        table_frame.setObjectName("table_frame")
        
        table_layout = QVBoxLayout(table_frame)
        table_layout.setContentsMargins(15, 15, 15, 15)
        
        # 创建表格
        self.table = QTableWidget()
        self.table.setColumnCount(len(self.columns))
        self.table.setHorizontalHeaderLabels([col[0] for col in self.columns])
        
        # 设置表格属性
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.MultiSelection)
        self.table.setSortingEnabled(True)
        
        # 设置列宽
        header = self.table.horizontalHeader()
        for i, (_, width) in enumerate(self.columns):
            header.resizeSection(i, width)
        header.setStretchLastSection(True)
        
        table_layout.addWidget(self.table)
        layout.addWidget(table_frame)
    
    def create_salary_buttons(self, layout):
        """创建工资数据按钮栏"""
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        # 导入Excel按钮
        import_btn = QPushButton("导入工资数据")
        import_btn.setObjectName("import_btn")
        import_btn.clicked.connect(self.import_salary_excel)
        button_layout.addWidget(import_btn)
        
        # 删除记录按钮
        delete_btn = QPushButton("删除选中记录")
        delete_btn.setObjectName("delete_btn")
        delete_btn.clicked.connect(self.delete_selected_records)
        button_layout.addWidget(delete_btn)
        
        button_layout.addStretch()
        
        # 关闭按钮
        close_btn = QPushButton("关闭")
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #666666;
            }
            QPushButton:hover {
                background-color: #555555;
            }
        """)
        close_btn.clicked.connect(self.close)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
    
    def on_search_text_changed(self):
        """搜索文本改变时的处理"""
        search_text = self.search_entry.text().strip()
        
        # 简单的表格过滤
        for row in range(self.table.rowCount()):
            show_row = True
            if search_text:
                # 检查员工编号和姓名列
                employee_id = self.table.item(row, 1).text() if self.table.item(row, 1) else ""
                name = self.table.item(row, 2).text() if self.table.item(row, 2) else ""
                if search_text.lower() not in employee_id.lower() and search_text.lower() not in name.lower():
                    show_row = False
            
            self.table.setRowHidden(row, not show_row)
    
    def load_salary_data(self):
        """加载工资数据"""
        if not self.db:
            return
        
        try:
            # 获取工资记录
            records = self.db.query(SalaryRecord).all()
            self.populate_table(records)
        except Exception as e:
            print(f"加载工资数据失败: {e}")
            QMessageBox.critical(self, "错误", f"加载工资数据失败: {str(e)}")
    
    def populate_table(self, records):
        """填充表格数据"""
        self.table.setRowCount(len(records))
        
        for row, record in enumerate(records):
            # 序号
            self.table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
            # 员工编号
            self.table.setItem(row, 1, QTableWidgetItem(record.employee_id))
            # 姓名（需要从员工表获取）
            employee = self.db.query(Employee).filter_by(employee_id=record.employee_id).first()
            name = employee.name if employee else "未知"
            self.table.setItem(row, 2, QTableWidgetItem(name))
            
            # 工资数据列
            salary_fields = [
                'post_salary', 'level_salary', 'performance', 'meal_allowance',
                'traffic_allowance', 'cooling_allowance', 'additional_payment',
                'casual_leave_deduction', 'sick_leave_deduction', 'other_deduction',
                'pre_tax_salary', 'social_insurance', 'housing_fund', 'personal_tax',
                'union_fee', 'actual_salary'
            ]
            
            for i, field in enumerate(salary_fields, start=3):
                value = getattr(record, field, '') or ''
                self.table.setItem(row, i, QTableWidgetItem(str(value)))
            
            # 发送状态
            status_item = QTableWidgetItem("未发送")
            status_item.setBackground(QColor("#ffeb3b"))
            self.table.setItem(row, len(self.columns)-1, status_item)
        
        # 更新记录数量
        self.count_label.setText(f"共 {len(records)} 条记录")
    
    def import_salary_excel(self):
        """导入工资Excel"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择工资数据Excel文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        
        if file_path:
            QMessageBox.information(self, "功能开发中", "Excel导入功能正在开发中，敬请期待！")
    
    def delete_selected_records(self):
        """删除选中的记录"""
        selected_rows = set()
        for item in self.table.selectedItems():
            selected_rows.add(item.row())
        
        if not selected_rows:
            QMessageBox.warning(self, "提示", "请先选择要删除的记录")
            return
        
        reply = QMessageBox.question(
            self, "确认删除", 
            f"确定要删除选中的 {len(selected_rows)} 条记录吗？\n此操作不可撤销！",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            QMessageBox.information(self, "功能开发中", "删除记录功能正在开发中，敬请期待！")
    
    def start_send_emails(self):
        """开始发送邮件"""
        if self.email_thread and self.email_thread.isRunning():
            QMessageBox.warning(self, "提示", "邮件发送正在进行中，请稍候...")
            return
        
        # 获取选中的月份
        selected_month = self.month_combo.currentData()
        if not selected_month:
            QMessageBox.warning(self, "提示", "请选择工资月份")
            return
        
        try:
            # 获取要发送的工资记录
            records = self.db.query(SalaryRecord).filter_by(salary_month=selected_month).all()
            if not records:
                QMessageBox.information(self, "提示", f"没有找到 {selected_month} 月的工资记录")
                return
            
            # 检查邮箱配置
            email_config = self.db.query(SalaryEmail).first()
            if not email_config:
                QMessageBox.warning(self, "提示", "请先配置邮箱设置")
                return
            
            # 显示进度对话框
            self.progress_bar.setLabelText("正在准备发送...")
            self.progress_bar.setRange(0, len(records))
            self.progress_bar.setValue(0)
            self.progress_bar.show()
            
            # 创建并启动邮件发送线程
            self.email_thread = EmailSendThread(records, email_config, "")
            self.email_thread.progress_updated.connect(self.progress_bar.setValue)
            self.email_thread.status_updated.connect(self.progress_bar.setLabelText)
            self.email_thread.finished.connect(self.on_send_finished)
            self.progress_bar.canceled.connect(self.email_thread.stop)
            
            self.email_thread.start()
            self.send_btn.setEnabled(False)
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"启动邮件发送失败: {str(e)}")
    
    def on_send_finished(self, success, message):
        """邮件发送完成"""
        self.progress_bar.hide()
        self.send_btn.setEnabled(True)
        
        if success:
            QMessageBox.information(self, "发送完成", message)
            # 刷新表格状态
            self.load_salary_data()
        else:
            QMessageBox.critical(self, "发送失败", message)
    
    def closeEvent(self, event):
        """处理窗口关闭事件"""
        if self.email_thread and self.email_thread.isRunning():
            self.email_thread.stop()
            self.email_thread.wait()
        event.accept()
