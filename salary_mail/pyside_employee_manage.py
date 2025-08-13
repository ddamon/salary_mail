# coding:utf-8
"""
PySide6版本的员工管理窗口
"""
import os
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, 
                               QTableWidgetItem, QLineEdit, QPushButton, QLabel,
                               QFrame, QHeaderView, QAbstractItemView, QMessageBox,
                               QFileDialog, QProgressDialog, QSpacerItem, QSizePolicy)
from PySide6.QtCore import Qt, QTimer, Signal, QThread
from PySide6.QtGui import QFont, QIcon
from sqlalchemy import or_
import openpyxl
from datetime import datetime

from salary_mail.db_instance import Employee


class PySideEmployeeManageWin(QDialog):
    """PySide6版本的员工管理窗口"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.db = parent.db if parent else None
        
        self.setWindowTitle('员工管理')
        self.setModal(False)  # 改为非模态窗口
        self.showMaximized()  # 最大化显示
        
        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(__file__), 'assets', 'employee.png')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        # 居中显示
        self.center_window()
        
        # 设置UI
        self.setup_ui()
        
        # 加载数据
        self.load_employees()
        
        # 设置搜索延迟
        self.search_timer = QTimer()
        self.search_timer.timeout.connect(self.perform_search)
        self.search_timer.setSingleShot(True)
    
    def center_window(self):
        """窗口居中显示"""
        if self.parent():
            parent_geo = self.parent().geometry()
            x = parent_geo.x() + (parent_geo.width() - self.width()) // 2
            y = parent_geo.y() + (parent_geo.height() - self.height()) // 2
            self.move(x, y)
        else:
            # 如果没有父窗口，在屏幕中央显示
            screen = self.screen()
            screen_rect = screen.geometry()
            x = (screen_rect.width() - self.width()) // 2
            y = (screen_rect.height() - self.height()) // 2
            self.move(x, y)
    
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
            QLineEdit {
                padding: 6px 12px;
                border: 2px solid #e8f4f8;
                border-radius: 12px;
                font-size: 14px;
                background: rgba(255, 255, 255, 0.9);
                color: #2c3e50;
            }
            QLineEdit:focus {
                border-color: #10b981;
                background: rgba(255, 255, 255, 1);
                box-shadow: 0 0 15px rgba(16, 185, 129, 0.3);
            }
            QLineEdit::placeholder {
                color: #94a3b8;
            }
            QPushButton {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #10b981, stop: 1 #059669);
                color: white;
                border: none;
                border-radius: 12px;
                padding: 12px 20px;
                font-size: 14px;
                font-weight: bold;
                box-shadow: 0 4px 15px rgba(16, 185, 129, 0.3);
            }
            QPushButton:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #059669, stop: 1 #047857);
                transform: translateY(-2px);
                box-shadow: 0 6px 20px rgba(16, 185, 129, 0.4);
            }
            QPushButton:pressed {
                transform: translateY(0px);
                box-shadow: 0 2px 10px rgba(16, 185, 129, 0.3);
            }
            QPushButton#import_btn {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #3b82f6, stop: 1 #1d4ed8);
                box-shadow: 0 4px 15px rgba(59, 130, 246, 0.3);
            }
            QPushButton#import_btn:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #1d4ed8, stop: 1 #1e3a8a);
                box-shadow: 0 6px 20px rgba(59, 130, 246, 0.4);
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
                padding: 12px;
                border-bottom: 1px solid #f1f5f9;
                color: #475569;
            }
            QTableWidget::item:selected {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 rgba(16, 185, 129, 0.1), stop: 1 rgba(5, 150, 105, 0.1));
                color: #047857;
            }
            QHeaderView::section {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #10b981, stop: 1 #059669);
                color: white;
                padding: 15px;
                border: none;
                font-weight: bold;
                font-size: 13px;
            }
            QHeaderView::section:first {
                border-top-left-radius: 15px;
            }
            QHeaderView::section:last {
                border-top-right-radius: 15px;
            }
            QLabel#count_label {
                color: #64748b;
                font-weight: 600;
            }
            QLabel {
                color: #374151;
            }
        """)
        
        # 主布局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # 创建各个部分
        self.create_toolbar(main_layout)
        self.create_table(main_layout)
        self.create_button_bar(main_layout)
    
    def create_toolbar(self, layout):
        """创建工具栏"""
        toolbar_frame = QFrame()
        toolbar_frame.setObjectName("toolbar_frame")
        toolbar_frame.setFixedHeight(80)
        
        toolbar_layout = QHBoxLayout(toolbar_frame)
        toolbar_layout.setContentsMargins(20, 15, 20, 15)
        toolbar_layout.setSpacing(15)
        
        # 标题
        title_label = QLabel("员工管理")
        title_label.setFont(QFont("Microsoft YaHei UI", 18, QFont.Bold))
        title_label.setStyleSheet("""
            color: #10b981; 
            font-weight: bold;
            text-shadow: 1px 1px 3px rgba(0, 0, 0, 0.1);
        """)
        toolbar_layout.addWidget(title_label)
        
        toolbar_layout.addStretch()
        
        # 搜索框
        search_label = QLabel("搜索:")
        search_label.setFont(QFont("Microsoft YaHei UI", 12))
        toolbar_layout.addWidget(search_label)
        
        self.search_entry = QLineEdit()
        self.search_entry.setPlaceholderText("输入员工编号、姓名或邮箱...")
        self.search_entry.setFixedWidth(250)
        self.search_entry.textChanged.connect(self.on_search_text_changed)
        toolbar_layout.addWidget(self.search_entry)
        
        # 员工数量标签
        self.count_label = QLabel("共 0 名员工")
        self.count_label.setObjectName("count_label")
        self.count_label.setFont(QFont("Microsoft YaHei UI", 12))
        toolbar_layout.addWidget(self.count_label)
        
        layout.addWidget(toolbar_frame)
    
    def create_table(self, layout):
        """创建表格"""
        table_frame = QFrame()
        table_frame.setObjectName("table_frame")
        
        table_layout = QVBoxLayout(table_frame)
        table_layout.setContentsMargins(15, 15, 15, 15)
        
        # 创建表格
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(['员工编号', '姓名', '邮箱', '手机号', '状态'])
        
        # 设置表格属性
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setSortingEnabled(True)
        
        # 设置列宽
        header = self.table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # 员工编号
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # 姓名
        header.setSectionResizeMode(2, QHeaderView.Stretch)           # 邮箱
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # 手机号
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # 状态
        
        table_layout.addWidget(self.table)
        layout.addWidget(table_frame)
    
    def create_button_bar(self, layout):
        """创建按钮栏"""
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        # 添加员工按钮
        add_btn = QPushButton("添加员工")
        add_btn.clicked.connect(self.add_employee)
        button_layout.addWidget(add_btn)
        
        # 编辑员工按钮
        edit_btn = QPushButton("编辑员工")
        edit_btn.clicked.connect(self.edit_employee)
        button_layout.addWidget(edit_btn)
        
        # 删除员工按钮
        delete_btn = QPushButton("删除员工")
        delete_btn.setObjectName("delete_btn")
        delete_btn.clicked.connect(self.delete_employee)
        button_layout.addWidget(delete_btn)
        
        button_layout.addStretch()
        
        # 导入Excel按钮
        import_btn = QPushButton("导入Excel")
        import_btn.setObjectName("import_btn")
        import_btn.clicked.connect(self.import_excel)
        button_layout.addWidget(import_btn)
        
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
        # 延迟搜索，避免频繁查询
        self.search_timer.stop()
        self.search_timer.start(300)  # 300ms延迟
    
    def perform_search(self):
        """执行搜索"""
        search_text = self.search_entry.text().strip()
        
        if not self.db:
            return
        
        try:
            if search_text:
                # 模糊搜索
                employees = self.db.query(Employee).filter(
                    or_(
                        Employee.employee_id.like(f'%{search_text}%'),
                        Employee.name.like(f'%{search_text}%'),
                        Employee.email.like(f'%{search_text}%'),
                        Employee.phone.like(f'%{search_text}%')
                    )
                ).all()
            else:
                # 显示所有员工
                employees = self.db.query(Employee).all()
            
            self.populate_table(employees)
            
        except Exception as e:
            print(f"搜索失败: {e}")
            QMessageBox.critical(self, "错误", f"搜索失败: {str(e)}")
    
    def load_employees(self):
        """加载员工数据"""
        if not self.db:
            return
        
        try:
            employees = self.db.query(Employee).all()
            self.populate_table(employees)
        except Exception as e:
            print(f"加载员工数据失败: {e}")
            QMessageBox.critical(self, "错误", f"加载员工数据失败: {str(e)}")
    
    def populate_table(self, employees):
        """填充表格数据"""
        self.table.setRowCount(len(employees))
        
        for row, employee in enumerate(employees):
            # 员工编号
            self.table.setItem(row, 0, QTableWidgetItem(str(employee.employee_id)))
            # 姓名
            self.table.setItem(row, 1, QTableWidgetItem(employee.name))
            # 邮箱
            self.table.setItem(row, 2, QTableWidgetItem(employee.email or ''))
            # 手机号
            self.table.setItem(row, 3, QTableWidgetItem(employee.phone or ''))
            # 状态
            status_text = "在职" if employee.status == 1 else "离职"
            status_item = QTableWidgetItem(status_text)
            if employee.status == 1:
                status_item.setBackground(Qt.green)
                status_item.setForeground(Qt.white)
            else:
                status_item.setBackground(Qt.red)
                status_item.setForeground(Qt.white)
            self.table.setItem(row, 4, status_item)
        
        # 更新员工数量
        self.count_label.setText(f"共 {len(employees)} 名员工")
    
    def add_employee(self):
        """添加员工"""
        QMessageBox.information(self, "功能开发中", "添加员工功能正在开发中，敬请期待！")
    
    def edit_employee(self):
        """编辑员工"""
        current_row = self.table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "提示", "请先选择要编辑的员工")
            return
        
        QMessageBox.information(self, "功能开发中", "编辑员工功能正在开发中，敬请期待！")
    
    def delete_employee(self):
        """删除员工"""
        current_row = self.table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "提示", "请先选择要删除的员工")
            return
        
        employee_name = self.table.item(current_row, 1).text()
        reply = QMessageBox.question(
            self, "确认删除", 
            f"确定要删除员工 {employee_name} 吗？\n此操作不可撤销！",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            QMessageBox.information(self, "功能开发中", "删除员工功能正在开发中，敬请期待！")
    
    def import_excel(self):
        """导入Excel"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        
        if file_path:
            QMessageBox.information(self, "功能开发中", "Excel导入功能正在开发中，敬请期待！")
    
    def closeEvent(self, event):
        """处理窗口关闭事件"""
        self.search_timer.stop()
        event.accept()
