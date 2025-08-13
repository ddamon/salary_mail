# coding:utf-8
"""
PySide6版本的系统设置窗口
"""
import os
import re
import base64
import hashlib
import tempfile
import webbrowser
from datetime import datetime

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLineEdit, 
                               QPushButton, QLabel, QFrame, QMessageBox,
                               QTextEdit, QComboBox, QCheckBox, QSpinBox,
                               QTabWidget, QWidget, QGridLayout, QGroupBox,
                               QScrollArea, QFileDialog)
from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtGui import QFont, QIcon, QPalette, QTextOption

from salary_mail.db_instance import SalaryEmail, User


class PySideEmailSettingWin(QDialog):
    """PySide6版本的邮箱设置窗口"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.db = parent.db if parent else None
        
        self.setWindowTitle('邮箱设置')
        self.setModal(True)
        self.resize(600, 700)
        
        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(__file__), 'assets', 'settings.png')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        # 居中显示
        self.center_window()
        
        self.setup_ui()
        self.load_email_config()
    
    def center_window(self):
        """窗口居中显示"""
        if self.parent():
            parent_geo = self.parent().geometry()
            x = parent_geo.x() + (parent_geo.width() - self.width()) // 2
            y = parent_geo.y() + (parent_geo.height() - self.height()) // 2
            self.move(x, y)
        else:
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
            QGroupBox {
                font-weight: bold;
                border: 2px solid #e2e8f0;
                border-radius: 15px;
                margin-top: 1ex;
                padding-top: 15px;
                background: rgba(255, 255, 255, 0.95);
                box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 10px 0 10px;
                color: #2c3e50;
                font-size: 14px;
            }
            QLineEdit, QComboBox, QSpinBox {
                padding: 6px 12px;
                border: 2px solid #e8f4f8;
                border-radius: 12px;
                font-size: 14px;
                background: rgba(255, 255, 255, 0.9);
                color: #2c3e50;
            }
            QLineEdit:focus, QComboBox:focus, QSpinBox:focus {
                border-color: #f59e0b;
                background: rgba(255, 255, 255, 1);
                box-shadow: 0 0 15px rgba(245, 158, 11, 0.3);
            }
            QLineEdit::placeholder {
                color: #94a3b8;
            }
            QTextEdit {
                border: 2px solid #e8f4f8;
                border-radius: 12px;
                background: rgba(255, 255, 255, 0.9);
                color: #2c3e50;
                font-family: "Consolas", "Microsoft YaHei UI";
            }
            QTextEdit:focus {
                border-color: #f59e0b;
                background: rgba(255, 255, 255, 1);
                box-shadow: 0 0 15px rgba(245, 158, 11, 0.3);
            }
            QPushButton {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #f59e0b, stop: 1 #d97706);
                color: white;
                border: none;
                border-radius: 12px;
                padding: 12px 20px;
                font-size: 14px;
                font-weight: bold;
                box-shadow: 0 4px 15px rgba(245, 158, 11, 0.3);
            }
            QPushButton:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #d97706, stop: 1 #b45309);
                transform: translateY(-2px);
                box-shadow: 0 6px 20px rgba(245, 158, 11, 0.4);
            }
            QPushButton:pressed {
                transform: translateY(0px);
                box-shadow: 0 2px 10px rgba(245, 158, 11, 0.3);
            }
            QPushButton#test_btn {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #10b981, stop: 1 #059669);
                box-shadow: 0 4px 15px rgba(16, 185, 129, 0.3);
            }
            QPushButton#test_btn:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #059669, stop: 1 #047857);
                box-shadow: 0 6px 20px rgba(16, 185, 129, 0.4);
            }
            QPushButton#save_btn {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #3b82f6, stop: 1 #1d4ed8);
                box-shadow: 0 4px 15px rgba(59, 130, 246, 0.3);
            }
            QPushButton#save_btn:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #1d4ed8, stop: 1 #1e3a8a);
                box-shadow: 0 6px 20px rgba(59, 130, 246, 0.4);
            }
            QLabel#title_label {
                color: #2c3e50;
                font-weight: bold;
                text-shadow: 1px 1px 3px rgba(0, 0, 0, 0.1);
            }
            QLabel {
                color: #374151;
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
                    stop: 0 #f59e0b, stop: 1 #d97706);
                border-color: #f59e0b;
            }
            QScrollArea {
                border: none;
                background: transparent;
            }
            QLabel {
                color: #374151;
            }
        """)
        
        # 主布局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # 标题
        title_label = QLabel("邮箱设置")
        title_label.setObjectName("title_label")
        title_label.setFont(QFont("Microsoft YaHei UI", 20, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 创建滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        # 滚动内容
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setSpacing(15)
        
        # 账号设置区域
        self.create_account_section(scroll_layout)
        
        # 服务器设置区域
        self.create_server_section(scroll_layout)
        
        # 发送设置区域
        self.create_send_section(scroll_layout)
        
        scroll_area.setWidget(scroll_content)
        main_layout.addWidget(scroll_area)
        
        # 按钮区域
        self.create_button_section(main_layout)
    
    def create_account_section(self, layout):
        """创建账号设置区域"""
        account_group = QGroupBox("📧 账号设置")
        account_group.setFont(QFont("Microsoft YaHei UI", 14, QFont.Bold))
        account_layout = QGridLayout(account_group)
        account_layout.setSpacing(10)
        
        # 发送邮箱
        account_layout.addWidget(QLabel("发送邮箱:"), 0, 0)
        self.email_entry = QLineEdit()
        self.email_entry.setPlaceholderText("example@company.com")
        account_layout.addWidget(self.email_entry, 0, 1)
        
        # 邮箱密码/授权码
        account_layout.addWidget(QLabel("密码/授权码:"), 1, 0)
        self.password_entry = QLineEdit()
        self.password_entry.setEchoMode(QLineEdit.Password)
        self.password_entry.setPlaceholderText("邮箱密码或授权码")
        account_layout.addWidget(self.password_entry, 1, 1)
        
        # 显示密码复选框
        self.show_password_cb = QCheckBox("显示密码")
        self.show_password_cb.toggled.connect(self.toggle_password_visibility)
        account_layout.addWidget(self.show_password_cb, 1, 2)
        
        # 发送者姓名
        account_layout.addWidget(QLabel("发送者姓名:"), 2, 0)
        self.sender_name_entry = QLineEdit()
        self.sender_name_entry.setPlaceholderText("公司财务部")
        account_layout.addWidget(self.sender_name_entry, 2, 1)
        
        layout.addWidget(account_group)
    
    def create_server_section(self, layout):
        """创建服务器设置区域"""
        server_group = QGroupBox("🔧 服务器设置")
        server_group.setFont(QFont("Microsoft YaHei UI", 14, QFont.Bold))
        server_layout = QGridLayout(server_group)
        server_layout.setSpacing(10)
        
        # SMTP服务器
        server_layout.addWidget(QLabel("SMTP服务器:"), 0, 0)
        self.smtp_server_entry = QLineEdit()
        self.smtp_server_entry.setPlaceholderText("smtp.company.com")
        server_layout.addWidget(self.smtp_server_entry, 0, 1)
        
        # SMTP端口
        server_layout.addWidget(QLabel("SMTP端口:"), 1, 0)
        self.smtp_port_spinbox = QSpinBox()
        self.smtp_port_spinbox.setRange(1, 65535)
        self.smtp_port_spinbox.setValue(587)
        server_layout.addWidget(self.smtp_port_spinbox, 1, 1)
        
        # 加密方式
        server_layout.addWidget(QLabel("加密方式:"), 2, 0)
        self.encryption_combo = QComboBox()
        self.encryption_combo.addItems(["TLS", "SSL", "无加密"])
        server_layout.addWidget(self.encryption_combo, 2, 1)
        
        # 常用邮箱快速设置
        server_layout.addWidget(QLabel("快速设置:"), 3, 0)
        self.quick_setup_combo = QComboBox()
        self.quick_setup_combo.addItems([
            "请选择...", "QQ邮箱", "163邮箱", "126邮箱", "Gmail", "Outlook", "企业邮箱"
        ])
        self.quick_setup_combo.currentTextChanged.connect(self.on_quick_setup_changed)
        server_layout.addWidget(self.quick_setup_combo, 3, 1)
        
        layout.addWidget(server_group)
    
    def create_send_section(self, layout):
        """创建发送设置区域"""
        send_group = QGroupBox("📤 发送设置")
        send_group.setFont(QFont("Microsoft YaHei UI", 14, QFont.Bold))
        send_layout = QGridLayout(send_group)
        send_layout.setSpacing(10)
        
        # 邮件主题
        send_layout.addWidget(QLabel("邮件主题:"), 0, 0)
        self.subject_entry = QLineEdit()
        self.subject_entry.setPlaceholderText("{company_name} {salary_month} 工资条")
        send_layout.addWidget(self.subject_entry, 0, 1)
        
        # 发送间隔
        send_layout.addWidget(QLabel("发送间隔(秒):"), 1, 0)
        self.interval_spinbox = QSpinBox()
        self.interval_spinbox.setRange(0, 60)
        self.interval_spinbox.setValue(1)
        send_layout.addWidget(self.interval_spinbox, 1, 1)
        
        # 批量发送数量
        send_layout.addWidget(QLabel("批量发送数量:"), 2, 0)
        self.batch_size_spinbox = QSpinBox()
        self.batch_size_spinbox.setRange(1, 100)
        self.batch_size_spinbox.setValue(10)
        send_layout.addWidget(self.batch_size_spinbox, 2, 1)
        
        layout.addWidget(send_group)
    
    def create_button_section(self, layout):
        """创建按钮区域"""
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        # 测试连接按钮
        test_btn = QPushButton("测试连接")
        test_btn.setObjectName("test_btn")
        test_btn.clicked.connect(self.test_connection)
        button_layout.addWidget(test_btn)
        
        button_layout.addStretch()
        
        # 保存按钮
        save_btn = QPushButton("保存设置")
        save_btn.setObjectName("save_btn")
        save_btn.clicked.connect(self.save_settings)
        button_layout.addWidget(save_btn)
        
        # 取消按钮
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.close)
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(button_layout)
    
    def toggle_password_visibility(self, checked):
        """切换密码可见性"""
        if checked:
            self.password_entry.setEchoMode(QLineEdit.Normal)
        else:
            self.password_entry.setEchoMode(QLineEdit.Password)
    
    def on_quick_setup_changed(self, text):
        """快速设置改变"""
        configs = {
            "QQ邮箱": ("smtp.qq.com", 587, "TLS"),
            "163邮箱": ("smtp.163.com", 587, "TLS"),
            "126邮箱": ("smtp.126.com", 587, "TLS"),
            "Gmail": ("smtp.gmail.com", 587, "TLS"),
            "Outlook": ("smtp.office365.com", 587, "TLS"),
            "企业邮箱": ("smtp.exmail.qq.com", 587, "TLS")
        }
        
        if text in configs:
            server, port, encryption = configs[text]
            self.smtp_server_entry.setText(server)
            self.smtp_port_spinbox.setValue(port)
            self.encryption_combo.setCurrentText(encryption)
    
    def load_email_config(self):
        """加载邮箱配置"""
        if not self.db:
            return
        
        try:
            config = self.db.query(SalaryEmail).first()
            if config:
                self.email_entry.setText(config.email or '')
                self.sender_name_entry.setText(config.sender_name or '')
                self.smtp_server_entry.setText(config.smtp_server or '')
                self.smtp_port_spinbox.setValue(config.smtp_port or 587)
                self.subject_entry.setText(config.email_subject or '')
                
                # 解码密码
                if config.password:
                    try:
                        password = base64.b64decode(config.password.encode()).decode()
                        self.password_entry.setText(password)
                    except:
                        pass
        except Exception as e:
            print(f"加载邮箱配置失败: {e}")
    
    def test_connection(self):
        """测试邮箱连接"""
        email = self.email_entry.text().strip()
        password = self.password_entry.text().strip()
        smtp_server = self.smtp_server_entry.text().strip()
        smtp_port = self.smtp_port_spinbox.value()
        
        if not all([email, password, smtp_server]):
            QMessageBox.warning(self, "提示", "请填写完整的邮箱配置信息")
            return
        
        # 这里应该实现真正的连接测试
        QMessageBox.information(self, "测试结果", "连接测试功能正在开发中，敬请期待！")
    
    def save_settings(self):
        """保存设置"""
        if not self.db:
            return
        
        email = self.email_entry.text().strip()
        password = self.password_entry.text().strip()
        
        if not email or not password:
            QMessageBox.warning(self, "提示", "邮箱和密码不能为空")
            return
        
        if not self.is_valid_email(email):
            QMessageBox.warning(self, "提示", "请输入有效的邮箱地址")
            return
        
        try:
            # 查找或创建邮箱配置
            config = self.db.query(SalaryEmail).first()
            if not config:
                config = SalaryEmail()
                self.db.add(config)
            
            # 更新配置
            config.email = email
            config.password = base64.b64encode(password.encode()).decode()
            config.sender_name = self.sender_name_entry.text().strip()
            config.smtp_server = self.smtp_server_entry.text().strip()
            config.smtp_port = self.smtp_port_spinbox.value()
            config.email_subject = self.subject_entry.text().strip()
            
            self.db.commit()
            
            QMessageBox.information(self, "成功", "邮箱设置保存成功！")
            self.accept()
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存邮箱设置失败: {str(e)}")
            self.db.rollback()
    
    def is_valid_email(self, email):
        """验证邮箱格式"""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None


class PySideTemplateSettingWin(QDialog):
    """PySide6版本的模板设置窗口"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.db = parent.db if parent else None
        
        self.setWindowTitle('邮件模板设置')
        self.setModal(True)
        self.resize(800, 600)
        
        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(__file__), 'assets', 'template.png')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        # 居中显示
        self.center_window()
        
        self.setup_ui()
        self.load_template()
    
    def center_window(self):
        """窗口居中显示"""
        if self.parent():
            parent_geo = self.parent().geometry()
            x = parent_geo.x() + (parent_geo.width() - self.width()) // 2
            y = parent_geo.y() + (parent_geo.height() - self.height()) // 2
            self.move(x, y)
        else:
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
            QGroupBox {
                font-weight: bold;
                border: 2px solid #e2e8f0;
                border-radius: 15px;
                margin-top: 1ex;
                padding-top: 15px;
                background: rgba(255, 255, 255, 0.95);
                box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 10px 0 10px;
                color: #2c3e50;
                font-size: 14px;
            }
            QTextEdit {
                border: 2px solid #e8f4f8;
                border-radius: 12px;
                background: rgba(255, 255, 255, 0.95);
                color: #2c3e50;
                font-family: "Consolas", "Microsoft YaHei UI", monospace;
                font-size: 12px;
                line-height: 1.5;
            }
            QTextEdit:focus {
                border-color: #a855f7;
                background: rgba(255, 255, 255, 1);
                box-shadow: 0 0 15px rgba(168, 85, 247, 0.3);
            }
            QPushButton {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #a855f7, stop: 1 #7c3aed);
                color: white;
                border: none;
                border-radius: 12px;
                padding: 12px 20px;
                font-size: 14px;
                font-weight: bold;
                box-shadow: 0 4px 15px rgba(168, 85, 247, 0.3);
            }
            QPushButton:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #7c3aed, stop: 1 #5b21b6);
                transform: translateY(-2px);
                box-shadow: 0 6px 20px rgba(168, 85, 247, 0.4);
            }
            QPushButton:pressed {
                transform: translateY(0px);
                box-shadow: 0 2px 10px rgba(168, 85, 247, 0.3);
            }
            QPushButton#preview_btn {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #10b981, stop: 1 #059669);
                box-shadow: 0 4px 15px rgba(16, 185, 129, 0.3);
            }
            QPushButton#preview_btn:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #059669, stop: 1 #047857);
                box-shadow: 0 6px 20px rgba(16, 185, 129, 0.4);
            }
            QPushButton#save_btn {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #3b82f6, stop: 1 #1d4ed8);
                box-shadow: 0 4px 15px rgba(59, 130, 246, 0.3);
            }
            QPushButton#save_btn:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #1d4ed8, stop: 1 #1e3a8a);
                box-shadow: 0 6px 20px rgba(59, 130, 246, 0.4);
            }
            QLabel#title_label {
                color: #2c3e50;
                font-weight: bold;
                text-shadow: 1px 1px 3px rgba(0, 0, 0, 0.1);
            }
            QLabel {
                color: #374151;
            }
        """)
        
        # 主布局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # 标题
        title_label = QLabel("邮件模板设置")
        title_label.setObjectName("title_label")
        title_label.setFont(QFont("Microsoft YaHei UI", 20, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 模板编辑区域
        template_group = QGroupBox("📄 HTML模板编辑")
        template_group.setFont(QFont("Microsoft YaHei UI", 14, QFont.Bold))
        template_layout = QVBoxLayout(template_group)
        
        # 模板编辑器
        self.template_editor = QTextEdit()
        self.template_editor.setMinimumHeight(300)
        self.template_editor.setPlaceholderText("请输入HTML邮件模板...")
        template_layout.addWidget(self.template_editor)
        
        # 变量说明
        var_info = QLabel("""
可用变量：
{employee_name} - 员工姓名    {employee_id} - 员工编号    {salary_month} - 工资月份
{post_salary} - 岗位工资      {level_salary} - 薪级工资   {performance} - 绩效工资
{actual_salary} - 实发工资    {company_name} - 公司名称
        """)
        var_info.setStyleSheet("color: #666666; background-color: #f9f9f9; padding: 10px; border-radius: 5px;")
        var_info.setFont(QFont("Microsoft YaHei UI", 10))
        template_layout.addWidget(var_info)
        
        main_layout.addWidget(template_group)
        
        # 按钮区域
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        # 预览按钮
        preview_btn = QPushButton("预览模板")
        preview_btn.setObjectName("preview_btn")
        preview_btn.clicked.connect(self.preview_template)
        button_layout.addWidget(preview_btn)
        
        # 加载默认模板按钮
        load_default_btn = QPushButton("加载默认模板")
        load_default_btn.clicked.connect(self.load_default_template)
        button_layout.addWidget(load_default_btn)
        
        button_layout.addStretch()
        
        # 保存按钮
        save_btn = QPushButton("保存模板")
        save_btn.setObjectName("save_btn")
        save_btn.clicked.connect(self.save_template)
        button_layout.addWidget(save_btn)
        
        # 取消按钮
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.close)
        button_layout.addWidget(cancel_btn)
        
        main_layout.addLayout(button_layout)
    
    def load_template(self):
        """加载模板"""
        if not self.db:
            return
        
        try:
            config = self.db.query(SalaryEmail).first()
            if config and config.email_template:
                # 解码模板
                try:
                    template = base64.b64decode(config.email_template.encode()).decode()
                    self.template_editor.setPlainText(template)
                except:
                    pass
        except Exception as e:
            print(f"加载模板失败: {e}")
    
    def load_default_template(self):
        """加载默认模板"""
        default_template = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>工资条</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { text-align: center; color: #2196F3; margin-bottom: 20px; }
        .content { background-color: #f9f9f9; padding: 20px; border-radius: 8px; }
        .salary-item { margin: 10px 0; padding: 8px; background-color: white; border-radius: 4px; }
        .total { font-weight: bold; font-size: 18px; color: #4CAF50; }
    </style>
</head>
<body>
    <div class="header">
        <h2>{company_name}</h2>
        <h3>{salary_month} 工资条</h3>
    </div>
    
    <div class="content">
        <p><strong>员工姓名：</strong>{employee_name}</p>
        <p><strong>员工编号：</strong>{employee_id}</p>
        
        <div class="salary-item">
            <strong>岗位工资：</strong>¥{post_salary}
        </div>
        <div class="salary-item">
            <strong>薪级工资：</strong>¥{level_salary}
        </div>
        <div class="salary-item">
            <strong>绩效工资：</strong>¥{performance}
        </div>
        
        <div class="salary-item total">
            <strong>实发工资：</strong>¥{actual_salary}
        </div>
    </div>
    
    <p style="margin-top: 20px; color: #666666; font-size: 12px;">
        此邮件由系统自动发送，请勿回复。如有疑问请联系财务部门。
    </p>
</body>
</html>
        """
        self.template_editor.setPlainText(default_template.strip())
    
    def preview_template(self):
        """预览模板"""
        template = self.template_editor.toPlainText()
        if not template.strip():
            QMessageBox.warning(self, "提示", "请先输入模板内容")
            return
        
        # 使用示例数据预览
        preview_data = {
            'company_name': '示例公司',
            'salary_month': '2025年01月',
            'employee_name': '张三',
            'employee_id': 'E001',
            'post_salary': '5000.00',
            'level_salary': '2000.00',
            'performance': '1000.00',
            'actual_salary': '7200.00'
        }
        
        try:
            preview_html = template.format(**preview_data)
            
            # 创建临时文件并在浏览器中打开
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(preview_html)
                temp_file = f.name
            
            webbrowser.open(f'file://{temp_file}')
            QMessageBox.information(self, "预览", "模板已在浏览器中打开预览")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"预览模板失败: {str(e)}")
    
    def save_template(self):
        """保存模板"""
        if not self.db:
            return
        
        template = self.template_editor.toPlainText().strip()
        if not template:
            QMessageBox.warning(self, "提示", "模板内容不能为空")
            return
        
        try:
            # 查找或创建邮箱配置
            config = self.db.query(SalaryEmail).first()
            if not config:
                config = SalaryEmail()
                self.db.add(config)
            
            # 编码并保存模板
            config.email_template = base64.b64encode(template.encode()).decode()
            
            self.db.commit()
            
            QMessageBox.information(self, "成功", "邮件模板保存成功！")
            self.accept()
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存模板失败: {str(e)}")
            self.db.rollback()


class PySideInfoManageWin(QDialog):
    """PySide6版本的信息管理窗口"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.db = parent.db if parent else None
        self.current_user = parent.current_user if parent else None
        
        self.setWindowTitle('信息管理')
        self.setModal(True)
        self.resize(500, 400)
        
        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(__file__), 'assets', 'settings.png')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        # 居中显示
        self.center_window()
        
        self.setup_ui()
        self.load_user_info()
    
    def center_window(self):
        """窗口居中显示"""
        if self.parent():
            parent_geo = self.parent().geometry()
            x = parent_geo.x() + (parent_geo.width() - self.width()) // 2
            y = parent_geo.y() + (parent_geo.height() - self.height()) // 2
            self.move(x, y)
        else:
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
            QGroupBox {
                font-weight: bold;
                border: 2px solid #e2e8f0;
                border-radius: 15px;
                margin-top: 1ex;
                padding-top: 15px;
                background: rgba(255, 255, 255, 0.95);
                box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 10px 0 10px;
                color: #2c3e50;
                font-size: 14px;
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
                border-color: #8b5cf6;
                background: rgba(255, 255, 255, 1);
                box-shadow: 0 0 15px rgba(139, 92, 246, 0.3);
            }
            QLineEdit::placeholder {
                color: #94a3b8;
            }
            QPushButton {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #8b5cf6, stop: 1 #7c3aed);
                color: white;
                border: none;
                border-radius: 12px;
                padding: 12px 20px;
                font-size: 14px;
                font-weight: bold;
                box-shadow: 0 4px 15px rgba(139, 92, 246, 0.3);
            }
            QPushButton:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #7c3aed, stop: 1 #5b21b6);
                transform: translateY(-2px);
                box-shadow: 0 6px 20px rgba(139, 92, 246, 0.4);
            }
            QPushButton:pressed {
                transform: translateY(0px);
                box-shadow: 0 2px 10px rgba(139, 92, 246, 0.3);
            }
            QPushButton#save_btn {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #3b82f6, stop: 1 #1d4ed8);
                box-shadow: 0 4px 15px rgba(59, 130, 246, 0.3);
            }
            QPushButton#save_btn:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #1d4ed8, stop: 1 #1e3a8a);
                box-shadow: 0 6px 20px rgba(59, 130, 246, 0.4);
            }
            QLabel#title_label {
                color: #2c3e50;
                font-weight: bold;
                text-shadow: 1px 1px 3px rgba(0, 0, 0, 0.1);
            }
            QLabel {
                color: #374151;
            }
        """)
        
        # 主布局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # 标题
        title_label = QLabel("信息管理")
        title_label.setObjectName("title_label")
        title_label.setFont(QFont("Microsoft YaHei UI", 20, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 公司信息区域
        company_group = QGroupBox("🏢 公司信息")
        company_group.setFont(QFont("Microsoft YaHei UI", 14, QFont.Bold))
        company_layout = QGridLayout(company_group)
        company_layout.setSpacing(10)
        
        # 公司名称
        company_layout.addWidget(QLabel("公司名称:"), 0, 0)
        self.company_name_entry = QLineEdit()
        self.company_name_entry.setPlaceholderText("请输入公司名称")
        company_layout.addWidget(self.company_name_entry, 0, 1)
        
        main_layout.addWidget(company_group)
        
        # 密码修改区域
        password_group = QGroupBox("🔐 密码修改")
        password_group.setFont(QFont("Microsoft YaHei UI", 14, QFont.Bold))
        password_layout = QGridLayout(password_group)
        password_layout.setSpacing(10)
        
        # 当前密码
        password_layout.addWidget(QLabel("当前密码:"), 0, 0)
        self.current_password_entry = QLineEdit()
        self.current_password_entry.setEchoMode(QLineEdit.Password)
        self.current_password_entry.setPlaceholderText("请输入当前密码")
        password_layout.addWidget(self.current_password_entry, 0, 1)
        
        # 新密码
        password_layout.addWidget(QLabel("新密码:"), 1, 0)
        self.new_password_entry = QLineEdit()
        self.new_password_entry.setEchoMode(QLineEdit.Password)
        self.new_password_entry.setPlaceholderText("请输入新密码")
        password_layout.addWidget(self.new_password_entry, 1, 1)
        
        # 确认密码
        password_layout.addWidget(QLabel("确认密码:"), 2, 0)
        self.confirm_password_entry = QLineEdit()
        self.confirm_password_entry.setEchoMode(QLineEdit.Password)
        self.confirm_password_entry.setPlaceholderText("请再次输入新密码")
        password_layout.addWidget(self.confirm_password_entry, 2, 1)
        
        main_layout.addWidget(password_group)
        
        main_layout.addStretch()
        
        # 按钮区域
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        button_layout.addStretch()
        
        # 保存按钮
        save_btn = QPushButton("保存修改")
        save_btn.setObjectName("save_btn")
        save_btn.clicked.connect(self.save_changes)
        button_layout.addWidget(save_btn)
        
        # 取消按钮
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.close)
        button_layout.addWidget(cancel_btn)
        
        main_layout.addLayout(button_layout)
    
    def load_user_info(self):
        """加载用户信息"""
        if self.current_user:
            self.company_name_entry.setText(self.current_user.company_name or '')
    
    def save_changes(self):
        """保存修改"""
        if not self.db or not self.current_user:
            return
        
        company_name = self.company_name_entry.text().strip()
        current_password = self.current_password_entry.text().strip()
        new_password = self.new_password_entry.text().strip()
        confirm_password = self.confirm_password_entry.text().strip()
        
        try:
            # 更新公司名称
            if company_name != self.current_user.company_name:
                self.current_user.company_name = company_name
            
            # 如果要修改密码
            if current_password or new_password or confirm_password:
                if not current_password:
                    QMessageBox.warning(self, "提示", "请输入当前密码")
                    return
                
                # 验证当前密码
                current_hash = hashlib.sha256(current_password.encode()).hexdigest()
                if current_hash != self.current_user.password:
                    QMessageBox.warning(self, "提示", "当前密码不正确")
                    return
                
                if not new_password:
                    QMessageBox.warning(self, "提示", "请输入新密码")
                    return
                
                if len(new_password) < 6:
                    QMessageBox.warning(self, "提示", "新密码长度不能少于6位")
                    return
                
                if new_password != confirm_password:
                    QMessageBox.warning(self, "提示", "两次输入的新密码不一致")
                    return
                
                # 更新密码
                new_hash = hashlib.sha256(new_password.encode()).hexdigest()
                self.current_user.password = new_hash
            
            self.db.commit()
            QMessageBox.information(self, "成功", "信息修改成功！")
            self.accept()
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存修改失败: {str(e)}")
            self.db.rollback()
