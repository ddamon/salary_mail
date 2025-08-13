# coding:utf-8
"""
PySide6版本的登录窗口
"""
import hashlib
from datetime import datetime
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                               QLineEdit, QPushButton, QFrame, QMessageBox,
                               QWidget, QSpacerItem, QSizePolicy)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont, QPalette, QIcon, QPixmap
import os

from salary_mail.db_instance import User


class PySideLoginWindow(QDialog):
    """PySide6版本的登录窗口"""
    
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.current_user = None
        
        self.setWindowTitle('登录 - 工资条管理系统')
        self.setFixedSize(420, 480)
        self.setWindowFlags(Qt.Dialog | Qt.WindowCloseButtonHint)
        
        # 居中显示
        self.center_window()
        
        # 设置UI
        self.setup_ui()
        
        # 检查首次使用
        self.check_first_time_use()
        
        # 设置焦点
        QTimer.singleShot(100, lambda: self.username_entry.setFocus())
    
    def center_window(self):
        """窗口居中显示"""
        # 获取屏幕几何
        screen = self.screen()
        screen_rect = screen.geometry()
        
        # 计算居中位置
        x = (screen_rect.width() - self.width()) // 2
        y = (screen_rect.height() - self.height()) // 2
        
        self.move(x, y)
    
    def setup_ui(self):
        """设置用户界面"""
        # 设置样式
        self.setStyleSheet("""
            QDialog {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1,
                    stop: 0 #667eea, stop: 1 #764ba2);
            }
            QFrame#header_frame {
                background: rgba(255, 255, 255, 0.95);
                border-radius: 20px;
                border: none;
                box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
            }
            QFrame#form_frame {
                background: rgba(255, 255, 255, 0.95);
                border-radius: 25px;
                border: none;
                box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
            }
            QLabel#title_label {
                color: #2c3e50;
                font-weight: bold;
                text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.1);
            }
            QLabel#subtitle_label {
                color: #34495e;
            }
            QLabel#field_label {
                color: #2c3e50;
                font-weight: bold;
            }
            QLineEdit {
                padding: 8px 15px;
                border: 2px solid #e8f4f8;
                border-radius: 15px;
                font-size: 14px;
                background: rgba(255, 255, 255, 0.9);
                color: #2c3e50;
                transition: all 0.3s ease;
            }
            QLineEdit:focus {
                border-color: #667eea;
                background: rgba(255, 255, 255, 1);
                box-shadow: 0 0 15px rgba(102, 126, 234, 0.3);
            }
            QLineEdit::placeholder {
                color: #95a5a6;
            }
            QPushButton#login_btn {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #667eea, stop: 1 #764ba2);
                color: white;
                border: none;
                border-radius: 15px;
                padding: 15px;
                font-size: 16px;
                font-weight: bold;
                text-transform: uppercase;
                letter-spacing: 1px;
            }
            QPushButton#login_btn:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #5a67d8, stop: 1 #6b46c1);
                transform: translateY(-2px);
                box-shadow: 0 8px 25px rgba(102, 126, 234, 0.4);
            }
            QPushButton#login_btn:pressed {
                transform: translateY(0px);
                box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
            }
            QLabel#copyright_label {
                color: rgba(255, 255, 255, 0.8);
                font-weight: 500;
            }
            QLabel {
                color: #374151;
            }
        """)
        
        # 主布局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(20)
        
        # 创建各个部分
        self.create_header(main_layout)
        self.create_login_form(main_layout)
        self.create_footer(main_layout)
    
    def create_header(self, layout):
        """创建标题区域"""
        header_frame = QFrame()
        header_frame.setObjectName("header_frame")
        header_frame.setFixedHeight(120)
        
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(20, 20, 20, 20)
        header_layout.setSpacing(10)
        
        # 主标题
        title_label = QLabel("工资条管理系统")
        title_label.setObjectName("title_label")
        title_label.setFont(QFont("Microsoft YaHei UI", 28, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        header_layout.addWidget(title_label)
        
        # 副标题
        subtitle_label = QLabel("请登录以继续使用")
        subtitle_label.setObjectName("subtitle_label")
        subtitle_label.setFont(QFont("Microsoft YaHei UI", 14))
        subtitle_label.setAlignment(Qt.AlignCenter)
        header_layout.addWidget(subtitle_label)
        
        layout.addWidget(header_frame)
    
    def create_login_form(self, layout):
        """创建登录表单"""
        # 表单容器
        form_frame = QFrame()
        form_frame.setObjectName("form_frame")
        
        form_layout = QVBoxLayout(form_frame)
        form_layout.setContentsMargins(30, 30, 30, 30)
        form_layout.setSpacing(20)
        
        # 用户名输入
        username_label = QLabel("用户名")
        username_label.setObjectName("field_label")
        username_label.setFont(QFont("Microsoft YaHei UI", 14, QFont.Bold))
        form_layout.addWidget(username_label)
        
        self.username_entry = QLineEdit()
        self.username_entry.setPlaceholderText("请输入用户名")
        self.username_entry.setFont(QFont("Microsoft YaHei UI", 14))
        self.username_entry.setFixedHeight(40)
        self.username_entry.returnPressed.connect(self.on_username_return)
        form_layout.addWidget(self.username_entry)
        
        # 密码输入
        password_label = QLabel("密码")
        password_label.setObjectName("field_label")
        password_label.setFont(QFont("Microsoft YaHei UI", 14, QFont.Bold))
        form_layout.addWidget(password_label)
        
        self.password_entry = QLineEdit()
        self.password_entry.setPlaceholderText("请输入密码")
        self.password_entry.setFont(QFont("Microsoft YaHei UI", 14))
        self.password_entry.setFixedHeight(40)
        self.password_entry.setEchoMode(QLineEdit.Password)
        self.password_entry.returnPressed.connect(self.login)
        form_layout.addWidget(self.password_entry)
        
        # 登录按钮
        self.login_btn = QPushButton("登 录")
        self.login_btn.setObjectName("login_btn")
        self.login_btn.setFont(QFont("Microsoft YaHei UI", 16, QFont.Bold))
        self.login_btn.setFixedHeight(45)
        self.login_btn.clicked.connect(self.login)
        form_layout.addWidget(self.login_btn)
        
        layout.addWidget(form_frame)
    
    def create_footer(self, layout):
        """创建底部信息"""
        # 版权信息
        copyright_label = QLabel("© 2025 工资条管理系统")
        copyright_label.setObjectName("copyright_label")
        copyright_label.setFont(QFont("Microsoft YaHei UI", 12))
        copyright_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(copyright_label)
    
    def check_first_time_use(self):
        """检查首次使用"""
        user_count = self.db.query(User).count()
        if user_count == 1:
            first_user = self.db.query(User).first()
            if first_user.username == 'admin' and first_user.last_login is None:
                QTimer.singleShot(500, self.show_first_time_dialog)
                self.username_entry.setText('admin')
                QTimer.singleShot(600, lambda: self.password_entry.setFocus())
    
    def show_first_time_dialog(self):
        """显示首次使用对话框"""
        msg = QMessageBox(self)
        msg.setWindowTitle("欢迎使用")
        msg.setIcon(QMessageBox.Information)
        msg.setText("首次使用系统")
        msg.setInformativeText("默认用户名：admin\n默认密码：admin123\n请登录后及时修改密码！")
        msg.exec()
    
    def on_username_return(self):
        """处理用户名输入框回车键事件"""
        self.password_entry.setFocus()
    
    def login(self):
        """登录验证"""
        username = self.username_entry.text().strip()
        password = self.password_entry.text().strip()
        
        if not username or not password:
            self.show_error_dialog('用户名和密码不能为空！')
            return
        
        # 密码加密
        pwd_hash = hashlib.sha256(password.encode()).hexdigest()
        
        # 验证用户
        user = self.db.query(User).filter_by(username=username).first()
        if not user or user.password != pwd_hash:
            self.show_error_dialog('用户名或密码错误！')
            return
        
        # 更新登录时间
        user.last_login = datetime.now()
        self.db.commit()
        
        # 保存登录用户信息
        self.current_user = user
        
        # 接受对话框
        self.accept()
    
    def show_error_dialog(self, message):
        """显示错误对话框"""
        msg = QMessageBox(self)
        msg.setWindowTitle("错误")
        msg.setIcon(QMessageBox.Critical)
        msg.setText(message)
        msg.exec()
