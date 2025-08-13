# coding:utf-8
"""
PySide6版本的主页面
"""
import os
import sys
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                               QGridLayout, QLabel, QPushButton, QFrame, 
                               QSpacerItem, QSizePolicy, QApplication, QMessageBox,
                               QDialog)
from PySide6.QtCore import Qt, QTimer, QSize
from PySide6.QtGui import QFont, QPalette, QIcon, QPixmap, QPainter
import traceback

from salary_mail.db_instance import set_db
from salary_mail.pyside_login_window import PySideLoginWindow


class PySideSettingsMenu(QDialog):
    """系统设置菜单"""
    
    def __init__(self, parent):
        super().__init__(parent)
        
        self.setWindowTitle('系统设置')
        self.setModal(True)
        self.setFixedSize(400, 420)
        
        # 居中显示
        self.center_on_parent(parent)
        
        self.parent = parent
        
        self.setup_ui()
    
    def center_on_parent(self, parent):
        """在父窗口中心显示"""
        parent_geo = parent.geometry()
        x = parent_geo.x() + (parent_geo.width() - self.width()) // 2
        y = parent_geo.y() + (parent_geo.height() - self.height()) // 2
        self.move(x, y)
    
    def setup_ui(self):
        """设置UI"""
        # 设置样式
        self.setStyleSheet("""
            QDialog {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1,
                    stop: 0 #f8fafc, stop: 1 #e2e8f0);
            }
            QFrame#main_frame {
                background: rgba(255, 255, 255, 0.95);
                border-radius: 25px;
                border: none;
                box-shadow: 0 10px 40px rgba(0, 0, 0, 0.1);
            }
            QPushButton {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #667eea, stop: 1 #764ba2);
                color: white;
                border: none;
                border-radius: 15px;
                padding: 18px;
                font-size: 14px;
                font-weight: bold;
                text-align: left;
                box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
            }
            QPushButton:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #5a67d8, stop: 1 #6b46c1);
                transform: translateY(-2px);
                box-shadow: 0 6px 20px rgba(102, 126, 234, 0.4);
            }
            QPushButton#email_btn {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #ff9500, stop: 1 #ff6b35);
                box-shadow: 0 4px 15px rgba(255, 149, 0, 0.3);
            }
            QPushButton#email_btn:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #e6850e, stop: 1 #e55a2b);
                box-shadow: 0 6px 20px rgba(255, 149, 0, 0.4);
            }
            QPushButton#template_btn {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #a855f7, stop: 1 #ec4899);
                box-shadow: 0 4px 15px rgba(168, 85, 247, 0.3);
            }
            QPushButton#template_btn:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #9333ea, stop: 1 #db2777);
                box-shadow: 0 6px 20px rgba(168, 85, 247, 0.4);
            }
            QPushButton#info_btn {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #8b5cf6, stop: 1 #06b6d4);
                box-shadow: 0 4px 15px rgba(139, 92, 246, 0.3);
            }
            QPushButton#info_btn:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #7c3aed, stop: 1 #0891b2);
                box-shadow: 0 6px 20px rgba(139, 92, 246, 0.4);
            }
            QPushButton#close_btn {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #6b7280, stop: 1 #4b5563);
                box-shadow: 0 4px 15px rgba(107, 114, 128, 0.3);
            }
            QPushButton#close_btn:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #4b5563, stop: 1 #374151);
                box-shadow: 0 6px 20px rgba(107, 114, 128, 0.4);
            }
            QLabel#title_label {
                color: #2c3e50;
                font-weight: bold;
                text-shadow: 1px 1px 3px rgba(0, 0, 0, 0.1);
            }
        """)
        
        # 主布局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)
        
        # 主容器
        main_frame = QFrame()
        main_frame.setObjectName("main_frame")
        frame_layout = QVBoxLayout(main_frame)
        frame_layout.setContentsMargins(20, 20, 20, 20)
        frame_layout.setSpacing(15)
        
        # 标题
        title_label = QLabel("系统设置")
        title_label.setObjectName("title_label")
        title_label.setFont(QFont("Microsoft YaHei UI", 18, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        frame_layout.addWidget(title_label)
        
        # 邮箱设置按钮
        email_btn = QPushButton("📧 邮箱设置")
        email_btn.setObjectName("email_btn")
        email_btn.clicked.connect(self.open_email_setting)
        frame_layout.addWidget(email_btn)
        
        # 模板设置按钮
        template_btn = QPushButton("📄 模板设置")
        template_btn.setObjectName("template_btn")
        template_btn.clicked.connect(self.open_template_setting)
        frame_layout.addWidget(template_btn)
        
        # 信息管理按钮
        info_btn = QPushButton("🏢 信息管理")
        info_btn.setObjectName("info_btn")
        info_btn.clicked.connect(self.open_info_manage)
        frame_layout.addWidget(info_btn)
        
        main_layout.addWidget(main_frame)
        
        # 关闭按钮
        close_btn = QPushButton("关闭")
        close_btn.setObjectName("close_btn")
        close_btn.clicked.connect(self.close)
        main_layout.addWidget(close_btn)
    
    def open_email_setting(self):
        """打开邮箱设置"""
        self.close()
        self.parent.show_email_setting()
    
    def open_template_setting(self):
        """打开模板设置"""
        self.close()
        self.parent.show_template_setting()
    
    def open_info_manage(self):
        """打开信息管理"""
        self.close()
        self.parent.show_info_manage()


class FeatureCard(QFrame):
    """功能卡片组件"""
    
    def __init__(self, title, description, icon_path, color, callback, parent=None):
        super().__init__(parent)
        self.callback = callback
        self.color = color
        self.title = title
        
        self.setFixedSize(300, 250)
        self.setCursor(Qt.PointingHandCursor)
        
        # 设置样式
        self.setStyleSheet(f"""
            FeatureCard {{
                background: rgba(255, 255, 255, 0.95);
                border: none;
                border-radius: 30px;
                box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
            }}
            FeatureCard:hover {{
                background: rgba(255, 255, 255, 1);
                box-shadow: 0 15px 45px rgba(0, 0, 0, 0.15);
                transform: translateY(-5px);
            }}
            QLabel#title_label {{
                color: {color};
                font-weight: bold;
                text-shadow: 1px 1px 3px rgba(0, 0, 0, 0.1);
            }}
            QLabel#desc_label {{
                color: #64748b;
                line-height: 1.6;
            }}
            QPushButton#action_btn {{
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 {color}, stop: 1 {self.darken_color(color)});
                color: white;
                border: none;
                border-radius: 20px;
                font-weight: bold;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
            }}
            QPushButton#action_btn:hover {{
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 {self.darken_color(color)}, stop: 1 {self.darken_color(self.darken_color(color))});
                transform: translateY(-2px);
                box-shadow: 0 6px 20px rgba(0, 0, 0, 0.3);
            }}
            QPushButton#action_btn:pressed {{
                transform: translateY(0px);
                box-shadow: 0 2px 10px rgba(0, 0, 0, 0.2);
            }}
        """)
        
        self.setup_ui(title, description, icon_path)
    
    def darken_color(self, color):
        """使颜色变暗用于悬停效果"""
        color_map = {
            "#4CAF50": "#388e3c",
            "#2196F3": "#1565c0", 
            "#FF9800": "#ef6c00",
            "#9C27B0": "#6a1b9a",
            "#388e3c": "#2e7d32",
            "#1565c0": "#0d47a1",
            "#ef6c00": "#e65100",
            "#6a1b9a": "#4a148c"
        }
        return color_map.get(color, color)
    
    def setup_ui(self, title, description, icon_path):
        """设置UI"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)
        
        # 图标和标题行
        header_layout = QHBoxLayout()
        header_layout.setSpacing(20)
        
        # 图标
        if icon_path and os.path.exists(icon_path):
            icon_label = QLabel()
            pixmap = QPixmap(icon_path)
            if not pixmap.isNull():
                scaled_pixmap = pixmap.scaled(32, 32, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                icon_label.setPixmap(scaled_pixmap)
            header_layout.addWidget(icon_label)
        
        # 标题
        title_label = QLabel(title)
        title_label.setObjectName("title_label")
        title_label.setFont(QFont("Microsoft YaHei UI", 24, QFont.Bold))
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        
        layout.addLayout(header_layout)
        
        # 描述文本
        desc_label = QLabel(description)
        desc_label.setObjectName("desc_label")
        desc_label.setFont(QFont("Microsoft YaHei UI", 16))
        desc_label.setWordWrap(True)
        desc_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        layout.addWidget(desc_label)
        
        layout.addStretch()
        
        # 操作按钮
        action_btn = QPushButton(f"打开{title}")
        action_btn.setObjectName("action_btn")
        action_btn.setFont(QFont("Microsoft YaHei UI", 16, QFont.Bold))
        action_btn.setFixedHeight(50)
        action_btn.clicked.connect(self.callback)
        layout.addWidget(action_btn)
    
    def mousePressEvent(self, event):
        """处理鼠标点击事件"""
        if event.button() == Qt.LeftButton:
            self.callback()
        super().mousePressEvent(event)


class PySideHomePage(QMainWindow):
    """PySide6版本的主页面"""
    
    def __init__(self):
        # 先登录
        self.db = set_db()
        login_window = PySideLoginWindow(self.db)
        
        if login_window.exec() != PySideLoginWindow.Accepted:
            # 登录取消，退出程序
            sys.exit()
        
        # 保存当前用户
        self.current_user = login_window.current_user
        
        super().__init__()
        
        self.setWindowTitle('工资条管理系统')
        self.setMinimumSize(1200, 800)
        
        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(__file__), 'assets', 'icon.ico')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        # 最大化窗口
        self.showMaximized()
        
        # 设置UI
        self.setup_ui()
        
        # 追踪打开的窗口
        self.open_windows = {
            'email': None,
            'info': None,
            'salary': None,
            'employee': None
        }
    
    def setup_ui(self):
        """设置用户界面"""
        # 中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 设置样式
        self.setStyleSheet("""
            QMainWindow {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1,
                    stop: 0 #f8fafc, stop: 1 #e2e8f0);
            }
            QFrame#header_frame {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                    stop: 0 #667eea, stop: 1 #764ba2);
                border-radius: 25px;
                border: none;
                box-shadow: 0 10px 40px rgba(102, 126, 234, 0.3);
            }
            QLabel#title_label {
                color: white;
                font-weight: bold;
                text-shadow: 2px 2px 8px rgba(0, 0, 0, 0.3);
            }
            QLabel#subtitle_label {
                color: rgba(255, 255, 255, 0.9);
                text-shadow: 1px 1px 4px rgba(0, 0, 0, 0.2);
            }
            QFrame#status_frame {
                background: rgba(255, 255, 255, 0.95);
                border-radius: 20px;
                border: none;
                box-shadow: 0 5px 20px rgba(0, 0, 0, 0.1);
            }
            QLabel#copyright_label {
                color: #64748b;
                font-weight: 500;
            }
            QLabel#status_label {
                color: #10b981;
                font-weight: 600;
            }
            QLabel {
                color: #374151;
            }
        """)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(40, 40, 40, 40)
        main_layout.setSpacing(40)
        
        # 创建各个部分
        self.create_header(main_layout)
        self.create_cards_section(main_layout)
        self.create_status_bar(main_layout)
    
    def create_header(self, layout):
        """创建顶部标题区域"""
        header_frame = QFrame()
        header_frame.setObjectName("header_frame")
        header_frame.setFixedHeight(150)
        
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(25, 25, 25, 25)
        header_layout.setSpacing(8)
        
        # 标题
        title_label = QLabel("工资条管理系统")
        title_label.setObjectName("title_label")
        title_label.setFont(QFont("Microsoft YaHei UI", 42, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        header_layout.addWidget(title_label)
        
        # 副标题
        subtitle_label = QLabel("现代化的工资条发送管理工具")
        subtitle_label.setObjectName("subtitle_label")
        subtitle_label.setFont(QFont("Microsoft YaHei UI", 18))
        subtitle_label.setAlignment(Qt.AlignCenter)
        header_layout.addWidget(subtitle_label)
        
        layout.addWidget(header_frame)
    
    def create_cards_section(self, layout):
        """创建功能卡片区域"""
        # 卡片容器
        cards_widget = QWidget()
        cards_layout = QGridLayout(cards_widget)
        cards_layout.setSpacing(30)
        
        # 获取图标路径
        assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
        
        # 员工管理卡片
        employee_icon = os.path.join(assets_dir, 'employee.png')
        employee_card = FeatureCard(
            "员工管理",
            "管理员工信息和邮箱\n支持批量导入和实时搜索",
            employee_icon,
            "#4CAF50",
            self.show_employee_manage
        )
        cards_layout.addWidget(employee_card, 0, 0)
        
        # 工资管理卡片
        salary_icon = os.path.join(assets_dir, 'salary.png')
        salary_card = FeatureCard(
            "工资管理",
            "导入工资数据并发送工资条\n支持批量发送和状态跟踪",
            salary_icon,
            "#2196F3",
            self.show_salary_manage
        )
        cards_layout.addWidget(salary_card, 0, 1)
        
        # 邮件设置卡片
        settings_icon = os.path.join(assets_dir, 'settings.png')
        email_card = FeatureCard(
            "邮件设置",
            "配置SMTP邮箱参数\n设置发送邮箱和服务器",
            settings_icon,
            "#FF9800",
            self.show_email_setting
        )
        cards_layout.addWidget(email_card, 1, 0)
        
        # 系统设置卡片
        template_icon = os.path.join(assets_dir, 'template.png')
        system_card = FeatureCard(
            "系统设置",
            "模板设置和信息管理\n密码修改和公司信息",
            template_icon,
            "#9C27B0",
            self.show_system_settings
        )
        cards_layout.addWidget(system_card, 1, 1)
        
        layout.addWidget(cards_widget)
    
    def create_status_bar(self, layout):
        """创建底部状态栏"""
        status_frame = QFrame()
        status_frame.setObjectName("status_frame")
        status_frame.setFixedHeight(60)
        
        status_layout = QHBoxLayout(status_frame)
        status_layout.setContentsMargins(20, 15, 20, 15)
        
        # 版权信息
        copyright_label = QLabel("© 2025 工资条管理系统")
        copyright_label.setObjectName("copyright_label")
        copyright_label.setFont(QFont("Microsoft YaHei UI", 12))
        status_layout.addWidget(copyright_label)
        
        status_layout.addStretch()
        
        # 状态信息
        self.status_label = QLabel("系统就绪")
        self.status_label.setObjectName("status_label")
        self.status_label.setFont(QFont("Microsoft YaHei UI", 12))
        status_layout.addWidget(self.status_label)
        
        layout.addWidget(status_frame)
    
    def update_status(self, text, color='#4CAF50'):
        """更新状态栏文本"""
        self.status_label.setText(text)
        self.status_label.setStyleSheet(f"color: {color};")
    
    # 功能窗口显示方法（暂时用消息框代替）
    def show_employee_manage(self):
        """显示员工管理窗口"""
        try:
            if self.open_windows.get('employee') and hasattr(self.open_windows['employee'], 'isVisible') and self.open_windows['employee'].isVisible():
                self.open_windows['employee'].raise_()
                self.open_windows['employee'].activateWindow()
                return
            
            self.update_status("正在打开员工管理...", "#FF9800")
            from salary_mail.pyside_employee_manage import PySideEmployeeManageWin
            dialog = PySideEmployeeManageWin(parent=self)
            self.open_windows['employee'] = dialog
            dialog.show()
            self.update_status("系统就绪")
        except Exception as e:
            print(f"打开员工管理窗口失败: {e}")
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"打开员工管理窗口失败: {str(e)}")
            self.update_status("系统就绪")
    
    def show_salary_manage(self):
        """显示工资管理窗口"""
        try:
            if self.open_windows.get('salary') and hasattr(self.open_windows['salary'], 'isVisible') and self.open_windows['salary'].isVisible():
                self.open_windows['salary'].raise_()
                self.open_windows['salary'].activateWindow()
                return
            
            self.update_status("正在打开工资管理...", "#FF9800")
            from salary_mail.pyside_salary_manage import PySideSalaryManageWin
            dialog = PySideSalaryManageWin(parent=self)
            self.open_windows['salary'] = dialog
            dialog.show()
            self.update_status("系统就绪")
        except Exception as e:
            print(f"打开工资管理窗口失败: {e}")
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"打开工资管理窗口失败: {str(e)}")
            self.update_status("系统就绪")
    
    def show_email_setting(self):
        """显示邮箱设置窗口"""
        try:
            if self.open_windows.get('email') and hasattr(self.open_windows['email'], 'isVisible') and self.open_windows['email'].isVisible():
                self.open_windows['email'].raise_()
                self.open_windows['email'].activateWindow()
                return
            
            self.update_status("正在打开邮箱设置...", "#FF9800")
            from salary_mail.pyside_settings import PySideEmailSettingWin
            dialog = PySideEmailSettingWin(parent=self)
            self.open_windows['email'] = dialog
            dialog.show()
            self.update_status("系统就绪")
        except Exception as e:
            print(f"打开邮箱设置窗口失败: {e}")
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"打开邮箱设置窗口失败: {str(e)}")
            self.update_status("系统就绪")
    
    def show_system_settings(self):
        """显示系统设置窗口"""
        # 创建设置菜单
        settings_menu = PySideSettingsMenu(self)
        settings_menu.exec()
    
    def show_template_setting(self):
        """显示模板设置窗口"""
        try:
            self.update_status("正在打开模板设置...", "#FF9800")
            from salary_mail.pyside_settings import PySideTemplateSettingWin
            dialog = PySideTemplateSettingWin(parent=self)
            dialog.exec()
            self.update_status("系统就绪")
        except Exception as e:
            print(f"打开模板设置窗口失败: {e}")
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"打开模板设置窗口失败: {str(e)}")
            self.update_status("系统就绪")
    
    def show_info_manage(self):
        """显示信息管理窗口"""
        try:
            if self.open_windows.get('info') and hasattr(self.open_windows['info'], 'isVisible') and self.open_windows['info'].isVisible():
                self.open_windows['info'].raise_()
                self.open_windows['info'].activateWindow()
                return
            
            self.update_status("正在打开信息管理...", "#FF9800")
            from salary_mail.pyside_settings import PySideInfoManageWin
            dialog = PySideInfoManageWin(parent=self)
            self.open_windows['info'] = dialog
            dialog.show()
            self.update_status("系统就绪")
        except Exception as e:
            print(f"打开信息管理窗口失败: {e}")
            traceback.print_exc()
            QMessageBox.critical(self, "错误", f"打开信息管理窗口失败: {str(e)}")
            self.update_status("系统就绪")
    
    def closeEvent(self, event):
        """处理窗口关闭事件"""
        # 关闭所有子窗口
        for window_key, window in self.open_windows.items():
            if window and hasattr(window, 'close'):
                try:
                    window.close()
                except:
                    pass
        
        event.accept()
