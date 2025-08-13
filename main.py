# coding:utf-8
"""
工资条邮件管理系统 - PySide6版本主程序入口
"""
import os
import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon

# 获取项目根目录路径
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# 将项目根目录添加到Python路径
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

try:
    # 导入PySide版本的主页面
    from salary_mail.pyside_home_page import PySideHomePage
    
    if __name__ == '__main__':
        # 创建QApplication实例
        app = QApplication(sys.argv)
        
        # 设置应用程序属性
        app.setApplicationName("工资条管理系统")
        app.setApplicationVersion("2.0")
        app.setOrganizationName("工资条管理系统")
        
        # 设置应用程序图标
        icon_path = os.path.join(BASE_DIR, 'salary_mail', 'assets', 'icon.ico')
        if os.path.exists(icon_path):
            app.setWindowIcon(QIcon(icon_path))
        
        # 设置高DPI支持（PySide6中这些属性已被弃用，Qt6自动处理高DPI）
        # Qt6自动处理高DPI缩放，不需要手动设置
        pass
        
        # 创建并显示主窗口
        main_window = PySideHomePage()
        main_window.show()
        
        # 运行应用程序
        sys.exit(app.exec())
        
except ImportError as e:
    print(f"导入错误: {e}")
    print("请确保已安装PySide6: pip install PySide6")
    print("如果仍有问题，请检查依赖包是否完整安装")
    input("按回车键退出...")
except Exception as e:
    print(f"程序启动失败: {e}")
    input("按回车键退出...")
