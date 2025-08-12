# coding:utf-8
"""
工资条邮件管理系统 - 主程序入口
现代化CustomTkinter界面版本
"""
import os
import sys
import customtkinter as ctk

# 获取项目根目录路径
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# 将项目根目录添加到Python路径
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

try:
    # 初始化主题管理器
    from salary_mail.ctk_theme_manager import theme_manager
    
    # 导入主页面
    from salary_mail.ctk_home_page import CTKHomePage
    
    if __name__ == '__main__':
        # 初始化主题管理器并加载保存的主题
        theme_manager.load_theme_config()
        theme_manager.apply_theme()
        
        # 启动应用
        app = CTKHomePage()
        app.mainloop()
        
except ImportError as e:
    print(f"导入错误: {e}")
    print("请确保已安装CustomTkinter: pip install customtkinter")
    print("如果仍有问题，请检查依赖包是否完整安装")
    input("按回车键退出...")
except Exception as e:
    print(f"程序启动失败: {e}")
    input("按回车键退出...")

