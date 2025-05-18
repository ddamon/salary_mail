# coding:utf-8
import os
import sys
import tkinter as tk

# 获取项目根目录路径
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# 将项目根目录添加到Python路径
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from salary_mail.home_page import HomePage

if __name__ == '__main__':
    app = HomePage()
    app.mainloop()
