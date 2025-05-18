import tkinter as tk
from tkinter import ttk
import hashlib
from datetime import datetime
from salary_mail.db_instance import User
import os

class LoginWindow(tk.Tk):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.title('登录 - 工资条管理系统')
        
        # 设置窗口样式
        self.configure(bg='white')
        self.resizable(False, False)  # 禁止调整大小
        
        self.setup_ui()
        self.center_window(400, 500)  # 增加窗口高度
        
        # 设置窗口焦点
        self.focus_force()
        self.username_entry.focus_set()

    def setup_ui(self):
        style = ttk.Style()
        style.configure('TFrame', background='white')
        style.configure('TLabel', background='white')
        
        # 主框架
        main_frame = ttk.Frame(self, padding=30)
        main_frame.pack(fill='both', expand=True)
        
        # 标题
        title_label = ttk.Label(
            main_frame,
            text="工资条管理系统",
            font=('Microsoft YaHei UI', 24, 'bold'),
            foreground='#1976d2'
        )
        title_label.pack(pady=(0, 10))
        
        # 副标题
        subtitle_label = ttk.Label(
            main_frame,
            text="请登录以继续使用",
            font=('Microsoft YaHei UI', 12),
            foreground='#666666'
        )
        subtitle_label.pack(pady=(0, 30))
        
        # 登录框区域
        login_frame = ttk.Frame(main_frame)
        login_frame.pack(fill='x', pady=10)
        
        # 用户名
        username_frame = ttk.Frame(login_frame)
        username_frame.pack(fill='x', pady=10)
        ttk.Label(
            username_frame, 
            text="用户名", 
            font=('Microsoft YaHei UI', 10),
            foreground='#333333'
        ).pack(anchor='w', pady=(0, 5))
        
        self.username_var = tk.StringVar()
        self.username_entry = ttk.Entry(
            username_frame,
            textvariable=self.username_var,
            width=30,
            font=('Microsoft YaHei UI', 11)
        )
        self.username_entry.pack(fill='x', ipady=5)
        
        # 密码
        password_frame = ttk.Frame(login_frame)
        password_frame.pack(fill='x', pady=10)
        ttk.Label(
            password_frame,
            text="密码",
            font=('Microsoft YaHei UI', 10),
            foreground='#333333'
        ).pack(anchor='w', pady=(0, 5))
        
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(
            password_frame,
            textvariable=self.password_var,
            show='●',  # 使用圆点代替星号
            width=30,
            font=('Microsoft YaHei UI', 11)
        )
        self.password_entry.pack(fill='x', ipady=5)
        
        # 登录按钮 - 使用 tk.Button 替代 ttk.Button
        self.login_btn = tk.Button(
            login_frame,
            text="登 录",
            command=self.login,
            font=('Microsoft YaHei UI', 11),
            fg='white',
            bg='#1976d2',  # 按钮背景色
            activebackground='#1565c0',  # 鼠标悬停时的颜色
            activeforeground='white',  # 鼠标悬停时的文字颜色
            relief='flat',  # 扁平化效果
            cursor='hand2',  # 鼠标悬停时显示手型
            width=20,  # 设置按钮宽度
            height=2   # 设置按钮高度
        )
        self.login_btn.pack(fill='x', pady=(20, 0))
        
        # 绑定鼠标事件以实现悬停效果
        def on_enter(e):
            self.login_btn['background'] = '#1565c0'

        def on_leave(e):
            self.login_btn['background'] = '#1976d2'

        self.login_btn.bind('<Enter>', on_enter)
        self.login_btn.bind('<Leave>', on_leave)
        
        # 版权信息
        ttk.Label(
            main_frame,
            text="© 2024 工资条管理系统",
            font=('Microsoft YaHei UI', 9),
            foreground='#999999'
        ).pack(side='bottom', pady=20)
        
        # 绑定事件
        self.bind('<Return>', lambda e: self.login())
        self.username_entry.bind('<Tab>', lambda e: self.password_entry.focus_set())
        self.password_entry.bind('<Tab>', lambda e: self.login_btn.focus_set())
        self.login_btn.bind('<Tab>', lambda e: self.username_entry.focus_set())
        
        # 首次使用提示
        user_count = self.db.query(User).count()
        if user_count == 1:
            first_user = self.db.query(User).first()
            if first_user.username == 'admin' and first_user.last_login is None:
                self.after(500, lambda: tk.messagebox.showinfo(
                    '欢迎使用',
                    '首次使用系统\n默认用户名：admin\n默认密码：admin123\n请登录后及时修改密码！',
                    parent=self
                ))
                self.username_var.set('admin')
                self.password_entry.focus()

    def login(self):
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        
        if not username or not password:
            tk.messagebox.showwarning('提示', '用户名和密码不能为空！', parent=self)
            return
            
        # 密码加密
        pwd_hash = hashlib.sha256(password.encode()).hexdigest()
        
        # 验证用户
        user = self.db.query(User).filter_by(username=username).first()
        if not user or user.password != pwd_hash:
            tk.messagebox.showerror('错误', '用户名或密码错误！', parent=self)
            return
            
        # 更新登录时间
        user.last_login = datetime.now()
        self.db.commit()
        
        # 保存登录用户信息
        self.current_user = user
        self.destroy()
    
    def center_window(self, width, height):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f'{width}x{height}+{x}+{y}') 