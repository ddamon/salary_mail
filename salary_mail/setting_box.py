import re
import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
from tkinter import filedialog
import base64
import json
import os
import sys
from threading import Timer
from datetime import datetime
import hashlib
import tempfile
import webbrowser
# 添加项目根目录到Python路径
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from salary_mail.db_instance import SalaryEmail  # 添加 SalaryRecord


class EmailSettingWin(tk.Toplevel):
    '''邮箱设置'''

    def __init__(self, parent):
        super(EmailSettingWin, self).__init__()
        self.title('邮箱设置')
        cx, cy = parent.get_center()
        self.geometry('400x400+{}+{}'.format(cx - 200, cy - 200))
        self.attributes("-topmost", 1)  # 保持在前
        self.resizable(width=False, height=False)  # 禁制拉伸大小
        self.parent = parent
        self.db = parent.db
        self.setupUI()

    def setupUI(self):
        # 创建主框架
        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(fill='both', expand=True)

        # 账号设置区域
        account_frame = ttk.LabelFrame(main_frame, text="账号设置", padding=10)
        account_frame.pack(fill='x', pady=(0, 15))

        self.email_address = tk.StringVar()
        self.password = tk.StringVar()
        self.sender_name = tk.StringVar()

        # 加载现有数据
        try:
            email_address = self.db.query(SalaryEmail).filter(SalaryEmail.field_name == 'sender').first()
            password = self.db.query(SalaryEmail).filter(SalaryEmail.field_name == 'password').first()
            sender_name = self.db.query(SalaryEmail).filter(SalaryEmail.field_name == 'sender_name').first()

            if email_address:
                self.email_address.set(email_address.field_value)
            if password:
                self.password.set(base64.decodebytes(password.field_value).decode('utf-8'))
            if sender_name:
                self.sender_name.set(sender_name.field_value)
        except Exception as e:
            tk.messagebox.showerror(title="错误", message="数据库错误，请重试！\n{}".format(e))

        # 发件邮箱
        email_frame = ttk.Frame(account_frame)
        email_frame.pack(fill='x', pady=5)
        ttk.Label(email_frame, text="发件邮箱：", width=15).pack(side=tk.LEFT)
        ttk.Entry(email_frame, textvariable=self.email_address, width=30).pack(side=tk.LEFT)

        # 密码
        pwd_frame = ttk.Frame(account_frame)
        pwd_frame.pack(fill='x', pady=5)
        ttk.Label(pwd_frame, text='密码(授权码)：', width=15).pack(side=tk.LEFT)
        ttk.Entry(pwd_frame, textvariable=self.password, show='*', width=30).pack(side=tk.LEFT)

        # 发件人名称
        name_frame = ttk.Frame(account_frame)
        name_frame.pack(fill='x', pady=5)
        ttk.Label(name_frame, text="发件人：", width=15).pack(side=tk.LEFT)
        ttk.Entry(name_frame, textvariable=self.sender_name, width=30).pack(side=tk.LEFT)

        # SMTP设置区域
        smtp_frame = ttk.LabelFrame(main_frame, text="SMTP设置", padding=10)
        smtp_frame.pack(fill='x', pady=(0, 15))

        self.smtp_server = tk.StringVar()
        self.port = tk.StringVar()

        try:
            smtp_server = self.db.query(SalaryEmail).filter(SalaryEmail.field_name=='smtp_server').first()
            port = self.db.query(SalaryEmail).filter(SalaryEmail.field_name=='port').first()
            if smtp_server:
                self.smtp_server.set(smtp_server.field_value)
            if port:
                self.port.set(port.field_value)
        except Exception as e:
            tk.messagebox.showerror(title="错误", message="数据库错误，请重试！\n{}".format(e))

        # SMTP服务器
        server_frame = ttk.Frame(smtp_frame)
        server_frame.pack(fill='x', pady=5)
        ttk.Label(server_frame, text="SMTP服务器：", width=15).pack(side=tk.LEFT)
        ttk.Entry(server_frame, textvariable=self.smtp_server, width=30).pack(side=tk.LEFT)

        # 端口
        port_frame = ttk.Frame(smtp_frame)
        port_frame.pack(fill='x', pady=5)
        ttk.Label(port_frame, text='SMTP端口：', width=15).pack(side=tk.LEFT)
        ttk.Entry(port_frame, textvariable=self.port, width=30).pack(side=tk.LEFT)

        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=(20, 0))
        ttk.Button(btn_frame, text='取消', command=self.cancel, width=10).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text='保存', command=self.save, width=10).pack(side=tk.RIGHT, padx=5)

    def save(self):
        """保存设置"""
        try:
            # 验证邮箱格式
            sender_text = self.email_address.get().strip()
            password_text = self.password.get().strip()
            sender_name_text = self.sender_name.get().strip()
            smtp_text = self.smtp_server.get().strip()
            port_text = self.port.get().strip()

            # 验证必填项
            if not all([sender_text, password_text, sender_name_text, smtp_text, port_text]):
                tk.messagebox.showwarning('提示', '所有字段都为必填项！', parent=self)
                return

            # 验证邮箱格式
            if not re.match(r'[0-9A-Za-z][\.-_0-9A-Za-z]*@[0-9A-Za-z]+(\.[0-9A-Za-z]+)+$', sender_text):
                tk.messagebox.showwarning('提示', '请输入正确的邮箱地址！', parent=self)
                return

            # 验证端口
            if port_text not in ('25', '465'):
                tk.messagebox.showwarning('提示', '请输入正确的端口号(25/465)！', parent=self)
                return

            # 保存所有设置
            fields = {
                'sender': sender_text,
                'password': base64.encodebytes(password_text.encode('utf-8')),
                'sender_name': sender_name_text,
                'smtp_server': smtp_text,
                'port': port_text
            }

            for field_name, value in fields.items():
                record = self.db.query(SalaryEmail).filter(SalaryEmail.field_name == field_name).first()
                if not record:
                    record = SalaryEmail(field_name=field_name)
                record.field_value = value
                self.db.add(record)

            self.db.commit()
            tk.messagebox.showinfo('成功', '设置保存成功！', parent=self)
            self.destroy()

        except Exception as e:
            tk.messagebox.showerror('错误', f'保存失败：\n{str(e)}', parent=self)
            self.db.rollback()

    def cancel(self):
        """取消设置"""
        if tk.messagebox.askyesno('确认', '确定要取消设置吗？', parent=self):
            self.destroy()


class TemplateSettingWin(tk.Toplevel):
    '''邮件模板设置'''

    def __init__(self, parent):
        super(TemplateSettingWin, self).__init__()
        self.title('邮件模板设置')
        cx, cy = parent.get_center()
        # 修改窗口大小
        self.geometry('800x600+{}+{}'.format(cx - 400, cy - 300))
        self.minsize(800, 600)
        self.attributes("-topmost", 1)
        self.resizable(width=True, height=True)
        self.parent = parent
        self.db = parent.db
        self.setupUI()

    def get_default_template(self):
        """获取默认模板"""
        return '''<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
.container {
    width: 100%;
    max-width: 800px;
    margin: 0 auto;
    font-family: "Microsoft YaHei", Arial, sans-serif;
    padding: 20px;
}
.header {
    text-align: center;
    margin-bottom: 30px;
    border-bottom: 2px solid #4299e1;
    padding-bottom: 20px;
}
.header h2 {
    color: #2b6cb0;
    font-size: 24px;
    margin: 0;
}
.greeting {
    margin: 20px 0;
    color: #2d3748;
    font-size: 16px;
    line-height: 1.6;
}
.salary-table {
    width: 100%;
    border-collapse: collapse;
    margin: 25px 0;
    border: 1px solid #e5e7eb;
}
.salary-table td {
    padding: 12px 15px;
    border: 1px solid #e5e7eb;
}
.salary-table tr:nth-child(even) {
    background-color: #f8fafc;
}
.salary-table td:last-child {
    text-align: right;
    font-family: Consolas, monospace;
}
.notice {
    margin: 20px 0;
    padding: 15px;
    background-color: #f0f9ff;
    border: 1px solid #bae6fd;
    border-radius: 4px;
    color: #0c4a6e;
}
.footer {
    margin-top: 30px;
    padding-top: 20px;
    border-top: 2px solid #e5e7eb;
    text-align: right;
    color: #4b5563;
}
</style>
</head>
<body>
<div class="container">
    <div class="header">
        <h2>{year}年{month}月工资条</h2>
    </div>
    
    <div class="greeting">
        <p>尊敬的 {name} 同志：</p>
        <p>以下是您{year}年{month}月份的工资明细，请查收。</p>
    </div>
    
    <table class="salary-table">
        {salary_items}
    </table>
    
    <div class="notice">
        <p><strong>温馨提示：</strong></p>
        <ol>
            <li>请严格遵守薪酬保密制度，不要以任何形式与他人提及个人的薪酬福利等内容；</li>
            <li>如对工资明细有任何疑问，请及时联系综合办公室。</li>
        </ol>
    </div>
    
    <div class="footer">
        <p>如有任何问题，欢迎随时与我们联系！</p>
        <p style="margin-top: 15px;">
            <strong>财务部</strong><br>
            <strong>综合办公室</strong>
        </p>
    </div>
</div>
</body>
</html>'''

    def setupUI(self):
        # 创建主框架，使用grid布局
        main_frame = ttk.Frame(self, padding=15)
        main_frame.grid(row=0, column=0, sticky='nsew')
        
        # 配置grid权重，使主框架可以填充整个窗口
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # 顶部说明区域
        help_frame = ttk.LabelFrame(main_frame, text="模板说明", padding=10)
        help_frame.grid(row=0, column=0, sticky='ew', pady=(0, 15))

        help_text = """可用的变量：
{name} - 员工姓名
{year} - 工资年份
{month} - 工资月份
{salary_items} - 工资条目表格
"""

        ttk.Label(
            help_frame,
            text=help_text,
            font=('Microsoft YaHei UI', 10),
            justify='left',
            wraplength=700
        ).pack(anchor='w')

        # 模板编辑区域
        template_frame = ttk.LabelFrame(main_frame, text="模板编辑", padding=10)
        template_frame.grid(row=1, column=0, sticky='nsew', pady=(0, 15))
        
        # 配置template_frame的grid权重
        main_frame.grid_rowconfigure(1, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        # 创建文本编辑器
        self.template_text = tk.Text(
            template_frame,
            wrap='word',
            font=('Consolas', 11),
            padx=10,
            pady=10
        )
        self.template_text.pack(fill='both', expand=True)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(template_frame, orient='vertical', command=self.template_text.yview)
        scrollbar.pack(side='right', fill='y')
        self.template_text.configure(yscrollcommand=scrollbar.set)

        # 加载现有模板
        try:
            template = self.db.query(SalaryEmail).filter(SalaryEmail.field_name=='email_template').first()
            if template and template.field_value:
                self.template_text.insert('1.0', template.field_value)
            else:
                self.template_text.insert('1.0', self.get_default_template())
        except Exception as e:
            tk.messagebox.showerror('错误', f'加载模板失败：\n{str(e)}', parent=self)

        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=2, column=0, sticky='ew', pady=(0, 5))

        ttk.Button(
            btn_frame,
            text="预览",
            command=self.preview_template,
            width=10
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            btn_frame,
            text="重置默认",
            command=self.reset_default,
            width=10
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            btn_frame,
            text="取消",
            command=self.cancel,
            width=10
        ).pack(side=tk.RIGHT, padx=5)

        ttk.Button(
            btn_frame,
            text="保存",
            command=self.save,
            width=10
        ).pack(side=tk.RIGHT, padx=5)

    def save(self):
        """保存模板"""
        try:
            template = self.db.query(SalaryEmail).filter(SalaryEmail.field_name=='email_template').first()
            if not template:
                template = SalaryEmail()
                template.field_name = 'email_template'
                template.memo = "邮件模板"

            template.field_value = self.template_text.get('1.0', 'end-1c')
            self.db.add(template)
            self.db.commit()
            tk.messagebox.showinfo('成功', '模板保存成功！', parent=self)
            self.destroy()
        except Exception as e:
            tk.messagebox.showerror('错误', f'保存失败：\n{str(e)}', parent=self)
            self.db.rollback()

    def cancel(self):
        """取消编辑"""
        if tk.messagebox.askyesno('确认', '确定要取消编辑吗？未保存的更改将丢失。', parent=self):
            self.destroy()

    def preview_template(self):
        """在浏览器中预览模板"""
        try:
            template = self.template_text.get('1.0', 'end-1c')

            # 确保模板中的大括号被正确转义
            template = template.replace('{', '{{').replace('}', '}}')
            # 恢复需要替换的变量占位符
            template = template.replace('{{name}}', '{name}')\
                             .replace('{{year}}', '{year}')\
                             .replace('{{month}}', '{month}')\
                             .replace('{{salary_items}}', '{salary_items}')

            preview_data = template.format(
                name="张三",
                year=datetime.now().year,
                month=datetime.now().month,
                salary_items=self.get_sample_table()
            )

            # 创建临时HTML文件
            with tempfile.NamedTemporaryFile('w', delete=False, suffix='.html', encoding='utf-8') as f:
                f.write(preview_data)
                temp_path = f.name

            # 在默认浏览器中打开
            webbrowser.open('file://' + temp_path)

            # 设置定时器删除临时文件
            def cleanup():
                try:
                    os.unlink(temp_path)
                except:
                    pass
            Timer(5, cleanup).start()  # 5秒后删除临时文件

        except Exception as e:
            tk.messagebox.showerror('预览失败',
                                  f'生成预览时出错：{str(e)}\n\n请检查模板中的变量是否正确。',
                                  parent=self)

    def get_sample_table(self):
        """生成示例表格数据"""
        return '''
        <tr><td>岗位工资</td><td>5,000.00</td></tr>
        <tr><td>薪级工资</td><td>2,000.00</td></tr>
        <tr><td>绩效工资</td><td>3,000.00</td></tr>
        <tr><td>餐补</td><td>500.00</td></tr>
        <tr><td>交补</td><td>300.00</td></tr>
        <tr><td>防暑降温费</td><td>200.00</td></tr>
        <tr><td>补发</td><td>0.00</td></tr>
        <tr><td>事假扣款</td><td>0.00</td></tr>
        <tr><td>病假扣款</td><td>0.00</td></tr>
        <tr><td>其他扣款</td><td>0.00</td></tr>
        <tr><td>税前工资</td><td>11,000.00</td></tr>
        <tr><td>社会保险</td><td>1,000.00</td></tr>
        <tr><td>公积金</td><td>600.00</td></tr>
        <tr><td>个人所得税</td><td>200.00</td></tr>
        <tr><td>代缴工会会费</td><td>20.00</td></tr>
        <tr style="background-color:#e8f5e9;font-weight:bold"><td>实发工资</td><td>9,180.00</td></tr>
        <tr>
            <td colspan="2" style="background-color:#fff3e0;padding:15px">
                <div style="font-weight:bold;margin-bottom:10px">备注</div>
                这是一个示例备注信息
            </td>
        </tr>
        '''

    def reset_default(self):
        self.template_text.delete('1.0', tk.END)
        self.template_text.insert('1.0', self.get_default_template())


class InfoManageWin(tk.Toplevel):
    '''信息管理窗口'''
    
    def __init__(self, parent):
        super().__init__()
        self.title('信息管理')
        cx, cy = parent.get_center()
        self.geometry('400x300+{}+{}'.format(cx - 200, cy - 150))
        self.attributes("-topmost", 1)
        self.resizable(False, False)
        self.parent = parent
        self.db = parent.db
        self.current_user = parent.current_user
        self.setupUI()
        
    def setupUI(self):
        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(fill='both', expand=True)
        
        # 公司名称
        company_frame = ttk.LabelFrame(main_frame, text="公司信息", padding=10)
        company_frame.pack(fill='x', pady=(0, 15))
        
        self.company_name = tk.StringVar(value=self.current_user.company_name)
        
        name_frame = ttk.Frame(company_frame)
        name_frame.pack(fill='x', pady=5)
        ttk.Label(name_frame, text="公司名称：", width=15).pack(side=tk.LEFT)
        ttk.Entry(name_frame, textvariable=self.company_name, width=30).pack(side=tk.LEFT)
        
        # 修改密码
        pwd_frame = ttk.LabelFrame(main_frame, text="修改密码", padding=10)
        pwd_frame.pack(fill='x', pady=(0, 15))
        
        self.old_pwd = tk.StringVar()
        self.new_pwd = tk.StringVar()
        self.confirm_pwd = tk.StringVar()
        
        # 原密码
        old_frame = ttk.Frame(pwd_frame)
        old_frame.pack(fill='x', pady=5)
        ttk.Label(old_frame, text="原密码：", width=15).pack(side=tk.LEFT)
        ttk.Entry(old_frame, textvariable=self.old_pwd, show='*', width=30).pack(side=tk.LEFT)
        
        # 新密码
        new_frame = ttk.Frame(pwd_frame)
        new_frame.pack(fill='x', pady=5)
        ttk.Label(new_frame, text="新密码：", width=15).pack(side=tk.LEFT)
        ttk.Entry(new_frame, textvariable=self.new_pwd, show='*', width=30).pack(side=tk.LEFT)
        
        # 确认密码
        confirm_frame = ttk.Frame(pwd_frame)
        confirm_frame.pack(fill='x', pady=5)
        ttk.Label(confirm_frame, text="确认新密码：", width=15).pack(side=tk.LEFT)
        ttk.Entry(confirm_frame, textvariable=self.confirm_pwd, show='*', width=30).pack(side=tk.LEFT)
        
        # 按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=(20, 0))
        ttk.Button(btn_frame, text="取消", command=self.destroy, width=10).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="保存", command=self.save, width=10).pack(side=tk.RIGHT, padx=5)
        
    def save(self):
        try:
            # 更新公司名称
            company_name = self.company_name.get().strip()
            if company_name:
                self.current_user.company_name = company_name
            
            # 修改密码
            old_pwd = self.old_pwd.get()
            new_pwd = self.new_pwd.get()
            confirm_pwd = self.confirm_pwd.get()
            
            if old_pwd or new_pwd or confirm_pwd:
                if not all([old_pwd, new_pwd, confirm_pwd]):
                    tk.messagebox.showwarning('提示', '请填写完整的密码信息！', parent=self)
                    return
                
                if new_pwd != confirm_pwd:
                    tk.messagebox.showwarning('提示', '两次输入的新密码不一致！', parent=self)
                    return
                
                # 验证原密码
                old_hash = hashlib.sha256(old_pwd.encode()).hexdigest()
                if old_hash != self.current_user.password:
                    tk.messagebox.showerror('错误', '原密码错误！', parent=self)
                    return
                
                # 更新密码
                self.current_user.password = hashlib.sha256(new_pwd.encode()).hexdigest()
            
            self.db.commit()
            tk.messagebox.showinfo('成功', '保存成功！', parent=self)
            self.destroy()
            
        except Exception as e:
            tk.messagebox.showerror('错误', f'保存失败：\n{str(e)}', parent=self)
            self.db.rollback()



