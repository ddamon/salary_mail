# coding:utf-8
"""
CustomTkinter版本的系统设置窗口
"""
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import re
import base64
import hashlib
import tempfile
import webbrowser
import os
from threading import Timer
from datetime import datetime

from salary_mail.db_instance import SalaryEmail, User
from salary_mail.ctk_components import CTKMessageBox, CTKWindowSizeManager
from salary_mail.theme_config import theme_manager as responsive_theme_manager

class CTKEmailSettingWin(ctk.CTkToplevel):
    """邮箱设置窗口"""
    
    def __init__(self, parent):
        super().__init__(parent)
        
        self.title('邮箱设置')
        
        # 设置响应式窗口配置
        self.responsive_config = responsive_theme_manager.setup_responsive_window(self, 'dialog_window')
        self.resizable(True, True)
        
        # 设置窗口属性
        self.transient(parent)
        self.grab_set()
        
        self.parent = parent
        self.db = parent.db
        
        self.setup_ui()
        
        # 确保窗口在父窗口中心显示
        self.after(100, lambda: responsive_theme_manager.center_window_on_parent(self, parent))
        
        # 设置焦点
        self.focus_force()
    
    
    def setup_ui(self):
        """设置UI"""
        # 主容器
        main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 标题
        title_label = ctk.CTkLabel(
            main_frame,
            text="邮箱设置",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=20, weight="bold")
        )
        title_label.pack(pady=(0, 20))
        
        # 创建滚动框架容器
        scroll_container = ctk.CTkFrame(main_frame, corner_radius=15)
        scroll_container.pack(fill="both", expand=True, pady=(0, 20))
        
        # 创建滚动框架
        scroll_width, scroll_height = CTKWindowSizeManager.get_scroll_frame_size(480, 400)
        self.scroll_frame = ctk.CTkScrollableFrame(
            scroll_container,
            width=scroll_width,
            height=scroll_height,
            corner_radius=10
        )
        self.scroll_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # 将所有表单内容放在滚动框架内
        content_frame = self.scroll_frame
        
        # 账号设置区域
        account_frame = ctk.CTkFrame(content_frame, corner_radius=10)
        account_frame.pack(fill="x", pady=(0, 15))
        
        # 账号设置标题
        account_title = ctk.CTkLabel(
            account_frame,
            text="📧 账号设置",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=16, weight="bold"),
            anchor="w"
        )
        account_title.pack(fill="x", padx=20, pady=(15, 10))
        
        # 账号设置内容
        account_content = ctk.CTkFrame(account_frame, fg_color="transparent")
        account_content.pack(fill="x", padx=20, pady=(0, 15))
        
        self.create_account_fields(account_content)
        
        # SMTP设置区域
        smtp_frame = ctk.CTkFrame(content_frame, corner_radius=10)
        smtp_frame.pack(fill="x", pady=(0, 20))
        
        # SMTP设置标题
        smtp_title = ctk.CTkLabel(
            smtp_frame,
            text="🌐 SMTP设置",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=16, weight="bold"),
            anchor="w"
        )
        smtp_title.pack(fill="x", padx=20, pady=(15, 10))
        
        # SMTP设置内容
        smtp_content = ctk.CTkFrame(smtp_frame, fg_color="transparent")
        smtp_content.pack(fill="x", padx=20, pady=(0, 15))
        
        self.create_smtp_fields(smtp_content)
        
        # 按钮区域（保持在main_frame中，不滚动）
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x", pady=10)
        
        # 取消按钮
        cancel_btn = ctk.CTkButton(
            button_frame,
            text="取消",
            command=self.cancel,
            width=100,
            height=40,
            fg_color="gray",
            hover_color="darkgray",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold")
        )
        cancel_btn.pack(side="right", padx=5)
        
        # 保存按钮
        save_btn = ctk.CTkButton(
            button_frame,
            text="保存",
            command=self.save,
            width=100,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold")
        )
        save_btn.pack(side="right", padx=5)
        
        # 绑定回车键
        self.bind('<Return>', lambda e: self.save())
    
    def create_account_fields(self, parent):
        """创建账号设置字段"""
        # 发件邮箱
        email_label = ctk.CTkLabel(
            parent,
            text="发件邮箱：",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            anchor="w"
        )
        email_label.pack(fill="x", pady=(10, 5))
        
        self.email_entry = ctk.CTkEntry(
            parent,
            placeholder_text="请输入发件邮箱地址",
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12)
        )
        self.email_entry.pack(fill="x", pady=(0, 10))
        
        # 密码
        password_label = ctk.CTkLabel(
            parent,
            text="密码(授权码)：",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            anchor="w"
        )
        password_label.pack(fill="x", pady=(0, 5))
        
        self.password_entry = ctk.CTkEntry(
            parent,
            placeholder_text="请输入SMTP授权码",
            show="●",
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12)
        )
        self.password_entry.pack(fill="x", pady=(0, 10))
        
        # 发件人名称
        sender_label = ctk.CTkLabel(
            parent,
            text="发件人：",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            anchor="w"
        )
        sender_label.pack(fill="x", pady=(0, 5))
        
        self.sender_entry = ctk.CTkEntry(
            parent,
            placeholder_text="请输入发件人显示名称",
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12)
        )
        self.sender_entry.pack(fill="x", pady=(0, 5))
        
        # 加载现有数据
        self.load_account_data()
    
    def create_smtp_fields(self, parent):
        """创建SMTP设置字段"""
        # SMTP服务器
        server_label = ctk.CTkLabel(
            parent,
            text="SMTP服务器：",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            anchor="w"
        )
        server_label.pack(fill="x", pady=(10, 5))
        
        self.server_entry = ctk.CTkEntry(
            parent,
            placeholder_text="例如：smtp.qq.com",
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12)
        )
        self.server_entry.pack(fill="x", pady=(0, 10))
        
        # SMTP端口
        port_label = ctk.CTkLabel(
            parent,
            text="SMTP端口：",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            anchor="w"
        )
        port_label.pack(fill="x", pady=(0, 5))
        
        # 端口选择框架
        port_frame = ctk.CTkFrame(parent, fg_color="transparent")
        port_frame.pack(fill="x", pady=(0, 10))
        
        self.port_var = tk.StringVar(value="465")
        
        # 465端口（推荐）
        port_465 = ctk.CTkRadioButton(
            port_frame,
            text="465 (SSL加密，推荐)",
            variable=self.port_var,
            value="465",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12)
        )
        port_465.pack(anchor="w", pady=2)
        
        # 25端口
        port_25 = ctk.CTkRadioButton(
            port_frame,
            text="25 (非加密)",
            variable=self.port_var,
            value="25",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12)
        )
        port_25.pack(anchor="w", pady=2)
        
        # 常用邮箱配置提示
        tip_label = ctk.CTkLabel(
            parent,
            text="💡 常用邮箱配置：\n• QQ邮箱：smtp.qq.com:465\n• 163邮箱：smtp.163.com:465\n• Gmail：smtp.gmail.com:465",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=10),
            text_color="gray",
            justify="left"
        )
        tip_label.pack(fill="x", pady=10)
        
        # 加载现有数据
        self.load_smtp_data()
    
    def load_account_data(self):
        """加载账号数据"""
        try:
            email_address = self.db.query(SalaryEmail).filter(SalaryEmail.field_name == 'sender').first()
            password = self.db.query(SalaryEmail).filter(SalaryEmail.field_name == 'password').first()
            sender_name = self.db.query(SalaryEmail).filter(SalaryEmail.field_name == 'sender_name').first()
            
            if email_address:
                self.email_entry.insert(0, email_address.field_value)
            if password:
                try:
                    decoded_password = base64.decodebytes(password.field_value).decode('utf-8')
                    self.password_entry.insert(0, decoded_password)
                except:
                    pass
            if sender_name:
                self.sender_entry.insert(0, sender_name.field_value)
        except Exception as e:
            print(f"加载账号数据失败: {e}")
    
    def load_smtp_data(self):
        """加载SMTP数据"""
        try:
            smtp_server = self.db.query(SalaryEmail).filter(SalaryEmail.field_name=='smtp_server').first()
            port = self.db.query(SalaryEmail).filter(SalaryEmail.field_name=='port').first()
            
            if smtp_server:
                self.server_entry.insert(0, smtp_server.field_value)
            if port:
                self.port_var.set(port.field_value)
        except Exception as e:
            print(f"加载SMTP数据失败: {e}")
    
    def save(self):
        """保存设置"""
        try:
            # 获取输入值
            sender_text = self.email_entry.get().strip()
            password_text = self.password_entry.get().strip()
            sender_name_text = self.sender_entry.get().strip()
            smtp_text = self.server_entry.get().strip()
            port_text = self.port_var.get()
            
            # 验证必填项
            if not all([sender_text, password_text, sender_name_text, smtp_text, port_text]):
                CTKMessageBox.show_warning(self, '提示', '所有字段都为必填项！')
                return
            
            # 验证邮箱格式
            if not re.match(r'[0-9A-Za-z][\.-_0-9A-Za-z]*@[0-9A-Za-z]+(\.[0-9A-Za-z]+)+$', sender_text):
                CTKMessageBox.show_warning(self, '提示', '请输入正确的邮箱地址！')
                return
            
            # 验证端口
            if port_text not in ('25', '465'):
                CTKMessageBox.show_warning(self, '提示', '请选择正确的端口号！')
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
            CTKMessageBox.show_info(self, '成功', '设置保存成功！')
            self.destroy()
            
        except Exception as e:
            CTKMessageBox.show_error(self, '错误', f'保存失败：\n{str(e)}')
            self.db.rollback()
    
    def cancel(self):
        """取消设置"""
        if CTKMessageBox.ask_yes_no(self, '确认', '确定要取消设置吗？') == "是":
            self.destroy()


class CTKTemplateSettingWin(ctk.CTkToplevel):
    """邮件模板设置窗口"""
    
    def __init__(self, parent):
        super().__init__(parent)
        
        self.title('邮件模板设置')
        # 使用窗口大小管理器适配高分辨率
        CTKWindowSizeManager.adjust_window_size(self, 1000, 800, 900, 700, (0.85, 0.8))
        
        # 设置窗口属性
        self.transient(parent)
        self.grab_set()
        
        # 确保窗口在父窗口中心显示
        self.after(100, lambda: responsive_theme_manager.center_window_on_parent(self, parent))
        
        self.parent = parent
        self.db = parent.db
        
        self.setup_ui()
        
        # 设置焦点
        self.focus_force()
    
    
    def setup_ui(self):
        """设置UI"""
        # 主容器
        main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 标题
        title_label = ctk.CTkLabel(
            main_frame,
            text="邮件模板设置",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=20, weight="bold")
        )
        title_label.pack(pady=(0, 20))
        
        # 说明区域
        help_frame = ctk.CTkFrame(main_frame, corner_radius=10)
        help_frame.pack(fill="x", pady=(0, 15))
        
        help_title = ctk.CTkLabel(
            help_frame,
            text="📝 模板说明",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=16, weight="bold"),
            anchor="w"
        )
        help_title.pack(fill="x", padx=20, pady=(15, 10))
        
        help_text = """可用的变量：
• {name} - 员工姓名
• {year} - 工资年份  
• {month} - 工资月份
• {salary_items} - 工资条目表格（自动生成）

支持HTML格式，可以使用CSS样式美化邮件外观。"""
        
        help_label = ctk.CTkLabel(
            help_frame,
            text=help_text,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12),
            justify="left",
            anchor="w"
        )
        help_label.pack(fill="x", padx=20, pady=(0, 15))
        
        # 模板编辑区域
        template_frame = ctk.CTkFrame(main_frame, corner_radius=10)
        template_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        template_title = ctk.CTkLabel(
            template_frame,
            text="📄 模板编辑",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=16, weight="bold"),
            anchor="w"
        )
        template_title.pack(fill="x", padx=20, pady=(15, 10))
        
        # 文本编辑区域
        text_frame = ctk.CTkFrame(template_frame, fg_color="transparent")
        text_frame.pack(fill="both", expand=True, padx=20, pady=(0, 15))
        
        # 创建文本编辑器
        self.template_text = ctk.CTkTextbox(
            text_frame,
            font=ctk.CTkFont(family="Consolas", size=11),
            corner_radius=5
        )
        self.template_text.pack(fill="both", expand=True)
        
        # 加载现有模板
        self.load_template()
        
        # 按钮区域
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x", pady=10)
        
        # 左侧按钮
        left_buttons = ctk.CTkFrame(button_frame, fg_color="transparent")
        left_buttons.pack(side="left")
        
        # 预览按钮
        preview_btn = ctk.CTkButton(
            left_buttons,
            text="👁️ 预览",
            command=self.preview_template,
            width=100,
            height=40,
            fg_color="#4CAF50",
            hover_color="#45a049",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold")
        )
        preview_btn.pack(side="left", padx=(0, 10))
        
        # 重置默认按钮
        reset_btn = ctk.CTkButton(
            left_buttons,
            text="🔄 重置默认",
            command=self.reset_default,
            width=120,
            height=40,
            fg_color="#FF9800",
            hover_color="#F57C00",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold")
        )
        reset_btn.pack(side="left")
        
        # 右侧按钮
        right_buttons = ctk.CTkFrame(button_frame, fg_color="transparent")
        right_buttons.pack(side="right")
        
        # 取消按钮
        cancel_btn = ctk.CTkButton(
            right_buttons,
            text="取消",
            command=self.cancel,
            width=100,
            height=40,
            fg_color="gray",
            hover_color="darkgray",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold")
        )
        cancel_btn.pack(side="right", padx=5)
        
        # 保存按钮
        save_btn = ctk.CTkButton(
            right_buttons,
            text="保存",
            command=self.save,
            width=100,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold")
        )
        save_btn.pack(side="right", padx=5)
    
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
    
    def load_template(self):
        """加载现有模板"""
        try:
            template = self.db.query(SalaryEmail).filter(SalaryEmail.field_name=='email_template').first()
            if template and template.field_value:
                self.template_text.insert("1.0", template.field_value)
            else:
                self.template_text.insert("1.0", self.get_default_template())
        except Exception as e:
            CTKMessageBox.show_error(self, '错误', f'加载模板失败：\n{str(e)}')
            self.template_text.insert("1.0", self.get_default_template())
    
    def save(self):
        """保存模板"""
        try:
            template = self.db.query(SalaryEmail).filter(SalaryEmail.field_name=='email_template').first()
            if not template:
                template = SalaryEmail()
                template.field_name = 'email_template'
                template.memo = "邮件模板"
            
            template.field_value = self.template_text.get("1.0", "end-1c")
            self.db.add(template)
            self.db.commit()
            
            CTKMessageBox.show_info(self, '成功', '模板保存成功！')
            self.destroy()
        except Exception as e:
            CTKMessageBox.show_error(self, '错误', f'保存失败：\n{str(e)}')
            self.db.rollback()
    
    def cancel(self):
        """取消编辑"""
        if CTKMessageBox.ask_yes_no(self, '确认', '确定要取消编辑吗？未保存的更改将丢失。') == "是":
            self.destroy()
    
    def preview_template(self):
        """在浏览器中预览模板"""
        try:
            template = self.template_text.get("1.0", "end-1c")
            
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
            Timer(5, cleanup).start()
            
        except Exception as e:
            CTKMessageBox.show_error(
                self,
                '预览失败',
                f'生成预览时出错：{str(e)}\n\n请检查模板中的变量是否正确。'
            )
    
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
        """重置为默认模板"""
        if CTKMessageBox.ask_yes_no(self, '确认', '确定要重置为默认模板吗？当前内容将丢失。') == "是":
            self.template_text.delete("1.0", "end")
            self.template_text.insert("1.0", self.get_default_template())


class CTKInfoManageWin(ctk.CTkToplevel):
    """信息管理窗口"""
    
    def __init__(self, parent):
        super().__init__(parent)
        
        self.title('信息管理')
        # 使用窗口大小管理器适配高分辨率
        CTKWindowSizeManager.adjust_window_size(self, 550, 650, 450, 500)
        self.resizable(True, True)
        
        # 设置窗口属性
        self.transient(parent)
        self.grab_set()
        
        # 确保窗口在父窗口中心显示
        self.after(100, lambda: responsive_theme_manager.center_window_on_parent(self, parent))
        
        self.parent = parent
        self.db = parent.db
        self.current_user = parent.current_user
        
        self.setup_ui()
        
        # 设置焦点
        self.focus_force()
    
    
    def setup_ui(self):
        """设置UI"""
        # 主容器
        main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 标题
        title_label = ctk.CTkLabel(
            main_frame,
            text="信息管理",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=20, weight="bold")
        )
        title_label.pack(pady=(0, 20))
        
        # 创建滚动框架
        self.scrollable_frame = ctk.CTkScrollableFrame(
            main_frame,
            corner_radius=10
        )
        self.scrollable_frame.pack(fill="both", expand=True, pady=(0, 20))
        
        # 公司信息区域
        company_frame = ctk.CTkFrame(self.scrollable_frame, corner_radius=10)
        company_frame.pack(fill="x", pady=(0, 15), padx=10)
        
        company_title = ctk.CTkLabel(
            company_frame,
            text="🏢 公司信息",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=16, weight="bold"),
            anchor="w"
        )
        company_title.pack(fill="x", padx=20, pady=(15, 10))
        
        # 公司名称
        company_content = ctk.CTkFrame(company_frame, fg_color="transparent")
        company_content.pack(fill="x", padx=20, pady=(0, 15))
        
        company_label = ctk.CTkLabel(
            company_content,
            text="公司名称：",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            anchor="w"
        )
        company_label.pack(fill="x", pady=(0, 5))
        
        self.company_entry = ctk.CTkEntry(
            company_content,
            placeholder_text="请输入公司名称",
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12)
        )
        self.company_entry.pack(fill="x")
        self.company_entry.insert(0, self.current_user.company_name or "")
        
        # 修改密码区域
        password_frame = ctk.CTkFrame(self.scrollable_frame, corner_radius=10)
        password_frame.pack(fill="x", pady=(0, 20), padx=10)
        
        password_title = ctk.CTkLabel(
            password_frame,
            text="🔒 修改密码",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=16, weight="bold"),
            anchor="w"
        )
        password_title.pack(fill="x", padx=20, pady=(15, 10))
        
        password_content = ctk.CTkFrame(password_frame, fg_color="transparent")
        password_content.pack(fill="x", padx=20, pady=(0, 15))
        
        # 原密码
        old_pwd_label = ctk.CTkLabel(
            password_content,
            text="原密码：",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            anchor="w"
        )
        old_pwd_label.pack(fill="x", pady=(0, 5))
        
        self.old_pwd_entry = ctk.CTkEntry(
            password_content,
            placeholder_text="请输入原密码",
            show="●",
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12)
        )
        self.old_pwd_entry.pack(fill="x", pady=(0, 10))
        
        # 新密码
        new_pwd_label = ctk.CTkLabel(
            password_content,
            text="新密码：",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            anchor="w"
        )
        new_pwd_label.pack(fill="x", pady=(0, 5))
        
        self.new_pwd_entry = ctk.CTkEntry(
            password_content,
            placeholder_text="请输入新密码",
            show="●",
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12)
        )
        self.new_pwd_entry.pack(fill="x", pady=(0, 10))
        
        # 确认密码
        confirm_pwd_label = ctk.CTkLabel(
            password_content,
            text="确认新密码：",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold"),
            anchor="w"
        )
        confirm_pwd_label.pack(fill="x", pady=(0, 5))
        
        self.confirm_pwd_entry = ctk.CTkEntry(
            password_content,
            placeholder_text="请再次输入新密码",
            show="●",
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12)
        )
        self.confirm_pwd_entry.pack(fill="x")
        
        # 按钮区域
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x", pady=10)
        
        # 取消按钮
        cancel_btn = ctk.CTkButton(
            button_frame,
            text="取消",
            command=self.destroy,
            width=100,
            height=40,
            fg_color="gray",
            hover_color="darkgray",
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold")
        )
        cancel_btn.pack(side="right", padx=5)
        
        # 保存按钮
        save_btn = ctk.CTkButton(
            button_frame,
            text="保存",
            command=self.save,
            width=100,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=14, weight="bold")
        )
        save_btn.pack(side="right", padx=5)
        
        # 绑定回车键
        self.bind('<Return>', lambda e: self.save())
    
    def save(self):
        """保存设置"""
        try:
            # 更新公司名称
            company_name = self.company_entry.get().strip()
            if company_name:
                self.current_user.company_name = company_name
            
            # 修改密码
            old_pwd = self.old_pwd_entry.get()
            new_pwd = self.new_pwd_entry.get()
            confirm_pwd = self.confirm_pwd_entry.get()
            
            if old_pwd or new_pwd or confirm_pwd:
                if not all([old_pwd, new_pwd, confirm_pwd]):
                    CTKMessageBox.show_warning(self, '提示', '请填写完整的密码信息！')
                    return
                
                if new_pwd != confirm_pwd:
                    CTKMessageBox.show_warning(self, '提示', '两次输入的新密码不一致！')
                    return
                
                # 验证原密码
                old_hash = hashlib.sha256(old_pwd.encode()).hexdigest()
                if old_hash != self.current_user.password:
                    CTKMessageBox.show_error(self, '错误', '原密码错误！')
                    return
                
                # 更新密码
                self.current_user.password = hashlib.sha256(new_pwd.encode()).hexdigest()
            
            self.db.commit()
            CTKMessageBox.show_info(self, '成功', '保存成功！')
            self.destroy()
            
        except Exception as e:
            CTKMessageBox.show_error(self, '错误', f'保存失败：\n{str(e)}')
            self.db.rollback()
