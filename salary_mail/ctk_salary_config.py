# coding:utf-8
"""
工资项配置管理窗口
"""
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import openpyxl
import json
from typing import List, Dict, Any

from salary_mail.db_instance import SalaryFieldConfig
from salary_mail.ctk_components import CTKTreeview, CTKMessageBox, CTKFileDialog, CTKWindowSizeManager
from salary_mail.theme_config import theme_manager as responsive_theme_manager

class CTKSalaryConfigWin(ctk.CTkToplevel):
    """工资项配置管理窗口"""
    
    def __init__(self, parent):
        super().__init__(parent)
        
        self.title('工资项配置管理')
        
        # 设置响应式窗口配置 - 全屏显示
        self.responsive_config = responsive_theme_manager.setup_responsive_window(self, 'fullscreen_window')
        
        # 设置窗口属性
        self.transient(parent)
        self.grab_set()
        
        self.parent = parent
        self.db = parent.db
        
        # 工资项配置列表
        self.salary_fields = []
        
        self.setup_ui()
        self.load_salary_fields()
        
        # 延迟设置全屏显示
        self.after(100, self.setup_fullscreen)
        
        # 设置焦点
        self.focus_force()
        
        # 绑定键盘快捷键
        self.bind('<Escape>', lambda event: self.close_window())
        self.bind('<Control-q>', lambda event: self.close_window())
    
    def setup_fullscreen(self):
        """设置全屏显示"""
        try:
            # 获取屏幕尺寸
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            
            # 设置窗口大小为屏幕尺寸
            self.geometry(f"{screen_width}x{screen_height}+0+0")
            
            # 最大化窗口
            self.state('zoomed')
            
            # 设置全屏属性
            self.attributes('-fullscreen', True)
            
        except Exception as e:
            print(f"设置全屏显示失败: {e}")
            # 如果全屏失败，至少最大化窗口
            try:
                self.state('zoomed')
            except:
                pass
    
    def close_window(self):
        """关闭窗口"""
        try:
            # 退出全屏模式
            self.attributes('-fullscreen', False)
            # 关闭窗口
            self.destroy()
        except Exception as e:
            print(f"关闭窗口失败: {e}")
            # 强制关闭
            self.quit()
    
    def setup_ui(self):
        """设置UI"""
        # 主框架
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 标题栏框架
        title_frame = ctk.CTkFrame(main_frame)
        title_frame.pack(fill='x', pady=(0, 20))
        
        # 标题
        title_label = ctk.CTkLabel(title_frame, text="工资项配置管理", font=("Microsoft YaHei UI", 16, "bold"))
        title_label.pack(side='left', padx=10, pady=10)
        
        # 关闭按钮
        close_btn = ctk.CTkButton(
            title_frame,
            text="✕",
            command=self.close_window,
            width=40,
            height=32,
            fg_color="#FF4444",  # 红色
            hover_color="#CC3333"
        )
        close_btn.pack(side='right', padx=10, pady=10)
        
        # 按钮框架 - 使用工具栏样式
        button_frame = ctk.CTkFrame(main_frame, height=60, corner_radius=10)
        button_frame.pack(fill='x', pady=(0, 15))
        button_frame.pack_propagate(False)
        
        # 按钮内容框架
        button_content = ctk.CTkFrame(button_frame, fg_color="transparent")
        button_content.pack(fill='both', expand=True, padx=15, pady=10)
        
        # 导入模板按钮
        import_btn = ctk.CTkButton(
            button_content, 
            text="📋 导入模板", 
            command=self.import_template,
            width=120,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold")
        )
        import_btn.pack(side='left', padx=(0, 10))
        
        # 重置默认按钮
        reset_btn = ctk.CTkButton(
            button_content, 
            text="🔄 重置默认", 
            command=self.reset_to_default,
            width=120,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold"),
            fg_color="#FF9800",
            hover_color="#F57C00"
        )
        reset_btn.pack(side='left', padx=(0, 10))
        
        # 导出配置按钮
        export_btn = ctk.CTkButton(
            button_content, 
            text="💾 导出配置", 
            command=self.export_config,
            width=120,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold"),
            fg_color="#2196F3",
            hover_color="#1976D2"
        )
        export_btn.pack(side='left', padx=(0, 10))
        
        # 下载模板示例按钮
        download_btn = ctk.CTkButton(
            button_content, 
            text="⬇️ 下载模板示例", 
            command=self.download_template_example,
            width=140,
            height=40,
            font=ctk.CTkFont(family="Microsoft YaHei UI", size=12, weight="bold"),
            fg_color="#4CAF50",  # 绿色
            hover_color="#45a049"
        )
        download_btn.pack(side='left', padx=(0, 10))
        
        # 工资项配置表格
        self.setup_salary_fields_table(main_frame)
        
        # 状态栏
        status_frame = ctk.CTkFrame(main_frame)
        status_frame.pack(fill='x', pady=(10, 0))
        
        self.status_label = ctk.CTkLabel(status_frame, text="就绪")
        self.status_label.pack(side='left', padx=10, pady=5)
    
    def setup_salary_fields_table(self, parent):
        """设置工资项配置表格"""
        # 表格框架
        table_frame = ctk.CTkFrame(parent)
        table_frame.pack(fill='both', expand=True)
        
        # 表格标题
        table_title = ctk.CTkLabel(table_frame, text="工资项配置列表", font=("Microsoft YaHei UI", 12, "bold"))
        table_title.pack(pady=(10, 5))
        
        # 创建表格
        columns = [
            ('字段名称', 120),
            ('字段键名', 120),
            ('字段类型', 80),
            ('显示顺序', 80),
            ('显示宽度', 80),
            ('是否必填', 80),
            ('是否固定', 80)
        ]
        
        self.salary_tree = CTKTreeview(table_frame, columns=columns)
        
        # 设置列标题
        for col, width in columns:
            self.salary_tree.tree.heading(col, text=col)
            self.salary_tree.tree.column(col, width=width, minwidth=width)
        
        # 布局 - CTKTreeview 内部已包含滚动条
        self.salary_tree.pack(fill='both', expand=True, padx=10, pady=(0, 10))
    
    def load_salary_fields(self):
        """加载工资项配置"""
        try:
            # 清空表格
            for item in self.salary_tree.tree.get_children():
                self.salary_tree.tree.delete(item)
            
            # 查询数据库
            fields = self.db.query(SalaryFieldConfig).order_by(SalaryFieldConfig.display_order).all()
            
            # 显示数据
            for field in fields:
                field_type_map = {
                    'fixed': '固定',
                    'income': '收入',
                    'deduction': '扣除',
                    'summary': '汇总',
                    'other': '其他'
                }
                
                required_map = {0: '否', 1: '是'}
                fixed_map = {0: '否', 1: '是'}
                
                self.salary_tree.tree.insert('', 'end', values=(
                    field.field_name,
                    field.field_key,
                    field_type_map.get(field.field_type, field.field_type),
                    field.display_order,
                    field.display_width,
                    required_map.get(field.is_required, '否'),
                    fixed_map.get(field.is_fixed, '否')
                ))
            
            self.status_label.configure(text=f"已加载 {len(fields)} 个工资项配置")
            
        except Exception as e:
            CTKMessageBox.show_error(self, '错误', f'加载工资项配置失败：{str(e)}')
    
    def import_template(self):
        """导入Excel模板"""
        file_path = CTKFileDialog.open_file(
            self,
            title='选择工资模板文件',
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
        
        try:
            # 读取Excel文件
            wb = openpyxl.load_workbook(file_path)
            sheet = wb.active
            
            # 获取表头（第一行）
            headers = []
            for cell in sheet[1]:
                if cell.value:
                    headers.append(str(cell.value).strip())
            
            if len(headers) < 4:
                CTKMessageBox.show_error(self, '错误', '模板文件格式错误：至少需要4列（序号、员工编号、姓名、月份）')
                return
            
            # 验证固定字段
            required_fields = ['序号', '员工编号', '姓名', '月份']
            missing_fields = [f for f in required_fields if f not in headers]
            if missing_fields:
                CTKMessageBox.show_error(self, '错误', f'缺少必要的固定字段：{", ".join(missing_fields)}')
                return
            
            # 确认导入
            result = CTKMessageBox.ask_yes_no(
                self, 
                '确认导入', 
                f'将导入以下工资项配置：\n{", ".join(headers)}\n\n是否继续？'
            )
            
            if result == "是":
                self.process_template_headers(headers)
                
        except Exception as e:
            CTKMessageBox.show_error(self, '错误', f'导入模板失败：{str(e)}')
    
    def process_template_headers(self, headers: List[str]):
        """处理模板表头，更新工资项配置"""
        try:
            # 检查是否已有事务在进行中
            if hasattr(self.db, 'in_transaction') and self.db.in_transaction():
                # 如果已有事务，先提交或回滚
                try:
                    self.db.commit()
                except:
                    self.db.rollback()
            
            # 开始新事务
            self.db.begin()
            
            # 删除现有配置（保留固定字段）
            self.db.query(SalaryFieldConfig).filter(
                SalaryFieldConfig.is_fixed == 0
            ).delete()
            
            # 创建新的工资项配置
            for i, header in enumerate(headers):
                if header in ['序号', '员工编号', '姓名', '月份']:
                    # 跳过固定字段
                    continue
                
                # 判断字段类型
                field_type = 'other'
                if any(keyword in header for keyword in ['工资', '补贴', '费', '补发']):
                    field_type = 'income'
                elif any(keyword in header for keyword in ['扣款', '扣除', '保险', '公积金', '税', '会费']):
                    field_type = 'deduction'
                elif any(keyword in header for keyword in ['税前', '实发']):
                    field_type = 'summary'
                
                # 创建配置项
                field_config = SalaryFieldConfig(
                    field_name=header,
                    field_key=self.generate_field_key(header),
                    field_type=field_type,
                    display_order=i + 1,
                    display_width=85,
                    is_required=0,
                    is_fixed=0
                )
                
                self.db.add(field_config)
            
            # 提交事务
            self.db.commit()
            
            # 重新加载数据
            self.load_salary_fields()
            
            CTKMessageBox.show_info(self, '成功', '工资项配置导入成功！')
            
        except Exception as e:
            # 确保回滚事务
            try:
                self.db.rollback()
            except:
                pass  # 如果回滚也失败，继续处理
            CTKMessageBox.show_error(self, '错误', f'处理模板表头失败：{str(e)}')
    
    def generate_field_key(self, field_name: str) -> str:
        """生成字段键名"""
        # 简单的键名生成规则
        key = field_name.lower().replace(' ', '_').replace('（', '').replace('）', '')
        return key
    

    
    def reset_to_default(self):
        """重置为默认配置"""
        result = CTKMessageBox.ask_yes_no(
            self, 
            '确认重置', 
            '确定要重置为默认工资项配置吗？这将删除所有自定义配置。'
        )
        
        if result == "是":
            try:
                # 检查是否已有事务在进行中
                if hasattr(self.db, 'in_transaction') and self.db.in_transaction():
                    # 如果已有事务，先提交或回滚
                    try:
                        self.db.commit()
                    except:
                        self.db.rollback()
                
                # 开始新事务
                self.db.begin()
                
                # 删除所有配置
                self.db.query(SalaryFieldConfig).delete()
                
                # 重新初始化默认配置
                from salary_mail.db_instance import init_default_salary_fields
                init_default_salary_fields(self.db)
                
                # 提交事务
                self.db.commit()
                
                # 重新加载数据
                self.load_salary_fields()
                
                CTKMessageBox.show_info(self, '成功', '已重置为默认配置！')
                
            except Exception as e:
                # 确保回滚事务
                try:
                    self.db.rollback()
                except:
                    pass  # 如果回滚也失败，继续处理
                CTKMessageBox.show_error(self, '错误', f'重置配置失败：{str(e)}')
    
    def export_config(self):
        """导出工资项配置"""
        try:
            fields = self.db.query(SalaryFieldConfig).order_by(SalaryFieldConfig.display_order).all()
            
            # 准备导出数据
            export_data = []
            for field in fields:
                export_data.append({
                    'field_name': field.field_name,
                    'field_key': field.field_key,
                    'field_type': field.field_type,
                    'display_order': field.display_order,
                    'display_width': field.display_width,
                    'is_required': field.is_required,
                    'is_fixed': field.is_fixed
                })
            
            # 保存为JSON文件
            file_path = CTKFileDialog.save_file(
                self,
                title='保存配置',
                defaultextension='.json',
                filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
            )
            
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(export_data, f, ensure_ascii=False, indent=2)
                
                CTKMessageBox.show_info(self, '成功', f'配置已导出到：{file_path}')
                
        except Exception as e:
            CTKMessageBox.show_error(self, '错误', f'导出配置失败：{str(e)}')
    
    def download_template_example(self):
        """下载模板示例文件"""
        try:
            import os
            import shutil
            
            # 获取模板文件路径（现在在项目根目录下）
            template_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'template.xlsx')
            
            # 如果路径不存在，尝试使用当前工作目录
            if not os.path.exists(template_path):
                template_path = os.path.join(os.getcwd(), 'template.xlsx')
            
            if not os.path.exists(template_path):
                CTKMessageBox.show_error(self, '错误', '模板示例文件不存在！')
                return
            
            # 选择保存位置
            save_path = CTKFileDialog.save_file(
                self,
                title='保存模板示例',
                defaultextension='.xlsx',
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            
            if save_path:
                # 复制文件
                shutil.copy2(template_path, save_path)
                
                CTKMessageBox.show_info(
                    self, 
                    '下载成功', 
                    f'模板示例文件已保存到：\n{save_path}\n\n'
                    '该文件包含了常见的工资项目，可以作为配置参考。'
                )
                
        except Exception as e:
            CTKMessageBox.show_error(self, '错误', f'下载模板示例失败：{str(e)}')
