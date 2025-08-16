# coding:utf-8
import os
import json
from sqlalchemy import create_engine, Column, Integer, String, Float, ForeignKey, DateTime, Text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
from datetime import datetime
import hashlib

Base = declarative_base()

class SalaryEmail(Base):
    __tablename__ = 'salary_email'
    id = Column(Integer, primary_key=True)
    field_name = Column(String(50))
    field_value = Column(String(500))
    memo = Column(String(200))

class Employee(Base):
    __tablename__ = 'employee'
    id = Column(Integer, primary_key=True)
    employee_id = Column(String(50), unique=True)
    name = Column(String(50))
    email = Column(String(100))
    phone = Column(String(20))
    status = Column(Integer, default=1)  # 1:在职 0:离职

class SalaryRecord(Base):
    """工资记录表 - 使用JSON存储工资项数据"""
    __tablename__ = 'salary_record'
    id = Column(Integer, primary_key=True)
    employee_id = Column(String(50), ForeignKey('employee.employee_id'))
    salary_month = Column(String(6))  # YYYYMM
    salary_data = Column(Text)  # JSON格式存储所有工资项数据
    remark = Column(String(1024))  # 备注
    send_status = Column(Integer, default=0)  # 0:未发送 1:发送成功 2:发送失败
    send_time = Column(String(20))  # 发送时间
    
    def get_salary_data(self):
        """获取工资数据字典"""
        if self.salary_data:
            try:
                return json.loads(self.salary_data)
            except json.JSONDecodeError:
                return {}
        return {}
    
    def set_salary_data(self, data_dict):
        """设置工资数据字典"""
        self.salary_data = json.dumps(data_dict, ensure_ascii=False)

class SalaryFieldConfig(Base):
    """工资项配置表"""
    __tablename__ = 'salary_field_config'
    
    id = Column(Integer, primary_key=True)
    field_name = Column(String(50))  # 显示名称
    field_key = Column(String(50))  # 字段键名
    field_type = Column(String(20))  # 字段类型: income(收入), deduction(扣除), other(其他)
    display_order = Column(Integer)  # 显示顺序
    display_width = Column(Integer, default=85)  # 显示宽度
    is_required = Column(Integer, default=0)  # 是否必填: 0否, 1是
    is_fixed = Column(Integer, default=0)  # 是否固定字段: 0否, 1是(如员工编号、姓名、月份等)
    created_at = Column(DateTime, default=datetime.now)

class SystemConfig(Base):
    """系统配置表"""
    __tablename__ = 'system_config'
    
    id = Column(Integer, primary_key=True)
    field_name = Column(String(50), unique=True)  # 字段名
    field_value = Column(String(200))  # 字段值
    memo = Column(String(200))  # 备注

class User(Base):
    """用户表"""
    __tablename__ = 'users'
    
    id = Column(Integer, primary_key=True)
    username = Column(String(50), unique=True)  # 用户名
    password = Column(String(200))  # 密码(加密存储)
    company_name = Column(String(100))  # 公司名称
    created_at = Column(DateTime, default=datetime.now)  # 创建时间
    last_login = Column(DateTime)  # 最后登录时间

def init_default_user(db_session):
    """初始化默认用户"""
    try:
        # 检查是否已存在用户
        user_count = db_session.query(User).count()
        if user_count == 0:
            # 创建默认用户
            default_password = "admin123"  # 默认密码
            password_hash = hashlib.sha256(default_password.encode()).hexdigest()
            
            default_user = User(
                username="admin",
                password=password_hash,
                company_name="默认公司",
                created_at=datetime.now()
            )
            
            db_session.add(default_user)
            db_session.commit()
            print("已创建默认用户 - 用户名: admin, 密码: admin123")
            return True
    except Exception as e:
        print(f"创建默认用户失败: {str(e)}")
        db_session.rollback()
    return False

def init_default_salary_fields(db_session):
    """初始化默认工资项配置"""
    try:
        # 检查是否已存在工资项配置
        field_count = db_session.query(SalaryFieldConfig).count()
        if field_count == 0:
            # 创建默认工资项配置
            default_fields = [
                # 固定字段
                SalaryFieldConfig(
                    field_name="序号",
                    field_key="serial_number",
                    field_type="fixed",
                    display_order=1,
                    display_width=50,
                    is_required=1,
                    is_fixed=1
                ),
                SalaryFieldConfig(
                    field_name="员工编号",
                    field_key="employee_id",
                    field_type="fixed",
                    display_order=2,
                    display_width=80,
                    is_required=1,
                    is_fixed=1
                ),
                SalaryFieldConfig(
                    field_name="姓名",
                    field_key="name",
                    field_type="fixed",
                    display_order=3,
                    display_width=70,
                    is_required=1,
                    is_fixed=1
                ),
                SalaryFieldConfig(
                    field_name="月份",
                    field_key="month",
                    field_type="fixed",
                    display_order=4,
                    display_width=70,
                    is_required=1,
                    is_fixed=1
                ),
                # 收入项
                SalaryFieldConfig(
                    field_name="岗位工资",
                    field_key="post_salary",
                    field_type="income",
                    display_order=5,
                    display_width=85,
                    is_required=0,
                    is_fixed=0
                ),
                SalaryFieldConfig(
                    field_name="薪级工资",
                    field_key="level_salary",
                    field_type="income",
                    display_order=6,
                    display_width=85,
                    is_required=0,
                    is_fixed=0
                ),
                SalaryFieldConfig(
                    field_name="绩效工资",
                    field_key="performance",
                    field_type="income",
                    display_order=7,
                    display_width=85,
                    is_required=0,
                    is_fixed=0
                ),
                SalaryFieldConfig(
                    field_name="餐补",
                    field_key="meal_allowance",
                    field_type="income",
                    display_order=8,
                    display_width=70,
                    is_required=0,
                    is_fixed=0
                ),
                SalaryFieldConfig(
                    field_name="交补",
                    field_key="traffic_allowance",
                    field_type="income",
                    display_order=9,
                    display_width=70,
                    is_required=0,
                    is_fixed=0
                ),
                SalaryFieldConfig(
                    field_name="防暑降温费",
                    field_key="cooling_allowance",
                    field_type="income",
                    display_order=10,
                    display_width=90,
                    is_required=0,
                    is_fixed=0
                ),
                SalaryFieldConfig(
                    field_name="补发",
                    field_key="additional_payment",
                    field_type="income",
                    display_order=11,
                    display_width=70,
                    is_required=0,
                    is_fixed=0
                ),
                # 扣除项
                SalaryFieldConfig(
                    field_name="事假扣款",
                    field_key="leave_deduct",
                    field_type="deduction",
                    display_order=12,
                    display_width=85,
                    is_required=0,
                    is_fixed=0
                ),
                SalaryFieldConfig(
                    field_name="病假扣款",
                    field_key="sick_deduct",
                    field_type="deduction",
                    display_order=13,
                    display_width=85,
                    is_required=0,
                    is_fixed=0
                ),
                SalaryFieldConfig(
                    field_name="其他扣款",
                    field_key="other_deduct",
                    field_type="deduction",
                    display_order=14,
                    display_width=85,
                    is_required=0,
                    is_fixed=0
                ),
                # 汇总项
                SalaryFieldConfig(
                    field_name="税前工资",
                    field_key="pre_tax_salary",
                    field_type="summary",
                    display_order=15,
                    display_width=85,
                    is_required=0,
                    is_fixed=0
                ),
                SalaryFieldConfig(
                    field_name="社会保险",
                    field_key="insurance",
                    field_type="deduction",
                    display_order=16,
                    display_width=85,
                    is_required=0,
                    is_fixed=0
                ),
                SalaryFieldConfig(
                    field_name="公积金",
                    field_key="house_fund",
                    field_type="deduction",
                    display_order=17,
                    display_width=80,
                    is_required=0,
                    is_fixed=0
                ),
                SalaryFieldConfig(
                    field_name="个人所得税",
                    field_key="tax",
                    field_type="deduction",
                    display_order=18,
                    display_width=85,
                    is_required=0,
                    is_fixed=0
                ),
                SalaryFieldConfig(
                    field_name="代缴工会会费",
                    field_key="union_fee",
                    field_type="deduction",
                    display_order=19,
                    display_width=90,
                    is_required=0,
                    is_fixed=0
                ),
                SalaryFieldConfig(
                    field_name="实发工资",
                    field_key="actual_salary",
                    field_type="summary",
                    display_order=20,
                    display_width=85,
                    is_required=0,
                    is_fixed=0
                )
            ]
            
            for field in default_fields:
                db_session.add(field)
            
            db_session.commit()
            print("已创建默认工资项配置")
            return True
    except Exception as e:
        print(f"创建默认工资项配置失败: {str(e)}")
        db_session.rollback()
    return False

def set_db():
    """创建数据库连接"""
    engine = create_engine('sqlite:///salary.db', echo=False)
    Base.metadata.create_all(engine)
    Session = sessionmaker(bind=engine)
    db_session = Session()
    
    # 检查并初始化默认用户
    init_default_user(db_session)
    
    # 不自动初始化默认工资项配置，让用户手动配置
    # init_default_salary_fields(db_session)
    
    return db_session

if __name__ == '__main__':
    db = set_db()
    Base.metadata.create_all(db.engine)
