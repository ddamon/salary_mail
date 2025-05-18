# coding:utf-8
import os
from sqlalchemy import create_engine, Column, Integer, String, Float, ForeignKey, DateTime
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
    __tablename__ = 'salary_record'
    id = Column(Integer, primary_key=True)
    employee_id = Column(String(50), ForeignKey('employee.employee_id'))
    salary_month = Column(String(6))  # YYYYMM
    post_salary = Column(String(20))
    level_salary = Column(String(20))
    performance = Column(String(20))
    meal_allowance = Column(String(20))
    traffic_allowance = Column(String(20))
    cooling_allowance = Column(String(20))
    additional_payment = Column(String(20))
    leave_deduct = Column(String(20))
    sick_deduct = Column(String(20))
    other_deduct = Column(String(20))
    pre_tax_salary = Column(String(20))
    insurance = Column(String(20))
    house_fund = Column(String(20))
    tax = Column(String(20))
    union_fee = Column(String(20))
    actual_salary = Column(String(20))
    remark = Column(String(1024))  # 备注
    send_status = Column(Integer, default=0)  # 0:未发送 1:发送成功 2:发送失败

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

def set_db():
    """创建数据库连接"""
    engine = create_engine('sqlite:///salary.db', echo=False)
    Base.metadata.create_all(engine)
    Session = sessionmaker(bind=engine)
    db_session = Session()
    
    # 检查并初始化默认用户
    init_default_user(db_session)
    
    return db_session

if __name__ == '__main__':
    db = set_db()
    Base.metadata.create_all(db.engine)
