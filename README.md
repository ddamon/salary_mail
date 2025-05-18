# 工资条发送系统

一个基于 Python 和 Tkinter 开发的桌面工资条邮件批量发送系统，支持员工信息管理、工资数据导入、邮件模板自定义与批量发送，适用于中小企业财务或人事场景。

---

## 目录结构

- `salary_mail/`：主程序代码
  - `EmployeeManageWin.py`：员工管理窗口
  - `SalaryManageWin.py`：工资管理窗口
  - `home_page.py`：主界面
  - `login_window.py`：登录窗口
  - `setting_box.py`：邮箱、模板、信息管理
  - `db_instance.py`：数据库模型与初始化
  - `assets/`：图标、logo等资源
  - `utils/`：工具模块
- `static/`：工资条邮件模板示例
- `release/`：打包发布目录，含说明文档
- `build.py`：一键打包脚本
- `requirements.txt`：依赖列表
- `salary.db`：SQLite 数据库（首次运行自动生成）

---

## 功能特点

- 员工信息管理（增删改查、Excel 导入）
- 工资数据导入（Excel）
- 邮件模板自定义与 HTML 预览
- 批量发送工资条，支持多线程
- 发送状态跟踪
- 系统参数、邮箱、公司信息管理
- 支持密码修改、公司名称自定义
- 现代化美观 UI，支持主题切换

---

## 安装与运行

### 环境要求
- Windows 7/8/10/11
- Python 3.7.9（推荐）
- Visual C++ Redistributable for Visual Studio 2015-2019

### 安装步骤
1. 安装 Python 3.7.9
2. 创建虚拟环境并激活
   ```powershell
   python -m venv venv
   .\venv\Scripts\activate
   ```
3. 安装依赖
   ```powershell
   python -m pip install -r requirements.txt
   ```
4. 运行主程序
   ```powershell
   python SalaryEmailRun.py
   ```
5. 打包发布
   ```powershell
   python build.py
   ```
   打包后文件在 `release/` 目录

---

## 使用说明
- 首次登录：用户名 `admin`，密码 `admin123`，请及时修改密码
- 员工、工资数据均支持 Excel 导入
- 邮件模板可自定义，支持 HTML 预览
- 邮箱配置、公司信息、密码修改等在"系统设置"中完成

---

## 依赖说明
- 详见 `requirements.txt`，如遇 Pillow 版本冲突请升级到 10.x

---

## 注意事项
1. Windows 7 用户必须安装 VC++ 运行库
2. 数据库文件 `salary.db` 位于程序目录，请定期备份
3. 如遇到 DLL 缺失，请安装 Visual C++ Redistributable
4. 建议使用 Windows 10 及以上版本运行
5. 安装依赖慢可使用：
   ```
   pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple/
   ```

---

## 许可证

本项目遵循 MIT License
