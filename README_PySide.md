# 工资条发送系统 - PySide6版本

这是工资条发送系统的PySide6版本，提供了现代化的Qt界面体验。

## 版本对比

| 特性 | CustomTkinter版本 | PySide6版本 |
|------|------------------|-------------|
| UI框架 | CustomTkinter | PySide6 (Qt6) |
| 启动文件 | `main.py` | `main_pyside.py` |
| 界面风格 | 现代化Tkinter | 原生Qt界面 |
| 性能 | 良好 | 优秀 |
| 跨平台 | Windows主要 | Windows/Linux/macOS |
| 打包大小 | 较小 | 较大 |

## 环境要求

- Windows 7/8/10/11
- Python 3.7.9+
- Visual C++ Redistributable for Visual Studio 2015-2019

## 安装与运行

### 1. 激活虚拟环境
```powershell
# Windows
myenv\Scripts\activate

# Linux/macOS
source myenv/bin/activate
```

### 2. 安装PySide6依赖
```powershell
pip install PySide6
```

### 3. 运行PySide6版本
```powershell
python main_pyside.py
```

### 4. 测试安装
```powershell
python test_pyside.py
```

## PySide6版本特性

### 已实现功能
- ✅ 现代化Qt界面设计
- ✅ 用户登录系统
- ✅ 主页面导航
- ✅ 员工管理界面框架
- ✅ 高DPI支持
- ✅ 响应式布局

### 开发中功能
- 🔄 员工增删改查功能
- 🔄 工资管理界面
- 🔄 邮件设置界面
- 🔄 系统设置界面
- 🔄 Excel导入导出
- 🔄 邮件发送功能

## 文件结构

```
salary_email/
├── main_pyside.py                    # PySide6版本启动文件
├── test_pyside.py                    # PySide6测试脚本
├── salary_mail/
│   ├── pyside_login_window.py        # PySide6登录窗口
│   ├── pyside_home_page.py           # PySide6主页面
│   ├── pyside_employee_manage.py     # PySide6员工管理
│   └── assets/                       # 共享资源文件
└── requirements.txt                  # 包含PySide6依赖
```

## 开发指南

### 添加新的PySide6窗口

1. 创建新的Python文件，如`pyside_new_window.py`
2. 继承适当的PySide6基类（QDialog, QMainWindow等）
3. 在`pyside_home_page.py`中导入并集成

### 样式设置

PySide6使用Qt样式表(QSS)进行界面美化，类似于CSS：

```python
self.setStyleSheet("""
    QWidget {
        background-color: #f5f5f5;
        font-family: "Microsoft YaHei UI";
    }
    QPushButton {
        background-color: #4CAF50;
        color: white;
        border: none;
        border-radius: 5px;
        padding: 8px 16px;
    }
""")
```

### 信号与槽

PySide6使用信号与槽机制进行事件处理：

```python
# 连接按钮点击事件
button.clicked.connect(self.on_button_click)

# 连接文本改变事件
line_edit.textChanged.connect(self.on_text_changed)
```

## 打包发布

### 使用PyInstaller打包PySide6版本

```powershell
# 安装PyInstaller
pip install pyinstaller

# 打包PySide6版本
pyinstaller --onefile --windowed --icon=salary_mail/assets/icon.ico main_pyside.py
```

### 注意事项

1. PySide6打包后的文件较大（约150MB+）
2. 需要确保所有Qt插件正确包含
3. 可能需要手动添加缺失的DLL文件

## 性能对比

| 指标 | CustomTkinter | PySide6 |
|------|---------------|---------|
| 启动速度 | 快 | 中等 |
| 内存占用 | 低 | 中等 |
| 界面响应 | 良好 | 优秀 |
| 渲染质量 | 良好 | 优秀 |
| 动画效果 | 基础 | 丰富 |

## 故障排除

### 常见问题

1. **PySide6导入失败**
   ```
   解决方案：pip install PySide6
   ```

2. **缺少Qt平台插件**
   ```
   错误：This application failed because no Qt platform plugin could be initialized
   解决方案：重新安装PySide6或设置QT_QPA_PLATFORM_PLUGIN_PATH环境变量
   ```

3. **高DPI显示问题**
   ```
   解决方案：在main_pyside.py中已设置高DPI支持
   ```

4. **字体显示问题**
   ```
   解决方案：确保系统安装了Microsoft YaHei UI字体
   ```

## 未来规划

### 短期目标（1-2周）
- 完成员工管理的增删改查功能
- 实现工资管理界面
- 添加邮件设置功能

### 中期目标（1个月）
- 完整的Excel导入导出功能
- 邮件发送功能
- 系统设置界面

### 长期目标（2-3个月）
- 数据可视化图表
- 多语言支持
- 主题切换功能
- 插件系统

## 贡献指南

1. Fork项目
2. 创建特性分支
3. 提交更改
4. 推送到分支
5. 创建Pull Request

## 许可证

本项目采用MIT许可证 - 查看LICENSE文件了解详情。

---

**注意**: PySide6版本与CustomTkinter版本可以并存，您可以根据需要选择合适的版本使用。
