import os
import shutil
import subprocess
import sys
import glob
from datetime import datetime

def build():
    # 设置环境变量，解决编码问题
    os.environ['PYTHONIOENCODING'] = 'utf-8'
    
    # 清理旧的构建文件
    for dir_name in ['build', 'dist', 'release']:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)

    try:
        # 执行打包命令
        result = subprocess.run(
            ['pyinstaller', 'build_config.spec', '--clean'],
            encoding='utf-8',
            check=True
        )
        
        if result.returncode != 0:
            print("打包失败！")
            return

        # 创建发布目录
        release_dir = os.path.join('release')
        os.makedirs(release_dir)

        # 复制可执行文件和资源
        dist_dir = 'dist'
        if not os.path.exists(dist_dir):
            print(f"错误：找不到目录 {dist_dir}")
            return
            
        # 查找实际的可执行文件
        exe_file = None
        for file in glob.glob(os.path.join(dist_dir, '*.exe')):
            exe_file = file
            break
            
        if not exe_file:
            print("错误：在dist目录中找不到可执行文件")
            return
            
        target_dir = os.path.join(release_dir, '工资条管理系统')
        os.makedirs(target_dir)
        
        # 复制所有文件到目标目录
        for item in os.listdir(dist_dir):
            s = os.path.join(dist_dir, item)
            d = os.path.join(target_dir, item)
            if os.path.isfile(s):
                shutil.copy2(s, d)
            else:
                shutil.copytree(s, d)
        
        # 重命名主程序
        new_exe_name = os.path.join(target_dir, '工资条管理系统.exe')
        if os.path.exists(new_exe_name):
            os.remove(new_exe_name)
        os.rename(
            os.path.join(target_dir, os.path.basename(exe_file)),
            new_exe_name
        )
        
        # 复制说明文档
        with open(os.path.join(release_dir, '说明.txt'), 'w', encoding='utf-8') as f:
            f.write(f'''工资条管理系统 v1.0.0

初始账号密码：
用户名：admin
密码：admin123

请在首次登录后及时修改密码。

注意事项：
1. 首次运行时会自动创建数据库
2. 数据库文件(salary.db)位于程序目录下
3. 请定期备份数据库文件

发布日期：{datetime.now().strftime('%Y-%m-%d')}
''')

        print('打包完成！发布文件在 release 目录下')
        
    except subprocess.CalledProcessError as e:
        print(f"打包过程出错：{e}")
    except Exception as e:
        print(f"发生错误：{e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    build() 