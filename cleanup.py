#!/usr/bin/env python3
# coding: utf-8
"""
项目清理脚本
清理临时文件、缓存文件和构建文件
"""
import os
import shutil
import glob
import sys

def cleanup_project():
    """清理项目目录"""
    print("🧹 开始清理项目目录...")
    
    # 清理Python缓存
    print("\n📂 清理Python缓存...")
    cache_count = 0
    for root, dirs, files in os.walk('.'):
        for dir_name in dirs:
            if dir_name == '__pycache__':
                cache_path = os.path.join(root, dir_name)
                print(f"  删除: {cache_path}")
                try:
                    shutil.rmtree(cache_path)
                    cache_count += 1
                except Exception as e:
                    print(f"  ⚠️  无法删除 {cache_path}: {e}")
    
    # 清理.pyc文件
    print("\n🐍 清理.pyc文件...")
    pyc_count = 0
    for pyc_file in glob.glob('**/*.pyc', recursive=True):
        print(f"  删除: {pyc_file}")
        try:
            os.remove(pyc_file)
            pyc_count += 1
        except Exception as e:
            print(f"  ⚠️  无法删除 {pyc_file}: {e}")
    
    # 清理构建文件
    print("\n🏗️  清理构建文件...")
    build_dirs = ['build', 'dist']
    build_count = 0
    for build_dir in build_dirs:
        if os.path.exists(build_dir):
            print(f"  删除构建目录: {build_dir}")
            try:
                shutil.rmtree(build_dir)
                build_count += 1
            except Exception as e:
                print(f"  ⚠️  无法删除 {build_dir}: {e}")
    
    # 清理spec文件
    print("\n📋 清理spec文件...")
    spec_count = 0
    for spec_file in glob.glob('*.spec'):
        print(f"  删除: {spec_file}")
        try:
            os.remove(spec_file)
            spec_count += 1
        except Exception as e:
            print(f"  ⚠️  无法删除 {spec_file}: {e}")
    
    # 清理临时文件
    print("\n🗂️  清理临时文件...")
    temp_dirs = ['temp', 'tmp']
    temp_count = 0
    for temp_dir in temp_dirs:
        if os.path.exists(temp_dir):
            print(f"  删除临时目录: {temp_dir}")
            try:
                shutil.rmtree(temp_dir)
                temp_count += 1
            except Exception as e:
                print(f"  ⚠️  无法删除 {temp_dir}: {e}")
    
    # 清理其他临时文件
    temp_patterns = ['*.tmp', '*.temp', '*.bak', '~*']
    file_count = 0
    for pattern in temp_patterns:
        for temp_file in glob.glob(pattern):
            print(f"  删除临时文件: {temp_file}")
            try:
                os.remove(temp_file)
                file_count += 1
            except Exception as e:
                print(f"  ⚠️  无法删除 {temp_file}: {e}")
    
    # 显示清理结果
    print("\n" + "="*50)
    print("✅ 清理完成！")
    print(f"📊 清理统计:")
    print(f"  - Python缓存目录: {cache_count}")
    print(f"  - .pyc文件: {pyc_count}")
    print(f"  - 构建目录: {build_count}")
    print(f"  - spec文件: {spec_count}")
    print(f"  - 临时目录: {temp_count}")
    print(f"  - 临时文件: {file_count}")
    print("="*50)

def check_virtual_envs():
    """检查虚拟环境"""
    print("\n🔍 检查虚拟环境...")
    
    venv_dirs = []
    for item in os.listdir('.'):
        if os.path.isdir(item) and item in ['venv', 'myenv', '.venv', 'env']:
            venv_dirs.append(item)
    
    if len(venv_dirs) > 1:
        print(f"⚠️  发现多个虚拟环境目录: {', '.join(venv_dirs)}")
        print("建议只保留一个虚拟环境目录以避免混淆。")
        
        for venv_dir in venv_dirs:
            size = get_dir_size(venv_dir)
            print(f"  - {venv_dir}: {format_size(size)}")
        
        print("\n💡 建议:")
        print("1. 选择一个主要使用的虚拟环境")
        print("2. 删除其他多余的虚拟环境目录")
        print("3. 重新安装依赖: pip install -r requirements.txt")
    elif len(venv_dirs) == 1:
        print(f"✅ 找到虚拟环境: {venv_dirs[0]}")
    else:
        print("❓ 未找到虚拟环境目录")

def get_dir_size(path):
    """获取目录大小"""
    total_size = 0
    try:
        for dirpath, dirnames, filenames in os.walk(path):
            for filename in filenames:
                file_path = os.path.join(dirpath, filename)
                try:
                    total_size += os.path.getsize(file_path)
                except (OSError, IOError):
                    pass
    except Exception:
        pass
    return total_size

def format_size(size):
    """格式化文件大小"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size < 1024.0:
            return f"{size:.1f} {unit}"
        size /= 1024.0
    return f"{size:.1f} TB"

def main():
    """主函数"""
    print("🧹 项目清理工具")
    print("=" * 50)
    
    # 确认当前目录
    current_dir = os.getcwd()
    print(f"📁 当前目录: {current_dir}")
    
    # 检查是否在项目根目录
    if not os.path.exists('salary_mail') or not os.path.exists('requirements.txt'):
        print("❌ 错误: 请在项目根目录下运行此脚本！")
        sys.exit(1)
    
    # 询问用户确认
    try:
        response = input("\n❓ 确定要清理项目目录吗？(y/N): ").strip().lower()
        if response not in ['y', 'yes', '是']:
            print("❌ 取消清理操作")
            return
    except KeyboardInterrupt:
        print("\n❌ 用户取消操作")
        return
    
    # 执行清理
    cleanup_project()
    
    # 检查虚拟环境
    check_virtual_envs()
    
    print("\n💡 清理完成后建议:")
    print("1. 重新测试程序是否正常运行")
    print("2. 如果删除了虚拟环境，请重新创建并安装依赖")
    print("3. 提交代码前检查.gitignore文件")

if __name__ == '__main__':
    main()
