import os
import subprocess
import sys

def build_exe():
    """打包成exe檔案"""
    
    # PyInstaller 參數
    args = [
        'pyinstaller',
        '--onefile',
        '--windowed',
        '--name=窗簾計價系統',
        '--add-data=data;data',
        '--add-data=assets;assets',
        '--icon=assets/logo.ico',
        'main.py'
    ]
    
    # 如果有圖示檔案
    if os.path.exists('assets/logo.ico'):
        print("使用自訂圖示")
    else:
        print("未找到圖示檔案，使用預設圖示")
        args.remove('--icon=assets/logo.ico')
    
    # 執行打包
    try:
        subprocess.run(args, check=True)
        print("打包完成！")
        print("可執行檔案位於: dist/窗簾計價系統.exe")
    except subprocess.CalledProcessError as e:
        print(f"打包失敗: {e}")
        return False
    
    return True

def install_requirements():
    """安裝所需套件"""
    requirements = [
        'openpyxl',
        'xlwings',
        'pyinstaller'
    ]
    
    for package in requirements:
        try:
            subprocess.run([sys.executable, '-m', 'pip', 'install', package], check=True)
            print(f"已安裝 {package}")
        except subprocess.CalledProcessError as e:
            print(f"安裝 {package} 失敗: {e}")

if __name__ == "__main__":
    print("窗簾計價系統打包工具")
    print("=" * 30)
    
    choice = input("1. 安裝所需套件\n2. 執行打包\n3. 全部執行\n請選擇 (1-3): ")
    
    if choice in ['1', '3']:
        print("\n正在安裝所需套件...")
        install_requirements()
    
    if choice in ['2', '3']:
        print("\n正在打包程式...")
        if build_exe():
            print("\n打包成功！")
        else:
            print("\n打包失敗！")
