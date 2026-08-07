import os
import tempfile
import subprocess
import shutil

def test_real_scenario():
    print("=" * 60)
    print("模拟真实更新场景测试")
    print("=" * 60)
    
    # 模拟临时目录（英文路径）
    tmp_base = tempfile.mkdtemp(prefix='doujiao_test_')
    extract_dir = os.path.join(tmp_base, 'extracted')
    os.makedirs(extract_dir)
    
    # 模拟目标目录（假设用户安装在英文路径）
    app_dir = os.path.join(tmp_base, 'app')
    os.makedirs(app_dir)
    
    exe_name = '豆脚AlteraExcel工具.exe'
    
    # 源文件（解压后的）
    with open(os.path.join(extract_dir, exe_name), 'w', encoding='utf-8') as f:
        f.write('new version ' * 1000)
    os.makedirs(os.path.join(extract_dir, 'data'))
    with open(os.path.join(extract_dir, 'data', 'utils.js'), 'w', encoding='utf-8') as f:
        f.write('// new utils')
    
    # 目标文件（旧版本）
    with open(os.path.join(app_dir, exe_name), 'w', encoding='utf-8') as f:
        f.write('old version')
    os.makedirs(os.path.join(app_dir, 'data'))
    with open(os.path.join(app_dir, 'data', 'old.txt'), 'w', encoding='utf-8') as f:
        f.write('old data')
    
    print(f'\n源目录: {extract_dir}')
    print(f'目标目录: {app_dir}')
    print(f'EXE 文件名: {exe_name}')
    print(f'\n源 EXE 大小: {os.path.getsize(os.path.join(extract_dir, exe_name))}')
    print(f'目标 EXE 大小(复制前): {os.path.getsize(os.path.join(app_dir, exe_name))}')
    
    # 测试1: bat 文件用 GBK 编码
    print('\n--- Test 1: bat with GBK encoding ---')
    bat_path1 = os.path.join(tmp_base, 'update_gbk.bat')
    bat_content1 = f'''@echo off
echo Updating...
timeout /t 1 /nobreak >nul
xcopy /e /y /h /r "{extract_dir}\\*" "{app_dir}\\"
echo xcopy errorlevel: %errorlevel%
if exist "{app_dir}\\{exe_name}" (
    echo EXE FOUND
) else (
    echo EXE NOT FOUND
)
'''
    
    with open(bat_path1, 'w', encoding='gbk') as f:
        f.write(bat_content1)
    
    print(f'bat 文件: {bat_path1}')
    
    result1 = subprocess.run([bat_path1], shell=True, capture_output=True, text=True, encoding='gbk', errors='replace')
    print(f'stdout: {result1.stdout}')
    print(f'stderr: {result1.stderr}')
    print(f'returncode: {result1.returncode}')
    
    if os.path.exists(os.path.join(app_dir, exe_name)):
        print(f'目标 EXE 大小(复制后): {os.path.getsize(os.path.join(app_dir, exe_name))}')
        content = open(os.path.join(app_dir, exe_name), 'r', encoding='utf-8').read(20)
        print(f'EXE 内容(前20): {content}')
    
    # 重置目标目录
    shutil.rmtree(app_dir)
    os.makedirs(app_dir)
    with open(os.path.join(app_dir, exe_name), 'w', encoding='utf-8') as f:
        f.write('old version')
    os.makedirs(os.path.join(app_dir, 'data'))
    with open(os.path.join(app_dir, 'data', 'old.txt'), 'w', encoding='utf-8') as f:
        f.write('old data')
    
    # 测试2: bat 文件用 UTF-8 编码 + chcp 65001
    print('\n\n--- Test 2: bat with UTF-8 + chcp 65001 ---')
    bat_path2 = os.path.join(tmp_base, 'update_utf8.bat')
    bat_content2 = f'''@echo off
chcp 65001 >nul
echo Updating...
timeout /t 1 /nobreak >nul
xcopy /e /y /h /r "{extract_dir}\\*" "{app_dir}\\"
echo xcopy errorlevel: %errorlevel%
if exist "{app_dir}\\{exe_name}" (
    echo EXE FOUND
) else (
    echo EXE NOT FOUND
)
'''
    
    with open(bat_path2, 'w', encoding='utf-8') as f:
        f.write(bat_content2)
    
    print(f'bat 文件: {bat_path2}')
    
    result2 = subprocess.run([bat_path2], shell=True, capture_output=True, text=True, encoding='utf-8', errors='replace')
    print(f'stdout: {result2.stdout}')
    print(f'stderr: {result2.stderr}')
    print(f'returncode: {result2.returncode}')
    
    if os.path.exists(os.path.join(app_dir, exe_name)):
        print(f'目标 EXE 大小(复制后): {os.path.getsize(os.path.join(app_dir, exe_name))}')
        content = open(os.path.join(app_dir, exe_name), 'r', encoding='utf-8').read(20)
        print(f'EXE 内容(前20): {content}')
    
    print('\n\n' + '=' * 60)
    print('测试完成')
    print(f'临时目录: {tmp_base}')

if __name__ == '__main__':
    test_real_scenario()
