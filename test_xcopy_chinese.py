import os
import tempfile
import subprocess
import shutil

def test_xcopy_chinese():
    print("=" * 60)
    print("测试中文路径下的 xcopy 行为")
    print("=" * 60)
    
    base = r'd:\Users\Administrator\Desktop\丁山\claude\测试更新'
    src = os.path.join(base, '源文件')
    dst = os.path.join(base, '目标目录')
    
    if os.path.exists(base):
        shutil.rmtree(base)
    os.makedirs(src)
    os.makedirs(dst)
    
    exe_name = '测试程序.exe'
    
    with open(os.path.join(src, exe_name), 'w', encoding='utf-8') as f:
        f.write('new version ' * 1000)
    os.makedirs(os.path.join(src, '数据'))
    with open(os.path.join(src, '数据', '新文件.txt'), 'w', encoding='utf-8') as f:
        f.write('new data')
    
    with open(os.path.join(dst, exe_name), 'w', encoding='utf-8') as f:
        f.write('old version')
    os.makedirs(os.path.join(dst, '数据'))
    with open(os.path.join(dst, '数据', '旧文件.txt'), 'w', encoding='utf-8') as f:
        f.write('old data')
    
    print(f'\n源目录: {src}')
    print(f'目标目录: {dst}')
    print(f'\n源文件大小: {os.path.getsize(os.path.join(src, exe_name))}')
    print(f'目标文件大小(复制前): {os.path.getsize(os.path.join(dst, exe_name))}')
    print(f'目标数据目录(复制前): {os.listdir(os.path.join(dst, "数据"))}')
    
    # 测试1: 直接用 subprocess 调用
    print('\n--- Test 1: subprocess call ---')
    cmd = ['xcopy', '/e', '/y', '/h', '/r', src + '\\*', dst + '\\']
    print(f'cmd: {cmd}')
    result = subprocess.run(cmd, capture_output=True, text=True, encoding='gbk', errors='replace')
    print(f'stdout: {result.stdout}')
    print(f'stderr: {result.stderr}')
    print(f'returncode: {result.returncode}')
    
    print(f'\n目标文件大小: {os.path.getsize(os.path.join(dst, exe_name))}')
    print(f'目标数据目录: {os.listdir(os.path.join(dst, "数据"))}')
    print(f'文件内容(前20字符): {open(os.path.join(dst, exe_name), "r", encoding="utf-8").read(20)}')
    
    # 测试2: 通过 bat 文件
    print('\n\n--- Test 2: bat file (gbk encoding) ---')
    dst2 = os.path.join(base, '目标目录2')
    os.makedirs(dst2)
    with open(os.path.join(dst2, exe_name), 'w', encoding='utf-8') as f:
        f.write('old version2')
    os.makedirs(os.path.join(dst2, '数据'))
    with open(os.path.join(dst2, '数据', '旧文件.txt'), 'w', encoding='utf-8') as f:
        f.write('old data2')
    
    bat_path = os.path.join(base, '更新.bat')
    bat_content = f'''@echo off
echo Starting xcopy...
xcopy /e /y /h /r "{src}\\*" "{dst2}\\"
echo xcopy done, errorlevel: %errorlevel%
if exist "{dst2}\\{exe_name}" (
    echo EXE FOUND
) else (
    echo EXE NOT FOUND
)
dir "{dst2}" /b
'''
    
    with open(bat_path, 'w', encoding='gbk') as f:
        f.write(bat_content)
    
    print(f'bat 文件: {bat_path}')
    bat_content_read = open(bat_path, "r", encoding="gbk").read()
    print(f'bat 内容(gbk): {repr(bat_content_read)}')
    
    result2 = subprocess.run([bat_path], shell=True, capture_output=True, text=True, encoding='gbk', errors='replace')
    print(f'stdout: {result2.stdout}')
    print(f'stderr: {result2.stderr}')
    print(f'returncode: {result2.returncode}')
    
    print(f'\n目标文件大小: {os.path.getsize(os.path.join(dst2, exe_name))}')
    print(f'目标数据目录: {os.listdir(os.path.join(dst2, "数据"))}')
    print(f'文件内容(前20字符): {open(os.path.join(dst2, exe_name), "r", encoding="utf-8").read(20)}')
    
    # 测试3: bat 文件用 utf-8 编码 + chcp 65001
    print('\n\n--- Test 3: bat file (utf-8 with chcp 65001) ---')
    dst3 = os.path.join(base, '目标目录3')
    os.makedirs(dst3)
    with open(os.path.join(dst3, exe_name), 'w', encoding='utf-8') as f:
        f.write('old version3')
    os.makedirs(os.path.join(dst3, '数据'))
    with open(os.path.join(dst3, '数据', '旧文件.txt'), 'w', encoding='utf-8') as f:
        f.write('old data3')
    
    bat_path3 = os.path.join(base, '更新3.bat')
    bat_content3 = f'''@echo off
chcp 65001 >nul
echo Starting xcopy...
xcopy /e /y /h /r "{src}\\*" "{dst3}\\"
echo xcopy done, errorlevel: %errorlevel%
if exist "{dst3}\\{exe_name}" (
    echo EXE FOUND
) else (
    echo EXE NOT FOUND
)
'''
    
    with open(bat_path3, 'w', encoding='utf-8') as f:
        f.write(bat_content3)
    
    print(f'bat 文件: {bat_path3}')
    
    result3 = subprocess.run([bat_path3], shell=True, capture_output=True, text=True, encoding='utf-8', errors='replace')
    print(f'stdout: {result3.stdout}')
    print(f'stderr: {result3.stderr}')
    print(f'returncode: {result3.returncode}')
    
    print(f'\n目标文件大小: {os.path.getsize(os.path.join(dst3, exe_name))}')
    print(f'目标数据目录: {os.listdir(os.path.join(dst3, "数据"))}')
    
    print('\n\n' + '=' * 60)
    print('测试完成')

if __name__ == '__main__':
    test_xcopy_chinese()
