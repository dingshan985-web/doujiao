import os
import sys
import tempfile
import shutil
import time

def _run_update_worker(extract_dir, app_dir, tmp_dir):
    """独立更新进程：等待主程序退出后，复制文件并重启。"""
    exe_name = "豆脚AlteraExcel工具.exe"
    try:
        time.sleep(0.5)
        for item in os.listdir(extract_dir):
            src = os.path.join(extract_dir, item)
            dst = os.path.join(app_dir, item)
            if os.path.isdir(src):
                if os.path.exists(dst):
                    shutil.rmtree(dst, ignore_errors=True)
                shutil.copytree(src, dst)
            else:
                shutil.copy2(src, dst)
        time.sleep(0.2)
        new_exe = os.path.join(app_dir, exe_name)
        if os.path.exists(new_exe):
            print(f"更新成功！新 exe: {new_exe}")
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
    except Exception as e:
        print(f"更新失败: {e}")

def test():
    print("=" * 60)
    print("测试更新函数")
    print("=" * 60)
    
    base = tempfile.mkdtemp(prefix='update_test_')
    extract_dir = os.path.join(base, 'extract')
    app_dir = os.path.join(base, 'app')
    tmp_dir = os.path.join(base, 'tmp')
    
    os.makedirs(extract_dir)
    os.makedirs(app_dir)
    os.makedirs(tmp_dir)
    
    exe_name = '豆脚AlteraExcel工具.exe'
    
    with open(os.path.join(extract_dir, exe_name), 'w', encoding='utf-8') as f:
        f.write('new version ' * 1000)
    os.makedirs(os.path.join(extract_dir, 'data'))
    with open(os.path.join(extract_dir, 'data', 'new.txt'), 'w', encoding='utf-8') as f:
        f.write('new data')
    os.makedirs(os.path.join(extract_dir, 'data', 'sub'))
    with open(os.path.join(extract_dir, 'data', 'sub', 'sub.txt'), 'w', encoding='utf-8') as f:
        f.write('sub data')
    
    with open(os.path.join(app_dir, exe_name), 'w', encoding='utf-8') as f:
        f.write('old version')
    os.makedirs(os.path.join(app_dir, 'data'))
    with open(os.path.join(app_dir, 'data', 'old.txt'), 'w', encoding='utf-8') as f:
        f.write('old data')
    
    print(f'\n更新前:')
    print(f'  exe 大小: {os.path.getsize(os.path.join(app_dir, exe_name))}')
    print(f'  data 目录: {os.listdir(os.path.join(app_dir, "data"))}')
    
    print(f'\n执行更新...')
    _run_update_worker(extract_dir, app_dir, tmp_dir)
    
    print(f'\n更新后:')
    if os.path.exists(os.path.join(app_dir, exe_name)):
        print(f'  exe 存在: 是')
        print(f'  exe 大小: {os.path.getsize(os.path.join(app_dir, exe_name))}')
        content = open(os.path.join(app_dir, exe_name), 'r', encoding='utf-8').read(20)
        print(f'  exe 内容(前20): {content}')
    else:
        print(f'  exe 存在: 否')
    print(f'  data 目录: {os.listdir(os.path.join(app_dir, "data"))}')
    print(f'  data/sub 目录: {os.listdir(os.path.join(app_dir, "data", "sub"))}')
    
    print(f'\ntmp_dir 还在吗: {os.path.exists(tmp_dir)}')
    
    print('\n' + '=' * 60)
    print('测试完成')

if __name__ == '__main__':
    test()
