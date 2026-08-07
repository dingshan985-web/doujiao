import os
import sys
import tempfile
import shutil

src_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, src_dir)

from altera_excel_generator import _run_update_worker

def test_update_worker():
    print("=" * 60)
    print("测试更新进程")
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
    
    with open(os.path.join(app_dir, exe_name), 'w', encoding='utf-8') as f:
        f.write('old version')
    os.makedirs(os.path.join(app_dir, 'data'))
    with open(os.path.join(app_dir, 'data', 'old.txt'), 'w', encoding='utf-8') as f:
        f.write('old data')
    
    print(f'\n更新前:')
    print(f'  exe 大小: {os.path.getsize(os.path.join(app_dir, exe_name))}')
    print(f'  data 目录: {os.listdir(os.path.join(app_dir, "data"))}')
    
    print(f'\n调用 _run_update_worker...')
    print(f'  extract_dir: {extract_dir}')
    print(f'  app_dir: {app_dir}')
    print(f'  tmp_dir: {tmp_dir}')
    
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
    
    print(f'\ntmp_dir 还在吗: {os.path.exists(tmp_dir)}')
    
    print('\n' + '=' * 60)
    print('测试完成')

if __name__ == '__main__':
    test_update_worker()
