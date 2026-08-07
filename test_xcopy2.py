import os
import tempfile
import subprocess

base = tempfile.mkdtemp(prefix='xcopy_test_')
src = os.path.join(base, 'src')
dst = os.path.join(base, 'dst')

os.makedirs(src)
os.makedirs(dst)

with open(os.path.join(src, 'app.exe'), 'w') as f:
    f.write('new version ' * 1000)
os.makedirs(os.path.join(src, 'data'))
with open(os.path.join(src, 'data', 'new.txt'), 'w') as f:
    f.write('new data')

with open(os.path.join(dst, 'app.exe'), 'w') as f:
    f.write('old version')
os.makedirs(os.path.join(dst, 'data'))
with open(os.path.join(dst, 'data', 'old.txt'), 'w') as f:
    f.write('old data')

print('Before xcopy:')
print(f'  dst/app.exe size: {os.path.getsize(os.path.join(dst, "app.exe"))}')
print(f'  dst/data files: {os.listdir(os.path.join(dst, "data"))}')

# 测试 xcopy 命令 - 用列表参数
print('\n--- Test 1: xcopy with src\\* ---')
cmd = ['xcopy', '/e', '/y', '/h', '/r', src + '\\*', dst + '\\']
print(f'cmd: {cmd}')
result = subprocess.run(cmd, capture_output=True, text=True, encoding='gbk')
print(f'stdout: {result.stdout}')
print(f'stderr: {result.stderr}')
print(f'returncode: {result.returncode}')

print('\nAfter xcopy:')
print(f'  dst/app.exe size: {os.path.getsize(os.path.join(dst, "app.exe"))}')
print(f'  dst/data files: {os.listdir(os.path.join(dst, "data"))}')
print(f'  dst/app.exe content (first 20 chars): {open(os.path.join(dst, "app.exe")).read(20)}')

# 测试2: 不带 \* 的写法
print('\n\n--- Test 2: xcopy with src\\ (no *) ---')
dst2 = os.path.join(base, 'dst2')
os.makedirs(dst2)
with open(os.path.join(dst2, 'app.exe'), 'w') as f:
    f.write('old version2')
os.makedirs(os.path.join(dst2, 'data'))
with open(os.path.join(dst2, 'data', 'old.txt'), 'w') as f:
    f.write('old data2')

cmd2 = ['xcopy', src + '\\', dst2 + '\\', '/e', '/y', '/h', '/r']
print(f'cmd: {cmd2}')
result2 = subprocess.run(cmd2, capture_output=True, text=True, encoding='gbk')
print(f'stdout: {result2.stdout}')
print(f'stderr: {result2.stderr}')
print(f'returncode: {result2.returncode}')

print('\nAfter xcopy:')
print(f'  dst2/app.exe size: {os.path.getsize(os.path.join(dst2, "app.exe"))}')
print(f'  dst2/data files: {os.listdir(os.path.join(dst2, "data"))}')
print(f'  dst2/app.exe content (first 20 chars): {open(os.path.join(dst2, "app.exe")).read(20)}')

# 测试3: 用 robocopy
print('\n\n--- Test 3: robocopy ---')
dst3 = os.path.join(base, 'dst3')
os.makedirs(dst3)
with open(os.path.join(dst3, 'app.exe'), 'w') as f:
    f.write('old version3')
os.makedirs(os.path.join(dst3, 'data'))
with open(os.path.join(dst3, 'data', 'old.txt'), 'w') as f:
    f.write('old data3')

cmd3 = ['robocopy', src, dst3, '/E', '/IS', '/IT', '/R:3', '/W:1']
print(f'cmd: {cmd3}')
result3 = subprocess.run(cmd3, capture_output=True, text=True, encoding='gbk')
print(f'stdout: {result3.stdout}')
print(f'stderr: {result3.stderr}')
print(f'returncode: {result3.returncode} (0-7=success, 8+=fail)')

print('\nAfter robocopy:')
print(f'  dst3/app.exe size: {os.path.getsize(os.path.join(dst3, "app.exe"))}')
print(f'  dst3/data files: {os.listdir(os.path.join(dst3, "data"))}')
print(f'  dst3/app.exe content (first 20 chars): {open(os.path.join(dst3, "app.exe")).read(20)}')

print('\n\n' + '=' * 60)
print('测试完成')
