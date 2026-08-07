import os
import tempfile
import subprocess
import shutil

def test_xcopy():
    print("=" * 60)
    print("测试 xcopy 行为")
    print("=" * 60)
    
    base = tempfile.mkdtemp(prefix="xcopy_test_")
    src = os.path.join(base, "src")
    dst = os.path.join(base, "dst")
    
    os.makedirs(src)
    os.makedirs(dst)
    
    with open(os.path.join(src, "app.exe"), "w") as f:
        f.write("new version " * 1000)
    os.makedirs(os.path.join(src, "data"))
    with open(os.path.join(src, "data", "new.txt"), "w") as f:
        f.write("new data")
    
    with open(os.path.join(dst, "app.exe"), "w") as f:
        f.write("old version")
    os.makedirs(os.path.join(dst, "data"))
    with open(os.path.join(dst, "data", "old.txt"), "w") as f:
        f.write("old data")
    
    print(f"\n源目录: {src}")
    print(f"目标目录: {dst}")
    print(f"\n源文件大小: {os.path.getsize(os.path.join(src, 'app.exe'))}")
    print(f"目标文件大小(复制前): {os.path.getsize(os.path.join(dst, 'app.exe'))}")
    
    bat_path = os.path.join(base, "test.bat")
    bat_content = f"""@echo off
echo Testing xcopy...
echo Source: {src}
echo Dest: {dst}
echo.
echo Running: xcopy /e /y /h /r "{src}\\*" "{dst}\\"
xcopy /e /y /h /r "{src}\\*" "{dst}\\"
echo.
echo xcopy exit code: %errorlevel%
echo.
if exist "{dst}\\app.exe" (
    echo app.exe exists
) else (
    echo app.exe NOT FOUND
)
dir "{dst}" /b
echo.
echo data dir:
dir "{dst}\\data" /b
pause
"""
    
    with open(bat_path, "w", encoding="gbk") as f:
        f.write(bat_content)
    
    print(f"\n启动 bat 测试: {bat_path}")
    subprocess.Popen(bat_path, shell=True)
    print("\n请查看弹出的 cmd 窗口...")

if __name__ == "__main__":
    test_xcopy()
