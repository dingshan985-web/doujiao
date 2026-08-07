import os
import sys
import tempfile
import zipfile
import shutil
import subprocess

def test_update_flow():
    print("=" * 60)
    print("测试更新流程")
    print("=" * 60)
    
    release_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "release")
    test_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_update_target")
    zip_path = os.path.join(release_dir, "update.zip")
    
    print(f"\n1. 检查 update.zip 是否存在: {zip_path}")
    if not os.path.exists(zip_path):
        print("   ❌ update.zip 不存在!")
        return False
    print(f"   ✅ 存在, 大小: {os.path.getsize(zip_path)} 字节")
    
    print(f"\n2. 准备测试目标目录: {test_dir}")
    if os.path.exists(test_dir):
        shutil.rmtree(test_dir)
    os.makedirs(test_dir)
    
    with open(os.path.join(test_dir, "豆脚AlteraExcel工具.exe"), "w") as f:
        f.write("old version")
    os.makedirs(os.path.join(test_dir, "data"), exist_ok=True)
    with open(os.path.join(test_dir, "data", "old_file.txt"), "w") as f:
        f.write("old data")
    print("   ✅ 测试目录准备完成")
    
    print("\n3. 解压 update.zip 到临时目录")
    tmp_dir = tempfile.mkdtemp(prefix="doujiao_test_update_")
    extract_dir = os.path.join(tmp_dir, "extracted")
    os.makedirs(extract_dir, exist_ok=True)
    
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(extract_dir)
    print(f"   ✅ 解压完成: {extract_dir}")
    
    exe_name = "豆脚AlteraExcel工具.exe"
    src_exe = os.path.join(extract_dir, exe_name)
    print(f"\n4. 检查 exe 是否在解压根目录: {src_exe}")
    if not os.path.exists(src_exe):
        print("   ❌ 根目录没有找到 exe，开始搜索...")
        found = False
        for root_dir, dirs, files in os.walk(extract_dir):
            if exe_name in files:
                src_exe = os.path.join(root_dir, exe_name)
                extract_dir = root_dir
                found = True
                print(f"   ✅ 在子目录找到: {src_exe}")
                print(f"   🔧 更新 extract_dir 为: {extract_dir}")
                break
        if not found:
            print("   ❌ 完全找不到 exe 文件!")
            return False
    else:
        print("   ✅ exe 在解压根目录")
    
    print(f"\n5. 列目录结构 (前20项):")
    count = 0
    for item in sorted(os.listdir(extract_dir)):
        item_path = os.path.join(extract_dir, item)
        if os.path.isdir(item_path):
            print(f"   📁 {item}/")
        else:
            print(f"   📄 {item}")
        count += 1
        if count >= 20:
            print(f"   ... 还有更多")
            break
    
    app_dir = test_dir
    sys_temp = tempfile.gettempdir()
    bat_path = os.path.join(sys_temp, "doujiao_test_update.bat")
    
    print(f"\n6. 生成 bat 更新脚本: {bat_path}")
    bat_content = f"""@echo off
echo Updating, please wait...
timeout /t 2 /nobreak >nul
echo Source: {extract_dir}
echo Target: {app_dir}
echo Running xcopy...
xcopy /e /y /h /r "{extract_dir}\\*" "{app_dir}\\"
echo xcopy exit code: %errorlevel%
if exist "{app_dir}\\{exe_name}" (
    echo Update complete!
    echo New exe size: 
    dir "{app_dir}\\{exe_name}" | find ".exe"
) else (
    echo Update FAILED - exe not found!
    pause
)
rmdir /s /q "{tmp_dir}"
echo Done. Press any key to exit...
pause >nul
del "%~f0"
"""
    
    with open(bat_path, "w", encoding="gbk") as f:
        f.write(bat_content)
    print("   ✅ bat 脚本生成完成")
    
    print(f"\n7. bat 脚本内容预览:")
    with open(bat_path, "r", encoding="gbk") as f:
        lines = f.readlines()
        for i, line in enumerate(lines[:20]):
            print(f"   {i+1:2d}: {line.rstrip()}")
    
    print(f"\n8. 启动 bat 脚本...")
    print("   (bat 脚本会暂停等待按键，测试完成后请关闭窗口)")
    
    result = subprocess.Popen(bat_path, shell=True, cwd=sys_temp)
    print(f"   ✅ bat 已启动, PID: {result.pid}")
    print("\n" + "=" * 60)
    print("测试启动完成！请查看弹出的 cmd 窗口检查 xcopy 是否成功")
    print("=" * 60)
    
    return True

if __name__ == "__main__":
    test_update_flow()
