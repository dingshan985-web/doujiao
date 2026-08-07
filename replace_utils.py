import os
import zipfile
import tempfile
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class ZipReplacerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("压缩包 utils.js 替换工具")
        self.root.geometry("520x320")
        self.root.resizable(True, False)

        self.zip_path = tk.StringVar()
        self.js_path = tk.StringVar()
        self.target_path = "assets/utils.js"

        self.create_widgets()

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(frame, text="压缩包 assets/utils.js 替换工具", font=("Microsoft YaHei", 14, "bold"))
        title.pack(pady=(0, 15))

        zip_frame = ttk.LabelFrame(frame, text="选择压缩包 (.zip)", padding=10)
        zip_frame.pack(fill=tk.X, pady=5)

        zip_row = ttk.Frame(zip_frame)
        zip_row.pack(fill=tk.X)

        ttk.Entry(zip_row, textvariable=self.zip_path).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(zip_row, text="浏览...", command=self.select_zip, width=10).pack(side=tk.RIGHT)

        js_frame = ttk.LabelFrame(frame, text="选择新的 utils.js 文件", padding=10)
        js_frame.pack(fill=tk.X, pady=5)

        js_row = ttk.Frame(js_frame)
        js_row.pack(fill=tk.X)

        ttk.Entry(js_row, textvariable=self.js_path).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(js_row, text="浏览...", command=self.select_js, width=10).pack(side=tk.RIGHT)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=20)

        ttk.Button(btn_frame, text="开始替换", command=self.replace, width=20).pack()

    def select_zip(self):
        path = filedialog.askopenfilename(
            title="选择压缩包",
            filetypes=[("ZIP 压缩包", "*.zip"), ("所有文件", "*.*")]
        )
        if path:
            self.zip_path.set(path)

    def select_js(self):
        path = filedialog.askopenfilename(
            title="选择 utils.js 文件",
            filetypes=[("JavaScript 文件", "*.js"), ("所有文件", "*.*")]
        )
        if path:
            self.js_path.set(path)

    def replace(self):
        zip_file = self.zip_path.get().strip()
        js_file = self.js_path.get().strip()

        if not zip_file:
            messagebox.showwarning("提示", "请选择压缩包文件")
            return
        if not os.path.exists(zip_file):
            messagebox.showerror("错误", "压缩包文件不存在")
            return
        if not js_file:
            messagebox.showwarning("提示", "请选择 utils.js 文件")
            return
        if not os.path.exists(js_file):
            messagebox.showerror("错误", "utils.js 文件不存在")
            return

        try:
            self._do_replace(zip_file, js_file)
            messagebox.showinfo("成功", f"替换完成！\n目标文件：{self.target_path}")
        except Exception as e:
            messagebox.showerror("错误", f"替换失败：{str(e)}")

    def _do_replace(self, zip_file, js_file):
        temp_dir = tempfile.mkdtemp()
        try:
            temp_zip = os.path.join(temp_dir, "temp.zip")

            with zipfile.ZipFile(zip_file, 'r') as zin:
                with zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED) as zout:
                    found = False
                    for item in zin.infolist():
                        if item.filename == self.target_path:
                            zout.write(js_file, self.target_path)
                            found = True
                        else:
                            data = zin.read(item.filename)
                            zout.writestr(item, data)

                    if not found:
                        raise RuntimeError(f"压缩包中未找到 {self.target_path}")

            shutil.move(temp_zip, zip_file)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        style.theme_use("clam")
    except Exception:
        pass
    ZipReplacerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
