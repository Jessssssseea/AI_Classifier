import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import Label, Entry, Button, Checkbutton

CONFIG_FILE = "config.json"

def load_config():
    if not os.path.exists(CONFIG_FILE):
        # 如果配置文件不存在，创建一个默认配置
        config = {
            "WATCH_FOLDER": os.path.expanduser("D:/定期练手/自动归纳/monitor"),
            "OUTPUT_BASE_FOLDER": os.path.expanduser("D:/定期练手/自动归纳/classified"),
            "DELAY_SECONDS": 5,
            "SUPPORTED_EXTS": [".docx", ".pdf", ".pptx", ".wps", ".mp4", ".wbd"],
            "AUTO_START": False,
            "TESSERACT_PATH": r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # 默认值
        }
        save_config(config)  # 保存默认配置
        return config
    else:
        # 如果配置文件存在，加载配置
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
        # 确保 TESSERACT_PATH 键存在
        if "TESSERACT_PATH" not in config:
            config["TESSERACT_PATH"] = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # 默认值
            save_config(config)  # 更新配置文件
        return config

def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)
    messagebox.showinfo("保存成功", "配置已保存")

def select_folder(entry):
    path = filedialog.askdirectory()
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

def select_tesseract_path(entry):
    path = filedialog.askopenfilename(title="选择 Tesseract.exe", filetypes=[("Executable files", "*.exe")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

def toggle_autostart(var):
    config = load_config()
    config["AUTO_START"] = var.get()
    save_config(config)

def save_and_exit(watch_entry, output_entry, delay_entry, tesseract_entry):
    try:
        config = load_config()
        config["WATCH_FOLDER"] = watch_entry.get()
        config["OUTPUT_BASE_FOLDER"] = output_entry.get()
        config["DELAY_SECONDS"] = int(delay_entry.get())
        config["TESSERACT_PATH"] = tesseract_entry.get()
        save_config(config)
    except Exception as e:
        messagebox.showerror("错误", f"保存失败: {e}")
    else:
        messagebox.showinfo("保存成功", "配置已保存")
    finally:
        root.destroy()

def run_config_editor():
    global root
    root = tk.Tk()
    root.title("AI课件分类器 - 配置界面")
    root.geometry("470x260")
    root.iconbitmap('icon.ico')  # 更改窗口图标

    config = load_config()

    Label(root, text="监控路径").grid(row=0, column=0, sticky="w", padx=10, pady=5)
    watch_entry = Entry(root, width=40)
    watch_entry.grid(row=0, column=1, padx=10, pady=5)
    watch_entry.insert(0, config["WATCH_FOLDER"])
    Button(root, text="选择", command=lambda: select_folder(watch_entry)).grid(row=0, column=2, padx=5, pady=5)

    Label(root, text="归类路径").grid(row=1, column=0, sticky="w", padx=10, pady=5)
    output_entry = Entry(root, width=40)
    output_entry.grid(row=1, column=1, padx=10, pady=5)
    output_entry.insert(0, config["OUTPUT_BASE_FOLDER"])
    Button(root, text="选择", command=lambda: select_folder(output_entry)).grid(row=1, column=2, padx=5, pady=5)

    Label(root, text="延迟时间（秒）").grid(row=2, column=0, sticky="w", padx=10, pady=5)
    delay_entry = Entry(root, width=10)
    delay_entry.grid(row=2, column=1, sticky="w", padx=10, pady=5)
    delay_entry.insert(0, str(config["DELAY_SECONDS"]))

    Label(root, text="Tesseract 路径").grid(row=3, column=0, sticky="w", padx=10, pady=5)
    tesseract_entry = Entry(root, width=40)
    tesseract_entry.grid(row=3, column=1, padx=10, pady=5)
    tesseract_entry.insert(0, config["TESSERACT_PATH"])
    Button(root, text="选择", command=lambda: select_tesseract_path(tesseract_entry)).grid(row=3, column=2, padx=5, pady=5)

    auto_var = tk.BooleanVar(value=config["AUTO_START"])
    Checkbutton(root, text="开机自启动", variable=auto_var, command=lambda: toggle_autostart(auto_var)).grid(
        row=4, column=1, sticky="w", padx=10, pady=5)

    Button(root, text="保存并退出", command=lambda: save_and_exit(watch_entry, output_entry, delay_entry, tesseract_entry)).grid(row=5, column=1, pady=20)

    root.mainloop()

if __name__ == "__main__":
    run_config_editor()