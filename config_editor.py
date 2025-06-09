import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import Label, Entry, Button, Checkbutton, messagebox, filedialog

CONFIG_FILE = "config.json"

def load_config():
    if not os.path.exists(CONFIG_FILE):
        return {
            "WATCH_FOLDER": os.path.expanduser("D:/定期练手/自动归纳/monitor"),
            "OUTPUT_BASE_FOLDER": os.path.expanduser("D:/定期练手/自动归纳/classified"),
            "DELAY_SECONDS": 5,
            "SUPPORTED_EXTS": [".docx", ".pdf", ".pptx", ".wps", ".mp4", ".wbd"],
            "AUTO_START": False
        }
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)
    messagebox.showinfo("保存成功", "配置已保存")

def select_folder(entry):
    path = filedialog.askdirectory()
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

def toggle_autostart(var):
    config = load_config()
    config["AUTO_START"] = var.get()
    save_config(config)

def run_config_editor():
    root = tk.Tk()
    root.title("AI文件分类器 - 配置界面")
    root.geometry("500x300")

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

    def save_and_exit():
        try:
            config["WATCH_FOLDER"] = watch_entry.get()
            config["OUTPUT_BASE_FOLDER"] = output_entry.get()
            config["DELAY_SECONDS"] = int(delay_entry.get())
            save_config(config)
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")
        finally:
            root.destroy()

    auto_var = tk.BooleanVar(value=config["AUTO_START"])
    Checkbutton(root, text="开机自启动", variable=auto_var, command=lambda: toggle_autostart(auto_var)).grid(
        row=3, column=1, sticky="w", padx=10, pady=5)

    Button(root, text="保存并退出", command=save_and_exit).grid(row=4, column=1, pady=20)

    root.mainloop()

if __name__ == "__main__":
    run_config_editor()
