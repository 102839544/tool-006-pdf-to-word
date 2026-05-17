#!/usr/bin/env python3
"""
PDF转Word工具 - 将PDF转换为可编辑的Word文档
"""
import sys, os, tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk

try:
    import pdf2docx
    HAS_PDF2DOCX = True
except ImportError:
    HAS_PDF2DOCX = False

class App:
    def __init__(self, root):
        self.root = root
        root.title("PDF转Word工具 v1.0")
        root.geometry("600x450")
        self.files = []
        self.build_ui()
    
    def build_ui(self):
        f = tk.Frame(self.root, bg="#c62828", height=50)
        f.pack(fill="x")
        tk.Label(f, text="📄 PDF → Word", font=("Arial",14,"bold"),
                 fg="white", bg="#c62828").pack(pady=12)
        
        main = tk.Frame(self.root, padx=15, pady=10)
        main.pack(fill="both", expand=True)
        
        bf = tk.Frame(main)
        bf.pack(fill="x", pady=5)
        tk.Button(bf, text="添加PDF文件", command=self.add_files,
                  bg="#c62828", fg="white", padx=12).pack(side="left", padx=5)
        tk.Button(bf, text="清空列表", command=self.clear,
                  bg="#d9534f", fg="white", padx=12).pack(side="left", padx=5)
        
        self.lb = tk.Listbox(main, font=("Consolas",10), bg="#ffebee", height=10)
        self.lb.pack(fill="both", expand=True, pady=10)
        
        tk.Button(main, text="开始转换", command=self.convert,
                  bg="#4caf50", fg="white", font=("Arial",11,"bold"),
                  padx=20, pady=8).pack(pady=10)
        
        self.status = tk.Label(main, text="请添加PDF文件进行转换",
                               font=("Arial",10), fg="gray")
        self.status.pack()
    
    def add_files(self):
        fs = filedialog.askopenfilenames(title="选择PDF文件",
             filetypes=[("PDF文件","*.pdf")])
        for f in fs:
            if f not in self.files:
                self.files.append(f)
                self.lb.insert("end", Path(f).name)
        self.status.config(text=f"已添加 {len(self.files)} 个文件")
    
    def clear(self):
        self.files.clear()
        self.lb.delete(0, "end")
        self.status.config(text="列表已清空")
    
    def convert(self):
        if not self.files:
            messagebox.showwarning("提示", "请先添加PDF文件")
            return
        if not HAS_PDF2DOCX:
            messagebox.showerror("缺少依赖", "请运行：pip install pdf2docx")
            return
        
        out_dir = filedialog.askdirectory(title="选择输出目录")
        if not out_dir: return
        
        ok = 0
        for f in self.files:
            try:
                out = str(Path(out_dir) / (Path(f).stem + ".docx"))
                cv = pdf2docx.Converter(f)
                cv.convert(out)
                cv.close()
                ok += 1
            except Exception as e:
                print(f"转换失败 {f}: {e}")
        
        messagebox.showinfo("完成", f"成功转换 {ok}/{len(self.files)} 个文件")
        self.status.config(text=f"✅ 完成：{ok} 个文件已转换")

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
