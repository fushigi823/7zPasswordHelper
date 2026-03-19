import tkinter as tk
from tkinter import ttk, messagebox
import os
import subprocess

from password_store import PasswordStore
from extractor import ArchiveExtractor


class UnlockDialog:
    """解锁对话框 - 选择密码并解压"""

    def __init__(self, file_path):
        self.file_path = file_path
        self.password_store = PasswordStore()
        self.extractor = ArchiveExtractor()
        self.passwords = self.password_store.list()
        self.selected_alias = None
        self.selected_password = None
        self.result = False
        self.entered_password = None  # 记录手动输入的密码

        self.root = tk.Toplevel()  # 使用 Toplevel
        self.root.title(f"解锁压缩包 - {os.path.basename(file_path)}")
        self.root.geometry("450x400")
        self.root.resizable(False, False)

        # 居中显示
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (450 // 2)
        y = (self.root.winfo_screenheight() // 2) - (400 // 2)
        self.root.geometry(f"450x400+{x}+{y}")

        self._create_widgets()
        self._setup_drag_drop()

    def _create_widgets(self):
        """创建窗口组件"""
        # 文件信息框架
        info_frame = ttk.Frame(self.root, padding=10)
        info_frame.pack(fill=tk.X)

        ttk.Label(info_frame, text="文件:", font=('微软雅黑', 9, 'bold')).grid(row=0, column=0, sticky=tk.W)
        ttk.Label(info_frame, text=os.path.basename(self.file_path), wraplength=380).grid(row=0, column=1, sticky=tk.W, padx=5)

        ttk.Label(info_frame, text="路径:", font=('微软雅黑', 9, 'bold')).grid(row=1, column=0, sticky=tk.NW, pady=(10, 0))
        path_label = ttk.Label(info_frame, text=self.file_path, wraplength=380, foreground='gray')
        path_label.grid(row=1, column=1, sticky=tk.W, padx=5, pady=(10, 0))

        # 密码选择框架
        list_frame = ttk.LabelFrame(self.root, text="选择密码", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 创建Treeview显示密码列表
        columns = ('别名',)
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=8)

        self.tree.heading('别名', text='密码别名')
        self.tree.column('别名', width=400)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 填充密码列表
        for alias in self.passwords.keys():
            self.tree.insert('', tk.END, values=(alias,))

        # 如果有密码，默认选中第一个
        if self.passwords:
            self.tree.selection_set(self.tree.get_children()[0])
            # 手动触发选中事件，初始化selected_password
            self._on_select(None)

        # 绑定选择事件
        self.tree.bind('<<TreeviewSelect>>', self._on_select)

        # 试用所有密码选项
        self.try_all_var = tk.BooleanVar(value=False)
        try_all_check = ttk.Checkbutton(
            list_frame,
            text="试用所有密码 (自动尝试已保存的密码)",
            variable=self.try_all_var
        )
        try_all_check.pack(anchor=tk.W, pady=(10, 0))

        # 按钮框架
        btn_frame = ttk.Frame(self.root, padding=10)
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="解压", command=self._on_unlock, width=15).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="取消", command=self.root.destroy, width=15).pack(side=tk.RIGHT)

        # 手动输入密码按钮
        ttk.Button(btn_frame, text="手动输入密码...", command=self._on_manual_input, width=15).pack(side=tk.LEFT)

    def _setup_drag_drop(self):
        """设置拖拽支持（此窗口不用于拖拽）"""
        pass

    def _on_select(self, event):
        """选中密码时保存选择"""
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            self.selected_alias = str(item['values'][0])  # 转为字符串
            self.selected_password = self.passwords.get(self.selected_alias)

    def _on_unlock(self):
        """执行解压"""
        # 直接从Treeview获取当前选中项
        selection = self.tree.selection()
        current_alias = None
        if selection:
            item = self.tree.item(selection[0])
            current_alias = item['values'][0]

        print(f"DEBUG unlock: passwords={self.passwords}, selected_alias={current_alias}")

        if self.try_all_var.get():
            # 试用所有密码
            print(f"DEBUG try_all: trying {len(self.passwords)} passwords")
            for alias, pwd in self.passwords.items():
                print(f"  trying alias={repr(alias)}, pwd={repr(pwd)}")
            success, used_password, output_dir, error = self.extractor.try_passwords(self.file_path, self.passwords)
            if success:
                self.result = True
                messagebox.showinfo("成功", f"解压成功！\n使用的密码: {used_password}\n\n输出目录: {output_dir}")
                self._open_output_dir(output_dir)
                self.root.destroy()
            else:
                messagebox.showerror("失败", f"无法解压: {error}")
        else:
            # 使用选中密码
            if not current_alias:
                messagebox.showwarning("请选择", "请选择一个密码或勾选「试用所有密码」")
                return

            # 确保 current_alias 是字符串类型
            current_alias = str(current_alias)
            password = self.passwords.get(current_alias)
            print(f"DEBUG: alias={repr(current_alias)}, password={repr(password)}, file={self.file_path}")
            success, output_dir, error = self.extractor.extract_with_password(
                self.file_path, password
            )

            if success:
                self.result = True
                messagebox.showinfo("成功", f"解压成功！\n\n输出目录: {output_dir}")
                self._open_output_dir(output_dir)
                self.root.destroy()
            else:
                messagebox.showerror("失败", f"解压失败: {error}")

    def _on_manual_input(self):
        """手动输入密码对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("手动输入密码")
        dialog.geometry("300x120")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (300 // 2)
        y = (dialog.winfo_screenheight() // 2) - (120 // 2)
        dialog.geometry(f"300x120+{x}+{y}")

        ttk.Label(dialog, text="密码:").pack(pady=(20, 5))
        password_entry = ttk.Entry(dialog, show='*', width=30)
        password_entry.pack(pady=5)
        password_entry.focus()

        def do_unlock():
            password = password_entry.get()
            if not password:
                messagebox.showwarning("请输入", "请输入密码", parent=dialog)
                return

            success, output_dir, error = self.extractor.extract_with_password(self.file_path, password)
            if success:
                self.result = True
                self.entered_password = password  # 记录输入的密码
                dialog.destroy()
                messagebox.showinfo("成功", f"解压成功！\n\n输出目录: {output_dir}")
                self._open_output_dir(output_dir)
                self.root.destroy()
            else:
                messagebox.showerror("失败", f"解压失败: {error}")

        def on_enter(event):
            do_unlock()

        password_entry.bind('<Return>', on_enter)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确定", command=do_unlock, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=5)

    def _open_output_dir(self, path):
        """打开输出目录"""
        try:
            subprocess.run(['explorer', path], check=False)
        except Exception:
            pass

    def show(self):
        """显示对话框"""
        self.root.wait_window()
        return self.result