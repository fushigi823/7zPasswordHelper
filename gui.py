import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import os
import subprocess
import sys
from tkinterdnd2 import DND_FILES, TkinterDnD

from password_store import PasswordStore
from extractor import ArchiveExtractor
from unlock_dialog import UnlockDialog
from batch_dialog import BatchDialog


class MainWindow:
    """主窗口 - 密码管理界面"""

    def __init__(self):
        self.password_store = PasswordStore()
        self.extractor = ArchiveExtractor()

        self.root = TkinterDnD.Tk()  # 使用支持拖拽的Tk
        self.root.title("7z 密码管理器 v1.0 | 开发者: fushigi823")
        self.root.geometry("550x450")
        self.root.minsize(400, 350)

        # 居中显示
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (550 // 2)
        y = (self.root.winfo_screenheight() // 2) - (450 // 2)
        self.root.geometry(f"550x450+{x}+{y}")

        self._create_menu()
        self._create_widgets()
        self._setup_drag_drop()
        self._refresh_list()

    def _create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="打开压缩包...", command=self._open_archive, accelerator="Ctrl+O")
        file_menu.add_command(label="批量解压...", command=self._batch_open_archives, accelerator="Ctrl+B")
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit, accelerator="Ctrl+Q")

        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="设置文件关联...", command=self._setup_file_association)
        help_menu.add_separator()
        help_menu.add_command(label="关于", command=self._show_about)

        # 绑定快捷键
        self.root.bind('<Control-o>', lambda e: self._open_archive())
        self.root.bind('<Control-q>', lambda e: self.root.quit())

    def _create_widgets(self):
        """创建窗口组件"""
        # 顶部说明
        info_label = ttk.Label(
            self.root,
            text="管理您的压缩包密码，双击或拖拽压缩包到窗口即可快速解压",
            wraplength=500,
            foreground='gray'
        )
        info_label.pack(pady=(15, 5))

        # 密码列表框架
        list_frame = ttk.LabelFrame(self.root, text="已保存的密码", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        # 创建Treeview显示密码列表
        columns = ('别名', '密码')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings')

        self.tree.heading('别名', text='密码别名')
        self.tree.heading('密码', text='密码')
        self.tree.column('别名', width=200)
        self.tree.column('密码', width=200)

        # 添加滚动条
        v_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')

        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        # 绑定双击事件
        self.tree.bind('<Double-1>', self._on_item_double_click)

        # 按钮框架
        btn_frame = ttk.Frame(self.root, padding=10)
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="添加", command=self._add_password, width=12).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="编辑", command=self._edit_password, width=12).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="删除", command=self._delete_password, width=12).pack(side=tk.LEFT, padx=3)

        ttk.Button(btn_frame, text="批量解压", command=self._batch_open_archives, width=12).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="打开压缩包", command=self._open_archive, width=12).pack(side=tk.RIGHT, padx=3)

    def _setup_drag_drop(self):
        """设置拖拽支持"""
        # 让窗口接受拖拽
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self._on_drop)

    def _on_drop(self, event):
        """处理拖拽的文件"""
        # event.data 包含拖拽的文件路径
        files = self.root.tk.splitlist(event.data)
        file_list = []
        for f in files:
            f = f.strip('{').strip('}')  # 清理路径
            if self.extractor.is_supported(f):
                file_list.append(f)

        if not file_list:
            return

        if len(file_list) == 1:
            # 单文件，弹出解锁对话框
            self._unlock_archive(file_list[0])
        else:
            # 多文件，批量解压
            if not self.password_store.list():
                messagebox.showwarning("无密码", "请先添加至少一个密码")
                return
            dialog = BatchDialog(file_list)
            dialog.show()
            # 刷新密码列表
            self._refresh_list()

    def _refresh_list(self):
        """刷新密码列表"""
        # 清空现有项
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 添加密码
        passwords = self.password_store.list()
        for alias, password in passwords.items():
            # 隐藏密码，只显示****
            hidden_password = '*' * min(len(password), 8)
            self.tree.insert('', tk.END, values=(alias, hidden_password))

    def _on_item_double_click(self, event):
        """双击项目时编辑"""
        self._edit_password()

    def _add_password(self):
        """添加密码对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("添加密码")
        dialog.geometry("350x180")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (350 // 2)
        y = (dialog.winfo_screenheight() // 2) - (180 // 2)
        dialog.geometry(f"350x180+{x}+{y}")

        ttk.Label(dialog, text="密码别名:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=(20, 5))
        alias_entry = ttk.Entry(dialog, width=30)
        alias_entry.grid(row=0, column=1, pady=(20, 5))
        alias_entry.focus()

        ttk.Label(dialog, text="密码:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=10)
        password_entry = ttk.Entry(dialog, show='*', width=30)
        password_entry.grid(row=1, column=1, pady=10)

        def do_save():
            alias = alias_entry.get().strip()
            password = password_entry.get()
            if not alias:
                messagebox.showwarning("请输入", "请输入密码别名", parent=dialog)
                return
            if not password:
                messagebox.showwarning("请输入", "请输入密码", parent=dialog)
                return

            self.password_store.add(alias, password)
            self._refresh_list()
            dialog.destroy()

        def on_enter(event):
            do_save()

        alias_entry.bind('<Return>', lambda e: password_entry.focus())
        password_entry.bind('<Return>', on_enter)

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=15)
        ttk.Button(btn_frame, text="保存", command=do_save, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=5)

    def _edit_password(self):
        """编辑密码对话框"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("请选择", "请先选择一个密码进行编辑")
            return

        item = self.tree.item(selection[0])
        old_alias = item['values'][0]
        old_password = self.password_store.get(old_alias)

        dialog = tk.Toplevel(self.root)
        dialog.title("编辑密码")
        dialog.geometry("350x180")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (350 // 2)
        y = (dialog.winfo_screenheight() // 2) - (180 // 2)
        dialog.geometry(f"350x180+{x}+{y}")

        ttk.Label(dialog, text="密码别名:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=(20, 5))
        alias_entry = ttk.Entry(dialog, width=30)
        alias_entry.grid(row=0, column=1, pady=(20, 5))
        alias_entry.insert(0, old_alias)
        alias_entry.focus()

        ttk.Label(dialog, text="密码:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=10)
        password_entry = ttk.Entry(dialog, show='*', width=30)
        password_entry.grid(row=1, column=1, pady=10)
        password_entry.insert(0, old_password)

        def do_save():
            new_alias = alias_entry.get().strip()
            new_password = password_entry.get()
            if not new_alias:
                messagebox.showwarning("请输入", "请输入密码别名", parent=dialog)
                return
            if not new_password:
                messagebox.showwarning("请输入", "请输入密码", parent=dialog)
                return

            # 删除旧记录，添加新记录
            self.password_store.delete(old_alias)
            self.password_store.add(new_alias, new_password)
            self._refresh_list()
            dialog.destroy()

        def on_enter(event):
            do_save()

        password_entry.bind('<Return>', on_enter)

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=15)
        ttk.Button(btn_frame, text="保存", command=do_save, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=5)

    def _delete_password(self):
        """删除密码"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("请选择", "请先选择一个密码进行删除")
            return

        item = self.tree.item(selection[0])
        alias = str(item['values'][0])  # 确保是字符串

        if messagebox.askyesno("确认删除", f"确定要删除密码「{alias}」吗？"):
            self.password_store.delete(alias)
            # 清除选择
            self.tree.selection_remove(selection)
            # 刷新列表
            self._refresh_list()
            # 强制更新UI
            self.root.update()

    def _open_archive(self):
        """打开压缩包"""
        file_path = filedialog.askopenfilename(
            title="选择压缩包",
            filetypes=[
                ("压缩文件", "*.7z;*.zip;*.rar;*.tar;*.tar.gz;*.tgz;*.tar.bz2;*.tbz2"),
                ("7z 文件", "*.7z"),
                ("ZIP 文件", "*.zip"),
                ("所有文件", "*.*")
            ]
        )

        if file_path:
            self._unlock_archive(file_path)

    def _batch_open_archives(self):
        """批量打开压缩包"""
        file_paths = filedialog.askopenfilenames(
            title="选择多个压缩包",
            filetypes=[
                ("压缩文件", "*.7z;*.zip;*.rar;*.tar;*.tar.gz;*.tgz;*.tar.bz2;*.tbz2"),
                ("7z 文件", "*.7z"),
                ("ZIP 文件", "*.zip"),
                ("所有文件", "*.*")
            ]
        )

        if file_paths:
            # 转换为列表
            file_list = list(file_paths)
            if not self.password_store.list():
                messagebox.showwarning("无密码", "请先添加至少一个密码")
                return
            print("DEBUG: 准备创建批量对话框...")
            dialog = BatchDialog(file_list)
            print("DEBUG: 批量对话框已创建，等待关闭...")
            dialog.show()
            print(f"DEBUG: 批量对话框已关闭，准备刷新密码列表，当前密码: {self.password_store.list()}")
            # 批量处理完成后刷新密码列表
            self._refresh_list()
            self.root.update()  # 强制刷新UI
            print("DEBUG: 密码列表已刷新")

    def _unlock_archive(self, file_path):
        """解锁压缩包"""
        if not os.path.exists(file_path):
            messagebox.showerror("文件不存在", f"文件不存在:\n{file_path}")
            return

        if not self.extractor.is_supported(file_path):
            messagebox.showerror("不支持的格式", "不支持的压缩包格式")
            return

        passwords = self.password_store.list()

        # 如果没有保存的密码，直接弹出手动输入
        if not passwords:
            self._manual_input_and_extract(file_path)
            return

        # 先尝试所有已保存的密码
        success, used_password, output_dir, error = self.extractor.try_passwords(file_path, passwords)

        if success:
            messagebox.showinfo("成功", f"解压成功！\n使用的密码: {used_password}\n\n输出目录: {output_dir}")
            subprocess.run(['explorer', output_dir], check=False)
        else:
            # 所有密码都失败，直接弹出手动输入对话框
            self._manual_input_and_extract(file_path)

    def _manual_input_and_extract(self, file_path):
        """手动输入密码并尝试解压"""
        dialog = tk.Toplevel(self.root)
        dialog.title("输入密码")
        dialog.geometry("300x120")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中
        dialog.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (300 // 2)
        y = (self.root.winfo_screenheight() // 2) - (120 // 2)
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

            success, output_dir, error = self.extractor.extract_with_password(file_path, password)
            if success:
                dialog.destroy()
                # 询问是否保存密码
                if messagebox.askyesno("保存密码", "密码正确！是否保存这个密码？"):
                    alias = simpledialog.askstring("保存密码", "请输入密码别名:", parent=self.root)
                    if alias:
                        self.password_store.add(alias, password)
                        self._refresh_list()
                        messagebox.showinfo("保存成功", f"密码已保存为「{alias}」")
                messagebox.showinfo("成功", f"解压成功！\n\n输出目录: {output_dir}")
                subprocess.run(['explorer', output_dir], check=False)
            else:
                messagebox.showerror("失败", f"解压失败: {error}")

        def on_enter(event):
            do_unlock()

        password_entry.bind('<Return>', on_enter)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确定", command=do_unlock, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=5)

        dialog.wait_window()

    def _setup_file_association(self):
        """设置文件关联"""
        dialog = tk.Toplevel(self.root)
        dialog.title("设置文件关联")
        dialog.geometry("450x280")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (450 // 2)
        y = (dialog.winfo_screenheight() // 2) - (280 // 2)
        dialog.geometry(f"450x280+{x}+{y}")

        info_text = """此功能将把此程序设置为以下压缩包格式的默认打开方式：
- .7z, .zip, .rar, .tar, .tar.gz 等

设置后，双击这些格式的压缩包文件时会自动启动此程序。"""

        ttk.Label(dialog, text=info_text, wraplength=400, justify=tk.LEFT).pack(pady=15, padx=15)

        # 格式列表
        formats_frame = ttk.LabelFrame(dialog, text="将关联的文件格式", padding=10)
        formats_frame.pack(pady=10, padx=20, fill=tk.X)

        formats = ['.7z', '.zip', '.rar', '.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2']
        for fmt in formats:
            ttk.Label(formats_frame, text=f"  {fmt}").pack(anchor=tk.W)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)

        def do_associate():
            try:
                import winreg
                app_path = os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else 'main.py')

                # 注册文件关联的函数
                def associate_extension(ext, app_path):
                    try:
                        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, f"Software\\Classes\\{ext}\\OpenWithProgids")
                        winreg.SetValueEx(key, "7zPasswordManager.exe", 0, winreg.REG_SZ, b"")
                        winreg.CloseKey(key)

                        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, f"Software\\Classes\\7zPasswordManager.exe")
                        winreg.SetValue(key, "", winreg.REG_SZ, f"7z Password Manager")
                        winreg.SetValueEx(key, "URL Protocol", 0, winreg.REG_SZ, "")
                        winreg.CloseKey(key)

                        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, f"Software\\Classes\\7zPasswordManager.exe\\shell\\open\\command")
                        winreg.SetValue(key, "", winreg.REG_SZ, f'"{app_path}" "%1"')
                        winreg.CloseKey(key)
                    except Exception as e:
                        return str(e)
                    return None

                # 关联所有格式
                errors = []
                for ext in formats:
                    err = associate_extension(ext, app_path)
                    if err:
                        errors.append(f"{ext}: {err}")

                if errors:
                    messagebox.showerror("部分失败", "部分格式关联失败:\n" + "\n".join(errors), parent=dialog)
                else:
                    messagebox.showinfo("成功", "文件关联设置成功！\n\n请重启电脑或注销后生效。", parent=dialog)
                    dialog.destroy()

            except ImportError:
                messagebox.showerror("不支持", "此功能仅在Windows系统下可用", parent=dialog)
            except Exception as e:
                messagebox.showerror("错误", f"设置失败:\n{str(e)}", parent=dialog)

        def do_open_settings():
            try:
                subprocess.run(['start', 'ms-settings:defaultapps'], shell=True)
            except:
                messagebox.showinfo("提示", "请手动在Windows设置中更改默认应用", parent=dialog)

        ttk.Button(btn_frame, text="设置关联", command=do_associate, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="打开系统设置", command=do_open_settings, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="关闭", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=5)

    def _show_about(self):
        """显示关于对话框"""
        messagebox.showinfo(
            "关于",
            "7z 密码管理器 v1.0\n\n"
            "用于管理压缩包密码的小工具\n"
            "支持 7z/ZIP/RAR 等格式\n\n"
            "使用方法:\n"
            "1. 先在此窗口添加密码\n"
            "2. 双击压缩包即可选择密码解压"
        )

    def run(self):
        """运行主窗口"""
        self.root.mainloop()