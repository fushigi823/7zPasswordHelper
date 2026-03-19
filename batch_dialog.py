import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import os
import subprocess
import threading

from password_store import PasswordStore
from extractor import ArchiveExtractor


class BatchDialog:
    """批量解压对话框"""

    def __init__(self, file_paths):
        self.file_paths = file_paths
        self.password_store = PasswordStore()
        self.extractor = ArchiveExtractor()
        self.passwords = self.password_store.list()
        self.results = []  # 存储解压结果
        self.cancelled = False
        self.window_exists = True

        self.root = tk.Toplevel()
        self.root.title(f"批量解压 ({len(file_paths)} 个文件)")
        self.root.geometry("550x450")
        self.root.minsize(450, 350)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # 居中显示
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (550 // 2)
        y = (self.root.winfo_screenheight() // 2) - (450 // 2)
        self.root.geometry(f"550x450+{x}+{y}")

        self._create_widgets()
        self._start_processing()

    def _create_widgets(self):
        """创建窗口组件"""
        # 文件列表框架
        list_frame = ttk.LabelFrame(self.root, text="待处理文件", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 5))

        # 创建Treeview显示文件列表
        columns = ('状态', '文件名', '结果')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=8)

        self.tree.heading('状态', text='状态')
        self.tree.heading('文件名', text='文件名')
        self.tree.heading('结果', text='结果')
        self.tree.column('状态', width=60, anchor='center')
        self.tree.column('文件名', width=200)
        self.tree.column('结果', width=250)

        # 添加滚动条
        v_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=v_scrollbar.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        # 添加文件到列表
        for file_path in self.file_paths:
            self.tree.insert('', tk.END, values=('⏳', os.path.basename(file_path), '等待处理...'))

        # 进度框架
        progress_frame = ttk.LabelFrame(self.root, text="进度", padding=10)
        progress_frame.pack(fill=tk.X, padx=10, pady=5)

        self.progress_label = ttk.Label(progress_frame, text="准备开始...")
        self.progress_label.pack(anchor=tk.W)

        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', length=400)
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))

        # 按钮框架
        btn_frame = ttk.Frame(self.root, padding=10)
        btn_frame.pack(fill=tk.X)

        self.retry_btn = ttk.Button(btn_frame, text="重试失败项", command=self._retry_failed, width=15, state='disabled')
        self.retry_btn.pack(side=tk.LEFT, padx=3)

        self.cancel_btn = ttk.Button(btn_frame, text="取消", command=self._on_cancel, width=15)
        self.cancel_btn.pack(side=tk.RIGHT, padx=5)

        self.close_btn = ttk.Button(btn_frame, text="关闭", command=self._on_close, width=15, state='disabled')
        self.close_btn.pack(side=tk.RIGHT, padx=5)

    def _start_processing(self):
        """开始后台处理"""
        thread = threading.Thread(target=self._process_files, daemon=True)
        thread.start()

    def _process_files(self):
        """后台处理文件"""
        if not self.window_exists:
            return

        total = len(self.file_paths)
        success_count = 0
        fail_count = 0

        for i, file_path in enumerate(self.file_paths):
            if not self.window_exists or self.cancelled:
                break

            # 更新进度
            self._safe_update_progress(i, total, file_path)

            # 处理当前文件
            result = self._process_single_file(file_path)
            self.results.append(result)

            if result['success']:
                success_count += 1
            else:
                fail_count += 1

            # 更新列表项状态
            self._safe_update_result(i, result)

        # 完成
        if self.window_exists:
            self._safe_complete(success_count, fail_count)

    def _process_single_file(self, file_path):
        """处理单个文件，返回结果"""
        result = {
            'file_path': file_path,
            'file_name': os.path.basename(file_path),
            'success': False,
            'password': None,
            'output_dir': None,
            'error': None
        }

        # 尝试所有密码
        for alias, password in self.passwords.items():
            if not self.window_exists or self.cancelled:
                break

            success, output_dir, error = self.extractor.extract_with_password(file_path, password)
            if success:
                result['success'] = True
                result['password'] = password
                result['output_dir'] = output_dir
                return result

        if not result['success'] and not self.cancelled:
            result['error'] = "所有密码都无法解压"

        return result

    def _retry_failed(self):
        """重试失败的项目"""
        # 获取所有失败的项目
        failed_results = [r for r in self.results if not r['success']]

        if not failed_results:
            messagebox.showinfo("提示", "没有失败的项目需要重试", parent=self.root)
            return

        for result in failed_results:
            if not self.window_exists:
                break

            file_path = result['file_path']
            file_name = result['file_name']

            # 弹出手动输入对话框
            password = self._show_manual_input_dialog(file_path)

            if password:
                success, output_dir, error = self.extractor.extract_with_password(file_path, password)
                if success:
                    # 更新结果
                    result['success'] = True
                    result['password'] = password
                    result['output_dir'] = output_dir
                    result['error'] = None

                    # 更新列表显示
                    idx = self.file_paths.index(file_path)
                    children = self.tree.get_children()
                    if idx < len(children):
                        self.tree.item(children[idx], values=('✓', file_name, f"成功 (手动输入)"))

                    # 询问是否保存密码
                    if messagebox.askyesno("保存密码", "密码正确！是否保存这个密码？", parent=self.root):
                        alias = simpledialog.askstring("保存密码", "请输入密码别名:", initialvalue=file_name, parent=self.root)
                        if alias:
                            self.password_store.add(alias, password)
                            self.passwords = self.password_store.list()
                            messagebox.showinfo("保存成功", f"密码已保存为「{alias}」", parent=self.root)

                    # 打开输出目录
                    subprocess.run(['explorer', output_dir], check=False)
                else:
                    messagebox.showerror("失败", f"密码错误，解压失败: {file_name}", parent=self.root)
            else:
                # 用户取消
                break

    def _show_manual_input_dialog(self, file_path):
        """显示手动输入密码对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("输入密码")
        dialog.geometry("320x180")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中
        dialog.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (320 // 2)
        y = (self.root.winfo_screenheight() // 2) - (180 // 2)
        dialog.geometry(f"320x180+{x}+{y}")

        ttk.Label(dialog, text=f"文件: {os.path.basename(file_path)}", wraplength=280).pack(pady=(15, 10))
        ttk.Label(dialog, text="密码:").pack(pady=(5, 2))
        password_entry = ttk.Entry(dialog, show='*', width=35)
        password_entry.pack(pady=2)
        password_entry.focus()

        result = [None]  # 使用列表存储结果

        def do_unlock():
            result[0] = password_entry.get()
            dialog.destroy()

        def on_enter(event):
            do_unlock()

        password_entry.bind('<Return>', on_enter)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="确定", command=do_unlock, width=10).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side=tk.LEFT, padx=10)

        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        dialog.wait_window()

        return result[0]

        return result[0]

    def _safe_update_progress(self, idx, total, file_path):
        """线程安全地更新进度"""
        try:
            if self.window_exists:
                self.root.after(0, lambda i=idx, fp=file_path, t=total: self._update_progress(i, t, fp))
        except Exception:
            pass

    def _safe_update_result(self, idx, result):
        """线程安全地更新结果"""
        try:
            if self.window_exists:
                self.root.after(0, lambda i=idx, r=result: self._update_result(i, r))
        except Exception:
            pass

    def _safe_complete(self, success_count, fail_count):
        """线程安全地完成"""
        try:
            if self.window_exists:
                self.root.after(0, lambda sc=success_count, fc=fail_count: self._on_complete(sc, fc))
        except Exception:
            pass

    def _update_progress(self, idx, total, file_path):
        """更新进度显示"""
        try:
            self.progress_bar['maximum'] = total
            self.progress_bar['value'] = idx + 1
            self.progress_label.config(text=f"正在处理: {os.path.basename(file_path)} ({idx + 1}/{total})")
        except Exception:
            pass

    def _update_result(self, idx, result):
        """更新结果列表"""
        try:
            children = self.tree.get_children()
            if idx < len(children):
                item = children[idx]
                if result['success']:
                    status = '✓'
                    result_text = f"成功 (密码: {result['password']})"
                else:
                    status = '✗'
                    result_text = result['error'] or '失败'
                self.tree.item(item, values=(status, result['file_name'], result_text))
        except Exception:
            pass

    def _on_cancel(self):
        """取消处理"""
        self.cancelled = True
        try:
            self.progress_label.config(text="已取消")
            self.cancel_btn.config(state='disabled')
            self.retry_btn.config(state='normal')
        except Exception:
            pass

    def _on_complete(self, success_count, fail_count):
        """处理完成"""
        try:
            self.cancel_btn.config(state='disabled')

            if self.cancelled:
                self.progress_label.config(text="已取消")
            else:
                self.progress_label.config(text=f"完成！成功: {success_count}, 失败: {fail_count}")

            # 设置最终进度
            self.progress_bar['value'] = len(self.file_paths)

            # 如果有失败的，启用重试按钮
            if fail_count > 0:
                self.retry_btn.config(state='normal')

            self.close_btn.config(state='normal')
        except Exception:
            pass

    def _on_close(self):
        """关闭窗口"""
        self.window_exists = False
        self.cancelled = True
        try:
            self.root.destroy()
        except Exception:
            pass

    def show(self):
        """显示对话框"""
        try:
            self.root.wait_window()
        except Exception:
            pass
        return self.results