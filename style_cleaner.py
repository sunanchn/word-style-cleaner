import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext

from batch import BatchResult, discover_docx, run_batch
from cleaner import StyleCategory

# 清理结果的中性类别 → 中文展示名；展示词汇只存在于 GUI adapter
CATEGORY_LABELS = {
    StyleCategory.PARAGRAPH: '段落',
    StyleCategory.CHARACTER: '字符',
    StyleCategory.TABLE: '表格',
    StyleCategory.OTHER: '其他',
}

class WordStyleCleaner:
    def __init__(self, root):
        self.root = root
        self.root.title("Word样式清理工具")
        self.root.geometry("700x350")  # 增加窗口高度以容纳样式列表区域

        # 创建UI组件
        self.create_widgets()

    def create_widgets(self):
        # 创建选择文件或者文件夹的框架
        selection_frame = tk.Frame(self.root)
        selection_frame.pack(pady=10, fill=tk.X, padx=10)

        file_path_label = tk.Label(selection_frame, text='选择文件或文件夹：')
        file_path_label.pack(side=tk.LEFT)

        self.file_path_var = tk.StringVar()
        file_path_entry = tk.Entry(selection_frame, textvariable=self.file_path_var)
        file_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        choose_file_button = tk.Button(selection_frame, text='选择文件', command=self.choose_file)
        choose_file_button.pack(side=tk.LEFT, padx=2)

        choose_folder_button = tk.Button(selection_frame, text='选择文件夹', command=self.choose_folder)
        choose_folder_button.pack(side=tk.LEFT, padx=2)

        # 创建处理按钮和进度条
        action_frame = tk.Frame(self.root)
        action_frame.pack(pady=10, fill=tk.X, padx=10)

        self.remove_button = tk.Button(action_frame, text='删除未使用的样式', command=self.remove_unused_styles)
        self.remove_button.pack(side=tk.LEFT)

        # 覆盖模式开关：默认关闭（默认产 _Q 副本，原文件不动）
        self.overwrite_var = tk.BooleanVar(value=False)
        overwrite_check = tk.Checkbutton(action_frame, text='覆盖原文件', variable=self.overwrite_var)
        overwrite_check.pack(side=tk.LEFT, padx=5)

        # 创建进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(action_frame, variable=self.progress_var, length=300)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)

        # 创建统计信息框架
        stats_frame = tk.Frame(self.root)
        stats_frame.pack(pady=5, fill=tk.X, padx=10)

        self.stats_label = tk.Label(stats_frame, text="统计信息：共处理 0 个文件，删除 0 个样式")
        self.stats_label.pack()

        # 创建样式列表区域
        style_list_frame = tk.LabelFrame(self.root, text="处理详情")
        style_list_frame.pack(pady=5, fill=tk.BOTH, expand=True, padx=10)

        # 创建文本显示区域（移除标签页）
        self.result_text = scrolledtext.ScrolledText(style_list_frame, wrap=tk.WORD)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.result_text.config(state=tk.DISABLED)

        # 创建状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_label = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def _clear_results(self):
        """选择新目标后清空上一轮的处理详情。"""
        self._render_results(BatchResult())

    def choose_file(self):
        file_path = filedialog.askopenfilename(filetypes=[('Word 文档', '*.docx')])
        if file_path:
            self.file_path_var.set(file_path)
            self.status_var.set(f"已选择文件: {os.path.basename(file_path)}")
            self._clear_results()

    def choose_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.file_path_var.set(folder_path)
            # 待处理数量与 batch module 用同一套发现规则（排除自产 _Q 输出）
            docx_count = len(discover_docx(folder_path))
            self.status_var.set(f"已选择文件夹，包含 {docx_count} 个Word文档")
            self._clear_results()

    def remove_unused_styles(self):
        target = self.file_path_var.get()
        if not target:
            messagebox.showwarning("警告", "请先选择文件或文件夹")
            return

        overwrite = self.overwrite_var.get()
        if overwrite and not self._confirm_overwrite(target):
            self.status_var.set("已取消覆盖模式清理，未做任何修改")
            return

        # 禁用按钮防止重复点击
        self.remove_button.config(state=tk.DISABLED)
        self.status_var.set("正在处理...")
        self.root.update()

        try:
            result = run_batch(target, on_progress=self._on_batch_progress, overwrite=overwrite)

            if not result.results:
                messagebox.showinfo("提示", "所选文件夹中没有找到Word文档(.docx)")
                self.status_var.set("就绪")
                return

            self._render_results(result)

            failed = result.failed
            if failed:
                summary = "\n".join(
                    f"{os.path.basename(r.input_path)}：{r.error}" for r in failed
                )
                messagebox.showwarning(
                    "完成",
                    f"样式清理完成：成功 {len(result.succeeded)} 个，失败 {len(failed)} 个\n\n{summary}",
                )
                self.status_var.set("处理完成（有失败）")
            else:
                done_msg = "样式清理完成（已覆盖原文件）！" if overwrite else "样式清理完成！"
                messagebox.showinfo("完成", done_msg)
                self.status_var.set("处理完成")

        except Exception as e:
            messagebox.showerror("错误", f"处理过程中发生错误：{str(e)}")
            self.status_var.set(f"处理失败: {str(e)}")
        finally:
            # 恢复按钮状态
            self.remove_button.config(state=tk.NORMAL)
            self.progress_var.set(0)

    def _confirm_overwrite(self, target) -> bool:
        """覆盖模式执行前的确认弹窗，返回是否继续。"""
        if os.path.isdir(target):
            detail = f"目标文件夹：{target}\n包含 {len(discover_docx(target))} 个Word文档(.docx)"
        else:
            detail = f"目标文件：{target}"
        return messagebox.askyesno(
            "确认覆盖",
            f"已开启覆盖模式：清理结果将直接替换原文件，不再生成 _Q 副本，此操作无法撤销。\n\n{detail}\n\n确定继续吗？",
            default=messagebox.NO,
        )

    def _on_batch_progress(self, index, total, input_path):
        """batch module 的进度回调：刷新状态栏、进度条并重绘，避免白屏。"""
        file_name = os.path.basename(input_path)
        self.status_var.set(f"正在处理 ({index}/{total}): {file_name}")
        self.progress_var.set(index / total * 100)
        self.root.update()

    def _render_results(self, result: BatchResult):
        """渲染批处理汇总：逐文件删除样式明细 + 失败文件与原因。"""
        total_deleted = sum(len(r.deleted) for r in result.succeeded)
        summary = f"统计信息：共处理 {len(result.succeeded)} 个文件，删除 {total_deleted} 个样式"
        if result.failed:
            summary += f"，{len(result.failed)} 个文件失败"
        self.stats_label.config(text=summary)

        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)

        if not result.results:
            self.result_text.insert(tk.END, "暂无处理结果")
        else:
            for file_result in result.results:
                file_name = os.path.basename(file_result.input_path)
                if not file_result.ok:
                    self.result_text.insert(tk.END, f"❌ 文件: {file_name}\n")
                    self.result_text.insert(tk.END, f"   失败原因: {file_result.error}\n\n")
                elif file_result.deleted:
                    self.result_text.insert(tk.END, f"📄 文件: {file_name}\n")
                    self.result_text.insert(tk.END, f"   删除样式数量: {len(file_result.deleted)}\n")
                    self.result_text.insert(tk.END, "   删除的样式：\n")
                    self._render_deleted_styles(file_result.deleted)
                    self.result_text.insert(tk.END, "\n")
                else:
                    self.result_text.insert(tk.END, f"📄 文件: {file_name}\n")
                    self.result_text.insert(tk.END, "   未删除任何样式\n\n")

        self.result_text.config(state=tk.DISABLED)

    def _render_deleted_styles(self, deleted):
        """按类别分组展示一个文件里删除的样式。"""
        by_label = {}
        for d in deleted:
            by_label.setdefault(CATEGORY_LABELS[d.category], []).append(d.name)

        for label in CATEGORY_LABELS.values():
            if by_label.get(label):
                self.result_text.insert(tk.END, f"     {label}样式：\n")
                for name in sorted(by_label[label]):
                    self.result_text.insert(tk.END, f"       • {name}\n")

def launch():
    """启动 GUI。由 cli.main() 在无参数时调用；保持 __main__ 入口单点分发。"""
    root = tk.Tk()
    WordStyleCleaner(root)
    root.mainloop()


if __name__ == "__main__":
    import sys

    import cli

    sys.exit(cli.main())
