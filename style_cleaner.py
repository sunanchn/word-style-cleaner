"""GUI adapter：Tk 界面、线程调度与进度分档渲染；清理逻辑在 batch/cleaner。

批量清理跑在后台线程（#9），进度与结果经队列回主线程轮询渲染——
worker 不触碰任何 widget，状态栏/进度条/弹窗只发生在主线程。
"""
import os
import queue
import sys
import threading
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

# 进度分档阈值（#10）：文件数超过它时状态栏改报「已处理 i 个」（已完成数），不报当前序号
MANY_FILES_THRESHOLD = 10

# 主线程轮询批处理事件队列的间隔（毫秒）
_PROGRESS_POLL_MS = 50

class WordStyleCleaner:
    def __init__(self, root):
        self.root = root
        self.root.title("Word样式清理工具")
        self.root.geometry("700x350")  # 增加窗口高度以容纳样式列表区域

        # 创建UI组件
        self.create_widgets()

        # 批处理状态（#9）：后台线程把事件塞进队列，主线程轮询渲染
        self.batch_queue = queue.Queue()
        self.batch_thread = None
        self._batch_running = False
        self._batch_overwrite = False

    def create_widgets(self):
        # 创建选择文件或者文件夹的框架
        selection_frame = tk.Frame(self.root)
        selection_frame.pack(pady=10, fill=tk.X, padx=10)

        file_path_label = tk.Label(selection_frame, text='选择文件或文件夹：')
        file_path_label.pack(side=tk.LEFT)

        self.file_path_var = tk.StringVar()
        file_path_entry = tk.Entry(selection_frame, textvariable=self.file_path_var)
        file_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        self.choose_file_button = tk.Button(selection_frame, text='选择文件', command=self.choose_file)
        self.choose_file_button.pack(side=tk.LEFT, padx=2)

        self.choose_folder_button = tk.Button(selection_frame, text='选择文件夹', command=self.choose_folder)
        self.choose_folder_button.pack(side=tk.LEFT, padx=2)

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
        if self._batch_running:
            return  # 批次进行中：防御路径（如渲染异常恢复按钮后）也不允许重入

        target = self.file_path_var.get()
        if not target:
            messagebox.showwarning("警告", "请先选择文件或文件夹")
            return

        overwrite = self.overwrite_var.get()
        if overwrite and not self._confirm_overwrite(target):
            self.status_var.set("已取消覆盖模式清理，未做任何修改")
            return

        # 批次开跑：禁用按钮防重复点击与中途换目标，进度条归零
        self._set_processing_state(processing=True)
        self.status_var.set("正在处理...")
        self._stop_progress_bar()

        # 批量清理放后台线程（#9），进度与结果经队列回主线程轮询渲染
        self.batch_queue = queue.Queue()
        self._batch_running = True
        self._batch_overwrite = overwrite
        self.batch_thread = threading.Thread(
            target=self._batch_worker,
            args=(target, overwrite, self.batch_queue),
            daemon=True,
        )
        self.batch_thread.start()
        self.root.after(_PROGRESS_POLL_MS, self._poll_batch_events)

    def _batch_worker(self, target, overwrite, events):
        """后台线程体：跑批量清理，把进度与结果事件塞进队列。只报数据，不碰 widget。"""
        try:
            result = run_batch(
                target,
                on_progress=lambda index, total, path: events.put(('start', index, total, path)),
                on_file_done=lambda completed, total: events.put(('file_done', completed, total)),
                overwrite=overwrite,
            )
            events.put(('finished', result))
        except Exception as e:
            events.put(('failed', f'{type(e).__name__}: {e}'))

    def _poll_batch_events(self):
        """主线程轮询批处理事件队列：一切 UI 更新（含弹窗）只发生在这里。"""
        try:
            while True:
                self._handle_batch_event(self.batch_queue.get_nowait())
        except queue.Empty:
            pass
        except Exception as e:
            self._fail_batch(f'{type(e).__name__}: {e}')
            return
        if self._batch_running:
            self.root.after(_PROGRESS_POLL_MS, self._poll_batch_events)

    def _handle_batch_event(self, event):
        kind = event[0]
        if kind == 'start':
            self._apply_start_progress(event[1], event[2], event[3])
        elif kind == 'file_done':
            self._apply_done_progress(event[1], event[2])
        elif kind == 'finished':
            self._finish_batch(event[1])
        elif kind == 'failed':
            self._fail_batch(event[1])

    def _apply_start_progress(self, index, total, input_path):
        """进度分档渲染（#10）：1 个文件忙碌条；2 至阈值个报当前序号；超阈值只报完成数。"""
        file_name = os.path.basename(input_path)
        if total == 1:
            self.progress_bar.config(mode='indeterminate')
            self.progress_bar.start()
            self.status_var.set(f"正在处理: {file_name}...")
        elif total <= MANY_FILES_THRESHOLD:
            self.status_var.set(f"正在处理 ({index}/{total}): {file_name}")
            self.progress_var.set(index / total * 100)
        # 超过阈值：开始事件不更新，等完成事件报已完成数

    def _apply_done_progress(self, completed, total):
        if total > MANY_FILES_THRESHOLD:
            self.status_var.set(f"已处理 {completed} 个，共 {total} 个")
            self.progress_var.set(completed / total * 100)

    def _end_batch(self):
        """批次收尾（主线程）：停止轮询、恢复按钮、进度条归位。"""
        self._batch_running = False
        self._set_processing_state(processing=False)
        self._stop_progress_bar()

    def _finish_batch(self, result: BatchResult):
        """批次正常结束（主线程）：渲染结果与弹窗，恢复按钮与进度条。"""
        self._end_batch()

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
            done_msg = "样式清理完成（已覆盖原文件）！" if self._batch_overwrite else "样式清理完成！"
            messagebox.showinfo("完成", done_msg)
            self.status_var.set("处理完成")

    def _fail_batch(self, message: str):
        """批次异常（主线程）：错误弹窗 + 恢复按钮与进度条。"""
        self._end_batch()
        messagebox.showerror("错误", f"处理过程中发生错误：{message}")
        self.status_var.set(f"处理失败: {message}")

    def _set_processing_state(self, processing: bool):
        """处理期间禁用删除与选择按钮，结束后恢复。"""
        state = tk.DISABLED if processing else tk.NORMAL
        self.remove_button.config(state=state)
        self.choose_file_button.config(state=state)
        self.choose_folder_button.config(state=state)

    def _stop_progress_bar(self):
        """进度条归位：停忙碌动画、回确定模式、清零。"""
        self.progress_bar.stop()
        self.progress_bar.config(mode='determinate')
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
    """启动 GUI。"""
    root = tk.Tk()
    WordStyleCleaner(root)
    root.mainloop()


# GUI 版收到任何参数时的指路文案：不接受参数，命令行用法请用 CLI 版（ADR-0002）
ARGS_TO_CLI_HINT = '图形界面版不接受参数，命令行请使用 word-style-cleaner-cli.exe'


def _warn_args_not_supported():
    """弹窗指路 CLI 版。独立成 seam 是为了让测试能隔离，不真弹窗。"""
    root = tk.Tk()
    root.withdraw()
    try:
        messagebox.showwarning('不支持参数', ARGS_TO_CLI_HINT)
    finally:
        root.destroy()


def main(argv: list[str] | None = None) -> int:
    """GUI 版入口（ADR-0002）：无参数启动 GUI；带任何参数弹窗指路 CLI 版并以非零码退出。"""
    args_list = sys.argv[1:] if argv is None else argv
    if args_list:
        _warn_args_not_supported()
        return 2
    launch()
    return 0


if __name__ == "__main__":
    sys.exit(main())
