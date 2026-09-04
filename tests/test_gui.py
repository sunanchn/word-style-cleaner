"""GUI 的 seam 测试：入口分发用假 launch/弹窗隔离，批处理接线用假 widget 组装——绝不真开窗口。"""
import queue

import pytest

import style_cleaner
from batch import BatchResult, FileResult
from style_cleaner import WordStyleCleaner


@pytest.fixture
def launched(monkeypatch):
    called = []
    monkeypatch.setattr(style_cleaner, 'launch', lambda: called.append(True))
    return called


def test_no_args_launches_gui(launched):
    assert style_cleaner.main([]) == 0
    assert launched == [True]


@pytest.mark.parametrize('args', [['文档.docx'], ['--overwrite'], ['--help']])
def test_any_args_warn_and_exit_nonzero_without_launching(args, launched, monkeypatch):
    # GUI 版单一职责：收到任何参数都不进 GUI，弹窗指路 CLI 版并以非零码退出
    warned = []
    monkeypatch.setattr(
        style_cleaner, '_warn_args_not_supported', lambda: warned.append(True)
    )

    code = style_cleaner.main(args)

    assert code != 0
    assert warned == [True]
    assert launched == []


def test_args_hint_points_to_cli_exe():
    assert '不接受参数' in style_cleaner.ARGS_TO_CLI_HINT
    assert 'word-style-cleaner-cli.exe' in style_cleaner.ARGS_TO_CLI_HINT


# ---------------------------------------------------------------------------
# 批处理接线：后台线程 + 事件队列 + 主线程分档渲染（#9/#10）。
# 用假 widget 组装 WordStyleCleaner（不走 __init__，不创建 Tk）。


class FakeVar:
    """StringVar/BooleanVar/DoubleVar 替身。"""

    def __init__(self, value=''):
        self.value = value

    def set(self, value):
        self.value = value

    def get(self):
        return self.value


class FakeProgressbar:
    """Progressbar 替身：记录模式切换与忙碌动画开关。"""

    def __init__(self):
        self.mode = 'determinate'
        self.running = False

    def config(self, **kw):
        if 'mode' in kw:
            self.mode = kw['mode']

    def start(self, interval=None):
        self.running = True

    def stop(self):
        self.running = False


class FakeButton:
    def __init__(self):
        self.state = 'normal'

    def config(self, **kw):
        if 'state' in kw:
            self.state = kw['state']


class FakeRoot:
    """root 替身：只记录 after 调度，不跑事件循环。"""

    def __init__(self):
        self.after_calls = []

    def after(self, ms, func):
        self.after_calls.append((ms, func))


class FakeThread:
    """Thread 替身：记录创建参数，不真开线程。"""

    def __init__(self, target=None, args=(), daemon=None):
        self.target = target
        self.args = args
        self.daemon = daemon
        self.started = False

    def start(self):
        self.started = True


def make_app():
    """用假 widget 组装 WordStyleCleaner 实例。"""
    app = object.__new__(WordStyleCleaner)
    app.root = FakeRoot()
    app.file_path_var = FakeVar()
    app.overwrite_var = FakeVar(False)
    app.status_var = FakeVar('就绪')
    app.progress_var = FakeVar(0.0)
    app.progress_bar = FakeProgressbar()
    app.remove_button = FakeButton()
    app.choose_file_button = FakeButton()
    app.choose_folder_button = FakeButton()
    app.batch_queue = queue.Queue()
    app._batch_running = False
    app._batch_overwrite = False
    return app


# 进度分档（#10）：1 个文件忙碌条；2–10 个报当前序号；超过阈值报已完成数


def test_single_file_tier_uses_busy_bar_without_percentage():
    app = make_app()
    app._apply_start_progress(1, 1, r'C:\docs\报告.docx')
    assert app.progress_bar.mode == 'indeterminate'
    assert app.progress_bar.running is True
    assert app.progress_var.value == 0.0  # 全程不出现百分比
    assert '正在处理' in app.status_var.value
    assert '报告.docx' in app.status_var.value


def test_few_files_tier_reports_current_index():
    app = make_app()
    app._apply_start_progress(3, 5, r'C:\docs\c.docx')
    assert app.status_var.value == '正在处理 (3/5): c.docx'
    assert app.progress_var.value == 60.0


def test_threshold_ten_files_still_reports_current_index():
    total = style_cleaner.MANY_FILES_THRESHOLD  # 阈值本身仍属"少量"档
    app = make_app()
    app._apply_start_progress(total, total, r'C:\docs\j.docx')
    assert app.status_var.value == f'正在处理 ({total}/{total}): j.docx'


def test_many_files_tier_reports_completed_count_only():
    total = style_cleaner.MANY_FILES_THRESHOLD + 1  # 超过阈值进"多文件"档
    app = make_app()
    # 超过阈值：开始事件不报当前序号
    app._apply_start_progress(1, total, r'C:\docs\a.docx')
    assert '正在处理 (' not in app.status_var.value
    # 完成事件报已完成数，随完成递增
    app._apply_done_progress(4, total)
    assert app.status_var.value == f'已处理 4 个，共 {total} 个'
    assert app.progress_var.value == pytest.approx(4 / total * 100)
    app._apply_done_progress(total, total)
    assert app.status_var.value == f'已处理 {total} 个，共 {total} 个'


def test_done_events_ignored_in_few_files_tier():
    app = make_app()
    app._apply_start_progress(2, 5, r'C:\docs\b.docx')
    before = (app.status_var.value, app.progress_var.value)
    app._apply_done_progress(2, 5)
    assert (app.status_var.value, app.progress_var.value) == before


# 启动批次：按钮禁用、后台线程 + 轮询排程


def test_remove_kicks_off_background_thread_and_polling(monkeypatch):
    app = make_app()
    app.file_path_var.set(r'C:\docs')
    monkeypatch.setattr(style_cleaner.threading, 'Thread', FakeThread)

    app.remove_unused_styles()

    assert app.remove_button.state == 'disabled'
    assert app.choose_file_button.state == 'disabled'
    assert app.choose_folder_button.state == 'disabled'
    assert app.status_var.value == '正在处理...'
    assert app._batch_running is True
    # 后台线程：worker 目标、daemon、带目标与覆盖模式参数
    assert app.batch_thread.started is True
    assert app.batch_thread.daemon is True
    assert app.batch_thread.target == app._batch_worker
    assert app.batch_thread.args == (r'C:\docs', False, app.batch_queue)
    # 主线程轮询已排程
    assert app.root.after_calls


def test_remove_without_target_warns_and_does_not_start(monkeypatch):
    app = make_app()
    warned = []
    monkeypatch.setattr(style_cleaner.messagebox, 'showwarning', lambda *a, **k: warned.append(a))
    monkeypatch.setattr(style_cleaner.threading, 'Thread', FakeThread)

    app.remove_unused_styles()

    assert warned
    assert app._batch_running is False
    assert app.remove_button.state == 'normal'


def test_remove_cancelled_overwrite_confirmation_does_not_start(monkeypatch):
    app = make_app()
    app.file_path_var.set(r'C:\docs\a.docx')
    app.overwrite_var.set(True)
    monkeypatch.setattr(style_cleaner.threading, 'Thread', FakeThread)
    monkeypatch.setattr(WordStyleCleaner, '_confirm_overwrite', lambda self, target: False)

    app.remove_unused_styles()

    assert app.status_var.value == '已取消覆盖模式清理，未做任何修改'
    assert app._batch_running is False
    assert app.remove_button.state == 'normal'


def test_remove_ignores_reentry_while_batch_running(monkeypatch):
    # 防御路径：渲染异常恢复按钮后若后台线程仍在跑，重入点击直接忽略
    app = make_app()
    app.file_path_var.set(r'C:\docs')
    app._batch_running = True
    created = []

    class RecordingThread(FakeThread):
        def __init__(self, **kw):
            super().__init__(**kw)
            created.append(self)

    monkeypatch.setattr(style_cleaner.threading, 'Thread', RecordingThread)

    app.remove_unused_styles()

    assert created == []
    assert app.root.after_calls == []


# 后台 worker：只往队列塞事件，不碰 widget


def test_batch_worker_queues_start_done_and_finished_events(monkeypatch):
    seen = {}

    def fake_run_batch(target, on_progress=None, overwrite=False, on_file_done=None):
        seen['target'] = target
        seen['overwrite'] = overwrite
        on_progress(1, 2, r'C:\docs\a.docx')
        on_file_done(1, 2)
        return BatchResult([FileResult(input_path=r'C:\docs\a.docx')])

    monkeypatch.setattr(style_cleaner, 'run_batch', fake_run_batch)
    app = make_app()
    q = queue.Queue()

    app._batch_worker(r'C:\docs', True, q)

    assert seen == {'target': r'C:\docs', 'overwrite': True}
    events = []
    while not q.empty():
        events.append(q.get_nowait())
    assert [e[0] for e in events] == ['start', 'file_done', 'finished']
    assert events[0] == ('start', 1, 2, r'C:\docs\a.docx')
    assert events[1] == ('file_done', 1, 2)
    assert isinstance(events[2][1], BatchResult)


def test_batch_worker_queues_failure_event_on_error(monkeypatch):
    def boom(target, on_progress=None, overwrite=False, on_file_done=None):
        raise ValueError(f'无效的文件或文件夹路径: {target}')

    monkeypatch.setattr(style_cleaner, 'run_batch', boom)
    app = make_app()
    q = queue.Queue()

    app._batch_worker(r'C:\不存在', False, q)

    kind, message = q.get_nowait()
    assert kind == 'failed'
    assert 'ValueError' in message


# 批次结束：渲染 + 弹窗 + 恢复按钮与进度条（既有行为保留）


def test_finish_batch_renders_results_and_restores_ui(monkeypatch):
    app = make_app()
    app._batch_running = True
    app.progress_bar.running = True  # 模拟单文件忙碌条还在走
    rendered = []
    info = []
    monkeypatch.setattr(app, '_render_results', lambda result: rendered.append(result))
    monkeypatch.setattr(style_cleaner.messagebox, 'showinfo', lambda title, msg, **k: info.append(title))

    result = BatchResult([FileResult(input_path=r'C:\docs\a.docx', output_path=r'C:\docs\a_Q.docx')])
    app._finish_batch(result)

    assert rendered == [result]
    assert info == ['完成']
    assert app.status_var.value == '处理完成'
    assert app._batch_running is False
    assert app.remove_button.state == 'normal'
    assert app.progress_bar.running is False
    assert app.progress_bar.mode == 'determinate'
    assert app.progress_var.value == 0


def test_finish_batch_with_failures_shows_summary(monkeypatch):
    app = make_app()
    app._batch_running = True
    ok = FileResult(input_path=r'C:\docs\a.docx', output_path=r'C:\docs\a_Q.docx')
    bad = FileResult(input_path=r'C:\docs\坏.docx', error='BadZipFile: 不是 docx')
    warnings = []
    monkeypatch.setattr(style_cleaner.messagebox, 'showwarning', lambda title, msg, **k: warnings.append(msg))
    monkeypatch.setattr(app, '_render_results', lambda result: None)

    app._finish_batch(BatchResult([ok, bad]))

    assert len(warnings) == 1
    assert '成功 1 个，失败 1 个' in warnings[0]
    assert '坏.docx' in warnings[0]
    assert app.status_var.value == '处理完成（有失败）'


def test_finish_batch_empty_folder_shows_hint_without_rendering(monkeypatch):
    app = make_app()
    app._batch_running = True
    rendered = []
    info = []
    monkeypatch.setattr(app, '_render_results', lambda result: rendered.append(result))
    monkeypatch.setattr(style_cleaner.messagebox, 'showinfo', lambda title, msg, **k: info.append((title, msg)))

    app._finish_batch(BatchResult())

    assert rendered == []
    assert info == [('提示', '所选文件夹中没有找到Word文档(.docx)')]
    assert app.status_var.value == '就绪'
    assert app.remove_button.state == 'normal'


def test_fail_batch_shows_error_dialog_and_restores(monkeypatch):
    app = make_app()
    app._batch_running = True
    errors = []
    monkeypatch.setattr(style_cleaner.messagebox, 'showerror', lambda title, msg, **k: errors.append(msg))

    app._fail_batch('ValueError: 无效的文件或文件夹路径: X')

    assert errors
    assert '处理失败' in app.status_var.value
    assert app._batch_running is False
    assert app.remove_button.state == 'normal'
    assert app.progress_var.value == 0


# 主线程轮询：消费队列事件，批次结束后停止排程


def test_poll_updates_ui_on_main_thread_and_stops_after_finish(monkeypatch):
    app = make_app()
    app._batch_running = True
    monkeypatch.setattr(app, '_render_results', lambda result: None)
    monkeypatch.setattr(style_cleaner.messagebox, 'showinfo', lambda title, msg, **k: None)

    app.batch_queue.put(('start', 1, 2, r'C:\docs\a.docx'))
    app.batch_queue.put(('file_done', 1, 2))
    app._poll_batch_events()

    # 队列里的事件已消费：状态与进度更新，批次未结束则继续排程
    assert app.status_var.value == '正在处理 (1/2): a.docx'
    assert app.progress_var.value == 50.0
    assert len(app.root.after_calls) == 1

    app.batch_queue.put(('finished', BatchResult([FileResult(input_path=r'C:\docs\a.docx')])))
    app._poll_batch_events()

    assert app._batch_running is False
    assert len(app.root.after_calls) == 1  # 不再排程


def test_poll_survives_handler_crash_by_failing_batch(monkeypatch):
    app = make_app()
    app._batch_running = True
    failures = []
    monkeypatch.setattr(app, '_fail_batch', lambda message: failures.append(message))

    app.batch_queue.put(('finished', '不是 BatchResult'))
    app._poll_batch_events()

    assert failures
    assert app._batch_running is False
