"""GUI 入口的 seam 测试：分发逻辑用假 launch 与假弹窗隔离，绝不真开窗口。"""
import pytest

import style_cleaner


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
