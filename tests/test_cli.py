"""CLI adapter 的 seam 测试：真实小 .docx + capsys，纯 CLI、零 GUI 依赖。"""
import sys
from pathlib import Path

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE

import cli


def make_docx(path, unused_style='未使用样式'):
    doc = Document()
    doc.styles.add_style(unused_style, WD_STYLE_TYPE.PARAGRAPH)
    doc.save(str(path))
    return path


def test_no_args_prints_usage_and_exits_2(capsys):
    # CLI 版单一职责：无参数即参数错误，usage 指路 GUI 版（ADR-0002）
    with pytest.raises(SystemExit) as exc_info:
        cli.main([])
    assert exc_info.value.code == 2
    err = capsys.readouterr().err
    assert 'word-style-cleaner.exe' in err  # usage 指路 GUI 版（prog 名不含 .exe）
    assert '目标路径' in err


def test_success_prints_output_path_and_deleted_styles(tmp_path, capsys):
    make_docx(tmp_path / 'a.docx', unused_style='甲样式')

    code = cli.main([str(tmp_path / 'a.docx')])

    out = capsys.readouterr().out
    assert code == 0
    assert str(tmp_path / 'a_Q.docx') in out
    assert '甲样式' in out
    assert '段落' in out


def test_folder_discovers_same_level_docx_excluding_q_outputs(tmp_path, capsys):
    make_docx(tmp_path / 'a.docx')
    make_docx(tmp_path / 'b_Q.docx')  # 自产输出，应被排除

    code = cli.main([str(tmp_path)])

    out = capsys.readouterr().out
    assert code == 0
    assert 'a.docx' in out
    assert '_Q.docx' not in out.replace('a_Q.docx', '')  # b 未被处理


def test_overwrite_flag_replaces_original_without_q_copy(tmp_path, capsys):
    src = make_docx(tmp_path / 'a.docx')

    code = cli.main([str(src), '--overwrite'])

    assert code == 0
    assert not (tmp_path / 'a_Q.docx').exists()
    names = [s.name for s in Document(str(src)).styles]
    assert '未使用样式' not in names
    assert '覆盖原文件' in capsys.readouterr().out


def test_invalid_path_exits_nonzero(tmp_path, capsys):
    code = cli.main([str(tmp_path / '不存在.docx')])

    assert code != 0
    err = capsys.readouterr().err
    assert '无效' in err


def test_empty_folder_exits_nonzero(tmp_path, capsys):
    code = cli.main([str(tmp_path)])

    assert code != 0
    assert 'Word 文档' in capsys.readouterr().err


def test_any_file_failure_exits_nonzero_but_reports_each(tmp_path, capsys):
    (tmp_path / '1坏.docx').write_bytes(b'not a real docx')
    make_docx(tmp_path / '2好.docx')

    code = cli.main([str(tmp_path)])

    out = capsys.readouterr().out
    assert code != 0
    assert '1坏.docx' in out and '失败' in out
    assert '2好.docx' in out
    assert str(tmp_path / '2好_Q.docx') in out


def test_all_success_exits_zero(tmp_path, capsys):
    make_docx(tmp_path / 'a.docx')
    make_docx(tmp_path / 'b.docx')

    assert cli.main([str(tmp_path)]) == 0


def test_param_error_exits_nonzero(capsys):
    with pytest.raises(SystemExit) as exc_info:
        cli.main(['--不存在的选项'])
    assert exc_info.value.code != 0


def test_flag_without_target_is_param_error(capsys):
    with pytest.raises(SystemExit) as exc_info:
        cli.main(['--overwrite'])
    assert exc_info.value.code != 0
    assert '目标路径' in capsys.readouterr().err


def test_cli_path_never_imports_tkinter(tmp_path):
    # GUI 只允许在 _launch_gui 里延迟 import，CLI 路径保持零 GUI 依赖
    make_docx(tmp_path / 'a.docx')
    sys.modules.pop('tkinter', None)

    cli.main([str(tmp_path)])

    assert 'tkinter' not in sys.modules
