"""批量处理 module 的 seam 测试：临时目录 + 真实小 .docx，不碰 GUI。"""
from pathlib import Path

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE

import batch
from batch import discover_docx, run_batch


def make_docx(path, unused_style='未使用样式'):
    """在 path 生成一个包含一条未使用自定义样式的真实 .docx。"""
    doc = Document()
    doc.styles.add_style(unused_style, WD_STYLE_TYPE.PARAGRAPH)
    doc.save(str(path))
    return path


def style_names(path):
    doc = Document(str(path))
    return [s.name for s in doc.styles]


def test_discover_excludes_self_outputs(tmp_path):
    make_docx(tmp_path / '报告.docx')
    make_docx(tmp_path / '旧输出_Q.docx')
    make_docx(tmp_path / '旧输出_Q(1).docx')
    (tmp_path / '说明.txt').write_text('不是 Word 文档')

    found = discover_docx(str(tmp_path))

    assert [Path(p).name for p in found] == ['报告.docx']


def test_folder_rerun_skips_q_outputs_without_suffix_stacking(tmp_path):
    make_docx(tmp_path / '报告.docx')
    # 上一轮已产出的 _Q 输出
    make_docx(tmp_path / '报告_Q.docx')

    result = run_batch(str(tmp_path))

    assert len(result.results) == 1
    report = result.results[0]
    assert Path(report.input_path).name == '报告.docx'
    assert report.ok
    # 防覆盖：_Q 已被占用，落位到 (1) 变体，而不是叠加成 _Q_Q
    assert Path(report.output_path).name == '报告_Q(1).docx'
    assert Path(report.output_path).exists()
    assert '未使用样式' not in style_names(report.output_path)
    assert not (tmp_path / '报告_Q_Q.docx').exists()


def test_single_file_output_gets_q_suffix(tmp_path):
    src = make_docx(tmp_path / 'a.docx')

    result = run_batch(str(src))

    assert len(result.results) == 1
    assert Path(result.results[0].output_path).name == 'a_Q.docx'


def test_output_counter_advances_past_existing_variants(tmp_path):
    src = make_docx(tmp_path / 'a.docx')
    make_docx(tmp_path / 'a_Q.docx')
    make_docx(tmp_path / 'a_Q(1).docx')

    result = run_batch(str(src))

    assert Path(result.results[0].output_path).name == 'a_Q(2).docx'


def test_failed_file_does_not_interrupt_batch(tmp_path):
    # 排序保证坏文件排在好文件前面：坏文件在前仍被继续处理
    (tmp_path / '1坏.docx').write_bytes(b'not a real docx')
    make_docx(tmp_path / '2好.docx')

    result = run_batch(str(tmp_path))

    assert len(result.results) == 2
    bad, good = result.results
    assert Path(bad.input_path).name == '1坏.docx'
    assert not bad.ok
    assert bad.error
    assert bad.output_path is None
    assert good.ok
    assert Path(good.output_path).exists()


def test_per_file_results_carry_deleted_styles(tmp_path):
    make_docx(tmp_path / 'a.docx', unused_style='甲样式')
    make_docx(tmp_path / 'b.docx', unused_style='乙样式')

    result = run_batch(str(tmp_path))

    # 默认模板自带的非内置 "… Char" 样式也会被清理（v0.1 起的既有行为），
    # 这里只断言每个文件的结果携带各自文件里删掉的样式
    by_name = {Path(r.input_path).name: r for r in result.results}
    assert '甲样式' in [d.name for d in by_name['a.docx'].deleted]
    assert '乙样式' in [d.name for d in by_name['b.docx'].deleted]
    assert '乙样式' not in [d.name for d in by_name['a.docx'].deleted]
    assert len(result.succeeded) == 2
    assert result.failed == []


def test_progress_callback_reports_every_file_before_processing(tmp_path):
    make_docx(tmp_path / 'a.docx')
    make_docx(tmp_path / 'b.docx')
    make_docx(tmp_path / 'c.docx')
    calls = []

    run_batch(str(tmp_path), on_progress=lambda i, total, path: calls.append((i, total, Path(path).name)))

    assert calls == [(1, 3, 'a.docx'), (2, 3, 'b.docx'), (3, 3, 'c.docx')]


def test_overwrite_single_file_replaces_original_without_q_copy(tmp_path):
    src = make_docx(tmp_path / 'a.docx')

    result = run_batch(str(src), overwrite=True)

    assert len(result.results) == 1
    report = result.results[0]
    assert report.ok
    assert report.output_path == str(src)
    assert '未使用样式' not in style_names(src)
    assert not (tmp_path / 'a_Q.docx').exists()


def test_overwrite_folder_replaces_originals_without_q_copies(tmp_path):
    make_docx(tmp_path / 'a.docx', unused_style='甲样式')
    make_docx(tmp_path / 'b.docx', unused_style='乙样式')

    result = run_batch(str(tmp_path), overwrite=True)

    assert len(result.succeeded) == 2
    for r in result.results:
        assert r.output_path == r.input_path
    assert '甲样式' not in style_names(tmp_path / 'a.docx')
    assert '乙样式' not in style_names(tmp_path / 'b.docx')
    assert list(tmp_path.glob('*_Q*.docx')) == []


def test_overwrite_save_failure_leaves_original_intact(tmp_path, monkeypatch):
    src = make_docx(tmp_path / 'a.docx')

    def half_written_save(self, path):
        with open(path, 'wb') as f:
            f.write('PK\x03\x04 写一半的损坏内容'.encode('utf-8'))
        raise OSError('模拟保存中途失败')

    monkeypatch.setattr('docx.document.Document.save', half_written_save)

    result = run_batch(str(src), overwrite=True)

    report = result.results[0]
    assert not report.ok
    assert report.output_path is None
    # 原文件完好无损，也没有留下临时文件
    assert '未使用样式' in style_names(src)
    assert list(tmp_path.glob('*.tmp')) == []


def test_invalid_target_raises_value_error(tmp_path):
    with pytest.raises(ValueError):
        run_batch(str(tmp_path / '不存在.docx'))


def test_batch_module_never_references_tkinter():
    source = Path(batch.__file__).read_text(encoding='utf-8')
    assert 'tkinter' not in source
