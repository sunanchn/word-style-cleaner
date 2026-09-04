"""清理 module interface 的测试：内存构造文档，不碰磁盘、不碰 GUI。"""
from docx import Document
from docx.enum.style import WD_STYLE_TYPE

from cleaner import StyleCategory, clean_document


def make_doc():
    return Document()


def test_unused_custom_paragraph_style_is_deleted():
    doc = make_doc()
    doc.styles.add_style('我的段落样式', WD_STYLE_TYPE.PARAGRAPH)

    deleted = clean_document(doc)

    mine = [s for s in deleted if s.name == '我的段落样式']
    assert len(mine) == 1
    assert mine[0].category is StyleCategory.PARAGRAPH
    assert '我的段落样式' not in [s.name for s in doc.styles]


def test_style_used_in_body_paragraph_is_kept():
    doc = make_doc()
    doc.styles.add_style('正文用到的样式', WD_STYLE_TYPE.PARAGRAPH)
    doc.add_paragraph('内容', style='正文用到的样式')

    deleted = clean_document(doc)

    assert '正文用到的样式' not in [s.name for s in deleted]
    assert '正文用到的样式' in [s.name for s in doc.styles]


def test_builtin_style_never_deleted_even_when_unused():
    doc = make_doc()
    # 'Heading 1' 是内置样式，默认模板中未被使用
    assert 'Heading 1' in [s.name for s in doc.styles]

    deleted = clean_document(doc)

    assert 'Heading 1' not in [s.name for s in deleted]
    assert 'Heading 1' in [s.name for s in doc.styles]


def test_character_style_used_by_run_is_kept():
    doc = make_doc()
    doc.styles.add_style('在用的字符样式', WD_STYLE_TYPE.CHARACTER)
    doc.styles.add_style('没用的字符样式', WD_STYLE_TYPE.CHARACTER)
    p = doc.add_paragraph('内容')
    run = p.add_run('片段')
    run.style = doc.styles['在用的字符样式']

    deleted = clean_document(doc)

    assert '在用的字符样式' in [s.name for s in doc.styles]
    mine = [s for s in deleted if s.name == '没用的字符样式']
    assert len(mine) == 1
    assert mine[0].category is StyleCategory.CHARACTER


def test_table_style_used_by_table_is_kept():
    doc = make_doc()
    doc.styles.add_style('在用的表格样式', WD_STYLE_TYPE.TABLE)
    doc.styles.add_style('没用的表格样式', WD_STYLE_TYPE.TABLE)
    table = doc.add_table(rows=1, cols=1)
    table.style = doc.styles['在用的表格样式']

    deleted = clean_document(doc)

    assert '在用的表格样式' in [s.name for s in doc.styles]
    mine = [s for s in deleted if s.name == '没用的表格样式']
    assert len(mine) == 1
    assert mine[0].category is StyleCategory.TABLE


def test_style_used_inside_table_cell_is_kept():
    doc = make_doc()
    doc.styles.add_style('表格内段落样式', WD_STYLE_TYPE.PARAGRAPH)
    table = doc.add_table(rows=1, cols=1)
    cell = table.rows[0].cells[0]
    cell.paragraphs[0].style = doc.styles['表格内段落样式']

    deleted = clean_document(doc)

    assert '表格内段落样式' not in [s.name for s in deleted]
    assert '表格内段落样式' in [s.name for s in doc.styles]
