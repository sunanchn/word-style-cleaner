"""样式清理 module：接收已打开的文档，返回结构化清理结果。

不碰磁盘、不碰 GUI——保存与展示由调用方（GUI/CLI adapter 或 batch）负责。
"""
import re
from dataclasses import dataclass
from enum import Enum, auto

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.opc.part import XmlPart


class StyleCategory(Enum):
    """清理结果中的中性样式类别；中文展示名由 adapter 映射。"""
    PARAGRAPH = auto()
    CHARACTER = auto()
    TABLE = auto()
    OTHER = auto()


@dataclass
class DeletedStyle:
    name: str
    category: StyleCategory


_CATEGORY_BY_WD_TYPE = {
    WD_STYLE_TYPE.PARAGRAPH: StyleCategory.PARAGRAPH,
    WD_STYLE_TYPE.CHARACTER: StyleCategory.CHARACTER,
    WD_STYLE_TYPE.TABLE: StyleCategory.TABLE,
}

# 承载文档文本的 part：正文、页眉/页脚、脚注/尾注。
# 命名如 /word/header1.xml、/word/footnotes.xml。
_STORY_PART_RE = re.compile(r'/word/(document|header|footer|footnotes|endnotes)\d*\.xml$')

# XML 中引用样式的三种元素：段落样式、字符（run）样式、表格样式。
# w:val 存的是 styleId，需要经 styles 部分换算成样式名。
_STYLE_REF_TAGS = (qn('w:pStyle'), qn('w:rStyle'), qn('w:tblStyle'))


def _category_of(style) -> StyleCategory:
    try:
        return _CATEGORY_BY_WD_TYPE.get(style.type, StyleCategory.OTHER)
    except Exception:
        return StyleCategory.OTHER


def _referenced_style_ids(element) -> set:
    """XML 树中引用的样式 ID（w:pStyle / w:rStyle / w:tblStyle 的 w:val）。"""
    ids = set()
    for tag in _STYLE_REF_TAGS:
        for ref in element.iter(tag):
            style_id = ref.get(qn('w:val'))
            if style_id:
                ids.add(style_id)
    return ids


def _story_parts(document):
    """产出承载文档文本的 part：正文、页眉/页脚、脚注/尾注。"""
    for part in document.part.package.parts:
        if _STORY_PART_RE.match(str(part.partname)):
            yield part


def _collect_used_style_names(document) -> set:
    """收集文档中引用的样式名——全位置扫描（CONTEXT.md「未使用的样式」定义）。

    覆盖：正文段落、表格（含单元格段落）、字符 run、页眉/页脚、脚注/尾注、文本框。
    实现上直接在 XML 层收集样式引用，天然覆盖 python-docx 未暴露 API 的位置
    （脚注/尾注 part、文本框内段落）；页眉/页脚按 part 发现，不经 sections API，
    避免其访问内容时隐式创建 part 的副作用。
    """
    id_to_name = {style.style_id: style.name for style in document.styles}

    used_ids = set()
    for part in _story_parts(document):
        # 主文档/页眉等 XmlPart 有实时 element；脚注/尾注等通用 Part 只能解析 blob
        element = part.element if isinstance(part, XmlPart) else parse_xml(part.blob)
        used_ids |= _referenced_style_ids(element)

    return {id_to_name[style_id] for style_id in used_ids if style_id in id_to_name}


def clean_document(document) -> list:
    """删除文档中未使用的自定义样式，返回 DeletedStyle 列表。

    内置样式无论是否使用都永不删除。
    """
    used = _collect_used_style_names(document)

    to_delete = [
        style for style in document.styles
        if style.name not in used and not style.builtin
    ]

    deleted = []
    for style in to_delete:
        try:
            document.styles[style.name].delete()
            deleted.append(DeletedStyle(name=style.name, category=_category_of(style)))
        except Exception:
            # 某些样式可能无法删除，与既有行为一致：忽略
            continue
    return deleted
