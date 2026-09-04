"""批量处理 module：.docx 发现、_Q 命名与防覆盖、逐文件清理与结果汇总。

不碰 GUI：进度通过回调上报，弹窗与展示由 adapter（GUI/CLI）负责。
自产输出指本工具生成的 `_Q` 副本及其防覆盖变体 `_Q(数字)`；文件夹发现时排除，
避免重跑时后缀叠加。显式指定的单文件不做排除。
"""
import os
import re
import tempfile
from dataclasses import dataclass, field

from docx import Document

from cleaner import DeletedStyle, clean_document

OUTPUT_SUFFIX = '_Q'

# 自产输出：*_Q.docx 与防覆盖计数变体 *_Q(1).docx
_SELF_OUTPUT_RE = re.compile(rf'{OUTPUT_SUFFIX}(\(\d+\))?$')


@dataclass
class FileResult:
    input_path: str
    output_path: str | None = None
    deleted: list[DeletedStyle] = field(default_factory=list)
    error: str | None = None

    @property
    def ok(self) -> bool:
        return self.error is None


@dataclass
class BatchResult:
    results: list[FileResult] = field(default_factory=list)

    @property
    def succeeded(self) -> list[FileResult]:
        return [r for r in self.results if r.ok]

    @property
    def failed(self) -> list[FileResult]:
        return [r for r in self.results if not r.ok]


def discover_docx(folder: str) -> list[str]:
    """列出文件夹中的 .docx（不递归），排除自产 _Q 输出；按文件名排序。"""
    names = [
        name for name in os.listdir(folder)
        if name.lower().endswith('.docx')
        and not _SELF_OUTPUT_RE.search(os.path.splitext(name)[0])
    ]
    return [os.path.join(folder, name) for name in sorted(names)]


def _output_path_for(input_path: str) -> str:
    """在原文件夹生成 `_Q` 输出路径；已存在则追加 (n) 计数。"""
    folder = os.path.dirname(input_path)
    base = os.path.splitext(os.path.basename(input_path))[0]
    candidate = os.path.join(folder, f'{base}{OUTPUT_SUFFIX}.docx')
    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(folder, f'{base}{OUTPUT_SUFFIX}({counter}).docx')
        counter += 1
    return candidate


def _save_in_place(document, input_path: str):
    """写回原文件：先写同目录临时文件，成功后原子替换，避免写一半损坏原文件。"""
    fd, temp_path = tempfile.mkstemp(
        prefix=os.path.basename(input_path) + '.',
        suffix='.tmp',
        dir=os.path.dirname(input_path) or '.',
    )
    os.close(fd)
    try:
        document.save(temp_path)
        os.replace(temp_path, input_path)
    except BaseException:
        try:
            os.remove(temp_path)
        except OSError:
            pass
        raise


def _clean_file(input_path: str, overwrite: bool = False) -> FileResult:
    """清理单个文件：默认保存为 _Q 副本，overwrite=True 写回原文件；失败记录原因，不抛异常。"""
    try:
        output_path = input_path if overwrite else _output_path_for(input_path)
        document = Document(input_path)
        deleted = clean_document(document)
        if overwrite:
            _save_in_place(document, output_path)
        else:
            document.save(output_path)
        return FileResult(input_path=input_path, output_path=output_path, deleted=deleted)
    except Exception as e:
        return FileResult(input_path=input_path, error=f'{type(e).__name__}: {e}')


def run_batch(target: str, on_progress=None, overwrite: bool = False) -> BatchResult:
    """批量清理 target（单个 .docx 或文件夹），返回逐文件的 BatchResult。

    on_progress(index, total, input_path) 在每个文件处理前回调一次。
    默认产出 `_Q` 副本；overwrite=True 时清理结果写回原文件路径，
    不产 `_Q` 副本（覆盖模式的确认由 adapter 负责，本 module 不弹窗）。
    单个文件失败记录到结果后继续，不中断批次；target 无效时抛 ValueError。
    """
    if os.path.isfile(target):
        input_paths = [target]
    elif os.path.isdir(target):
        input_paths = discover_docx(target)
    else:
        raise ValueError(f'无效的文件或文件夹路径: {target}')

    results = []
    total = len(input_paths)
    for index, input_path in enumerate(input_paths, start=1):
        if on_progress:
            on_progress(index, total, input_path)
        results.append(_clean_file(input_path, overwrite=overwrite))
    return BatchResult(results=results)
