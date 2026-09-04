"""CLI adapter：参数解析、逐文件结果输出与退出码。

单 exe 双模式入口（ADR-0001）：无参数启动 GUI；带参数进入 CLI。
全程不弹任何 GUI 窗口——本 module 不 import tkinter，GUI 模式通过
延迟 import style_cleaner 进入。--overwrite flag 即覆盖确认，无交互。
"""
import argparse
import os
import sys

from batch import FileResult, run_batch
from cleaner import StyleCategory

# 清理结果的中性类别 → 中文展示名；展示词汇只存在于 CLI adapter
CATEGORY_LABELS = {
    StyleCategory.PARAGRAPH: '段落',
    StyleCategory.CHARACTER: '字符',
    StyleCategory.TABLE: '表格',
    StyleCategory.OTHER: '其他',
}

# 参数错误由 argparse 以 2 退出；业务失败（无效路径/无可处理文件/任一文件失败）统一为 1
EXIT_FAILURE = 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog='word-style-cleaner',
        description='Word 样式清理工具：删除文档中未使用的样式。无参数启动图形界面。',
    )
    parser.add_argument(
        'target', nargs='?',
        help='目标路径：单个 .docx 文件或文件夹（同层 .docx，排除自产 _Q 输出）',
    )
    parser.add_argument(
        '--overwrite', action='store_true',
        help='覆盖模式：清理结果直接写回原文件，不生成 _Q 副本（本 flag 即确认）',
    )
    return parser


def _format_deleted(deleted) -> list[str]:
    """按类别分组渲染一个文件删除的样式，顺序与类别枚举一致。"""
    lines = []
    by_label = {}
    for d in deleted:
        by_label.setdefault(CATEGORY_LABELS[d.category], []).append(d.name)
    for label in CATEGORY_LABELS.values():
        names = by_label.get(label)
        if names:
            lines.append(f'  {label}样式: {", ".join(sorted(names))}')
    return lines


def _print_file_result(result: FileResult, index: int, total: int):
    name = os.path.basename(result.input_path)
    print(f'[{index}/{total}] {name}')
    if not result.ok:
        print(f'  失败: {result.error}')
        return
    print(f'  输出: {result.output_path}')
    if result.deleted:
        for line in _format_deleted(result.deleted):
            print(line)
    else:
        print('  未删除任何样式')


def main(argv: list[str] | None = None) -> int:
    """双模式入口：无参数启动 GUI；带参数进入 CLI。返回进程退出码。"""
    args_list = sys.argv[1:] if argv is None else argv
    if not args_list:
        return _launch_gui()

    args = build_parser().parse_args(args_list)
    if args.target is None:
        build_parser().error('缺少目标路径：单个 .docx 文件或文件夹')

    def on_progress(index, total, input_path):
        print(f'正在处理 ({index}/{total}): {os.path.basename(input_path)}')

    try:
        result = run_batch(args.target, on_progress=on_progress, overwrite=args.overwrite)
    except ValueError as e:
        print(f'错误: {e}', file=sys.stderr)
        return EXIT_FAILURE

    if not result.results:
        print('未发现可处理的 Word 文档(.docx)', file=sys.stderr)
        return EXIT_FAILURE

    for index, file_result in enumerate(result.results, start=1):
        _print_file_result(file_result, index, len(result.results))

    mode = '覆盖原文件' if args.overwrite else '生成 _Q 副本'
    summary = f'完成: 成功 {len(result.succeeded)} 个, 失败 {len(result.failed)} 个 ({mode})'
    print(summary)

    return 0 if not result.failed else EXIT_FAILURE


def _launch_gui() -> int:
    """无参数路径：延迟 import GUI（tkinter），CLI 路径保持零 GUI 依赖。"""
    import style_cleaner
    style_cleaner.launch()
    return 0


if __name__ == '__main__':
    sys.exit(main())
