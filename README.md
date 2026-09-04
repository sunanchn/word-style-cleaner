# Word 样式清理工具（word-style-cleaner）

一个批量删除 Word 文档中**未使用样式**的小工具。提供图形界面和命令行两个程序：图形界面版 `word-style-cleaner.exe` 双击即用；命令行版 `word-style-cleaner-cli.exe` 供终端与脚本使用。不需要安装 Word，不需要写代码。

## 它解决什么问题

从别处复制粘贴内容多了，Word 文档的样式列表会积累几百条垃圾样式：样式面板又长又乱、文件莫名变大、真正想找的样式被淹没。手动一个个删非常痛苦。

本工具扫描 `.docx` 文档正文中实际引用到的样式，把没有被引用的样式全部删除。

## 功能特点

- 两个单文件 exe：图形界面版 `word-style-cleaner.exe`（双击即用，无控制台窗口）+ 命令行版 `word-style-cleaner-cli.exe`
- 图形界面与命令行都支持**单个文件**或**整个文件夹**批量处理
- 两种输出模式：默认生成 `_Q` 副本（原文件不动）；可开启**覆盖原文件**模式（GUI 需弹窗确认，CLI 需显式 `--overwrite`）
- 实时进度条，处理过程中始终显示当前文件名
- 重复处理同一文件夹时，自动跳过上轮生成的 `*_Q` 副本，不会叠加后缀
- 单个文件失败不会中断批处理，结束后汇总失败文件与原因
- 处理完成后显示统计：共处理多少个文件、删除多少个样式
- 删除明细列表：每个文件删掉了哪些样式，一目了然

<!-- 建议以后在这里补一张运行截图，效果胜过千言 -->

## 安装与运行

### 方式一：免安装 exe（推荐，无需 Python）

在仓库根目录执行构建（构建机需装 Python 和 PyInstaller）：

```bash
pip install -r requirements-dev.txt          # 含 python-docx + pytest + pyinstaller
python -m PyInstaller --clean --noconfirm word-style-cleaner.spec
```

一次构建产出两个单文件 exe（`dist/word-style-cleaner.exe` 图形界面版、`dist/word-style-cleaner-cli.exe` 命令行版），免安装，可直接拷到任何 Windows 机器（无需 Python）使用。

### 方式二：从源码运行

需要 Python 3.8 及以上（[下载地址](https://www.python.org/downloads/)，安装时勾选 "Add Python to PATH"）。

```bash
# 1. 进入本目录，安装依赖（只有一个：python-docx）
pip install -r requirements.txt

# 2. 运行图形界面（与 GUI 版 exe 行为一致）
python style_cleaner.py
```

命令行版对应 `python cli.py <参数>`，参数用法见下文 CLI 部分。

## 使用方法（GUI）

1. 双击 `word-style-cleaner.exe`（源码方式为 `python style_cleaner.py`），启动后点「选择文件」或「选择文件夹」
2. 点「删除未使用的样式」；如需直接写回原文件，先勾选「覆盖原文件」（会弹窗确认）
3. 等进度条走完，查看统计信息和「处理详情」（含失败文件与原因）

> 说明：图形界面版不接受命令行参数——带参数启动会弹窗提示改用 `word-style-cleaner-cli.exe`。命令行用法请看下一节。

## 使用方法（CLI）

使用命令行版 `word-style-cleaner-cli.exe`（源码方式为 `python cli.py <参数>`）：

```bash
# 清理单个文件，生成 xxx_Q.docx 副本（原文件不动）
word-style-cleaner-cli.exe 文档.docx

# 清理整个文件夹下所有 .docx（自动排除本工具生成的 _Q 副本）
word-style-cleaner-cli.exe D:\docs

# 覆盖模式：清理结果直接写回原文件，不生成副本（--overwrite 本身即确认，无交互）
word-style-cleaner-cli.exe 文档.docx --overwrite
```

处理结果逐文件输出（输出路径、按类别分组的删除样式明细），结束后汇总成功/失败数量。退出码：`0` 全部成功；`1` 存在失败或路径无效；`2` 参数错误（无参数同此）。

查看帮助：`word-style-cleaner-cli.exe --help`

> 自 v0.2 升级的用户注意：命令行用法由 `word-style-cleaner.exe <参数>` 改为 `word-style-cleaner-cli.exe <参数>`；旧 exe 带参数启动现在只会弹窗指路，不再进入命令行模式。

## 注意事项

- 仅支持 `.docx` 格式（不支持老式 `.doc`，可先用 Word 另存为 `.docx`）
- 清理效果很彻底：表格样式也会被删除，但表格中文字的格式不受影响
- 建议处理前先备份原文档（或先在副本上试运行）

## 开发与测试

```bash
pip install -r requirements-dev.txt   # 运行依赖 + pytest
python -m pytest tests/               # 运行测试（无需图形界面）
```

## 项目状态

个人学习与自用项目，按需维护。发现问题或有想法，欢迎[提 issue](https://github.com/sunanchn/word-style-cleaner/issues)。

潜在方向（欢迎 PR，非承诺）：英文界面、清理前预览确认、`.doc`/WPS 支持。

已知限制：python-docx 模板自带的约 19 个 "… Char" 影子样式（如 Header Char）会出现在删除明细中——无害，属既有行为。

## 许可证

[MIT License](./LICENSE)
