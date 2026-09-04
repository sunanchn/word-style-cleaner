# Word 样式清理工具（word-style-cleaner）

一个批量删除 Word 文档中**未使用样式**的小工具。提供图形界面和命令行两种用法：双击 exe 或无参数运行进入图形界面；带参数运行进入命令行模式。不需要安装 Word，不需要写代码。

## 它解决什么问题

从别处复制粘贴内容多了，Word 文档的样式列表会积累几百条垃圾样式：样式面板又长又乱、文件莫名变大、真正想找的样式被淹没。手动一个个删非常痛苦。

本工具扫描 `.docx` 文档正文中实际引用到的样式，把没有被引用的样式全部删除。

## 功能特点

- 双模式单文件 exe：双击进图形界面（基于 tkinter），带参数进命令行
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
python -m PyInstaller --onefile --name word-style-cleaner --hidden-import style_cleaner --clean --noconfirm cli.py
```

产物为 `dist/word-style-cleaner.exe`，单文件、免安装，可直接拷到任何 Windows 机器（无需 Python）使用。

### 方式二：从源码运行

需要 Python 3.8 及以上（[下载地址](https://www.python.org/downloads/)，安装时勾选 "Add Python to PATH"）。

```bash
# 1. 进入本目录，安装依赖（只有一个：python-docx）
pip install -r requirements.txt

# 2. 运行（无参数进 GUI，与 exe 行为一致）
python style_cleaner.py
```

## 使用方法（GUI）

1. 双击 `word-style-cleaner.exe`（或无参数运行），启动后点「选择文件」或「选择文件夹」
2. 点「删除未使用的样式」；如需直接写回原文件，先勾选「覆盖原文件」（会弹窗确认）
3. 等进度条走完，查看统计信息和「处理详情」（含失败文件与原因）

> 说明：exe 采用带控制台的打包方式（CLI 模式需要显示输出），因此双击进入图形界面时会同时出现一个控制台窗口，属正常现象，关闭它即退出程序。

## 使用方法（CLI）

带参数运行同一 exe 即进入命令行模式（源码方式为 `python cli.py <参数>`）：

```bash
# 清理单个文件，生成 xxx_Q.docx 副本（原文件不动）
word-style-cleaner.exe 文档.docx

# 清理整个文件夹下所有 .docx（自动排除本工具生成的 _Q 副本）
word-style-cleaner.exe D:\docs

# 覆盖模式：清理结果直接写回原文件，不生成副本（--overwrite 本身即确认，无交互）
word-style-cleaner.exe 文档.docx --overwrite
```

处理结果逐文件输出（输出路径、按类别分组的删除样式明细），结束后汇总成功/失败数量。退出码：`0` 全部成功；`1` 存在失败或路径无效；`2` 参数错误。

查看帮助：`word-style-cleaner.exe --help`

## 注意事项

- 仅支持 `.docx` 格式（不支持老式 `.doc`，可先用 Word 另存为 `.docx`）
- 清理效果很彻底：表格样式也会被删除，但表格中文字的格式不受影响
- 建议处理前先备份原文档（或先在副本上试运行）

## 开发与测试

```bash
pip install -r requirements-dev.txt   # 运行依赖 + pytest
python -m pytest tests/               # 运行测试（无需图形界面）
```

## 后续计划

- [x] 打包成免安装的单文件 exe（本地构建自用，暂不打 tag、不发 Release）
- [ ] 界面支持英文
- [ ] 清理前预览：先列出"将要删除的样式"供确认
- [ ] 支持 `.doc` / WPS 格式

## 许可证

[MIT License](./LICENSE)
