# 产品形态：双 exe（GUI windowed + CLI console）

取代 [0001](0001-product-form-single-exe-dual-mode.md) 的"单 exe 双模式"：同一次构建产出两个单文件 exe——GUI 版 `word-style-cleaner.exe` 用 windowed 子系统（彻底无控制台），CLI 版 `word-style-cleaner-cli.exe` 用 console 子系统（原生输出、退出码与 shell 等待语义）。动因（#8）：实测单文件 exe 从启动到用户代码执行约 1.4–1.6 秒，console 子系统下双击进 GUI 必然伴随这段黑框闪现（进程内只能尽早释放控制台，无法阻止其出现），对面向非技术用户的双击主路径，观感伤害大于多发一个文件的分发成本。v0.2 已按单 exe 发布，CLI 用法从下一版本起改为 `word-style-cleaner-cli.exe`。

各 exe 单一职责：GUI 版收到参数时弹窗指路 CLI 版并以非零码退出（不静默忽略）；CLI 版无参数时打印 usage 并按参数错误退出（exit 2），usage 指路 GUI 版；CLI 路径保持零 GUI 依赖（不 import tkinter）。

## Considered Options

- **单 exe + FreeConsole 尽早释放控制台**：被否——控制台由 Windows 在任何用户代码执行前创建，启动期约 1.5 秒的黑框闪现无法消除，只能把"全程伴随"降级为"每次闪现"。
- **单 exe windowed + CLI 模式 AttachConsole**：被否——windowed 子系统在 cmd/PowerShell 中无控制台输出，等待与退出码语义因 shell 而异且不可靠，破坏脚本化场景。
- **单 exe 双模式（原 0001）**：被否——其"双倍发布物与维护成本，无对应收益"的前提被 #8 推翻；单 spec 定义两个 EXE 目标、一次构建产出双 exe 后，维护成本可忽略。
