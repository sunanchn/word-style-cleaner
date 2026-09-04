# 产品形态：单 exe 双模式（GUI 默认 + CLI 参数模式）

> 状态：已被 [0002](0002-product-form-dual-exe-gui-cli.md) 取代。

工具面向非技术用户，GUI 为主；CLI 保留为次要形态，供脚本化使用，也让清理 module 的 seam 上有两个真 adapter（GUI 与 CLI）。分发为同一个 PyInstaller 单文件 exe：无参数双击进 GUI，带参数进 CLI（`--overwrite` 显式 flag 即确认，无交互弹窗）。只保证 Windows 可用，Tkinter 代码不刻意破坏跨平台。

## Considered Options

- **仅 GUI**：被否——丢失脚本化场景，seam 只剩一个 adapter（hypothetical）。
- **GUI 与 CLI 两个 exe**：被否——双倍发布物与维护成本，无对应收益。
