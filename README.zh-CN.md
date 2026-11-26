# dsh-preview-printer

[English](README.md) | 中文

在 DSH 右侧文档预览里直接打印，并顺手救回预览打不开的旧格式文件。

## 它加了什么

文档工具栏里始终有一个按钮（`sidebar.right.tab.document.actions`），按文件类型呈现三种状态：

| 状态 | 适用文件 | 行为 |
|---|---|---|
| 正常 | PDF、DOC、DOCX、PPT、PPTX | 页码 + 清晰度 → Windows 打印对话框；或把 PDF 存到源文件旁；或「高清全速打印」（矢量、不弹框、直连默认打印机） |
| 正常 | PNG、JPG、WEBP、BMP | 一图一页 A4，按比例自动横竖；「存为 PDF」在浏览器合成单页 PDF 下载 |
| 正常 | Markdown | 渲染标题、列表、引用、代码块后排版打印 |
| 正常 | TXT、PY、JSON、JS、BAT | 按文本打印；UTF-8 失败则按 GBK 解码；JSON 自动缩进；超过 2 MB 或 5000 行会拒绝 |
| 变灰 | C、CPP、H、ASM、SH、XLSX、CSV、TSV、GIF | 不可打印；按钮变灰并悬停说明原因（如"请在 Excel 或 WPS 中打印"） |
| 米色药丸 | 预览失败的旧版 DOC/PPT、XLS | 「已修复」/「已转化」：点一下生成可读副本并在预览中打开 |

按钮**始终出现**，所以一眼就能确认插件是否安装生效。

## 安装

1. 下载 Release 里的 `dsh-preview-printer-<版本>.tgz`，或在 DSH 插件管理器里选择「从目录安装」并指向本仓库目录
2. 重启 DSH（Host 侧插件需要重启才加载）
3. 打开任意文档，工具栏应出现打印按钮

需要 DSH Desktop 0.2.2-rc.5 或更新的版本。

## 已知限制

- 旧版 DOC/PPT 容器残缺时，插件会在源文件旁另存 `<名字>-修复.<扩展名>`；原文件不会被修改
- XLS 预览失败时会生成 `<名字>-转换.xlsx`
- 存为 PDF 会在源文件旁生成 `<名字>-print.pdf`
- 打印临时页放在工作区 `.dsh-print`，两小时后自动清理
- 只能操作当前工作区内的文件
- 对话框路径是 200–400 DPI 位图；只有「高清全速打印」是矢量

## 隐私

插件完全在本机运行：**不发起任何网络请求、不做遥测、不上传对话或文件内容**。打印走本机 Windows 对话框，文档转换使用 DSH 自带的 LibreOffice sidecar；只在打开预览或点击按钮时读取当前工作区的文件。

## 许可

MIT，见 [LICENSE](LICENSE)。版权所有 (c) 2026 Qingguo Li。

## 更多文档

- 操作手册（中文，随包分发）：`docs/dsh-preview-printer-manual-zh.docx`
- 英文说明：[README.md](README.md)
