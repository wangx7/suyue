# 素阅 · 文档阅读器

在 VS Code 里安静地阅读 Word、Excel、CSV 和 Markdown。文件在本地解析，不上传；跟随 VS Code 的深浅色主题。

![Markdown](https://raw.githubusercontent.com/wangx7/suyue/main/store/screenshots/md-light.png)

![表格](https://raw.githubusercontent.com/wangx7/suyue/main/store/screenshots/sheet-dark.png)

## 打开方式

| 文件 | 怎么打开 |
|---|---|
| Word：.docx、.doc | 直接点开 |
| 表格：.xlsx、.xlsm、.xlsb、.xls、.ods | 直接点开 |
| Markdown：.md 等 | 资源管理器里右键 →“用素阅打开”，或文本编辑器右上角的书本图标 |
| CSV / TSV | 同上 |

Markdown 和 CSV 默认仍用文本编辑器打开。想一直用素阅，可以右键 →“打开方式…”→“配置默认编辑器”选“素阅”。

## 功能

- 左侧自动生成目录：Word 按标题样式，Markdown 按 `#` 标题，表格按工作表
- Word 有阅读视图和页面视图，深色下文档颜色自动调整
- 表格用 Canvas 绘制，几十万行也能流畅滚动；⌘/Ctrl + F 在表格里搜索
- 双指捏合或 Ctrl + 滚轮缩放
- 记住每个文件的阅读位置和缩放
- 文件在磁盘上改动后自动刷新，适合一边编辑 Markdown 一边看效果
- Markdown 里相对路径的图片和链接都能用

也可以在浏览器里使用：https://wangx7.github.io/suyue/

## 已知限制

- .doc 只显示文字
- 表格不显示单元格填充色和字体样式
- Markdown 不渲染数学公式、Mermaid 和脚注
