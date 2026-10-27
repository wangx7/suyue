# 素阅

安静地阅读 Word、Excel、CSV 和 Markdown。文件在浏览器里解析，不会上传；浅色、深色都是原生配色。
同一份代码既能当网页用，也能打包成 Chrome / Edge 扩展。

在线使用：https://wangx7.github.io/suyue/

界面只有三样东西：左侧目录、正文、右上角几个按钮。向下阅读时按钮自动淡出，向上滚动或把鼠标移到顶部时出现。

| 格式 | 实现 | 目录 | 深色 |
|---|---|---|---|
| docx | docx-preview，默认阅读视图，可切换页面视图 | 标题样式（标题 1–9） | 文字、底纹、边框颜色重映射，图片保持原色 |
| doc（97-2003） | 自己解析 piece table，只取正文 | 按“第一章 / 一、/ 1.1”这类编号推测 | 纯文字 |
| xlsx / xls / ods / csv | SheetJS 在 Web Worker 里解析，Canvas 虚拟网格 | 工作表列表 | 网格按主题绘制 |
| md / markdown | markdown-it（GFM 表格、任务列表、自动链接），内嵌 HTML 经 DOMPurify 过滤，highlight.js 着色 | `#` 标题 | 主题变量 |

## 操作

| 操作 | 作用 |
|---|---|
| 拖入文件 / ⌘/Ctrl + O | 打开文件 |
| 右上角主页按钮 | 回到首页，查看最近打开的文件 |
| ⌘/Ctrl + \ | 显示 / 隐藏目录 |
| 双指捏合 / Ctrl + 滚轮 | 缩放，指针下的内容不动；点右上角的百分比恢复 100% |
| ⌘/Ctrl + F | 表格里搜索；Word 和 Markdown 用浏览器自带的查找 |

每个文件自动记住阅读位置、缩放、视图和选中格。Word 阅读视图把正文放大到约 16px（最多 1.2 倍）；文档里的 Windows 字体（宋体、仿宋、微软雅黑、Calibri 等）在没有安装的系统上换成字形相近的本机字体。

## 开发

```bash
npm install
npm run dev      # 浏览器打开提示的地址，把文件拖进页面
npm run build    # 输出到 dist/
```

`?url=` 参数可以直接打开远程文件，例如 `http://localhost:5173/?url=https://example.com/a.docx`（对方服务器需允许跨域）。

`patches/` 是对 docx-preview 的修补，`npm install` 时由 patch-package 自动打上：
- 表格 `tblLayout` 读错属性；缺省的 `lineRule` 被当成固定值（规范默认是 auto）。
- 支持“字符”单位的缩进和“行”单位的段前段后间距。
- 读出文档网格（`docGrid`）、“对齐到网格”、“相同样式段落间不加空格”和行距规则，交给 `docx-layout.ts` 按实际字号计算行高。
- 元素的 `--docx-*` 自定义属性改用 `setProperty` 设置。

升级 docx-preview 前先确认这些是否已修复或改动。本地存储的键名仍以 `anyue` 开头（项目旧名），改了会丢失用户的设置和最近打开。

## 安装为浏览器扩展

1. 运行 `npm run build`。
2. 打开 `chrome://extensions`（Edge 是 `edge://extensions`），开启“开发者模式”。
3. 点“加载已解压的扩展程序”，选择 `dist/` 目录。

之后点工具栏图标打开，或在网页上右键 .docx / .xlsx / .md 等链接选“用素阅打开”。右键本地 `file://` 链接需要在扩展详情页开启“允许访问文件网址”。

## 已知限制

- .doc 只显示文字，没有排版和图片。
- docx 页面视图只在明确的分页符或分节处分页；文本框、域代码等复杂排版还原有限。
- 表格不显示单元格填充色和字体样式（SheetJS 社区版不读取）。
- Markdown 不渲染数学公式、Mermaid 和脚注；本地打开时读不到同目录的图片。
- 不支持 pptx。PDF、图片、音视频、网页、JSON 等浏览器自己能显示的文件，会直接在新标签页里交给浏览器打开；本地网页在沙箱里运行，读不到同目录的图片和脚本。
