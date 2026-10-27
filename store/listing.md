# 商店上架资料（Chrome / Edge 通用）

## 名称
素阅 · 文档阅读器

## 简短说明（≤ 132 字符）
安静地阅读 Word、Excel、CSV 和 Markdown。文件在本地解析，不上传，支持深色模式。

## 详细说明
素阅是一个安静的文档阅读器，在浏览器里直接打开 Word、Excel、CSV 和 Markdown 文件，不需要安装 Office，也不需要把文件上传到任何地方。

支持的格式：
• Word：.docx（阅读视图 / 页面视图）、.doc（正文）
• 表格：.xlsx、.xlsm、.xlsb、.xls、.ods、.csv、.tsv，大表格也能流畅滚动和搜索
• Markdown：表格、任务列表、代码高亮
• PDF、图片、音视频、网页等浏览器能直接显示的文件，自动在新标签页打开

特点：
• 文件只在本地解析，不上传、不收集任何数据
• 浅色和深色都是原生配色，Word 文档里的颜色也会跟着调整
• 左侧自动生成目录；向下阅读时按钮自动淡出
• 双指捏合或 Ctrl + 滚轮缩放
• 自动记住每个文件的阅读位置和缩放
• 在网页上右键文档链接，选“用素阅打开”

## 类别
工作效率（Productivity）

## 隐私政策网址
https://github.com/wangx7/suyue/blob/main/PRIVACY.md

## 网站
https://wangx7.github.io/suyue/

## 单一用途说明
在浏览器中本地查看 Word、Excel、CSV 和 Markdown 文件。

## 权限理由
- contextMenus：在网页上的文档链接（.docx、.xlsx、.md 等）右键菜单中加入“用素阅打开”。
- downloads：用户把本地的 Word、Excel 等文件拖进浏览器标签页时，Chrome 会把它当作下载处理；扩展只拦截这类本地文件（file://）的下载，取消它并改用素阅打开，网上文件的下载不受影响。
- host_permissions（file:///*）：读取用户拖进浏览器或右键打开的本地文件。只有用户在扩展详情页打开“允许访问文件网址”后才生效，不访问任何网站。
- optional_host_permissions（<all_urls>）：可选权限，安装时不申请。只有当用户打开某个网站上的文档、下载因跨域失败、并主动点击“允许读取该网站”时，才针对这一个网站申请，用于下载要显示的文档。

## 远程代码
不使用。所有脚本都打包在扩展内。

## 数据使用
不收集任何用户数据。

## 给审核人员的说明（Edge：Notes for certification）
无需登录。点击工具栏图标打开阅读器，把任意 .docx / .xlsx / .csv / .md 文件拖进页面即可查看。也可以在网页上右键这类文件的链接，选“用素阅打开”。

## 截图（1280×800，store/screenshots/）
1. md-light.png — Markdown，浅色，左侧目录
2. sheet-dark.png — 表格，深色
3. md-dark.png — Markdown，深色
4. sheet-light.png — 表格，浅色
5. home-light.png — 首页
