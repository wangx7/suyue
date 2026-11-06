"""生成测试文档：python3 test-files/gen.py（需要 reportlab、openpyxl）"""
import os
import random

from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas

OUT = os.path.dirname(os.path.abspath(__file__))
random.seed(7)


def make_pdf():
    # STSong-Light 不内嵌字体，依赖 CMap，用来测试中文渲染
    pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
    c = canvas.Canvas(os.path.join(OUT, '测试文档.pdf'), pagesize=A4)
    c.setTitle('素阅 PDF 测试')
    w, h = A4
    for i in range(1, 31):
        c.bookmarkPage(f'p{i}')
        c.addOutlineEntry(f'第 {i} 页', f'p{i}', level=0)
        c.setFont('STSong-Light', 22)
        c.drawString(60, h - 80, f'第 {i} 页：素阅测试文档')
        c.setFont('STSong-Light', 12)
        for n in range(30):
            c.drawString(60, h - 130 - n * 20,
                         f'第 {n + 1} 行：这是用于测试选中复制与搜索的中文文本。Hello PDF {i}-{n + 1}，关键词：苹果、橙子。')
        c.showPage()
    c.save()


def make_xlsx():
    wb = Workbook()
    ws = wb.active
    ws.title = '订单'
    ws.append(['订单号', '城市', '品类', '数量', '金额', '日期', '状态'])
    cities = ['北京', '上海', '广州', '深圳', '杭州', '成都']
    cats = ['手机', '电脑', '耳机', '平板', '手表']
    status = ['已付款', '待发货', '已完成', '已退款']
    for i in range(1, 3001):
        ws.append([f'NO{i:05d}', random.choice(cities), random.choice(cats),
                   random.randint(1, 20), round(random.uniform(20, 9999), 2),
                   f'2026-{random.randint(1, 12):02d}-{random.randint(1, 28):02d}',
                   random.choice(status) if i % 50 else ''])
    s2 = wb.create_sheet('含特殊字符')
    s2.append(['说明', '内容'])
    s2.append(['含制表符', 'a\tb'])
    s2.append(['含换行', '第一行\n第二行'])
    s2.append(['含引号', '他说"你好"'])
    wb.save(os.path.join(OUT, '测试表格.xlsx'))


def make_text():
    rows = ['姓名,部门,工资,备注']
    for i in range(1, 201):
        rows.append(f'员工{i},{random.choice(["研发", "市场", "财务"])},{random.randint(5000, 30000)},"含逗号, 与""引号"""' if i % 10 == 0
                    else f'员工{i},{random.choice(["研发", "市场", "财务"])},{random.randint(5000, 30000)},')
    open(os.path.join(OUT, '测试.csv'), 'w', encoding='utf-8-sig').write('\n'.join(rows) + '\n')
    open(os.path.join(OUT, '测试.tsv'), 'w', encoding='utf-8').write(
        '\n'.join(r.replace(',', '\t') for r in ['产品,价格,库存', '苹果,3.5,100', '橙子,4.2,80', '香蕉,2.8,0']) + '\n')
    md = """# Markdown 测试

## 文本样式
**粗体**、*斜体*、~~删除线~~、`行内代码`、[链接](https://example.com)

> 这是一段引用

## 列表
- 项目一
  - 子项目
- 项目二

1. 第一步
2. 第二步

- [x] 已完成
- [ ] 待办

## 代码
```python
def hello():
    print("你好，素阅")
```

## 表格
| 名称 | 数量 | 单价 |
|------|------|------|
| 苹果 | 3 | 5.5 |
| 橙子 | 5 | 4.0 |

---

脚注与数学：行内 $E=mc^2$。
"""
    open(os.path.join(OUT, '测试.md'), 'w', encoding='utf-8').write(md)


def make_docx():
    from docx import Document
    d = Document()
    d.add_heading('Word 测试文档', 0)
    d.add_paragraph('这是正文，包含 ').add_run('粗体').bold = True
    d.add_heading('列表', 1)
    d.add_paragraph('第一项', style='List Bullet')
    d.add_paragraph('第二项', style='List Bullet')
    d.add_heading('表格', 1)
    t = d.add_table(rows=3, cols=3)
    t.style = 'Table Grid'
    for r in range(3):
        for c in range(3):
            t.cell(r, c).text = f'单元格 {r + 1}-{c + 1}'
    for i in range(20):
        d.add_paragraph(f'第 {i + 1} 段：用于测试滚动、搜索与目录的中文内容。')
    d.save(os.path.join(OUT, '测试.docx'))


def make_pptx():
    from pptx import Presentation
    p = Presentation()
    s = p.slides.add_slide(p.slide_layouts[0])
    s.shapes.title.text = 'PPT 测试'
    s.placeholders[1].text = '素阅演示文稿预览'
    for i in range(2, 6):
        s = p.slides.add_slide(p.slide_layouts[1])
        s.shapes.title.text = f'第 {i} 页'
        s.placeholders[1].text = f'要点 A\n要点 B\n要点 C（第 {i} 页）'
    p.save(os.path.join(OUT, '测试.pptx'))


def make_epub():
    from ebooklib import epub
    b = epub.EpubBook()
    b.set_identifier('suyue-test')
    b.set_title('EPUB 测试')
    b.set_language('zh')
    chs = []
    for i in range(1, 4):
        c = epub.EpubHtml(title=f'第 {i} 章', file_name=f'ch{i}.xhtml', lang='zh')
        c.content = f'<h1>第 {i} 章</h1>' + ''.join(f'<p>这是第 {i} 章第 {n} 段，用于测试阅读与目录。</p>' for n in range(1, 15))
        b.add_item(c)
        chs.append(c)
    b.toc = tuple(chs)
    b.add_item(epub.EpubNcx())
    b.add_item(epub.EpubNav())
    b.spine = ['nav'] + chs
    epub.write_epub(os.path.join(OUT, '测试.epub'), b)


make_pdf()
make_xlsx()
make_text()
make_docx()
make_pptx()
make_epub()
print('done')
