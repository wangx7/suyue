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


make_pdf()
make_xlsx()
print('done')
