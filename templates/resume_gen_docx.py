# -*- coding: utf-8 -*-
"""ToB 销售简历 python-docx 生成模板

使用方式：
1. 复制此文件到目标目录，重命名为 gen_final.py
2. 修改下方 CONTENT 部分为实际简历内容
3. 运行：python gen_final.py
4. 运行后删除脚本

依赖：pip install python-docx（Python 3.14 环境）
"""

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import os

# ── 工具函数（无需修改） ──

def set_run_font(run, name_cn='微软雅黑', name_en='Calibri', size=10.5, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = name_en
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name_cn)

def add_heading_text(doc, text, size=12):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    set_run_font(r, size=size, bold=True)

def add_sub_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(text)
    set_run_font(r, size=10.5, bold=True)

def add_bullet(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.first_line_indent = Cm(-0.4)
    r = p.add_run('• ' + text)
    set_run_font(r, size=10.5)

def add_exp_header(doc, company, role, date):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(1)
    r1 = p.add_run(company)
    set_run_font(r1, size=10.5, bold=True)
    r2 = p.add_run(f'　{role}　{date}')
    set_run_font(r2, size=10.5)

def add_advantage_line(doc, label, desc):
    """个人优势总结格式：**小标题**:描述"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    r1 = p.add_run(f'{label}:')
    set_run_font(r1, size=10.5, bold=True)
    r2 = p.add_run(desc)
    set_run_font(r2, size=10.5)

# ── 简历正文（修改此部分） ──

doc = Document()
for s in doc.sections:
    s.top_margin = Cm(2); s.bottom_margin = Cm(2)
    s.left_margin = Cm(2); s.right_margin = Cm(2)

# 姓名
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
r = p.add_run('候选人姓名')
set_run_font(r, size=16, bold=True)

# 求职意向
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
r = p.add_run('求职意向：ToB 销售 / 目标岗位')
set_run_font(r, size=11)

# ── 在此添加各板块 ──
# add_heading_text(doc, '教育背景')
# add_exp_header(doc, '学校名', '专业/学位', '时间')
# add_bullet(doc, '核心课程：...')
#
# add_heading_text(doc, '工作/实习经历')
# add_exp_header(doc, '公司名', '岗位', '时间')
# add_sub_heading(doc, '小标题')
# add_bullet(doc, '具体描述...')
#
# add_heading_text(doc, '个人优势总结')
# add_advantage_line(doc, '特质小标题', '1-2句佐证描述')

# ── 保存 ──
# out = r'路径\文件名.docx'
# doc.save(out)
# print(f'✅ 已保存: {out}')
