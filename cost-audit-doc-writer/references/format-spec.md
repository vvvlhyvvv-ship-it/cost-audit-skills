# 造价审核公文 —— 完整生成脚本模板（python-docx 直设格式化）

> 配套 Skill `cost-audit-doc-writer`。本文件提供可直接运行的生成脚本骨架，按需替换 `<...>` 占位内容即可。

```python
# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

doc = Document()

# ---- 页面：A4，页边距 上下 2.54 / 左右 3.17cm（两类公文统一；2026-09-11 实测）
# 注：旧版此处曾写"情况说明改用 上/下1.0 左/右1.25" —— 已证伪作废，勿再使用 ----
sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin = Cm(2.54)
sec.bottom_margin = Cm(2.54)
sec.left_margin = Cm(3.17)
sec.right_margin = Cm(3.17)

def add_para(text, font_name='仿宋', size=16, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=True):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    if indent:
        p.paragraph_format.first_line_indent = Pt(32)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)  # 关键：中文 eastAsia 字体
    return p

def add_title(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    run.font.size = Pt(18)
    run.font.name = '黑体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

# ---- 标题 ----
add_title('<项目名>')
add_title('<报告名>')
doc.add_paragraph()

# ---- 引言（具名主体 + 现场复核）----
add_para('根据<报送主体全称>报送的<资料名>及相关资料，我们对<项目名>工程进度情况进行了现场复核，现将复核情况说明如下：')

# ---- 一、项目基本情况 ----
add_para('一、项目基本情况', font_name='楷体')
add_para('工程名称：<…>。')
add_para('工程地点：<…>。')
add_para('项目内容：<…>。')

# ---- 二、工程进度复核情况说明（已完成+未完成合并）----
add_para('二、工程进度复核情况说明', font_name='楷体')
add_para('经现场复核，本项目已完成的主要工程内容如下：')
add_para('1、<已完成项1>；')
add_para('2、<已完成项2>；')
add_para('3、<已完成项3>。')
add_para('4、<未完成项>尚未完成。')

# ---- 三、复核结论（含比例 + 现场证据）----
add_para('三、复核结论', font_name='楷体')
add_para('经复核，本项目实际完成情况与申报进度一致，已完成工程量属实，整体完成比例约<XX>%，未完成部分主要为<…>，建议按合同约定继续推进后续工作，完成后及时组织验收。（后附现场施工影像）')

# ---- 落款 ----
doc.add_paragraph()
doc.add_paragraph()
add_para('复核单位：<真实单位全称>', align=WD_ALIGN_PARAGRAPH.RIGHT, indent=False)
p_date = doc.add_paragraph()
p_date.alignment = WD_ALIGN_PARAGRAPH.LEFT
p_date.paragraph_format.left_indent = Pt(224)  # 用户习惯：左缩进靠右
run = p_date.add_run('<YYYY年MM月DD日>')
run.font.size = Pt(16)
run.font.name = '仿宋'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
doc.add_paragraph()

doc.save(r'<输出路径>.docx')
print('done')
```

## 校验清单（生成后必查）

- [ ] 页边距为 上下2.54 / 左右3.17cm（**现读母本实测复核**，不凭记忆）
- [ ] 标题黑体18居中；一级标题楷体16；正文仿宋16
- [ ] 正文首行缩进 32pt、两端对齐
- [ ] 引言含报送主体全称 + "现场复核"
- [ ] 完成比例在"复核结论"章
- [ ] 结论含"（后附现场施工影像）"
- [ ] 落款为真实单位全称、日期填实
- [ ] 页数 1–2 页
