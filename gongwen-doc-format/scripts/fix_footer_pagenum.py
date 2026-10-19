# -*- coding: utf-8 -*-
"""
把「浮动文本框页脚」的页码修复为普通段落 PAGE 域 —— 附：页脚/页码相关工具集。

背景（2026-09-21 某智慧停车报告实测）：
  页脚页码若被放在浮动文本框（<mc:AlternateContent><wps:txbx>）内，
  会出现**页码不逐页刷新**——本案例正文第 10 页起全部显示「1」。
  强制 Fields.Update() 无效。改回普通段落内的 PAGE 域后恢复正常。

用法：
    python fix_footer_pagenum.py <docx> [--section 2] [--font 仿宋] [--size 12]
                                 [--blank-section 1] [--out 报告.txt]

    --section       要写入页码的节序号（1-based），默认 2（即"正文节"）
    --blank-section 要清空页脚的节序号，默认 1（目录/封面节不显示页码）
    --font/--size   页码字体与字号，默认 仿宋 12pt

注意：
  * 本脚本只改页脚，不动正文。执行后仍需用 Word COM 更新域并**逐页核对页码**。
  * 清空前**不要用 p.runs** —— 域内容与 w:drawing 都不在 runs 里。
"""
import os
import re
import sys
import shutil
import datetime

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'


def mk_run(font, size_pt):
    r = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    rf = OxmlElement('w:rFonts')
    for a in ('w:eastAsia', 'w:ascii', 'w:hAnsi'):
        rf.set(qn(a), font)
    rpr.append(rf)
    for tag in ('w:sz', 'w:szCs'):
        el = OxmlElement(tag)
        el.set(qn('w:val'), str(int(size_pt * 2)))
        rpr.append(el)
    r.append(rpr)
    return r


def add_page_field(p, font="仿宋", size_pt=12, cached="1"):
    """在段落末尾插入 { PAGE \\* MERGEFORMAT } 域（含缓存值）"""
    def fld(t):
        r = mk_run(font, size_pt)
        e = OxmlElement('w:fldChar')
        e.set(qn('w:fldCharType'), t)
        r.append(e)
        p._element.append(r)
    fld('begin')
    r = mk_run(font, size_pt)
    it = OxmlElement('w:instrText')
    it.set(XMLSPACE, 'preserve')
    it.text = ' PAGE  \\* MERGEFORMAT '
    r.append(it)
    p._element.append(r)
    fld('separate')
    r = mk_run(font, size_pt)
    t = OxmlElement('w:t')
    t.text = cached
    r.append(t)
    p._element.append(r)
    fld('end')


def clear_para(p):
    """清空段落除 pPr 外的全部子元素（含浮动文本框 drawing / 域）。返回删除数。"""
    n = 0
    for child in list(p._element):
        if child.tag != qn('w:pPr'):
            p._element.remove(child)
            n += 1
    return n


def fix(path, sec_no=2, blank_sec_no=1, font="仿宋", size_pt=12, log=print):
    out = []
    def w(s=""): out.append(str(s))

    d = Document(path)
    w("节数：%d" % len(d.sections))

    # 目标节：写页码
    idx = sec_no - 1
    if not (0 <= idx < len(d.sections)):
        raise ValueError("节序号 %d 超出范围（共 %d 节）" % (sec_no, len(d.sections)))
    p = d.sections[idx].footer.paragraphs[0]
    had = 'txbxContent' in p._element.xml
    n = clear_para(p)
    add_page_field(p, font, size_pt)
    pf = p.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    w("节%d 页脚：原含浮动文本框=%s，清除 %d 个子元素，已插入 PAGE 域（%s %gpt、居中）"
      % (sec_no, had, n, font, size_pt))

    # 清空节：不显示页码
    if blank_sec_no:
        bi = blank_sec_no - 1
        if 0 <= bi < len(d.sections) and bi != idx:
            nb = clear_para(d.sections[bi].footer.paragraphs[0])
            w("节%d 页脚：已清空（清除 %d 个子元素）" % (blank_sec_no, nb))
        elif bi == idx:
            w("⚠️ --blank-section 与 --section 相同，已跳过")

    d.save(path)
    w("✅ 已保存")
    w("")
    w("下一步：Word COM 更新域 → 导出 PDF → **逐页**核对页码（见 update_toc_fields.py）")
    return "\n".join(out)


if __name__ == "__main__":
    argv = sys.argv[1:]
    if not argv:
        print(__doc__)
        sys.exit(1)
    doc = argv[0]
    kw = {}
    for flag, key, cast in (("--section", "sec_no", int),
                            ("--blank-section", "blank_sec_no", int),
                            ("--font", "font", str),
                            ("--size", "size_pt", float)):
        if flag in argv:
            i = argv.index(flag)
            kw[key] = cast(argv[i + 1])
            del argv[i:i + 2]

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = os.path.splitext(doc)[0] + ".页码修复前.%s.bak.docx" % ts
    shutil.copy2(doc, bak)
    txt = fix(doc, **kw)
    print(txt)
    print("\n备份：%s" % bak)
