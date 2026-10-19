# -*- coding: utf-8 -*-
"""
给 docx 加/修目录（TOC）—— 一次做对，含四个必踩的坑。

用法：
    python build_toc.py <docx路径> [--levels 1-2] [--h1-font 黑体] [--h2-font 楷体]
                         [--size 14] [--page-break]

背景 / 为何要这么写（详见 SKILL.md §四 坑 11）：
  * Word 实际使用**小写**内置样式 `toc 1` / `toc 2`，用 python-docx 建大写 `TOC 1` 不生效。
  * 目录条目的蓝色下划线来自 `Hyperlink` **字符样式**，只改 TOC 段落样式无效；
    且 `Hyperlink` 是 CHARACTER 型样式，**没有 `.paragraph_format`**。
  * 样式重名会静默生成 `TOC 11` / `TOC 21` 副本，须清理。
  * 只改样式压不住行距，必须逐条写段落级 `w:spacing`（固定行距 exact）。

流程：
    ① 逐条清理 TOC 条目上的残留字符格式（color/u/shd）并重设字体
    ② 逐条写段落级固定行距（h1=19pt, h2=18pt），段前后 0
    ③ 同步小写 `toc 1` / `toc 2` 样式（防 Word 重生成时回弹）
    ④ 修 `Hyperlink` 字符样式 → 目标字体 / 黑色 / 无下划线
    ⑤ 清掉误建的大写 TOC 样式副本
    ⑥ 可选：主标题后加分页符，让目录独立起页
完成后**必须**调用 update_toc_fields.py 用 Word 更新域生成真实页码。
"""
import os, re, sys, shutil, datetime
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

H1_RE = re.compile(r'^[一二三四五六七八九十]+、')
H2_RE = re.compile(r'^（[一二三四五六七八九十]+）')


# ───────────────────────── 工具函数 ─────────────────────────

def force_spacing(p, line_pt, before_pt=0, after_pt=0):
    """直接写 w:spacing，行距=固定值（exact）。这是压住目录行距的唯一可靠办法。"""
    ppr = p._element.get_or_add_pPr()
    sp = ppr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing')
        ppr.append(sp)
    sp.set(qn('w:line'), str(int(round(line_pt * 20))))   # twips
    sp.set(qn('w:lineRule'), 'exact')
    sp.set(qn('w:before'), str(int(round(before_pt * 20))))
    sp.set(qn('w:after'), str(int(round(after_pt * 20))))
    sp.set(qn('w:beforeAutospacing'), '0')
    sp.set(qn('w:afterAutospacing'), '0')


def clear_run_fmt(run):
    """清掉 run 上残留的颜色/下划线/底纹/字距。"""
    rpr = run._element.find(qn('w:rPr'))
    if rpr is None:
        return
    for tag in ('w:color', 'w:u', 'w:highlight', 'w:shd', 'w:spacing'):
        el = rpr.find(qn(tag))
        if el is not None:
            rpr.remove(el)


def set_run_font(run, font, size, bold=False):
    """中文字体必须设 w:eastAsia，否则无效。"""
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.underline = False
    run.font.color.rgb = RGBColor(0, 0, 0)
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rpr.insert(0, rf)
    for a in ('w:eastAsia', 'w:ascii', 'w:hAnsi'):
        rf.set(qn(a), font)


def fix_hyperlink_style(doc, font, size):
    """把 Hyperlink 字符样式改为黑色无下划线 —— 目录蓝色下划线的真正来源。"""
    try:
        st = doc.styles["Hyperlink"]
    except KeyError:
        return False
    st.font.name = font
    st.font.size = Pt(size)
    st.font.underline = False
    st.font.color.rgb = RGBColor(0, 0, 0)
    rpr = st.element.get_or_add_rPr()
    for tag in ('w:color', 'w:u'):
        el = rpr.find(qn(tag))
        if el is not None:
            rpr.remove(el)
    cel = OxmlElement('w:color'); cel.set(qn('w:val'), '000000'); rpr.append(cel)
    uel = OxmlElement('w:u'); uel.set(qn('w:val'), 'none'); rpr.append(uel)
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts'); rpr.insert(0, rf)
    for a in ('w:eastAsia', 'w:ascii', 'w:hAnsi'):
        rf.set(qn(a), font)
    return True


def set_toc_style(doc, name, font, size, left_cm, line_pt):
    """同步小写内置样式。"""
    try:
        st = doc.styles[name]
    except KeyError:
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    st.font.name = font
    st.font.size = Pt(size)
    st.font.bold = False
    st.font.underline = False
    st.font.color.rgb = RGBColor(0, 0, 0)
    st.paragraph_format.line_spacing = Pt(line_pt)     # 固定值
    st.paragraph_format.left_indent = Cm(left_cm)
    st.paragraph_format.first_line_indent = Pt(0)
    st.paragraph_format.space_before = Pt(0)
    st.paragraph_format.space_after = Pt(0)
    return st


def drop_style(doc, name):
    """删除误建的样式副本（如 TOC 11 / TOC 21）。"""
    try:
        st = doc.styles[name]
    except KeyError:
        return False
    st.element.getparent().remove(st.element)
    return True


# ───────────────────────── 主流程 ─────────────────────────

def build_toc(doc_path, h1_font="黑体", h2_font="楷体", size=14, page_break=False):
    out = []
    def w(s=""): out.append(str(s))

    d = Document(doc_path)
    w("=" * 74); w("目录构建 / 修复"); w("=" * 74)

    # ①~③ 逐条处理 TOC 条目
    n1 = n2 = 0
    for p in d.paragraphs:
        sn = (p.style.name if p.style else "") or ""
        if not sn.lower().startswith("toc"):
            continue
        t = p.text.strip()
        if H1_RE.match(t):
            font, indent_cm, line_pt = h1_font, 0.0, 19
            n1 += 1
        elif H2_RE.match(t):
            font, indent_cm, line_pt = h2_font, 0.74, 18
            n2 += 1
        else:
            continue
        force_spacing(p, line_pt, 0, 0)
        p.paragraph_format.left_indent = Cm(indent_cm)
        p.paragraph_format.first_line_indent = Pt(0)
        for r in p.runs:
            clear_run_fmt(r)
            set_run_font(r, font, size, bold=False)

    w("① 一级条目 %d 条：%s%d / 固定行距19pt" % (n1, h1_font, size))
    w("   二级条目 %d 条：%s%d / 固定行距18pt / 左缩进0.74cm" % (n2, h2_font, size))

    # ③ 样式同步
    for nm, font, indent, lp in (("toc 1", h1_font, 0.0, 19), ("toc 2", h2_font, 0.74, 18)):
        set_toc_style(d, nm, font, size, indent, lp)
        w("② 样式同步：%s → %s%d / 固定行距%dpt" % (nm, font, size, lp))

    # ④ Hyperlink 字符样式
    if fix_hyperlink_style(d, "仿宋", size):
        w("③ Hyperlink 字符样式 → 仿宋%d / 黑色 / 无下划线" % size)
    else:
        w("③ 无 Hyperlink 样式，跳过")

    # ⑤ 清理误建样式副本
    dropped = [n for n in ("TOC 1", "TOC 2", "TOC 11", "TOC 21", "TOC 12", "TOC 22")
               if drop_style(d, n)]
    if dropped:
        w("④ 已删除误建样式：%s" % "、".join(dropped))

    # ⑥ 主标题后分页（目录独立起页）
    if page_break:
        idx = None
        for i, p in enumerate(d.paragraphs):
            if p.text.strip() == "投资风险评估报告" or (p.style and p.style.name == "Title"):
                idx = i
        if idx is not None:
            tp = d.paragraphs[idx]
            brk = d.add_paragraph()
            r = brk.add_run()
            set_run_font(r, "仿宋", 16)
            brk.paragraph_format.line_spacing = 1.5
            r.add_break(WD_BREAK.PAGE)
            brk._element.getparent().remove(brk._element)
            tp._element.addnext(brk._element)
            w("⑤ 已在主标题后插入分页符（目录独立起页）")

    d.save(doc_path)
    w("✅ 已保存")
    w("")
    w("下一步：python update_toc_fields.py <docx> --pdf <目录>   （Word 更新域生成真实页码）")
    return "\n".join(out)


if __name__ == "__main__":
    argv = sys.argv[1:]
    if not argv:
        print(__doc__); sys.exit(1)
    doc = argv[0]
    do_break = "--page-break" in argv
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = os.path.splitext(doc)[0] + ".目录处理前.%s.bak.docx" % ts
    shutil.copy2(doc, bak)
    print(build_toc(doc, page_break=do_break))
    print("\n备份：%s" % bak)
