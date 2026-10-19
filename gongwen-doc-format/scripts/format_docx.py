# -*- coding: utf-8 -*-
"""
公文 / 长文档套版 —— python-docx 直接格式化（幂等，可重复运行）。

用法：
    1) 修改下方 CONFIG；
    2) python format_docx.py
    或 import 后调用 apply_one(path, CONFIG)

规格基准见 SKILL.md §一 / §二。所有字体均为 run 级直设（eastAsia/ascii/hAnsi）。
"""
import os, re, sys

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================== 配置 ==============================
CONFIG = dict(
    # 目标文档（原地覆盖，请先自行备份）
    docs=[
        # r"C:\Users\lihaoyuan\Desktop\示例（送审稿）.docx",
    ],
    # 页面（cm）—— 通用公文母本实测值
    margins=dict(top=2.54, bottom=2.54, left=3.17, right=3.17,
                 header=1.50, footer=1.75),
    page=(21.0, 29.7),          # A4
    # 行距：None = 不动；"single" = 单倍；1.5 = 1.5 倍（长文档推荐）
    line_spacing=1.5,
    # 首行缩进（pt），32pt ≈ 2 字
    first_line_indent=32,
    # 字体规格 (中文字体, 字号pt, 加粗; None=保留原加粗)
    spec=dict(
        cover=("黑体", 18, False),   # 封面主标题
        toc_head=("黑体", 18, False),  # "目 录"
        l1=("黑体", 16, False),      # 第X章
        l2=("楷体", 16, False),      # X.Y
        l3=("仿宋", 16, True),       # X.Y.Z / （一）
        cap=("仿宋", 14, False),     # 表N—N
        body=("仿宋", 16, None),     # 正文
    ),
    # 目录条目样式 (字体, 字号, 加粗, 左缩进pt)
    toc_style={
        "TOC1": ("黑体", 14, True, 0),
        "TOC2": ("楷体", 14, False, 28),
        "TOC3": ("仿宋", 14, False, 56),
    },
    keep_next_headings=True,   # 标题设"与下段同页"（防落单）
    fit_tables=True,           # 表格列宽等比缩放到版心
)

# ============================== 层级识别 ==============================
RE_L1 = re.compile(r'^第[一二三四五六七八九十]{1,3}章[\s\u3000]')
RE_L3N = re.compile(r'^\d{1,2}\.\d{1,2}\.\d{1,2}')          # X.Y.Z
RE_L2 = re.compile(r'^\d{1,2}\.\d{1,2}[\s\u3000]')          # X.Y
RE_L3C = re.compile(r'^（[一二三四五六七八九十]{1,3}）')      # （一）
RE_CAP = re.compile(r'^表\s?\d+\s?[—\-–]\s?\d+')            # 表N—N
RE_TOCHEAD = re.compile(r'^目[\s\u3000]*录$')
HEAD_MAXLEN = 60   # 超长则视为正文（防误吞长句）


# ============================== 工具 ==============================
def set_font(p, cn, size, bold=None):
    """逐 run 直设中英文字体（中文字体必须写 w:eastAsia）。"""
    for r in p.runs:
        r.font.name = cn
        r.font.size = Pt(size)
        if bold is not None:
            r.font.bold = bold
        rpr = r._element.get_or_add_rPr()
        rf = rpr.find(qn('w:rFonts'))
        if rf is None:
            rf = OxmlElement('w:rFonts'); rpr.insert(0, rf)
        for a in ('w:eastAsia', 'w:ascii', 'w:hAnsi'):
            rf.set(qn(a), cn)


def set_multiple(ppr, mult):
    """行距设为自动多倍（1.5 -> line=360 lineRule=auto）。"""
    sp = ppr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing'); ppr.append(sp)
    sp.set(qn('w:line'), str(int(round(mult * 240))))
    sp.set(qn('w:lineRule'), 'auto')


def set_single(ppr):
    sp = ppr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing'); ppr.append(sp)
    sp.set(qn('w:line'), '240'); sp.set(qn('w:lineRule'), 'auto')


def set_keep_next(ppr, on=True):
    """标题'与下段同页'。注意：<w:keepNext/> 无 val 即默认生效。"""
    k = ppr.find(qn('w:keepNext'))
    if k is None:
        k = OxmlElement('w:keepNext'); ppr.insert(0, k)
    k.set(qn('w:val'), '1' if on else '0')


def set_outline(p, lvl):
    """重建大纲级别（1级->0 / 2级->1 / 3级->2）。"""
    ppr = p._element.get_or_add_pPr()
    old = ppr.find(qn('w:outlineLvl'))
    if old is not None:
        ppr.remove(old)
    if lvl is None:
        return
    o = OxmlElement('w:outlineLvl'); o.set(qn('w:val'), str(lvl))
    ppr.insert_element_before(o, 'w:divId', 'w:cnfStyle', 'w:rPr',
                              'w:sectPr', 'w:pPrChange')


def pstyle_id(p):
    ppr = p._element.find(qn('w:pPr'))
    if ppr is None:
        return None
    ps = ppr.find(qn('w:pStyle'))
    return ps.get(qn('w:val')) if ps is not None else None


def clear_outline(p):
    """清掉大纲级别。空段落若带 outlineLvl（历史遗留），会污染大纲与目录。"""
    ppr = p._element.find(qn('w:pPr'))
    if ppr is None:
        return
    old = ppr.find(qn('w:outlineLvl'))
    if old is not None:
        ppr.remove(old)


def toc_region(doc):
    """返回 TOC 域结果覆盖的段落索引集合（含 begin 段）。"""
    toks = []
    for i, p in enumerate(doc.paragraphs):
        for el in p._element.iter():
            if el.tag == qn('w:fldChar'):
                toks.append(('f', i, el.get(qn('w:fldCharType'))))
            elif el.tag == qn('w:instrText'):
                toks.append(('i', i, (el.text or '').strip()))
    stack, region = [], set()
    for kind, i, v in toks:
        if kind == 'f':
            if v == 'begin':
                stack.append({'instr': None, 'start': i})
            elif v == 'end' and stack:
                f = stack.pop()
                if f['instr'] and 'TOC' in f['instr']:
                    region.update(range(f['start'], i + 1))
        else:
            if stack and stack[-1]['instr'] is None:
                stack[-1]['instr'] = v
    return region


def upsert_style(doc, sid, name, cn, size, bold, left_pt, mult):
    """重建 TOC 样式（关键：Word 更新域后 TOC 条目会回到样式，样式必须设对）。"""
    styles = doc.styles.element
    el = None
    for st in styles.findall(qn('w:style')):
        if st.get(qn('w:styleId')) == sid:
            el = st; break
    if el is None:
        el = OxmlElement('w:style')
        el.set(qn('w:type'), 'paragraph'); el.set(qn('w:styleId'), sid)
        n = OxmlElement('w:name'); n.set(qn('w:val'), name); el.append(n)
        b = OxmlElement('w:basedOn'); b.set(qn('w:val'), 'Normal'); el.append(b)
        styles.append(el)
    for tag in ('w:pPr', 'w:rPr'):
        old = el.find(qn(tag))
        if old is not None:
            el.remove(old)
    ppr = OxmlElement('w:pPr')
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:line'), str(int(round((mult or 1) * 240))))
    sp.set(qn('w:lineRule'), 'auto')
    sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '0'); ppr.append(sp)
    ind = OxmlElement('w:ind')
    ind.set(qn('w:left'), str(int(left_pt * 20)))
    ind.set(qn('w:firstLine'), '0'); ppr.append(ind)
    jc = OxmlElement('w:jc'); jc.set(qn('w:val'), 'left'); ppr.append(jc)
    el.append(ppr)
    rpr = OxmlElement('w:rPr')
    rf = OxmlElement('w:rFonts')
    for a in ('w:ascii', 'w:eastAsia', 'w:hAnsi'):
        rf.set(qn(a), cn)
    rpr.append(rf)
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(int(size * 2))); rpr.append(sz)
    szcs = OxmlElement('w:szCs'); szcs.set(qn('w:val'), str(int(size * 2))); rpr.append(szcs)
    if bold:
        rpr.append(OxmlElement('w:b'))
    el.append(rpr)


def fit_tables(doc, avail_cm):
    """把每张表的列宽等比缩放到版心宽度，避免溢出。返回处理张数。"""
    n = 0
    for tb in doc.tables:
        grid = tb._tbl.find(qn('w:tblGrid'))
        if grid is None:
            continue
        cols = grid.findall(qn('w:gridCol'))
        widths = [int(c.get(qn('w:w')) or 0) for c in cols]
        total = sum(widths)
        if total <= 0:
            continue
        factor = (avail_cm * 567.0) / total
        new = [max(1, int(round(w * factor))) for w in widths]
        for c, w in zip(cols, new):
            c.set(qn('w:w'), str(w))
        for tr in tb._tbl.findall(qn('w:tr')):
            gi = 0
            for tc in tr.findall(qn('w:tc')):
                tcPr = tc.find(qn('w:tcPr'))
                span = 1
                if tcPr is not None:
                    gs = tcPr.find(qn('w:gridSpan'))
                    if gs is not None:
                        try:
                            span = int(gs.get(qn('w:val')))
                        except Exception:
                            span = 1
                wsum = sum(new[gi:gi + span]) if gi < len(new) else 0
                gi += span
                if tcPr is None:
                    tcPr = OxmlElement('w:tcPr'); tc.insert(0, tcPr)
                tcW = tcPr.find(qn('w:tcW'))
                if tcW is None:
                    tcW = OxmlElement('w:tcW'); tcPr.insert(0, tcW)
                tcW.set(qn('w:w'), str(wsum)); tcW.set(qn('w:type'), 'dxa')
        tblPr = tb._tbl.find(qn('w:tblPr'))
        if tblPr is not None:
            tblW = tblPr.find(qn('w:tblW'))
            if tblW is None:
                tblW = OxmlElement('w:tblW'); tblPr.append(tblW)
            tblW.set(qn('w:w'), str(int(avail_cm * 567))); tblW.set(qn('w:type'), 'dxa')
            lay = tblPr.find(qn('w:tblLayout'))
            if lay is None:
                lay = OxmlElement('w:tblLayout'); tblPr.append(lay)
            lay.set(qn('w:type'), 'fixed')
        n += 1
    return n


# ============================== 主流程 ==============================
def apply_one(path, cfg=CONFIG, log=print):
    doc = Document(path)
    ps = doc.paragraphs
    spec, toc_spec = cfg['spec'], cfg['toc_style']
    mult = cfg['line_spacing']
    indent = cfg['first_line_indent']

    # 1) 页面
    for s in doc.sections:
        m = cfg['margins']
        s.page_width, s.page_height = Cm(cfg['page'][0]), Cm(cfg['page'][1])
        s.top_margin, s.bottom_margin = Cm(m['top']), Cm(m['bottom'])
        s.left_margin, s.right_margin = Cm(m['left']), Cm(m['right'])
        s.header_distance, s.footer_distance = Cm(m['header']), Cm(m['footer'])
    log("  页面：%.2f/%.2f/%.2f/%.2f cm（%d 节）"
        % (cfg['margins']['top'], cfg['margins']['bottom'],
           cfg['margins']['left'], cfg['margins']['right'], len(doc.sections)))

    toc_idx = toc_region(doc)
    avail = cfg['page'][0] - cfg['margins']['left'] * 2
    stat = dict(cover=0, toch=0, L1=0, L2=0, L3=0, CAP=0, BODY=0, TOC=0)

    # 封面主标题 = 首个非空段落（通常为文档标题）
    first_nonempty = next((i for i, p in enumerate(ps) if p.text.strip()), None)

    for i, p in enumerate(ps):
        t = p.text.strip()
        ppr = p._element.get_or_add_pPr()

        # 行距（全段，含空段）
        if mult:
            set_multiple(ppr, mult)
        elif cfg['line_spacing'] == 'single':
            set_single(ppr)

        if not t:
            clear_outline(p)      # 空段落不得带大纲级别（否则污染目录）
            continue

        head = False
        if pstyle_id(p) and (pstyle_id(p) or '').upper().startswith('TOC'):
            sid = (pstyle_id(p) or 'TOC3').upper()
            cn, sz, bd, left = toc_spec.get(sid, toc_spec['TOC3'])
            set_font(p, cn, sz, bd)
            pf = p.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.first_line_indent = Pt(0); pf.left_indent = Pt(left)
            stat['TOC'] += 1
            continue
        if i in toc_idx:
            stat['TOC'] += 1
            continue

        if i == first_nonempty:
            cn, sz, bd = spec['cover']
            set_font(p, cn, sz, bd)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            stat['cover'] += 1; head = True
        elif RE_TOCHEAD.match(t):
            cn, sz, bd = spec['toc_head']
            set_font(p, cn, sz, bd)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            stat['toch'] += 1; head = True
        elif len(t) <= HEAD_MAXLEN and RE_L1.match(t):
            cn, sz, bd = spec['l1']; set_font(p, cn, sz, bd)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Pt(indent)
            set_outline(p, 0); stat['L1'] += 1; head = True
        elif len(t) <= HEAD_MAXLEN and RE_L3N.match(t):
            cn, sz, bd = spec['l3']; set_font(p, cn, sz, bd)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Pt(indent)
            set_outline(p, 2); stat['L3'] += 1; head = True
        elif len(t) <= HEAD_MAXLEN and RE_L2.match(t):
            cn, sz, bd = spec['l2']; set_font(p, cn, sz, bd)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Pt(indent)
            set_outline(p, 1); stat['L2'] += 1; head = True
        elif len(t) <= HEAD_MAXLEN and RE_L3C.match(t) and '。' not in t:
            cn, sz, bd = spec['l3']; set_font(p, cn, sz, bd)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Pt(indent)
            set_outline(p, 2); stat['L3'] += 1; head = True
        elif len(t) <= 30 and RE_CAP.match(t):
            cn, sz, bd = spec['cap']; set_font(p, cn, sz, bd)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_outline(p, None); stat['CAP'] += 1; head = True
        else:
            cn, sz, bd = spec['body']; set_font(p, cn, sz, bd)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Pt(indent)
            set_outline(p, None); stat['BODY'] += 1

        if head and cfg['keep_next_headings']:
            set_keep_next(ppr, True)

    log("  段落：封面%d 目录标题%d 一级%d 二级%d 三级%d 表题%d 正文%d 目录条目%d"
        % (stat['cover'], stat['toch'], stat['L1'], stat['L2'], stat['L3'],
           stat['CAP'], stat['BODY'], stat['TOC']))

    # 2) TOC 样式（行距与字体必须设，Word 更新域后条目回到样式）
    for sid, (cn, sz, bd, left) in toc_spec.items():
        upsert_style(doc, sid, sid.lower(), cn, sz, bd, left, mult or 1)
    log("  TOC1/2/3 样式已重建")

    # 3) 表格列宽
    if cfg['fit_tables']:
        n = fit_tables(doc, avail)
        log("  表格 %d 张，列宽已缩放至版心 %.2fcm" % (n, avail))

    doc.save(path)
    log("  已保存：%s" % path)


if __name__ == '__main__':
    if not CONFIG['docs']:
        print("请先在 CONFIG['docs'] 中填入目标文档路径。")
        sys.exit(1)
    for path in CONFIG['docs']:
        print("=" * 72); print(os.path.basename(path))
        apply_one(path)
    print("\n完成。下一步请运行 update_toc_fields.py 更新目录域与页码。")
