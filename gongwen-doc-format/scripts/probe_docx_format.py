# -*- coding: utf-8 -*-
"""
docx 格式全量探测 / 与参照版差异定位 —— 用于"学习用户手改后的格式"。

用法：
    # 仅探测当前格式
    python probe_docx_format.py <docx>

    # 与参照版对比（定位用户改了什么）
    python probe_docx_format.py <docx> --ref <参照docx> --offset 1

参数：
    --ref     参照版本（例如我方上一版交付）
    --offset  当前版段落索引 = 参照版索引 + offset（结构变动时用，默认 0 表示自动按文本匹配）
    --out     输出 txt 路径（默认 <docx同目录>/格式探测.txt）

为什么要单独写脚本：
  * 本机 PowerShell 工具不回显 stdout → 一律 print() + 写 .txt
  * "用户说改了格式"时，必须**分区定位**：页面 / 样式 / 逐段 / 分节 / 页码域，
    不能只看肉眼可感的字体字号。实测中用户改动常落在：
    分节符与页码重编、目录条目行距（仅段落未改样式）、序号顿号化、引号全角化。

产出：一份 txt，含
  1) 页面与分节（含 pgNumType / 页眉页脚域）
  2) 关键样式定义（含 toc 1 / toc 2 / Hyperlink）
  3) 逐段明细（样式/大纲级别/字体/字号/行距/缩进/段前/keepNext/文本）
  4) 分组统计
  5) 若给 --ref：逐段格式差异 + 文本差异
"""
import os
import re
import sys
import difflib

from docx import Document
from docx.oxml.ns import qn

H1_RE = re.compile(r'^[一二三四五六七八九十]+、')
H2_RE = re.compile(r'^（[一二三四五六七八九十]+）')


def q(elm, tag):
    return elm.find(qn(tag)) if elm is not None else None


def ppr_val(p, tag):
    ppr = p._element.find(qn('w:pPr'))
    el = q(ppr, tag)
    return el.get(qn('w:val')) if el is not None else None


def ppr_exists(p, tag):
    """keepNext 常不带 val，必须判元素存在性"""
    ppr = p._element.find(qn('w:pPr'))
    return q(ppr, tag) is not None


def run_font(p):
    """取第一个有文本 run 的字体/字号/粗体；域内容读不到时返回 None（从样式继承）"""
    for r in p.runs:
        if r.text.strip():
            rpr = r._element.find(qn('w:rPr'))
            ea = sz = None
            if rpr is not None:
                rf = q(rpr, 'w:rFonts')
                if rf is not None:
                    ea = rf.get(qn('w:eastAsia'))
                szel = q(rpr, 'w:sz')
                if szel is not None:
                    sz = int(szel.get(qn('w:val'))) / 2
            return (ea or r.font.name,
                    sz if sz is not None else (r.font.size.pt if r.font.size else None),
                    r.font.bold)
    return (None, None, None)


def facts(p):
    pf = p.paragraph_format
    ls = pf.line_spacing
    if ls is None:
        ls_s = "None"
    elif isinstance(ls, float):
        ls_s = "%.2fx" % ls
    else:
        ls_s = "%.1fpt" % ls.pt
    ea, sz, _ = run_font(p)
    return {
        'style': (p.style.name if p.style else ""),
        'lvl': ppr_val(p, 'w:outlineLvl') or "-",
        'font': ea or "-",
        'size': ("%g" % sz) if sz else "-",
        'line': ls_s,
        'first': ("%.1f" % pf.first_line_indent.pt) if pf.first_line_indent is not None else "-",
        'left': ("%.2f" % pf.left_indent.cm) if pf.left_indent is not None else "-",
        'before': ("%.1f" % pf.space_before.pt) if pf.space_before is not None else "-",
        'after': ("%.1f" % pf.space_after.pt) if pf.space_after is not None else "-",
        'keep': "Y" if ppr_exists(p, 'w:keepNext') else "-",
        'align': str(pf.alignment).split(' ')[0] if pf.alignment is not None else "-",
    }


KEYS = ('style', 'lvl', 'font', 'size', 'line', 'first', 'left', 'before', 'after', 'keep', 'align')


def probe(path, ref_path=None, offset=None):
    out = []
    def w(s=""): out.append(str(s))

    d = Document(path)

    # ── 1. 页面与分节
    w("=" * 96)
    w("一、页面与分节")
    w("=" * 96)
    for i, s in enumerate(d.sections):
        w("节 %d  纸张 %.2f×%.2f cm  边距 上%.3f 下%.3f 左%.3f 右%.3f cm  页眉%.3f 页脚%.3f"
          % (i, s.page_width.cm, s.page_height.cm, s.top_margin.cm, s.bottom_margin.cm,
             s.left_margin.cm, s.right_margin.cm, s.header_distance.cm, s.footer_distance.cm))
        t = q(s._sectPr, 'w:type')
        w("     节起始=%s" % (t.get(qn('w:val')) if t is not None else "默认(nextPage)"))
        pn = q(s._sectPr, 'w:pgNumType')
        w("     页码 pgNumType: %s" % ("start=%s fmt=%s" % (pn.get(qn('w:start')), pn.get(qn('w:fmt')))
                                       if pn is not None else "未设（承接前节）"))
        for nm, part in (("页眉", s.header), ("页脚", s.footer)):
            info = []
            for p in part.paragraphs:
                if p.text.strip():
                    info.append(repr(p.text.strip()[:30]))
                flds = [it.text.strip() for it in p._element.iter(qn('w:instrText')) if it.text]
                if flds:
                    info.append("<域:%s>" % " | ".join(set(flds)))
                # 浮动文本框检测（页码陷阱）：python-docx 的 nsmap 无 wps 前缀，改用字符串判据
                if 'txbxContent' in p._element.xml:
                    info.append("⚠️含浮动文本框")
            w("     %s linked=%s %s" % (nm, part.is_linked_to_previous,
                                        " ".join(info) if info else "(空)"))

    # ── 2. 关键样式
    w("")
    w("=" * 96)
    w("二、关键样式定义")
    w("=" * 96)
    for nm in ("Normal", "toc 1", "toc 2", "TOC 1", "TOC 2", "Hyperlink",
               "Heading 1", "Heading 2"):
        try:
            st = d.styles[nm]
        except KeyError:
            w("  [%s] 不存在" % nm)
            continue
        rpr = st.element.find(qn('w:rPr'))
        ea = None
        if rpr is not None:
            rf = q(rpr, 'w:rFonts')
            ea = rf.get(qn('w:eastAsia')) if rf is not None else None
        line = "-"
        try:
            ls = st.paragraph_format.line_spacing
            line = ("%.1fpt" % ls.pt) if (ls is not None and not isinstance(ls, float)) else str(ls)
        except AttributeError:
            line = "(字符型样式，无段落格式)"
        w("  [%s] type=%s eastAsia=%s size=%s line=%s"
          % (nm, st.type, ea, st.font.size.pt if st.font.size else None, line))

    # ── 3. 逐段明细
    w("")
    w("=" * 96)
    w("三、逐段明细")
    w("=" * 96)
    for i, p in enumerate(d.paragraphs):
        t = p.text.strip()
        f = facts(p)
        brk = ""
        for r in p.runs:
            for b in r._element.findall('.//' + qn('w:br')):
                if b.get(qn('w:type')) == 'page':
                    brk = "[分页]"
        sect = "[分节]" if ppr_exists(p, 'w:sectPr') else ""
        w("%-4d %-10s lvl=%-3s %-5s %-4s %-8s 首%-7s 左%-6s 前%-5s %s %s %s%s"
          % (i, f['style'][:10], f['lvl'], f['font'], f['size'], f['line'],
             f['first'], f['left'], f['before'], f['keep'], f['align'],
             sect + brk, (t[:44] or "(空)")))

    # ── 4. 分组统计
    w("")
    w("=" * 96)
    w("四、分组统计（仅非空段）")
    w("=" * 96)
    from collections import Counter
    cs = {k: Counter() for k in ('font', 'size', 'line', 'first', 'align', 'style')}
    for p in d.paragraphs:
        if not p.text.strip():
            continue
        f = facts(p)
        for k in cs:
            cs[k][f[k]] += 1
    for k, c in cs.items():
        w("  【%s】%s" % (k, dict(c.most_common())))

    # ── 5. 与参照版对比
    if ref_path and os.path.exists(ref_path):
        w("")
        w("=" * 96)
        w("五、与参照版差异")
        w("=" * 96)
        d2 = Document(ref_path)
        P1, P2 = d.paragraphs, d2.paragraphs
        w("当前版 %d 段 / 参照版 %d 段" % (len(P1), len(P2)))

        # 自动推定偏移：用**长段落**作锚点
        # （比"首个非空段"稳健：目录条目短、且结构重排时首段必变）
        off = offset
        if off is None:
            off = 0
            anchors = sorted(
                [(len(p.text.strip()), i2, p.text.strip())
                 for i2, p in enumerate(P2) if len(p.text.strip()) >= 40],
                reverse=True)[:25]
            for _, i2, txt in anchors:
                hit = next((i1 for i1, p in enumerate(P1) if p.text.strip() == txt), None)
                if hit is not None:
                    off = hit - i2
                    w("自动推定偏移 = %d（锚点 参照#%d，段长%d）" % (off, i2, len(txt)))
                    break
            else:
                w("⚠️ 未找到匹配锚点，偏移按 0 处理；结构变动较大时请显式指定 --offset")

        # 目录条目（含制表符+页码）单独说明
        w("")
        w("【段落顺序变化】当前版前 8 段：")
        for i, p in enumerate(P1[:8]):
            t = p.text.strip()
            w("   %2d  %s" % (i, (t[:52] or "(空)")))
        w("【段落顺序变化】参照版前 8 段：")
        for i, p in enumerate(P2[:8]):
            t = p.text.strip()
            w("   %2d  %s" % (i, (t[:52] or "(空)")))

        w("")
        w("【逐段差异】")
        nd = nt = 0
        for i2 in range(len(P2)):
            i1 = i2 + off
            if i1 >= len(P1):
                break
            p1, p2 = P1[i1], P2[i2]
            t1, t2 = p1.text.strip(), p2.text.strip()
            f1, f2 = facts(p1), facts(p2)
            dfs = [(k, f2[k], f1[k]) for k in KEYS if f2[k] != f1[k]]
            if dfs:
                nd += 1
                w("  [格式] 参照#%d → 当前#%d  %s" % (i2, i1, (t1[:34] or "(空)")))
                for k, a, b in dfs:
                    w("        %-7s 参照=%-8s 当前=%s" % (k, a, b))
            if t1 != t2:
                nt += 1
                if nt <= 60:
                    sm = difflib.SequenceMatcher(None, t2, t1)
                    chg = []
                    for tag, a1, a2, b1, b2 in sm.get_opcodes():
                        if tag == 'replace':
                            chg.append("%r→%r" % (t2[a1:a2], t1[b1:b2]))
                        elif tag == 'delete':
                            chg.append("删%r" % t2[a1:a2])
                        elif tag == 'insert':
                            chg.append("增%r" % t1[b1:b2])
                    w("  [文本] 参照#%d → 当前#%d  %s" % (i2, i1, "  ".join(chg)[:110]))
        w("")
        w("格式有差异段数 %d ；文本有差异段数 %d" % (nd, nt))

        # 标点形态统计（序号/引号）
        w("")
        w("【标点形态】")
        pats = [("序号·顿号 1、", re.compile(r'^\d+、')),
                ("序号·半角点 1.", re.compile(r'^\d+\.[^0-9]')),
                ("序号·全角点 1．", re.compile(r'^\d+．'))]
        for label, rx in pats:
            c1 = sum(1 for p in P1 if rx.match(p.text.strip()))
            c2 = sum(1 for p in P2 if rx.match(p.text.strip()))
            w("   %-16s 当前=%-4d 参照=%d" % (label, c1, c2))
        q1 = sum(p.text.count('"') for p in P1)
        q2 = sum(p.text.count('"') for p in P2)
        w("   %-16s 当前=%-4d 参照=%d" % ("半角直引号 \"", q1, q2))

    return "\n".join(out)


if __name__ == "__main__":
    argv = sys.argv[1:]
    if not argv:
        print(__doc__)
        sys.exit(1)
    doc = argv[0]
    ref = None
    off = None
    outp = None
    for flag in ("--ref", "--offset", "--out"):
        if flag in argv:
            i = argv.index(flag)
            v = argv[i + 1] if i + 1 < len(argv) else None
            if flag == "--ref":
                ref = v
            elif flag == "--offset":
                off = int(v)
            else:
                outp = v
            del argv[i:i + 2]
    if outp is None:
        outp = os.path.join(os.path.dirname(doc), "格式探测.txt")
    try:
        txt = probe(doc, ref, off)
        with open(outp, "w", encoding="utf-8") as f:
            f.write(txt)
        print(txt[:1500])
        print("\n...\n完整结果 -> %s" % outp)
    except Exception:
        import traceback
        tb = traceback.format_exc()
        errp = outp + ".err.txt"
        with open(errp, "w", encoding="utf-8") as f:
            f.write(tb)
        print("FAILED -> %s" % errp)
        print(tb)
        sys.exit(1)
