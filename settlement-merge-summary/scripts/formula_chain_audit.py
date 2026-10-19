# -*- coding: utf-8 -*-
# [skill: settlement-merge-summary] 公式链展开（支持区域引用）+ 全表父行一致性审计
# 用法：改顶部 BASE/SRC/DST 路径常量后直接运行；输出写入 .txt 便于 Read 查看。
"""精确判定：r4 审定列公式链的覆盖范围（支持区域引用），找出未被覆盖的叶子行"""
import openpyxl, re, os, datetime
from openpyxl.utils import column_index_from_string as ci

BASE = r"<项目工作区>\审核过程资料"
SRC = os.path.join(BASE, "工程结算书8-28 - 副本.xlsx")
DST = os.path.join(BASE, "工程结算书合并汇总.xlsx")
OUT = r"D:\Workbuddy文件\2026-07-14-15-47-54\hb9_chain2.txt"
L = []
def P(s=""): L.append(str(s))

P("=== 文件时间戳 ===")
for p in [SRC, DST]:
    P("  %s\n     mtime=%s  size=%.1fKB" % (
        os.path.basename(p),
        datetime.datetime.fromtimestamp(os.path.getmtime(p)).strftime("%Y-%m-%d %H:%M:%S"),
        os.path.getsize(p)/1024))

wbf = openpyxl.load_workbook(SRC)
wbv = openpyxl.load_workbook(SRC, data_only=True)
wf, wsv = wbf["Sheet2"], wbv["Sheet2"]
MAXR = wsv.max_row

TOKEN = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?")
def refs_in(f):
    """返回 {行号: {列号,...}}"""
    if not isinstance(f, str) or not f.startswith("="): return {}
    out = {}
    for m in TOKEN.finditer(f):
        c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), m.group(4)
        if c2:
            rlo, rhi = sorted((r1, int(r2))); cs = sorted({ci(c1), ci(c2)})
            cols = set(range(cs[0], cs[-1]+1))
        else:
            rlo = rhi = r1; cols = {ci(c1)}
        for rr in range(rlo, rhi+1):
            out.setdefault(rr, set()).update(cols)
    return out

def chain(start, col):
    """沿 col 列公式展开，返回 (叶子行集, 已访问表头行集)"""
    seen, leafset = set(), set()
    stack = [start]
    while stack:
        r = stack.pop()
        if r in seen: continue
        seen.add(r)
        f = wf.cell(r, col).value
        rs = sorted([rr for rr, cs in refs_in(f).items() if col in cs])
        if rs:
            stack.extend(rs)
        else:
            leafset.add(r)
    return leafset, seen

leafH, seenH = chain(4, 8)      # 8 = H 列
leafG, seenG = chain(4, 7)

# 全部叶子行（单价数值 且 有量）
leaves = []
for r in range(4, MAXR+1):
    c4 = wsv.cell(r,4).value; c5 = wsv.cell(r,5).value; c6 = wsv.cell(r,6).value
    if wsv.cell(r,2).value is None: continue
    if isinstance(c6,(int,float)) and (isinstance(c4,(int,float)) or isinstance(c5,(int,float))):
        leaves.append(r)
allset = set(leaves)

P("")
P("=== 链条覆盖 ===")
P("  全表叶子行 = %d" % len(leaves))
P("  H链(审定) 覆盖叶子 = %d     G链(报送) 覆盖叶子 = %d" % (
    len(allset & leafH), len(allset & leafG)))

def rng(rs):
    rs = sorted(rs); out=[]; 
    for x in rs:
        if out and x == out[-1][1]+1: out[-1][1]=x
        else: out.append([x,x])
    return ",".join(str(a) if a==b else "%d-%d"%(a,b) for a,b in out)

missH = allset - leafH
missG = allset - leafG
P("")
P("=== 未被 H链 覆盖的叶子行: %d 个  → 行号 %s" % (len(missH), rng(missH)))
sm_h = 0.0
for r in sorted(missH):
    v7 = wsv.cell(r,7).value or 0.0; v8 = wsv.cell(r,8).value or 0.0
    sm_h += v8
    P("   r%-4d %-32s 单位=%-8s 单价=%-11s 报送额=%12.2f 审定额=%12.4f" % (
        r, str(wsv.cell(r,2).value)[:30], str(wsv.cell(r,3).value)[:8],
        wsv.cell(r,6).value, v7, v8))
P("   未被覆盖叶子 审定合计 = %.4f" % sm_h)
P("")
P("=== 未被 G链 覆盖的叶子行: %d 个  → 行号 %s" % (len(missG), rng(missG)))

# 反证：r4 父链口径 vs 明细口径
tot8 = sum((wsv.cell(r,8).value or 0.0) for r in leaves)
tot7 = sum((wsv.cell(r,7).value or 0.0) for r in leaves)
P("")
P("=== 两个口径 ===")
P("  明细口径（叶子求和）: 报送=%.4f  审定=%.4f" % (tot7, tot8))
P("  源表父链口径（r4）  : 报送=%s  审定=%s" % (wsv.cell(4,7).value, wsv.cell(4,8).value))
P("  差异（明细 - r4）   : 报送=%.4f  审定=%.4f" % (
    tot7 - (wsv.cell(4,7).value or 0), tot8 - (wsv.cell(4,8).value or 0)))

# 目标文件当前状态（我上轮写的值）
P("")
P("=== 目标 合并汇总 当前状态（上轮落盘）===")
wn = openpyxl.load_workbook(DST, data_only=True)["合并汇总"]
for r in range(85, 105):
    nm = wn.cell(r,2).value; c6 = wn.cell(r,6).value
    tag = "叶子" if isinstance(c6,(int,float)) else ("合计" if nm and "计" in str(nm) else "表头")
    P("   r%-3d %-6s %-30s 单价=%-10s 报送额=%-14s 审核额=%-14s" % (
        r, tag, str(nm)[:28], c6, wn.cell(r,7).value, wn.cell(r,8).value))

open(OUT,"w",encoding="utf-8").write("\n".join(L))
print("done ->", OUT)
