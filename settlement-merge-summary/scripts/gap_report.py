# -*- coding: utf-8 -*-
# [skill: settlement-merge-summary] 审定合计差异核查报告生成器（全自动）
# 用法：改 SRC / DST / REPORT 三个路径常量后直接运行。
#   源表 = 结算书（Sheet2，col4报送量 col5审定量 col6单价 col7报送额 col8审定额）
#   目标表 = 合并汇总（可选，仅用于读取 r102/r103 做交叉印证）
# 产出：[差异溯源] + [差异明细] + [两口径对照] 三张 sheet 的 xlsx，并输出 .txt 日志。
"""自动完成：
 ① 算两个口径：Σ(叶子行 col8) vs 源表项目级合计 → 差异 gap
 ② 对 G 列 / H 列分别做公式链展开（支持区域引用），求各自的"未被覆盖叶子行"集合
 ③ 差集 = H 未覆盖但 G 覆盖 → 「同处插入只改了一列」的**直接证据**
 ④ 自动抽取父行 G/H 公式原文做对照、统计该表"不予认可=审定量置0"的既有写法
 ⑤ 生成两口径对照（核减额差多少），供用户拍板
"""
import openpyxl, re, os
from decimal import Decimal
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ============ 配置区 ============
BASE   = r"<项目工作区>\审核过程资料"
SRC    = os.path.join(BASE, "工程结算书8-28 - 副本.xlsx")
DST    = os.path.join(BASE, "工程结算书合并汇总.xlsx")
REPORT = r"D:\Workbuddy文件\2026-07-14-15-47-54\审定合计差异核查报告.xlsx"
LOG    = os.path.splitext(REPORT)[0] + "_log.txt"
SRC_SHEET = "Sheet2"
# ===============================

L = []
def P(s=""): L.append(str(s))
D = lambda x: Decimal(repr(x)) if isinstance(x, (int, float)) else Decimal(0)

# ---------- 公式引用解析（必须支持 H48:H52 区域引用） ----------
TOKEN = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?")
def ci(col):
    n = 0
    for ch in col: n = n * 26 + (ord(ch) - 64)
    return n

def refs_in(f):
    """返回 {行号: {列号集合}}；区域引用展开为逐行"""
    if not isinstance(f, str) or not f.startswith("="): return {}
    out = {}
    for m in TOKEN.finditer(f):
        c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), m.group(4)
        if c2:
            rlo, rhi = sorted((r1, int(r2))); cs = sorted({ci(c1), ci(c2)})
            cols = set(range(cs[0], cs[-1] + 1))
        else:
            rlo = rhi = r1; cols = {ci(c1)}
        for rr in range(rlo, rhi + 1):
            out.setdefault(rr, set()).update(cols)
    return out

def chain(wf, ws_name, start_row, col):
    """沿公式链从 start_row 向下展开，返回 (叶子行集合, 访问过的全部行)"""
    wf_sheet = wf[ws_name]
    seen, leaves, stack = set(), set(), [start_row]
    while stack:
        r = stack.pop()
        if r in seen: continue
        seen.add(r)
        rs = sorted([rr for rr, cs in refs_in(wf_sheet.cell(r, col).value).items() if col in cs])
        if rs: stack.extend(rs)
        else: leaves.add(r)
    return leaves, seen

# ---------- 读源表 ----------
wv = openpyxl.load_workbook(SRC, data_only=True)
wf = openpyxl.load_workbook(SRC)
sv, sf = wv[SRC_SHEET], wf[SRC_SHEET]
maxr = sv.max_row

# 叶子行 = col6 有数值
leaf_rows = [r for r in range(4, maxr + 1)
             if isinstance(sv.cell(r, 6).value, (int, float)) and sv.cell(r, 2).value]
sum_h = sum((D(sv.cell(r, 8).value) for r in leaf_rows), Decimal(0))
sum_g = sum((D(sv.cell(r, 7).value) for r in leaf_rows), Decimal(0))

# 项目级合计行：取 col8 有公式且覆盖行数最多的顶层行（通常 r4）
top_row = 4
proj_h = D(sv.cell(top_row, 8).value); proj_g = D(sv.cell(top_row, 7).value)
gap = sum_h - proj_h
NEGLIGIBLE = abs(gap) < Decimal("0.01")   # 1 分钱以内视为自洽（浮点级尾差）

P("=" * 76)
P("审定合计差异核查（自动生成）  " + os.path.basename(SRC))
P("=" * 76)
P("")
P("【口径对比】")
P("  源表叶子行数（col6 有单价）      = %d" % len(leaf_rows))
P("  ① 明细口径 Σ(叶子 col8 审定合计) = %s" % sum_h)
P("  ② 合计口径 项目级 r%d col8       = %s" % (top_row, proj_h))
P("  差异 ①−②                        = %s" % gap)
P("  （报送侧：Σ叶子 col7 = %s ；项目级 r%d col7 = %s ；差 = %s）"
  % (sum_g, top_row, proj_g, sum_g - proj_g))
P("")

# ---------- 公式链展开：H 列 vs G 列 ----------
leaves_h, seen_h = chain(wf, SRC_SHEET, top_row, 8)
leaves_g, seen_g = chain(wf, SRC_SHEET, top_row, 7)
leaf_set = set(leaf_rows)

# 方向要分清："X 链覆盖" = 行在 leaves_X 里；"X 链漏了" = 行在 leaf_set 但不在 leaves_X 里
h_cov = leaf_set & leaves_h; g_cov = leaf_set & leaves_g
h_miss = leaf_set - leaves_h                      # H 链未覆盖 → 这些行的金额没进项目级审定合计
g_miss = leaf_set - leaves_g                      # G 链未覆盖
only_h_miss = h_miss & g_cov                      # ★ H 漏、G 有 = "同处插入只改一列"的铁证
only_g_miss = g_miss & h_cov                      # G 漏、H 有
both_miss   = h_miss & g_miss                     # 两侧都漏

P("【公式链展开】（从 r%d 沿 col 公式链向下，区域引用已逐行展开）" % top_row)
P("  叶子行总数                            : %d" % len(leaf_set))
P("  H 列链覆盖 / 漏                      : %d / %d" % (len(h_cov), len(h_miss)))
P("  G 列链覆盖 / 漏                      : %d / %d" % (len(g_cov), len(g_miss)))
P("  ├ ★H 漏、G 有（只改了报送列的铁证） : %d 行 %s" % (len(only_h_miss), sorted(only_h_miss)))
P("  ├ G 漏、H 有                        : %d 行 %s" % (len(only_g_miss), sorted(only_g_miss)))
P("  └ 两侧都漏                          : %d 行 %s" % (len(both_miss), sorted(both_miss)))
gap_rows = sorted(h_miss)
gap_sum_h = sum((D(sv.cell(r, 8).value) for r in gap_rows), Decimal(0))
gap_sum_g = sum((D(sv.cell(r, 7).value) for r in gap_rows), Decimal(0))
P("")
P("  H 链漏掉行的审定金额合计 = %s" % gap_sum_h)
P("  >>> 与口径差异是否吻合：%s（差 %s）"
  % ("✅ 完全吻合" if abs(gap_sum_h - gap) < Decimal("0.01") else "❌ 不吻合", gap_sum_h - gap))
if not gap_rows:
    P("  ⚠️ 未发现漏覆盖叶子 —— 差异可能来自：① 公式链不全；② 重复计入（同一行被多个父行引用）；")
    P("     ③ 合计行是硬编码常量而非公式。此时应改查『重复引用』：统计每个叶子行被引用的父行数。")
P("")

# ---------- 定位真正的缺陷处：G/H 两列公式链的「分叉点」 ----------
def find_divergence(start_row, col_g=7, col_h=8):
    """从 start_row 沿报送列(G)与审定列(H)的公式链**同时**向下，
       返回第一个『两边引用集合不一致』的行 —— 那才是缺陷所在。
       （不要用"引用目标行的最小行号"：项8 内部行 r295 的 G/H 公式是对称的，
         真正的漏引发生在更上层的汇总行 r238。）
       返回 (分叉行, G 多引用的行号, H 多引用的行号)"""
    stack, seen = [start_row], set()
    while stack:
        r = stack.pop()
        if r in seen: continue
        seen.add(r)
        sg = {rr for rr, cs in refs_in(sf.cell(r, col_g).value).items() if col_g in cs}
        sh = {rr for rr, cs in refs_in(sf.cell(r, col_h).value).items() if col_h in cs}
        if sg != sh:
            return r, sorted(sg - sh), sorted(sh - sg)
        stack.extend(sorted(sg))
    return None, [], []

div_row, g_extra, h_extra = find_divergence(top_row)
P("【缺陷定位】G/H 公式链分叉点：%s" % ("r%d" % div_row if div_row else "未发现"))
if div_row:
    P("  r%d G = %s" % (div_row, sf.cell(div_row, 7).value))
    P("  r%d H = %s" % (div_row, sf.cell(div_row, 8).value))
    P("  → G 列引用了 %s，H 列没有  ← 缺陷就在这个汇总行"
      % (["r%d" % r for r in g_extra] or "—"))
    P("     即：G 列链的覆盖比 H 列多这些行，正是导致审定合计少算的原因")
P("")

# ---------- "不予认可"的既有写法统计（审定量置 0） ----------
zero_rows = [r for r in leaf_rows
             if isinstance(sv.cell(r, 5).value, (int, float)) and abs(sv.cell(r, 5).value) < 1e-12
             and isinstance(sv.cell(r, 4).value, (int, float)) and abs(sv.cell(r, 4).value) > 1e-12]
P("【旁证】该表『不予认可』的既有写法 = 审定工程量置 0 的行：%d 行 %s"
  % (len(zero_rows), ["r%d" % r for r in zero_rows[:20]]))
if gap_rows:
    eq_cnt = sum(1 for r in gap_rows
                 if isinstance(sv.cell(r, 5).value, (int, float)) and isinstance(sv.cell(r, 4).value, (int, float))
                 and abs(D(sv.cell(r, 5).value) - D(sv.cell(r, 4).value)) < Decimal("1e-9")
                 and D(sv.cell(r, 5).value) != 0)
    P("        而漏覆盖的 %d 行中，有 %d 行是『审定量 = 报送量』（并非置 0）"
      % (len(gap_rows), eq_cnt))
P("")

# ---------- 目标表印证（可选） ----------
dst_ok = os.path.exists(DST)
r102 = r103 = None
if dst_ok:
    try:
        dv = openpyxl.load_workbook(DST, data_only=True)
        dsh = dv[dv.sheetnames[0]]
        for r in range(dsh.max_row, 0, -1):
            nm = dsh.cell(r, 2).value
            if nm and "计" in str(nm):
                r102 = (r, D(dsh.cell(r, 8).value)); break
        # 对账基准：合计行下面第一个有值的 col8
        if r102:
            for r in range(r102[0] + 1, dsh.max_row + 1):
                v = dsh.cell(r, 8).value
                if isinstance(v, (int, float)):
                    r103 = (r, D(v)); break
        P("【目标表印证】%s" % os.path.basename(DST))
        if r102: P("  合计行 r%d col8 = %s" % r102)
        if r103: P("  对账基准 r%d col8 = %s（合计 − 基准 = %s）" % (r103[0], r103[1], r102[1] - r103[1]))
        P("")
    except Exception as e:
        P("【目标表印证】读取失败：%s\n" % e)

# ---------- 两口径数值（供用户拍板） ----------
# 口径 A = 不计入漏覆盖行；口径 B = 计入
sb_total = sum_g
shr_total = sum_h
A_h = sum_h - gap_sum_h          # 剔除后
B_h = sum_h                      # 维持
P("【两口径对照】")
P("  报送合计            = %s" % sb_total)
P("  口径 A（%d 行不计入审定）= %s   核减额 = %s" % (len(gap_rows), A_h, sb_total - A_h))
P("  口径 B（计入审定）      = %s   核减额 = %s" % (B_h, sb_total - B_h))
P("  两口径差            = %s" % (B_h - A_h))
P("")

# ================= 生成报告 =================
thin = Side(style="thin", color="BFBFBF")
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
HDR = PatternFill("solid", fgColor="DDEBF7")
OK = PatternFill("solid", fgColor="E2EFDA")
WARN = PatternFill("solid", fgColor="FFF2CC")
BAD = PatternFill("solid", fgColor="FCE4E4")
F_T = Font(name="微软雅黑", size=11, bold=True)
F_H = Font(name="微软雅黑", size=10, bold=True)
F_N = Font(name="微软雅黑", size=10)
F_M = Font(name="Consolas", size=9)
AW = Alignment(vertical="top", wrap_text=True)
AC = Alignment(vertical="center", horizontal="center", wrap_text=True)

wb = openpyxl.Workbook()

# ---- Sheet1 差异溯源 ----
ws = wb.active; ws.title = "差异溯源"
for c, h in enumerate(["序号", "核查项", "数值 / 说明", "依据"], 1):
    cell = ws.cell(1, c, h); cell.font = F_H; cell.fill = HDR; cell.alignment = AC; cell.border = BD

def add(idx, item, val, src, tag="N"):
    global _r
    _r += 1
    ws.cell(_r, 1, idx).font = F_N
    ws.cell(_r, 2, item).font = F_N
    ws.cell(_r, 3, val).font = F_N
    ws.cell(_r, 4, src).font = F_M
    fill = {"N": None, "OK": OK, "W": WARN, "X": BAD}[tag]
    for c in range(1, 5):
        ws.cell(_r, c).border = BD; ws.cell(_r, c).alignment = AW
        if fill: ws.cell(_r, c).fill = fill

_r = 1
add(1, "源表叶子行数（col6 有单价）", str(len(leaf_rows)), os.path.basename(SRC) + " " + SRC_SHEET)
add(2, "口径① 明细：Σ(叶子行 col8 审定合计)", "%.6f" % sum_h, "源表全部叶子行 col8 求和")
add(3, "口径② 合计：项目级 r%d col8" % top_row, "%.6f" % proj_h, "源表 r%d（公式链顶端）" % top_row)
add(4, "差异 ①−②", "%.6f" % gap, "待解释的缺口", "W")
if NEGLIGIBLE:
    add(5, "结论：源表自洽，无实质差异", "差异 %.9f 元（分位以下，属浮点尾差）" % gap,
        "项目级合计已覆盖全部叶子行", "OK")
else:
    add(5, "差异来源：H 列公式链未覆盖的叶子行", "%d 行 %s" % (len(gap_rows), ["r%d" % r for r in gap_rows]),
        "G 列链已覆盖这些行、H 列链未覆盖", "X")
    add(6, "这些行的审定金额合计", "%.6f" % gap_sum_h,
        "与口径差异%s" % ("完全吻合" if abs(gap_sum_h - gap) < Decimal("0.01") else "不吻合"), "OK")
if gap_rows:
    add(7, "G 列链覆盖叶子数 / 总数", "%d / %d" % (len(g_cov), len(leaf_set)), "报送侧是否完整", "OK")
    add(8, "H 列链覆盖叶子数 / 总数", "%d / %d" % (len(h_cov), len(leaf_set)), "审定侧是否完整", "X")
    if div_row:
        add(9, "★缺陷定位：G/H 公式链分叉点 = r%d" % div_row,
            "r%d G 引用了 %s，H 没引用" % (div_row, ["r%d" % r for r in g_extra] or "—"),
            "同处插入只改了报送列 → 典型『漏改公式』", "X")
        add(10, "证据② 该行两列公式原文", "G = %s" % sf.cell(div_row, 7).value, "报送列已含新增行", "OK")
        add(11, "", "H = %s" % sf.cell(div_row, 8).value, "审定列漏了新增行", "X")
else:
    add(7, "G 列链覆盖叶子数 / 总数", "%d / %d" % (len(g_cov), len(leaf_set)), "报送侧", "OK")
    add(8, "H 列链覆盖叶子数 / 总数", "%d / %d" % (len(h_cov), len(leaf_set)), "审定侧", "OK")
add(12, "证据③ 漏覆盖行的审定金额取值",
    "行 %s，金额 %s" % (["r%d" % r for r in gap_rows[:8]], "%.6f" % gap_sum_h),
    "非 0、非空 → 若『有意不计』应为 0", "W")
add(13, "证据④ 该表『不予认可』的既有写法",
    "审定工程量置 0 的行：%d 行 %s" % (len(zero_rows), ["r%d" % r for r in zero_rows[:10]]),
    "与此对照，漏覆盖行并未置 0", "N")
if dst_ok and r102 and r103:
    add(14, "目标表印证：合计行 / 对账基准", "r%d=%s  vs  r%d=%s" % (r102[0], r102[1], r103[0], r103[1]),
        "合计 − 基准 = %.6f（归零说明已自洽）" % (r102[1] - r103[1]), "OK")
add(15, "❓ 待你裁定：两种口径" if not NEGLIGIBLE else "（源表自洽，以下口径对照仅供参考）",
    "A 不计入 %d 行 → 审核合计 %s（核减 %s）" % (len(gap_rows), A_h, sb_total - A_h),
    "口径差异 = %.2f" % (B_h - A_h), "W")
add(16, "", "B 计入        → 审核合计 %s（核减 %s）" % (B_h, sb_total - B_h),
    "选定后：① 落盘目标表使 Σ行=合计=r103；② 修源表汇总行公式使源表自洽", "W")
ws.column_dimensions["A"].width = 6
ws.column_dimensions["B"].width = 34
ws.column_dimensions["C"].width = 62
ws.column_dimensions["D"].width = 44

# ---- Sheet2 差异明细 ----
ws2 = wb.create_sheet("差异明细")
ws2.cell(1, 1, "被漏覆盖的叶子行明细（源表 %s）" % SRC_SHEET).font = F_T
cols = ["源表行", "父类别（最近表头 col2）", "定额名称", "单位", "综合单价",
        "报送工程量", "审定工程量", "报送金额", "审定金额"]
for c, h in enumerate(cols, 1):
    cell = ws2.cell(2, c, h); cell.font = F_H; cell.fill = HDR; cell.alignment = AC; cell.border = BD

# 找每行的父类别（最近一个 col6 为空、col2 有值的上方表头）
def parent_name(r):
    for rr in range(r - 1, 3, -1):
        if not isinstance(sv.cell(rr, 6).value, (int, float)) and sv.cell(rr, 2).value:
            return str(sv.cell(rr, 2).value)
    return ""

rr = 2
for r in gap_rows:
    rr += 1
    vals = [r, parent_name(r), str(sv.cell(r, 2).value), str(sv.cell(r, 3).value),
            sv.cell(r, 6).value, sv.cell(r, 4).value, sv.cell(r, 5).value,
            sv.cell(r, 7).value, sv.cell(r, 8).value]
    for c, v in enumerate(vals, 1):
        cell = ws2.cell(rr, c, v); cell.font = F_M if c in (1, 5, 6, 7, 8, 9) else F_N
        cell.border = BD; cell.alignment = AW
    for c in (6, 7): ws2.cell(rr, c).number_format = "#,##0.0000"
    for c in (5, 8, 9): ws2.cell(rr, c).number_format = "#,##0.00"
rr += 1
ws2.cell(rr, 3, "合计（金额）").font = F_H
ws2.cell(rr, 8, float(gap_sum_g)).font = F_H
ws2.cell(rr, 9, float(gap_sum_h)).font = F_H
for c in range(1, 10):
    ws2.cell(rr, c).border = BD; ws2.cell(rr, c).fill = OK; ws2.cell(rr, c).alignment = AW
for c in (8, 9): ws2.cell(rr, c).number_format = "#,##0.00"
for c, w in zip("ABCDEFGHI", [8, 24, 30, 10, 12, 12, 14, 14, 14]):
    ws2.column_dimensions[c].width = w

# ---- Sheet3 两口径对照 ----
ws3 = wb.create_sheet("两口径对照")
ws3.cell(1, 1, "两种口径的金额后果（请择一裁定）").font = F_T
for c, h in enumerate(["口径", "报送合计", "审定/审核合计", "核减额", "落盘做法"], 1):
    cell = ws3.cell(2, c, h); cell.font = F_H; cell.fill = HDR; cell.alignment = AC; cell.border = BD
rows3 = [
    ("A 漏覆盖的 %d 行不计入审定" % len(gap_rows), float(sb_total), float(A_h), float(sb_total - A_h),
     "从目标表相应行扣减审核量与审核金额；r103 对账基准随之设为 A；源表父行公式不必改（保持现状即口径A）"),
    ("B 漏覆盖的 %d 行计入审定" % len(gap_rows), float(sb_total), float(B_h), float(sb_total - B_h),
     "维持目标表全额落盘；r103 设为 B；并用 patch_parent_formula.py 修源表父行公式补上漏引行，使源表自洽"),
]
for i, row in enumerate(rows3, 3):
    for c, v in enumerate(row, 1):
        cell = ws3.cell(i, c, v); cell.font = F_N if c != 1 else F_H
        cell.border = BD; cell.alignment = AW; cell.fill = OK if i == 4 else BAD
    for c in (2, 3, 4): ws3.cell(i, c).number_format = "#,##0.00"
ws3.cell(5, 1, "※ 两口径差额 = %.2f，即漏覆盖行是否认可；该差异同时决定核减额。" % float(B_h - A_h)).font = F_N
ws3.cell(6, 1, "※ 口径 A 与源表当前状态一致（源表父行公式就是漏的）；口径 B 才是『明细口径』，"
               "需同步修源表父行公式，否则源表自身矛盾会一直存在。").font = F_N
ws3.cell(6, 1).alignment = AW
for c, w in zip("ABCDE", [36, 16, 18, 16, 62]):
    ws3.column_dimensions[c].width = w

os.makedirs(os.path.dirname(REPORT), exist_ok=True)
wb.save(REPORT)
P("[已生成] %s" % REPORT)
open(LOG, "w", encoding="utf-8").write("\n".join(L))
print("done -> %s" % REPORT)
print("  差异 = %s ; 漏覆盖 %d 行 ; 金额 %s" % (gap, len(gap_rows), gap_sum_h))
