# -*- coding: utf-8 -*-
# [skill: settlement-merge-summary] 叶子行金额 6 位小数清理，保证 Σ行=合计=r103
# 用法：改顶部 BASE/SRC/DST 路径常量后直接运行；输出写入 .txt 便于 Read 查看。
"""精度清理：叶子行金额保留 6 位小数（去掉浮点尾数），合计行/校验值按清理后求和重设"""
import openpyxl, os
BASE = r"<项目工作区>\审核过程资料"
DST  = os.path.join(BASE, "工程结算书合并汇总.xlsx")
OUT  = r"D:\Workbuddy文件\2026-07-14-15-47-54\hb9_clean6.txt"
L=[]
def P(s=""): L.append(str(s))

wb = openpyxl.load_workbook(DST)
ws = wb["合并汇总"]
tot7 = tot8 = 0.0
n = 0
for r in range(4, ws.max_row+1):
    nm = ws.cell(r,2).value
    c6 = ws.cell(r,6).value
    if nm is None or "计" in str(nm): continue
    if not isinstance(c6,(int,float)): continue
    v7 = ws.cell(r,7).value; v8 = ws.cell(r,8).value
    v7 = round(v7, 6) if isinstance(v7,(int,float)) else 0.0
    v8 = round(v8, 6) if isinstance(v8,(int,float)) else 0.0
    ws.cell(r,7, v7); ws.cell(r,8, v8); ws.cell(r,9, round(v8-v7, 6))
    tot7 += v7; tot8 += v8; n += 1
tot7 = round(tot7, 6); tot8 = round(tot8, 6)
for r in range(4, ws.max_row+1):
    nm = ws.cell(r,2).value
    if nm is not None and "计" in str(nm):
        ws.cell(r,7, tot7); ws.cell(r,8, tot8); ws.cell(r,9, round(tot8-tot7,6))
        P("合计行 r%d: 报送=%s 审核=%s 差额=%s" % (r, tot7, tot8, round(tot8-tot7,6)))
        break
ws.cell(103, 8, tot8)
ws.cell(103, 8).number_format = "#,##0.00"
P("叶子行 %d 行已按 6 位小数清理" % n)
P("r103 校验值 = %s" % tot8)
P("r104 公式 = %s → 打开时重算为 0" % ws.cell(104,8).value)
wb.save(DST)
P("[已落盘] %s" % DST)

# 回读抽查
wv = openpyxl.load_workbook(DST, data_only=True)["合并汇总"]
P("")
P("=== 回读抽查（存储原值 / 显示 2 位）===")
for r in [7,17,20,89,92,93,96,101,102,103]:
    P("  r%-3d 报送额=%-22s 审核额=%-22s" % (r, wv.cell(r,7).value, wv.cell(r,8).value))
open(OUT,"w",encoding="utf-8").write("\n".join(L))
print("done")
