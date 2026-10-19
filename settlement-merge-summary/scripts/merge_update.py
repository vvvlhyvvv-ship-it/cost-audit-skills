# -*- coding: utf-8 -*-
# [skill: settlement-merge-summary] 读源 Sheet2 → 按四元组主键聚合 → 更新合并汇总
# 用法：改顶部 BASE/SRC/DST 路径常量后直接运行；输出写入 .txt 便于 Read 查看。
"""
用「调整报送2026.9.9\工程结算书8-28 - 副本.xlsx」的 Sheet2（15村明细）
合并汇总更新到「调整报送2026.9.9\工程结算书合并汇总.xlsx」的「合并汇总」sheet。

主键 = (父类别名, 定额名, 单位, 单价) 四元组。
金额直接搬运源表 报送合计/审定合计（避免"工程量×单价再取整"的行级偏差）。
仅更新叶子行 + 合计行；表头行的 亩/m 面积量不做 roll-up（单位不同），保留原值。
DRY_RUN=False 就地落盘。
"""
import openpyxl, re, shutil
from collections import defaultdict

SRC = r"<项目工作区>\审核过程资料\工程结算书8-28 - 副本.xlsx"
DST = r"<项目工作区>\审核过程资料\工程结算书合并汇总.xlsx"
OUT = r"D:\Workbuddy文件\2026-07-14-15-47-54\hb9_final_log.txt"
DRY_RUN = False

L = []
def P(s=""): L.append(str(s))

def nos(s):
    if s is None: return ""
    return re.sub(r"\s+", "", str(s))

def nu(u):
    if u is None: return ""
    s = re.sub(r"\s+", "", str(u))
    return s.replace("m3","m³").replace("m2","m²").replace("hm2","hm²").replace("km2","km²")

ALIAS = [("装石碴自卸汽车运输","石碴运输"), ("挖掘机装石碴","石碴"),
         ("挖掘机挖装运土","挖装自卸汽车运土"), ("伸缩缝","道路伸缩缝")]
SUFFIXES = [
    r"自卸汽车柴油型载重量[\d\.]+t", r"柴油型载重量[\d\.]+t", r"载重量[\d\.]+t",
    r"推土机功率\d+kW", r"功率\d+kW",
    r"1m³挖掘机", r"挖掘机装石碴自卸汽车运输", r"装石碴自卸汽车运输",
    r"运距[0-9.~～\-]+km", r"运距[0-9.~～\-]+", r"运距$",
    r"推土距离.*",
    r"厚度\d+cm", r"增加[\-\d\.]+cm", r"三类土", r"机械夯填", r"其他机械",
    r"φ[\d\.]+mm", r"平段公称直径\d+mm以内", r"（51m）", r"\(51m\)",
    r"机械摊铺压实", r"机械摊铺", r"压实", r"沥青木板",
    r"\[[^\]]*\]", r"（[^）]*[：:][^）]*）", r"\([^)]*[:：][^)]*\)",
]
def nq(s):
    n = nos(s)
    for a, b in ALIAS: n = n.replace(a, b)
    for suf in SUFFIXES: n = re.sub(suf, "", n)
    return re.sub(r"(C25混凝土渠).*", r"\1", n)

# ---------- 1. 源 Sheet2 聚合 ----------
wb_src = openpyxl.load_workbook(SRC, data_only=True)
ws2 = wb_src["Sheet2"]
agg = defaultdict(lambda: {"sb":0.0,"sh":0.0,"jeb":0.0,"jeh":0.0,"n":0})
current_parent = None
src_total_jeb = src_total_jeh = 0.0
for r in range(4, ws2.max_row+1):
    nm = ws2.cell(r,2).value
    u  = ws2.cell(r,3).value
    c4 = ws2.cell(r,4).value; c5 = ws2.cell(r,5).value; c6 = ws2.cell(r,6).value
    c7 = ws2.cell(r,7).value; c8 = ws2.cell(r,8).value
    if nm is None: continue
    if isinstance(c6,(int,float)):
        if not isinstance(c4,(int,float)) and not isinstance(c5,(int,float)): continue
        key = (nq(current_parent) if current_parent else "", nq(str(nm)), nu(u), round(c6,2))
        a = agg[key]
        a["sb"] += c4 if isinstance(c4,(int,float)) else 0.0
        a["sh"] += c5 if isinstance(c5,(int,float)) else 0.0
        a["jeb"] += c7 if isinstance(c7,(int,float)) else 0.0
        a["jeh"] += c8 if isinstance(c8,(int,float)) else 0.0
        a["n"] += 1
        src_total_jeb += c7 if isinstance(c7,(int,float)) else 0.0
        src_total_jeh += c8 if isinstance(c8,(int,float)) else 0.0
    else:
        current_parent = str(nm)
P("源 Sheet2 聚合项: %d" % len(agg))
P("源 叶子求和: 报送=%.2f 审定=%.2f  (源 r4: 报送=%s 审定=%s)" % (
    src_total_jeb, src_total_jeh, ws2.cell(4,7).value, ws2.cell(4,8).value))

# ---------- 2. 目标 合并汇总 解析（读值用 data_only；写用普通模式保留公式） ----------
wb_v = openpyxl.load_workbook(DST, data_only=True)
wb_w = openpyxl.load_workbook(DST)
wmv = wb_v["合并汇总"]
wmw = wb_w["合并汇总"]

def level_of(seq):
    s = str(seq)
    if s.startswith("（") or s.startswith("("): return 0
    return s.count(".")+1

stack = []
items = []
for r in range(4, wmv.max_row+1):
    idx = wmv.cell(r,1).value; nm = wmv.cell(r,2).value; u = wmv.cell(r,3).value
    c4 = wmv.cell(r,4).value; c5 = wmv.cell(r,5).value; c6 = wmv.cell(r,6).value
    if nm is None: continue
    nm_s = str(nm)
    if nm_s.strip() == "合  计":
        items.append({"r":r,"kind":"total"}); continue
    if isinstance(c6,(int,float)):
        parent = stack[-1][1] if stack else None
        items.append({"r":r,"idx":idx,"nm":nm_s,"u":u,"price":round(c6,2),
                      "parent":parent,"kind":"leaf",
                      "old_c4":c4,"old_c5":c5,
                      "old_c7":wmv.cell(r,7).value,"old_c8":wmv.cell(r,8).value})
    else:
        lv = level_of(idx) if idx is not None else 999
        while stack and stack[-1][0] >= lv: stack.pop()
        stack.append((lv, nm_s))
        items.append({"r":r,"idx":idx,"nm":nm_s,"u":u,"kind":"header","old_c4":c4,"old_c5":c5})

leaf_total = sum(1 for i in items if i.get("kind")=="leaf")
P("目标 合并汇总: 叶子=%d 表头=%d 合计行=%d" % (
    leaf_total, sum(1 for i in items if i.get("kind")=="header"),
    sum(1 for i in items if i.get("kind")=="total")))

# ---------- 3. 匹配 ----------
used = set()
def find_match(name_n, unit, price, parent_n):
    key = (parent_n, name_n, unit, price)
    if key in agg:
        used.add(key); return agg[key]
    if "/" in name_n:
        tot = {"sb":0.0,"sh":0.0,"jeb":0.0,"jeh":0.0,"n":1}
        hit = False
        for part in name_n.split("/"):
            k2 = (parent_n, nq(part), unit, price)
            if k2 in agg:
                for f in ("sb","sh","jeb","jeh"): tot[f] += agg[k2][f]
                used.add(k2); hit = True
        if hit: return tot
    return None

matched = 0; unmatched = []
for it in items:
    if it.get("kind") != "leaf": continue
    m = find_match(nq(it["nm"]), nu(it["u"]), it["price"], nq(it["parent"]) if it["parent"] else "")
    if m:
        matched += 1
        it["new_c4"] = round(m["sb"],3); it["new_c5"] = round(m["sh"],3)
        it["new_c7"] = round(m["jeb"],2); it["new_c8"] = round(m["jeh"],2)
        it["new_c9"] = round(m["jeh"]-m["jeb"],2)
    else:
        it["new_c4"]=0.0; it["new_c5"]=0.0; it["new_c7"]=0.0; it["new_c8"]=0.0; it["new_c9"]=0.0
        unmatched.append(it)

P("叶子匹配: %d / %d" % (matched, leaf_total))
unused = [k for k in agg if k not in used]
P("源未被消费聚合项: %d 项, 审定金额=%.2f" % (len(unused), sum(agg[k]["jeh"] for k in unused)))
for k in sorted(unused, key=lambda x:-agg[x]["jeh"]):
    P("   父=%-22s 名=%-26s 单位=%-8s 单价=%-10s 审定=%.2f" % (k[0][:20],k[1][:24],k[2],k[3],agg[k]["jeh"]))

tot_sb = sum(it.get("new_c7") or 0.0 for it in items if it.get("kind")=="leaf")
tot_sh = sum(it.get("new_c8") or 0.0 for it in items if it.get("kind")=="leaf")
P("")
P("新合计: 报送=%.2f 审核=%.2f 差额=%.2f" % (tot_sb, tot_sh, tot_sh-tot_sb))
P("源叶子求和: 报送=%.2f 审定=%.2f" % (src_total_jeb, src_total_jeh))
P("一致性: 报送差=%.2f  审定差=%.2f" % (tot_sb-src_total_jeb, tot_sh-src_total_jeh))

P("")
P("==== 变化行 ====")
for it in items:
    if it.get("kind")!="leaf": continue
    ch = (abs((it.get("new_c4") or 0)-(it.get("old_c4") or 0))>0.005 or
          abs((it.get("new_c5") or 0)-(it.get("old_c5") or 0))>0.005 or
          abs((it.get("new_c7") or 0)-(it.get("old_c7") or 0))>0.005 or
          abs((it.get("new_c8") or 0)-(it.get("old_c8") or 0))>0.005)
    if ch:
        P("  r%-3d %-26s | 审核量 %-11s→%-11s | 审核额 %-13s→%-13s | 报送额 %-13s→%-13s" % (
            it["r"], it["nm"][:24],
            ("%.3f"%it["old_c5"]) if isinstance(it["old_c5"],(int,float)) else "-", "%.3f"%it["new_c5"],
            ("%.2f"%it["old_c8"]) if isinstance(it["old_c8"],(int,float)) else "-", "%.2f"%it["new_c8"],
            ("%.2f"%it["old_c7"]) if isinstance(it["old_c7"],(int,float)) else "-", "%.2f"%it["new_c7"]))

if not DRY_RUN:
    for it in items:
        if it.get("kind") != "leaf": continue
        wmw.cell(it["r"], 4, it["new_c4"])
        wmw.cell(it["r"], 5, it["new_c5"])
        wmw.cell(it["r"], 7, it["new_c7"])
        wmw.cell(it["r"], 8, it["new_c8"])
        wmw.cell(it["r"], 9, it["new_c9"])
    for it in items:
        if it.get("kind") == "total":
            wmw.cell(it["r"], 7, round(tot_sb,2))
            wmw.cell(it["r"], 8, round(tot_sh,2))
            wmw.cell(it["r"], 9, round(tot_sh-tot_sb,2))
            P("")
            P("合计行写入 r%d: 报送=%.2f 审核=%.2f 差额=%.2f" % (it["r"], tot_sb, tot_sh, tot_sh-tot_sb))
    # 校验单元格 r103 col8（原 54,565,093.0325637）
    if wmv.cell(103,8).value is not None:
        wmw.cell(103, 8, round(tot_sh,4))
        P("校验单元格 r103 col8 写入 = %.4f" % round(tot_sh,4))
    wb_w.save(DST)
    P("")
    P("[已落盘] %s" % DST)
else:
    P("[DRY_RUN 未落盘]")

open(OUT,"w",encoding="utf-8").write("\n".join(L))
print("done")
