# -*- coding: utf-8 -*-
"""
人日法计费独立复算（Decimal 精算 + 断言 + 金额大写）
用法：改 CONFIG 后运行。不复用构建脚本的任何中间变量，从费率表与配置重算。
"""
from decimal import Decimal as D

# =================== CONFIG ===================
# 中咨协政〔2015〕46号 费率表（元/人工日）
RATE_TABLE = {
    "规划咨询":      {"外部高级专家": D("3000"), "高级职称及以上": D("2210"), "中级职称": D("1730"), "初级职称及以下": D("1230"), "系数": D("2.44")},
    "项目建议书编制": {"外部高级专家": D("2780"), "高级职称及以上": D("2040"), "中级职称": D("1610"), "初级职称及以下": D("1200"), "系数": D("2.51")},
    "可行性研究报告编制": {"外部高级专家": D("2920"), "高级职称及以上": D("2100"), "中级职称": D("1630"), "初级职称及以下": D("1230"), "系数": D("2.55")},
    "其他咨询服务":  {"外部高级专家": D("2760"), "高级职称及以上": D("1990"), "中级职称": D("1470"), "初级职称及以下": D("1070"), "系数": D("2.28")},
}
CATEGORY = "规划咨询"                 # ← 选定口径，须与《编制说明》一致

# (工作阶段, 人员类别, 人日数)
CONFIG = [
    ("项目调研与基础数据采集", "高级职称及以上", 2),
    ("项目调研与基础数据采集", "中级职称", 3),
    ("成本测算与方案编制", "高级职称及以上", 3),
    ("成本测算与方案编制", "中级职称", 3),
    ("内部审核与专家评审", "外部高级专家", 1),
    ("内部审核与专家评审", "高级职称及以上", 1),
]

EXPECT_TOTAL_RI = 13
EXPECT_DEPRECATE_AMT = D("26640")
EXPECT_RAW = D("65001.60")
EXPECT_TIDY = D("65000")
# ==============================================


def upper_cn(n):
    """整数金额转中文大写（元）"""
    digits = "零壹贰叁肆伍陆柒捌玖"
    units = ["", "拾", "佰", "仟"]
    groups = ["", "万", "亿", "万亿"]
    s = str(int(n))
    out = ""
    ln = len(s)
    for i, ch in enumerate(s):
        pos = ln - 1 - i
        g, u = pos // 4, pos % 4
        d = int(ch)
        out += digits[d] + (units[u] if d else "")
        if u == 0:
            out += groups[g]
    while "零零" in out:
        out = out.replace("零零", "零")
    out = out.rstrip("零").replace("零万", "万").replace("零亿", "亿")
    return out + "元整"


def main():
    tab = RATE_TABLE[CATEGORY]
    coef = tab["系数"]
    total_ri, total_amt = 0, D("0")
    print("口径：%s（综合计算系数 %s）\n" % (CATEGORY, coef))
    print("%-22s %-14s %8s %6s %12s" % ("工作阶段", "人员类别", "单价", "人日", "金额"))
    for stage, role, days in CONFIG:
        amt = tab[role] * days
        total_ri += days
        total_amt += amt
        print("%-22s %-14s %8s %6d %12s" % (stage, role, tab[role], days, amt))
    raw = total_amt * coef
    tidy = (raw / D("100")).quantize(D("1"), rounding="ROUND_HALF_UP") * D("100")
    print("-" * 68)
    print("直接人工成本 =", total_amt, "｜ 总人日 =", total_ri)
    print("编制费 = %s × %s = %s" % (total_amt, coef, raw))
    print("取整至百元 =", tidy)
    print("大写 =", upper_cn(tidy))

    assert total_ri == EXPECT_TOTAL_RI, ("人日合计", total_ri)
    assert total_amt == EXPECT_DEPRECATE_AMT, ("人工成本", total_amt)
    assert raw == EXPECT_RAW, ("编制费原值", raw)
    assert tidy == EXPECT_TIDY, ("取定值", tidy)
    print("\n全部断言通过 ✅")


if __name__ == "__main__":
    main()
