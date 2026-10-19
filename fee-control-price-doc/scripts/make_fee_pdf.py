# -*- coding: utf-8 -*-
r"""
三件套 → 单一 PDF（页序：封皮 / 编制说明 / 计费清单），并逐页渲染 PNG 供目视复核。
用法：改 OUT 后运行。依赖 win32com(pywin32) + pymupdf，托管 venv 已装。
"""
import os, sys

OUT = r"<项目工作区>\XXXX\二类费用控制价\XXXX费"
PNG_DIR = os.path.join(OUT, "_preview")

TMP = os.path.join(OUT, "_pdf")
FINAL = os.path.join(OUT, os.path.basename(OUT.rstrip("\\")) + "控制价.pdf")

import win32com.client as win32
import fitz

os.makedirs(TMP, exist_ok=True)

# ---------- Word ----------
word = win32.DispatchEx("Word.Application")
word.Visible = False
try:
    word.DisplayAlerts = 0
except Exception:
    pass
for name in ["封皮", "编制说明"]:
    doc = word.Documents.Open(os.path.join(OUT, name + ".docx"), ReadOnly=True)
    doc.SaveAs(os.path.join(TMP, name + ".pdf"), FileFormat=17)   # wdFormatPDF
    print("[PDF]", name, "pages =", doc.ComputeStatistics(2))
    doc.Close(False)
word.Quit()

# ---------- Excel ----------
xl = win32.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
wb = xl.Workbooks.Open(os.path.join(OUT, "计费清单.xlsx"))
ws = wb.Worksheets("计费清单")
ws.PageSetup.Orientation = 2            # 横向
ws.PageSetup.Zoom = False
ws.PageSetup.FitToPagesWide = 1
ws.PageSetup.FitToPagesTall = 1
ws.PageSetup.CenterHorizontally = True
wb.ExportAsFixedFormat(0, os.path.join(TMP, "计费清单.pdf"))       # xlTypePDF
wb.Close(False)
xl.Quit()

# ---------- 合并 ----------
merged = fitz.open()
for name in ["封皮", "编制说明", "计费清单"]:
    d = fitz.open(os.path.join(TMP, name + ".pdf"))
    merged.insert_pdf(d)
    d.close()
merged.save(FINAL)
print("[OK]", FINAL, "共", merged.page_count, "页（应为 3）")

# ---------- 渲染复核 ----------
os.makedirs(PNG_DIR, exist_ok=True)
for i, page in enumerate(merged):
    page.get_pixmap(dpi=110).save(os.path.join(PNG_DIR, "page%d.png" % (i + 1)))
merged.close()
print("[OK] 预览图 ->", PNG_DIR)

import shutil
shutil.rmtree(TMP, ignore_errors=True)
