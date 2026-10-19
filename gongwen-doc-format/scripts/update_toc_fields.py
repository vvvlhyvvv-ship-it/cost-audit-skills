# -*- coding: utf-8 -*-
"""
用 Word COM 更新目录域并生成真实页码 —— 套版后必跑。

用法：
    python update_toc_fields.py <docx路径> [<docx路径2> ...] [--pdf 输出目录]

要点（踩过的坑见 SKILL.md §四）：
  * 必须 Repaginate 后再取页数；改版式会使页码变化，故每次套版后都要重跑。
  * 只更新目录域不够，TOC 条目字体/行距回到 TOC 样式 —— 样式须在套版脚本里设好。
  * 结束后务必清理 WINWORD 进程与桌面 ~$ 锁文件。
"""
import os, sys, io
import pythoncom
import win32com.client as wc

WD_EXPORT_FORMAT_PDF = 17


def update(paths, pdf_dir=None, log=print):
    out = io.StringIO()
    pythoncom.CoInitialize()
    word = wc.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for path in paths:
            log("=" * 72); log(os.path.basename(path))
            doc = word.Documents.Open(path, ConfirmConversions=False, ReadOnly=False,
                                      AddToRecentFiles=False, Revert=False)
            try:
                n = doc.TablesOfContents.Count
                log("  TOC 数量 = %d" % n)
                for i in range(1, n + 1):
                    doc.TablesOfContents(i).Update()
                doc.Fields.Update()
                doc.Repaginate()
                log("  页数 = %d" % doc.ComputeStatistics(2))   # wdStatisticPages
                doc.Save()
                log("  已保存")
                if pdf_dir:
                    os.makedirs(pdf_dir, exist_ok=True)
                    pdf = os.path.join(pdf_dir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
                    doc.ExportAsFixedFormat(
                        OutputFileName=pdf, ExportFormat=WD_EXPORT_FORMAT_PDF,
                        OpenAfterExport=False, OptimizeFor=0, Range=0, Item=0,
                        IncludeDocProps=False, KeepIRM=True, CreateBookmarks=0,
                        DocStructureTags=False, BitmapMissingFonts=True,
                        UseISO19005_1=False)
                    log("  已导出 PDF：%s" % pdf)
            finally:
                doc.Close(0)
    finally:
        word.Quit()
    return out.getvalue()


if __name__ == '__main__':
    argv = sys.argv[1:]
    pdf_dir = None
    if '--pdf' in argv:
        i = argv.index('--pdf')
        pdf_dir = argv[i + 1] if i + 1 < len(argv) else None
        del argv[i:i + 2]
    paths = [a for a in argv if not a.startswith('--')]
    if not paths:
        print(__doc__); sys.exit(1)
    update(paths, pdf_dir)
    print("\n完成。请运行 verify_format.py 核验行距 / keepNext / 落单标题。")
