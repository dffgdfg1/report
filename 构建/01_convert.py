# -*- coding: utf-8 -*-
"""一次性：把原始 .doc 模板转成完整 .docx，存入 模板库/主体_full.docx"""
import win32com.client as win32, os

SRC = r"C:\Users\admin\Desktop\测试项目\O1D\MEWTD20251230012 O1D试验报告.doc"
DST = r"C:\Users\admin\Desktop\报告生成器\模板库\主体_full.docx"

app = win32.gencache.EnsureDispatch("KWps.Application")
app.Visible = False
try:
    app.DisplayAlerts = False
except Exception:
    pass
doc = app.Documents.Open(SRC, ReadOnly=True)
doc.SaveAs2(DST, FileFormat=16)  # 16 = docx
doc.Close(False)
app.Quit()
print("OK", DST, os.path.getsize(DST))
