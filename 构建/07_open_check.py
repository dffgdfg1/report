# -*- coding: utf-8 -*-
import win32com.client as win32
p=r"C:\Users\admin\Desktop\报告生成器\输出\_测试输出.docx"
app=win32.gencache.EnsureDispatch("KWps.Application")
app.Visible=False
try: app.DisplayAlerts=False
except: pass
doc=app.Documents.Open(p, ReadOnly=True)
try:
    pages=doc.ComputeStatistics(2)  # wdStatisticPages=2
except Exception as e:
    pages="?"
print("OPENED OK, pages=", pages)
doc.Close(False)
app.Quit()
