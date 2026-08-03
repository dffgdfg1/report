# -*- coding: utf-8 -*-
"""分析 主体_full.docx 的块序列与各表格单元格，结果写入 分析结果.txt（UTF-8）"""
import io
from docx import Document
from docx.shared import Emu

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A_BLIP='{http://schemas.openxmlformats.org/drawingml/2006/main}blip'
V_IMG='{urn:schemas-microsoft-com:vml}imagedata'

doc = Document(r"C:\Users\admin\Desktop\报告生成器\模板库\主体_full.docx")
out = io.StringIO()

def imgn(el):
    return len(el.findall('.//'+A_BLIP))+len(el.findall('.//'+V_IMG))

# 1) block sequence
out.write("===== 块序列 =====\n")
i=0
tbl_idx=0
for child in doc.element.body.iterchildren():
    tag=child.tag.split('}')[-1]
    if tag=='p':
        t="".join([(x.text or "") for x in child.findall('.//'+W+'t')]).strip()
        n=imgn(child)
        out.write(f"{i:3} P   {'[IMG%d]'%n if n else ''} {t[:70]}\n")
    elif tag=='tbl':
        rows=child.findall('./'+W+'tr')
        out.write(f"{i:3} TBL#{tbl_idx} rows={len(rows)} imgs={imgn(child)}\n")
        tbl_idx+=1
    else:
        out.write(f"{i:3} {tag}\n")
    i+=1

# 2) dump every table cell grid (text only)
out.write("\n===== 各表格单元格 =====\n")
for ti,t in enumerate(doc.tables):
    out.write(f"\n--- TBL#{ti}  {len(t.rows)}行 x {len(t.columns)}列  imgs={imgn(t._tbl)} ---\n")
    for ri,row in enumerate(t.rows):
        cells=[]
        seen=set()
        for c in row.cells:
            key=id(c._tc)
            if key in seen:
                cells.append("<merge>")
                continue
            seen.add(key)
            has_img = imgn(c._tc)>0
            txt=c.text.strip().replace("\n","/")
            cells.append(("[图]" if has_img else txt[:24]))
        out.write(f"  R{ri}: " + " | ".join(cells) + "\n")

with open(r"C:\Users\admin\Desktop\报告生成器\构建\分析结果.txt","w",encoding="utf-8") as f:
    f.write(out.getvalue())
print("done")
