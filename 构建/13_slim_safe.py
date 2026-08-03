# -*- coding: utf-8 -*-
"""安全瘦身：只删 document.xml 里已不再引用的图片(被删测试段落遗留的照片)。
绝不触碰页眉/页脚/主题等其它 .rels，保留封面 logo、底纹、公章。"""
import zipfile, re, os, shutil, posixpath

SRC = r"C:\Users\admin\Desktop\报告生成器\模板库\骨架.docx"
TMP = SRC + ".slim.tmp"

zin = zipfile.ZipFile(SRC, "r")
data = {n: zin.read(n) for n in zin.namelist()}
zin.close()

doc = data["word/document.xml"].decode("utf-8", "ignore")
ref = set(re.findall(r'r:(?:id|embed|link)="(rId\d+)"', doc)) | set(re.findall(r'o:relid="(rId\d+)"', doc))

rels_name = "word/_rels/document.xml.rels"
rels = data[rels_name].decode("utf-8", "ignore")

drop_targets = set()
def repl(m):
    tag = m.group(0)
    if "/image" not in tag:
        return tag
    rid = re.search(r'Id="([^"]+)"', tag)
    tgt = re.search(r'Target="([^"]+)"', tag)
    rid = rid.group(1) if rid else ""
    if rid and rid not in ref:  # 孤儿图片关系 -> 删除
        if tgt:
            base = posixpath.normpath(posixpath.join("word", tgt.group(1)))
            drop_targets.add(base.replace("\\", "/"))
        return ""
    return tag

rels_new = re.sub(r'<Relationship\b[^>]*/>', repl, rels)
data[rels_name] = rels_new.encode("utf-8")

# 删除对应 media 文件（仅当没有任何其它 .rels 仍引用它）
still_used = set()
for n, b in data.items():
    if n.endswith(".rels") and n != rels_name:
        for mm in re.findall(rb'Target="([^"]+)"', b):
            t = mm.decode("utf-8", "ignore")
            base = posixpath.normpath(posixpath.join(posixpath.dirname(posixpath.dirname(n)), t))
            still_used.add(base.replace("\\", "/"))

removed = []
for t in list(drop_targets):
    if t in data and t not in still_used:
        del data[t]; removed.append(t)

if os.path.exists(TMP): os.remove(TMP)
zout = zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED)
for n, b in data.items():
    zout.writestr(n, b)
zout.close()
shutil.move(TMP, SRC)
print("删除孤儿图片数:", len(removed))
print("新骨架大小: %.2f MB" % (os.path.getsize(SRC) / 1048576))
