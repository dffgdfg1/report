# -*- coding: utf-8 -*-
import io, os
P = os.path.join(os.path.dirname(os.path.abspath(__file__)), "raw_record_engine.py")
src = io.open(P, encoding="utf-8").read()
orig = src

a1 = 'SKELETON = os.path.join(BASE, "模板库", "原始记录骨架.docx")'
add1 = a1 + '\n# 每张原始记录末尾另起一页追加的「试验过程点检记录」附页\nCHECKLIST = os.path.join(BASE, "模板库", "试验过程点检记录.docx")'
assert a1 in src, "a1"
assert "CHECKLIST =" not in src, "already patched"
src = src.replace(a1, add1, 1)

a2 = "import report_engine as E"
src = src.replace(a2, a2 + "\nfrom docxcompose.composer import Composer", 1)

a3 = "        doc.save(file_path)"
assert a3 in src, "a3"
src = src.replace(a3, "        _append_checklist(doc)\n" + a3, 1)

a4 = "def _make_page_break(doc):"
helper = (
    'def _append_checklist(doc):\n'
    '    """在原始记录末尾另起一页追加「试验过程点检记录」附页。\n'
    '    附页模板缺失时静默跳过，不影响原始记录生成。"""\n'
    '    if not os.path.exists(CHECKLIST):\n'
    '        return\n'
    '    # 在正文 sectPr 之前插入分页符，使附页从新一页开始\n'
    '    body = doc.element.body\n'
    "    sectPr = body.find(qn('w:sectPr'))\n"
    '    pb = _make_page_break(doc)\n'
    '    if sectPr is not None:\n'
    '        sectPr.addprevious(pb)\n'
    '    else:\n'
    '        body.append(pb)\n'
    '    Composer(doc).append(Document(CHECKLIST))\n'
    '\n\n'
)
assert a4 in src, "a4"
src = src.replace(a4, helper + a4, 1)

assert src != orig
io.open(P, "w", encoding="utf-8", newline="\n").write(src)
print("patched OK; len", len(orig), "->", len(src))
