# -*- coding: utf-8 -*-
"""解析「试验申请单」PDF（文字层），提取首页信息字段。
依赖 pypdf。OA 导出用私有区(PUA)字符表示勾选状态：
   = 选中的单选(◉)      = 未选的单选(○)
   = 勾选的复选(☑)      = 未勾的复选(□)
"""
import re

RADIO_ON = ""   # ◉ 选中的单选
CHECK_ON = ""   # ☑ 勾选的复选
# 常见 PUA 标记字符，清理普通文本时统一去掉
_PUA = re.compile(r"[-]")

def _clean(s):
    return _PUA.sub("", s or "").strip()

# 勾选式字段里，被选中的往往是“其他/其它”这个填空项，行首会带上选项标签。
# 真正内容在标签之后，这里把开头的“其他/其它”及紧跟的标点、空格剥掉。
_OPT_PREFIX = re.compile(r"^其[他它][\s.．、,，:：]*")
def _strip_opt(s):
    return _OPT_PREFIX.sub("", s or "").strip()

def _line_after(lines, label):
    """返回以 label 开头那行的下一行内容（清理后）。"""
    for i, ln in enumerate(lines):
        if ln.strip().startswith(label) and i + 1 < len(lines):
            return _clean(lines[i + 1])
    return ""
def _selected_radio(line):
    """从一行(含多个单选项)里取出被选中的那项文字。
    例：'验证阶段 EV○ DV○ PV○ 量产○ 其它◉ 客户回退的异常件' -> ('其它', '客户回退的异常件')"""
    body = line.split(None, 1)[1] if " " in line.strip() else line
    toks = body.split()
    sel, tail_idx = "", -1
    for i, tk in enumerate(toks):
        if RADIO_ON in tk:
            sel = _clean(tk); tail_idx = i; break
    tail = ""
    if tail_idx >= 0 and tail_idx + 1 < len(toks):
        rest = " ".join(toks[tail_idx + 1:])
        # 后面若还有别的单选项则不算说明文字
        if RADIO_ON not in rest and "" not in rest:
            tail = _clean(rest)
    return sel, tail

def _checked_value(lines, label, maxscan=4):
    """label 那行之后、下一节之前，返回带勾选标记☑那行的文字；没有则返回紧邻下一行。"""
    for i, ln in enumerate(lines):
        if ln.strip().startswith(label):
            first = ""
            for ln2 in lines[i + 1:i + 1 + maxscan]:
                s = ln2.strip()
                if not first and _clean(s):
                    first = _clean(s)
                if CHECK_ON in ln2:
                    return _clean(s)
            return first
    return ""

# 末尾体积标注，如 (249KB) / (1.2 MB)，提取附件名时去掉
_SIZE_TAIL = re.compile(r"\s*\([\d.]+\s*[KMGkmg]?B\)\s*$")
def _attachment_name(lines):
    """从“附件”那行取附件文件名：去行首“附件”标签、去 PUA、去末尾体积标注。"""
    for ln in lines:
        s = ln.strip()
        if s.startswith("附件"):
            s = _clean(s[len("附件"):])          # 去标签 + PUA
            return _SIZE_TAIL.sub("", s).strip()  # 去 (249KB) 之类
    return ""

def parse_form(text):
    """输入 PDF 首页文字，返回 {字段: 值}（只含解析到的，键与前端 info 一致）。"""
    lines = text.split("\n")
    out = {}

    def find(prefix):
        for ln in lines:
            if ln.strip().startswith(prefix):
                return ln.strip()
        return ""

    # 申请单编号 -> 委托单号
    m = re.search(r"申请单编号\s*(\S+)", find("申请单编号"))
    if m: out["commission_no"] = _clean(m.group(1))
    # 样品名称
    m = re.match(r"样品名称\s*(.+)", find("样品名称"))
    if m: out["sample_name"] = _clean(m.group(1))
    # 样品型号 + 零件号（同一行）
    m = re.match(r"样品型号\s*(.+?)\s+样品零件号\s*(.+)", find("样品型号"))
    if m:
        out["sample_model"] = _clean(m.group(1))
        out["sample_no"] = _clean(m.group(2))
    # 样本数量 -> 样品数量（值夹在“样本数量”和“辅材/配件信息”之间）
    qline = find("样本数量")
    if qline:
        m = re.match(r"样本数量\s*(.+?)(?:\s+辅材.*)?$", qline)
        if m and _clean(m.group(1)):
            out["sample_qty"] = _clean(m.group(1))
    # 验证阶段（单选）
    vline = find("验证阶段")
    if vline:
        sel, _ = _selected_radio(vline)   # 只取选中的选项，后面的说明文字不要
        if sel:
            out["verify_phase"] = sel
    # 委托方名称 / 地址（勾选）
    cn = _strip_opt(_checked_value(lines, "委托方名称")); ca = _strip_opt(_checked_value(lines, "委托方地址"))
    if cn: out["client_name"] = cn
    if ca: out["client_addr"] = ca
    # 制造商：勾了“同委托方”则同委托方，否则读独立行
    mline = find("制造商信息选择")
    same = "同委托方" in mline and CHECK_ON in mline and mline.split("同委托方")[1][:1] == CHECK_ON
    if same:
        if cn: out["maker_name"] = cn
        if ca: out["maker_addr"] = ca
    else:
        mn = _strip_opt(_line_after(lines, "制造商名称")); ma = _strip_opt(_line_after(lines, "制造商地址"))
        if mn: out["maker_name"] = mn
        if ma: out["maker_addr"] = ma
    # 检测项目 / 检测依据 = “参考”+附件全名（拼接，不带“+”号）
    att = _attachment_name(lines)
    if att:
        out["test_items"] = "参考" + att
        out["test_basis"] = "参考" + att
    return out

def parse_pdf(path_or_bytes):
    """从 PDF 路径或字节流解析首页。返回 dict。"""
    from pypdf import PdfReader
    import io
    src = io.BytesIO(path_or_bytes) if isinstance(path_or_bytes, (bytes, bytearray)) else path_or_bytes
    reader = PdfReader(src)
    text = reader.pages[0].extract_text() or ""
    return parse_form(text)
