# -*- coding: utf-8 -*-
"""原始记录生成引擎：从「原始记录骨架.docx」按测试项逐个生成原始记录。

每个试验项目生成一张原始记录表（15 行 4 列）：
- 样品基础信息来自首页 info（可由申请单自动导入）
- 仪器设备复用设备库（test.equipment）
- 试验项目/标准/条件/要求复用试验项目库（test）

多个测试项时，每张记录之间插入分页符，一份文档包含全部原始记录。
"""
import os
import copy
from docx import Document
from docx.oxml.ns import qn

# 复用主报告引擎的低层助手，保证字体/换行处理一致（宋体五号、\r\n 归一等）
import report_engine as E

W = E.W

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKELETON = os.path.join(BASE, "模板库", "原始记录骨架.docx")


def _row_cells(row_el):
    return row_el.findall(W + 'tc')


def _set(tc, text):
    """填单元格文本，沿用主引擎的换行/字体处理。"""
    E.set_tc_text(tc, "" if text is None else str(text))


def _join(items):
    """多台设备等多值信息按换行拼接（模板里就是多行显示）。"""
    return "\n".join(s for s in items if str(s or "").strip())


def _fill_table(tbl_el, info, test, doc):
    """填一张原始记录表（15 行）。tbl_el 为 <w:tbl> 元素。"""
    rows = tbl_el.findall(W + 'tr')

    def cells(ri):
        return _row_cells(rows[ri])

    equ = test.get("equipment", []) or []
    eq_names = _join([e.get("name", "") for e in equ])
    eq_models = _join([e.get("model", "") for e in equ])
    eq_nos = _join([e.get("mgmt_no", "") for e in equ])
    eq_cal = _join([e.get("cal_valid", "") for e in equ])

    # R0 样品名称 / 样品型号
    c = cells(0); _set(c[1], info.get("sample_name", "")); _set(c[3], info.get("sample_model", ""))
    # R1 样品零件号 / 委托单号
    c = cells(1); _set(c[1], info.get("sample_no", "")); _set(c[3], info.get("commission_no", ""))
    # R2 样品数量 / 样品编号（样品编号取该测试项的样机编号）
    c = cells(2); _set(c[1], info.get("sample_qty", "")); _set(c[3], test.get("sample_no", ""))
    # R3 额定电压 / 环境条件
    c = cells(3); _set(c[1], info.get("rated_volt", "")); _set(c[3], test.get("env", ""))
    # R4 设备名称 / 设备编号
    c = cells(4); _set(c[1], eq_names); _set(c[3], eq_nos)
    # R5 设备型号 / 校准周期
    c = cells(5); _set(c[1], eq_models); _set(c[3], eq_cal)
    # R6 试验项目 / 试验标准
    c = cells(6); _set(c[1], test.get("title", "")); _set(c[3], test.get("standard", ""))
    # R7 试验条件（跨列）：文字上下居中，其下追加试验条件配图
    _cond_text = test.get("condition", "")
    _cond_imgs = test.get("condition_images", [])
    c = cells(7); _set(c[1], _cond_text); _vcenter(c[1])
    _append_cond_images(c[1], doc, _cond_imgs)
    # 仅当试验条件内容较多（有配图或文字超阈值）时，才分页+撕框；
    # 内容少时保持默认行高、不插分页符，避免产生空白页。
    _need_split = bool(_cond_imgs) or _condition_is_long(_cond_text)
    if _need_split:
        h7 = _page1_r7_height(doc, rows)
        if h7 > 0:
            _row_min_height(rows[7], h7)
    # R8 试验要求（跨列）
    c = cells(8); _set(c[1], test.get("requirement", ""))
    if _need_split:
        _page_break_before_cell(c[0])
    # R10 试验状态：仅分页模式下才拉高
    if _need_split:
        h10 = _page2_status_height(doc, rows)
        if h10 > 0:
            _row_min_height(rows[10], h10)
        _row_min_height(rows[10], h10)
    # R9~R12、R14 为签字/状态/判定/备注等手填栏，保持模板原样不动
    # R13 测试日期：把开始/结束时间填进模板那句固定格式的话术里
    _fill_test_date(cells(13)[1], test)


def _vcenter(tc):
    """单元格内容垂直居中（试验条件文字上下居中）。"""
    from docx.oxml.ns import qn
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = tc.makeelement(qn('w:tcPr'), {})
        tc.insert(0, tcPr)
    vA = tcPr.find(qn('w:vAlign'))
    if vA is None:
        vA = tcPr.makeelement(qn('w:vAlign'), {})
        tcPr.append(vA)
    vA.set(qn('w:val'), 'center')


def _row_min_height(row_el, twips):
    """给行设最小行高（atLeast），使单元格纵向拉高填满页面。twips：1cm≈567。"""
    trPr = row_el.find(qn('w:trPr'))
    if trPr is None:
        trPr = row_el.makeelement(qn('w:trPr'), {})
        row_el.insert(0, trPr)
    trH = trPr.find(qn('w:trHeight'))
    if trH is None:
        trH = trPr.makeelement(qn('w:trHeight'), {})
        trPr.append(trH)
    trH.set(qn('w:val'), str(int(twips)))
    trH.set(qn('w:hRule'), 'atLeast')


def _usable_text_height(doc):
    """正文可用高度(twips) = 页高 - 上下页边距。"""
    sec = doc.sections[0]
    try:
        return int(sec.page_height.twips) - int(sec.top_margin.twips) - int(sec.bottom_margin.twips)
    except Exception:
        return 14500  # A4 常见值兜底


# 模板各行默认行高(twips)，当行元素没有显式 w:trHeight 时用作兆底
_DEFAULT_ROW_HEIGHTS = {
    0: 510, 1: 510, 2: 510, 3: 510,
    4: 637, 5: 680, 6: 510, 7: 2283,
    8: 631, 9: 567, 10: 3331, 11: 601,
    12: 400, 13: 417, 14: 455,
}


def _rows_height_sum(rows, idxs):
    """若干行的最小行高(trHeight)之和(twips)。
    有显式 w:trHeight 用显式值；否则回退到模板默认行高。"""
    tot = 0
    for i in idxs:
        val = None
        trPr = rows[i].find(qn('w:trPr'))
        if trPr is not None:
            trH = trPr.find(qn('w:trHeight'))
            if trH is not None and trH.get(qn('w:val')):
                try:
                    val = int(trH.get(qn('w:val')))
                except Exception:
                    pass
        if val is None:
            val = _DEFAULT_ROW_HEIGHTS.get(i, 0)
        tot += val
    return tot


# 标题(试验记录表, sz30)+编号行(sz24)+段间距，实测约占这么多 twips；
# 加一点安全余量，宁可 R7 稍矮留白，也不要溢出把整行挤到下一页。
_HEADER_PARA_TWIPS = 1400
_PAGE1_SAFETY_TWIPS = 300


def _condition_is_long(text):
    """判断试验条件文字是否足够长，需要分页排版。
    有3行及以上换行，或总字符数超 150，视为「较长」。"""
    if not text:
        return False
    if text.count('\n') >= 3:
        return True
    return len(text) > 150






def _page1_r7_height(doc, rows):
    """R7(试验条件)在第一页能占的最大行高：
    可用高度 - 标题/编号 - 信息行(R0~R6) - 安全余量。
    这样 R7 拉到最大又不会脱离第一页。"""
    usable = _usable_text_height(doc)
    info_h = _rows_height_sum(rows, range(0, 7))
    h = usable - _HEADER_PARA_TWIPS - info_h - _PAGE1_SAFETY_TWIPS
    return max(0, h)


def _page2_status_height(doc, rows):
    """R10(试验状态)在分页模式下的合理高度。
    不撑满整页（会把 R11~R14 挤走产生空白），而是固定一个较大值。"""
    # 固定拉高到约 6000 twips (约 10.5cm)，足够容纳较多内容，
    # 又不至于把签字行挤到单独页造成空白。
    return 6000


def _page_break_before_cell(tc):
    """给单元格首段加 pageBreakBefore，使该行从新一页开始。"""
    p = tc.find(qn('w:p'))
    if p is None:
        return
    ppr = p.find(qn('w:pPr'))
    if ppr is None:
        ppr = p.makeelement(qn('w:pPr'), {})
        p.insert(0, ppr)
    if ppr.find(qn('w:pageBreakBefore')) is None:
        ppr.insert(0, ppr.makeelement(qn('w:pageBreakBefore'), {}))


def _append_cond_images(tc, doc, imgs):
    """在试验条件文字下方追加配图：每行横排 2 张，避免竖排把版面撑乱。
    在单元格内嵌一张 2 列表格承载图片（无边框、居中）。"""
    if not imgs:
        return
    from docx.table import Table, _Cell
    cell = _Cell(tc, Table(tc.getparent().getparent(), doc))
    nrows = (len(imgs) + 1) // 2
    nested = cell.add_table(rows=nrows, cols=2)
    try:
        nested.autofit = True
    except Exception:
        pass
    _no_borders(nested._tbl)
    for idx, im in enumerate(imgs):
        r, c = idx // 2, idx % 2
        ncell = nested.cell(r, c)
        p = ncell.paragraphs[0]
        p.alignment = 1  # center
        run = p.add_run()
        try:
            stream, size = E.normalize_image(im["path"])
            w, h = E._target_size(size)
            run.add_picture(stream, width=w, height=h)
        except Exception:
            w, h = E._target_size(None)
            run.add_picture(im["path"], width=w, height=h)
        cap = im.get("caption", "")
        if cap:
            cp = ncell.add_paragraph(cap)
            cp.alignment = 1
            for rr in cp.runs:
                E.force_song5(rr._r)


def _no_borders(tbl_el):
    """把内嵌表格的四周及内部边框全部设为 none。"""
    tblPr = tbl_el.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = tbl_el.makeelement(qn('w:tblPr'), {})
        tbl_el.insert(0, tblPr)
    borders = tblPr.find(qn('w:tblBorders'))
    if borders is None:
        borders = tblPr.makeelement(qn('w:tblBorders'), {})
        tblPr.append(borders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = borders.find(qn('w:' + edge))
        if e is None:
            e = borders.makeelement(qn('w:' + edge), {})
            borders.append(e)
        e.set(qn('w:val'), 'none')
        e.set(qn('w:sz'), '0')
        e.set(qn('w:space'), '0')


def _fill_test_date(tc, test):
    """R13：模板文本为『测试开始时间：   测试结束时间：   时间：  小时』，
    把开始/结束日期插到对应标签后面，保留其余格式与空格。"""
    sd = str(test.get("start_date", "") or "").strip()
    ed = str(test.get("end_date", "") or "").strip()
    if not sd and not ed:
        return
    cur = "".join(n.text or "" for n in tc.iter(W + 't'))
    if "测试开始时间" not in cur:
        # 模板话术若被改动，退化为直接写区间
        _set(tc, "测试开始时间：%s   测试结束时间：%s" % (sd, ed))
        return
    new = cur
    new = new.replace("测试开始时间：", "测试开始时间：%s " % sd, 1)
    new = new.replace("测试结束时间：", "测试结束时间：%s " % ed, 1)
    _set(tc, new)


def _fill_record_no(doc, no):
    """把正文「编号：」段落补成「编号：<委托单号>」，保留原字体。"""
    no = str(no or '').strip()
    if not no:
        return
    for p in doc.paragraphs:
        txt = ''.join(r.text or '' for r in p.runs)
        if txt.strip().startswith('编号'):
            runs = p.runs
            if not runs:
                continue
            # 保留首个 run 的「编号：」标签，把单号接到其后
            base = runs[0].text or ''
            if '：' in base:
                runs[0].text = base.split('：')[0] + '：' + no
            else:
                runs[0].text = base + no
            for r in runs[1:]:
                r.text = ''
            break


def generate_raw_records(project, out_path):
    """project: {info:{...}, tests:[{...}]} -> 生成 out_path(.docx)。
    每个测试项一张原始记录表，多项之间分页。"""
    info = project.get("info", {}) or {}
    tests = project.get("tests", []) or []
    if not tests:
        tests = [{}]  # 没有测试项也产出一张空表，避免空文档

    doc = Document(SKELETON)
    # 「编号：」段落填委托单号
    _fill_record_no(doc, info.get("commission_no", ""))
    tmpl_tbl = doc.tables[0]._tbl
    # 先留一份未填写的空白模板副本，供后续每张记录克隆（否则会克隆到已填数据）
    blank_tbl = copy.deepcopy(tmpl_tbl)

    # 第一张：直接填模板里已有的表
    _fill_table(tmpl_tbl, info, tests[0], doc)
    last_el = tmpl_tbl

    # 其余每张：分页符 + 克隆空白模板表再填
    for t in tests[1:]:
        pbreak = _make_page_break(doc)
        last_el.addnext(pbreak)
        new_tbl = copy.deepcopy(blank_tbl)
        pbreak.addnext(new_tbl)
        _fill_table(new_tbl, info, t, doc)
        last_el = new_tbl

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    return out_path


def _make_page_break(doc):
    """构造一个只含分页符的段落元素。"""
    p = doc.element.body.makeelement(W + 'p', {})
    r = p.makeelement(W + 'r', {})
    br = r.makeelement(W + 'br', {})
    br.set(qn('w:type'), 'page')
    r.append(br)
    p.append(r)
    return p
