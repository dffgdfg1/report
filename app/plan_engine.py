# -*- coding: utf-8 -*-
"""试验计划生成引擎：从「试验计划模板.xlsx」按测试项填出一张试验计划表。

模板结构（第一个 sheet「试验计划」）：
- 行1  标题「试验计划」（合并 C1:N1）
- 行2  表头信息：产品名称/型号、试验阶段、试验日期、计划编制人、计划核准、客户批准
- 行3~4 列标题（No./试验项目/标准号/试验标准/判定依据/样机数量/试验周期/
        计划时间(开始,结束)/实际时间(开始,结束)/试验结果/试验机构/备注）
- 行5  表格编号行（整行合并）
- 行6  分组行「第N组（样机号）」（整行合并）
- 行7  数据行样例

数据来自报告项目 project = {info:{...}, tests:[{...}]}。
"""
import os
import copy
import openpyxl
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(BASE, "模板库", "试验计划模板.xlsx")

# 数据行列 -> (取值函数)；1-based 列号
NCOLS = 14
DATA_ROW_TEMPLATE = 7   # 模板里数据行样例的行号
GROUP_ROW_TEMPLATE = 6  # 模板里分组行样例的行号
FORMNO_ROW = 5          # 表格编号行


def _fmt_date(s):
    """日期原样输出（报告里已是 2026.08.14 点分格式）。"""
    return str(s or "").strip()


def _days_between(a, b):
    """由起止日期估算试验周期天数（含首尾）。解析失败返回空串。"""
    import re
    def parse(s):
        m = re.findall(r"\d+", str(s or ""))
        if len(m) >= 3:
            return tuple(int(x) for x in m[:3])
        return None
    pa, pb = parse(a), parse(b)
    if not pa or not pb:
        return ""
    import datetime
    try:
        da = datetime.date(*pa); db = datetime.date(*pb)
        n = (db - da).days + 1
        return f"{n}天" if n > 0 else ""
    except Exception:
        return ""


def _copy_style(src_cell, dst_cell):
    """复制单元格样式（字体/边框/对齐/填充/数字格式）。"""
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy.copy(src_cell.protection)


def _set_header(ws, info, applicant):
    """填第2行表头信息，保留原有「标签：」前缀，只把值接到冒号后。"""
    name = (info.get("sample_name", "") or "").strip()
    model = (info.get("sample_model", "") or "").strip()
    no = (info.get("sample_no", "") or "").strip()
    phase = (info.get("verify_phase", "") or "").strip()
    model_no = model + ("/" + no if no else "")

    def put(cell_ref, label, value):
        cell = ws[cell_ref]
        raw = str(cell.value or "")
        if not value:
            return  # 无值：保留模板原样（标签+空白）
        # 保留原首尾换行样式（模板里表头带前后换行做垂直留白）
        lead = "\n" if raw.startswith("\n") else ""
        tail = "\n" if raw.endswith("\n") else ""
        # 取「标签：」前缀（去掉原有换行后再统一补回，避免换行叠加）
        core = raw.strip()
        prefix = core
        for sep in ("：", ":"):
            if sep in core:
                prefix = core.split(sep)[0] + sep
                break
        cell.value = f"{lead}{prefix} {value}{tail}"

    put("A2", "产品名称/零件名称：", name)
    put("D2", "产品型号/零件号：", model_no)
    put("E2", "试验阶段：", phase)
    # F2 试验日期、I2 计划编制人、K2 计划核准：按需求留空（保留标签）
    put("M2", "客户批准：", applicant)


def _write_data_rows(ws, tests):
    """把 tests 写进数据区。分组行样例(行6)+数据行样例(行7)作为样式来源。
    按每个测试项的样机号(sample_no)分组：同组共用一个「第N组（样机号）」标题。"""
    # 记录样式来源
    grp_styles = [ws.cell(GROUP_ROW_TEMPLATE, c) for c in range(1, NCOLS + 1)]
    dat_styles = [ws.cell(DATA_ROW_TEMPLATE, c) for c in range(1, NCOLS + 1)]
    grp_h = ws.row_dimensions[GROUP_ROW_TEMPLATE].height
    dat_h = ws.row_dimensions[DATA_ROW_TEMPLATE].height

    # 清掉模板里的分组行(6)与数据行样例(7)内容和它们的合并
    for mc in list(ws.merged_cells.ranges):
        if mc.min_row >= GROUP_ROW_TEMPLATE:
            ws.unmerge_cells(str(mc))
    # 删除第6、7行（样例），后面从第6行起重建
    ws.delete_rows(GROUP_ROW_TEMPLATE, 2)

    # 按样机号分组，保持首次出现顺序
    groups = []           # [(sample_no, [test,...])]
    index = {}
    for t in tests:
        key = (t.get("sample_no", "") or "").strip()
        if key not in index:
            index[key] = len(groups)
            groups.append((key, []))
        groups[index[key]][1].append(t)

    # 清掉模板中数据样例以下遗留的空行内容（模板原本排到第15行左右）
    last = ws.max_row
    for rr in range(GROUP_ROW_TEMPLATE, last + 1):
        for cc in range(1, NCOLS + 1):
            ws.cell(rr, cc).value = None

    r = GROUP_ROW_TEMPLATE  # 从第6行开始重建
    seq = 1
    for gi, (sample_no, gtests) in enumerate(groups, 1):
        # 分组标题行：整行合并，「第N组（样机号）」
        for c in range(1, NCOLS + 1):
            cell = ws.cell(r, c)
            _copy_style(grp_styles[c - 1], cell)
        ws.cell(r, 1).value = f"第{gi}组（{sample_no}）"
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NCOLS)
        if grp_h:
            ws.row_dimensions[r].height = grp_h
        r += 1
        # 该组各测试项数据行
        for t in gtests:
            vals = _row_values(t, seq)
            for c in range(1, NCOLS + 1):
                cell = ws.cell(r, c)
                _copy_style(dat_styles[c - 1], cell)
                cell.value = vals[c - 1]
            # 数据行高留空 -> 打开时按内容自动调整（试验标准/方法可能很长）
            r += 1
            seq += 1


def _row_values(t, seq):
    """返回一行 14 列的值列表。"""
    samples = t.get("samples", []) or []
    qty = len(samples) if samples else ""
    start = _fmt_date(t.get("start_date", ""))
    end = _fmt_date(t.get("end_date", ""))
    cycle = _days_between(start, end)
    return [
        seq,                                  # 1 No.
        t.get("title", ""),                   # 2 试验项目
        t.get("standard", ""),                # 3 标准号/条款号
        t.get("condition", ""),               # 4 试验标准/方法
        t.get("requirement", ""),             # 5 判定依据
        qty,                                  # 6 样机数量
        cycle,                                # 7 试验周期
        "",                                   # 8 计划时间-开始（留空）
        "",                                   # 9 计划时间-结束（留空）
        start,                                # 10 实际时间-开始
        end,                                  # 11 实际时间-结束
        t.get("overall_result", ""),          # 12 试验结果
        "minieye",                            # 13 试验机构
        "",                                   # 14 备注
    ]


def generate_plan(project, out_path):
    """project: {info:{...}, tests:[...]} -> 生成试验计划 .xlsx，返回文件路径。"""
    info = project.get("info", {}) or {}
    tests = project.get("tests", []) or []
    applicant = (info.get("applicant", "") or "").strip()

    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb["试验计划"] if "试验计划" in wb.sheetnames else wb.active

    _set_header(ws, info, applicant)
    if tests:
        _write_data_rows(ws, tests)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)
    return out_path
