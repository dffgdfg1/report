# -*- coding: utf-8 -*-
"""测试类型预设：新增测试类型只需在此加一条。图组标题决定图片分几组。"""

# 内置预设已清空。测试类型改为使用"已保存方案"（用户自存，见 测试方案.json）。
PRESETS = {}

# 通用空白类型（模板未收录的测试）
GENERIC = {
    "standard": "参考客户要求",
    "env": "18℃-28℃、25%RH-75%RH",
    "condition": "",
    "requirement": "功能等级A。",
    "equipment": [],
    "image_group_titles": ["试验前图片", "试验中图片", "试验后图片"],
}

def get_preset(name):
    return PRESETS.get(name, GENERIC)

def list_types():
    return list(PRESETS.keys())
