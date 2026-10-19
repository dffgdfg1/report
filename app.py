#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Gunicorn 入口点：导入 app/app.py 中的 Flask 应用实例
"""
import sys
import os

# 获取项目根目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
APP_DIR = os.path.join(BASE_DIR, 'app')

# 确保 app 目录在路径开头
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

# 直接导入 app 模块中的 app 对象
# 使用明确的模块导入来避免循环依赖问题
import importlib.util
spec = importlib.util.spec_from_file_location("app_module", os.path.join(APP_DIR, "app.py"))
app_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app_module)
app = app_module.app

if __name__ == "__main__":
    # 直接运行此文件时，使用 Flask 开发服务器
    app.run()
