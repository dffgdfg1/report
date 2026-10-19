#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Gunicorn 入口点：导入 app/app.py 中的 Flask 应用实例
"""
import sys
import os

# 将 app 目录添加到 Python 路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'app'))

# 导入真正的 Flask 应用实例
from app import app

if __name__ == "__main__":
    # 直接运行此文件时，使用 Flask 开发服务器
    app.run()
