# -*- coding: utf-8 -*-
"""把工具打包成一个干净的分发 zip，放到桌面。
包含：app代码、模板库、离线依赖、启动.bat、说明。
排除：你自己的项目数据/输出/备份/开发脚本/缓存。"""
import os, shutil, zipfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STAGE = os.path.join(os.path.dirname(BASE), "_打包暂存_报告生成器")
ZIP = os.path.join(os.path.expanduser("~"), "Desktop", "试验报告生成器_分发版.zip")

INCLUDE_DIRS = ["app", "模板库", "依赖"]
INCLUDE_FILES = ["启动.bat", "分发说明_给同事.txt", "使用说明.txt"]
# app 内要跳过的
SKIP = {"__pycache__", ".pyc"}

def copy_tree(src, dst):
    for root, dirs, files in os.walk(src):
        dirs[:] = [d for d in dirs if d not in SKIP]
        rel = os.path.relpath(root, src)
        os.makedirs(os.path.join(dst, rel), exist_ok=True)
        for f in files:
            if f.endswith(".pyc"):
                continue
            shutil.copy2(os.path.join(root, f), os.path.join(dst, rel, f))

def main():
    if os.path.exists(STAGE):
        shutil.rmtree(STAGE)
    os.makedirs(STAGE)
    # 目录
    for d in INCLUDE_DIRS:
        s = os.path.join(BASE, d)
        if os.path.isdir(s):
            copy_tree(s, os.path.join(STAGE, d))
    # 把测试方案清空为干净起点（同事从空白方案库开始）
    sf = os.path.join(STAGE, "模板库", "测试方案.json")
    if os.path.exists(sf):
        with open(sf, "w", encoding="utf-8") as f:
            f.write("{}")
    # 建空的 项目/ 输出/ 占位
    for d in ["项目", "输出"]:
        os.makedirs(os.path.join(STAGE, d), exist_ok=True)
        with open(os.path.join(STAGE, d, "说明.txt"), "w", encoding="utf-8") as f:
            f.write("此文件夹用于存放数据，请勿删除。")
    # 文件
    for f in INCLUDE_FILES:
        s = os.path.join(BASE, f)
        if os.path.exists(s):
            shutil.copy2(s, os.path.join(STAGE, f))
    # 打 zip
    if os.path.exists(ZIP):
        os.remove(ZIP)
    with zipfile.ZipFile(ZIP, "w", zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(STAGE):
            for f in files:
                fp = os.path.join(root, f)
                arc = os.path.join("试验报告生成器", os.path.relpath(fp, STAGE))
                z.write(fp, arc)
    shutil.rmtree(STAGE)
    print("打包完成:", ZIP)
    print("大小: %.1f MB" % (os.path.getsize(ZIP) / 1048576))

if __name__ == "__main__":
    main()
