# -*- coding: utf-8 -*-
"""首次运行检测依赖，缺失则优先用本地 依赖/ 离线安装，失败再联网兜底。"""
import sys, os, subprocess, importlib.util

# 需要的顶层模块名 -> pip 包名
NEEDS = {
    "flask": "flask",
    "docx": "python-docx",
    "PIL": "Pillow",
    "win32com": "pywin32",
    "lxml": "lxml",
    "openpyxl": "openpyxl",
    "pypdf": "pypdf",
}

def missing():
    miss = []
    for mod, pkg in NEEDS.items():
        if importlib.util.find_spec(mod) is None:
            miss.append(pkg)
    return miss

def main():
    miss = missing()
    if not miss:
        return 0
    print("=" * 52)
    print("  首次运行：正在安装所需组件（约 1 分钟，仅这一次）")
    print("  缺少：", ", ".join(miss))
    print("=" * 52)
    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    dep = os.path.join(base, "依赖")
    py = sys.executable
    # 1) 优先离线安装
    if os.path.isdir(dep):
        try:
            subprocess.check_call([py, "-m", "pip", "install",
                                   "--no-index", "--find-links", dep] + miss)
        except Exception:
            pass
    # 2) 仍缺则联网兜底
    miss2 = missing()
    if miss2:
        print("离线安装未完全成功，尝试联网安装：", ", ".join(miss2))
        try:
            subprocess.check_call([py, "-m", "pip", "install"] + miss2)
        except Exception:
            pass
    # 3) pywin32 需要跑一次 post-install（注册 COM，供 WPS 刷新目录用）
    if "pywin32" in miss:
        try:
            subprocess.call([py, os.path.join(os.path.dirname(py), "Scripts", "pywin32_postinstall.py"), "-install"])
        except Exception:
            pass
    left = missing()
    if left:
        print("\n[提示] 以下组件仍未装好：", ", ".join(left))
        print("目录页码自动刷新等功能可能受影响，但报告仍可生成。")
        return 1
    print("组件安装完成！\n")
    return 0

if __name__ == "__main__":
    sys.exit(main())
