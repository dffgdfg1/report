# -*- coding: utf-8 -*-
"""Build a portable Windows folder containing the local EXE and templates."""
import os
import shutil
import subprocess
import sys


BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIST = os.path.join(BASE, "dist", "报告生成器")
WORK = os.path.join(BASE, "构建", "pyinstaller-work")
SPEC = os.path.join(BASE, "构建", "pyinstaller-spec")


def main():
    shutil.rmtree(DIST, ignore_errors=True)
    cmd = [
        sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--onefile",
        "--name", "报告生成器", "--paths", os.path.join(BASE, "app"),
        "--distpath", DIST, "--workpath", WORK, "--specpath", SPEC,
        "--add-data", os.path.join(BASE, "app", "templates") + ";templates",
        "--add-data", os.path.join(BASE, "app", "static") + ";static",
        "--hidden-import", "docxcompose.composer",
        "--hidden-import", "win32timezone",
        os.path.join(BASE, "构建", "exe_launcher.py"),
    ]
    subprocess.check_call(cmd, cwd=BASE)

    shutil.copytree(os.path.join(BASE, "模板库"), os.path.join(DIST, "模板库"),
                    ignore=shutil.ignore_patterns("_probe.py", "__pycache__", "*.pyc"))
    for folder in ("项目", "输出", "构建"):
        os.makedirs(os.path.join(DIST, folder), exist_ok=True)
    print("Build ready:", DIST)


if __name__ == "__main__":
    main()
