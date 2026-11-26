# -*- coding: utf-8 -*-
"""Start the local report generator from a PyInstaller distribution."""
import os
import socket
import sys
import threading
import webbrowser


if getattr(sys, "frozen", False):
    HOME = os.path.dirname(sys.executable)
    APP_DIR = os.path.join(sys._MEIPASS, "app")
else:
    HOME = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    APP_DIR = os.path.join(HOME, "app")

os.environ["REPORT_GENERATOR_HOME"] = HOME
sys.path.insert(0, APP_DIR)

import app as app_module


def _available_port(preferred):
    for port in (preferred, 0):
        with socket.socket() as probe:
            try:
                probe.bind(("127.0.0.1", port))
                return probe.getsockname()[1]
            except OSError:
                continue
    raise RuntimeError("无法分配本地服务端口")


if __name__ == "__main__":
    port = _available_port(app_module.PORT)
    app_module.PORT = port
    threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:%d/" % port)).start()
    print("报告生成器已启动: http://127.0.0.1:%d/" % port)
    print("关闭此窗口即可停止服务。")
    app_module.app.run(host=app_module.HOST, port=port, debug=False, threaded=True, use_reloader=False)
