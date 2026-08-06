# -*- coding: utf-8 -*-
"""目录更新服务（在装了 WPS/Word 的 Windows 机器上运行）。

用途：Linux 报告服务器生成 docx 后，把文件 POST 到本服务，本服务用
WPS/Word COM 在 1-2 秒内把目录(TOC)与所有域/页码刷新好，再把成品返回。
这样绕开慢虚拟机上 LibreOffice 要 4 分钟的问题。

接口：
    POST /update_toc
        头：X-Token: <共享口令>
        体：docx 二进制
        返回：更新后的 docx 二进制（失败返回 4xx/5xx + 文本原因）
    GET /health  -> 200 "ok"（探活，不需口令）

安全：
    - 共享口令校验（口令由 TOC_SERVICE_TOKEN 环境变量或 --token 指定）
    - COM 打开文件时强制禁用宏（防带宏的恶意 docx）
    - 建议只在局域网内使用
"""
import os
import sys
import time
import tempfile
import threading

from flask import Flask, request, Response

app = Flask(__name__)

# WPS/Word 同一时刻只能稳定开一个自动化实例，用锁把并发请求排队
_COM_LOCK = threading.Lock()

TOKEN = ""   # 由 main 设置

def _update_com(path):
    """用 WPS(优先)/Word 打开 docx，刷新目录与所有域后保存。禁用宏。
    每次都新建实例、用完退出，避免残留进程。"""
    import pythoncom
    pythoncom.CoInitialize()
    application = None
    used = None
    try:
        import win32com.client as win32
        # DispatchEx 强制新建独立进程：否则 Dispatch 会附着到用户已打开的
        # WPS/Word 实例，退出时把用户正在看的文档一起关掉。
        try:
            application = win32.DispatchEx("KWps.Application"); used = "WPS"
        except Exception:
            try:
                application = win32.DispatchEx("Word.Application"); used = "Word"
            except Exception:
                # 极少数环境 DispatchEx 不可用时退回 Dispatch（下面用文档计数兜底保护）
                try:
                    application = win32.Dispatch("KWps.Application"); used = "WPS"
                except Exception:
                    application = win32.Dispatch("Word.Application"); used = "Word"
        try: application.Visible = False
        except Exception: pass
        try: application.DisplayAlerts = False
        except Exception: pass
        # 强制禁用宏：msoAutomationSecurityForceDisable = 3
        try: application.AutomationSecurity = 3
        except Exception: pass

        doc = application.Documents.Open(path)
        try:
            for toc in doc.TablesOfContents:
                toc.Update()
        except Exception:
            pass
        try:
            doc.Fields.Update()
        except Exception:
            pass
        try:
            for sec in doc.Sections:
                for hf in (sec.Headers, sec.Footers):
                    for h in hf:
                        h.Range.Fields.Update()
        except Exception:
            pass
        doc.Save()
        doc.Close(False)
        doc = None
        return used
    finally:
        try:
            if application is not None:
                # 兜底保护：万一附着到了用户的实例，只有在没有其它文档打开时才退出，
                # 避免关掉用户正在编辑的文档。
                try:
                    remaining = application.Documents.Count
                except Exception:
                    remaining = 0
                if remaining <= 0:
                    application.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


@app.route("/health", methods=["GET"])
def health():
    return Response("ok", mimetype="text/plain")


@app.route("/update_toc", methods=["POST"])
def update_toc():
    # 口令校验
    if TOKEN and request.headers.get("X-Token", "") != TOKEN:
        return Response("bad token", status=403, mimetype="text/plain")
    data = request.get_data()
    if not data:
        return Response("empty body", status=400, mimetype="text/plain")

    tmpdir = tempfile.mkdtemp(prefix="toc_svc_")
    fp = os.path.join(tmpdir, "in.docx")
    with open(fp, "wb") as f:
        f.write(data)

    t0 = time.time()
    with _COM_LOCK:   # 串行处理，WPS/Word 一次只开一个
        try:
            used = _update_com(fp)
        except Exception as e:
            _cleanup(tmpdir)
            sys.stderr.write("COM 更新失败: %s\n" % e)
            return Response("com failed: %s" % e, status=500, mimetype="text/plain")

    try:
        with open(fp, "rb") as f:
            out = f.read()
    finally:
        _cleanup(tmpdir)

    dt = time.time() - t0
    print("[toc_service] %s 更新完成 %.1fs, %d bytes" % (used, dt, len(out)))
    resp = Response(out, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    resp.headers["X-Engine"] = used or "?"
    return resp


def _cleanup(d):
    try:
        import shutil
        shutil.rmtree(d, ignore_errors=True)
    except Exception:
        pass


def _lan_ips():
    ips = []
    try:
        import socket
        host = socket.gethostname()
        for info in socket.getaddrinfo(host, None):
            ip = info[4][0]
            if "." in ip and ip not in ips:
                ips.append(ip)
    except Exception:
        pass
    return ips


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--host", default="0.0.0.0")
    ap.add_argument("--port", type=int, default=int(os.environ.get("TOC_SERVICE_PORT", "8765")))
    ap.add_argument("--token", default=os.environ.get("TOC_SERVICE_TOKEN", ""))
    args = ap.parse_args()
    TOKEN = args.token

    print("=" * 56)
    print(" 目录更新服务已启动")
    print(" 监听: %s:%d" % (args.host, args.port))
    for ip in _lan_ips():
        print("   局域网地址: http://%s:%d" % (ip, args.port))
    print(" 口令: %s" % ("(已设置)" if TOKEN else "(局域网免口令)"))
    print(" 让此窗口保持开着；虚拟机会把报告发到这里刷新目录。")
    print("=" * 56)
    # 单线程即可：COM 本就串行；关掉 reloader 防重复起进程
    app.run(host=args.host, port=args.port, threaded=True, use_reloader=False)
