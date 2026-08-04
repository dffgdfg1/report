# -*- coding: utf-8 -*-
"""用 LibreOffice(无头模式) 打开 docx，更新目录(TOC)与所有域/页码后存回。
在 Linux 服务器上把页码算好写死进文件，用户用 WPS/Word 打开即成品、零操作。

用法（由 report_engine 以子进程调用，需用带 uno 的 python 运行）：
    <python-with-uno> toc_lo.py <docx路径> [soffice可执行路径]

每一步都打印带时间戳的日志到 stderr，便于定位卡点。
"""
import os
import sys
import time
import socket
import tempfile
import subprocess

_T0 = time.time()
def log(msg):
    sys.stderr.write("[toc_lo +%5.1fs] %s\n" % (time.time() - _T0, msg))
    sys.stderr.flush()

def _free_port():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def _find_soffice(explicit=None):
    if explicit and os.path.exists(explicit):
        return explicit
    import shutil
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in ("/usr/bin/soffice", "/usr/bin/libreoffice",
              "/opt/libreoffice/program/soffice", "/snap/bin/libreoffice"):
        if os.path.exists(p):
            return p
    return None


def _wait_port(port, proc, timeout=40):
    """轮询 TCP 端口是否已监听。比盲目 resolve 更快更准，且能及时发现 soffice 退出。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if proc.poll() is not None:
            raise RuntimeError("soffice 进程已退出，返回码=%s" % proc.returncode)
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(0.5)
        try:
            s.connect(("127.0.0.1", port))
            s.close()
            return True
        except Exception:
            s.close()
            time.sleep(0.3)
    return False


def update_toc(path, soffice_bin=None):
    import uno
    from com.sun.star.beans import PropertyValue

    path = os.path.abspath(path)
    log("目标文件: %s" % path)
    soffice_bin = _find_soffice(soffice_bin)
    if not soffice_bin:
        raise RuntimeError("找不到 soffice/libreoffice")
    log("soffice: %s" % soffice_bin)

    port = _free_port()
    profile = tempfile.mkdtemp(prefix="lo_toc_")
    profile_url = "file://" + profile.replace("\\", "/")
    log("端口=%d profile=%s" % (port, profile))

    # 关键：给 soffice 子进程一个可写的 HOME，避免在受限环境(systemd)下卡住
    env = dict(os.environ)
    env.setdefault("HOME", profile)

    log("启动 soffice(监听模式)…")
    proc = subprocess.Popen([
        soffice_bin, "--headless", "--invisible", "--nodefault",
        "--norestore", "--nologo", "--nofirststartwizard", "--nocrashreport",
        "-env:UserInstallation=" + profile_url,
        "--accept=socket,host=127.0.0.1,port=%d;urp;" % port,
    ], env=env)

    try:
        log("等待端口 %d 就绪…" % port)
        if not _wait_port(port, proc, timeout=40):
            raise RuntimeError("等待 soffice 端口超时")
        log("端口已就绪，建立 UNO 连接…")

        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx)
        conn = ("uno:socket,host=127.0.0.1,port=%d;urp;"
                "StarOffice.ComponentContext" % port)
        ctx = None
        last = None
        for _ in range(20):
            try:
                ctx = resolver.resolve(conn); break
            except Exception as e:
                last = e; time.sleep(0.5)
        if ctx is None:
            raise RuntimeError("UNO 连接失败: %s" % last)
        log("UNO 已连接，打开桌面服务…")

        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        url = uno.systemPathToFileUrl(path)
        # 加载参数：隐藏窗口 + 禁止跑宏 + 不弹更新链接对话框(避免 headless 卡死)
        def pv(n, v):
            p = PropertyValue(); p.Name = n; p.Value = v; return p
        props = (pv("Hidden", True), pv("MacroExecutionMode", 0), pv("UpdateDocMode", 0))
        log("加载文档中…")
        doc = desktop.loadComponentFromURL(url, "_blank", 0, props)
        log("文档已加载，更新目录/索引…")

        try:
            idxs = doc.getDocumentIndexes()
            log("索引数量=%d" % idxs.getCount())
            for i in range(idxs.getCount()):
                idxs.getByIndex(i).update()
        except Exception as e:
            log("更新索引出错(忽略): %s" % e)
        try:
            doc.getTextFields().refresh()
        except Exception as e:
            log("刷新文本域出错(忽略): %s" % e)
        try:
            doc.refresh()
        except Exception as e:
            log("全局刷新出错(忽略): %s" % e)

        log("保存回 docx…")
        doc.storeToURL(url, (pv("FilterName", "MS Word 2007 XML"), pv("Overwrite", True)))
        log("已保存，关闭文档…")
        try:
            doc.close(False)
        except Exception:
            pass
        try:
            desktop.terminate()
        except Exception:
            pass
    finally:
        _kill(proc, profile)


def _kill(proc, profile):
    try:
        proc.terminate()
        try:
            proc.wait(timeout=8)
        except Exception:
            proc.kill()
    except Exception:
        pass
    try:
        import shutil
        shutil.rmtree(profile, ignore_errors=True)
    except Exception:
        pass


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.stderr.write("用法: toc_lo.py <docx路径> [soffice路径]\n")
        sys.exit(2)
    p = sys.argv[1]
    sof = sys.argv[2] if len(sys.argv) > 2 else None
    try:
        update_toc(p, sof)
        log("完成 OK")
        print("OK")
        sys.exit(0)
    except Exception as e:
        log("失败: %s" % e)
        sys.exit(1)
