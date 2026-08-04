# -*- coding: utf-8 -*-
"""用 LibreOffice(无头模式) 打开 docx，更新目录(TOC)与所有域/页码后存回。
在 Linux 服务器上把页码算好写死进文件，用户用 WPS/Word 打开即成品、零操作。

用法（由 report_engine 以子进程调用，需用带 uno 的 python 运行）：
    <python-with-uno> toc_lo.py <docx路径> [soffice可执行路径]

设计要点：
- 自己拉起一个 headless soffice 监听 socket，用完杀掉；用独立临时 profile 避免与
  桌面实例/并发调用冲突。
- 全程异常都打印到 stderr 并以非 0 退出，调用方据此判断成败、失败则回退。
"""
import os
import sys
import time
import socket
import tempfile
import subprocess

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
    # 常见安装路径兜底
    for p in ("/usr/bin/soffice", "/usr/bin/libreoffice",
              "/opt/libreoffice/program/soffice",
              "/snap/bin/libreoffice"):
        if os.path.exists(p):
            return p
    return None


def update_toc(path, soffice_bin=None):
    import uno
    from com.sun.star.beans import PropertyValue

    path = os.path.abspath(path)
    soffice_bin = _find_soffice(soffice_bin)
    if not soffice_bin:
        raise RuntimeError("找不到 soffice/libreoffice 可执行文件")

    port = _free_port()
    profile = tempfile.mkdtemp(prefix="lo_toc_")
    profile_url = "file://" + profile.replace("\\", "/")
    conn = ("socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext" % port)

    proc = subprocess.Popen([
        soffice_bin, "--headless", "--invisible", "--nodefault",
        "--norestore", "--nologo", "--nofirststartwizard",
        "-env:UserInstallation=" + profile_url,
        "--accept=socket,host=127.0.0.1,port=%d;urp;" % port,
    ])

    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx)

    # 等 soffice 起来并能连上（最多约 60 秒）
    ctx = None
    last_err = None
    for _ in range(120):
        try:
            ctx = resolver.resolve("uno:" + conn)
            break
        except Exception as e:
            last_err = e
            time.sleep(0.5)
    if ctx is None:
        _kill(proc, profile)
        raise RuntimeError("连接 LibreOffice 失败: %s" % last_err)

    try:
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx)

        url = uno.systemPathToFileUrl(path)
        hidden = PropertyValue(); hidden.Name = "Hidden"; hidden.Value = True
        doc = desktop.loadComponentFromURL(url, "_blank", 0, (hidden,))

        # 1) 刷新所有目录/索引（TOC 在这里被算出页码）
        try:
            idxs = doc.getDocumentIndexes()
            for i in range(idxs.getCount()):
                idxs.getByIndex(i).update()
        except Exception:
            pass
        # 2) 刷新文本域（页码等）
        try:
            doc.getTextFields().refresh()
        except Exception:
            pass
        # 3) 全局重算
        try:
            doc.refresh()
        except Exception:
            pass

        # 存回为 docx（Word 2007 XML 过滤器）
        flt = PropertyValue(); flt.Name = "FilterName"; flt.Value = "MS Word 2007 XML"
        ow = PropertyValue(); ow.Name = "Overwrite"; ow.Value = True
        doc.storeToURL(url, (flt, ow))
        doc.close(False)
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        _kill(proc, profile)


def _kill(proc, profile):
    try:
        proc.terminate()
        try:
            proc.wait(timeout=10)
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
        print("OK")
        sys.exit(0)
    except Exception as e:
        sys.stderr.write("目录更新失败: %s\n" % e)
        sys.exit(1)
