# -*- coding: utf-8 -*-
"""跨进程文件锁：多 worker(gunicorn 多进程)下保护共享 JSON 的读改写。

用法：给"读整份→修改→写回整份"的写接口套 @locked(锁名)，
同一把锁同一时刻只允许一个请求进入，其余排队，杜绝并发覆盖丢数据。

Linux 用 fcntl.flock(进程间有效)；无 fcntl 的平台(如 Windows 本地开发)
退化为线程锁，保证功能与单元测试可用。
"""
import os, functools, threading

try:
    import fcntl  # Linux/macOS
    _HAS_FCNTL = True
except ImportError:
    _HAS_FCNTL = False

_LOCK_DIR = None
_thread_locks = {}
_thread_locks_guard = threading.Lock()


def init(lock_dir):
    """指定锁文件存放目录(建议项目内一个专用目录)。"""
    global _LOCK_DIR
    _LOCK_DIR = lock_dir
    os.makedirs(lock_dir, exist_ok=True)


def _thread_lock(name):
    with _thread_locks_guard:
        lk = _thread_locks.get(name)
        if lk is None:
            lk = threading.Lock()
            _thread_locks[name] = lk
        return lk


class _Guard:
    """上下文管理器：进入即加锁，退出即解锁。"""
    def __init__(self, name):
        self.name = name
        self._fh = None
        self._tlk = None

    def __enter__(self):
        # 线程锁：即便同一 worker 内多线程也互斥
        self._tlk = _thread_lock(self.name)
        self._tlk.acquire()
        if _HAS_FCNTL and _LOCK_DIR:
            path = os.path.join(_LOCK_DIR, self.name + ".lock")
            self._fh = open(path, "w")
            fcntl.flock(self._fh.fileno(), fcntl.LOCK_EX)  # 阻塞式排他锁
        return self

    def __exit__(self, *exc):
        try:
            if self._fh is not None:
                fcntl.flock(self._fh.fileno(), fcntl.LOCK_UN)
                self._fh.close()
        finally:
            if self._tlk is not None:
                self._tlk.release()
        return False


def guard(name):
    """返回一个可用于 with 的锁上下文。"""
    return _Guard(name)


def locked(name):
    """装饰器：整个视图函数在名为 name 的锁保护下执行。"""
    def deco(fn):
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            with _Guard(name):
                return fn(*args, **kwargs)
        return wrapper
    return deco
