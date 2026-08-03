# -*- coding: utf-8 -*-
"""飞书（Lark）网页应用免登 + 访问控制。

设计目标：
- 零新增依赖，仅用标准库 urllib 调用飞书开放接口。
- 默认关闭：不配置时整套逻辑不生效，本地/内网使用完全不受影响。
- 开启后：只有本企业飞书成员登录后才能访问页面与 API。

开启方式：在项目根目录放 feishu_config.json（见 feishu_config.example.json），
或用环境变量 FEISHU_APP_ID / FEISHU_APP_SECRET / FEISHU_ENABLE=1。
"""
import os, json, time, ipaddress, urllib.request, urllib.parse
from flask import request, session, redirect, jsonify

_OPEN = "https://open.feishu.cn"

class Config:
    def __init__(self, d):
        self.enable = bool(d.get("enable"))
        self.app_id = (d.get("app_id") or "").strip()
        self.app_secret = (d.get("app_secret") or "").strip()
        # 反代/穿透后对外的 https 根地址，例如 https://report.example.com
        # 留空则自动根据 X-Forwarded-* 头推断
        self.redirect_base = (d.get("redirect_base") or "").strip().rstrip("/")
        # 会话密钥；留空则随机生成（重启后已登录用户需重新登录）
        self.secret_key = (d.get("secret_key") or "").strip()
        # 局域网直连是否免登：内网用户直连 VM 时跳过飞书登录（默认开）。
        # 直连速度快、绕开公网隧道；隧道进来的请求不受影响，照常走飞书。
        self.lan_bypass = d.get("lan_bypass", True)

    @property
    def ready(self):
        return bool(self.enable and self.app_id and self.app_secret)


def load_config(base_dir):
    d = {}
    fp = os.path.join(base_dir, "feishu_config.json")
    if os.path.exists(fp):
        try:
            with open(fp, "r", encoding="utf-8") as f:
                d = json.load(f) or {}
        except Exception:
            d = {}
    # 环境变量覆盖
    if os.environ.get("FEISHU_APP_ID"):
        d["app_id"] = os.environ["FEISHU_APP_ID"]
    if os.environ.get("FEISHU_APP_SECRET"):
        d["app_secret"] = os.environ["FEISHU_APP_SECRET"]
    if os.environ.get("FEISHU_ENABLE"):
        d["enable"] = os.environ["FEISHU_ENABLE"] not in ("0", "", "false", "False")
    if os.environ.get("FEISHU_REDIRECT_BASE"):
        d["redirect_base"] = os.environ["FEISHU_REDIRECT_BASE"]
    return Config(d)


def _post_json(url, payload, headers=None):
    data = json.dumps(payload).encode("utf-8")
    h = {"Content-Type": "application/json; charset=utf-8"}
    if headers:
        h.update(headers)
    req = urllib.request.Request(url, data=data, headers=h, method="POST")
    with urllib.request.urlopen(req, timeout=10) as r:
        return json.loads(r.read().decode("utf-8"))


def _app_access_token(cfg):
    j = _post_json(_OPEN + "/open-apis/auth/v3/app_access_token/internal",
                   {"app_id": cfg.app_id, "app_secret": cfg.app_secret})
    if j.get("code") != 0:
        raise RuntimeError("获取 app_access_token 失败：%s" % j.get("msg"))
    return j["app_access_token"]


def _user_by_code(cfg, code):
    """用登录预授权码换取用户身份。"""
    token = _app_access_token(cfg)
    j = _post_json(_OPEN + "/open-apis/authen/v1/access_token",
                   {"grant_type": "authorization_code", "code": code},
                   {"Authorization": "Bearer " + token})
    if j.get("code") != 0:
        raise RuntimeError("换取用户身份失败：%s" % j.get("msg"))
    return j.get("data", {}) or {}


def _redirect_base(cfg):
    if cfg.redirect_base:
        return cfg.redirect_base
    # 反代场景优先用转发头，确保是对外的 https 域名
    proto = request.headers.get("X-Forwarded-Proto", request.scheme)
    host = request.headers.get("X-Forwarded-Host", request.host)
    return "%s://%s" % (proto, host)


def _authorize_url(cfg):
    redirect_uri = _redirect_base(cfg) + "/feishu/callback"
    q = urllib.parse.urlencode({
        "app_id": cfg.app_id,
        "redirect_uri": redirect_uri,
        "state": "report",
    })
    return _OPEN + "/open-apis/authen/v1/index?" + q


# 无需登录即可访问的路径前缀
_EXEMPT = ("/feishu/callback", "/feishu/login", "/api/health", "/static/", "/favicon")


def _is_lan_direct():
    """请求是否来自局域网直连（而非经 cloudflared 隧道转发）。
    直连：remote_addr 是私有网段 IP（192.168/10/172.16-31），且无转发头。
    隧道：cloudflared 本机转发，remote_addr=127.0.0.1 且带 X-Forwarded-For。"""
    # 经过任何反代/隧道都会带转发头，此时不认为是内网直连
    if request.headers.get("X-Forwarded-For") or request.headers.get("X-Forwarded-Host"):
        return False
    addr = request.remote_addr or ""
    try:
        ip = ipaddress.ip_address(addr)
    except ValueError:
        return False
    # 私有网段但排除本机回环（回环通常是隧道转发的来源）
    return ip.is_private and not ip.is_loopback


def init_app(app, cfg):
    """把飞书免登接入 Flask 应用。cfg 未就绪时不做任何拦截。"""
    if not cfg.ready:
        app.logger.info("飞书免登未启用（未配置 app_id/app_secret 或 enable=false）")
        return

    app.secret_key = cfg.secret_key or os.urandom(24)

    @app.route("/feishu/login")
    def feishu_login():
        return redirect(_authorize_url(cfg))

    @app.route("/feishu/callback")
    def feishu_callback():
        code = request.args.get("code", "")
        if not code:
            return "缺少授权码，请从飞书工作台重新打开本应用。", 400
        try:
            info = _user_by_code(cfg, code)
        except Exception as ex:
            return "飞书登录失败：%s" % ex, 500
        session["fs_user"] = {
            "name": info.get("name", ""),
            "open_id": info.get("open_id", ""),
            "user_id": info.get("user_id", ""),
            "ts": int(time.time()),
        }
        return redirect("/")

    @app.before_request
    def _guard():
        p = request.path or ""
        if any(p.startswith(x) for x in _EXEMPT):
            return None
        # 局域网直连免登：内网用户直连 VM 时放行，享受局域网直连速度
        if cfg.lan_bypass and _is_lan_direct():
            return None
        if session.get("fs_user"):
            return None
        # 未登录：页面导航跳登录，API 返回 401 供前端处理
        if p.startswith("/api/"):
            return jsonify({"ok": False, "login_required": True,
                            "error": "请在飞书中重新登录。"}), 401
        return redirect("/feishu/login")

    app.logger.info("飞书免登已启用，app_id=%s", cfg.app_id[:8] + "…")
