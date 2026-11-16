"""把数据填进 CuisineTemplate.html，渲染成 PNG。

渲染方式（按优先级）：
  1. 常驻 Chrome + CDP（默认）—— 启动一次反复出图，约 85~125ms/张
  2. 一次性 `chrome --screenshot` —— 常驻不可用时回落，约 1.2~1.8s/张
  3. Playwright —— 把 USE_PLAYWRIGHT 改成 True 才用，需自行安装
渲染结果可以按内容哈希做磁盘缓存，重复渲染同一张图约 15ms。

用法:
    python render.py                        # 用示例数据渲染到 预览/示例@0.8x.png
    python render.py --scale 2              # 2 倍图（1822x1472）
    from render import render; render({...}, "out.png")
"""
import base64
import atexit
import hashlib
import html as html_mod
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.request
import uuid

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "CuisineTemplate.html")
WORK = HERE                                   # 临时 html 必须和 assets 同级，相对路径才能用
A = "assets/CuisineHtml/"                     # 模板静态资源（背景/底盘/图标/字体）都在这里
# 相对路径前面紧跟着引号或括号的地方（src="assets/…" 和 url('assets/…') 两种写法都要盖）
ASSET_REF = re.compile(r"(?<=[\"'(])assets/")
W, H = 911, 736

USE_PLAYWRIGHT = False

BROWSERS = [
    os.environ.get("CHROME_PATH", ""),
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]

# ---------------------------------------------------------------- 数据

EXAMPLE = {
    "bg": A + "bg.png",
    "plate": A + "bg-quality-blue.png",
    "dish": A + "cook-dish.png",
    "card": A + "cook-card.png",
    "title": "梦幻和煦花果茶",
    "row1_pill": A + "cook-bg-icon-gold.png",    "row1_icon": A + "cook-foodtype-fowerfruit.png", "row1_name": "问号",
    "row2_pill": A + "cook-bg-icon-purple.png",  "row2_icon": A + "cook-foodtype-fish.png",       "row2_name": "问号/问号",
    "row3_pill": A + "cook-bg-icon-white.png",   "row3_icon": A + "cook-foodtype-jellyfish.png",  "row3_name": "问号问号",
    "row4_pill": A + "cook-bg-icon-white.png",   "row4_icon": A + "cook-foodtype-mushroom.png",   "row4_name": "问号色问号",
    "row5_icon": A + "cook-cookertype-zhuguo.png", "row5_name": "烤问号",
    "note_icon": A + "icon-warning.png", "note": "该信息来自于三测",
}


def fill(data, template=TEMPLATE):
    """把 {{key}} 换成 data[key]（会做 HTML 转义）。"""
    with open(template, encoding="utf-8") as f:
        out = f.read()
    for k, v in data.items():
        out = out.replace("{{%s}}" % k, html_mod.escape(str(v), quote=True))
    return out


# ---------------------------------------------------------------- 渲染

def find_browser():
    for p in BROWSERS:
        if p and os.path.exists(p):
            return p
    raise RuntimeError("没找到 Chrome / Edge，可用环境变量 CHROME_PATH 指定")


def _render_playwright(html_path, png_path, scale):
    # 可选路径：只有把 USE_PLAYWRIGHT 改成 True 才会走到这里，需要自行 pip install playwright
    from playwright.sync_api import sync_playwright  # type: ignore[import-not-found]

    with sync_playwright() as p:
        b = p.chromium.launch()
        page = b.new_page(viewport={"width": W, "height": H}, device_scale_factor=scale)
        page.goto("file:///" + html_path.replace("\\", "/"))
        page.wait_for_timeout(300)          # 等字体加载
        page.screenshot(path=png_path)
        b.close()


def _render_chrome(html_path, png_path, scale):
    exe = find_browser()
    profile = tempfile.mkdtemp(dir=WORK, prefix="profile-")
    cmd = [
        exe,
        "--headless=new",
        "--disable-gpu",
        "--disable-extensions",
        "--hide-scrollbars",
        "--no-first-run",
        "--disable-crash-reporter",
        "--user-data-dir=" + profile,
        "--force-device-scale-factor=%s" % scale,
        "--window-size=%d,%d" % (W, H),
        "--virtual-time-budget=3000",
        "--screenshot=" + png_path,
        "file:///" + html_path.replace("\\", "/"),
    ]
    try:
        subprocess.run(cmd, check=False, capture_output=True, timeout=120)
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    if not os.path.exists(png_path):
        raise RuntimeError("渲染失败，没有生成截图")


# ---------------------------------------------------------------- 常驻 Chrome（CDP）
#
# 一次性 `chrome --screenshot` 每次要 1.2~1.8 秒，绝大部分是进程启动 + profile 初始化，
# 真正画一页只要几十毫秒。这里改成「启动一次、常驻、用 CDP 反复出图」：
# 实测 ~1.8s -> ~88ms，且输出与一次性方案逐像素一致（差异 0）。
#
# 代价是常驻内存（实测，渲染 15 张后的稳态）：
#   工作集 ~490 MB / 私有 ~208 MB（已是优化后；默认启动 ~590 MB / ~253 MB）
# 内存紧张时可把 USE_RESIDENT 改成 False，或由 main.py 根据配置调 set_resident(False)：
# 那样每张图会多花 ~1.1 秒冷启动，但进程用完就退，不占常驻内存。
#
# 不做空闲超时：菜谱卡片本来就是低频功能，空闲几分钟后关掉反而要重付 1.1s 冷启动，
# 常驻的意义就没了。Chrome 会一直活到插件卸载 / 进程退出（shutdown）。
#
# 依赖 websockets（同步客户端，>=12）；没装就自动回落一次性渲染，不影响可用性。

USE_RESIDENT = True           # False = 每张都用一次性渲染（不占常驻内存）
_CDP_LOCK = threading.RLock()
_CDP = None                   # 当前 _ChromeCDP 实例

# 降内存：单渲染进程 + GPU 进程内联。实测 590/253 -> 489/208 MB（约 -18%），
# 连渲 15 张与默认参数的输出逐像素一致。
_LEAN_FLAGS = ["--renderer-process-limit=1", "--in-process-gpu"]

# 等字体和图片都真正就绪再截图：只等 loadEventFired 会截到没图的半成品
_READY_JS = """
(async () => {
  await document.fonts.ready;
  const imgs = [...document.images];
  await Promise.all(imgs.map(i => i.complete ? Promise.resolve()
      : new Promise(r => { i.onload = r; i.onerror = r; })));
  await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
  return imgs.filter(i => !i.complete || i.naturalWidth === 0).length;
})()
"""

# 只换 body 内容，省掉整页重导航（实测 ~83ms vs ~160ms）
_SWAP_JS = """
(async () => {
  document.body.innerHTML = %s;
  await document.fonts.ready;
  const imgs = [...document.images];
  await Promise.all(imgs.map(i => i.complete ? Promise.resolve()
      : new Promise(r => { i.onload = r; i.onerror = r; })));
  await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
  return imgs.filter(i => !i.complete || i.naturalWidth === 0).length;
})()
"""

_ASSET_BASE = "file:///" + (HERE.replace("\\", "/") + "/")
_BODY_RE = re.compile(r"<body[^>]*>(.*)</body>", re.S | re.I)
_PLACEHOLDER_RE = re.compile(r"\{\{\w+\}\}")


def _template_mtime():
    try:
        return os.path.getmtime(TEMPLATE)
    except OSError:
        return 0.0


def _body_inner(html):
    """取出 <body> 内部片段，并把 assets/ 改成绝对 file:// URI。

    改成绝对是必需的：换 body 后页面仍以「骨架页」为基准解析相对路径，
    而骨架页在临时目录，相对路径会变成 404（字体也会跟着加载失败）。
    """
    m = _BODY_RE.search(html)
    inner = m.group(1) if m else html
    return ASSET_REF.sub(_ASSET_BASE + "assets/", inner)


def _base_document():
    """模板骨架：保留 <style> 和 @font-face，body 清空，供后续反复替换。"""
    with open(TEMPLATE, encoding="utf-8") as f:
        doc = f.read()
    doc = _PLACEHOLDER_RE.sub("", doc)
    return ASSET_REF.sub(_ASSET_BASE + "assets/", doc)


def _free_port():
    s = socket.socket()
    try:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]
    finally:
        s.close()


class _ChromeCDP:
    """一个常驻的 headless Chrome，复用同一个标签页反复出图。"""

    def __init__(self, exe, scale):
        self.exe = exe
        self.scale = scale
        self.proc = None
        self.ws = None
        self._mid = 0
        self._work = None
        self._scale = None
        self._base_loaded = False
        self._base_mtime = 0.0

    # ---- 生命周期 ----
    def start(self):
        from websockets.sync.client import connect

        self.close()
        self._work = tempfile.mkdtemp(prefix="cdp-")
        port = _free_port()
        self.proc = subprocess.Popen(
            [
                self.exe,
                "--headless=new",
                "--disable-gpu",
                "--disable-extensions",
                "--hide-scrollbars",
                "--no-first-run",
                "--disable-crash-reporter",
                "--no-sandbox",
                "--user-data-dir=" + os.path.join(self._work, "prof"),
                "--remote-debugging-port=%d" % port,
                # 启动时就定下倍数：只用 Emulation 覆盖的话，文字抗锯齿会和
                # 一次性渲染（--force-device-scale-factor）差 ~2.4% 像素，肉眼看不出但没必要。
                "--force-device-scale-factor=%s" % (("%g" % self.scale)),
                "--window-size=%d,%d" % (W, H),
            ] + _LEAN_FLAGS + [
                "about:blank",
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

        # 等调试端口就绪（通常 ~0.7s，只在第一次/重启时发生）
        deadline = time.time() + 20
        while time.time() < deadline:
            if self.proc.poll() is not None:
                raise RuntimeError("Chrome 启动后立即退出")
            try:
                with urllib.request.urlopen("http://127.0.0.1:%d/json/version" % port, timeout=1):
                    break
            except Exception:
                time.sleep(0.05)
        else:
            raise RuntimeError("Chrome 调试端口超时未就绪")

        # 开一个标签页，拿到它的 WS 地址
        req = urllib.request.Request(
            "http://127.0.0.1:%d/json/new?about:blank" % port, method="PUT"
        )
        with urllib.request.urlopen(req, timeout=10) as r:
            target = json.load(r)
        # proxy=None：websockets>=15 默认会读系统代理，本地回环必须绕开
        self.ws = connect(target["webSocketDebuggerUrl"], max_size=64 * 1024 * 1024, proxy=None)
        self._mid = 0
        self._send("Page.enable")
        self._scale = None

    def close(self):
        ws, self.ws = self.ws, None
        if ws is not None:
            try:
                ws.close()
            except Exception:
                pass
        proc, self.proc = self.proc, None
        if proc is not None and proc.poll() is None:
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except Exception:
                proc.kill()
        if self._work:
            shutil.rmtree(self._work, ignore_errors=True)
            self._work = None

    def alive(self):
        return self.ws is not None and self.proc is not None and self.proc.poll() is None

    # ---- CDP 收发 ----
    def _send(self, method, params=None, wait_event=None, timeout=30):
        """发一条 CDP 命令并等它的响应；顺带把 wait_event 的事件也等到。"""
        ws = self.ws
        if ws is None:
            raise RuntimeError("CDP 未连接（Chrome 可能已退出）")
        self._mid += 1
        mid = self._mid
        ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
        result = None
        event_seen = wait_event is None
        deadline = time.time() + timeout
        while time.time() < deadline:
            msg = json.loads(ws.recv(timeout=max(0.1, deadline - time.time())))
            if msg.get("id") == mid:
                if "error" in msg:
                    raise RuntimeError("CDP %s 失败：%s" % (method, msg["error"]))
                result = msg.get("result", {})
                if event_seen:
                    return result
            elif wait_event and msg.get("method") == wait_event:
                event_seen = True
                if result is not None:
                    return result
        raise TimeoutError("CDP %s 超时" % method)

    def _evaluate(self, expression, await_promise=False):
        r = self._send(
            "Runtime.evaluate",
            {"expression": expression, "awaitPromise": await_promise, "returnByValue": True},
        )
        return r.get("result", {}).get("value")

    # ---- 出图 ----
    def _set_scale(self, scale):
        if scale != self._scale:
            self._send(
                "Emulation.setDeviceMetricsOverride",
                {"width": W, "height": H, "deviceScaleFactor": scale, "mobile": False},
            )
            self._scale = scale

    def _screenshot(self, png_path):
        # optimizeForSpeed 编不出更小的文件，但快一倍（152ms -> 70ms），图片仍无损
        shot = self._send("Page.captureScreenshot", {"format": "png", "optimizeForSpeed": True})
        with open(png_path, "wb") as f:
            f.write(base64.b64decode(shot["data"]))

    def _load_base(self):
        """把模板骨架导航进这个标签页（只做一次），后续只换 body 内容。"""
        if not self._work:
            raise RuntimeError("Chrome 工作目录未创建")
        if self._base_loaded and self._base_mtime == _template_mtime():
            return
        base = _base_document()
        p = os.path.join(self._work, "base.html")
        with open(p, "w", encoding="utf-8") as f:
            f.write(base)
        self._send("Page.navigate", {"url": "file:///" + p.replace("\\", "/")},
                   wait_event="Page.loadEventFired")
        self._evaluate("document.fonts.ready.then(() => 1)", await_promise=True)
        self._base_loaded = True
        self._base_mtime = _template_mtime()

    def shoot(self, html, png_path, scale):
        """html = 已填好的完整文档。优先只换 body（~83ms），否则整页导航（~160ms）。"""
        self._set_scale(scale)
        self._load_base()

        body = _body_inner(html)
        try:
            # 只替换 body：省掉重新加载页面、重新解析 CSS/字体的开销
            self._evaluate(_SWAP_JS % json.dumps(body), await_promise=True)
        except Exception:
            # 替换失败就老实整页导航
            if not self._work:
                raise
            p = os.path.join(self._work, "_p_%s.html" % uuid.uuid4().hex[:8])
            with open(p, "w", encoding="utf-8") as f:
                f.write(ASSET_REF.sub(_ASSET_BASE, html))
            self._send("Page.navigate", {"url": "file:///" + p.replace("\\", "/")},
                       wait_event="Page.loadEventFired")
            self._evaluate(_READY_JS, await_promise=True)
            try:
                os.remove(p)
            except OSError:
                pass

        self._screenshot(png_path)


def _render_cdp(html, png_path, scale):
    """用常驻 Chrome 出图；不可用/出错时返回 False，由调用方回落一次性渲染。"""
    global _CDP
    try:
        from websockets.sync.client import connect  # noqa: F401  仅探测依赖
    except Exception:
        return False

    with _CDP_LOCK:
        try:
            # 进程挂了或倍数变了才重建；否则一直复用同一个 Chrome
            if _CDP is not None and (not _CDP.alive() or _CDP.scale != scale):
                _CDP.close()
                _CDP = None
            if _CDP is None:
                _CDP = _ChromeCDP(find_browser(), scale)
                _CDP.start()
            _CDP.shoot(html, png_path, scale)
            return True
        except Exception as e:
            print("常驻 Chrome 出图失败，回落一次性渲染：%r" % (e,), file=sys.stderr)
            if _CDP is not None:
                _CDP.close()
                _CDP = None
            return False


def shutdown():
    """关掉常驻 Chrome（插件卸载 / 进程退出时调用）。"""
    global _CDP
    with _CDP_LOCK:
        if _CDP is not None:
            _CDP.close()
            _CDP = None


# ---------------------------------------------------------------- 结果缓存
#
# 同一道菜的内容是确定的，渲染一次后没必要再算。键 = 最终 HTML + 倍数的哈希，
# 模板改了或数据变了哈希自然会变，不会命中过期图。

_CACHE_DIR = os.path.join(tempfile.gettempdir(), "petitplanet_card_cache")
_CACHE_MAX = 200            # 超过就清掉一半，防止无限增长


def _cache_key(html, scale):
    h = hashlib.sha256()
    h.update(html.encode("utf-8"))
    h.update(("|scale=%s" % scale).encode("utf-8"))
    return h.hexdigest()[:32]


def _cache_get(html, scale):
    p = os.path.join(_CACHE_DIR, _cache_key(html, scale) + ".png")
    if os.path.isfile(p) and os.path.getsize(p) > 0:
        return p
    return None


def _cache_put(html, scale, png_path, touch=True):
    try:
        os.makedirs(_CACHE_DIR, exist_ok=True)
        dst = os.path.join(_CACHE_DIR, _cache_key(html, scale) + ".png")
        if os.path.abspath(dst) != os.path.abspath(png_path):
            shutil.copyfile(png_path, dst)
        if touch:
            files = [os.path.join(_CACHE_DIR, f) for f in os.listdir(_CACHE_DIR)]
            if len(files) > _CACHE_MAX:
                files.sort(key=os.path.getmtime)
                for f in files[: len(files) // 2]:
                    try:
                        os.remove(f)
                    except OSError:
                        pass
    except OSError:
        pass  # 缓存失败不影响出图


def clear_cache():
    shutil.rmtree(_CACHE_DIR, ignore_errors=True)


# 解释器退出时把常驻 Chrome 收掉，别留孤儿进程
atexit.register(shutdown)


def render(data=None, png_path=None, scale: float = 1, use_cache: bool = False):
    """把数据渲染成 PNG，返回图片的绝对路径。

    use_cache=True 时按「HTML + 倍数」做磁盘缓存，重复渲染同一张图直接拷贝（~15ms）。
    渲染优先走常驻 Chrome（~85ms），不可用时回落一次性启动（~1.2s）。
    """
    os.makedirs(WORK, exist_ok=True)
    if png_path is None:
        # 倍数可能是小数（如 0.8），文件名里用 0.8x / 2x 这种写法，不能 %d
        tag = "" if scale == 1 else "@%sx" % (("%g" % scale))
        png_path = os.path.join(HERE, "预览", "示例%s.png" % tag)
    png_path = os.path.abspath(png_path)      # Chrome 拿到相对路径会写到别处，必须转绝对路径
    os.makedirs(os.path.dirname(png_path), exist_ok=True)

    html = fill(data or EXAMPLE)

    # 命中缓存就直接拷贝，不碰浏览器
    if use_cache:
        hit = _cache_get(html, scale)
        if hit:
            if os.path.abspath(hit) != png_path:
                shutil.copyfile(hit, png_path)
            print("已命中缓存 ->", png_path)
            return png_path

    # 1) 优先走常驻 Chrome：直接把 HTML 交给它，不需要落盘
    done = False
    if not USE_PLAYWRIGHT:
        done = _render_cdp(html, png_path, scale)

    # 2) 回落：写中间 html 后走一次性渲染（~1.2s）
    if not done:
        # 唯一文件名：并发时共用一个 _job.html 会互相覆盖
        stem = "_job_%s" % uuid.uuid4().hex[:10]
        try:
            html_path = os.path.join(WORK, stem + ".html")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html)
        except OSError:
            # 插件目录不可写（部署时可能只读）→ 临时目录 + 绝对 file:// 路径。
            # 模板里既有 src="assets/…" 也有 @font-face 的 url('assets/…')，都要换。
            html_path = os.path.join(tempfile.gettempdir(), stem + ".html")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(ASSET_REF.sub(_ASSET_BASE + "assets/", html))
        try:
            if USE_PLAYWRIGHT:
                _render_playwright(html_path, png_path, scale)
            else:
                _render_chrome(html_path, png_path, scale)
        finally:
            if os.path.exists(html_path):
                os.remove(html_path)

    if use_cache and os.path.isfile(png_path):
        _cache_put(html, scale, png_path)
    print("已渲染 ->", png_path)
    return png_path


if __name__ == "__main__":
    scale = 0.8
    if "--scale" in sys.argv:
        scale = float(sys.argv[sys.argv.index("--scale") + 1])
    render(scale=scale, use_cache=False)
