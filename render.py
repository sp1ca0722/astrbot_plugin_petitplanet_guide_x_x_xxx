"""把数据填进 CuisineTemplate.html，渲染成 PNG。

渲染方式（按优先级）：
  1. 常驻 Chrome + CDP —— 仅当 USE_RESIDENT=True，约 88ms/张，但常驻约 490MB
  2. 一次性 `chrome --screenshot` —— 默认方式，约 1.2~1.8s/张，用完即退
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
import signal
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
    # Linux / 容器（Debian 的 chromium、Ubuntu 的 snap、以及 Chrome 官方包）
    "/usr/bin/chromium",
    "/usr/bin/chromium-browser",
    "/usr/bin/google-chrome",
    "/usr/bin/google-chrome-stable",
    "/snap/bin/chromium",
    "/opt/google/chrome/chrome",
    # Windows
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]
# 即使不在上述路径，也试着从 PATH 里找（容器里 apt 装的通常已在 PATH）
BROWSER_NAMES = ["chromium", "chromium-browser", "google-chrome", "google-chrome-stable"]

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
    """按 CHROME_PATH → 常见安装路径 → PATH 的顺序找一个可用的浏览器。"""
    for p in BROWSERS:
        if p and os.path.exists(p):
            return p
    for name in BROWSER_NAMES:
        p = shutil.which(name)
        if p:
            return p
    raise RuntimeError(
        "没找到 Chrome / Chromium。可用环境变量 CHROME_PATH 指定；"
        "Debian/Ubuntu 容器里可 `apt-get install -y --no-install-recommends chromium`。"
    )


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
    """一次性渲染：起一个 headless Chrome 截完就退。

    慢（~1.2~1.8s，几乎全是进程启动），但峰值过后不占任何常驻内存，
    适合 4GB 这类内存紧张的场景（USE_RESIDENT=False 时走这里）。
    """
    exe = find_browser()
    # 临时 profile 放工作目录（部署时通常在数据盘），不污染插件目录
    os.makedirs(WORK_ROOT, exist_ok=True)
    profile = tempfile.mkdtemp(dir=WORK_ROOT, prefix="profile-")
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
    # 降峰值内存：单渲染进程 + GPU 内联；再补上环境相关参数（root 需 --no-sandbox）
    cmd[1:1] = _LEAN_FLAGS + _env_flags()
    try:
        proc = subprocess.run(cmd, check=False, capture_output=True, timeout=120)
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    if not os.path.exists(png_path):
        # 把 Chrome 的 stderr 带出来，否则只能看到「没生成截图」很难排
        err = (proc.stderr or b"").decode("utf-8", "replace").strip()
        tail = "\n".join(err.splitlines()[-6:]) if err else "（无输出）"
        raise RuntimeError(f"渲染失败，没有生成截图（rc={proc.returncode}）\n{tail}")


# ---------------------------------------------------------------- 临时 Chrome（CDP 单次）
#
# 关键结论（4GB ARM 服务器实测，容器内 root）：
#   一次性 `--screenshot`            ≈ 4.0s / 张（空页面也要 3.0s，全是进程启动）
#   常驻 Chrome + CDP（只换 body）    ≈ 0.33~0.73s / 张，冷启动 1.17s
# 也就是说即使内存紧张，用「临时起一个 CDP 会话、渲完立刻退出」也比
# `--screenshot` 快一倍以上，因为省掉了 Chrome 主进程的重复初始化。
#
# 所以策略是三级：
#   1. USE_RESIDENT=True  → 常驻（最快，约 88ms/张，但常驻 ~490MB）
#   2. USE_RESIDENT=False → 每次起临时 CDP 会话，用完即退（约 0.4~2s，不常驻）
#   3. 上面都不可用       → 一次性 `--screenshot`（约 4s）

CDP_IDLE_SECONDS = 0          # 保留字段，已不用于自动关闭（见下方说明）


# ---------------------------------------------------------------- 常驻 Chrome（CDP）
#
# 一次性 `chrome --screenshot` 每次要 1.2~1.8 秒，绝大部分是进程启动 + profile 初始化，
# 真正画一页只要几十毫秒。这里改成「启动一次、常驻、用 CDP 反复出图」：
# 实测 ~1.8s -> ~88ms，且输出与一次性方案逐像素一致（差异 0）。
#
# 代价是常驻内存（实测，渲染 15 张后的稳态）：
#   工作集 ~490 MB / 私有 ~208 MB（已是优化后；默认启动 ~590 MB / ~253 MB）
# 所以 USE_RESIDENT 默认关闭；内存充裕（例如 16GB 桌面机）时改回 True 可让每张快 19 倍。
#
# 不做空闲超时：菜谱卡片本来就是低频功能，空闲几分钟后关掉反而要重付 1.1s 冷启动，
# 常驻的意义就没了。Chrome 会一直活到插件卸载 / 进程退出（shutdown）。
#
# 依赖 websockets（同步客户端，>=12）；没装就自动回落一次性渲染，不影响可用性。

USE_RESIDENT = True           # True  = 常驻 Chrome：首次 ~1.2s 启动，之后 ~0.3~0.7s/张
                              # False = 每次临时起会话（约 3.5s/张），不占常驻内存
                              # 4GB 的 ARM 服务器实测：常驻比临时会话快 5~10 倍，
                              # 内存代价约 490MB 工作集，可用内存充足时建议开着。
_CDP_LOCK = threading.RLock()
_CDP = None                   # 当前 _ChromeCDP 实例

# 降内存：单渲染进程 + GPU 进程内联。实测 590/253 -> 489/208 MB（约 -18%），
# 连渲 15 张与默认参数的输出逐像素一致。
_LEAN_FLAGS = ["--renderer-process-limit=1", "--in-process-gpu"]


def _env_flags():
    """按运行环境决定的启动参数。

    - 容器/Docker 里常以 root 运行，Chrome 在这种情况下**必须**加 --no-sandbox，
      否则直接报错退出："Running as root without --no-sandbox is not supported"。
    - 容器默认 /dev/shm 只有 64MB，渲染大页面容易崩，用 --disable-dev-shm-usage
      改成写临时文件。
    非 root（普通 Linux 用户 / Windows）不加 --no-sandbox，保留沙箱。
    """
    flags = []
    try:
        geteuid = getattr(os, "geteuid", None)   # Windows 上没有这个函数
        if geteuid is not None and geteuid() == 0:
            flags.append("--no-sandbox")
    except Exception:
        pass
    flags.append("--disable-dev-shm-usage")
    return flags

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
        self._pgid = None
        self._scale = None
        self._base_loaded = False
        self._base_mtime = 0.0

    # ---- 生命周期 ----
    def start(self):
        from websockets.sync.client import connect

        self.close()
        # 工作目录放 WORK_ROOT（部署时是数据盘），而不是 /tmp：
        # 容器里的 /tmp 是 overlay 层，写在那儿既占内存盘、又跟主体产物分家。
        os.makedirs(WORK_ROOT, exist_ok=True)
        self._work = tempfile.mkdtemp(prefix="cdp-", dir=WORK_ROOT)
        tmpdir = os.path.join(self._work, "tmp")
        os.makedirs(tmpdir, exist_ok=True)
        port = _free_port()
        argv = [
            self.exe,
            "--headless=new",
            "--disable-gpu",
            "--disable-extensions",
            "--hide-scrollbars",
            "--no-first-run",
            "--disable-crash-reporter",
            "--user-data-dir=" + os.path.join(self._work, "prof"),
            "--remote-debugging-port=%d" % port,
            # 启动时就定下倍数：只用 Emulation 覆盖的话，文字抗锯齿会和
            # 一次性渲染（--force-device-scale-factor）差 ~2.4% 像素，肉眼看不出但没必要。
            "--force-device-scale-factor=%s" % (("%g" % self.scale)),
            "--window-size=%d,%d" % (W, H),
        ] + _env_flags() + _LEAN_FLAGS + ["about:blank"]

        # 把子进程的 TMPDIR 指到自己的会话目录：
        # Chrome 会在 /tmp 里建 org.chromium.Chromium.* 临时目录，而被 SIGKILL 的进程
        # 来不及自己清理，这些目录会一直堆在 /tmp。指进来后随 _work 一起删。
        env = dict(os.environ)
        env["TMPDIR"] = tmpdir
        env["TMP"] = tmpdir
        env["TEMP"] = tmpdir

        # 关键：让 Chrome 自成进程组/会话。
        # 直接 Popen 时 Chrome 会 fork 出 zygote/renderer 等子进程，只 terminate 主进程
        # 会把它们变成孤儿；而容器里 PID 1 是 AstrBot 自己的 python，不回收孤儿，
        # 于是攒出一堆僵尸（实测连渲几次就堆了 50 个 Z 状态进程）。
        #
        # 用 start_new_session=True（内部就是 setsid(2)），子进程成为新会话/进程组的组长，
        # close() 里 killpg(pgid) 能一次收干净。
        # 注意：不要在前加 "setsid" 命令去包一层，那样第一个参数会被当成程序名。
        # 注意：不要把 stdout/stderr 塞进 **kwargs。
        # subprocess.DEVNULL 的值是整数 -3，写成 {"stdout": subprocess.DEVNULL}
        # 会被类型检查推断成 dict[str, int]，展开后每个参数都报「int 不可分配给 bool/str」。
        self._pgid = None
        if os.name == "posix":
            self.proc = subprocess.Popen(
                argv, start_new_session=True, env=env,
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
        else:
            # Windows：CREATE_NEW_PROCESS_GROUP，配合 taskkill /T 杀整棵树
            self.proc = subprocess.Popen(
                argv, creationflags=subprocess.CREATE_NEW_PROCESS_GROUP, env=env,
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
        self._pgid = self.proc.pid

        # 等调试端口就绪（ARM 容器实测约 1.2s）
        deadline = time.time() + 30
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
        pgid, self._pgid = self._pgid, None
        if proc is not None:
            # 先按进程组整体收，再兜底 kill 主进程
            self._kill_group(pgid)
            if proc.poll() is None:
                try:
                    proc.terminate()
                except Exception:
                    pass
                try:
                    proc.wait(timeout=3)
                except Exception:
                    try:
                        proc.kill()
                    except Exception:
                        pass
            # 必须 wait 一次：poll()/wait() 才会真正回收子进程，
            # 否则哪怕进程已死，本进程里也会留下 <defunct>。
            try:
                proc.wait(timeout=2)
            except Exception:
                pass
            # 组内可能还有没退干净的子孙，再补一刀并等它们消失
            self._kill_group(pgid)
            self._wait_group_gone(pgid)

        if self._work:
            shutil.rmtree(self._work, ignore_errors=True)
            self._work = None

    @staticmethod
    def _kill_group(pgid):
        if not pgid:
            return
        try:
            if os.name == "posix":
                os.killpg(pgid, signal.SIGKILL)
            else:
                subprocess.run(["taskkill", "/F", "/T", "/PID", str(pgid)],
                               capture_output=True, timeout=10)
        except Exception:
            pass

    @staticmethod
    def _wait_group_gone(pgid, timeout=2.0):
        """等进程组里只剩本进程自己的（已 wait 过的）僵尸，最多等 timeout 秒。

        Linux 上直接扫 /proc 统计同 PGID 的存活进程；其它平台直接返回。
        """
        if not pgid or os.name != "posix":
            return
        deadline = time.time() + timeout
        while time.time() < deadline:
            alive = 0
            try:
                for pid in os.listdir("/proc"):
                    if not pid.isdigit():
                        continue
                    try:
                        with open("/proc/%s/stat" % pid) as f:
                            line = f.read()
                    except OSError:
                        continue
                    if ")" not in line:
                        continue
                    f3 = line.rsplit(")", 1)[-1].split()
                    if len(f3) < 3:
                        continue
                    state, group = f3[0], f3[2]
                    # 只算活着的（Z 是已死等待回收，不算占用）
                    if group == str(pgid) and state != "Z":
                        alive += 1
            except OSError:
                return
            if alive == 0:
                return
            time.sleep(0.05)

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
            print("常驻 Chrome 出图失败，回落临时会话：%r" % (e,), file=sys.stderr)
            if _CDP is not None:
                _CDP.close()
                _CDP = None
            return False


def _render_cdp_oneshot(html, png_path, scale):
    """临时起一个 CDP 会话，渲完立刻退出。

    比一次性 `--screenshot` 快一倍以上（省掉 Chrome 主进程重复初始化），
    又不像常驻那样长期占内存，是低内存环境下的默认路径。
    失败返回 False，由调用方继续回落到 `--screenshot`。
    """
    try:
        from websockets.sync.client import connect  # noqa: F401  仅探测依赖
    except Exception:
        return False

    session = None
    try:
        session = _ChromeCDP(find_browser(), scale)
        session.start()
        session.shoot(html, png_path, scale)
        return os.path.isfile(png_path)
    except Exception as e:
        print("临时 CDP 会话失败，回落 chrome --screenshot：%r" % (e,), file=sys.stderr)
        return False
    finally:
        if session is not None:
            session.close()


def shutdown():
    """关掉常驻 Chrome（插件卸载 / 进程退出时调用）。"""
    global _CDP
    with _CDP_LOCK:
        if _CDP is not None:
            _CDP.close()
            _CDP = None


# ---------------------------------------------------------------- 工作目录
#
# 部署时往往希望缓存/产物落在数据盘上，而不是容器的 /tmp（overlay 层，占内存盘、重启就没）。
# 优先级：
#   1. 环境变量 PETITPLANET_WORK_DIR（显式指定，最优先）
#   2. AstrBot 的数据目录：插件在 <root>/data/plugins/<name> 时用 <root>/data/plugin_data/petitplanet
#      （AstrBot 的 data 目录通常已经是挂载的数据盘，比如宿主 /mnt/ssd/astrbot/data）
#   3. 系统临时目录

def _pick_work_root() -> str:
    env = os.environ.get("PETITPLANET_WORK_DIR", "").strip()
    if env:
        try:
            os.makedirs(env, exist_ok=True)
            return env
        except OSError:
            pass

    # HERE 是插件目录：<root>/data/plugins/<插件名>
    # 目标    ：<root>/data/plugin_data/petitplanet
    # 必须严格校验这两级目录名，否则「上两级刚好存在」的任意布局都会被误判：
    # 比如插件直接放在 ~/xxx 时，HERE/../.. 就是整个用户目录，会在那里乱建目录。
    if (os.path.basename(os.path.dirname(HERE)) == "plugins"
            and os.path.basename(os.path.dirname(os.path.dirname(HERE))) == "data"):
        cand = os.path.normpath(os.path.join(HERE, "..", "..", "plugin_data", "petitplanet"))
        try:
            os.makedirs(cand, exist_ok=True)
            probe = os.path.join(cand, ".probe")
            with open(probe, "w") as f:
                f.write("1")
            os.remove(probe)
            return cand
        except OSError:
            pass

    return os.path.join(tempfile.gettempdir(), "petitplanet")


WORK_ROOT = _pick_work_root()


def _cleanup_stale_sessions():
    """清掉上次异常退出（被 SIGKILL / 崩溃）留下的 Chrome 会话目录。

    正常的 close() 会自己删；但进程被强杀时删不掉，会在数据盘和 /tmp 里攒目录。
    这里在模块加载时扫一遍「超过 1 小时没动过」的目录删掉。
    """
    import glob

    cutoff = time.time() - 3600
    patterns = [
        os.path.join(WORK_ROOT, "cdp-*"),
        os.path.join(tempfile.gettempdir(), "cdp-*"),
        os.path.join(tempfile.gettempdir(), "org.chromium.Chromium.*"),
    ]
    for pat in patterns:
        for p in glob.glob(pat):
            try:
                if os.path.isdir(p) and os.path.getmtime(p) < cutoff:
                    shutil.rmtree(p, ignore_errors=True)
            except OSError:
                pass


try:
    _cleanup_stale_sessions()
except Exception:
    pass

# ---------------------------------------------------------------- 结果缓存
#
# 同一道菜的内容是确定的，渲染一次后没必要再算。键 = 最终 HTML + 倍数的哈希，
# 模板改了或数据变了哈希自然会变，不会命中过期图。

_CACHE_DIR = os.path.join(WORK_ROOT, "card_cache")
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
    渲染分三级，前一级不可用/失败就退到下一级：
      1. 常驻 Chrome（USE_RESIDENT=True）—— ~88ms/张，常驻 ~490MB
      2. 临时 CDP 会话（默认）  —— 约 0.4~2s/张，用完即退
      3. chrome --screenshot     —— 约 4s/张（容器里），最慢但兼容性最好
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

    done = False
    if not USE_PLAYWRIGHT:
        # 1) 常驻（打开了才走）
        if USE_RESIDENT:
            done = _render_cdp(html, png_path, scale)
        # 2) 临时 CDP 会话：比 --screenshot 快一倍以上
        if not done:
            done = _render_cdp_oneshot(html, png_path, scale)

    # 3) 最后回落：写中间 html 后走 --screenshot / playwright
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
