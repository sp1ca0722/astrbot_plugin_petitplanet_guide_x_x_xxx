import asyncio
import hashlib
import json
import os
import random
import re
import unicodedata
from collections import OrderedDict
from datetime import datetime, timedelta, timezone
from difflib import SequenceMatcher

import pandas as pd

from astrbot.api.event import filter, AstrMessageEvent, MessageChain
from astrbot.api.message_components import Image, Plain
from astrbot.api.star import Context, Star, register
from astrbot.api import AstrBotConfig, logger

CUTOFF = 0.3
# 花名模糊匹配的阈值：花名很短（如「白百合」），阈值要比菜谱高，
# 否则「紫霞」这类前缀会以低分匹配到不着边的花
FLOWER_CUTOFF = 0.6
# 删掉品质
NAME_TAIL = re.compile(r"\s*\([^()]*\)\s*$")
# 从消息里剥掉 "/recipe"、"/菜谱" 这类指令前缀，兼容带不带斜杠、带不带空格
# 指令名要与 @filter.command("recipe", alias={'菜谱', '食谱'}) 保持一致（cuisine 是旧命令名，留着兼容）
# 用后瞻 (?![A-Za-z0-9_]) 代替 \b：\b 对中文不生效（中文算 word 字符），会导致 “/菜谱和煦花果茶” 剥不掉
CMD_PREFIX = re.compile(
    r"^\s*/?\s*(?:recipe|cuisine|食谱|菜谱)(?![A-Za-z0-9_])[\s:：]*", re.IGNORECASE
)
# 数据表跟 main.py 放一起，避免受 AstrBot 启动目录影响
PLUGIN_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(PLUGIN_DIR, "cuisine.xlsx")
FLOWER_PATH = os.path.join(PLUGIN_DIR, "flower.xlsx")

# ---- 菜谱卡片的静态资源（CuisineTemplate.html + render.py）----
CARD_ASSET_DIR = "assets/CuisineHtml"  # 相对插件目录，写进 HTML 的路径
CARD_TEMPLATE = os.path.join(PLUGIN_DIR, "CuisineTemplate.html")
# cuisine.xlsx 里的第二张表：食材值 → 图标名
ICON_SHEET = "图标"
ICON_FILE_PREFIX = "cook-foodtype-"
# 品质 → (底盘, 食材底框)。表里没有品质标注的基础菜按「白」处理
QUALITY_ASSETS = {
    "白": ("bg-quality-white.png", "cook-bg-icon-white.png"),
    "蓝": ("bg-quality-blue.png", "cook-bg-icon-white.png"),  # 无蓝底框，用白
    "紫": ("bg-quality-purple.png", "cook-bg-icon-purple.png"),
    "金": ("bg-quality-gold.png", "cook-bg-icon-gold.png"),
}
# 厨具 → 图标（「图标」表只盖食材，厨具在这里兜底）
COOKER_ASSETS = {
    "煮锅": "cook-cookertype-zhuguo.png",
    "烤箱": "cook-cookertype-kaoxiang.png",
    "榨汁机": "cook-cookertype-zhazhiji.png",
}
CARD_NOTE_TEXT = "该信息来自于三测"
# 卡片消息里最多列几个「其它的菜」（防止一次刷出几十个名字）。
# 指令按钮每行最多 5 个、最多 5 行，这里取 10 正好两行。
CARD_HINT_LIMIT = 10

# ---- QQ 官方：一条消息里「内嵌图 + 指令按钮」----
# 图片会被发成富媒体（msg_type=7）并剥掉 markdown 字段，导致 ** 显示成原样。
# 想同时要图片和按钮，只能走 markdown（msg_type=2）内嵌图：
#   1) POST /v2/media-images/prepare  拿预签名 URL
#   2) PUT 预签名 URL 直传二进制
#   3) POST /v2/media-images/complete 拿 qqbot-media:// 引用
#   4) 发送 markdown（可带 keyboard），发完 DELETE 释放配额（上限 100 张）
# 卡片图用 JPEG 而非 PNG：卡片背景整幅不透明、不需要 alpha，
# 实测同一张卡 JPEG(q95) 体积只有 PNG 的 1/4.8，编码还更快，
# 连带让素材上传和消息下发（平台要转存/分发图片）都变快。
CARD_MEDIA_FILE_NAME = "card.jpg"
CARD_MEDIA_CONTENT_TYPE = "image/jpeg"
CARD_OUTPUT_EXT = ".jpg"  # 渲染产物扩展名，要与上面的 Content-Type 一致
CARD_MEDIA_MAX_BYTES = 10 * 1024 * 1024  # 官方单张上限
CARD_BUTTON_PER_ROW = 5  # 官方：每行最多 5 个
CARD_BUTTON_MAX_ROWS = 5  # 官方：最多 5 行
CARD_BUTTON_LABEL_MAX = 12  # 按钮文字太长会被截断，这里自己先截
# 素材引用缓存条数上限。官方配额是 100 张，而菜有 104 道，留点余量取 90。
# 缓存落盘（media_cache.json）：qqbot-media 引用有效期是 90 天（滑动），
# 存到磁盘后插件重启也不用重新上传，直接省掉 prepare/PUT/complete 约 1.45s。
CARD_MEDIA_CACHE_MAX = 90
MEDIA_CACHE_FILE = "media_cache.json"
# 卡片输出倍数：0.8 → 画布 911x736 出图 729x589，群里刷屏不至于太大
CARD_RENDER_SCALE = 0.8

# render.py 的模块缓存（按文件路径加载，见 load_render_module）
_RENDER_MODULE = None
# 图片测试用的样图，同样按插件目录定位（/imagetest）
TEST_IMAGE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "assets",
    "cuisine",
    "2000050_GameplayEntity_Cook_Meat_b.png",
)
# 花配方指令前缀，剥掉 “/花配方”“/flower” 这类头，兼容带不带斜杠/空格
FLOWER_CMD_PREFIX = re.compile(
    r"^\s*/?\s*(?:花配方|flowerrecipe|flower)(?![A-Za-z0-9_])[\s:：]*", re.IGNORECASE
)
# 花组合指令前缀，剥掉 “/花组合”“/flowercombo” 这类头
COMBO_CMD_PREFIX = re.compile(
    r"^\s*/?\s*(?:花组合|flowercombo)(?![A-Za-z0-9_])[\s:：]*", re.IGNORECASE
)
# 菜谱卡片指令前缀，剥掉 “/卡片”“/card”“/菜谱卡片” 这类头
CARD_CMD_PREFIX = re.compile(
    r"^\s*/?\s*(?:card|菜谱卡片|卡片)(?![A-Za-z0-9_])[\s:：]*", re.IGNORECASE
)

# 开服倒计时：固定按北京时间（UTC+8）算，不依赖服务器时区（容器里常是 UTC）
SERVER_OPEN_AT = datetime(2026, 12, 4, 10, 0, tzinfo=timezone(timedelta(hours=8)))
# 提示还剩不到一天时精确到时分秒，否则只报天数
COUNTDOWN_SECONDS_PER_DAY = 24 * 60 * 60
# 剩余不足 3 小时时，额外提示一句可能提前开服
COUNTDOWN_SECONDS_BEFORE_HINT = 3 * 60 * 60
# 倒计时配图（表情包）：放在 assets/countdown 下，每次随机抽一张
COUNTDOWN_IMAGE_DIR = os.path.join(PLUGIN_DIR, "assets", "countdown")
COUNTDOWN_IMAGE_EXTS = (".jpg", ".jpeg", ".png")

# 平台模式（_conf_schema.json 里的 platform_mode）：只有 QQ 官方机器人渲染 Markdown
MODE_QQ_OFFICIAL = "qq_official"
MODE_ONEBOT = "onebot"
DEFAULT_PLATFORM_MODE = MODE_ONEBOT

# 降级纯文本时要去掉的 Markdown 标记，都是「保留内容、只剥标记」
MD_IMAGE = re.compile(r"!\[([^\]]*)\]\([^)\s]*\)")
MD_LINK = re.compile(r"\[([^\]]*)\]\([^)\s]*\)")
MD_FENCE = re.compile(r"^[ \t]*```.*$\n?", re.MULTILINE)
MD_HR = re.compile(r"^[ \t]{0,3}(?:[-*_][ \t]*){3,}$\n?", re.MULTILINE)
MD_HEADING = re.compile(r"^[ \t]{0,3}#{1,6}[ \t]*", re.MULTILINE)
MD_QUOTE = re.compile(r"^[ \t]{0,3}>[ \t]?", re.MULTILINE)
MD_LIST = re.compile(r"^([ \t]*)[-*+][ \t]+", re.MULTILINE)
MD_CODE = re.compile(r"`([^`\n]+)`")
MD_BOLD = re.compile(r"\*\*(.+?)\*\*")
MD_ITALIC = re.compile(r"(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)")


def base_name(text: str) -> str:
    """全角转半角 → 去首尾空白 → 剥掉尾部的括号标注。"""
    s = unicodedata.normalize("NFKC", str(text)).strip()
    return NAME_TAIL.sub("", s).strip() or s


# 表格里存在 '[ 贝类]'、'荤 菜'、' 煮锅' 这类夹空格/首空格的脏数据，比对前一律抹平
WS_ALL = re.compile(r"\s+")
# 尾部品质标注：(紫) ( 金) [*紫*]
QUALITY_PAREN = re.compile(r"\(\s*\*?([^()\[\]]*?)\*?\s*\)$")
QUALITY_BRACKET = re.compile(r"\[\s*\*?([^()\[\]]*?)\*?\s*\]$")


def squash(text) -> str:
    """全角转半角 + 抹掉所有空白，用于把表格里的脏数据归一化。"""
    return WS_ALL.sub("", unicodedata.normalize("NFKC", str(text if text is not None else "")))


def ingredient_key(text) -> str:
    """食材值 → 「图标」表的匹配键：抹空白 + 剥掉尾部品质标注。

    例：[水果类](紫) → [水果类]；茄子[*紫*] → 茄子；[素菜]/[荤菜] 原样保留
    （以 [ 开头的类别值，末尾的方括号是类别本身，不能剥）。
    """
    s = squash(text)
    m = QUALITY_PAREN.search(s)
    if m and m.start() > 0:
        s = s[: m.start()]
    m = QUALITY_BRACKET.search(s)
    if m and m.start() > 0 and not s.startswith("["):
        s = s[: m.start()]
    return s


def ingredient_quality(text) -> str:
    """从食材值里取出品质（白/蓝/紫/金），没有标注就当白。"""
    s = squash(text)
    for pattern in (QUALITY_PAREN, QUALITY_BRACKET):
        m = pattern.search(s)
        if m and m.start() > 0:
            q = m.group(1)
            return q if q in QUALITY_ASSETS else "白"
    return "白"


def ingredient_label(text) -> str:
    """食材值的显示名：剥品质标注 + 去方括号，如 [素菜]/[荤菜](紫) → 素菜/荤菜。"""
    return ingredient_key(text).replace("[", "").replace("]", "")


def dish_quality(name) -> str:
    """从菜名尾部取品质：…(紫) → 紫；没有标注 → 白。"""
    m = QUALITY_PAREN.search(squash(name))
    if m and m.group(1) in QUALITY_ASSETS:
        return m.group(1)
    return "白"


def load_render_module():
    """按文件路径加载同目录的 render.py，并缓存。

    不能写 `from render import render`：AstrBot 把插件当**包**加载，
    本文件的实际模块名是 `<插件目录名>.main`，而插件目录并不在 sys.path 里，
    所以顶层的 `import render` 会抛 ModuleNotFoundError。
    按文件路径加载（importlib）不依赖 sys.path，也不会与其它插件的同名模块冲突。
    """
    global _RENDER_MODULE
    if _RENDER_MODULE is not None:
        return _RENDER_MODULE

    import importlib.util

    path = os.path.join(PLUGIN_DIR, "render.py")
    if not os.path.isfile(path):
        raise FileNotFoundError(f"找不到渲染脚本：{path}")
    spec = importlib.util.spec_from_file_location("petitplanet_card_render", path)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载渲染脚本：{path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)  # __file__ 指向该路径，render 内部的相对目录才正确
    _RENDER_MODULE = module
    return module


def strip_markdown(text: str) -> str:
    """把 Markdown 标记去掉、只留内容，给不渲染 Markdown 的平台用。

    顺序有讲究：图片要先于链接（否则 `![x](y)` 会剩下一个 `!`），
    行内代码要先于加粗/斜体（避免代码里的 `*` 被当成强调标记）。
    """
    for pattern, repl in (
        (MD_IMAGE, r"\1"),
        (MD_LINK, r"\1"),
        (MD_FENCE, ""),
        (MD_HR, ""),
        (MD_HEADING, ""),
        (MD_QUOTE, ""),
        (MD_LIST, r"\1"),  # 去掉列表的 "- " 标记，保留缩进
        (MD_CODE, r"\1"),
        (MD_BOLD, r"\1"),
        (MD_ITALIC, r"\1"),
    ):
        text = pattern.sub(repl, text)
    return text


def build_countdown_text(now: datetime | None = None) -> str:
    """生成开服倒计时文案：剩余 24 小时以上按日期差算天数，不足 24 小时精确到时分秒。

    天数用「开服日期 - 今天日期」，所以 12/3 查就是「还有 1 天」；
    不能按 24 小时整除向上取整，那样 12/3 会算成「2 天」。
    开服后改为「恭喜星布谷地开服：N 天」，开服当天算第 1 天。
    """
    now = now or datetime.now(SERVER_OPEN_AT.tzinfo)
    left = int((SERVER_OPEN_AT - now).total_seconds())
    if left <= 0:
        days = (now.date() - SERVER_OPEN_AT.date()).days + 1
        return f"\n恭喜星布谷地开服：{days}天"
    if left < COUNTDOWN_SECONDS_PER_DAY:
        hours, rem = divmod(left, 3600)
        minutes, seconds = divmod(rem, 60)
        text = f"\n距离星布谷地开服还有\n{hours}小时{minutes}分{seconds}秒"
        # 剩余不足 3 小时时，多提醒一句可能提前开服
        if left < COUNTDOWN_SECONDS_BEFORE_HINT:
            text += "\n可能会提前开服哦~"
        return text
    return f"\n距离星布谷地开服还有\n{(SERVER_OPEN_AT.date() - now.date()).days}天"


@register("astrbot_plugin_petitplanet_guide_x_x_xxx", "spica", "星布谷地漫游指北", "1.2.0 beta2")
class MyPlugin(Star):
    def __init__(self, context: Context, config: AstrBotConfig | None = None):
        super().__init__(context)
        # 老版本 AstrBot 不传 config，此时按默认模式（默认 onebot = 去 Markdown）处理
        self.config = config or {}
        self._rows = []
        self._names = []
        self._base_names = []
        self._source = ""  # 表头 L 列：版本（兼作数据来源说明）
        # 食材图标（cuisine.xlsx 的「图标」表）：食材值 → 图标名
        self._icon_map = {}
        self._icon_fallback = ""  # 「*」行，未匹配时用
        # 花配方表（flower.xlsx）：列 = 类别|花色|品质|配方1|配方2|获取方式|版本
        self._flower_rows = []
        self._flower_names = []  # 第 1 列：花的完整名（配色+品类）
        self._flower_base = []  # 剥掉标注后的主体名，用于精确匹配
        self._flower_cats = []  # 去重后的品类名，用于把输入的品类模糊匹配到表里的品类
        # 图片内容哈希 → {url, image_id}：QQ 素材库引用缓存（LRU，复用可省一次上传）
        self._media_cache: OrderedDict = OrderedDict()

    async def initialize(self):
        """插件实例化后自动调用：把菜谱表读进内存，之后每次查询不再碰磁盘。"""
        try:
            df = await asyncio.to_thread(pd.read_excel, EXCEL_PATH, header=None)
        except Exception as e:  # 文件缺失 / 没装 openpyxl
            logger.error(f"读取 cuisine.xlsx 失败（{EXCEL_PATH}）：{e}")
            return

        # 表格第 1 行是表头，其余行是数据；菜名空白的行跳过（表格常有多余空行）
        header = df.iloc[0].tolist()
        # 用列表下标定位，避免再拿标签回查 DataFrame（下标是 int，类型干净）
        self._rows = [
            row
            for row in df.iloc[1:].values.tolist()
            if not pd.isna(row[1]) and str(row[1]).strip()
        ]
        self._names = [str(row[1]).strip() for row in self._rows]  # 菜品名在第 1 列
        self._base_names = [base_name(name) for name in self._names]  # 剥掉标注后的主体名
        # 表头 L 列是版本（兼作数据来源），供 /菜谱 源 查询
        # 注意：M 列已被「图片」占用，不能再取 header[11:13]
        self._source = "\n".join(
            str(cell).strip() for cell in header[11:12] if not pd.isna(cell) and str(cell).strip()
        )
        logger.info(f"菜谱加载完成，共 {len(self._rows)} 条；平台模式：{self._platform_mode()}")

        # 食材图标映射：读失败只影响卡片，不影响文本查询
        await self._load_icons()

        # 预热常驻 Chrome：冷启动实测 6.6s（ARM 容器），提前在后台拉起来，
        # 别让第一个用 /卡片 的人白等。预热失败无所谓，按需启动仍然可用。
        try:
            render_mod = load_render_module()
            prewarm = getattr(render_mod, "prewarm", None)
            if prewarm is not None and getattr(render_mod, "USE_RESIDENT", False):
                asyncio.create_task(asyncio.to_thread(prewarm, CARD_RENDER_SCALE))
                logger.info("已在后台预热渲染用 Chrome（首次冷启动约 6s，之后直接复用）")
        except Exception as e:
            logger.warning(f"预热渲染用 Chrome 失败（不影响按需启动）：{e}")

        # 把上次落盘的素材引用读回来：命中的菜直接跳过上传（省 prepare/PUT/complete 约 1.45s）
        self._load_media_cache()

        # 花配方表单独加载：读失败不影响菜谱功能
        try:
            fdf = await asyncio.to_thread(pd.read_excel, FLOWER_PATH, header=None)
        except Exception as e:
            logger.error(f"读取 flower.xlsx 失败（{FLOWER_PATH}）：{e}")
            return
        # 跳过表头行（第 0 行是「类别/花色/…」）
        self._flower_rows = fdf.iloc[1:].values.tolist()
        self._flower_names = [str(row[1]).strip() for row in self._flower_rows]
        self._flower_base = [base_name(name) for name in self._flower_names]
        # 去重后的品类（保持表内先后），供「/花配方 大花马齿 蓝」这类输入做模糊匹配
        self._flower_cats = list(dict.fromkeys(str(row[0]).strip() for row in self._flower_rows))
        logger.info(f"花配方加载完成，共 {len(self._flower_rows)} 条")

    async def _load_icons(self):
        """读「图标」sheet，建立 食材值 → 图标名 的映射（供菜谱卡片用）。

        表结构：匹配键 | 图标 | 类型 | 出现次数 | 说明
        - 匹配键是食材值（`[蔬菜类]`、`番茄`）或兜底 `*`
        - 图标列只填图标名（`veg`），实际文件是 assets/CuisineHtml/cook-foodtype-<名>.png
        读失败或表不存在时只打日志，菜谱文本查询不受影响。
        """
        try:
            gdf = await asyncio.to_thread(
                pd.read_excel, EXCEL_PATH, sheet_name=ICON_SHEET, header=None
            )
        except Exception as e:
            logger.warning(f"读取「{ICON_SHEET}」表失败，菜谱卡片将不显示食材图标：{e}")
            return

        mapping, fallback, missing = {}, "", set()
        # 第 1 行是表头（匹配键|图标|类型|出现次数|说明），其余是映射
        for row in gdf.iloc[1:].values.tolist():
            key = ingredient_key(row[0]) if not pd.isna(row[0]) else ""
            if not key:
                continue
            stem = str(row[1]).strip() if len(row) > 1 and not pd.isna(row[1]) else ""
            if stem and not os.path.isfile(os.path.join(PLUGIN_DIR, CARD_ASSET_DIR, ICON_FILE_PREFIX + stem + ".png")):
                missing.add(stem)
            if key == "*":
                fallback = stem
            else:
                mapping[key] = stem

        self._icon_map = mapping
        self._icon_fallback = fallback
        logger.info(f"食材图标加载完成，共 {len(mapping)} 条映射，兜底：{fallback or '无'}")
        if missing:
            logger.warning(f"「{ICON_SHEET}」表引用了不存在的图标文件：{sorted(missing)}")

    def ingredient_icon(self, text) -> str:
        """食材值 → 图标的相对路径（assets/CuisineHtml/cook-foodtype-*.png）。

        匹配顺序：完整食材值 → 兜底 `*`；都没有返回空字符串（模板会隐藏空 src 的图）。
        """
        stem = self._icon_map.get(ingredient_key(text))
        if stem is None:
            stem = self._icon_fallback
        return f"{CARD_ASSET_DIR}/{ICON_FILE_PREFIX}{stem}.png" if stem else ""

    def build_card_data(self, row: list, note: str = CARD_NOTE_TEXT) -> dict:
        """把一行菜谱数据转成 CuisineTemplate.html 需要的占位符字典。

        row 的列序（0 起）：0 分类 1 菜名 2-5 食材1-4 6 烹饪方式 …
        10 备注 11 版本 12 图片路径
        """
        name = str(row[1]).strip()
        plate = QUALITY_ASSETS[dish_quality(name)][0]
        # 菜品大图：M 列的相对路径（assets/cuisine/…）；没填就用模板自带样图兜底
        dish = str(row[12]).strip() if len(row) > 12 and not pd.isna(row[12]) else ""
        if not dish:
            dish = f"{CARD_ASSET_DIR}/cook-dish.png"

        data = {
            "bg": f"{CARD_ASSET_DIR}/bg.png",
            "plate": f"{CARD_ASSET_DIR}/{plate}",
            "dish": dish,
            "card": f"{CARD_ASSET_DIR}/cook-card.png",
            "title": base_name(name),
            "note_icon": f"{CARD_ASSET_DIR}/icon-warning.png",
            "note": note,
        }

        # 先收集非空食材（跳过中间的空格，否则会出现「第 2 行空、第 3 行有值」的断档）
        ingredients = []
        for i in range(4):
            cell = row[2 + i] if len(row) > 2 + i else None
            text = "" if cell is None or pd.isna(cell) else str(cell).strip()
            if text:
                ingredients.append(text)
        cooker = "" if pd.isna(row[6]) else str(row[6]).strip()

        # 行压缩：食材依次占前几行，厨具紧跟最后一个食材。
        # 例：只有 2 个食材时，厨具出现在第 3 行，而不是空两行再显示。
        layout = [("ing", t) for t in ingredients]
        if cooker:
            layout.append(("cooker", cooker))

        for i in range(5):
            key = i + 1          # row1..row5
            if i >= len(layout):
                # 剩余行全部留空；第 1~4 行的底框也清空，模板会连同图标一起隐藏
                if key <= 4:
                    data[f"row{key}_pill"] = ""
                data[f"row{key}_icon"] = ""
                data[f"row{key}_name"] = ""
                continue

            kind, text = layout[i]
            if kind == "ing":
                # 底框跟着食材自身的品质走（表里用 (紫)/(金) 标注）
                data[f"row{key}_pill"] = (
                    f"{CARD_ASSET_DIR}/{QUALITY_ASSETS[ingredient_quality(text)][1]}"
                )
                data[f"row{key}_icon"] = self.ingredient_icon(text)
                data[f"row{key}_name"] = ingredient_label(text)
            else:
                # 厨具没有品质底框：落在第 1~4 行时把 pill 清空，模板会隐藏它
                if key <= 4:
                    data[f"row{key}_pill"] = ""
                data[f"row{key}_icon"] = (
                    f"{CARD_ASSET_DIR}/{COOKER_ASSETS.get(text, 'cook-cookertype-zhuguo.png')}"
                )
                data[f"row{key}_name"] = text
        return data

    @filter.command("flower", alias={'花配方'})
    async def flower(self, event: AstrMessageEvent):
        """查询花配方：/花配方 花名 或品类（品类列出整类配方），也可 /花配方 品类 品质（如：/花配方 百合 蓝）"""
        target = FLOWER_CMD_PREFIX.sub("", event.message_str).strip()
        text = self._render(await self._build_flower_reply(target))
        yield event.plain_result(text)

    @filter.command("flowercombo", alias={'花组合'})
    async def flowercombo(self, event: AstrMessageEvent):
        """查询杂交组合的后代：/花组合 品类 色1 色2（如：/花组合 百合 白 黄）"""
        target = COMBO_CMD_PREFIX.sub("", event.message_str).strip()
        text = self._render(self._build_combo_reply(target))
        yield event.plain_result(text)

    @filter.command("countdown", alias={'倒计时'})
    async def countdown(self, event: AstrMessageEvent):
        """开服倒计时：/倒计时（文字 + 一张随机表情包）"""
        # 不走 markdown 富文本：带图片的消息本来就会被平台剥掉 markdown 字段，
        # 直接用「纯文字 + 富媒体图片」，图片排在文字下方。
        parts: list = [Plain(build_countdown_text())]
        image_path = self._pick_countdown_image()
        if image_path:
            parts.append(Image.fromFileSystem(image_path))
        try:
            await event.send(MessageChain(parts))
        except Exception as e:
            logger.error(f"发送倒计时失败：{e}")
        event.stop_event()
        yield

    def _pick_countdown_image(self) -> str | None:
        """从 assets/countdown 里随机抽一张配图；目录缺失或为空时返回 None。"""
        try:
            names = [
                name
                for name in os.listdir(COUNTDOWN_IMAGE_DIR)
                if name.lower().endswith(COUNTDOWN_IMAGE_EXTS)
            ]
        except OSError as e:
            logger.warning(f"读取倒计时配图目录失败（{COUNTDOWN_IMAGE_DIR}）：{e}")
            return None
        if not names:
            logger.warning(f"倒计时配图目录里没有图片：{COUNTDOWN_IMAGE_DIR}")
            return None
        return os.path.join(COUNTDOWN_IMAGE_DIR, random.choice(names))

    @filter.command("imagetest", alias={'图片测试'})
    async def imagetest(self, event: AstrMessageEvent):
        """测试发送本地图片：/imagetest"""
        if not os.path.isfile(TEST_IMAGE_PATH):
            logger.error(f"测试图片不存在：{TEST_IMAGE_PATH}")
            yield event.plain_result(f"测试图片不存在：{TEST_IMAGE_PATH}")
            return
        # fromFileSystem 会把路径转成 file:///...，OneBot 适配器（NapCat）可直接读取本机文件
        yield event.chain_result([Image.fromFileSystem(TEST_IMAGE_PATH)])

    @filter.command("linktest", alias={'链接测试'})
    async def linktest(self, event: AstrMessageEvent):
        """实验：两条「可点击交互」路线各发一条，真机点一下就知道哪条能用。

        第 1 条：markdown 链接（对照组）。官方文档说 markdown 链接点击 = 打开网页，
                 如果你实测那个 bot 的蓝字真是 markdown 链接且能填入输入框，这条会复现。
        第 2 条：keyboard 指令按钮（action.type=2）。官方文档原话：「自动在输入框插入 @bot data」，
                 和你截图里输入框出现「@小呱呱 /人气精灵」的形态完全一致 —— 主候选。
                 AstrBot 消息链没有 Keyboard 组件，这里绕过它直接调 botpy 的 post_group_message。
        """
        # —— 路线 1：markdown 链接 ——
        md_text = "\n".join(
            [
                "【实验 1/2】markdown 链接（对照组）",
                "点下面蓝字，预期只会打开网页，不会填入输入框：",
                "[🔗 点我试试（https 链接）](https://www.qq.com)",
            ]
        )
        chain = MessageChain([Plain(md_text)])
        # 显式强制 markdown（msg_type=2），不依赖平台默认。
        # 用 setattr：服务器上的 astrbot 有这个字段，但本地 Pylance 解析的是旧版类型桩，会误报。
        setattr(chain, "use_markdown_", True)
        try:
            await event.send(chain)
        except Exception as e:
            logger.error(f"发送 markdown 链接测试失败：{e}")

        # —— 路线 2：keyboard 指令按钮 ——
        try:
            # getattr 间接调用：服务器上的 AstrBot 有这个方法，本地 Pylance 解析的是旧版类型桩，会误报
            get_inst = getattr(self.context, "get_platform_inst", None)
            if get_inst is None:
                raise RuntimeError("当前 AstrBot 版本不支持 get_platform_inst")
            adapter = get_inst(event.get_platform_id())
            client = adapter.get_client()
            source = event.message_obj
            keyboard = {
                "rows": [
                    {
                        "buttons": [
                            {
                                "id": "probe1",
                                "render_data": {
                                    "label": "🌸 花配方 百合",
                                    "visited_label": "已填入",
                                    "style": 1,
                                },
                                "action": {
                                    "type": 2,  # 指令按钮：自动在输入框插入 @bot + data
                                    "permission": {"type": 2},  # 所有人可点
                                    "data": "/花配方 百合",
                                    "unsupport_tips": "当前QQ版本不支持，请升级",
                                },
                            },
                            {
                                "id": "probe2",
                                "render_data": {
                                    "label": "🍳 卡片 杏鲍菇",
                                    "visited_label": "已填入",
                                    "style": 1,
                                },
                                "action": {
                                    "type": 2,
                                    "permission": {"type": 2},
                                    "data": "/卡片 杏鲍菇",
                                    "unsupport_tips": "当前QQ版本不支持，请升级",
                                },
                            },
                        ]
                    }
                ]
            }
            kb_md = {"content": "【实验 2/2】keyboard 指令按钮\n点下面的蓝色按钮，预期把「/花配方 百合」填入输入框：\n"}
            # botpy 不对 keyboard 做序列化包装，直接透传；官方 API 要求 keyboard = {"content": {rows}}
            keyboard_payload = {"content": keyboard}
            # event.message_obj 是 AstrBotMessage（AstrBot 自己的包装），没有 botpy 的
            # group_openid / user_openid 字段 —— 上一版就是因此走了「认不出目标」分支。
            # 适配器自己的做法：群消息时 abm.group_id = message.group_openid，
            # 私聊时 abm.session_id = sender.user_id。所以用 is_private_chat 分流，
            # 群取 get_group_id()（空则退回 session_id），私聊取 session_id。
            msg_id = getattr(source, "message_id", None)
            seq = random.randint(1, 10000)
            if event.is_private_chat():
                await client.api.post_c2c_message(
                    openid=event.session_id,
                    msg_type=2,
                    markdown=kb_md,
                    keyboard=keyboard_payload,
                    msg_id=msg_id,
                    msg_seq=seq,
                )
            else:
                gid = event.get_group_id() or event.session_id
                await client.api.post_group_message(
                    group_openid=gid,
                    msg_type=2,
                    markdown=kb_md,
                    keyboard=keyboard_payload,
                    msg_id=msg_id,
                    msg_seq=seq,
                )
        except Exception as e:
            logger.error(f"发送 keyboard 按钮测试失败：{e}")

        event.stop_event()
        yield

    # alias 同时保留「卡片」和「菜谱卡片」：按钮里填的是 /卡片 菜名，
    # 若只留「菜谱卡片」按钮就会被填成一条不存在的指令（点了没反应）。
    @filter.command("card", alias={'卡片', '菜谱卡片'})
    async def card(self, event: AstrMessageEvent):
        """渲染菜谱卡片图片：/卡片 菜名（渲染失败时退化成文本结果）"""
        target = CARD_CMD_PREFIX.sub("", event.message_str).strip()
        # 先算一次匹配，拿到「可能还想要的其它菜」；精确命中时没有候选
        row, best_score, matched = (
            self._match_dish(target) if self._rows and target else (None, 0.0, [])
        )
        candidates = self._candidate_names(best_score, matched, limit=CARD_HINT_LIMIT)

        png = await self._render_card(target)
        if not png:
            # 渲染不可用（没装 Chrome / 模板缺失 / 菜名没匹配上）时退回文本，不让命令彻底失败
            text = self._render(await self._build_reply(target))
            yield event.plain_result(text)
            return

        # 首选：QQ 官方平台上发一条 markdown 消息，图片内嵌 + 候选菜做成「指令按钮」。
        # 这样既保住一条消息，又不会像富媒体那样把 ** 和 > 原样显示出来。
        if row is not None and self._platform_mode() == MODE_QQ_OFFICIAL:
            sent = await self._send_card_markdown(event, png, row, candidates)
            if sent:
                event.stop_event()
                yield
                return
            logger.warning("markdown 卡片没发出去，改用富媒体兜底")

        # 兜底：内嵌图/按钮不可用（非 QQ 官方、上传被拒、版本太旧……）时，
        # 退回「富媒体图片 + 纯文字提示」，功能不至于丢。
        # 注意：带图片时适配器会强制走 msg_type=7 并剥掉 markdown 字段，
        # 所以这里只能用纯文字，不能再塞 ** 和 >。
        parts: list = [Image.fromFileSystem(png)]
        plain_hint = self._plain_candidate_hint(candidates)
        if plain_hint:
            parts.append(Plain(plain_hint))
        try:
            await event.send(MessageChain(parts))
        except Exception as e:
            logger.error(f"发送菜谱卡片失败：{e}")
        event.stop_event()
        yield

    async def _render_card(self, target: str) -> str | None:
        """把匹配到的菜渲染成 PNG，返回图片路径；任何环节出问题都返回 None。"""
        if not self._rows or not target:
            return None
        row, _, _ = self._match_dish(target)
        if row is None:
            logger.info("渲染：没匹配到菜名，跳过渲染")
            return None

        try:
            render_mod = load_render_module()
            render_png = render_mod.render
        except Exception as e:
            logger.error(f"加载 render.py 失败，无法渲染菜谱卡片：{e}")
            return None

        data = self.build_card_data(row)
        # 输出到 render 的工作根目录（部署时通常已在数据盘上，如 AstrBot 的 data 目录），
        # 而不是插件目录（避免越写越大）也不是 /tmp（容器里会在 overlay 层）
        out_dir = os.path.join(render_mod.WORK_ROOT, "cards")
        os.makedirs(out_dir, exist_ok=True)
        safe = re.sub(r"[^\w\u4e00-\u9fff-]+", "_", base_name(str(row[1])))[:40] or "card"
        png_path = os.path.join(out_dir, f"{safe}{CARD_OUTPUT_EXT}")

        logger.info(f"开始渲染菜谱卡片：{row[1]} -> {png_path}")
        # 这一步是大头：拉 Chrome、切 DOM、截图；use_cache=True 命中磁盘缓存时会快很多
        try:
            # render 内部会拉 Chrome 子进程，扔到线程里避免阻塞事件循环
            # use_cache=True：同一道菜重复查询直接复用上次的 PNG
            await asyncio.to_thread(render_png, data, png_path, CARD_RENDER_SCALE, True)
        except Exception as e:
            logger.error(f"渲染菜谱卡片失败：{e}")
            return None

        if not os.path.isfile(png_path):
            logger.error(f"渲染结束但没生成文件：{png_path}")
            return None
        return png_path

    # ---- QQ 官方：一条消息里「内嵌图 + 候选菜指令按钮」----

    def _get_qq_client(self, event: AstrMessageEvent):
        """拿到 QQ 官方适配器的 botpy client；拿不到就返回 None。

        AstrBot 的消息链里没有 Keyboard 组件、适配器发送时也不读 keyboard，
        所以要发按钮必须绕过 message chain，直接调 botpy 的接口。
        """
        get_inst = getattr(self.context, "get_platform_inst", None)
        if get_inst is None:
            return None
        adapter = get_inst(event.get_platform_id())
        if adapter is None:
            return None
        get_client = getattr(adapter, "get_client", None)
        return get_client() if get_client else None

    async def _send_card_markdown(
        self,
        event: AstrMessageEvent,
        png_path: str,
        row: list,
        candidates: list,
    ) -> bool:
        """用一条 markdown 消息发送卡片：图片内嵌 + 候选菜做成指令按钮。

        成功返回 True；任何环节不行都返回 False，由调用方退回富媒体方案。
        """
        client = self._get_qq_client(event)
        if client is None:
            logger.error("发送：拿不到 QQ client，改用兜底方案")
            return False

        # 素材引用会缓存复用（省掉 prepare/PUT/complete 约 1.5s）。
        # 但缓存可能已过期（官方 90 天滑动过期），所以发送失败时丢缓存、重新上传再试一次。
        for attempt in (1, 2):
            media = await self._upload_card_image(client, png_path, allow_cache=attempt == 1)
            if not media:
                return False
            try:
                content = self._card_markdown(row, png_path, media["url"], candidates)
                keyboard = self._card_keyboard(candidates)

                await self._post_markdown(event, client, content, keyboard)
                return True
            except Exception as e:
                logger.error(f"发送菜谱 markdown 卡片失败（第 {attempt} 次）：{e}")
                if attempt == 2:
                    return False
                # 大概率是缓存里的 qqbot-media 引用失效了，丢掉它换新的重试
                await self._forget_media(client, media.get("digest"))
        return False

    @staticmethod
    def _image_size(path: str) -> tuple[int, int] | None:
        """读图片宽高（PNG 走 IHDR，JPEG 走 SOF 段），读不出返回 None。"""
        try:
            with open(path, "rb") as f:
                data = f.read()
        except OSError:
            return None

        # PNG：IHDR 固定在第 16/20 字节
        if data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) >= 24:
            return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")

        # JPEG：扫到 SOF 段才有尺寸（SOF0..SOF15，但 C4/C8/CC 不是 SOF）
        if data[:2] != b"\xff\xd8":
            return None
        i, n = 2, len(data)
        while i + 9 < n:
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                height = int.from_bytes(data[i + 5 : i + 7], "big")
                width = int.from_bytes(data[i + 7 : i + 9], "big")
                return width, height
            seg_len = int.from_bytes(data[i + 2 : i + 4], "big")
            if seg_len < 2:
                return None
            i += 2 + seg_len
        return None

    def _card_markdown(
        self, row: list, png_path: str, media_url: str, candidates: list
    ) -> str:
        """拼 markdown 正文：标题 + 内嵌图片 + 候选菜提示（按钮挂在消息底部）。"""
        title = f"**[{row[0]}] {row[1]}**"
        size = self._image_size(png_path)
        alt = str(row[1]).replace("[", "(").replace("]", ")")
        # 宽高可选，但显式写出来可以让客户端按原尺寸渲染、避免二次缩放
        if size:
            image = f"![{alt} #{size[0]}px #{size[1]}px]({media_url})"
        else:
            image = f"![{alt}]({media_url})"
        lines = [title, "", image]
        if candidates:
            lines += ["", "**或是其他的菜？**"]
        return "\n".join(lines)

    def _card_keyboard(self, candidates: list) -> dict | None:
        """把候选菜名做成「指令按钮」（action.type=2：点击后自动在输入框插入 @bot + data）。"""
        if not candidates:
            return None
        buttons = []
        for index, name in enumerate(candidates):
            label = name
            if len(label) > CARD_BUTTON_LABEL_MAX:
                label = label[: CARD_BUTTON_LABEL_MAX - 1] + "…"
            buttons.append(
                {
                    "id": f"dish{index}",
                    "render_data": {
                        "label": label,
                        "visited_label": label,
                        "style": 1,
                    },
                    "action": {
                        "type": 2,  # 指令按钮：自动在输入框插入 @bot + data
                        "permission": {"type": 2},  # 所有人可点
                        "data": f"/卡片 {name}",
                        "unsupport_tips": "当前QQ版本不支持按钮，请升级",
                    },
                }
            )
        rows = [
            {"buttons": buttons[i : i + CARD_BUTTON_PER_ROW]}
            for i in range(0, len(buttons), CARD_BUTTON_PER_ROW)
        ][:CARD_BUTTON_MAX_ROWS]
        return {"content": {"rows": rows}} if rows else None

    async def _post_markdown(
        self,
        event: AstrMessageEvent,
        client,
        content: str,
        keyboard: dict | None,
    ) -> None:
        """直接调 botpy 发一条 markdown 消息（可选挂按钮）。"""
        kwargs: dict = {
            "msg_type": 2,  # markdown
            "markdown": {"content": content},
            # 被动回复要带 msg_id；同一 msg_id 重复用同一个 msg_seq 会失败，所以随机取
            "msg_id": getattr(event.message_obj, "message_id", None),
            "msg_seq": random.randint(1, 10000),
        }
        if keyboard:
            kwargs["keyboard"] = keyboard
        if event.is_private_chat():
            await client.api.post_c2c_message(openid=event.session_id, **kwargs)
        else:
            # 群聊时适配器把 group_openid 存在 group_id 上，取不到再退回 session_id
            gid = event.get_group_id() or event.session_id
            await client.api.post_group_message(group_openid=gid, **kwargs)

    async def _upload_card_image(
        self, client, png_path: str, allow_cache: bool = True
    ) -> dict | None:
        """取得卡片图在素材库的引用，返回 {url, image_id, digest}。

        url 形如 qqbot-media://image/v1/...，可直接写进 markdown 内嵌图。
        已上传过的相同内容会直接复用缓存（省掉 prepare/PUT/complete 约 1.5s）；
        没命中则走官方三段式：prepare → PUT 直传 → complete。
        """
        try:
            data = await asyncio.to_thread(self._read_bytes, png_path)
        except OSError as e:
            logger.error(f"读取卡片图片失败：{e}")
            return None
        if not data:
            return None
        if len(data) > CARD_MEDIA_MAX_BYTES:
            logger.error(f"卡片图片超过素材上限（{len(data)} 字节），改用富媒体发送")
            return None

        # 同一张图（内容哈希相同）的引用可以反复用，不必每次重传
        digest = hashlib.sha256(data).hexdigest()
        if allow_cache:
            cached = self._media_cache.get(digest)
            if cached:
                # 命中了就挪到末尾，表示“最近用过”，淘汰时优先删早的
                self._media_cache.move_to_end(digest)
                logger.info("素材命中缓存，跳过上传")
                return {**cached, "digest": digest}
            logger.info("素材没命中缓存，需要上传")

        # 延迟导入：botpy / aiohttp 在 AstrBot 里一定有，但本地静态检查环境未必有
        import importlib

        try:
            route_cls = importlib.import_module("botpy.http").Route
            aiohttp = importlib.import_module("aiohttp")
        except Exception as e:
            logger.error(f"素材上传依赖缺失：{e}")
            return None

        http = getattr(getattr(client, "api", None), "_http", None)
        if http is None:
            logger.error("素材：拿不到 botpy 的 http 对象")
            return None

        try:
            prepared = await http.request(
                route_cls("POST", "/v2/media-images/prepare"),
                json={
                    "file_name": CARD_MEDIA_FILE_NAME,
                    "file_size": len(data),
                    "content_type": CARD_MEDIA_CONTENT_TYPE,
                    # 官方要求同时给 MD5 与 SHA1
                    "md5": hashlib.md5(data).hexdigest(),
                    "sha1": hashlib.sha1(data).hexdigest(),
                },
            )
            upload_id = prepared.get("upload_id")
            presigned = prepared.get("presigned_url")
            if not upload_id or not presigned:
                logger.error(f"申请素材上传失败：{prepared}")
                return None

            # PUT 预签名 URL：不能带 botpy 的 Authorization 头，否则 COS 会因签名不符拒绝
            # 这一步把图片二进制直传到对象存储，上传带宽差的时候会明显耗时
            async with aiohttp.ClientSession() as session:
                async with session.put(
                    presigned,
                    data=data,
                    headers={"Content-Type": CARD_MEDIA_CONTENT_TYPE},
                ) as resp:
                    if resp.status >= 300:
                        logger.error(f"上传素材失败：HTTP {resp.status}")
                        return None
                    # COS 的 ETag 就是文件 MD5（带引号）
                    etag = (resp.headers.get("ETag") or "").strip('"')

            completed = await http.request(
                route_cls("POST", "/v2/media-images/complete"),
                json={"upload_id": upload_id, "md5": etag},
            )
            url = completed.get("url")
            if not url:
                logger.error(f"完成素材上传失败：{completed}")
                return None
            entry = {"url": url, "image_id": completed.get("image_id")}
            # 写缓存；超限时把最久没用的那几条删掉，避免撞 100 张配额
            await self._remember_media(client, digest, entry)
            return {**entry, "digest": digest}
        except Exception as e:
            logger.error(f"上传卡片图片到素材库失败：{e}")
            return None

    def _media_cache_path(self) -> str:
        """素材缓存的落盘路径：优先放 render 的数据目录（通常已在数据盘），退化为插件目录。"""
        try:
            root = load_render_module().WORK_ROOT
        except Exception:
            root = PLUGIN_DIR
        return os.path.join(root, MEDIA_CACHE_FILE)

    def _load_media_cache(self) -> None:
        """把上次落盘的素材引用读回内存；文件不存在或坏了都只记日志，按空缓存继续。"""
        path = self._media_cache_path()
        try:
            with open(path, encoding="utf-8") as f:
                saved = json.load(f)
        except FileNotFoundError:
            return
        except Exception as e:
            logger.warning(f"读取素材缓存失败（忽略，按空缓存继续）：{e}")
            return

        if not isinstance(saved, dict):
            return
        for digest, entry in saved.items():
            if not isinstance(entry, dict):
                continue
            url = entry.get("url")
            if url:
                self._media_cache[digest] = {"url": url, "image_id": entry.get("image_id")}
        if self._media_cache:
            logger.info(
                f"素材引用缓存已加载 {len(self._media_cache)} 条"
                f"（这些菜不用重新上传，直接省约 1.45s）"
            )

    def _save_media_cache(self) -> None:
        """把素材缓存写盘（先写临时文件再替换，避免半个文件）。失败不影响发送。"""
        path = self._media_cache_path()
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            tmp = path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self._media_cache, f, ensure_ascii=False)
            os.replace(tmp, path)
        except Exception as e:
            logger.warning(f"写入素材缓存失败（不影响发送）：{e}")

    async def _remember_media(self, client, digest: str, entry: dict) -> None:
        """把素材引用记进 LRU 缓存；超出上限时删掉最久没用的，释放配额。"""
        if not digest:
            return
        self._media_cache[digest] = entry
        self._media_cache.move_to_end(digest)
        while len(self._media_cache) > CARD_MEDIA_CACHE_MAX:
            _, old = self._media_cache.popitem(last=False)
            logger.info(f"素材缓存超限，删掉最久没用的一条：{old.get('image_id')}")
            await self._delete_card_image(client, old.get("image_id"))
        self._save_media_cache()

    async def _forget_media(self, client, digest) -> None:
        """丢弃一条素材缓存（引用可能已过期），并顺便从平台删除。"""
        if not digest:
            return
        old = self._media_cache.pop(digest, None)
        self._save_media_cache()
        if old:
            await self._delete_card_image(client, old.get("image_id"))

    async def _delete_card_image(self, client, image_id) -> None:
        """删除素材释放配额（上限 100 张）；失败只记日志，不影响发送结果。"""
        if not image_id:
            return
        import importlib

        try:
            route_cls = importlib.import_module("botpy.http").Route
            http = getattr(getattr(client, "api", None), "_http", None)
            if http is None:
                return
            await http.request(
                route_cls("DELETE", f"/v2/media-images/{image_id}/delete")
            )
        except Exception as e:
            logger.warning(f"删除卡片素材失败（不影响消息）：{e}")

    @staticmethod
    def _read_bytes(path: str) -> bytes:
        with open(path, "rb") as f:
            return f.read()

    def _plain_candidate_hint(self, candidates: list) -> str:
        """兜底用的纯文字候选提示（富媒体消息里 markdown 不生效，只能给纯文本）。"""
        if not candidates:
            return ""
        return "或是其他的菜？\n" + " ".join(candidates)

    # ---- /卡片 具体实现结束 ----

    @filter.command("recipe", alias={'菜谱', '食谱'})
    async def cuisine(self, event: AstrMessageEvent):
        """查询菜谱：/recipe 菜名"""
        # 输入内容来自 message_str；剥掉指令前缀（含别名「菜谱」），兼容不剥的情况
        target = CMD_PREFIX.sub("", event.message_str).strip()
        text = self._render(await self._build_reply(target))

        # 这里手动发送、不走 `yield event.plain_result(text)`：
        # AstrBot 的「引用回复」（platform_settings.reply_with_quote）会往消息链头部插入 Reply 段，
        # 该段序列化后带一个空的 chain 字段，部分 NapCat 版本会直接拒绝整条消息
        # （ActionFailed retcode=1400, 'message segment "reply" field "chain" must be a scalar value'），
        # 结果消息完全发不出去。自己发送就不含 Reply 段，保证送达（代价是没有引用效果）。
        # 若日后 AstrBot / NapCat 修好，可换回 yield event.plain_result(text)。
        try:
            await event.send(MessageChain([Plain(text)]))
        except Exception as e:
            logger.error(f"发送菜谱结果失败：{e}")
        # 已经自己发过了，终止事件传播，避免 AstrBot 再把结果拿去（带引用地）发送一遍
        event.stop_event()
        yield

    @filter.command("helpme", alias={'帮助'})
    async def help(self, event: AstrMessageEvent):
        """查询菜谱帮助：/helpme"""
        text = (
            "### 使用帮助\n\n"
            "`/食谱 菜名`(可模糊搜索)\n"
            "- 返回这道菜的食材和做法\n\n"
            "`/卡片 菜名`(可模糊搜索)\n"
            "- 把这道菜渲染成图片卡片\n\n"
            "`/倒计时`\n"
            "- 距离星布谷地开服还有多久（附一张表情包）\n\n"
#            "`/菜详情 菜名`(可模糊搜索)\n"
#            "- 返回这道菜具体信息\n"
            "`/许愿 许愿内容`\n"
            "- 可以进行对bot功能的许愿（）"
        )
        yield event.plain_result(self._render(text))

    def _platform_mode(self) -> str:
        """读配置里的平台模式，兼容大小写和空格；取值不认识时回落到默认值。"""
        mode = str(self.config.get("platform_mode") or DEFAULT_PLATFORM_MODE).strip().lower()
        return mode if mode in (MODE_QQ_OFFICIAL, MODE_ONEBOT) else DEFAULT_PLATFORM_MODE

    def _render(self, text: str) -> str:
        """按平台模式整理输出：OneBot 不会渲染 Markdown，得先把标记去掉。"""
        return strip_markdown(text) if self._platform_mode() == MODE_ONEBOT else text

    async def _build_reply(self, target: str) -> str:
        """把查询结果整理成要回复的纯文本。"""
        if not self._rows:
            return "菜谱数据没加载成功，请检查 cuisine.xlsx 是否放在插件目录下。"
        if not target:
            return "请输入菜名，例如：/食谱 和煦花果茶"
        # 「源」是固定关键词：直接回表头的 L1（版本/数据来源），不走模糊匹配
        if target == "源":
            return self._source or "表格里没有填写数据来源。"

        row, best_score, matched = self._match_dish(target)
        if row is None:
            return await self._ask_ai(target)

        lines = []
        if best_score < 1:
            lines.append("猜你想搜：")
        lines.append(f"**[{row[0]}] {row[1]}**")
        ingredients = [str(cell).strip() for cell in row[2:6]]
        ingredients = [f"`{name}`" for name in ingredients if name]
        if ingredients:
            lines.append("食材：" + " ".join(ingredients))
        lines.append(f"厨具：`{row[6]}`")
        if row[10] != "":
            lines.append(str(row[10]))
        # 只有一项匹配时没有「其他的菜」可列，不输出这一行
        hint = self._candidate_hint(best_score, matched)
        if hint:
            lines.append(hint)
        return "\n".join(lines)

    def _candidate_names(self, best_score: float, matched: list, limit: int | None = None) -> list:
        """取出「可能还想要的其它菜」的菜名列表；没有候选时返回空列表。

        和 _candidate_hint 同一套判定：精确命中（1.0）或只有一个结果时不给候选。
        卡片命令要用菜名去生成按钮，所以单独返回 list 而不是拼好的文字。
        """
        if best_score >= 1 or len(matched) < 2:
            return []
        names = [str(self._names[pos]) for _, pos in matched[1:]]
        return names[:limit] if limit is not None else names

    def _candidate_hint(self, best_score: float, matched: list, limit: int | None = None) -> str:
        """把「可能还想要的其它菜」整理成一段带 Markdown 的文字；没有候选时返回空串。

        只在模糊匹配（best_score < 1）且确实有别的候选时才输出，精确命中不啰嗦。
        返回的是引用块形式，和文本回复里的「或是其他的菜？」完全一致：
            > **或是其他的菜？**
            > 甜瓜花茶 甜饼果茶 馥郁瓜果茶

        limit 限制列出多少个（卡片消息里防止刷屏），None 表示全列。
        """
        if best_score >= 1 or len(matched) < 2:
            return ""
        names = [self._names[pos] for _, pos in matched[1:]]
        if limit is not None and len(names) > limit:
            names = names[:limit]
            tail = " 等"
        else:
            tail = ""
        return "> **或是其他的菜？**\n> " + " ".join(names) + tail

    def _match_dish(self, target: str) -> tuple[list | None, float, list]:
        """模糊匹配菜名，返回 (行数据, 匹配分, [(分, 下标)...])；没匹配上时行是 None。

        文本回复和菜谱卡片都要用同一套匹配，抽出来避免两处逻辑跑偏。
        """
        # 用主体名算相似度：输入“梦幻番茄汤汁面”和表里的“梦幻番茄汤汁面(金)”就是完全一致(1.0)
        target_base = base_name(target)
        scored = [
            (SequenceMatcher(None, target_base, base).ratio(), pos)
            for pos, base in enumerate(self._base_names)
        ]
        # key 只按分数排，稳定性保证同分时仍是表里的先后顺序
        matched = sorted(
            (item for item in scored if item[0] >= CUTOFF),
            key=lambda item: item[0],
            reverse=True,
        )
        if not matched:
            return None, 0.0, []
        best_score, best_pos = matched[0]
        # 有特殊效果才打印；注意判断和打印要用同一个下标，否则会打印出空行
        row = ["" if pd.isna(cell) else cell for cell in self._rows[best_pos]]
        return row, best_score, matched

    async def _build_flower_reply(self, target: str) -> str:
        """把花配方查询结果整理成要回复的纯文本。

        /花配方 的三种查法（输入都支持模糊匹配，直接取最相似项）：
        1. 品类（/花配方 百合）→ 列出整类配方
        2. 具体花名（/花配方 紫霞百合）→ 只列这一朵
        3. 品类 + 品质（/花配方 百合 蓝）→ 只列该品质
        """
        if not self._flower_rows:
            return "花数据没加载成功，请检查 flower.xlsx 是否放在插件目录下。"
        if not target:
            return (
                "请输入花名或品类，例如：/花配方 紫霞百合、/花配方 百合\n"
                "查品类会列出整类的杂交配方；也可以按品类+品质批量查，例如：/花配方 百合 蓝"
            )

        parts = target.split()
        # 两根词 = 品类 + 品质（如 /花配方 百合 蓝）
        if len(parts) == 2:
            return self._flower_quality_reply(parts[0], parts[1])
        # 其余 = 品类或具体花名（如 /花配方 百合、/花配方 紫霞百合）
        return self._flower_single_reply(target)

    def _flower_quality_reply(self, category_text: str, quality: str) -> str:
        """「品类 + 品质」查询：先模糊匹配品类（「大花马齿」→「大花马齿苋」），再列该品质的花。"""
        cat = self._flower_category(category_text)
        if cat is None:
            return f"根本没有这种花"
        hits = [
            row
            for row in self._flower_rows
            if str(row[0]).strip() == cat and str(row[2]).strip() == quality
        ]
        if not hits:
            return f"「{cat}」里没有 {quality} 品质的花（白/蓝/紫/金）"
        return "\n\n".join(self._format_flower_row(row) for row in hits)

    def _flower_single_reply(self, target: str) -> str:
        """单段输入的查询：更像品类就列整类（「百合」），更像具体花名就查单朵（「紫霞百合」）。

        为什么要比分数：花名本身含品类（「紫霞百合」里也有「百合」），
        只看「有没有匹配到品类」会把「紫霞百合」误判成品类查询。
        所以品类、花名各取最相似项，谁分高听谁的：
        「百合」品类分 1.0 > 花名分（「紫霞百合」≈0.67）→ 列整类；
        「紫霞百合」花名分 1.0 > 品类分（品类「百合」≈0.67）→ 查单朵。
        """
        cat_score, cat = self._flower_best_category(target)
        flower_score, best_pos = self._flower_best_flower(target)

        if cat is not None and cat_score >= FLOWER_CUTOFF and cat_score >= flower_score:
            return self._build_flower_list(cat)
        if flower_score >= FLOWER_CUTOFF:
            return self._format_flower_row(self._flower_rows[best_pos])
        return "根本没有这种花，或者没有输入完整"

    def _format_flower_row(self, raw_row: list) -> str:
        """把一行花数据格式化成「(品质)花名：\n父本1 ⊗/× 父本2\n→ 花名」；白品质不标品质。"""
        row = ["" if pd.isna(cell) else cell for cell in raw_row]
        category = str(row[0]).strip()
        name = str(row[1]).strip()
        quality = '('+str(row[2]).strip()+')'
        if quality == '(白)':
            quality = ''
        color1, color2 = str(row[3]).strip(), str(row[4]).strip()
        # 基础花（白品质）没有杂交配方，回复获取方式；「？？？」占位行两者皆空
        if not color1 and not color2:
            way = str(row[5]).strip()
            return f"### {quality} {name}：\n{way or '还没有大佬研究出它的配方orz'}"
        return f"### {quality} {name}：\n{self._flower_process(color1, color2, category)}\n→ {self._flower_display_name(name, category)}"

    def _build_flower_list(self, category: str) -> str:
        """列出一个品类的全部杂交配方，一行一组短名：「亲本1 ⊗/× 亲本2 → 后代」。

        例：百合 → 「黄 × 白 → 灿黄」「青 ⊗ 青 → 紫霞」……
        短名 = 花名去掉品类后缀（「灿黄百合」→「灿黄」），紧凑得能一眼看完整个品类；
        单朵详情仍用完整花名。品类里一条配方都没有（比如全是基础花）时回复一句提示。
        """
        lines = []
        for raw_row in self._flower_rows:
            row = ["" if pd.isna(cell) else cell for cell in raw_row]
            if str(row[0]).strip() != category:
                continue
            color1, color2 = str(row[3]).strip(), str(row[4]).strip()
            if not color1 or not color2:
                continue  # 基础花没有配方；「？？？」占位行两者皆空
            sym = self._flower_symbol(color1, color2)
            quality = str(row[2]).strip()
            result = self._flower_short_name(str(row[1]).strip(), category)
            lines.append(f"{color1} {sym} {color2} → (**{quality}**){result}")
        if not lines:
            return f"「{category}」还没有收集到配方"
        return f"### {category}\n" + "\n".join(lines)

    def _flower_short_name(self, name: str, category: str) -> str:
        """显示短名：能剥掉品类后缀就剥（「紫霞百合」→「紫霞」），剥不掉就保持原名。

        「紫罗兰」品类里基底花名就是品类名本身，剥不了，原样返回。
        """
        if name.endswith(category) and len(name) > len(category):
            return name[: -len(category)]
        return name

    def _flower_process(self, color1: str, color2: str, category: str) -> str:
        """杂交过程那一行「亲本1 ⊗/× 亲本2」，只把花色部分加粗。

        亲本名 = 花色 + 品类（「青」+「百合」），只加粗花色那一段 → **青**百合。
        单朵详情与 /花组合 都走这里，保证样式一致；
        品类总览（_build_flower_list）用紧凑短名，不走这里。
        OneBot 等不渲染 Markdown 的平台，_render 会把 ** 标记去掉。
        """
        sym = self._flower_symbol(color1, color2)
        return f"{self._flower_parent(color1, category)} {sym} {self._flower_parent(color2, category)}"

    def _flower_symbol(self, color1: str, color2: str) -> str:
        """杂交符号：两个亲本花色相同 = 自交，用 ⊗；不同则是杂交，用 ×。"""
        return "⊗" if color1 == color2 else "×"

    def _flower_parent(self, color: str, category: str) -> str:
        """亲本名只给花色加粗：正常「青百合」→「**青**百合」；紫罗兰特例还原成「紫罗兰」不加粗。"""
        name = self._flower_full_name(color, category)
        if name == color + category:
            return f"**{color}**{category}"
        return name

    def _flower_display_name(self, name: str, category: str) -> str:
        """箭头后的结果名也只给花色加粗：「紫霞百合」→「**紫霞**百合」。

        花名就是品类名本身时（如「紫罗兰」）没有花色前缀，不加粗。
        """
        color = name[: -len(category)]
        if color and name.endswith(category):
            return f"**{color}**{category}"
        return name

    def _build_combo_reply(self, target: str) -> str:
        """「/花组合 百合 白 黄」：在百合品类里找 白×黄（不分先后）的组合，回它的后代。

        只给一个花色（如 /花组合 百合 青）视为自交，等价于 青 青。
        """
        if not self._flower_rows:
            return "花数据没加载成功，请检查 flower.xlsx 是否放在插件目录下。"
        parts = target.split()
        if len(parts) < 2:
            return "请输入品类和亲本花色，例如：/花组合 百合 白 黄"
        # 品类也走模糊匹配（「大花马齿」→「大花马齿苋」）
        category = self._flower_category(parts[0])
        if category is None:
            return f"没有「{parts[0]}」这个品类哦"
        colors = parts[1:3]
        # 容错：把「青百合」这类完整名剥掉品类尾巴，和表里的花色前缀列对齐
        colors = [
            c[: -len(category)]
            if len(c) > len(category) and c.endswith(category)
            else c
            for c in colors
        ]
        if len(colors) == 1:
            colors *= 2  # 只给一个色 = 自交
        want = sorted(colors)

        hits = []
        for raw_row in self._flower_rows:
            row = ["" if pd.isna(cell) else cell for cell in raw_row]
            if str(row[0]).strip() != category:
                continue
            color1, color2 = str(row[3]).strip(), str(row[4]).strip()
            if not color1 or not color2:
                continue  # 基础花没有杂交配方，不参与组合反查
            if sorted((color1, color2)) != want:
                continue
            quality = str(row[2]).strip()
            hits.append(f"{self._flower_process(color1, color2, category)}\n→ (**{quality}**){self._flower_display_name(str(row[1]).strip(), category)}")

        if not hits:
            return "根本没有这种组合"
        return "\n\n".join(hits)

    def _flower_best_flower(self, text: str) -> tuple[float, int]:
        """在花名（色+品类，如「紫霞百合」）里找与输入最相似的一朵，返回（相似度, 表内行下标）。"""
        target = base_name(text)
        return max(
            (
                (SequenceMatcher(None, target, base).ratio(), pos)
                for pos, base in enumerate(self._flower_base)
            ),
            key=lambda item: item[0],
        )

    def _flower_best_category(self, text: str) -> tuple[float, str | None]:
        """在表内品类里找与输入最相似的一个，返回（相似度, 品类）。

        相似度是原始分数、不设阈值，交给调用方按场景判断；没有品类时返回 (0.0, None)。
        """
        if not self._flower_cats:
            return 0.0, None
        target = base_name(text)
        return max(
            (
                (SequenceMatcher(None, target, cat).ratio(), cat)
                for cat in self._flower_cats
            ),
            key=lambda item: item[0],
        )

    def _flower_category(self, text: str) -> str | None:
        """把用户输入的品类模糊匹配到表里的品类（「大花马齿」→「大花马齿苋」）。

        相似度低于 FLOWER_CUTOFF 时返回 None，让调用方提示没这个品类。
        """
        score, cat = self._flower_best_category(text)
        return cat if score >= FLOWER_CUTOFF else None

    def _flower_full_name(self, color: str, category: str) -> str:
        """把配方里的「色」按品类拼回完整花名：色 + 品类（「青」+「百合」→「青百合」）。
        唯一的例外是紫罗兰品类：配方「紫」指「紫罗兰」本身，直拼会得到
        不存在的「紫紫罗兰」，所以色恰好是品类名前缀时直接返回品类名。
        """
        if category.startswith(color):
            return category
        return color + category

    async def _ask_ai(self, target: str) -> str:
        """用ai来查一些奇怪的菜"""
        # openai 的 import 要 1.2s 左右，只有真需要问 AI 时才导入；命中菜品时可以完全跳过
        from openai import AsyncOpenAI

        api_key = os.environ.get("DEEPSEEK_API_KEY")
        if not api_key:
            logger.error("缺少环境变量 DEEPSEEK_API_KEY，无法调用 AI 兜底")
            return "没有这道菜哦，也不知道它要什么食材……（请管理员配置 DEEPSEEK_API_KEY）"

        client = AsyncOpenAI(api_key=api_key, base_url="https://api.deepseek.com")

        try:
            response = await client.chat.completions.create(
                model="deepseek-flash",
                messages=[
                    {
                        "role": "system",
                        "content": "你是一个厨艺大师，擅长根据食材和厨具推荐菜谱。",
                    },
                    {
                        "role": "user",
                        "content": f"现在用户询问这道菜{target}，但现有的菜单里没有这个菜，请用以下说法来回答：没有这道菜哦，但基本上就是【食材】【食材】（或者更多，最多4个），然后使用【厨具】。语言请一定要简洁。",
                    },
                ],
                stream=False,
                # thinking 关闭 = 非思考模式；reasoning_effort="high" 会把思考模式重新打开，两者冲突，所以不传
                extra_body={"thinking": {"type": "disabled"}},
            )
        except Exception as e:
            logger.error(f"调用 DeepSeek 失败：{e}")
            return "没有这道菜哦，查询服务暂时有点问题，稍后再试试。"

        return response.choices[0].message.content or "没有这道菜哦。"

    async def terminate(self):
        """插件被卸载/停用时会调用：把常驻的渲染用 Chrome 收掉，别留孤儿进程。"""
        try:
            load_render_module().shutdown()
        except Exception as e:
            logger.warning(f"关闭渲染用 Chrome 失败：{e}")
