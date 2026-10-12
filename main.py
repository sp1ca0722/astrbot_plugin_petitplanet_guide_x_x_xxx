import asyncio
import os
import re
import unicodedata
from difflib import SequenceMatcher

import pandas as pd

from astrbot.api.event import filter, AstrMessageEvent, MessageChain
from astrbot.api.message_components import Plain
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
EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cuisine.xlsx")
FLOWER_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "flower.xlsx")
# 花配方指令前缀，剥掉 “/花配方”“/flower” 这类头，兼容带不带斜杠/空格
FLOWER_CMD_PREFIX = re.compile(
    r"^\s*/?\s*(?:花配方|flowerrecipe|flower)(?![A-Za-z0-9_])[\s:：]*", re.IGNORECASE
)
# 花组合指令前缀，剥掉 “/花组合”“/flowercombo” 这类头
COMBO_CMD_PREFIX = re.compile(
    r"^\s*/?\s*(?:花组合|flowercombo)(?![A-Za-z0-9_])[\s:：]*", re.IGNORECASE
)

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


@register("astrbot_plugin_petitplanet_guide_x_x_xxx", "spica", "星布谷地漫游指北", "1.1.2")
class MyPlugin(Star):
    def __init__(self, context: Context, config: AstrBotConfig | None = None):
        super().__init__(context)
        # 老版本 AstrBot 不传 config，此时按默认模式（默认 onebot = 去 Markdown）处理
        self.config = config or {}
        self._rows = []
        self._names = []
        self._base_names = []
        self._source = ""  # 表头 L1、M1：数据来源与更新日期
        # 花配方表（flower.xlsx）：列 = 类别|花色|品质|配方1|配方2|获取方式|版本
        self._flower_rows = []
        self._flower_names = []  # 第 1 列：花的完整名（配色+品类）
        self._flower_base = []  # 剥掉标注后的主体名，用于精确匹配
        self._flower_cats = []  # 去重后的品类名，用于把输入的品类模糊匹配到表里的品类

    async def initialize(self):
        """插件实例化后自动调用：把菜谱表读进内存，之后每次查询不再碰磁盘。"""
        try:
            df = await asyncio.to_thread(pd.read_excel, EXCEL_PATH, header=None)
        except Exception as e:  # 文件缺失 / 没装 openpyxl
            logger.error(f"读取 cuisine.xlsx 失败（{EXCEL_PATH}）：{e}")
            return

        # 用列表下标定位，避免再拿标签回查 DataFrame（下标是 int，类型干净）
        self._rows = df.iloc[1:].values.tolist()
        self._names = [str(row[1]).strip() for row in self._rows]  # 菜品名在第 1 列
        self._base_names = [base_name(name) for name in self._names]  # 剥掉标注后的主体名
        # 表头首行的 L、M 两格是「数据来源」「更新日期」，供 /菜谱 源 查询
        header = df.iloc[0].tolist()
        self._source = "\n".join(
            str(cell).strip() for cell in header[11:13] if not pd.isna(cell) and str(cell).strip()
        )
        logger.info(f"菜谱加载完成，共 {len(self._rows)} 条；平台模式：{self._platform_mode()}")

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
        # 「源」是固定关键词：直接回表头的 L1、M1（数据来源、更新日期），不走模糊匹配
        if target == "源":
            return self._source or "表格里没有填写数据来源与更新日期。"

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
            return await self._ask_ai(target)

        best_score, best_pos = matched[0]
        # 有特殊效果才打印；注意判断和打印要用同一个下标，否则会打印出空行
        row = ["" if pd.isna(cell) else cell for cell in self._rows[best_pos]]

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
        if best_score < 1 and len(matched) > 1:
            lines.append("> **或是其他的菜？**")
            lines.append("> " + " ".join(self._names[pos] for _, pos in matched[1:]))
        return "\n".join(lines)

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
            return f"没有「{category_text}」这个品类哦"
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
            lines.append(f"{color1} {sym} {color2} → **({quality})**{result}")
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
            hits.append(f"{self._flower_process(color1, color2, category)}\n→ **({quality})**{self._flower_display_name(str(row[1]).strip(), category)}")

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
        """可选择实现异步的插件销毁方法，当插件被卸载/停用时会调用。"""
