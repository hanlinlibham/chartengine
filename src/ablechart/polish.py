"""报告级图表收尾（polish pass）。

在图表 XML 构建完成后统一执行的"设计规范"处理，解决默认 Office 外观的几个硬伤：

1. **智能值轴范围**：按数据计算 nice 刻度（1/2/2.5/5 × 10^k），数据远离 0 时自动
   收窄区间（净值 1.0~1.24 不再画成 0~1.4）；柱状系列强制包含 0。
2. **双轴网格对齐**：左右轴使用相同的分格数（默认 5 格），网格线严格对齐。
3. **柱宽**：gapWidth 默认 80%（堆叠 50%），告别细柱子。
4. **排版**：标题 13pt 加粗深灰、左对齐；轴标签 9pt 灰；统一无衬线字体。
5. **降噪**：值轴不画轴线，分类轴线用浅灰；网格线（如启用）用极浅灰细线。
6. **plot 区域钉定**：manualLayout 固定绘图区，使 slide 级 overlay（瀑布图
   数值标签/连接线）可以精确对位。

所有函数只在"用户未显式配置"时填充默认值，不覆盖显式配置。
"""

from __future__ import annotations

import math
from typing import Dict, List, Optional, Tuple

from lxml import etree

from .oxml_ns import NAMESPACES

C = NAMESPACES["c"]
A = NAMESPACES["a"]

# 设计常量（颜色已迁移到 tokens.CHART_TOKENS，按调用时读取以支持运行时覆盖）
from .tokens import get_chart_token  # noqa: E402

TITLE_FONT = "微软雅黑"
TITLE_SIZE_PT = 13
SUBTITLE_SIZE_PT = 9.5
AXIS_FONT = "微软雅黑"
AXIS_SIZE_PT = 9
GAP_WIDTH_CLUSTERED = 80
GAP_WIDTH_STACKED = 50
DEFAULT_TICK_COUNT = 5

# CT_CatAx / CT_ValAx 子元素的规范顺序（用于按 schema 插入）
_AXIS_CHILD_ORDER = [
    "axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines",
    "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr",
    "txPr", "crossAx", "crosses", "crossesAt", "auto", "lblAlgn", "lblOffset",
    "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl", "crossBetween",
    "majorUnit", "minorUnit", "dispUnits",
]


# ============================================================================
# nice 刻度计算
# ============================================================================

def _nice_unit(raw: float) -> float:
    """向上取最近的 nice 步长（1 / 1.5 / 2 / 2.5 / 3 / 5 × 10^k）。

    包含 1.5 和 3：GTM 常用 3% 步长刻度，且 1→2 的跳档过猛会造成
    大量留白（如 max=62 时 15 步长给出 0–75，而 20 步长只能 0–80/100）。
    """
    if raw <= 0:
        return 1.0
    exp = math.floor(math.log10(raw))
    base = raw / (10 ** exp)
    for nice in (1.0, 1.5, 2.0, 2.5, 3.0, 5.0, 10.0):
        if base <= nice + 1e-9:
            return nice * (10 ** exp)
    return 10.0 * (10 ** exp)


def nice_range(
    vmin: float,
    vmax: float,
    ticks: int = DEFAULT_TICK_COUNT,
    *,
    include_zero: str = "auto",
) -> Tuple[float, float, float]:
    """计算 nice 轴范围，保证恰好 `ticks` 个分格（刻度线 = ticks+1 条）。

    include_zero:
        - "always": 范围必须包含 0（柱状/面积系列）
        - "never":  紧贴数据
        - "auto":   数据离 0 较近时包含 0，远离时收窄区间（折线/散点）
    """
    if math.isnan(vmin) or math.isnan(vmax):
        return 0.0, 1.0, 1.0 / ticks
    if vmin > vmax:
        vmin, vmax = vmax, vmin
    if vmin == vmax:
        pad = abs(vmin) * 0.1 or 1.0
        vmin, vmax = vmin - pad, vmax + pad

    if include_zero == "always":
        vmin = min(vmin, 0.0)
        vmax = max(vmax, 0.0)
    elif include_zero == "auto":
        # 数据全为正且波动幅度小于整体量级的 60% → 收窄；否则归零起步
        if vmin > 0 and (vmax - vmin) < 0.6 * vmax:
            pass  # 紧贴数据
        elif vmin >= 0:
            vmin = 0.0
        elif vmax <= 0:
            vmax = 0.0

    span = vmax - vmin
    unit = _nice_unit(span / ticks)
    nmin = math.floor(vmin / unit) * unit
    nmax = nmin + unit * ticks
    while nmax < vmax - 1e-9:
        unit = _nice_unit(unit * 1.01)  # 升一档
        nmin = math.floor(vmin / unit) * unit
        nmax = nmin + unit * ticks
    # 浮点修整
    digits = max(0, -int(math.floor(math.log10(unit))) + 2)
    return round(nmin, digits), round(nmax, digits), round(unit, digits)


_EXCEL_DATE_TOKENS = (
    ("yyyy", "%Y"), ("yy", "%y"), ("mm", "%m"), ("dd", "%d"), ("m", "%m"), ("d", "%d"),
)


def strftime_from_excel(fmt: Optional[str]) -> str:
    """Excel 日期格式码 → strftime（yyyy/mm → %Y/%m）。

    逐字符扫描而非级联 replace：否则 mm→%m 之后 m→%m 会把
    已生成的 %m 再次替换，产出 %%mm 这类字面量。
    """
    if not fmt:
        return "%Y/%m"
    out = []
    i = 0
    while i < len(fmt):
        for token, rep in _EXCEL_DATE_TOKENS:
            if fmt[i:i + len(token)].lower() == token:
                out.append(rep)
                i += len(token)
                break
        else:
            out.append(fmt[i])
            i += 1
    return "".join(out)


# ============================================================================
# XML 原语
# ============================================================================

def _axis_insert(ax, tag: str):
    """按 CT_*Ax 规范顺序查找或创建轴子元素。"""
    el = ax.find(f"c:{tag}", namespaces=NAMESPACES)
    if el is not None:
        return el
    el = etree.Element(f"{{{C}}}{tag}")
    order = _AXIS_CHILD_ORDER.index(tag) if tag in _AXIS_CHILD_ORDER else 99
    for i, child in enumerate(ax):
        name = child.tag.split("}")[-1]
        child_order = _AXIS_CHILD_ORDER.index(name) if name in _AXIS_CHILD_ORDER else 99
        if child_order > order:
            ax.insert(i, el)
            return el
    ax.append(el)
    return el


def find_axes(chart) -> Dict[str, list]:
    """返回 {'cat': [...], 'val': [(ax, pos), ...]}；pos 为 'l'/'r'/'b' 等。"""
    plot_area = chart._chartSpace.find(f".//{{{C}}}plotArea")
    cat_axes, val_axes = [], []
    for ax in plot_area:
        name = ax.tag.split("}")[-1]
        if name in ("catAx", "dateAx"):
            cat_axes.append(ax)
        elif name == "valAx":
            pos_el = ax.find(f"{{{C}}}axPos")
            val_axes.append((ax, pos_el.get("val") if pos_el is not None else None))
    return {"cat": cat_axes, "val": val_axes}


def axis_has_explicit_scale(ax) -> bool:
    scaling = ax.find(f"{{{C}}}scaling")
    if scaling is None:
        return False
    return (
        scaling.find(f"{{{C}}}min") is not None
        or scaling.find(f"{{{C}}}max") is not None
    )


def set_axis_scale(ax, vmin: float, vmax: float, unit: Optional[float] = None) -> None:
    scaling = _axis_insert(ax, "scaling")
    if scaling.find(f"{{{C}}}orientation") is None:
        orient = etree.SubElement(scaling, f"{{{C}}}orientation")
        orient.set("val", "minMax")
    for tag, value in (("max", vmax), ("min", vmin)):
        el = scaling.find(f"{{{C}}}{tag}")
        if el is None:
            el = etree.SubElement(scaling, f"{{{C}}}{tag}")
        el.set("val", repr(float(value)))
    if unit is not None:
        major = _axis_insert(ax, "majorUnit")
        major.set("val", repr(float(unit)))


def set_axis_text(ax, font: str = AXIS_FONT, size_pt: float = AXIS_SIZE_PT, color: str = None) -> None:
    """轴刻度标签字体/字号/颜色（覆盖已有设置以保证一致性）。"""
    color = color or get_chart_token("axis_text")
    old = ax.find(f"{{{C}}}txPr")
    # Preserve tick-label rotation set by config (bodyPr@rot/@vert) — polish
    # rebuilds txPr for font/color consistency but must not drop rotation.
    preserved = {}
    if old is not None:
        old_bodyPr = old.find(f"{{{A}}}bodyPr")
        if old_bodyPr is not None:
            for attr in ("rot", "vert"):
                if old_bodyPr.get(attr) is not None:
                    preserved[attr] = old_bodyPr.get(attr)
        ax.remove(old)
    txPr = _axis_insert(ax, "txPr")
    bodyPr = etree.SubElement(txPr, f"{{{A}}}bodyPr")
    for attr, val in preserved.items():
        bodyPr.set(attr, val)
    etree.SubElement(txPr, f"{{{A}}}lstStyle")
    p = etree.SubElement(txPr, f"{{{A}}}p")
    pPr = etree.SubElement(p, f"{{{A}}}pPr")
    defRPr = etree.SubElement(pPr, f"{{{A}}}defRPr")
    defRPr.set("sz", str(int(size_pt * 100)))
    fill = etree.SubElement(defRPr, f"{{{A}}}solidFill")
    etree.SubElement(fill, f"{{{A}}}srgbClr").set("val", color)
    latin = etree.SubElement(defRPr, f"{{{A}}}latin")
    latin.set("typeface", font)
    ea = etree.SubElement(defRPr, f"{{{A}}}ea")
    ea.set("typeface", font)
    etree.SubElement(p, f"{{{A}}}endParaRPr")


def style_axis_line(ax, color: Optional[str]) -> None:
    """轴线颜色；None = 不画轴线。"""
    old = ax.find(f"{{{C}}}spPr")
    if old is not None:
        ax.remove(old)
    spPr = _axis_insert(ax, "spPr")
    ln = etree.SubElement(spPr, f"{{{A}}}ln")
    if color is None:
        etree.SubElement(ln, f"{{{A}}}noFill")
    else:
        ln.set("w", "9525")  # 0.75pt
        fill = etree.SubElement(ln, f"{{{A}}}solidFill")
        etree.SubElement(fill, f"{{{A}}}srgbClr").set("val", color)


def style_gridlines(ax, *, on: bool, color: str = None) -> None:
    color = color or get_chart_token("gridline")
    old = ax.find(f"{{{C}}}majorGridlines")
    if old is not None:
        ax.remove(old)
    if not on:
        return
    grid = _axis_insert(ax, "majorGridlines")
    spPr = etree.SubElement(grid, f"{{{C}}}spPr")
    ln = etree.SubElement(spPr, f"{{{A}}}ln")
    ln.set("w", "6350")  # 0.5pt
    fill = etree.SubElement(ln, f"{{{A}}}solidFill")
    etree.SubElement(fill, f"{{{A}}}srgbClr").set("val", color)


def hide_axis(ax) -> None:
    delete = _axis_insert(ax, "delete")
    delete.set("val", "1")


def set_gap_width(chart, percent: int, *, stacked_percent: int = GAP_WIDTH_STACKED) -> None:
    """为所有 barChart 设置 gapWidth（schema 位置：series 之后、axId 之前）。"""
    for bar_chart in chart._chartSpace.findall(f".//{{{C}}}barChart"):
        grouping = bar_chart.find(f"{{{C}}}grouping")
        is_stacked = grouping is not None and grouping.get("val") in ("stacked", "percentStacked")
        value = stacked_percent if is_stacked else percent

        existing = bar_chart.find(f"{{{C}}}gapWidth")
        if existing is not None:
            existing.set("val", str(value))
            continue
        gap = etree.Element(f"{{{C}}}gapWidth")
        gap.set("val", str(value))
        # 插到第一个 axId 之前；若存在 overlap 则插到 overlap 之前
        anchor = bar_chart.find(f"{{{C}}}overlap")
        if anchor is None:
            anchor = bar_chart.find(f"{{{C}}}axId")
        if anchor is not None:
            bar_chart.insert(list(bar_chart).index(anchor), gap)
        else:
            bar_chart.append(gap)


def apply_chart_title(chart, title: str, subtitle: str = None) -> None:
    """设置标题（可带 GTM 风格单位副标题行）并应用 house style。"""
    chart.has_title = True
    tf = chart.chart_title.text_frame
    tf.text = title
    if subtitle:
        para = tf.add_paragraph()
        para.text = subtitle
    style_chart_title(chart)


def style_chart_title(
    chart,
    *,
    font: str = TITLE_FONT,
    size_pt: float = TITLE_SIZE_PT,
    bold: bool = True,
    color: str = None,
    align_left: bool = True,
) -> None:
    """标题字体 + 左上对齐（manualLayout）。

    首段按标题样式，后续段落按副标题样式（9.5pt 常规灰 — GTM 单位行）。
    无标题时跳过。
    """
    if not chart.has_title:
        return
    color = color or get_chart_token("title")
    from pptx.util import Pt
    from pptx.dml.color import RGBColor

    from pptx.enum.text import PP_ALIGN

    tf = chart.chart_title.text_frame
    for i, para in enumerate(tf.paragraphs):
        is_title = i == 0
        para.alignment = PP_ALIGN.LEFT  # 副标题与标题同左对齐（默认会在文本框内居中）
        for run in para.runs or [para.add_run()]:
            run.font.name = font
            run.font.size = Pt(size_pt if is_title else SUBTITLE_SIZE_PT)
            run.font.bold = bold if is_title else False
            run.font.color.rgb = RGBColor.from_string(color if is_title else get_chart_token("subtitle"))

    if align_left:
        title_el = chart._chartSpace.find(f".//{{{C}}}title")
        if title_el is not None:
            old = title_el.find(f"{{{C}}}layout")
            if old is not None:
                title_el.remove(old)
            layout = etree.Element(f"{{{C}}}layout")
            manual = etree.SubElement(layout, f"{{{C}}}manualLayout")
            etree.SubElement(manual, f"{{{C}}}xMode").set("val", "edge")
            etree.SubElement(manual, f"{{{C}}}yMode").set("val", "edge")
            etree.SubElement(manual, f"{{{C}}}x").set("val", "0.015")
            etree.SubElement(manual, f"{{{C}}}y").set("val", "0.02")
            title_el.insert(0, layout)
            overlay = title_el.find(f"{{{C}}}overlay")
            if overlay is None:
                overlay = etree.SubElement(title_el, f"{{{C}}}overlay")
            overlay.set("val", "0")


def pin_plot_area(chart, *, x: float, y: float, w: float, h: float) -> None:
    """manualLayout 固定绘图区（inner plot），坐标为图表区域的比例。"""
    plot_area = chart._chartSpace.find(f".//{{{C}}}plotArea")
    old = plot_area.find(f"{{{C}}}layout")
    if old is not None:
        plot_area.remove(old)
    layout = etree.Element(f"{{{C}}}layout")
    manual = etree.SubElement(layout, f"{{{C}}}manualLayout")
    etree.SubElement(manual, f"{{{C}}}layoutTarget").set("val", "inner")
    etree.SubElement(manual, f"{{{C}}}xMode").set("val", "edge")
    etree.SubElement(manual, f"{{{C}}}yMode").set("val", "edge")
    etree.SubElement(manual, f"{{{C}}}x").set("val", repr(x))
    etree.SubElement(manual, f"{{{C}}}y").set("val", repr(y))
    etree.SubElement(manual, f"{{{C}}}w").set("val", repr(w))
    etree.SubElement(manual, f"{{{C}}}h").set("val", repr(h))
    plot_area.insert(0, layout)


# ============================================================================
# combo 收尾
# ============================================================================

_ZERO_BASED_TYPES = {"bar", "column", "area"}


def polish_combo_chart(
    chart,
    df,
    categories_col: str,
    series_config: List[Dict],
    *,
    skip_primary_scale: bool = False,
    skip_secondary_scale: bool = False,
    ticks: int = DEFAULT_TICK_COUNT,
    cat_text: Optional[Tuple[str, float]] = None,
    val_text: Optional[Tuple[str, float]] = None,
    sec_text: Optional[Tuple[str, float]] = None,
) -> None:
    """组合图收尾：智能轴范围 + 双轴对齐 + 柱宽 + 排版降噪。

    cat_text/val_text/sec_text: (字体, 字号) 覆盖，用于尊重调用方在
    layout_config 中显式指定的字体；None 用 house style。
    """
    axes = find_axes(chart)

    # 1. 智能值轴范围：4/5/6 格联合寻优，双轴共用同一格数 → 网格线严格对齐
    primary_series = [s for s in series_config if s.get("axis", "primary") == "primary"]
    secondary_series = [s for s in series_config if s.get("axis") == "secondary"]

    scalable = []  # [(ax, vmin, vmax, zero_mode), ...]
    for ax, pos in axes["val"]:
        if axis_has_explicit_scale(ax):
            continue
        group = primary_series if pos != "r" else secondary_series
        skip = skip_primary_scale if pos != "r" else skip_secondary_scale
        if skip or not group:
            continue
        if any(s.get("grouping") == "percent_stacked" for s in group):
            continue
        bounds = _series_bounds(df, group)
        if bounds is None:
            continue
        zero = "always" if any(s.get("type", "bar") in _ZERO_BASED_TYPES for s in group) else "auto"
        scalable.append((ax, bounds[0], bounds[1], zero))

    if scalable:
        best_n = _choose_tick_count([item[1:] for item in scalable])
        for ax, vmin, vmax, zero in scalable:
            nmin, nmax, unit = nice_range(vmin, vmax, best_n, include_zero=zero)
            set_axis_scale(ax, nmin, nmax, unit)
            _ensure_format_resolution(ax, unit)

    # 2. 柱宽
    set_gap_width(chart, GAP_WIDTH_CLUSTERED)

    # 3. 排版与降噪
    horizontal = any(
        bd.get("val") == "bar"
        for bd in chart._chartSpace.findall(f".//{{{C}}}barDir")
    )
    style_chart_title(chart)

    # 数据跨零 → 零轴线（即分类轴线）加深（GTM 的黑色零基线惯例）
    crosses_zero = False
    for ax, pos in axes["val"]:
        if pos == "r" or horizontal:
            continue
        scaling = ax.find(f"{{{C}}}scaling")
        if scaling is None:
            continue
        vmin_el = scaling.find(f"{{{C}}}min")
        vmax_el = scaling.find(f"{{{C}}}max")
        if vmin_el is not None and vmax_el is not None:
            if float(vmin_el.get("val")) < 0 < float(vmax_el.get("val")):
                crosses_zero = True

    axis_line_color = get_chart_token("axis_line")
    cat_line_color = get_chart_token("zero_axis") if crosses_zero else axis_line_color
    cat_font, cat_size = cat_text or (AXIS_FONT, AXIS_SIZE_PT)
    for ax in axes["cat"]:
        set_axis_text(ax, font=cat_font, size_pt=cat_size)
        style_axis_line(ax, cat_line_color)
    for ax, pos in axes["val"]:
        font, size = (sec_text if pos == "r" else val_text) or (AXIS_FONT, AXIS_SIZE_PT)
        set_axis_text(ax, font=font, size_pt=size)
        # 横向条形图的值轴在底部，保留轴线作为基准；纵向值轴不画线
        style_axis_line(ax, axis_line_color if horizontal else None)


def _ensure_format_resolution(ax, unit: float) -> None:
    """轴数字格式精度自适应步长：步长 0.5% 配 "0%" 会产生 1%,1%,2%,2% 重复标签。"""
    num_fmt = ax.find(f"{{{C}}}numFmt")
    if num_fmt is None:
        return
    code = num_fmt.get("formatCode", "")
    if "." in code or unit <= 0:
        return  # 已带小数位的不动
    if code.endswith("%"):
        unit_scaled = unit * 100
    elif code in ("0", "#,##0"):
        unit_scaled = unit
    else:
        return
    if unit_scaled >= 1 - 1e-9:
        return
    decimals = max(1, math.ceil(-math.log10(unit_scaled)))
    body = code[:-1] if code.endswith("%") else code
    num_fmt.set("formatCode", body + "." + "0" * decimals + ("%" if code.endswith("%") else ""))


def _choose_tick_count(axis_specs: List[Tuple[float, float, str]], options=(4, 5, 6)) -> int:
    """为一组值轴选统一格数：最大化各轴的数据填充率之和。

    固定单一格数时 nice 步长可能跳档过猛（如 max=140、5 格 → 步长 50 → 轴到 250，
    留白 44%；6 格 → 步长 25 → 轴到 150，留白 7%）。多格数寻优避免这种浪费，
    且所有轴共用同一格数保证双轴网格线对齐。
    """
    best_n, best_fill = options[0], -1.0
    for n in options:
        total_fill = 0.0
        for vmin, vmax, zero in axis_specs:
            nmin, nmax, _unit = nice_range(vmin, vmax, n, include_zero=zero)
            span = nmax - nmin
            if span <= 0:
                continue
            data_low = nmin if vmin >= 0 and nmin == 0 else vmin
            total_fill += (vmax - data_low) / span
        if total_fill > best_fill + 1e-9:
            best_fill, best_n = total_fill, n
    return best_n


def _series_bounds(df, series_group: List[Dict]) -> Optional[Tuple[float, float]]:
    """系列组的数据范围；堆叠组按行求和（正负分开累计）。"""
    import pandas as pd

    keys = [s["key"] for s in series_group if s.get("key") in df.columns]
    if not keys:
        return None
    frame = df[keys].apply(pd.to_numeric, errors="coerce")

    stacked_keys = [
        s["key"] for s in series_group
        if s.get("grouping") == "stacked"
        and s.get("type", "bar") in ("bar", "column", "area")
        and s["key"] in frame
    ]
    plain_keys = [k for k in keys if k not in stacked_keys]

    candidates_max, candidates_min = [], []
    if stacked_keys:
        sub = frame[stacked_keys]
        candidates_max.append(sub.clip(lower=0).sum(axis=1).max())
        candidates_min.append(sub.clip(upper=0).sum(axis=1).min())
    if plain_keys:
        sub = frame[plain_keys]
        candidates_max.append(sub.max().max())
        candidates_min.append(sub.min().min())

    vmax = max(candidates_max)
    vmin = min(candidates_min)
    if math.isnan(vmin) or math.isnan(vmax):
        return None
    return float(vmin), float(vmax)


# ============================================================================
# scatter / bubble 收尾
# ============================================================================

def polish_xy_chart(chart, x_values, y_values, *, ticks: int = DEFAULT_TICK_COUNT) -> None:
    """散点/气泡收尾：两轴 nice 缩放（不强制含 0）+ 排版降噪。"""
    axes = find_axes(chart)
    # scatter/bubble 的 X 轴也是 valAx；按 axPos 区分: 'b' = X 轴
    for ax, pos in axes["val"]:
        values = x_values if pos == "b" else y_values
        values = [v for v in values if v is not None and not math.isnan(float(v))]
        if not values or axis_has_explicit_scale(ax):
            continue
        nmin, nmax, unit = nice_range(min(values), max(values), ticks, include_zero="auto")
        set_axis_scale(ax, nmin, nmax, unit)

    style_chart_title(chart)
    for ax, pos in axes["val"]:
        set_axis_text(ax)
        if pos == "b":
            style_axis_line(ax, get_chart_token("axis_line"))
            style_gridlines(ax, on=False)
            # 轴穿过 0 时刻度文字会贴在零轴上被数据点遮挡 → 固定到底部
            tick_pos = _axis_insert(ax, "tickLblPos")
            tick_pos.set("val", "low")
        else:
            style_axis_line(ax, None)
            style_gridlines(ax, on=True)
