"""估值区间图（range chart）— GTM 估值页的标志图型。

形态：每个类目一根灰色浮动条（历史区间 低→高），叠加
绿色横杠（历史均值）与蓝色菱形（当前值）。用于
「当前估值在历史区间中的位置」这类真实金融场景：
PE/PB 区间、利差区间、波动率锥、行业仓位历史分位等。

实现：隐形基底 + 可见区间段的堆叠柱（与 waterfall 同一招），
均值/当前值用「只显示 marker 的折线系列」，全部原生可编辑。
"""

from __future__ import annotations

from typing import Optional

import pandas as pd
from pptx.slide import Slide
from pptx.util import Inches

from .api import create_combo_chart
from .annotations import delete_legend_entry, find_series_element, style_marker_only_series
from .oxml_ns import NAMESPACES

RANGE_BASE_KEY = "_range_base"
RANGE_SPAN_KEY = "_range_span"

# 区间/均值/当前默认色现集中在 tokens.CHART_TOKENS（range_band/range_avg/range_current）
from .tokens import get_chart_token  # 颜色真源


def create_range_chart(
    slide: Slide,
    df: pd.DataFrame,
    categories_col: str,
    low_col: str,
    high_col: str,
    *,
    current_col: Optional[str] = None,
    average_col: Optional[str] = None,
    position: tuple = (Inches(1), Inches(2)),
    size: tuple = (Inches(8), Inches(4.5)),
    title: Optional[str] = None,
    subtitle: Optional[str] = None,
    layout_config=None,
    number_format: Optional[str] = None,
    sort: Optional[str] = None,
    range_color: str = None,
    average_color: str = None,
    current_color: str = None,
    range_name: str = "历史区间",
    average_name: str = "历史均值",
    current_name: str = "当前",
):
    """创建估值区间图。low/high 必填，current/average 可选。

    sort: 'asc'/'desc' 按 current（无 current 按 high）排序类目。
    number_format: 值轴格式，如 '0"x"'（PE 倍数）。
    """
    range_color = range_color or get_chart_token("range_band")
    average_color = average_color or get_chart_token("range_avg")
    current_color = current_color or get_chart_token("range_current")
    for col in [low_col, high_col, current_col, average_col]:
        if col and col not in df.columns:
            raise ValueError(f"range chart 缺少列: {col}")

    if sort:
        sort_col = current_col or high_col
        df = df.sort_values(sort_col, ascending=str(sort).lower() != "desc").reset_index(drop=True)

    low = pd.to_numeric(df[low_col], errors="raise")
    high = pd.to_numeric(df[high_col], errors="raise")

    plot_df = pd.DataFrame({
        categories_col: df[categories_col].values,
        RANGE_BASE_KEY: low.values,
        RANGE_SPAN_KEY: (high - low).values,
    })
    series_config = [
        {"key": RANGE_BASE_KEY, "name": RANGE_BASE_KEY, "type": "bar", "axis": "primary", "grouping": "stacked"},
        {"key": RANGE_SPAN_KEY, "name": range_name, "type": "bar", "axis": "primary", "grouping": "stacked"},
    ]
    if average_col:
        plot_df[average_col] = pd.to_numeric(df[average_col], errors="raise").values
        series_config.append({"key": average_col, "name": average_name, "type": "line", "axis": "primary"})
    if current_col:
        plot_df[current_col] = pd.to_numeric(df[current_col], errors="raise").values
        series_config.append({"key": current_col, "name": current_name, "type": "line", "axis": "primary"})

    if layout_config is None:
        from .layout import ChartLayoutConfig, ValueAxisConfig
        layout_config = ChartLayoutConfig(  # 默认图例置底，区间/均值/当前需要图例解释
            value_axis_config=ValueAxisConfig(number_format=number_format) if number_format else None,
        )

    metadata = {
        "chart_family": "range",
        "spec_version": 1,
        "categories_col": categories_col,
        "low_col": low_col,
        "high_col": high_col,
        "current_col": current_col,
        "average_col": average_col,
    }

    chart = create_combo_chart(
        slide=slide,
        df=plot_df,
        categories_col=categories_col,
        series_config=series_config,
        position=position,
        size=size,
        layout_config=layout_config,
        metadata=metadata,
    )

    # 基底隐形 + 区间段灰色
    _hide_series(chart, RANGE_BASE_KEY)
    _fill_series(chart, range_name, range_color)
    delete_legend_entry(chart, 0)

    # GTM 几何：柱宽 40% 槽宽（gapWidth=150）；均值杠按槽宽动态取尺寸，
    # 略宽于柱体、从两侧伸出（JPM 估值图的标志细节）
    n = max(1, len(plot_df))
    slot_pt = (size[0].inches if hasattr(size[0], "inches") else float(size[0]) / 914400.0) * 72 / n
    dash_size = int(min(72, max(12, slot_pt * 0.55)))

    if average_col:
        ser = find_series_element(chart, average_name)
        if ser is not None:
            style_marker_only_series(ser, symbol="dash", size=dash_size, color=average_color)
    if current_col:
        ser = find_series_element(chart, current_name)
        if ser is not None:
            # 白描边让菱形从深色区间条上跳出来
            style_marker_only_series(ser, symbol="diamond", size=9, color=current_color,
                                     border_color="FFFFFF", border_width_pt=1.0)

    from .polish import set_gap_width
    set_gap_width(chart, 150, stacked_percent=150)

    if title:
        from .polish import apply_chart_title
        apply_chart_title(chart, title, subtitle)

    return chart


def _hide_series(chart, name: str) -> None:
    from lxml import etree

    ser = find_series_element(chart, name)
    if ser is None:
        return
    spPr = ser.find(f"{{{NAMESPACES['c']}}}spPr")
    if spPr is not None:
        ser.remove(spPr)
    spPr = etree.Element(f"{{{NAMESPACES['c']}}}spPr")
    etree.SubElement(spPr, f"{{{NAMESPACES['a']}}}noFill")
    ln = etree.SubElement(spPr, f"{{{NAMESPACES['a']}}}ln")
    etree.SubElement(ln, f"{{{NAMESPACES['a']}}}noFill")
    tx = ser.find(f"{{{NAMESPACES['c']}}}tx")
    ser.insert(list(ser).index(tx) + 1 if tx is not None else 0, spPr)


def _fill_series(chart, name: str, color: str) -> None:
    from lxml import etree

    ser = find_series_element(chart, name)
    if ser is None:
        return
    spPr = ser.find(f"{{{NAMESPACES['c']}}}spPr")
    if spPr is not None:
        ser.remove(spPr)
    spPr = etree.Element(f"{{{NAMESPACES['c']}}}spPr")
    fill = etree.SubElement(spPr, f"{{{NAMESPACES['a']}}}solidFill")
    etree.SubElement(fill, f"{{{NAMESPACES['a']}}}srgbClr").set("val", color.lstrip("#").upper())
    ln = etree.SubElement(spPr, f"{{{NAMESPACES['a']}}}ln")
    etree.SubElement(ln, f"{{{NAMESPACES['a']}}}noFill")
    tx = ser.find(f"{{{NAMESPACES['c']}}}tx")
    ser.insert(list(ser).index(tx) + 1 if tx is not None else 0, spPr)
