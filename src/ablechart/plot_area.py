"""Plot-area geometry heuristics for chart overlay alignment.

PowerPoint chart XML rarely exposes exact rendered plot bounds unless a manual
layout is set, so we use a structure-aware estimate based on chart title,
legend placement, axis label positions, and chart family.
"""

from __future__ import annotations

from typing import Any

from pptx.chart.chart import Chart
from pptx.enum.chart import XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION


def estimate_chart_plot_area(
    chart: Chart,
    *,
    left: float,
    top: float,
    width: float,
    height: float,
    family_hint: str = "combo",
) -> dict[str, float]:
    left_pad = 0.42
    right_pad = 0.16
    top_pad = 0.14
    bottom_pad = 0.20

    if family_hint in {"combo", "event_timeline", "range_snapshot_vertical"}:
        left_pad = 0.58
        right_pad = 0.22
        top_pad = 0.22
        bottom_pad = 0.34
    elif family_hint in {"scatter", "style_box"}:
        left_pad = 0.60
        right_pad = 0.28
        top_pad = 0.24
        bottom_pad = 0.44
    elif family_hint == "range_snapshot_horizontal":
        left_pad = 1.10
        right_pad = 0.24
        top_pad = 0.22
        bottom_pad = 0.28

    if _safe_attr(chart, "has_title", False):
        top_pad += 0.24

    if _safe_attr(chart, "has_legend", False):
        try:
            legend_pos = chart.legend.position
        except Exception:
            legend_pos = None
        if legend_pos == XL_LEGEND_POSITION.TOP:
            top_pad += 0.30
        elif legend_pos == XL_LEGEND_POSITION.BOTTOM:
            bottom_pad += 0.34
        elif legend_pos == XL_LEGEND_POSITION.LEFT:
            left_pad += min(max(width * 0.16, 0.9), 1.6)
        elif legend_pos == XL_LEGEND_POSITION.RIGHT:
            right_pad += min(max(width * 0.16, 0.95), 1.75)
        elif legend_pos == XL_LEGEND_POSITION.CORNER:
            top_pad += 0.18
            right_pad += 0.72

    category_axis = _safe_axis(chart, "category_axis")
    value_axis = _safe_axis(chart, "value_axis")

    if category_axis is not None:
        try:
            tick_pos = category_axis.tick_label_position
        except Exception:
            tick_pos = None
        if tick_pos == XL_TICK_LABEL_POSITION.LOW:
            bottom_pad += 0.16
        elif tick_pos == XL_TICK_LABEL_POSITION.HIGH:
            top_pad += 0.12
        elif tick_pos == XL_TICK_LABEL_POSITION.NEXT_TO_AXIS:
            bottom_pad += 0.08

    if value_axis is not None:
        try:
            tick_pos = value_axis.tick_label_position
        except Exception:
            tick_pos = None
        if tick_pos in {XL_TICK_LABEL_POSITION.LOW, XL_TICK_LABEL_POSITION.NEXT_TO_AXIS}:
            left_pad += 0.10
        elif tick_pos == XL_TICK_LABEL_POSITION.HIGH:
            right_pad += 0.10

    if family_hint in {"style_box", "scatter"}:
        # Scatter charts tend to keep a slightly more centered inner plot.
        top_pad += 0.02
        right_pad += 0.04

    if family_hint == "event_timeline":
        # Leave a little more bottom room for category labels plus event labels.
        bottom_pad += 0.06

    return {
        "left": left + left_pad,
        "top": top + top_pad,
        "width": max(width - left_pad - right_pad, 1.0),
        "height": max(height - top_pad - bottom_pad, 1.0),
    }


def _safe_axis(chart: Chart, axis_name: str):
    try:
        return getattr(chart, axis_name)
    except Exception:
        return None


def _safe_attr(obj: Any, name: str, default: Any) -> Any:
    try:
        return getattr(obj, name)
    except Exception:
        return default
