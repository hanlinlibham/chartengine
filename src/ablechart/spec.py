"""声明式 Chart Spec 层 — 低门槛、JSON 友好的统一入口。

设计目标（面向能力较弱的 LLM 或非专家调用方）：

1. **一个 dict 进、图表出**：不需要 import 任何配置类、不需要 pptx 枚举。
2. **容错归一化**：类型/轴/图例位置/颜色等接受大量别名（含中文），大小写不敏感。
3. **智能推断**：不给 series 就自动取全部数值列；不给 categories 就取第一个非数值列；
   日期列自动启用日期轴。
4. **友好报错**：一次性收集所有错误，列名拼错会给出"是否想用 xxx"建议。
5. **灵活定制**：支持自定义调色板、逐系列覆盖颜色/线宽/标记，所有底层配置仍可触达。

最小示例（一个 JSON 就能画双轴组合图）::

    from ablechart import render_chart

    render_chart(slide, {
        "title": "营收与增速",
        "data": {"年份": [2021, 2022, 2023], "营收": [100, 120, 150], "增速": [0.2, 0.2, 0.25]},
        "series": ["营收", {"column": "增速", "type": "line", "axis": "right"}],
        "layout": {"y2_axis": {"format": "percent"}},
    })
"""

from __future__ import annotations

import contextlib
import difflib
import io
import math
import re
import warnings
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import pandas as pd
from lxml import etree
from pptx.util import Cm, Inches, Pt

from .oxml_ns import NAMESPACES


class SpecError(ValueError):
    """Spec 校验失败。message 中一次性列出全部错误。"""

    def __init__(self, errors: List[str]):
        self.errors = list(errors)
        lines = [f"Chart spec 校验失败 ({len(self.errors)} 处):"]
        lines += [f"  {i}. {msg}" for i, msg in enumerate(self.errors, 1)]
        super().__init__("\n".join(lines))


# ============================================================================
# 别名表（全部小写比较）
# ============================================================================

_CHART_KIND_ALIASES = {
    "combo": "combo", "mixed": "combo", "组合": "combo", "组合图": "combo",
    "dual": "combo", "dual_axis": "combo", "双轴": "combo", "双轴图": "combo",
    "bar": "combo", "column": "combo", "柱状图": "combo", "柱形图": "combo",
    "line": "combo", "折线图": "combo", "area": "combo", "面积图": "combo",
    "timeseries": "combo", "时间序列": "combo",
    "waterfall": "waterfall", "bridge": "waterfall", "瀑布": "waterfall",
    "瀑布图": "waterfall", "桥图": "waterfall", "归因图": "waterfall",
    "scatter": "scatter", "xy": "scatter", "散点": "scatter", "散点图": "scatter",
    "bubble": "bubble", "气泡": "bubble", "气泡图": "bubble",
    "range": "range", "区间": "range", "区间图": "range",
    "估值区间": "range", "valuation_range": "range",
    "contribution": "combo", "贡献": "combo", "贡献图": "combo", "贡献分解": "combo",
}

# contribution 别名：堆叠分项 + 合计线（GTM 宏观分解图标配）
# 颜色取自 tokens（contribution_line / contribution_parts），集中管理。
_CONTRIBUTION_ALIASES = {"contribution", "贡献", "贡献图", "贡献分解"}

_HORIZONTAL_ALIASES = {"horizontal", "h", "bar_h", "横向", "水平", "条形", "条形图"}

# chart 别名携带的默认系列类型（chart: "line" → 系列默认 line）
_CHART_DEFAULT_SERIES_TYPE = {
    "bar": "bar", "column": "bar", "柱状图": "bar", "柱形图": "bar",
    "line": "line", "折线图": "line",
    "area": "area", "面积图": "area",
}

_SERIES_TYPE_ALIASES = {
    "bar": "bar", "column": "bar", "col": "bar", "柱": "bar",
    "柱状": "bar", "柱状图": "bar", "柱形图": "bar",
    "line": "line", "折线": "line", "折线图": "line", "曲线": "line",
    "area": "area", "面积": "area", "面积图": "area",
    "scatter": "scatter", "散点": "scatter", "散点图": "scatter",
    "bubble": "bubble", "气泡": "bubble", "气泡图": "bubble",
}

_AXIS_ALIASES = {
    "primary": "primary", "left": "primary", "l": "primary", "y": "primary",
    "y1": "primary", "main": "primary", "first": "primary", "1": "primary",
    "主": "primary", "主轴": "primary", "左": "primary", "左轴": "primary",
    "secondary": "secondary", "right": "secondary", "r": "secondary",
    "y2": "secondary", "second": "secondary", "2": "secondary",
    "次": "secondary", "次轴": "secondary", "右": "secondary", "右轴": "secondary",
}

_GROUPING_ALIASES = {
    "clustered": None, "cluster": None, "standard": None, "normal": None,
    "并列": None, "簇状": None, "none": None, "": None,
    "stacked": "stacked", "stack": "stacked", "堆叠": "stacked", "堆积": "stacked",
    "percent_stacked": "percent_stacked", "percentstacked": "percent_stacked",
    "percent": "percent_stacked", "100%": "percent_stacked",
    "百分比": "percent_stacked", "百分比堆叠": "percent_stacked",
}

_LEGEND_OFF_VALUES = {"none", "off", "hide", "hidden", "false", "no", "无", "隐藏"}
_LEGEND_POSITION_NAMES = {
    "bottom": "bottom", "底部": "bottom", "下": "bottom",
    "top": "top", "顶部": "top", "上": "top",
    "left": "left", "左侧": "left",
    "right": "right", "右侧": "right",
    "corner": "corner", "右上角": "corner",
}

_NUMBER_FORMAT_ALIASES = {
    "percent": "0%", "percentage": "0%", "百分比": "0%", "%": "0%",
    "percent1": "0.0%", "percent2": "0.00%",
    "int": "#,##0", "integer": "#,##0", "thousands": "#,##0", "千分位": "#,##0",
    "decimal1": "0.0", "decimal2": "0.00", "两位小数": "0.00",
    "currency": "#,##0.00",
    "times": '0"x"', "倍": '0"x"', "x": '0"x"',  # PE/PB 倍数轴
    "general": "General", "auto": "General",
}

_NAMED_COLORS = {
    "red": "C00000", "blue": "0070C0", "green": "1A5C2A", "orange": "ED7D31",
    "gray": "808080", "grey": "808080", "black": "000000", "white": "FFFFFF",
    "gold": "C9A84C", "navy": "1B3D6E", "lightgray": "C0C0C0", "lightgrey": "C0C0C0",
    "darkred": "8B0000", "darkblue": "1E2761", "teal": "00838F", "purple": "7B68EE",
    "红": "C00000", "蓝": "0070C0", "绿": "1A5C2A", "橙": "ED7D31",
    "灰": "808080", "黑": "000000", "金": "C9A84C", "深蓝": "1B3D6E",
}

# series dict 中各字段的别名
_SERIES_KEY_FIELDS = ("key", "column", "col", "field", "y")
_SERIES_NAME_FIELDS = ("name", "label", "title")
_SERIES_TYPE_FIELDS = ("type", "kind", "chart")
_SERIES_AXIS_FIELDS = ("axis", "yaxis", "y_axis", "side")
_SERIES_LEGEND_FIELDS = ("legend", "show_in_legend", "in_legend")

_KNOWN_SERIES_FIELDS = set(
    _SERIES_KEY_FIELDS + _SERIES_NAME_FIELDS + _SERIES_TYPE_FIELDS + _SERIES_AXIS_FIELDS
    + _SERIES_LEGEND_FIELDS
) | {"grouping", "stacked", "color", "line_width", "marker", "marker_size", "x_key", "size_key",
     "labels", "label_format", "last_point_label"}

_KNOWN_TOP_FIELDS = {
    "chart", "type", "kind", "title", "data", "df",
    "categories", "category", "x", "categories_col", "labels",
    "series", "columns", "y",
    "values", "value", "value_col", "measures", "measure", "measure_col",
    "totals", "total", "total_categories",
    "size", "size_col", "name", "series_name", "color", "marker_size",
    "style", "layout", "legend", "position", "chart_size",
    "stacked", "grouping", "colors", "connectors", "value_labels", "polish",
    "x_col", "y_col",
    "orientation", "horizontal", "annotations", "highlight", "forecast_from",
    "forecast_label", "low", "high", "current", "average", "y_axis",
    "subtitle", "sort", "format",
    "range_color", "average_color", "current_color",
}


# ============================================================================
# 基础工具
# ============================================================================

def _norm_token(value: Any) -> str:
    return str(value).strip().lower().replace("-", "_").replace(" ", "_")


def _lookup_alias(table: Dict[str, Any], value: Any, *, what: str, errors: List[str]) -> Any:
    token = _norm_token(value)
    if token in table:
        return table[token]
    suggestion = difflib.get_close_matches(token, list(table), n=1, cutoff=0.6)
    hint = f" — 是否想用 '{suggestion[0]}'?" if suggestion else ""
    allowed = "/".join(sorted({str(k) for k in table if k and not _is_cjk(k)})[:12])
    errors.append(f"{what}: 无法识别 '{value}'{hint} 允许的值如: {allowed}")
    return None


def _is_cjk(text: str) -> bool:
    return any("一" <= ch <= "鿿" for ch in text)


def normalize_color(value: Any, *, what: str = "color", errors: Optional[List[str]] = None) -> Optional[str]:
    """颜色容错归一化：'#C00000' / 'c00000' / '#f00' / 'red' / '红' → 'C00000'。"""
    if value is None:
        return None
    token = str(value).strip().lstrip("#")
    lowered = token.lower()
    if lowered in _NAMED_COLORS:
        return _NAMED_COLORS[lowered]
    if len(token) == 3 and all(c in "0123456789abcdefABCDEF" for c in token):
        token = "".join(ch * 2 for ch in token)
    if len(token) == 6 and all(c in "0123456789abcdefABCDEF" for c in token):
        return token.upper()
    if errors is not None:
        errors.append(f"{what}: 无法识别颜色 '{value}'（支持 #RRGGBB / RRGGBB / 常用颜色名）")
    return None


# A string "looks like" an Excel formatCode if it carries a digit placeholder,
# a percent/currency/text symbol, or a date/time pattern (repeated date letters,
# or date letters on both sides of a separator). Bare words like "number" do
# not — passing those straight through renders garbage on the axis.
_FORMAT_CODE_RE = re.compile(
    r'[0#?%$¥€£@]|[ymdhs]{2,}|[ymd]\s?[/\-.]\s?[ymd]', re.IGNORECASE
)


def supported_number_format_tokens() -> List[str]:
    """Aliases accepted by ``format`` (besides raw Excel formatCodes)."""
    return sorted(set(_NUMBER_FORMAT_ALIASES))


def _looks_like_format_code(text: str) -> bool:
    return bool(_FORMAT_CODE_RE.search(text))


def _normalize_number_format(value: Any, *, errors: Optional[List[str]] = None) -> Optional[str]:
    """Resolve a ``format`` value to an Excel formatCode.

    Order: known alias (e.g. ``percent`` -> ``0%``, ``times`` -> ``0"x"``) →
    raw formatCode passthrough (``0.0%``, ``#,##0``, ``yyyy-mm-dd``, ``"¥"0`` …).
    An unrecognized token that is *not* a plausible formatCode (e.g. ``number``)
    is **not** passed through — it would render as literal garbage — instead it
    yields a did-you-mean error (when ``errors`` is provided) or a warning, and
    falls back to no explicit number format.

    Note on percentages: ``0%`` (and friends) multiply by 100, so pass a
    fraction (``0.27`` -> ``27%``). A value already in percentage points
    (``27``) needs a literal format like ``0"%"`` instead.
    """
    if value is None:
        return None
    token = _norm_token(value)
    if token in _NUMBER_FORMAT_ALIASES:
        return _NUMBER_FORMAT_ALIASES[token]
    text = str(value)
    if _looks_like_format_code(text):
        return text  # advanced: caller passed a real Excel formatCode
    allowed = ", ".join(supported_number_format_tokens())
    suggestion = difflib.get_close_matches(token, list(_NUMBER_FORMAT_ALIASES), n=1)
    hint = f"（是否想用 '{suggestion[0]}'？）" if suggestion else ""
    msg = (f"format: 无法识别数字格式 '{value}'{hint} —— 既不是别名也不像 Excel 格式码，"
           f"已忽略。可用别名: {allowed}；或直接传格式码如 '0'、'0.0%'、'#,##0'、'yyyy-mm-dd'")
    if errors is not None:
        errors.append(msg)
    else:
        warnings.warn(msg, stacklevel=2)
    return None


def _coerce_length(value: Any, *, what: str, errors: List[str]):
    """长度容错：数字按英寸；'2.5cm' / '1.5in' / '36pt' 字符串也接受。"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return Inches(float(value))
    text = str(value).strip().lower()
    try:
        if text.endswith("cm"):
            return Cm(float(text[:-2]))
        if text.endswith("in") or text.endswith("inch"):
            return Inches(float(text.rstrip("inch")))
        if text.endswith("pt"):
            return Pt(float(text[:-2]))
        return Inches(float(text))
    except ValueError:
        errors.append(f"{what}: 无法解析长度 '{value}'（数字按英寸，或 '2.5cm'/'1.5in'）")
        return None


def _coerce_bool(value: Any) -> bool:
    if isinstance(value, str):
        return _norm_token(value) not in {"false", "no", "off", "0", "none", "否", "无"}
    return bool(value)


def _coerce_dataframe(spec: Dict, df: Optional[pd.DataFrame], errors: List[str]) -> Optional[pd.DataFrame]:
    data = spec.get("data", spec.get("df"))
    if data is None:
        if df is None:
            errors.append("data: 缺少数据 — 在 spec['data'] 中提供，或通过 df 参数传入")
            return None
        return df
    if isinstance(data, pd.DataFrame):
        return data
    if isinstance(data, dict):
        if "columns" in data and "rows" in data:
            try:
                return pd.DataFrame(data["rows"], columns=data["columns"])
            except Exception as exc:
                errors.append(f"data: columns/rows 形式解析失败 — {exc}")
                return None
        try:
            return pd.DataFrame(data)
        except Exception as exc:
            errors.append(f"data: dict-of-lists 解析失败（各列长度需一致）— {exc}")
            return None
    if isinstance(data, list):
        if not data:
            errors.append("data: 数据为空列表")
            return None
        try:
            return pd.DataFrame(data)
        except Exception as exc:
            errors.append(f"data: records 列表解析失败 — {exc}")
            return None
    if isinstance(data, str):
        try:
            if data.lower().endswith((".xlsx", ".xls")):
                return pd.read_excel(data)
            return pd.read_csv(data)
        except Exception as exc:
            errors.append(f"data: 读取文件 '{data}' 失败 — {exc}")
            return None
    errors.append(f"data: 不支持的数据类型 {type(data).__name__}（支持 DataFrame / dict / records / CSV 路径）")
    return None


def _resolve_column(df: pd.DataFrame, name: Any, *, what: str, errors: List[str]) -> Optional[str]:
    """列名解析，带 strip 容错和 did-you-mean 建议。"""
    if name in df.columns:
        return name
    stripped = str(name).strip()
    if stripped in df.columns:
        return stripped
    # 大小写不敏感匹配（英文列名常见笔误）
    lowered = {str(c).lower(): c for c in df.columns}
    if stripped.lower() in lowered:
        return lowered[stripped.lower()]
    matches = difflib.get_close_matches(stripped, [str(c) for c in df.columns], n=1, cutoff=0.4)
    hint = f" — 是否想用 '{matches[0]}'?" if matches else ""
    cols = ", ".join(str(c) for c in df.columns)
    errors.append(f"{what}: 找不到列 '{name}'{hint} 可用列: {cols}")
    return None


def _first(spec: Dict, fields, default=None):
    for f in fields:
        if f in spec and spec[f] is not None:
            return spec[f]
    return default


# ============================================================================
# 归一化结果
# ============================================================================

@dataclass
class NormalizedSpec:
    """归一化后的渲染计划：kind + 直接可调用底层 API 的参数。"""

    kind: str
    df: Optional[pd.DataFrame] = None
    kwargs: Dict[str, Any] = field(default_factory=dict)
    series_overrides: List[Dict] = field(default_factory=list)
    post_build: List = field(default_factory=list)  # [(callable(chart, slide)), ...]
    warnings: List[str] = field(default_factory=list)

    def summary(self) -> Dict[str, Any]:
        out = {"chart": self.kind, "warnings": self.warnings}
        if self.df is not None:
            out["rows"] = len(self.df)
            out["columns"] = [str(c) for c in self.df.columns]
        if self.kind == "combo":
            out["categories"] = self.kwargs.get("categories_col")
            out["series"] = [
                {k: s.get(k) for k in ("key", "name", "type", "axis", "grouping") if s.get(k)}
                for s in self.kwargs.get("series_config", [])
            ]
        return out


# ============================================================================
# 主入口
# ============================================================================

def render_chart(slide, spec: Dict, df: Optional[pd.DataFrame] = None, *, quiet: bool = True):
    """根据声明式 spec 在 slide 上创建图表。

    Args:
        slide: pptx Slide 对象
        spec: 图表 spec dict（见 SPEC.md 或 chart_spec_reference()）
        df: 可选的数据 DataFrame；spec['data'] 缺省时使用
        quiet: 默认 True，吞掉引擎内部的调试输出

    Returns:
        Chart 对象

    Raises:
        SpecError: spec 不合法时，一次性报告全部问题
    """
    plan = normalize_spec(spec, df)
    return _execute(slide, plan, quiet=quiet)


def validate_spec(spec: Dict, df: Optional[pd.DataFrame] = None) -> Dict[str, Any]:
    """只校验不渲染。返回 {ok, errors, warnings, normalized}，供调用方自检。"""
    try:
        plan = normalize_spec(spec, df)
    except SpecError as exc:
        return {"ok": False, "errors": exc.errors, "warnings": [], "normalized": None}
    return {"ok": True, "errors": [], "warnings": plan.warnings, "normalized": plan.summary()}


def normalize_spec(spec: Dict, df: Optional[pd.DataFrame] = None) -> NormalizedSpec:
    """归一化 + 校验 spec，返回可执行的渲染计划。"""
    if not isinstance(spec, dict):
        raise SpecError([f"spec 必须是 dict，收到 {type(spec).__name__}"])

    errors: List[str] = []
    warnings: List[str] = []

    for key in spec:
        if key not in _KNOWN_TOP_FIELDS:
            matches = difflib.get_close_matches(str(key), list(_KNOWN_TOP_FIELDS), n=1, cutoff=0.6)
            hint = f"（是否想用 '{matches[0]}'?）" if matches else ""
            warnings.append(f"忽略未知字段 '{key}' {hint}")

    raw_kind = _first(spec, ("chart", "type", "kind"), "combo")
    kind = _lookup_alias(_CHART_KIND_ALIASES, raw_kind, what="chart", errors=errors)
    if errors:
        raise SpecError(errors)

    frame = _coerce_dataframe(spec, df, errors)
    if frame is None:
        raise SpecError(errors)

    if kind == "combo":
        plan = _normalize_combo(spec, frame, raw_kind, errors, warnings)
    elif kind == "waterfall":
        plan = _normalize_waterfall(spec, frame, errors, warnings)
    elif kind == "range":
        plan = _normalize_range(spec, frame, errors, warnings)
    else:
        plan = _normalize_xy(spec, frame, kind, errors, warnings)

    if errors:
        raise SpecError(errors)

    # GTM 单位副标题行（range 在 create_range_chart 内原生处理）
    subtitle = spec.get("subtitle")
    if subtitle and kind != "range":
        plan.post_build.insert(
            0, lambda chart, slide, _s=str(subtitle): _append_subtitle(chart, _s))

    plan.warnings = warnings
    return plan


def _append_subtitle(chart, subtitle: str) -> None:
    from .polish import style_chart_title

    if not chart.has_title:
        return
    para = chart.chart_title.text_frame.add_paragraph()
    para.text = subtitle
    style_chart_title(chart)


def chart_spec_reference() -> str:
    """返回 spec 格式速查表（可直接拼进 LLM 提示词）。"""
    return _SPEC_REFERENCE


# ============================================================================
# combo 归一化
# ============================================================================

def _normalize_combo(spec, df, raw_kind, errors, warnings) -> NormalizedSpec:
    from .layout import (
        CategoryAxisConfig, ChartLayoutConfig, LegendConfig, ValueAxisConfig,
    )
    from .date_axis import DateAxisConfig
    from .styles import StyleConfig

    df = df.copy()

    # ---- categories 列 ----
    cat_raw = _first(spec, ("categories", "category", "x", "categories_col", "labels"))
    if cat_raw is None:
        cat_col = _infer_categories_col(df)
        warnings.append(f"categories 未指定，自动使用列 '{cat_col}'")
    else:
        cat_col = _resolve_column(df, cat_raw, what="categories", errors=errors)
    if cat_col is None:
        raise SpecError(errors)

    # ---- 全局 grouping（spec 级 stacked 便捷开关） ----
    default_grouping = None
    if "stacked" in spec and _coerce_bool(spec["stacked"]):
        default_grouping = "stacked"
    if spec.get("grouping") is not None:
        default_grouping = _lookup_alias(
            _GROUPING_ALIASES, spec["grouping"], what="grouping", errors=errors)

    default_type = _CHART_DEFAULT_SERIES_TYPE.get(_norm_token(raw_kind), "bar")

    # ---- contribution 模式：堆叠分项 + 合计线（GTM 贡献分解图） ----
    is_contribution = _norm_token(raw_kind) in _CONTRIBUTION_ALIASES
    raw_series = _first(spec, ("series", "columns", "y"))
    series_inferred = raw_series is None and not is_contribution
    if is_contribution:
        default_grouping = default_grouping or "stacked"
        total_raw = spec.get("total")
        if raw_series is None:
            total_col = None
            if total_raw is not None:
                total_col = _resolve_column(df, total_raw, what="total", errors=errors)
            parts = [
                c for c in df.columns
                if c != cat_col and c != total_col and pd.api.types.is_numeric_dtype(df[c])
            ]
            from .tokens import get_chart_palette, get_chart_token
            part_colors = get_chart_palette("contribution_parts")
            raw_series = [
                {"column": c, "type": "bar", "stacked": True,
                 "color": part_colors[i % len(part_colors)]}
                for i, c in enumerate(parts)
            ]
            if total_col:
                raw_series.append({
                    "column": total_col, "type": "line",
                    "color": get_chart_token("contribution_line"), "line_width": 2,
                })
            warnings.append(
                f"contribution: 分项={', '.join(map(str, parts))}"
                + (f"，合计线='{total_col}'" if total_col else ""))

    # ---- series ----
    series_config, overrides, extras = _normalize_series_list(
        raw_series, df, cat_col, default_type, default_grouping, errors, warnings)
    if not series_config and not errors:
        errors.append("series: 没有可绘制的数值列")
    if errors:
        raise SpecError(errors)

    # ---- 推断模式下的自动双轴：两列量纲差 50 倍以上 → 小量纲改右轴折线 ----
    if series_inferred and len(series_config) == 2:
        mags = [
            float(pd.to_numeric(df[s["key"]], errors="coerce").abs().max() or 0)
            for s in series_config
        ]
        if min(mags) > 0 and max(mags) / min(mags) > 50:
            small = series_config[0] if mags[0] < mags[1] else series_config[1]
            small["axis"] = "secondary"
            small["type"] = "line"
            warnings.append(
                f"两列量纲差异大（{max(mags):.0f} vs {min(mags):.3g}），"
                f"'{small['key']}' 自动改为右轴折线；显式给 series 可关闭")

    # ---- 日期轴判断 ----
    layout_spec = spec.get("layout") or {}
    if not isinstance(layout_spec, dict):
        errors.append(f"layout: 应为 dict，收到 {type(layout_spec).__name__}")
        layout_spec = {}
    x_axis_spec = layout_spec.get("x_axis") or {}
    if isinstance(x_axis_spec, str):
        x_axis_spec = {"format": x_axis_spec}

    date_format = _first(x_axis_spec, ("date_format", "format")) if x_axis_spec else None
    wants_date = (
        _coerce_bool(x_axis_spec.get("date", False))
        or (date_format and any(ch in str(date_format).lower() for ch in "ymd"))
        or pd.api.types.is_datetime64_any_dtype(df[cat_col])
    )
    if not wants_date and pd.api.types.is_string_dtype(df[cat_col]):
        # CSV 里的日期常是字符串：形如 yyyy-mm(-dd) / yyyy/mm(/dd) 时自动转日期轴
        sample = df[cat_col].astype(str).str.strip()
        if sample.str.match(r"^\d{4}[-/]\d{1,2}([-/]\d{1,2})?$").all():
            try:
                df[cat_col] = pd.to_datetime(sample)
                wants_date = True
                warnings.append(f"categories: 列 '{cat_col}' 识别为日期字符串，已启用日期轴")
            except Exception:
                pass
    if wants_date and not pd.api.types.is_datetime64_any_dtype(df[cat_col]):
        try:
            df[cat_col] = pd.to_datetime(df[cat_col])
        except Exception:
            warnings.append(f"x_axis: 列 '{cat_col}' 无法转换为日期，按普通分类轴处理")
            wants_date = False

    n = len(df)
    interval = x_axis_spec.get("interval")
    max_ticks = int(x_axis_spec.get("max_ticks", 7))

    if wants_date and not date_format:
        # 按数据粒度选默认格式：日频数据用 yyyy/mm/dd（月格式 + 等间隔抽稀
        # 必然产生 2024/01、2024/01 重复标签），月频及以上用 yyyy/mm
        deltas = df[cat_col].sort_values().diff().dropna()
        median_days = deltas.median().days if len(deltas) else 31
        date_format = "yyyy/mm/dd" if median_days <= 14 else "yyyy/mm"

    if interval is None and wants_date and n > max_ticks:
        # 按"格式化后唯一标签数"定间隔，避免重复标签
        from .polish import strftime_from_excel

        fmt = strftime_from_excel(str(date_format) if date_format else None)
        unique_labels = df[cat_col].dt.strftime(fmt).nunique()
        target = max(2, min(max_ticks, unique_labels))
        interval = max(1, math.ceil(n / target))
        warnings.append(f"x_axis: 自动设置标签间隔为每 {interval} 个点（约 {target} 个标签）")
    elif interval is None and "max_ticks" in x_axis_spec:
        interval = max(1, math.ceil(n / max_ticks))

    # ---- 布局 ----
    layout_kwargs: Dict[str, Any] = {}
    title = _first(spec, ("title",)) or layout_spec.get("title")
    if title:
        layout_kwargs["title"] = str(title)

    legend_cfg, legend_off = _normalize_legend(
        _first(spec, ("legend",), layout_spec.get("legend")), LegendConfig, errors)
    if legend_cfg is not None:
        layout_kwargs["legend_config"] = legend_cfg

    y_axis = layout_spec.get("y_axis") or layout_spec.get("value_axis")
    if y_axis is not None:
        layout_kwargs["value_axis_config"] = _normalize_value_axis(
            y_axis, ValueAxisConfig, "layout.y_axis", errors)
    y2_axis = layout_spec.get("y2_axis") or layout_spec.get("secondary_axis")
    if y2_axis is not None:
        layout_kwargs["secondary_value_axis_config"] = _normalize_value_axis(
            y2_axis, ValueAxisConfig, "layout.y2_axis", errors)

    if wants_date:
        layout_kwargs["date_axis_config"] = DateAxisConfig(
            major_unit=float(interval or 1),
            number_format=str(date_format) if date_format else "yyyy/mm",
        )
    elif x_axis_spec:
        layout_kwargs["category_axis_config"] = CategoryAxisConfig(
            number_format=str(date_format) if date_format else None,
            font_size_pt=float(x_axis_spec.get("font_size", 9)),
            font_name=str(x_axis_spec.get("font", "微软雅黑")),
        )

    layout_config = ChartLayoutConfig(**layout_kwargs)

    # ---- 样式 ----
    style_config = _normalize_style(spec.get("style"), StyleConfig, errors, warnings)
    if style_config is None and default_type == "line":
        # 纯折线图默认 1.75pt，1pt 单线在投影/打印下太弱
        style_config = StyleConfig(line_width_pt=1.75)

    # ---- 位置 / 大小 ----
    position, size = _normalize_geometry(spec, errors)

    # ---- 方向 ----
    orientation = "vertical"
    raw_orient = spec.get("orientation")
    if raw_orient is None and _coerce_bool(spec.get("horizontal", False)):
        orientation = "horizontal"
    elif raw_orient is not None and _norm_token(raw_orient) in _HORIZONTAL_ALIASES:
        orientation = "horizontal"

    # ---- 排序（单系列排名图）：desc = 显示上从大到小 ----
    sort_raw = spec.get("sort")
    if sort_raw and len(series_config) == 1:
        asc_intent = _norm_token(sort_raw) in ("asc", "ascending", "升序")
        # 横向条形图第一行类目画在底部 → 数据序与视觉序相反
        data_ascending = asc_intent if orientation != "horizontal" else not asc_intent
        df = df.sort_values(series_config[0]["key"], ascending=data_ascending).reset_index(drop=True)
    elif sort_raw and len(series_config) > 1:
        warnings.append("sort: 多系列图不支持自动排序，已忽略")

    kwargs = {
        "categories_col": cat_col,
        "series_config": series_config,
        "position": position,
        "size": size,
        "style_config": style_config,
        "layout_config": layout_config,
        "polish": _coerce_bool(spec.get("polish", True)),
        "orientation": orientation,
    }

    post_build = []
    if legend_off:
        post_build.append(lambda chart, slide: setattr(chart, "has_legend", False))
    if interval and not wants_date:
        post_build.append(lambda chart, slide, _i=int(interval): _set_tick_label_skip(chart, _i))

    # ---- GTM 元素：数值标签 / 类目高亮 / 预测斜纹 / 注释 ----
    annotations = [dict(a) for a in (spec.get("annotations") or [])]
    _resolve_annotation_values(annotations, df, series_config, errors, warnings)

    for extra in extras:
        if extra.get("labels") is not None:
            opts = extra["labels"] if isinstance(extra["labels"], dict) else {}
            post_build.append(
                lambda chart, slide, _n=extra["name"], _o=dict(opts): _add_series_labels(chart, _n, _o))
        lp = extra.get("last_point_label")
        if lp:
            ann = {"type": "last_point", "series": extra["name"]}
            if isinstance(lp, dict):
                ann.update(lp)
            annotations.append(ann)

    highlight = spec.get("highlight")
    if highlight is not None:
        if not isinstance(highlight, dict):
            highlight = {"category": highlight}
        target_series = highlight.get("series") or next(
            (s["name"] for s in series_config if s.get("type", "bar") == "bar"), None)
        h_color = normalize_color(highlight.get("color", "6BA43A"), what="highlight.color", errors=errors)
        if target_series and highlight.get("category") is not None:
            post_build.append(
                lambda chart, slide, _s=target_series, _c=highlight["category"], _col=h_color:
                    _highlight(chart, _s, df, cat_col, _c, _col))

    forecast_from = spec.get("forecast_from")
    if forecast_from is not None:
        matches = df.index[df[cat_col].astype(str) == str(forecast_from)].tolist()
        if matches:
            f_idx, n_pts = int(matches[0]), len(df)
            bar_names = [s["name"] for s in series_config if s.get("type", "bar") == "bar"]
            post_build.append(
                lambda chart, slide, _names=bar_names, _i=f_idx, _n=n_pts: _apply_forecast(chart, _names, _i, _n))
            annotations.append({"type": "vline", "index": f_idx, "label": spec.get("forecast_label", "预测")})
        else:
            warnings.append(f"forecast_from: 找不到类目 '{forecast_from}'，已忽略")

    if annotations:
        from .polish import strftime_from_excel

        strf = strftime_from_excel(str(date_format)) if date_format else "%Y/%m"
        post_build.append(
            lambda chart, slide, _ann=annotations, _df=df, _cat=cat_col, _sc=series_config,
                   _pos=position, _size=size, _fmt=strf:
                _annotate(chart, slide, _df, _cat, _sc, _ann, _pos, _size, _fmt))

    return NormalizedSpec(
        kind="combo", df=df, kwargs=kwargs,
        series_overrides=overrides, post_build=post_build,
    )


def _resolve_annotation_values(annotations, df, series_config, errors, warnings):
    """注释免算数：average 缺 value 时引擎算均值；band 给 quantiles 时算分位数。

    模型自己算均值/分位数是最不可靠的环节——能让引擎算的都让引擎算。
    """
    def col_of(ref):
        for s in series_config:
            if ref in (s["name"], s["key"]):
                return s["key"]
        return ref if ref in df.columns else None

    for ann in annotations:
        kind = _norm_token(ann.get("type", "hline"))
        series_ref = ann.get("series")

        if kind in ("average", "hline", "reference") and ann.get("value") is None:
            col = col_of(series_ref) if series_ref else None
            if col is None:
                errors.append(f"annotations[{kind}]: 缺少 value，或用 'series' 指定数据列由引擎计算均值")
                continue
            mean = float(pd.to_numeric(df[col], errors="coerce").mean())
            ann["value"] = mean
            if not ann.get("label"):
                from .annotations import format_value
                ann["label"] = f"均值 {format_value(mean, ann.get('format'))}"

        if kind == "band" and ann.get("from") is None and ann.get("quantiles"):
            col = col_of(series_ref) if series_ref else None
            if col is None:
                errors.append("annotations[band]: 用 quantiles 时需要 'series' 指定数据列")
                continue
            q_lo, q_hi = ann["quantiles"]
            values = pd.to_numeric(df[col], errors="coerce")
            ann["from"] = float(values.quantile(float(q_lo)))
            ann["to"] = float(values.quantile(float(q_hi)))
            warnings.append(
                f"annotations[band]: 按 {q_lo}~{q_hi} 分位自动计算区间 "
                f"[{ann['from']:.4g}, {ann['to']:.4g}]")


def _add_series_labels(chart, name, opts):
    from .annotations import add_value_labels

    add_value_labels(
        chart, name,
        number_format=_normalize_number_format(opts.get("format")),
        position=_norm_token(opts.get("position", "outside")),
        color=opts.get("color"),
        font_size_pt=float(opts.get("font_size", 9)),
    )


def _highlight(chart, series_name, df, cat_col, category, color):
    from .annotations import highlight_category

    highlight_category(chart, series_name, df, cat_col, category, color or "6BA43A")


def _apply_forecast(chart, series_names, from_index, n_points):
    from .annotations import apply_forecast_pattern

    for name in series_names:
        apply_forecast_pattern(chart, name, from_index, n_points)


def _annotate(chart, slide, df, cat_col, series_config, annotations, position, size, date_fmt):
    from .annotations import apply_annotations

    apply_annotations(
        chart, slide,
        df=df, categories_col=cat_col, series_config=series_config,
        annotations=annotations, position=position, size=size, date_format=date_fmt)


def _normalize_range(spec, df, errors, warnings) -> NormalizedSpec:
    """估值区间图：low/high 必填，current/average 可选。"""
    df = df.copy()
    cat_raw = _first(spec, ("categories", "category", "x", "categories_col", "labels"))
    if cat_raw is None:
        cat_col = _infer_categories_col(df)
        warnings.append(f"categories 未指定，自动使用列 '{cat_col}'")
    else:
        cat_col = _resolve_column(df, cat_raw, what="categories", errors=errors)

    numeric = [c for c in df.columns if c != cat_col and pd.api.types.is_numeric_dtype(df[c])]

    def pick(fields, what, auto_index=None):
        raw = _first(spec, fields)
        if raw is None:
            if auto_index is not None and len(numeric) > auto_index:
                col = numeric[auto_index]
                warnings.append(f"{what} 未指定，自动使用数值列 '{col}'")
                return col
            return None
        return _resolve_column(df, raw, what=what, errors=errors)

    low_col = pick(("low", "min"), "low", auto_index=0)
    high_col = pick(("high", "max"), "high", auto_index=1)
    current_col = pick(("current", "latest"), "current",
                       auto_index=2 if len(numeric) > 2 else None)
    average_col = pick(("average", "avg", "mean"), "average",
                       auto_index=3 if len(numeric) > 3 else None)
    if low_col is None or high_col is None:
        errors.append("range: 需要 low 和 high 两列（或至少两个数值列）")

    position, size = _normalize_geometry(spec, errors)
    if errors:
        raise SpecError(errors)

    kwargs = dict(
        categories_col=cat_col,
        low_col=low_col,
        high_col=high_col,
        current_col=current_col,
        average_col=average_col,
        position=position,
        size=size,
        title=str(spec["title"]) if spec.get("title") else None,
        subtitle=str(spec["subtitle"]) if spec.get("subtitle") else None,
        number_format=_normalize_number_format(spec.get("format"), errors=errors),
        sort=spec.get("sort"),
    )
    # 图例名（非估值场景可改，如「配置区间」）；缺省走 create_range_chart 默认
    for spec_key, kw in (("range_name", "range_name"), ("average_name", "average_name"),
                         ("current_name", "current_name")):
        if spec.get(spec_key) is not None:
            kwargs[kw] = str(spec[spec_key])
    # legend: "none"/"off"/"hide"/"无" 隐藏整个图例
    if _norm_token(spec.get("legend", "")) in {"none", "off", "hide", "no", "false", "隐藏", "无"}:
        kwargs["show_legend"] = False
    for spec_key, kw in (("range_color", "range_color"), ("average_color", "average_color"),
                         ("current_color", "current_color")):
        if spec.get(spec_key) is not None:
            c = normalize_color(spec[spec_key], what=spec_key, errors=errors)
            if c:
                kwargs[kw] = c

    return NormalizedSpec(kind="range", df=df, kwargs=kwargs)


def _infer_categories_col(df: pd.DataFrame) -> str:
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]) or not pd.api.types.is_numeric_dtype(df[col]):
            return col
    return df.columns[0]


def _normalize_series_list(raw_series, df, cat_col, default_type, default_grouping, errors, warnings):
    series_config: List[Dict] = []
    overrides: List[Dict] = []
    extras: List[Dict] = []  # labels / last_point_label 等 GTM 元素

    if raw_series is None:
        numeric_cols = [
            c for c in df.columns
            if c != cat_col and pd.api.types.is_numeric_dtype(df[c])
        ]
        if numeric_cols:
            warnings.append(f"series 未指定，自动绘制数值列: {', '.join(map(str, numeric_cols))}")
        raw_series = numeric_cols

    if isinstance(raw_series, (str, dict)):
        raw_series = [raw_series]

    for i, item in enumerate(raw_series):
        where = f"series[{i}]"
        if isinstance(item, str):
            item = {"key": item}
        if not isinstance(item, dict):
            errors.append(f"{where}: 应为字符串或 dict，收到 {type(item).__name__}")
            continue

        for k in item:
            if k not in _KNOWN_SERIES_FIELDS:
                matches = difflib.get_close_matches(str(k), sorted(_KNOWN_SERIES_FIELDS), n=1, cutoff=0.6)
                hint = f"（是否想用 '{matches[0]}'?）" if matches else ""
                warnings.append(f"{where}: 忽略未知字段 '{k}' {hint}")

        key_raw = _first(item, _SERIES_KEY_FIELDS)
        if key_raw is None:
            errors.append(f"{where}: 缺少列名（用 'column' 或 'key' 指定）")
            continue
        key = _resolve_column(df, key_raw, what=f"{where}.column", errors=errors)
        if key is None:
            continue
        if not pd.api.types.is_numeric_dtype(df[key]):
            coerced = pd.to_numeric(df[key], errors="coerce")
            if coerced.isna().any():
                errors.append(f"{where}: 列 '{key}' 不是数值列，无法作为系列绘制")
                continue
            df[key] = coerced

        type_raw = _first(item, _SERIES_TYPE_FIELDS, default_type)
        series_type = _lookup_alias(_SERIES_TYPE_ALIASES, type_raw, what=f"{where}.type", errors=errors)
        axis_raw = _first(item, _SERIES_AXIS_FIELDS, "primary")
        axis = _lookup_alias(_AXIS_ALIASES, axis_raw, what=f"{where}.axis", errors=errors)

        final_type = series_type or default_type
        # 全局 stacked 只作用于柱/面积系列；叠加的合计线不参与堆叠
        grouping = default_grouping if final_type in ("bar", "area") else None
        if "stacked" in item and _coerce_bool(item["stacked"]):
            grouping = "stacked"
        if item.get("grouping") is not None:
            grouping = _lookup_alias(_GROUPING_ALIASES, item["grouping"], what=f"{where}.grouping", errors=errors)

        name = str(_first(item, _SERIES_NAME_FIELDS, key))
        cfg = {"key": key, "name": name, "type": series_type or default_type, "axis": axis or "primary"}
        if grouping:
            cfg["grouping"] = grouping
        legend_raw = _first(item, _SERIES_LEGEND_FIELDS, None)
        if legend_raw is not None and not _coerce_bool(legend_raw):
            cfg["show_in_legend"] = False
        series_config.append(cfg)

        ov = {}
        if item.get("color") is not None:
            color = normalize_color(item["color"], what=f"{where}.color", errors=errors)
            if color:
                ov["color"] = color
        if item.get("line_width") is not None:
            try:
                ov["line_width_pt"] = float(item["line_width"])
            except (TypeError, ValueError):
                errors.append(f"{where}.line_width: 应为数字（pt），收到 '{item['line_width']}'")
        if item.get("marker") is not None:
            ov["marker_style"] = _norm_token(item["marker"])
        if item.get("marker_size") is not None:
            ov["marker_size"] = int(item["marker_size"])
        if ov:
            ov["name"] = name
            overrides.append(ov)

        extra = {}
        if item.get("labels") is not None and item["labels"] is not False:
            extra["labels"] = item["labels"] if isinstance(item["labels"], dict) else (
                {"format": item["label_format"]} if item.get("label_format") else {})
        elif item.get("label_format") is not None:
            extra["labels"] = {"format": item["label_format"]}
        if item.get("last_point_label"):
            extra["last_point_label"] = item["last_point_label"]
        if extra:
            extra["name"] = name
            extras.append(extra)

    return series_config, overrides, extras


def _normalize_legend(raw, LegendConfig, errors):
    """返回 (legend_config, legend_off)。raw 为 None 时返回 (None, False) 用默认。"""
    from pptx.enum.chart import XL_LEGEND_POSITION

    if raw is None:
        return None, False
    if raw is False:
        return None, True
    if isinstance(raw, str):
        token = _norm_token(raw)
        if token in _LEGEND_OFF_VALUES:
            return None, True
        raw = {"position": raw}
    if not isinstance(raw, dict):
        errors.append(f"legend: 应为位置字符串、false 或 dict，收到 {type(raw).__name__}")
        return None, False

    pos_name = _lookup_alias(
        _LEGEND_POSITION_NAMES, raw.get("position", "bottom"), what="legend.position", errors=errors)
    pos_map = {
        "bottom": XL_LEGEND_POSITION.BOTTOM, "top": XL_LEGEND_POSITION.TOP,
        "left": XL_LEGEND_POSITION.LEFT, "right": XL_LEGEND_POSITION.RIGHT,
        "corner": XL_LEGEND_POSITION.CORNER,
    }
    return LegendConfig(
        position=pos_map.get(pos_name, XL_LEGEND_POSITION.BOTTOM),
        font_size_pt=float(raw.get("font_size", 9)),
        font_name=str(raw.get("font", "微软雅黑")),
    ), False


def _normalize_value_axis(raw, ValueAxisConfig, where, errors):
    if isinstance(raw, str):
        raw = {"format": raw}
    if not isinstance(raw, dict):
        errors.append(f"{where}: 应为 dict 或数字格式字符串，收到 {type(raw).__name__}")
        return None
    fmt = _normalize_number_format(_first(raw, ("format", "number_format")), errors=errors)
    return ValueAxisConfig(
        number_format=fmt,
        font_size_pt=float(raw.get("font_size", 9)),
        font_name=str(raw.get("font", "微软雅黑")),
        has_major_gridlines=_coerce_bool(raw.get("gridlines", False)),
        min_value=raw.get("min"),
        max_value=raw.get("max"),
        major_unit=_first(raw, ("unit", "major_unit", "interval")),
    )


def _normalize_style(raw, StyleConfig, errors, warnings):
    if raw is None:
        return None
    if isinstance(raw, str):
        raw = {"theme": raw}
    if not isinstance(raw, dict):
        errors.append(f"style: 应为 dict 或主题名字符串，收到 {type(raw).__name__}")
        return None

    from .styles import COLOR_SCHEMES

    scheme = _first(raw, ("theme", "color_scheme", "scheme", "palette"), "default")
    if isinstance(scheme, str) and scheme not in COLOR_SCHEMES:
        matches = difflib.get_close_matches(scheme, list(COLOR_SCHEMES), n=1, cutoff=0.5)
        hint = f"（是否想用 '{matches[0]}'? 已回退 default）" if matches else "（已回退 default）"
        warnings.append(f"style.theme: 未知主题 '{scheme}' {hint} 可用: {', '.join(COLOR_SCHEMES)}")
        scheme = "default"

    colors = raw.get("colors")
    normalized_colors = None
    if colors is not None:
        if isinstance(colors, str):
            colors = [colors]
        normalized_colors = [
            c for c in (normalize_color(c, what="style.colors", errors=errors) for c in colors) if c
        ] or None

    try:
        line_width = float(raw.get("line_width", 1.0))
    except (TypeError, ValueError):
        errors.append(f"style.line_width: 应为数字（pt），收到 '{raw.get('line_width')}'")
        line_width = 1.0

    return StyleConfig(
        color_scheme=scheme,
        line_width_pt=line_width,
        marker_style=_norm_token(raw.get("marker", "none")),
        marker_size=int(raw.get("marker_size", 5)),
        colors=normalized_colors,
    )


def _normalize_geometry(spec, errors):
    pos_raw = spec.get("position", (1, 2))
    size_raw = spec.get("chart_size") or _default_size_field(spec)
    if isinstance(pos_raw, dict):
        pos_raw = (pos_raw.get("left", 1), pos_raw.get("top", 2))
    if isinstance(size_raw, dict):
        size_raw = (size_raw.get("width", 8), size_raw.get("height", 4.5))

    def pair(raw, what, defaults):
        if not isinstance(raw, (list, tuple)) or len(raw) != 2:
            errors.append(f"{what}: 应为 [a, b] 两元素（单位英寸），收到 {raw!r}")
            return tuple(Inches(d) for d in defaults)
        a = _coerce_length(raw[0], what=f"{what}[0]", errors=errors)
        b = _coerce_length(raw[1], what=f"{what}[1]", errors=errors)
        return (a or Inches(defaults[0]), b or Inches(defaults[1]))

    return pair(pos_raw, "position", (1, 2)), pair(size_raw, "size", (8, 4.5))


def _default_size_field(spec):
    """spec['size'] 在 bubble 中表示 size 列，仅当它像几何尺寸时才用作图表大小。"""
    size = spec.get("size")
    if isinstance(size, (list, tuple, dict)):
        return size
    return (8, 4.5)


# ============================================================================
# waterfall 归一化
# ============================================================================

def _normalize_waterfall(spec, df, errors, warnings) -> NormalizedSpec:
    df = df.copy()

    cat_raw = _first(spec, ("categories", "category", "x", "categories_col", "labels"))
    val_raw = _first(spec, ("values", "value", "y", "value_col"))
    measure_raw = _first(spec, ("measures", "measure", "measure_col"))

    if cat_raw is None:
        cat_col = _infer_categories_col(df)
        warnings.append(f"categories 未指定，自动使用列 '{cat_col}'")
    else:
        cat_col = _resolve_column(df, cat_raw, what="categories", errors=errors)

    if val_raw is None:
        numeric = [c for c in df.columns if c != cat_col and pd.api.types.is_numeric_dtype(df[c])]
        if numeric:
            val_col = numeric[0]
            warnings.append(f"values 未指定，自动使用数值列 '{val_col}'")
        else:
            val_col = None
            errors.append("values: 找不到数值列，请用 'values' 指定")
    else:
        val_col = _resolve_column(df, val_raw, what="values", errors=errors)

    measure_col = None
    if measure_raw is not None:
        measure_col = _resolve_column(df, measure_raw, what="measures", errors=errors)
    elif cat_col is not None and val_col is not None:
        measure_col = _infer_measure_col(df, cat_col, val_col)
        if measure_col:
            warnings.append(f"measures 未指定，自动识别列 '{measure_col}'")

    totals = _first(spec, ("totals", "total", "total_categories"))
    if isinstance(totals, str):
        totals = [totals]

    colors = spec.get("colors") or {}
    if not isinstance(colors, dict):
        errors.append("colors: waterfall 的 colors 应为 dict，如 {'positive': '#1A5C2A', 'negative': '#C00000'}")
        colors = {}
    color_kwargs = {}
    for spec_key, kw in (("positive", "positive_color"), ("negative", "negative_color"), ("total", "total_color")):
        value = colors.get(spec_key) or spec.get(f"{spec_key}_color")
        if value is not None:
            c = normalize_color(value, what=f"colors.{spec_key}", errors=errors)
            if c:
                color_kwargs[kw] = c

    position, size = _normalize_geometry(spec, errors)
    if errors:
        raise SpecError(errors)

    kwargs = dict(
        categories_col=cat_col,
        value_col=val_col,
        measure_col=measure_col,
        total_categories=list(totals) if totals else None,
        position=position,
        size=size,
        show_legend=_coerce_bool(spec.get("legend", False)),
        show_connectors=_coerce_bool(spec.get("connectors", True)),
        show_value_labels=_coerce_bool(spec.get("value_labels", True)),
        show_y_axis=_coerce_bool(spec.get("y_axis", False)),
        title=str(spec["title"]) if spec.get("title") else None,
        **color_kwargs,
    )

    return NormalizedSpec(kind="waterfall", df=df, kwargs=kwargs)


_MEASURE_TOKENS = {
    "relative", "rel", "delta", "total", "subtotal", "absolute", "sum",
    "相对", "增减", "合计", "总计", "小计",
}


def _infer_measure_col(df, cat_col, val_col) -> Optional[str]:
    for col in df.columns:
        if col in (cat_col, val_col):
            continue
        values = df[col].dropna().astype(str).str.strip().str.lower()
        if len(values) and values.isin(_MEASURE_TOKENS).all():
            return col
    return None


# ============================================================================
# scatter / bubble 归一化
# ============================================================================

def _normalize_xy(spec, df, kind, errors, warnings) -> NormalizedSpec:
    df = df.copy()
    numeric = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

    x_raw = _first(spec, ("x", "x_col", "categories"))
    y_raw = _first(spec, ("y", "y_col", "values"))

    if x_raw is None and len(numeric) >= 2:
        x_raw = numeric[0]
        warnings.append(f"x 未指定，自动使用数值列 '{x_raw}'")
    if y_raw is None and len(numeric) >= 2:
        y_raw = numeric[1]
        warnings.append(f"y 未指定，自动使用数值列 '{y_raw}'")

    x_col = _resolve_column(df, x_raw, what="x", errors=errors) if x_raw is not None else None
    y_col = _resolve_column(df, y_raw, what="y", errors=errors) if y_raw is not None else None
    if x_raw is None or y_raw is None:
        errors.append(f"{kind}: 需要至少两个数值列，或显式指定 'x' 和 'y'")

    kwargs: Dict[str, Any] = {}
    if kind == "bubble":
        size_raw = _first(spec, ("size", "size_col"))
        if isinstance(size_raw, (list, tuple, dict)):
            size_raw = None  # 是几何尺寸不是列名
        if size_raw is None:
            remaining = [c for c in numeric if c not in (x_col, y_col)]
            if remaining:
                size_raw = remaining[0]
                warnings.append(f"size 未指定，自动使用数值列 '{size_raw}'")
            else:
                errors.append("bubble: 缺少气泡大小列，请用 'size' 指定")
        size_col = _resolve_column(df, size_raw, what="size", errors=errors) if size_raw else None
        kwargs["size_col"] = size_col

    name = _first(spec, ("name", "series_name", "title"))
    if name:
        kwargs["series_name"] = str(name)
    color = spec.get("color")
    if color is not None:
        c = normalize_color(color, what="color", errors=errors)
        if c:
            kwargs["color"] = c
    if spec.get("marker_size") is not None:
        kwargs["marker_size"] = int(spec["marker_size"])

    position, size = _normalize_geometry(spec, errors)
    if errors:
        raise SpecError(errors)

    kwargs.update({"x_col": x_col, "y_col": y_col, "position": position, "size": size})

    post_build = []
    title = spec.get("title")
    if title:
        post_build.append(lambda chart, slide, _t=str(title): _set_chart_title(chart, _t))

    return NormalizedSpec(kind=kind, df=df, kwargs=kwargs, post_build=post_build)


# ============================================================================
# 执行
# ============================================================================

def _execute(slide, plan: NormalizedSpec, *, quiet: bool):
    from .api import create_combo_chart
    from .range_chart import create_range_chart
    from .scatter import create_bubble_chart, create_scatter_chart
    from .waterfall import create_waterfall_chart

    sink = io.StringIO()
    ctx = contextlib.redirect_stdout(sink) if quiet else contextlib.nullcontext()

    with ctx:
        if plan.kind == "combo":
            chart = create_combo_chart(slide=slide, df=plan.df, **plan.kwargs)
        elif plan.kind == "waterfall":
            chart = create_waterfall_chart(slide, plan.df, **plan.kwargs)
        elif plan.kind == "range":
            chart = create_range_chart(slide, plan.df, **plan.kwargs)
        elif plan.kind == "scatter":
            chart = create_scatter_chart(slide, plan.df, **plan.kwargs)
        elif plan.kind == "bubble":
            chart = create_bubble_chart(slide, plan.df, **plan.kwargs)
        else:  # pragma: no cover - normalize_spec 已保证
            raise SpecError([f"未知图表类型: {plan.kind}"])

        for override in plan.series_overrides:
            _apply_series_override(chart, override)
        for hook in plan.post_build:
            hook(chart, slide)

    return chart


def _set_chart_title(chart, title: str):
    from .polish import style_chart_title

    chart.has_title = True
    chart.chart_title.text_frame.text = title
    style_chart_title(chart)


def _set_tick_label_skip(chart, interval: int):
    """对普通分类轴设置标签间隔（每 N 个数据点显示一个标签）。"""
    cat_ax_list = chart._element.findall(f".//{{{NAMESPACES['c']}}}catAx")
    if not cat_ax_list:
        return
    cat_ax = cat_ax_list[0]
    skip = cat_ax.find(f"{{{NAMESPACES['c']}}}tickLblSkip")
    if skip is None:
        skip = etree.Element(f"{{{NAMESPACES['c']}}}tickLblSkip")
        auto = cat_ax.find(f"{{{NAMESPACES['c']}}}auto")
        if auto is not None:
            cat_ax.insert(list(cat_ax).index(auto) + 1, skip)
        else:
            cat_ax.append(skip)
    skip.set("val", str(int(interval)))


def _apply_series_override(chart, override: Dict):
    """按系列名定位 <c:ser> 并覆盖颜色/线宽/标记。"""
    from .styles import apply_series_style

    target = override["name"]
    for ser in chart._chartSpace.findall(f".//{{{NAMESPACES['c']}}}ser"):
        name = _series_name(ser)
        if name != target:
            continue
        line_width = None
        if override.get("line_width_pt") is not None:
            line_width = int(override["line_width_pt"] * 12700)
        elif override.get("color"):
            line_width = _existing_line_width(ser)
        # 只改线宽/标记时必须把已有颜色带回去，否则 ln 重建后颜色丢失、回退主题色
        color = override.get("color") or _existing_series_color(ser)
        apply_series_style(
            ser,
            color=color,
            line_width=line_width,
            marker_style=override.get("marker_style"),
            marker_size=override.get("marker_size"),
        )


def _series_name(ser) -> Optional[str]:
    for path in ("c:tx/c:strRef/c:strCache/c:pt/c:v", "c:tx/c:v"):
        node = ser.find(path, namespaces=NAMESPACES)
        if node is not None and node.text:
            return node.text
    return None


def _existing_series_color(ser) -> Optional[str]:
    for path in ("c:spPr/a:ln/a:solidFill/a:srgbClr", "c:spPr/a:solidFill/a:srgbClr"):
        node = ser.find(path, namespaces=NAMESPACES)
        if node is not None:
            return node.get("val")
    return None


def _existing_line_width(ser) -> Optional[int]:
    ln = ser.find("c:spPr/a:ln", namespaces=NAMESPACES)
    if ln is not None and ln.get("w"):
        return int(ln.get("w"))
    return None


# ============================================================================
# Spec 速查表（供 LLM 提示词使用）
# ============================================================================

_SPEC_REFERENCE = """\
# ablechart Chart Spec 速查

一个 JSON dict 描述一张图。所有字段大小写不敏感、支持中英文别名、未知字段只警告不报错。

## 通用字段
- chart: combo(默认)|waterfall|scatter|bubble|range|contribution。
  别名: bar/line/area→combo, bridge/瀑布→waterfall, 区间/估值区间→range, 贡献图→contribution
- title: 图表标题
- data: dict-of-lists | records 列表 | {"columns":[...],"rows":[[...]]} | CSV/XLSX 路径 | 省略则用 df 参数
- position: [left, top] 英寸（默认 [1,2]）; chart_size: [width, height]（默认 [8,4.5]）
- orientation: "horizontal"/"横向" → 横向条形图（排名图）

## combo（柱/线/面积，单轴或双轴）
- categories: X 轴列名（省略→自动取第一个非数值列；日期列自动启用日期轴）
- series: 列名字符串数组，或对象数组:
  {"column": 列名, "name": 显示名, "type": "bar|line|area", "axis": "left|right",
   "stacked": true, "color": "#C00000", "line_width": 2, "marker": "circle"}
- stacked: true → 全部柱状系列堆叠; grouping: "percent" → 百分比堆叠
- style: {"theme": "able_finance", "colors": ["#1B3D6E","#C9A84C"], "line_width": 1.5, "marker": "none"}
  可用主题: advisory/gtm/midnight/charcoal/able_finance/able_warm/tech_blue/state_red/esg_green/dark_pro/daybreak/macro_research
- layout:
  - legend: "bottom|top|left|right|corner|none" 或 {"position":..,"font_size":9}
  - y_axis / y2_axis: {"format": "percent|0.00|#,##0", "min":0, "max":1, "unit":0.2, "gridlines": false}
  - x_axis: {"date_format": "yyyy/mm", "max_ticks": 7} 或 {"interval": 5}

## waterfall（瀑布桥图）
- categories: 标签列; values: 数值列; measures: 度量列(relative/total，可省略自动识别)
- totals: ["期末收益"] 指定合计类目（替代 measures 列）
- colors: {"positive": "#1A5C2A", "negative": "#C00000", "total": "#1B3D6E"}
- value_labels: true(默认); connectors: false(默认); legend: false(默认)

## scatter / bubble
- x, y: 数值列名（省略→自动取前两个数值列）; bubble 加 size: 大小列名
- name: 系列名; color: "#1E2761"; marker_size: 9

## range（估值区间图：灰色区间条 + 均值横杠 + 当前菱形）
- low/high: 区间列（必填，省略时按数值列顺序自动取）; average/current: 可选
- 场景: PE/PB 历史分位、利差区间、波动率锥、仓位历史区间

## contribution（贡献分解：堆叠分项 + 合计橙线，GTM 宏观图标配）
- total: 合计列名 → 自动画成橙色折线，其余数值列自动堆叠（配色自动跳过橙色）

## GTM 元素（combo 系列可用）
- series 内: "labels": {"format": "0%", "position": "outside|inside|center", "color": "white"} 柱上数值标签
- series 内: "last_point_label": {"format": "0.0%"} 末点圆点+「日期 数值」标注
- 顶层 highlight: {"category": "全部A股", "color": "#6BA43A"} 排名图单类目高亮
- 顶层 forecast_from: "2025F" → 该类目起柱体变斜纹 + 虚线分隔（forecast_label 自定义文字）
- 顶层 annotations: [
    {"type": "average", "value": 0.024, "label": "长期均值", "label_at": "left|center|right"},
    {"type": "band", "from": 0.02, "to": 0.03, "label": "目标区间"},
    {"type": "vline", "category": "2024Q1", "label": "政策转向"}
  ]

## 校验
- validate_spec(spec, df) → {"ok", "errors", "warnings", "normalized"}，可先自检再渲染
- 列名拼错会返回"是否想用 'xxx'"建议
"""
