"""
预设图表配置库

包含常用的图表配置，可直接用于模板替换
"""

from pptx.enum.chart import XL_LEGEND_POSITION

from .date_axis import DateAxisConfig
from .layout import ChartLayoutConfig, LegendConfig, ValueAxisConfig
from .styles import StyleConfig


# ============================================================================
# 预设图表配置
# ============================================================================

def get_chart1_config(df):
    """
    图表1：当年收益率走势（柱状图+折线图）
    
    特点：
    - 柱状图（沪深300指数，右轴）
    - 折线图（组合收益率，左轴）
    - 日期轴：2024/01 格式
    - 双轴配置
    """
    label_interval = len(df) // 7
    
    return {
        "df": df,
        "categories_col": "分类",
        "series_config": [
            {
                'name': '沪深300指数(收盘价)',
                'key': '沪深300指数(收盘价)',
                'type': 'bar',
                'axis': 'secondary'
            },
            {
                'name': '组合收益率（左轴）',
                'key': '组合收益率（左轴）',
                'type': 'line',
                'axis': 'primary'
            },
        ],
        "style_config": StyleConfig(
            color_scheme="aim00",
            line_width_pt=2.0,
            marker_style="none",
        ),
        "layout_config": ChartLayoutConfig(
            title="组合2024年以来收益率走势图",
            legend_config=LegendConfig(
                position=XL_LEGEND_POSITION.TOP,
                font_size_pt=9,
                font_name="黑体",
            ),
            value_axis_config=ValueAxisConfig(
                number_format="0%",
                font_size_pt=9,
                font_name="黑体",
                has_major_gridlines=False,
            ),
            secondary_value_axis_config=ValueAxisConfig(
                number_format="#,##0",
                font_size_pt=9,
                font_name="黑体",
                has_major_gridlines=False,
                min_value=2950,
                max_value=4450,
                major_unit=200,
            ),
            date_axis_config=DateAxisConfig(
                major_unit=label_interval,
                number_format='yyyy/mm',
            ),
        ),
    }


def get_chart2_config(df):
    """
    图表2：组合成立以来收益率走势（面积图+折线图）
    
    特点：
    - 面积图（组合规模，右轴）
    - 折线图（累计收益率，左轴）
    - 长期数据（2013-2024，2800+个点）
    - 日期轴自动间隔
    """
    label_interval = len(df) // 7
    
    return {
        "df": df,
        "categories_col": "分类",
        "series_config": [
            {
                'name': '组合规模(万元)',
                'key': '组合规模(万元)',
                'type': 'area',
                'axis': 'secondary'
            },
            {
                'name': '累计收益率',
                'key': '累计收益率',
                'type': 'line',
                'axis': 'primary'
            },
        ],
        "style_config": StyleConfig(
            color_scheme="aim00",
            line_width_pt=2.0,
            marker_style="none",
        ),
        "layout_config": ChartLayoutConfig(
            title="组合成立以来收益率走势",
            legend_config=LegendConfig(
                position=XL_LEGEND_POSITION.TOP,
                font_size_pt=9,
                font_name="黑体",
            ),
            value_axis_config=ValueAxisConfig(
                number_format="0%",
                font_size_pt=9,
                font_name="黑体",
                has_major_gridlines=True,
            ),
            secondary_value_axis_config=ValueAxisConfig(
                number_format="#,##0",
                font_size_pt=9,
                font_name="黑体",
                has_major_gridlines=False,
                min_value=0,
                max_value=160000,
                major_unit=24000,
            ),
            date_axis_config=DateAxisConfig(
                major_unit=label_interval,
                number_format='yyyy/mm',
            ),
        ),
    }


def get_chart3_config(df):
    """
    图表3：权益仓位走势（面积图+折线图）
    
    特点：
    - 面积图（沪深300指数，右轴）
    - 折线图（权益仓位，左轴）
    - 百分比格式
    """
    label_interval = len(df) // 7
    
    return {
        "df": df,
        "categories_col": "分类",
        "series_config": [
            {
                'name': '沪深300指数(收盘价)',
                'key': '沪深300指数(收盘价)',
                'type': 'area',
                'axis': 'secondary'
            },
            {
                'name': '权益仓位（人社部口径）',
                'key': '权益仓位（人社部口径）',
                'type': 'line',
                'axis': 'primary'
            },
        ],
        "style_config": StyleConfig(
            color_scheme="aim00",
            line_width_pt=2.0,
            marker_style="none",
        ),
        "layout_config": ChartLayoutConfig(
            title="组合权益仓位走势图",
            legend_config=LegendConfig(
                position=XL_LEGEND_POSITION.TOP,
                font_size_pt=9,
                font_name="黑体",
            ),
            value_axis_config=ValueAxisConfig(
                number_format="0%",
                font_size_pt=9,
                font_name="黑体",
                has_major_gridlines=False,
                min_value=0.15,
                max_value=0.30,
            ),
            secondary_value_axis_config=ValueAxisConfig(
                number_format="#,##0",
                font_size_pt=9,
                font_name="黑体",
                has_major_gridlines=False,
                min_value=2950,
                max_value=4450,
                major_unit=200,
            ),
            date_axis_config=DateAxisConfig(
                major_unit=label_interval,
                number_format='yyyy/mm',
            ),
        ),
    }


def get_chart4_config(df):
    """
    图表4：久期走势（柱状图+折线图）
    
    特点：
    - 柱状图（中债指数，右轴）
    - 折线图（久期，左轴）
    - 两位小数格式
    """
    label_interval = len(df) // 7
    
    # 动态计算右轴范围
    bond_index_data = df['中债新综合总财富指数(收盘价)']
    bond_min = bond_index_data.min()
    bond_max = bond_index_data.max()
    bond_range = bond_max - bond_min
    bond_unit = bond_range / 6
    
    return {
        "df": df,
        "categories_col": "分类",
        "series_config": [
            {
                'name': '中债新综合总财富指数(收盘价)',
                'key': '中债新综合总财富指数(收盘价)',
                'type': 'bar',
                'axis': 'secondary'
            },
            {
                'name': '久期',
                'key': '久期',
                'type': 'line',
                'axis': 'primary'
            },
        ],
        "style_config": StyleConfig(
            color_scheme="aim00",
            line_width_pt=2.0,
            marker_style="none",
        ),
        "layout_config": ChartLayoutConfig(
            title="组合久期走势图",
            legend_config=LegendConfig(
                position=XL_LEGEND_POSITION.TOP,
                font_size_pt=9,
                font_name="黑体",
            ),
            value_axis_config=ValueAxisConfig(
                number_format="0.00",
                font_size_pt=9,
                font_name="黑体",
                has_major_gridlines=False,
                min_value=0.2,
                max_value=1.0,
            ),
            secondary_value_axis_config=ValueAxisConfig(
                number_format="#,##0.00",
                font_size_pt=9,
                font_name="黑体",
                has_major_gridlines=False,
                min_value=int(bond_min - bond_unit),
                max_value=int(bond_max + bond_unit),
                major_unit=int(bond_unit),
            ),
            date_axis_config=DateAxisConfig(
                major_unit=label_interval,
                number_format='yyyy/mm',
            ),
        ),
    }


def build_range_snapshot_preset(
    df,
    *,
    title: str,
    subtitle: str | None = None,
    insight: str | None = None,
    footnote: str | None = None,
    categories_col: str = "category",
    min_col: str = "range_min",
    max_col: str = "range_max",
    average_col: str = "average",
    current_col: str = "current",
    orientation: str = "vertical",
    range_color: str = "5F6772",
    average_color: str = "87A330",
    current_color: str = "1E88E5",
    number_format: str = "0.0x",
    axis_break: dict | None = None,
):
    """Generic range snapshot preset builder for valuation-like pages."""

    return {
        "df": df,
        "title": title,
        "subtitle": subtitle,
        "insight": insight,
        "footnote": footnote,
        "categories_col": categories_col,
        "min_col": min_col,
        "max_col": max_col,
        "average_col": average_col,
        "current_col": current_col,
        "orientation": orientation,
        "range_color": range_color,
        "average_color": average_color,
        "current_color": current_color,
        "number_format": number_format,
        "show_average_ticks": True,
        "show_current_markers": True,
        "show_current_labels": True,
        "axis_break": axis_break,
    }


def _round_up_to_step(value: float, step: float) -> float:
    import math

    return math.ceil(value / step) * step


def _build_sector_axis_break_preset(
    df,
    *,
    categories_col: str,
    max_col: str,
    break_value: float,
    tick_step: float,
) -> dict | None:
    max_series = df[max_col].astype(float)
    exceeding_categories = df.loc[max_series >= break_value, categories_col].astype(str).tolist()
    if not exceeding_categories:
        return None

    upper_max = max_series.max()
    upper_tick = _round_up_to_step(upper_max, tick_step if upper_max <= 60 else 10.0)
    lower_ticks = list(range(0, int(break_value) + 1, int(tick_step)))
    tick_values = [float(value) for value in lower_ticks if value <= break_value]
    if upper_tick not in tick_values:
        tick_values.append(float(upper_tick))

    return {
        "value": break_value,
        "categories": exceeding_categories,
        "compress_ratio": 0.22,
        "tick_values": tick_values,
        "size_inches": 0.18,
        "gap_inches": 0.08,
        "line_width_pt": 2.6,
    }


def get_vertical_global_valuation_snapshot_preset(df):
    """Vertical global valuation preset close to `jp_demo` page-level grammar."""

    return build_range_snapshot_preset(
        df,
        title="Global valuations",
        subtitle="Current and historical price-to-earning valuations",
        categories_col="market",
        min_col="range_min",
        max_col="range_max",
        average_col="average",
        current_col="current",
        orientation="vertical",
        number_format="0.0x",
    )


def get_vertical_sector_valuation_snapshot_preset(
    df,
    *,
    title: str = "Sector valuation",
    subtitle: str = "Forward price-to-earnings ratio",
    categories_col: str = "sector",
    min_col: str = "range_min",
    max_col: str = "range_max",
    average_col: str = "average",
    current_col: str = "current",
    break_value: float = 30.0,
    tick_step: float = 5.0,
):
    """Vertical sector valuation preset with default axis-break compression."""

    axis_break = _build_sector_axis_break_preset(
        df,
        categories_col=categories_col,
        max_col=max_col,
        break_value=break_value,
        tick_step=tick_step,
    )

    return build_range_snapshot_preset(
        df,
        title=title,
        subtitle=subtitle,
        categories_col=categories_col,
        min_col=min_col,
        max_col=max_col,
        average_col=average_col,
        current_col=current_col,
        orientation="vertical",
        number_format="0.0x",
        axis_break=axis_break,
    )


def get_asx200_sector_valuation_snapshot_preset(df):
    """Preset matching the `ASX 200 valuation` page grammar from jp_demo."""

    return get_vertical_sector_valuation_snapshot_preset(
        df,
        title="ASX 200 valuation",
        subtitle="Forward price-to-earnings ratio",
        break_value=50.0,
        tick_step=10.0,
    )


def get_sp500_sector_valuation_snapshot_preset(df):
    """Preset matching the `S&P 500 valuation` page grammar from jp_demo."""

    return get_vertical_sector_valuation_snapshot_preset(
        df,
        title="S&P 500 valuation",
        subtitle="Forward price-to-earnings ratio",
        break_value=30.0,
        tick_step=5.0,
    )


def get_msci_emu_sector_valuation_snapshot_preset(df):
    """Preset matching the `MSCI EMU valuation` page grammar from jp_demo."""

    return build_range_snapshot_preset(
        df,
        title="MSCI EMU valuation",
        subtitle="Forward price-to-earnings ratio",
        categories_col="sector",
        min_col="range_min",
        max_col="range_max",
        average_col="average",
        current_col="current",
        orientation="vertical",
        number_format="0.0x",
        axis_break=None,
    )


def get_msci_japan_sector_valuation_snapshot_preset(df):
    """Preset matching the `MSCI Japan valuation` page grammar from jp_demo."""

    return build_range_snapshot_preset(
        df,
        title="MSCI Japan valuation",
        subtitle="Trailing price-to-book ratio",
        categories_col="sector",
        min_col="range_min",
        max_col="range_max",
        average_col="average",
        current_col="current",
        orientation="vertical",
        number_format="0.0x",
        axis_break=None,
    )


# ============================================================================
# 预设图表映射
# ============================================================================

CHART_PRESET_FUNCTIONS = {
    "图表1 - 当年收益率走势": get_chart1_config,
    "图表2 - 成立以来收益率": get_chart2_config,
    "图表3 - 权益仓位走势": get_chart3_config,
    "图表4 - 久期走势": get_chart4_config,
}

FINANCE_PRESET_FUNCTIONS = CHART_PRESET_FUNCTIONS

RANGE_SNAPSHOT_PRESET_FUNCTIONS = {
    "估值快照 - 全市场竖版": get_vertical_global_valuation_snapshot_preset,
    "估值快照 - 行业竖版带轴断裂": get_vertical_sector_valuation_snapshot_preset,
}

VERTICAL_VALUATION_PRESET_FUNCTIONS = {
    "ASX 200 valuation": get_asx200_sector_valuation_snapshot_preset,
    "S&P 500 valuation": get_sp500_sector_valuation_snapshot_preset,
    "MSCI EMU valuation": get_msci_emu_sector_valuation_snapshot_preset,
    "MSCI Japan valuation": get_msci_japan_sector_valuation_snapshot_preset,
}


def get_chart_config(chart_key: str, df):
    """
    获取图表配置
    
    Args:
        chart_key: 图表键名
        df: DataFrame
    
    Returns:
        图表配置字典
    """
    if chart_key in CHART_PRESET_FUNCTIONS:
        return CHART_PRESET_FUNCTIONS[chart_key](df)
    else:
        raise ValueError(f"未找到图表配置: {chart_key}")
