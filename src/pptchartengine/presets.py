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
