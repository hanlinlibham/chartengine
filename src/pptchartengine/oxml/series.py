"""
系列 (Series) XML 操作模块

负责创建图表系列 (<c:ser>) 和数据链接。
"""

from lxml import etree
from typing import Dict, List, Optional
import pandas as pd

from ..oxml_ns import NAMESPACES


def get_excel_col_name(col_idx: int) -> str:
    """
    将 0-based 索引转换为 Excel 列名 (A, B, ..., Z, AA, AB, ...)
    
    Args:
        col_idx: 0-based 列索引 (0='A', 1='B', ..., 25='Z', 26='AA', ...)
        
    Returns:
        Excel 列名字符串
        
    Examples:
        >>> get_excel_col_name(0)   # 'A'
        >>> get_excel_col_name(25)  # 'Z'
        >>> get_excel_col_name(26)  # 'AA'
        >>> get_excel_col_name(27)  # 'AB'
        >>> get_excel_col_name(701) # 'ZZ'
        >>> get_excel_col_name(702) # 'AAA'
    """
    col_name = ""
    while col_idx >= 0:
        col_name = chr(col_idx % 26 + 65) + col_name
        col_idx = col_idx // 26 - 1
    return col_name


def add_series_to_plot(
    plot_element,
    chart_type: str,
    series_cfg: Dict,
    series_idx: int,
    df: pd.DataFrame,
    categories_col: str,
    style_config=None,
):
    """
    向绘图元素添加一个系列
    
    Args:
        plot_element: 绘图元素 (<c:barChart>, <c:lineChart> 等)
        chart_type: 图表类型 ('bar', 'line', 'area')
        series_cfg: 系列配置 {"key": "col_name", "name": "Series Name"}
        series_idx: 系列索引 (0-based，用于 Excel 列引用)
        df: 数据源 DataFrame
        categories_col: 分类列名
        style_config: 样式配置对象（可选）
        
    Returns:
        创建的系列元素（用于后续样式应用）
        
    Notes:
        - series_idx 用于生成 Excel 列引用
        - 假设分类在 A 列，数据从 B 列开始
        - series_idx=0 对应 B 列，series_idx=1 对应 C 列，以此类推
    """
    chart_type = chart_type.lower()
    
    if chart_type in ('bar', 'column'):
        return _add_bar_series(plot_element, series_cfg, series_idx, df, categories_col, style_config)
    elif chart_type == 'line':
        return _add_line_series(plot_element, series_cfg, series_idx, df, categories_col, style_config)
    elif chart_type == 'area':
        return _add_area_series(plot_element, series_cfg, series_idx, df, categories_col, style_config)
    elif chart_type == 'scatter':
        return _add_scatter_series(plot_element, series_cfg, series_idx, df, categories_col, style_config)
    elif chart_type == 'bubble':
        return _add_bubble_series(plot_element, series_cfg, series_idx, df, categories_col, style_config)
    else:
        raise ValueError(f"不支持的图表类型: {chart_type}")


def _add_bar_series(plot_element, series_cfg: Dict, series_idx: int, df: pd.DataFrame, categories_col: str, style_config=None):
    """添加柱状图系列"""
    values = df[series_cfg["key"]].tolist()
    categories = df[categories_col].tolist()
    series_name = series_cfg["name"]
    
    # 创建系列元素
    ser = etree.SubElement(plot_element, f"{{{NAMESPACES['c']}}}ser")
    
    # idx 和 order
    _add_series_index(ser, series_idx)
    
    # tx (系列名称)
    _add_series_title(ser, series_name, series_idx)
    
    # ⭐ 应用样式（如果提供）
    if style_config is not None:
        style_config.apply_to_series(ser, series_idx)
    
    # ⭐ cat (分类数据) - 每个系列都必须有！
    _add_series_categories(ser, categories, series_idx)
    
    # val (数值)
    _add_series_values(ser, values, series_idx)
    
    # 柱状图特有：不添加 marker 和 smooth
    
    return ser


def _add_line_series(plot_element, series_cfg: Dict, series_idx: int, df: pd.DataFrame, categories_col: str, style_config=None):
    """添加折线图系列"""
    values = df[series_cfg["key"]].tolist()
    categories = df[categories_col].tolist()
    series_name = series_cfg["name"]
    
    # 创建系列元素
    ser = etree.SubElement(plot_element, f"{{{NAMESPACES['c']}}}ser")
    
    # idx 和 order
    _add_series_index(ser, series_idx)
    
    # tx (系列名称)
    _add_series_title(ser, series_name, series_idx)
    
    # ⭐ 应用样式（如果提供）- 会自动处理 marker
    if style_config is not None:
        style_config.apply_to_series(ser, series_idx)
    else:
        # 默认：圆形标记点
        marker = etree.SubElement(ser, f"{{{NAMESPACES['c']}}}marker")
        symbol = etree.SubElement(marker, f"{{{NAMESPACES['c']}}}symbol")
        symbol.set('val', 'circle')
    
    # ⭐ cat (分类数据) - 每个系列都必须有！
    _add_series_categories(ser, categories, series_idx)
    
    # val (数值)
    _add_series_values(ser, values, series_idx)
    
    # smooth (平滑)
    smooth = etree.SubElement(ser, f"{{{NAMESPACES['c']}}}smooth")
    smooth.set('val', '0')
    
    return ser


def _add_bubble_series(
    plot_element,
    series_cfg: Dict,
    series_idx: int,
    df: pd.DataFrame,
    categories_col: str,
    style_config=None
):
    """添加气泡图系列"""
    y_values = df[series_cfg["key"]].tolist()
    size_values = df[series_cfg["size_key"]].tolist()
    series_name = series_cfg["name"]
    x_key = series_cfg.get("x_key", categories_col)
    x_values = df[x_key].tolist()

    ser = etree.SubElement(plot_element, f"{{{NAMESPACES['c']}}}ser")
    _add_series_index(ser, series_idx)
    _add_xy_series_title(ser, series_name, series_idx, len(x_values))

    invert_if_negative = etree.SubElement(ser, f"{{{NAMESPACES['c']}}}invertIfNegative")
    invert_if_negative.set('val', '0')

    _add_scatter_x_values(ser, x_values, series_idx, x_key)
    _add_scatter_y_values(ser, y_values, series_idx)
    _add_bubble_size_values(ser, size_values, series_idx)

    bubble3D = etree.SubElement(ser, f"{{{NAMESPACES['c']}}}bubble3D")
    bubble3D.set('val', '0')

    return ser


def _add_area_series(plot_element, series_cfg: Dict, series_idx: int, df: pd.DataFrame, categories_col: str, style_config=None):
    """添加面积图系列"""
    values = df[series_cfg["key"]].tolist()
    categories = df[categories_col].tolist()
    series_name = series_cfg["name"]
    
    # 创建系列元素
    ser = etree.SubElement(plot_element, f"{{{NAMESPACES['c']}}}ser")
    
    # idx 和 order
    _add_series_index(ser, series_idx)
    
    # tx (系列名称)
    _add_series_title(ser, series_name, series_idx)
    
    # ⭐ 应用样式（如果提供）
    if style_config is not None:
        style_config.apply_to_series(ser, series_idx)
    
    # ⭐ cat (分类数据) - 每个系列都必须有！
    _add_series_categories(ser, categories, series_idx)
    
    # val (数值)
    _add_series_values(ser, values, series_idx)
    
    return ser


def _add_scatter_series(
    plot_element,
    series_cfg: Dict,
    series_idx: int,
    df: pd.DataFrame,
    categories_col: str,
    style_config=None
):
    """
    添加散点图系列
    
    Note: 
        - 散点图需要 X 和 Y 两个数值序列
        - 与其他图表类型不同，散点图使用 <c:xVal> 和 <c:yVal>
        - 如果 series_cfg 中指定了 'x_key'，使用它作为 X 轴数据
        - 否则使用 categories_col 作为 X 轴数据（必须是数值型）
        
    Args:
        plot_element: 散点图元素
        series_cfg: 系列配置，必须包含 'key' (Y轴数据)，可选 'x_key' (X轴数据)
        series_idx: 系列索引
        df: 数据 DataFrame
        categories_col: 默认的 X 轴列名
        style_config: 样式配置对象（可选）
    """
    # Y 轴数据（必需）
    y_values = df[series_cfg["key"]].tolist()
    series_name = series_cfg["name"]
    
    # X 轴数据（可选，默认使用 categories_col）
    x_key = series_cfg.get("x_key", categories_col)
    x_values = df[x_key].tolist()
    
    # 创建系列元素
    ser = etree.SubElement(plot_element, f"{{{NAMESPACES['c']}}}ser")
    
    # idx 和 order
    _add_series_index(ser, series_idx)
    
    # tx (系列名称)
    _add_series_title(ser, series_name, series_idx)
    
    # ⭐ 应用样式（如果提供）- 会自动处理 marker
    if style_config is not None:
        style_config.apply_to_series(ser, series_idx)
    else:
        # 默认：圆形标记
        marker = etree.SubElement(ser, f"{{{NAMESPACES['c']}}}marker")
        symbol = etree.SubElement(marker, f"{{{NAMESPACES['c']}}}symbol")
        symbol.set('val', 'circle')
    
    # ⭐ xVal (X轴数值) - 散点图特有
    _add_scatter_x_values(ser, x_values, series_idx, x_key)
    
    # ⭐ yVal (Y轴数值) - 散点图特有
    _add_scatter_y_values(ser, y_values, series_idx)
    
    # smooth (平滑)
    smooth = etree.SubElement(ser, f"{{{NAMESPACES['c']}}}smooth")
    smooth.set('val', '0')
    
    return ser


def _add_scatter_x_values(ser, x_values: list, series_idx: int, x_col_name: str):
    """
    为散点图添加 X 轴数值数据
    
    Note: 散点图使用 <c:xVal> 而不是 <c:cat>
    """
    xVal = etree.SubElement(ser, f"{{{NAMESPACES['c']}}}xVal")
    numRef = etree.SubElement(xVal, f"{{{NAMESPACES['c']}}}numRef")
    
    # f (公式引用)
    f_elem = etree.SubElement(numRef, f"{{{NAMESPACES['c']}}}f")
    start_row = _get_xy_series_first_data_row(series_idx, len(x_values))
    end_row = start_row + len(x_values) - 1
    data_range = f"Sheet1!$A${start_row}:$A${end_row}"
    f_elem.text = data_range
    
    # numCache (数值缓存)
    numCache = etree.SubElement(numRef, f"{{{NAMESPACES['c']}}}numCache")
    formatCode = etree.SubElement(numCache, f"{{{NAMESPACES['c']}}}formatCode")
    formatCode.text = 'General'
    ptCount = etree.SubElement(numCache, f"{{{NAMESPACES['c']}}}ptCount")
    ptCount.set('val', str(len(x_values)))
    
    # 添加每个数据点
    for i, x_value in enumerate(x_values):
        pt = etree.SubElement(numCache, f"{{{NAMESPACES['c']}}}pt")
        pt.set('idx', str(i))
        v = etree.SubElement(pt, f"{{{NAMESPACES['c']}}}v")
        v.text = str(x_value)


def _add_scatter_y_values(ser, y_values: list, series_idx: int):
    """
    为散点图添加 Y 轴数值数据
    
    Note: 散点图使用 <c:yVal> 而不是 <c:val>
    """
    yVal = etree.SubElement(ser, f"{{{NAMESPACES['c']}}}yVal")
    numRef = etree.SubElement(yVal, f"{{{NAMESPACES['c']}}}numRef")
    
    # f (公式引用)
    f_elem = etree.SubElement(numRef, f"{{{NAMESPACES['c']}}}f")
    start_row = _get_xy_series_first_data_row(series_idx, len(y_values))
    end_row = start_row + len(y_values) - 1
    data_range = f"Sheet1!$B${start_row}:$B${end_row}"
    f_elem.text = data_range
    
    # numCache (数值缓存)
    numCache = etree.SubElement(numRef, f"{{{NAMESPACES['c']}}}numCache")
    formatCode = etree.SubElement(numCache, f"{{{NAMESPACES['c']}}}formatCode")
    formatCode.text = 'General'
    ptCount = etree.SubElement(numCache, f"{{{NAMESPACES['c']}}}ptCount")
    ptCount.set('val', str(len(y_values)))
    
    # 添加每个数据点
    for i, y_value in enumerate(y_values):
        pt = etree.SubElement(numCache, f"{{{NAMESPACES['c']}}}pt")
        pt.set('idx', str(i))
        v = etree.SubElement(pt, f"{{{NAMESPACES['c']}}}v")
        v.text = str(y_value)


def _add_bubble_size_values(ser, size_values: list, series_idx: int):
    """为气泡图添加 size 数据"""
    bubbleSize = etree.SubElement(ser, f"{{{NAMESPACES['c']}}}bubbleSize")
    numRef = etree.SubElement(bubbleSize, f"{{{NAMESPACES['c']}}}numRef")

    f_elem = etree.SubElement(numRef, f"{{{NAMESPACES['c']}}}f")
    start_row = _get_xy_series_first_data_row(series_idx, len(size_values))
    end_row = start_row + len(size_values) - 1
    f_elem.text = f"Sheet1!$C${start_row}:$C${end_row}"

    numCache = etree.SubElement(numRef, f"{{{NAMESPACES['c']}}}numCache")
    formatCode = etree.SubElement(numCache, f"{{{NAMESPACES['c']}}}formatCode")
    formatCode.text = 'General'
    ptCount = etree.SubElement(numCache, f"{{{NAMESPACES['c']}}}ptCount")
    ptCount.set('val', str(len(size_values)))

    for i, size_value in enumerate(size_values):
        pt = etree.SubElement(numCache, f"{{{NAMESPACES['c']}}}pt")
        pt.set('idx', str(i))
        v = etree.SubElement(pt, f"{{{NAMESPACES['c']}}}v")
        v.text = str(size_value)


# ============================================================================
# 辅助函数
# ============================================================================

def _add_series_index(ser, series_idx: int):
    """添加系列索引和顺序"""
    idx = etree.SubElement(ser, f"{{{NAMESPACES['c']}}}idx")
    idx.set('val', str(series_idx))
    
    order = etree.SubElement(ser, f"{{{NAMESPACES['c']}}}order")
    order.set('val', str(series_idx))


def _add_series_categories(ser, categories: list, series_idx: int):
    """
    为系列添加分类数据 (cat)
    
    ⭐ 关键修复：每个系列都必须有自己的 cat 元素，PowerPoint 才能正确显示标签
    """
    from datetime import datetime
    
    # 检测是否为日期类型
    is_date_data = False
    if categories and isinstance(categories[0], (datetime, float)):
        is_date_data = True
    
    cat = etree.SubElement(ser, f"{{{NAMESPACES['c']}}}cat")
    
    if is_date_data:
        # 使用 strRef + strCache（格式化为字符串）
        strRef = etree.SubElement(cat, f"{{{NAMESPACES['c']}}}strRef")
        
        # f (公式引用)
        f_elem = etree.SubElement(strRef, f"{{{NAMESPACES['c']}}}f")
        f_elem.text = f"Sheet1!$A$2:$A${len(categories) + 1}"
        
        # strCache (字符串缓存)
        strCache = etree.SubElement(strRef, f"{{{NAMESPACES['c']}}}strCache")
        ptCount = etree.SubElement(strCache, f"{{{NAMESPACES['c']}}}ptCount")
        ptCount.set('val', str(len(categories)))
        
        # 添加每个分类点（格式化为字符串）
        for i, cat_value in enumerate(categories):
            pt = etree.SubElement(strCache, f"{{{NAMESPACES['c']}}}pt")
            pt.set('idx', str(i))
            v = etree.SubElement(pt, f"{{{NAMESPACES['c']}}}v")
            
            if isinstance(cat_value, datetime):
                # 格式化为 "yyyy/mm"（年份/月份）
                v.text = cat_value.strftime('%Y/%m')
            elif isinstance(cat_value, float):
                # Excel 日期序列号，转换为日期字符串
                base_date = datetime(1899, 12, 30)
                from datetime import timedelta
                actual_date = base_date + timedelta(days=cat_value)
                v.text = actual_date.strftime('%Y/%m')
            else:
                v.text = str(cat_value)
    else:
        # 普通字符串分类
        strRef = etree.SubElement(cat, f"{{{NAMESPACES['c']}}}strRef")
        
        # f (公式引用)
        f_elem = etree.SubElement(strRef, f"{{{NAMESPACES['c']}}}f")
        f_elem.text = f"Sheet1!$A$2:$A${len(categories) + 1}"
        
        # strCache (字符串缓存)
        strCache = etree.SubElement(strRef, f"{{{NAMESPACES['c']}}}strCache")
        ptCount = etree.SubElement(strCache, f"{{{NAMESPACES['c']}}}ptCount")
        ptCount.set('val', str(len(categories)))
        
        # 添加每个分类点
        for i, cat in enumerate(categories):
            pt = etree.SubElement(strCache, f"{{{NAMESPACES['c']}}}pt")
            pt.set('idx', str(i))
            v = etree.SubElement(pt, f"{{{NAMESPACES['c']}}}v")
            v.text = str(cat)


def _add_series_title(ser, series_name: str, series_idx: int):
    """添加系列标题 (tx)"""
    tx = etree.SubElement(ser, f"{{{NAMESPACES['c']}}}tx")
    strRef = etree.SubElement(tx, f"{{{NAMESPACES['c']}}}strRef")
    f_elem = etree.SubElement(strRef, f"{{{NAMESPACES['c']}}}f")
    
    # 使用修复后的 Excel 列名生成
    # series_idx 从 0 开始，分类在 A 列(索引0)，数据从 B 列(索引1)开始
    col_letter = get_excel_col_name(series_idx + 1)
    f_elem.text = f"Sheet1!${col_letter}$1"
    
    # strCache (缓存的字符串值)
    strCache = etree.SubElement(strRef, f"{{{NAMESPACES['c']}}}strCache")
    ptCount = etree.SubElement(strCache, f"{{{NAMESPACES['c']}}}ptCount")
    ptCount.set('val', '1')
    pt = etree.SubElement(strCache, f"{{{NAMESPACES['c']}}}pt")
    pt.set('idx', '0')
    v = etree.SubElement(pt, f"{{{NAMESPACES['c']}}}v")
    v.text = series_name


def _add_xy_series_title(ser, series_name: str, series_idx: int, point_count: int):
    """添加 XY 图表系列标题 (tx)"""
    tx = etree.SubElement(ser, f"{{{NAMESPACES['c']}}}tx")
    strRef = etree.SubElement(tx, f"{{{NAMESPACES['c']}}}strRef")
    f_elem = etree.SubElement(strRef, f"{{{NAMESPACES['c']}}}f")
    f_elem.text = f"Sheet1!$B${_get_xy_series_title_row(series_idx, point_count)}"

    strCache = etree.SubElement(strRef, f"{{{NAMESPACES['c']}}}strCache")
    ptCount = etree.SubElement(strCache, f"{{{NAMESPACES['c']}}}ptCount")
    ptCount.set('val', '1')
    pt = etree.SubElement(strCache, f"{{{NAMESPACES['c']}}}pt")
    pt.set('idx', '0')
    v = etree.SubElement(pt, f"{{{NAMESPACES['c']}}}v")
    v.text = series_name


def _add_series_values(ser, values: List, series_idx: int):
    """添加系列数值 (val)"""
    val = etree.SubElement(ser, f"{{{NAMESPACES['c']}}}val")
    numRef = etree.SubElement(val, f"{{{NAMESPACES['c']}}}numRef")
    f_elem = etree.SubElement(numRef, f"{{{NAMESPACES['c']}}}f")
    
    # 数据范围引用
    col_letter = get_excel_col_name(series_idx + 1)
    data_range = f"Sheet1!${col_letter}$2:${col_letter}${len(values) + 1}"
    f_elem.text = data_range
    
    # numCache (缓存的数值)
    numCache = etree.SubElement(numRef, f"{{{NAMESPACES['c']}}}numCache")
    formatCode = etree.SubElement(numCache, f"{{{NAMESPACES['c']}}}formatCode")
    formatCode.text = 'General'
    ptCount = etree.SubElement(numCache, f"{{{NAMESPACES['c']}}}ptCount")
    ptCount.set('val', str(len(values)))
    
    # 添加每个数据点
    for i, value in enumerate(values):
        pt = etree.SubElement(numCache, f"{{{NAMESPACES['c']}}}pt")
        pt.set('idx', str(i))
        v = etree.SubElement(pt, f"{{{NAMESPACES['c']}}}v")
        v.text = str(value)


def _get_xy_series_title_row(series_idx: int, point_count: int) -> int:
    return 1 + series_idx * (point_count + 2)


def _get_xy_series_first_data_row(series_idx: int, point_count: int) -> int:
    return _get_xy_series_title_row(series_idx, point_count) + 1
