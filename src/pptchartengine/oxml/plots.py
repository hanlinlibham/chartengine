"""
绘图区 (Plot) XML 操作模块

负责创建不同类型的图表绘图区 (<c:barChart>, <c:lineChart> 等)。
"""

from lxml import etree
from typing import Literal

from ..oxml_ns import NAMESPACES

ChartType = Literal['bar', 'column', 'line', 'area', 'scatter', 'bubble']
ChartGrouping = Literal['clustered', 'standard', 'stacked', 'percent_stacked']


def create_plot_element(
    plotArea,
    chart_type: ChartType,
    cat_ax_id: int,
    val_ax_id: int,
    order_index: int = 0,
    grouping: ChartGrouping | None = None,
):
    """
    创建图表绘图区元素
    
    Args:
        plotArea: 父绘图区元素
        chart_type: 图表类型 ('bar', 'line', 'area', 'scatter', 'bubble')
        cat_ax_id: 分类轴 ID
        val_ax_id: 值轴 ID
        order_index: 绘图顺序索引（0=最底层，越大越在上层）
        
    Returns:
        创建的绘图元素 (lxml Element)
        
    Raises:
        ValueError: 如果图表类型不支持
        
    Notes:
        - 每个绘图元素会自动关联指定的坐标轴
        - 调用方需要自己添加系列 (<c:ser>)
        - order_index 决定图表的堆叠顺序
    """
    chart_type = chart_type.lower()
    
    if chart_type in ('bar', 'column'):
        return _create_bar_plot(plotArea, cat_ax_id, val_ax_id, order_index, grouping or 'clustered')
    elif chart_type == 'line':
        return _create_line_plot(plotArea, cat_ax_id, val_ax_id, order_index, grouping or 'standard')
    elif chart_type == 'area':
        return _create_area_plot(plotArea, cat_ax_id, val_ax_id, order_index, grouping or 'standard')
    elif chart_type == 'scatter':
        return _create_scatter_plot(plotArea, cat_ax_id, val_ax_id, order_index)
    elif chart_type == 'bubble':
        return _create_bubble_plot(plotArea, cat_ax_id, val_ax_id, order_index)
    else:
        raise ValueError(f"不支持的图表类型: {chart_type}")


def _normalize_grouping(chart_type: str, grouping: str) -> str:
    mapping = {
        "clustered": "clustered",
        "standard": "standard",
        "stacked": "stacked",
        "percent_stacked": "percentStacked",
        "percentStacked": "percentStacked",
    }
    normalized = mapping.get(grouping, grouping)

    if chart_type in ("bar", "column"):
        return normalized if normalized in {"clustered", "stacked", "percentStacked"} else "clustered"
    if chart_type == "area":
        return normalized if normalized in {"standard", "stacked", "percentStacked"} else "standard"
    if chart_type == "line":
        return normalized if normalized in {"standard", "stacked", "percentStacked"} else "standard"
    return normalized


def _create_bar_plot(plotArea, cat_ax_id: int, val_ax_id: int, order_index: int, grouping_value: str):
    """创建柱状图元素"""
    barChart = etree.SubElement(plotArea, f"{{{NAMESPACES['c']}}}barChart")
    
    # barDir: 柱状图方向 ('col' = 垂直柱状, 'bar' = 水平条形)
    barDir = etree.SubElement(barChart, f"{{{NAMESPACES['c']}}}barDir")
    barDir.set('val', 'col')
    
    # grouping: 分组方式 ('clustered' = 簇状, 'stacked' = 堆叠)
    grouping = etree.SubElement(barChart, f"{{{NAMESPACES['c']}}}grouping")
    normalized = _normalize_grouping("bar", grouping_value)
    grouping.set('val', normalized)
    
    # varyColors: 是否每个系列使用不同颜色
    varyColors = etree.SubElement(barChart, f"{{{NAMESPACES['c']}}}varyColors")
    varyColors.set('val', '0')

    if normalized in {"stacked", "percentStacked"}:
        overlap = etree.SubElement(barChart, f"{{{NAMESPACES['c']}}}overlap")
        overlap.set('val', '100')
    
    # ⭐ 绘图顺序（决定堆叠层次，数字越小越在底层）
    # OOXML 规范建议在 varyColors 之后添加
    # 注意：这里不是 <c:ser> 的 order，而是整个 plot 的渲染顺序
    # 但 PowerPoint 实际使用 XML 元素出现的顺序来决定堆叠
    # 所以这个标签主要是语义化，真正的顺序由 XML 元素在 plotArea 中的位置决定
    
    # ⚠️ 注意：不在这里添加轴引用！
    # 轴引用应该在所有系列之后添加，由调用方在添加完系列后调用 add_axis_refs()
    
    return barChart


def _create_line_plot(plotArea, cat_ax_id: int, val_ax_id: int, order_index: int, grouping_value: str):
    """创建折线图元素"""
    lineChart = etree.SubElement(plotArea, f"{{{NAMESPACES['c']}}}lineChart")
    
    # grouping: 分组方式 ('standard' = 标准)
    grouping = etree.SubElement(lineChart, f"{{{NAMESPACES['c']}}}grouping")
    grouping.set('val', _normalize_grouping("line", grouping_value))
    
    # varyColors: 是否每个系列使用不同颜色
    varyColors = etree.SubElement(lineChart, f"{{{NAMESPACES['c']}}}varyColors")
    varyColors.set('val', '0')
    
    # ⚠️ 注意：不在这里添加轴引用！
    # 轴引用应该在所有系列之后添加
    
    return lineChart


def _create_area_plot(plotArea, cat_ax_id: int, val_ax_id: int, order_index: int, grouping_value: str):
    """创建面积图元素"""
    areaChart = etree.SubElement(plotArea, f"{{{NAMESPACES['c']}}}areaChart")
    
    # grouping: 分组方式 ('standard' = 标准)
    grouping = etree.SubElement(areaChart, f"{{{NAMESPACES['c']}}}grouping")
    grouping.set('val', _normalize_grouping("area", grouping_value))
    
    # varyColors
    varyColors = etree.SubElement(areaChart, f"{{{NAMESPACES['c']}}}varyColors")
    varyColors.set('val', '0')
    
    # ⚠️ 注意：不在这里添加轴引用！
    
    return areaChart


def _create_scatter_plot(plotArea, cat_ax_id: int, val_ax_id: int, order_index: int):
    """创建散点图元素"""
    scatterChart = etree.SubElement(plotArea, f"{{{NAMESPACES['c']}}}scatterChart")
    
    # scatterStyle: 散点样式 ('lineMarker' = 带线和标记)
    scatterStyle = etree.SubElement(scatterChart, f"{{{NAMESPACES['c']}}}scatterStyle")
    scatterStyle.set('val', 'lineMarker')
    
    # varyColors
    varyColors = etree.SubElement(scatterChart, f"{{{NAMESPACES['c']}}}varyColors")
    varyColors.set('val', '0')
    
    # ⚠️ 注意：不在这里添加轴引用！
    
    return scatterChart


def _create_bubble_plot(plotArea, cat_ax_id: int, val_ax_id: int, order_index: int):
    """创建气泡图元素"""
    bubbleChart = etree.SubElement(plotArea, f"{{{NAMESPACES['c']}}}bubbleChart")

    varyColors = etree.SubElement(bubbleChart, f"{{{NAMESPACES['c']}}}varyColors")
    varyColors.set('val', '0')

    dLbls = etree.SubElement(bubbleChart, f"{{{NAMESPACES['c']}}}dLbls")
    for tag_name in ("showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize"):
        label_elem = etree.SubElement(dLbls, f"{{{NAMESPACES['c']}}}{tag_name}")
        label_elem.set('val', '0')

    bubbleScale = etree.SubElement(bubbleChart, f"{{{NAMESPACES['c']}}}bubbleScale")
    bubbleScale.set('val', '100')

    showNegBubbles = etree.SubElement(bubbleChart, f"{{{NAMESPACES['c']}}}showNegBubbles")
    showNegBubbles.set('val', '0')

    return bubbleChart


def add_axis_refs(plot_element, cat_ax_id: int, val_ax_id: int):
    """
    为绘图元素添加坐标轴引用（应该在所有系列之后调用）
    
    Args:
        plot_element: 绘图元素 (<c:barChart>, <c:lineChart> 等)
        cat_ax_id: 分类轴 ID
        val_ax_id: 值轴 ID
    """
    axId1 = etree.SubElement(plot_element, f"{{{NAMESPACES['c']}}}axId")
    axId1.set('val', str(cat_ax_id))
    
    axId2 = etree.SubElement(plot_element, f"{{{NAMESPACES['c']}}}axId")
    axId2.set('val', str(val_ax_id))


def add_plot_categories(plot_element, categories: list):
    """
    为绘图元素添加共享的分类数据（在所有系列之前调用）
    
    Args:
        plot_element: 绘图元素 (<c:barChart>, <c:lineChart> 等)
        categories: 分类列表（X轴数据）
        
    Notes:
        - 在 OOXML 规范中，<c:cat> 是图表级别的共享元素
        - 应该在添加任何 <c:ser> 系列之前调用
        - 所有系列共享同一组分类数据
        - ⭐ 自动检测日期类型，使用 numCache（数值缓存）而非 strCache
    """
    from datetime import datetime
    
    # ⭐ 检测是否为日期类型
    is_date_data = False
    if categories and isinstance(categories[0], (datetime, float)):
        # datetime 对象或浮点数（Excel 日期序列号）
        is_date_data = True
    
    cat = etree.SubElement(plot_element, f"{{{NAMESPACES['c']}}}cat")
    
    if is_date_data:
        # ⭐ 新方案：将日期格式化为字符串，使用 strRef + strCache
        # 这样 PowerPoint 就会将其作为文本标签显示，不会出现 1900 年问题
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
                # 假设是 Excel 日期序列号，转换为日期字符串
                base_date = datetime(1899, 12, 30)
                from datetime import timedelta
                actual_date = base_date + timedelta(days=cat_value)
                v.text = actual_date.strftime('%Y/%m')
            else:
                v.text = str(cat_value)
    else:
        # ⭐ 使用 strRef + strCache（普通分类轴）
        strRef = etree.SubElement(cat, f"{{{NAMESPACES['c']}}}strRef")
        
        # f (公式引用)
        f_elem = etree.SubElement(strRef, f"{{{NAMESPACES['c']}}}f")
        f_elem.text = f"Sheet1!$A$2:$A${len(categories) + 1}"
        
        # strCache (字符串缓存)
        strCache = etree.SubElement(strRef, f"{{{NAMESPACES['c']}}}strCache")
        ptCount = etree.SubElement(strCache, f"{{{NAMESPACES['c']}}}ptCount")
        ptCount.set('val', str(len(categories)))
        
        # 添加每个分类点
        for i, cat_value in enumerate(categories):
            pt = etree.SubElement(strCache, f"{{{NAMESPACES['c']}}}pt")
            pt.set('idx', str(i))
            v = etree.SubElement(pt, f"{{{NAMESPACES['c']}}}v")
            v.text = str(cat_value)
