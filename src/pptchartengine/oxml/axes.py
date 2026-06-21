"""
坐标轴 XML 操作模块

负责创建、提取和优化坐标轴元素。
"""

from lxml import etree
from typing import Tuple

from .._log import debug_print as print
from ..oxml_ns import NAMESPACES


def extract_axis_ids(plotArea) -> Tuple[int, int]:
    """
    提取现有坐标轴的 ID
    
    Args:
        plotArea: 绘图区元素 (lxml Element)
        
    Returns:
        (cat_ax_id, val_ax_id) 元组
        
    Raises:
        ValueError: 如果无法找到坐标轴 ID
    """
    # python-pptx 的 BaseOxmlElement.xpath() 已经注册了命名空间
    cat_ax_elements = plotArea.xpath('.//c:catAx/c:axId')
    val_ax_elements = plotArea.xpath('.//c:valAx/c:axId')
    
    if cat_ax_elements and val_ax_elements:
        cat_ax_id = int(cat_ax_elements[0].get('val'))
        val_ax_id = int(val_ax_elements[0].get('val'))
        return cat_ax_id, val_ax_id

    if len(val_ax_elements) >= 2:
        x_axis_id = int(val_ax_elements[0].get('val'))
        y_axis_id = int(val_ax_elements[1].get('val'))
        return x_axis_id, y_axis_id

    raise ValueError("无法找到现有坐标轴 ID")


def create_value_axis(
    plotArea,
    ax_id: int,
    cross_ax_id: int,
    position: str = 'r',
    tick_label_position: str = 'high',
    crosses_at: str = 'max',
) -> int:
    """
    创建一个新的值轴 (Y轴)
    
    Args:
        plotArea: 绘图区元素 (lxml Element)
        ax_id: 新轴的 ID
        cross_ax_id: 交叉轴的 ID（通常是分类轴）
        position: 轴位置 ('l'=左, 'r'=右, 't'=顶, 'b'=底)
        tick_label_position: 标签位置 ('low'=左/底, 'high'=右/顶, 'nextTo'=靠近轴)
        crosses_at: 交叉位置 ('min'=最小值/左边, 'max'=最大值/右边)
        
    Returns:
        创建的轴 ID
        
    Notes:
        - 严格按照 OOXML 规范 (ISO/IEC 29500-1:2016) 的元素顺序
        - 'low' 和 'high' 用于双轴图，确保标签不重叠
        - crosses_at='min' 让轴线在图表左边，'max' 让轴线在图表右边
    """
    # 创建值轴元素
    valAx = etree.SubElement(plotArea, f"{{{NAMESPACES['c']}}}valAx")
    
    # ⭐ 严格按照 OOXML 规范顺序添加子元素
    
    # 1. axId (轴 ID) - 必需
    axId_elem = etree.SubElement(valAx, f"{{{NAMESPACES['c']}}}axId")
    axId_elem.set('val', str(ax_id))
    
    # 2. scaling (缩放) - 必需
    scaling = etree.SubElement(valAx, f"{{{NAMESPACES['c']}}}scaling")
    orientation = etree.SubElement(scaling, f"{{{NAMESPACES['c']}}}orientation")
    orientation.set('val', 'minMax')
    
    # 3. delete (是否隐藏) - 必需
    delete = etree.SubElement(valAx, f"{{{NAMESPACES['c']}}}delete")
    delete.set('val', '0')
    
    # 4. axPos (轴位置) - 必需
    axPos = etree.SubElement(valAx, f"{{{NAMESPACES['c']}}}axPos")
    axPos.set('val', position)
    
    # 5. majorGridlines (主网格线) - 可选
    # 次轴通常不显示网格线，避免与主轴重叠
    # 如果需要，调用方可以手动添加
    
    # 6. numFmt (数字格式) - 可选
    numFmt = etree.SubElement(valAx, f"{{{NAMESPACES['c']}}}numFmt")
    numFmt.set('formatCode', 'General')
    numFmt.set('sourceLinked', '0')
    
    # 7. majorTickMark (主刻度线) - 可选
    majorTickMark = etree.SubElement(valAx, f"{{{NAMESPACES['c']}}}majorTickMark")
    majorTickMark.set('val', 'out')
    
    # 8. minorTickMark (次刻度线) - 可选
    minorTickMark = etree.SubElement(valAx, f"{{{NAMESPACES['c']}}}minorTickMark")
    minorTickMark.set('val', 'none')
    
    # 9. tickLblPos (标签位置) - 可选
    tickLblPos = etree.SubElement(valAx, f"{{{NAMESPACES['c']}}}tickLblPos")
    tickLblPos.set('val', tick_label_position)
    
    # 10. crossAx (交叉轴 ID) - 必需
    crossAx = etree.SubElement(valAx, f"{{{NAMESPACES['c']}}}crossAx")
    crossAx.set('val', str(cross_ax_id))
    
    # 11. crosses (交叉方式) - 可选
    # ⭐ 关键修复：根据 crosses_at 参数决定交叉位置
    # 'min' = 在最小值（左边）交叉，'max' = 在最大值（右边）交叉
    crosses = etree.SubElement(valAx, f"{{{NAMESPACES['c']}}}crosses")
    crosses.set('val', crosses_at)
    
    # 12. crossBetween (交叉位置) - 可选
    crossBetween = etree.SubElement(valAx, f"{{{NAMESPACES['c']}}}crossBetween")
    crossBetween.set('val', 'between')
    
    return ax_id


def optimize_axis_labels(
    plotArea,
    ax_id: int,
    tick_label_position: str = 'low',
    crosses_at: str = 'min',
    remove_gridlines: bool = True,
):
    """
    优化现有轴的标签位置和交叉位置
    
    Args:
        plotArea: 绘图区元素
        ax_id: 要优化的轴 ID
        tick_label_position: 标签位置 ('low'=左/底, 'high'=右/顶)
        crosses_at: 交叉位置 ('min'=最小值/左边, 'max'=最大值/右边)
        remove_gridlines: 是否移除网格线（默认移除）
        
    Notes:
        主要用于优化主值轴，使其与次值轴协调
        crosses_at='min' 让主轴线在图表左边，'max' 让次轴线在图表右边
    """
    # 查找指定的值轴
    val_ax_elements = plotArea.xpath(f'.//c:valAx[c:axId[@val="{ax_id}"]]')
    if not val_ax_elements:
        return  # 轴不存在，跳过
    
    val_ax = val_ax_elements[0]
    
    # ⭐ 设置 crosses 位置
    crosses_elements = val_ax.xpath('./c:crosses')
    if crosses_elements:
        crosses_elements[0].set('val', crosses_at)
    
    # 设置标签位置
    tickLblPos_elements = val_ax.xpath('./c:tickLblPos')
    if tickLblPos_elements:
        tickLblPos_elements[0].set('val', tick_label_position)
    else:
        # 如果不存在，创建一个
        tickLblPos = etree.Element(f"{{{NAMESPACES['c']}}}tickLblPos")
        tickLblPos.set('val', tick_label_position)
        # 插入到 crossAx 之前（保持正确顺序）
        cross_ax_elements = val_ax.xpath('./c:crossAx')
        if cross_ax_elements:
            cross_ax_elements[0].addprevious(tickLblPos)
    
    # ⭐ 移除网格线（取消内部横框）
    if remove_gridlines:
        gridlines = val_ax.xpath('./c:majorGridlines')
        for gridline in gridlines:
            val_ax.remove(gridline)
            print(f"  → 已移除主值轴网格线")


def optimize_category_axis(
    plotArea,
    cat_ax_id: int,
    remove_tick_marks: bool = True,
):
    """
    优化分类轴（X轴）的显示
    
    Args:
        plotArea: 绘图区元素
        cat_ax_id: 分类轴 ID
        remove_tick_marks: 是否移除主刻度线（默认移除，即取消日期间的小竖线）
        
    Notes:
        用于清理分类轴的视觉元素，让图表更简洁
    """
    # 查找分类轴
    cat_ax_elements = plotArea.xpath(f'.//c:catAx[c:axId[@val="{cat_ax_id}"]]')
    if not cat_ax_elements:
        return  # 轴不存在，跳过
    
    cat_ax = cat_ax_elements[0]
    
    # ⭐ 移除或设置主刻度线为 'none'（取消日期间的小竖线）
    if remove_tick_marks:
        majorTickMark_elements = cat_ax.xpath('./c:majorTickMark')
        if majorTickMark_elements:
            # 修改为 'none' 而不是删除元素
            majorTickMark_elements[0].set('val', 'none')
            print(f"  → 已移除分类轴主刻度线（日期间的小竖线）")
        
        # 同时也设置次刻度线为 'none'
        minorTickMark_elements = cat_ax.xpath('./c:minorTickMark')
        if minorTickMark_elements:
            minorTickMark_elements[0].set('val', 'none')
