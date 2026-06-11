"""
图表样式配置模块

定义默认的颜色方案、线型、标记点等样式。
"""

from typing import Dict, List, Tuple
from lxml import etree

from .oxml_ns import NAMESPACES

# ============================================================================
# 颜色方案定义（RGB 格式）
# ============================================================================

# 深色系列（主色）
DARK_RED = "C00000"      # 深红色
DARK_GRAY = "595959"     # 深灰色
DARK_BLUE = "0070C0"     # 深蓝色
DARK_ORANGE = "ED7D31"   # 深橙色

# 浅色系列（辅助色）
LIGHT_RED = "FF9999"     # 浅红色
LIGHT_GRAY = "BFBFBF"    # 浅灰色
LIGHT_BLUE = "9DC3E6"    # 浅蓝色
LIGHT_ORANGE = "F4B183"  # 浅橙色

# ⭐ aim00 风格颜色（对比色方案：浅灰柱状图 + 深红折线）
AIM00_LIGHT_GRAY = "C0C0C0"  # 浅灰色 - 用于柱状图/面积图（第一系列）
AIM00_DARK_RED = "C00000"    # 深红色 - 用于折线图（第二系列）
AIM00_DARK_BLUE = "305496"   # 深蓝色 - 备用
AIM00_GRAY = "808080"        # 中灰色 - 备用

# 默认颜色序列（深色 + 浅色交替）
DEFAULT_COLOR_SEQUENCE = [
    DARK_RED,
    DARK_BLUE,
    DARK_ORANGE,
    DARK_GRAY,
    LIGHT_RED,
    LIGHT_BLUE,
    LIGHT_ORANGE,
    LIGHT_GRAY,
]

# ── 主题配套图表系列色（每个主题一套 8 色） ──
COLOR_SCHEMES = {
    # 咨询报告风格：深海军蓝主色 + 亮蓝/青绿辅助 + 灰阶，深色优先保证单系列可读
    "advisory":      ["1F3864", "2E9BD6", "00A398", "8C8C8C", "9DC3E6", "C9A84C", "5C7A93", "404040"],
    # 市场指南（GTM，General Theme for Markets）风格：灰+青为主对，橙色留给"合计/净值"线，
    # 海军蓝/紫/橄榄绿做分项（市场指南类报告堆叠贡献图的惯用色序）
    "gtm":           ["595959", "29ABE2", "F5821F", "1F3864", "7B5EA7", "6BA43A", "00838F", "A6A6A6"],
    "midnight":      ["CADCFC", "1E2761", "E8B931", "4A6FA5", "9BB5D6", "B08C28", "8896AB", "484E5C"],
    "charcoal":      ["A8B4BE", "36454F", "E8B931", "607D8B", "B0BEC5", "C4A035", "78909C", "4A565E"],
    "jp_finance":    ["A8D5DC", "1B3D6E", "C9A84C", "4A8FB8", "7FBFCF", "A07840", "8FA8C0", "505868"],
    "pension_warm":  ["A8C4D4", "2E5FA3", "D4903F", "7BA7BC", "B0C9B0", "9C7B48", "8A9DB8", "606870"],
    "tech_blue":     ["90CAF9", "1565C0", "00BFA5", "7E57C2", "29B6F6", "FF7043", "90A4AE", "546E7A"],
    "state_red":     ["CDAC60", "8B0000", "B8860B", "4A6FA5", "C07840", "507050", "907878", "585048"],
    "esg_green":     ["A5D6A7", "1A5C2A", "8BC34A", "00838F", "C8E6C9", "795548", "78909C", "607D8B"],
    "dark_pro":      ["00BFFF", "FFD700", "00E676", "7B68EE", "FF6D00", "E040FB", "A0B0C0", "4A5868"],
    "daybreak":      ["A4C2D8", "1D2B3A", "E67E22", "5A9BD5", "95C8D8", "B87333", "7B8794", "495057"],
    "macro_research":["AED6F1", "2C3E50", "3498DB", "95A5A6", "85C1E9", "1ABC9C", "BDC3C7", "566573"],
    # 兜底别名：深色优先（浅色开头会让单系列折线图几乎不可见）
    "default":       ["1F3864", "2E9BD6", "00A398", "8C8C8C", "9DC3E6", "C9A84C", "5C7A93", "404040"],
}

# ============================================================================
# 线型配置
# ============================================================================

# 线宽（EMUs: English Metric Units, 1 pt = 12700 EMUs）
LINE_WIDTH_PT = {
    0.5: 6350,    # 0.5 pt
    0.75: 9525,   # 0.75 pt
    1.0: 12700,   # 1 pt (默认)
    1.5: 19050,   # 1.5 pt
    2.0: 25400,   # 2 pt
    2.25: 28575,  # 2.25 pt
    3.0: 38100,   # 3 pt
}

DEFAULT_LINE_WIDTH = LINE_WIDTH_PT[1.0]  # 1 pt

# ============================================================================
# 标记点配置
# ============================================================================

# 标记点样式
MARKER_STYLES = {
    "none": None,        # 无标记
    "circle": "circle",  # 圆形
    "square": "square",  # 方形
    "diamond": "diamond", # 菱形
    "triangle": "triangle", # 三角形
}

DEFAULT_MARKER_STYLE = "none"  # 默认无标记点
DEFAULT_MARKER_SIZE = 5        # 标记点大小（pt）

# ============================================================================
# 样式应用函数
# ============================================================================

def apply_series_style(
    ser_element,
    color: str = None,
    line_width: int = None,
    marker_style: str = None,
    marker_size: int = None,
):
    """
    为系列元素应用样式
    
    Args:
        ser_element: 系列元素 (<c:ser>)
        color: RGB 颜色（6位十六进制，如 "C00000"）
        line_width: 线宽（EMUs）
        marker_style: 标记点样式（"none", "circle", "square" 等）
        marker_size: 标记点大小（pt）
    """
    # ⭐ 检测图表类型（通过父元素）
    parent = ser_element.getparent()
    parent_tag = parent.tag.split('}')[1] if '}' in parent.tag else parent.tag
    is_area_or_bar = parent_tag in ('barChart', 'areaChart')
    # OOXML schema: <c:marker> is only a valid child of <c:ser> in lineChart,
    # scatterChart, and radarChart series. CT_BarSer / CT_AreaSer / CT_BubbleSer
    # do not allow it, and PowerPoint refuses to open files where it appears.
    marker_allowed = parent_tag in ('lineChart', 'scatterChart', 'radarChart')

    # 查找或创建 spPr (shape properties) 元素
    spPr = ser_element.find('c:spPr', namespaces=NAMESPACES)
    if spPr is None:
        # 在 idx/order/tx 之后插入 spPr
        insert_index = _find_insert_position(ser_element, 'spPr')
        spPr = etree.Element(f"{{{NAMESPACES['c']}}}spPr")
        ser_element.insert(insert_index, spPr)

    # 应用线条样式（传入图表类型信息）
    if line_width is not None or color is not None:
        _apply_line_style(spPr, color, line_width, is_area_or_bar)

    # 应用标记点样式（仅对支持 marker 的图表类型）
    if marker_style is not None and marker_allowed:
        _apply_marker_style(ser_element, marker_style, marker_size, color)


def _apply_line_style(spPr, color: str = None, line_width: int = None, is_area_or_bar: bool = False):
    """
    应用线条样式
    
    Args:
        spPr: shape properties 元素
        color: 颜色
        line_width: 线宽
        is_area_or_bar: 是否为面积图或柱状图（这些图表类型需要设置填充色而非边框色）
    """
    # 清除现有的 ln 与 solidFill（重新构造）
    for tag in ('ln',):
        for el in spPr.findall(f'a:{tag}', namespaces=NAMESPACES):
            spPr.remove(el)
    for fill in spPr.findall('a:solidFill', namespaces=NAMESPACES):
        spPr.remove(fill)

    # ⭐ OOXML CT_ShapeProperties 要求 fill 选项在 <a:ln> 之前。
    # 顺序错了会让 PowerPoint 忽略 fill / 用主题色回退。
    if is_area_or_bar:
        # 柱状图 / 面积图：先 fill 再 ln（无边框）
        if color is not None:
            solidFill = etree.SubElement(spPr, f"{{{NAMESPACES['a']}}}solidFill")
            srgbClr = etree.SubElement(solidFill, f"{{{NAMESPACES['a']}}}srgbClr")
            srgbClr.set('val', color)
        ln = etree.SubElement(spPr, f"{{{NAMESPACES['a']}}}ln")
        etree.SubElement(ln, f"{{{NAMESPACES['a']}}}noFill")
    else:
        # 折线图 / 散点图：颜色在 ln 内（描边色），spPr 不需要顶层 fill
        ln = etree.SubElement(spPr, f"{{{NAMESPACES['a']}}}ln")
        if line_width is not None:
            ln.set('w', str(line_width))
        if color is not None:
            solidFill = etree.SubElement(ln, f"{{{NAMESPACES['a']}}}solidFill")
            srgbClr = etree.SubElement(solidFill, f"{{{NAMESPACES['a']}}}srgbClr")
            srgbClr.set('val', color)


def _apply_marker_style(
    ser_element,
    marker_style: str,
    marker_size: int = None,
    color: str = None
):
    """应用标记点样式"""
    # 清除现有的 marker 元素
    for marker in ser_element.findall('c:marker', namespaces=NAMESPACES):
        ser_element.remove(marker)
    
    if marker_style == "none" or marker_style is None:
        # 创建无标记的 marker 元素
        insert_index = _find_insert_position(ser_element, 'marker')
        marker = etree.Element(f"{{{NAMESPACES['c']}}}marker")
        ser_element.insert(insert_index, marker)
        
        symbol = etree.SubElement(marker, f"{{{NAMESPACES['c']}}}symbol")
        symbol.set('val', 'none')
    else:
        # 创建有标记的 marker 元素
        insert_index = _find_insert_position(ser_element, 'marker')
        marker = etree.Element(f"{{{NAMESPACES['c']}}}marker")
        ser_element.insert(insert_index, marker)
        
        # 标记样式
        symbol = etree.SubElement(marker, f"{{{NAMESPACES['c']}}}symbol")
        symbol.set('val', marker_style)
        
        # 标记大小
        if marker_size is not None:
            size = etree.SubElement(marker, f"{{{NAMESPACES['c']}}}size")
            size.set('val', str(marker_size))
        
        # 标记颜色（如果提供）
        if color is not None:
            spPr = etree.SubElement(marker, f"{{{NAMESPACES['c']}}}spPr")
            solidFill = etree.SubElement(spPr, f"{{{NAMESPACES['a']}}}solidFill")
            srgbClr = etree.SubElement(solidFill, f"{{{NAMESPACES['a']}}}srgbClr")
            srgbClr.set('val', color)


def _find_insert_position(parent, target_element: str) -> int:
    """
    查找元素的正确插入位置（按照 OOXML 规范的顺序）
    
    OOXML 规范中 <c:ser> 的子元素顺序：
    1. idx
    2. order
    3. tx
    4. spPr (shape properties)
    5. marker
    6. cat / xVal
    7. val / yVal
    8. smooth
    """
    element_order = {
        'idx': 0,
        'order': 1,
        'tx': 2,
        'spPr': 3,
        'marker': 4,
        'cat': 5,
        'xVal': 5,
        'val': 6,
        'yVal': 6,
        'smooth': 7,
    }
    
    target_order = element_order.get(target_element, 99)
    
    # 查找第一个顺序大于 target 的元素位置
    for i, child in enumerate(parent):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        child_order = element_order.get(tag, 99)
        if child_order > target_order:
            return i
    
    # 如果没找到，返回末尾位置
    return len(parent)


def get_color_for_series(series_index: int, color_scheme: str = "default") -> str:
    """
    为系列获取颜色
    
    Args:
        series_index: 系列索引（0-based）
        color_scheme: 颜色方案名称
        
    Returns:
        RGB 颜色字符串（6位十六进制）
    """
    colors = COLOR_SCHEMES.get(color_scheme, DEFAULT_COLOR_SEQUENCE)
    return colors[series_index % len(colors)]


# ============================================================================
# 预设样式配置
# ============================================================================

class StyleConfig:
    """样式配置类"""

    def __init__(
        self,
        color_scheme: str = "default",
        line_width_pt: float = 1.0,
        marker_style: str = "none",
        marker_size: int = 5,
        colors: List[str] = None,
    ):
        """
        初始化样式配置

        Args:
            color_scheme: 颜色方案（"default", "dark_only", "light_only" 等）
            line_width_pt: 线宽（pt），任意正数均可
            marker_style: 标记点样式（"none", "circle", "square" 等）
            marker_size: 标记点大小（pt）
            colors: 自定义调色板（6 位十六进制 RGB 列表），提供时优先于 color_scheme
        """
        self.color_scheme = color_scheme
        self.line_width = LINE_WIDTH_PT.get(line_width_pt) or max(1, int(line_width_pt * 12700))
        self.marker_style = marker_style
        self.marker_size = marker_size
        self.colors = [c.lstrip("#").upper() for c in colors] if colors else None

    def apply_to_series(self, ser_element, series_index: int):
        """
        将样式应用到系列元素

        Args:
            ser_element: 系列元素
            series_index: 系列索引（用于选择颜色）
        """
        if self.colors:
            color = self.colors[series_index % len(self.colors)]
        else:
            color = get_color_for_series(series_index, self.color_scheme)
        
        apply_series_style(
            ser_element,
            color=color,
            line_width=self.line_width,
            marker_style=self.marker_style,
            marker_size=self.marker_size,
        )


# 默认样式配置
DEFAULT_STYLE_CONFIG = StyleConfig(
    color_scheme="default",
    line_width_pt=1.0,
    marker_style="none",
    marker_size=5,
)

