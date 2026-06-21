"""
图表布局配置模块

控制图表的结构性属性：
- 图例位置、大小
- 横轴配置（日期轴、刻度）
- 纵轴配置（位置、刻度）
"""

from pptx.enum.chart import XL_LEGEND_POSITION, XL_TICK_MARK, XL_TICK_LABEL_POSITION
from pptx.util import Pt
from typing import Optional, Dict, Any

from ._log import debug_print as print

_C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

# CT_CatAx / CT_ValAx child order; ``title`` must sit after these and before
# numFmt/majorTickMark/... (OOXML is order-sensitive).
_AXIS_PRE_TITLE_TAGS = ("axId", "scaling", "delete", "axPos",
                        "majorGridlines", "minorGridlines")


def _write_axis_title(ax, text, *, font_name=None, font_size_pt=None):
    """Insert a ``c:title`` into a value/category axis lxml element.

    Works on raw ``c:valAx`` / ``c:catAx`` elements so it covers the primary,
    secondary, and category axes uniformly (python-pptx only exposes the
    primary pair). Idempotent: replaces any existing axis title.
    """
    from lxml import etree

    def c(tag):
        return f"{{{_C_NS}}}{tag}"

    def a(tag):
        return f"{{{_A_NS}}}{tag}"

    existing = ax.find(c("title"))
    if existing is not None:
        ax.remove(existing)

    title = etree.Element(c("title"))
    tx = etree.SubElement(title, c("tx"))
    rich = etree.SubElement(tx, c("rich"))
    etree.SubElement(rich, a("bodyPr"))
    etree.SubElement(rich, a("lstStyle"))
    p = etree.SubElement(rich, a("p"))
    pPr = etree.SubElement(p, a("pPr"))
    defRPr = etree.SubElement(pPr, a("defRPr"))
    defRPr.set("b", "0")
    if font_size_pt:
        defRPr.set("sz", str(int(font_size_pt * 100)))
    if font_name:
        etree.SubElement(defRPr, a("latin")).set("typeface", font_name)
    r = etree.SubElement(p, a("r"))
    rPr = etree.SubElement(r, a("rPr"))
    rPr.set("lang", "en-US")
    if font_size_pt:
        rPr.set("sz", str(int(font_size_pt * 100)))
    if font_name:
        etree.SubElement(rPr, a("latin")).set("typeface", font_name)
    etree.SubElement(r, a("t")).text = text
    overlay = etree.SubElement(title, c("overlay"))
    overlay.set("val", "0")

    # Insert right after the last pre-title child so element order stays valid.
    insert_at = 0
    for i, child in enumerate(ax):
        if etree.QName(child).localname in _AXIS_PRE_TITLE_TAGS:
            insert_at = i + 1
    ax.insert(insert_at, title)


# ============================================================================
# 图例配置
# ============================================================================

class LegendConfig:
    """图例配置"""
    
    # 预设位置
    BOTTOM = XL_LEGEND_POSITION.BOTTOM      # 底部
    TOP = XL_LEGEND_POSITION.TOP            # 顶部
    LEFT = XL_LEGEND_POSITION.LEFT          # 左侧
    RIGHT = XL_LEGEND_POSITION.RIGHT        # 右侧
    CORNER = XL_LEGEND_POSITION.CORNER      # 右上角
    
    def __init__(
        self,
        position = XL_LEGEND_POSITION.BOTTOM,
        font_size_pt: float = 10,
        font_name: str = "微软雅黑",
        include_in_layout: bool = False,
    ):
        """
        初始化图例配置
        
        Args:
            position: 图例位置
                - XL_LEGEND_POSITION.BOTTOM (默认)
                - XL_LEGEND_POSITION.TOP
                - XL_LEGEND_POSITION.LEFT
                - XL_LEGEND_POSITION.RIGHT
                - XL_LEGEND_POSITION.CORNER
            font_size_pt: 字体大小（pt）
            font_name: 字体名称（默认：黑体）
            include_in_layout: 是否包含在布局中（False 防止图例覆盖图表）
        """
        self.position = position
        self.font_size_pt = font_size_pt
        self.font_name = font_name
        self.include_in_layout = include_in_layout
    
    def apply_to_chart(self, chart):
        """应用到图表"""
        chart.has_legend = True
        chart.legend.position = self.position
        chart.legend.include_in_layout = self.include_in_layout
        
        if self.font_size_pt:
            chart.legend.font.size = Pt(self.font_size_pt)
        if self.font_name:
            chart.legend.font.name = self.font_name
        
        print(f"  - 图例配置: 位置={self._position_name()}, 字体={self.font_name} {self.font_size_pt}pt")
    
    def _position_name(self):
        """获取位置名称"""
        position_names = {
            XL_LEGEND_POSITION.BOTTOM: "底部",
            XL_LEGEND_POSITION.TOP: "顶部",
            XL_LEGEND_POSITION.LEFT: "左侧",
            XL_LEGEND_POSITION.RIGHT: "右侧",
            XL_LEGEND_POSITION.CORNER: "右上角",
        }
        return position_names.get(self.position, "未知")


# ============================================================================
# 横轴配置
# ============================================================================

class CategoryAxisConfig:
    """横轴（分类轴）配置"""
    
    def __init__(
        self,
        is_date_axis: bool = False,
        major_unit: Optional[float] = None,
        major_unit_days: Optional[int] = None,
        tick_label_position = XL_TICK_LABEL_POSITION.LOW,
        major_tick_mark = XL_TICK_MARK.NONE,
        minor_tick_mark = XL_TICK_MARK.NONE,
        number_format: Optional[str] = None,
        font_size_pt: float = 10,
        font_name: str = "微软雅黑",
        axis_title: Optional[str] = None,
    ):
        """
        初始化横轴配置
        
        Args:
            is_date_axis: 是否为日期轴
            major_unit: 主刻度单位（用于日期轴，单位为天）
            major_unit_days: 主刻度单位（天）- 便捷参数，与 major_unit 相同
            tick_label_position: 刻度标签位置
            major_tick_mark: 主刻度线样式
            minor_tick_mark: 次刻度线样式
            number_format: 数字格式（如 'yyyy-mm-dd'）
            font_size_pt: 字体大小
            font_name: 字体名称（默认：黑体）
        """
        self.is_date_axis = is_date_axis
        self.major_unit = major_unit or major_unit_days
        self.tick_label_position = tick_label_position
        self.major_tick_mark = major_tick_mark
        self.minor_tick_mark = minor_tick_mark
        self.number_format = number_format
        self.font_size_pt = font_size_pt
        self.font_name = font_name
        self.axis_title = axis_title

    def apply_to_chart(self, chart):
        """应用到图表"""
        try:
            category_axis = chart.category_axis
            
            # 设置日期轴
            if self.is_date_axis:
                from pptx.enum.chart import XL_CATEGORY_TYPE
                category_axis.category_type = XL_CATEGORY_TYPE.TIME_SCALE
                print(f"  - 横轴: 日期轴")
                
                # 设置刻度单位
                if self.major_unit:
                    from pptx.enum.chart import XL_TIME_UNIT
                    category_axis.major_unit = self.major_unit
                    category_axis.major_unit_scale = XL_TIME_UNIT.DAYS
                    print(f"  - 刻度单位: {self.major_unit} 天")
            else:
                print(f"  - 横轴: 分类轴")
            
            # 设置刻度标签位置
            category_axis.tick_label_position = self.tick_label_position
            
            # 设置刻度线
            category_axis.major_tick_mark = self.major_tick_mark
            category_axis.minor_tick_mark = self.minor_tick_mark
            
            # 设置数字格式
            if self.number_format:
                category_axis.tick_labels.number_format = self.number_format
                print(f"  - 日期格式: {self.number_format}")
            
            # 设置字体大小和字体名称
            if self.font_size_pt:
                category_axis.tick_labels.font.size = Pt(self.font_size_pt)
            if self.font_name:
                category_axis.tick_labels.font.name = self.font_name
            print(f"  - 横轴字体: {self.font_name} {self.font_size_pt}pt")

            # 轴标题
            if self.axis_title:
                _write_axis_title(category_axis._element, self.axis_title,
                                  font_name=self.font_name, font_size_pt=self.font_size_pt)
                print(f"  - 横轴标题: {self.axis_title}")

        except Exception as e:
            print(f"  ⚠️ 横轴配置失败: {e}")


# ============================================================================
# 纵轴配置
# ============================================================================

class ValueAxisConfig:
    """纵轴（数值轴）配置"""
    
    def __init__(
        self,
        tick_label_position = XL_TICK_LABEL_POSITION.LOW,
        major_tick_mark = XL_TICK_MARK.OUTSIDE,
        minor_tick_mark = XL_TICK_MARK.NONE,
        number_format: Optional[str] = None,
        font_size_pt: float = 10,
        font_name: str = "微软雅黑",
        has_major_gridlines: bool = True,
        min_value: Optional[float] = None,
        max_value: Optional[float] = None,
        major_unit: Optional[float] = None,
        axis_title: Optional[str] = None,
    ):
        """
        初始化纵轴配置
        
        Args:
            tick_label_position: 刻度标签位置
            major_tick_mark: 主刻度线样式
            minor_tick_mark: 次刻度线样式
            number_format: 数字格式（如 '0.00%'）
            font_size_pt: 字体大小
            font_name: 字体名称（默认：黑体）
            has_major_gridlines: 是否显示主网格线
            min_value: 最小值（可选）
            max_value: 最大值（可选）
            major_unit: 主刻度单位（可选）
        """
        self.tick_label_position = tick_label_position
        self.major_tick_mark = major_tick_mark
        self.minor_tick_mark = minor_tick_mark
        self.number_format = number_format
        self.font_size_pt = font_size_pt
        self.font_name = font_name
        self.has_major_gridlines = has_major_gridlines
        self.min_value = min_value
        self.max_value = max_value
        self.major_unit = major_unit
        self.axis_title = axis_title

    def apply_to_chart(self, chart):
        """应用到图表（主值轴）"""
        try:
            # ⭐ 通过 XML 直接设置主值轴，确保格式生效
            chart_element = chart._element
            val_ax_elements = chart_element.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}valAx')
            
            # 主值轴：纵向图在左侧（'l'），横向条形图在底部（'b'）
            primary_ax = None
            for ax in val_ax_elements:
                ax_pos = ax.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}axPos')
                if ax_pos is not None and ax_pos.get('val') in ('l', 'b'):
                    primary_ax = ax
                    break
            
            if primary_ax is None:
                print(f"  ⚠️ 未找到主值轴")
                return
            
            from lxml import etree
            
            # 设置数字格式（通过 XML）
            if self.number_format:
                num_fmt = primary_ax.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}numFmt')
                if num_fmt is None:
                    # 创建 numFmt 元素
                    num_fmt = etree.Element('{http://schemas.openxmlformats.org/drawingml/2006/chart}numFmt')
                    # 在 tickLblPos 之后插入
                    tick_lbl_pos = primary_ax.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}tickLblPos')
                    if tick_lbl_pos is not None:
                        parent_list = list(primary_ax)
                        pos_index = parent_list.index(tick_lbl_pos)
                        primary_ax.insert(pos_index + 1, num_fmt)
                    else:
                        primary_ax.append(num_fmt)
                
                num_fmt.set('formatCode', self.number_format)
                num_fmt.set('sourceLinked', '0')
                print(f"  - 主值轴格式: {self.number_format}")
            
            # 设置字体（通过 XML）
            if self.font_size_pt or self.font_name:
                self._apply_font_to_axis_xml(primary_ax, self.font_name, self.font_size_pt)
                print(f"  - 主值轴字体: {self.font_name} {self.font_size_pt}pt")

            # 轴标题
            if self.axis_title:
                _write_axis_title(primary_ax, self.axis_title,
                                  font_name=self.font_name, font_size_pt=self.font_size_pt)
                print(f"  - 主值轴标题: {self.axis_title}")
            
            # 设置轴范围和刻度间隔（通过 XML）
            if self.min_value is not None or self.max_value is not None:
                scaling = primary_ax.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}scaling')
                if scaling is None:
                    scaling = etree.Element('{http://schemas.openxmlformats.org/drawingml/2006/chart}scaling')
                    # 在 axId 之后插入
                    ax_id = primary_ax.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}axId')
                    if ax_id is not None:
                        parent_list = list(primary_ax)
                        ax_id_index = parent_list.index(ax_id)
                        primary_ax.insert(ax_id_index + 1, scaling)
                    else:
                        primary_ax.insert(0, scaling)
                    
                    # 添加 orientation
                    orientation = etree.SubElement(scaling, '{http://schemas.openxmlformats.org/drawingml/2006/chart}orientation')
                    orientation.set('val', 'minMax')
                
                if self.min_value is not None:
                    min_elem = scaling.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}min')
                    if min_elem is None:
                        min_elem = etree.SubElement(scaling, '{http://schemas.openxmlformats.org/drawingml/2006/chart}min')
                    min_elem.set('val', str(self.min_value))
                    print(f"  - 主值轴最小值: {self.min_value}")
                
                if self.max_value is not None:
                    max_elem = scaling.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}max')
                    if max_elem is None:
                        max_elem = etree.SubElement(scaling, '{http://schemas.openxmlformats.org/drawingml/2006/chart}max')
                    max_elem.set('val', str(self.max_value))
                    print(f"  - 主值轴最大值: {self.max_value}")
            
            if self.major_unit is not None:
                major_unit_elem = primary_ax.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}majorUnit')
                if major_unit_elem is None:
                    major_unit_elem = etree.SubElement(primary_ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}majorUnit')
                major_unit_elem.set('val', str(self.major_unit))
                print(f"  - 主值轴刻度间隔: {self.major_unit}")
            
        except Exception as e:
            print(f"  ⚠️ 主值轴配置失败: {e}")
            import traceback
            traceback.print_exc()
    
    def _apply_font_to_axis_xml(self, axis_element, font_name: str, font_size_pt: float):
        """通过 XML 设置轴字体"""
        from lxml import etree
        
        # 查找或创建 txPr
        txPr = axis_element.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}txPr')
        if txPr is None:
            txPr = etree.SubElement(axis_element, '{http://schemas.openxmlformats.org/drawingml/2006/chart}txPr')
            bodyPr = etree.SubElement(txPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
            lstStyle = etree.SubElement(txPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}lstStyle')
        
        # 设置字体
        p = txPr.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}p')
        if p is None:
            p = etree.SubElement(txPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}p')
        
        pPr = p.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
        if pPr is None:
            pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
        
        defRPr = pPr.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr')
        if defRPr is None:
            defRPr = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr')
        
        if font_size_pt:
            defRPr.set('sz', str(int(font_size_pt * 100)))
        
        if font_name:
            latin = defRPr.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}latin')
            if latin is None:
                latin = etree.SubElement(defRPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}latin')
            latin.set('typeface', font_name)



# ============================================================================
# 默认配置
# ============================================================================

# 默认图例配置：底部、9pt 黑体
DEFAULT_LEGEND_CONFIG = LegendConfig(
    position=XL_LEGEND_POSITION.BOTTOM,
    font_size_pt=9,
    font_name="微软雅黑",
    include_in_layout=False,
)

# 默认横轴配置：普通分类轴、外侧刻度线、9pt 黑体
DEFAULT_CATEGORY_AXIS_CONFIG = CategoryAxisConfig(
    is_date_axis=False,
    major_tick_mark=XL_TICK_MARK.OUTSIDE,  # ⭐ 改为 OUTSIDE（与原图一致）
    minor_tick_mark=XL_TICK_MARK.NONE,
    font_size_pt=9,
    font_name="微软雅黑",
)

# 默认纵轴配置：无网格线、9pt 黑体
DEFAULT_VALUE_AXIS_CONFIG = ValueAxisConfig(
    tick_label_position=XL_TICK_LABEL_POSITION.LOW,
    major_tick_mark=XL_TICK_MARK.OUTSIDE,
    minor_tick_mark=XL_TICK_MARK.NONE,
    font_size_pt=9,
    font_name="微软雅黑",
    has_major_gridlines=False,  # ⭐ 改为 False（原图无网格线）
)


# ============================================================================
# 完整布局配置
# ============================================================================

class ChartLayoutConfig:
    """完整的图表布局配置"""
    
    def __init__(
        self,
        title: Optional[str] = None,
        legend_config: Optional[LegendConfig] = None,
        category_axis_config: Optional[CategoryAxisConfig] = None,
        value_axis_config: Optional[ValueAxisConfig] = None,
        secondary_value_axis_config: Optional[ValueAxisConfig] = None,  # ⭐ 新增次值轴配置
        date_axis_config = None,  # DateAxisConfig（避免循环导入）
    ):
        """
        初始化布局配置
        
        Args:
            title: 图表标题
            legend_config: 图例配置
            category_axis_config: 横轴配置
            value_axis_config: 主值轴（左轴）配置
            secondary_value_axis_config: 次值轴（右轴）配置
            date_axis_config: 日期轴配置（DateAxisConfig，用于精确控制日期轴）
        """
        self.title = title
        self.legend_config = legend_config or DEFAULT_LEGEND_CONFIG
        self.category_axis_config = category_axis_config or DEFAULT_CATEGORY_AXIS_CONFIG
        self.value_axis_config = value_axis_config or DEFAULT_VALUE_AXIS_CONFIG
        self.secondary_value_axis_config = secondary_value_axis_config  # ⭐ 新增
        self.date_axis_config = date_axis_config  # ⭐ 新增
    
    def apply_to_chart(self, chart):
        """应用所有配置到图表"""
        print(f"\n⚙️ 应用布局配置:")
        
        # 应用标题
        if self.title:
            try:
                chart.has_title = True
                chart.chart_title.text_frame.text = self.title
                print(f"  - 标题: {self.title}")
            except Exception as e:
                print(f"  ⚠️ 标题设置失败: {e}")
        
        # 应用图例配置
        if self.legend_config:
            self.legend_config.apply_to_chart(chart)
        
        # ⭐ 应用日期轴配置（优先于普通横轴配置）
        if self.date_axis_config:
            self.date_axis_config.apply_to_chart(chart)
        # 应用横轴配置（如果没有日期轴配置）
        elif self.category_axis_config:
            self.category_axis_config.apply_to_chart(chart)
        
        # 应用主值轴配置（左轴）
        if self.value_axis_config:
            self.value_axis_config.apply_to_chart(chart)
        
        # ⭐ 应用次值轴配置（右轴）
        if self.secondary_value_axis_config:
            self._apply_secondary_axis_config(chart)
    
    def _apply_secondary_axis_config(self, chart):
        """应用次值轴配置（通过 XML）"""
        try:
            # 通过 XML 查找次值轴
            chart_element = chart._element
            val_ax_elements = chart_element.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}valAx')
            
            # 次值轴通常是第二个 valAx（position='r'）
            secondary_ax = None
            for ax in val_ax_elements:
                ax_pos = ax.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}axPos')
                if ax_pos is not None and ax_pos.get('val') == 'r':
                    secondary_ax = ax
                    break
            
            if secondary_ax is None:
                print(f"  ⚠️ 未找到次值轴")
                return
            
            config = self.secondary_value_axis_config
            
            # 设置数字格式
            if config.number_format:
                num_fmt = secondary_ax.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}numFmt')
                if num_fmt is not None:
                    num_fmt.set('formatCode', config.number_format)
                    num_fmt.set('sourceLinked', '0')
                print(f"  - 次值轴格式: {config.number_format}")

            # 轴标题
            if config.axis_title:
                _write_axis_title(secondary_ax, config.axis_title,
                                  font_name=config.font_name, font_size_pt=config.font_size_pt)
                print(f"  - 次值轴标题: {config.axis_title}")
            
            # 设置字体
            if config.font_size_pt or config.font_name:
                from pptx.util import Pt
                from lxml import etree
                
                # 查找或创建 txPr (文本属性)
                txPr = secondary_ax.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}txPr')
                if txPr is None:
                    txPr = etree.SubElement(secondary_ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}txPr')
                    bodyPr = etree.SubElement(txPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
                    lstStyle = etree.SubElement(txPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}lstStyle')
                
                # 设置字体
                p = txPr.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}p')
                if p is None:
                    p = etree.SubElement(txPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}p')
                
                pPr = p.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
                if pPr is None:
                    pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/drawingml/2006/main}pPr')
                
                defRPr = pPr.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr')
                if defRPr is None:
                    defRPr = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr')
                
                if config.font_size_pt:
                    defRPr.set('sz', str(int(config.font_size_pt * 100)))
                
                if config.font_name:
                    latin = defRPr.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}latin')
                    if latin is None:
                        latin = etree.SubElement(defRPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}latin')
                    latin.set('typeface', config.font_name)
                
                print(f"  - 次值轴字体: {config.font_name} {config.font_size_pt}pt")
            
            # ⭐ 设置轴范围和刻度间隔（通过 XML）
            from lxml import etree
            
            if config.min_value is not None or config.max_value is not None:
                # 查找或创建 scaling 元素
                scaling = secondary_ax.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}scaling')
                if scaling is None:
                    scaling = etree.SubElement(secondary_ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}scaling')
                    # 添加 orientation
                    orientation = etree.SubElement(scaling, '{http://schemas.openxmlformats.org/drawingml/2006/chart}orientation')
                    orientation.set('val', 'minMax')
                
                if config.min_value is not None:
                    min_elem = scaling.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}min')
                    if min_elem is None:
                        min_elem = etree.SubElement(scaling, '{http://schemas.openxmlformats.org/drawingml/2006/chart}min')
                    min_elem.set('val', str(config.min_value))
                    print(f"  - 次值轴最小值: {config.min_value}")
                
                if config.max_value is not None:
                    max_elem = scaling.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}max')
                    if max_elem is None:
                        max_elem = etree.SubElement(scaling, '{http://schemas.openxmlformats.org/drawingml/2006/chart}max')
                    max_elem.set('val', str(config.max_value))
                    print(f"  - 次值轴最大值: {config.max_value}")
            
            if config.major_unit is not None:
                major_unit_elem = secondary_ax.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}majorUnit')
                if major_unit_elem is None:
                    major_unit_elem = etree.SubElement(secondary_ax, '{http://schemas.openxmlformats.org/drawingml/2006/chart}majorUnit')
                major_unit_elem.set('val', str(config.major_unit))
                print(f"  - 次值轴刻度间隔: {config.major_unit}")
            
        except Exception as e:
            print(f"  ⚠️ 次值轴配置失败: {e}")
            import traceback
            traceback.print_exc()
