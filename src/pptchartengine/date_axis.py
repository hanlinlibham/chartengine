"""
日期轴 XML 直接操作模块

专门用于设置日期轴的底层 XML 属性，绕过 python-pptx 的限制。

使用场景：
- 横轴确定为日期类型
- 需要精确控制刻度密度
- 需要设置日期范围
"""

from lxml import etree
from typing import Optional, Literal
from datetime import datetime, timedelta

from ._log import debug_print as print
from .oxml_ns import NAMESPACES


def excel_date(date: datetime) -> float:
    """
    将 Python datetime 转换为 Excel 日期序号
    
    Excel 的日期基准：1900-01-01 = 1
    """
    base_date = datetime(1899, 12, 30)  # Excel 的实际基准日期
    delta = date - base_date
    return float(delta.days)


def format_category_label(value, number_format: str = "yyyy-mm-dd") -> str:
    """Serialize date-like category values into stable human-readable labels.

    Rules:
    - when the value is date-like, suppress time-of-day by default
    - honor coarse patterns like year / year-month / month-day
    - fall back to plain string for non-date-like values
    """

    dt = _coerce_datetime(value)
    if dt is None:
        return str(value)

    fmt = (number_format or "yyyy-mm-dd").lower()
    if "yyyy/mm" in fmt and "dd" not in fmt:
        return dt.strftime("%Y/%m")
    if "yyyy-mm" in fmt and "dd" not in fmt:
        return dt.strftime("%Y-%m")
    if "yyyy" in fmt and "mm" not in fmt and "dd" not in fmt:
        return dt.strftime("%Y")
    if "mm-dd" in fmt and "yyyy" not in fmt:
        return dt.strftime("%m-%d")
    return dt.strftime("%Y-%m-%d")


def _coerce_datetime(value):
    if hasattr(value, "to_pydatetime"):
        value = value.to_pydatetime()
    if isinstance(value, datetime):
        return value
    if isinstance(value, float):
        try:
            return datetime(1899, 12, 30) + timedelta(days=value)
        except Exception:
            return None
    if isinstance(value, str):
        candidate = value.strip()
        if not candidate:
            return None
        candidate = candidate.replace("T", " ")
        try:
            return datetime.fromisoformat(candidate)
        except ValueError:
            for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y/%m", "%Y-%m", "%Y%m%d"):
                try:
                    return datetime.strptime(candidate, fmt)
                except ValueError:
                    continue
    return None


class DateAxisConfig:
    """
    日期轴 XML 配置类
    
    直接操作 XML 以实现完整的日期轴控制
    """
    
    def __init__(
        self,
        base_unit: Literal["days", "months", "years"] = "days",
        major_unit: float = 7.0,
        major_unit_scale: Literal["days", "months", "years"] = "days",
        minor_unit: Optional[float] = None,
        minor_unit_scale: Literal["days", "months", "years"] = "days",
        min_date: Optional[datetime] = None,
        max_date: Optional[datetime] = None,
        number_format: str = "yyyy-mm-dd",
        auto_adjust: bool = False,
    ):
        """
        初始化日期轴配置
        
        Args:
            base_unit: 基础时间单位（days/months/years）
            major_unit: 主刻度单位数值（如 7 表示每 7 个单位一个刻度）
            major_unit_scale: 主刻度单位类型
            minor_unit: 次刻度单位数值
            minor_unit_scale: 次刻度单位类型
            min_date: 最小日期（用于固定范围）
            max_date: 最大日期（用于固定范围）
            number_format: 日期显示格式
            auto_adjust: 是否允许 PowerPoint 自动调整
        
        Examples:
            >>> # 每周显示一个刻度
            >>> config = DateAxisConfig(
            ...     base_unit="days",
            ...     major_unit=7.0,
            ...     major_unit_scale="days",
            ... )
            
            >>> # 每月显示一个刻度
            >>> config = DateAxisConfig(
            ...     base_unit="months",
            ...     major_unit=1.0,
            ...     major_unit_scale="months",
            ... )
            
            >>> # 固定日期范围
            >>> config = DateAxisConfig(
            ...     base_unit="days",
            ...     major_unit=14.0,
            ...     min_date=datetime(2024, 1, 1),
            ...     max_date=datetime(2024, 12, 31),
            ... )
        """
        self.base_unit = base_unit
        self.major_unit = major_unit
        self.major_unit_scale = major_unit_scale
        self.minor_unit = minor_unit or (major_unit / 7)  # 默认为主刻度的 1/7
        self.minor_unit_scale = minor_unit_scale
        self.min_date = min_date
        self.max_date = max_date
        self.number_format = number_format
        self.auto_adjust = auto_adjust
    
    def apply_to_chart(self, chart):
        """
        应用日期轴配置到图表
        
        Args:
            chart: python-pptx 的 Chart 对象
        """
        print(f"\n📅 应用日期轴配置（XML 直接操作）:")
        
        try:
            # 1. 获取 XML 元素
            chart_element = chart._element
            
            # ⭐ 新方案：数据已在 plots.py 中格式化为字符串并使用 strCache
            print(f"  → 数据已格式化为字符串标签（如 '2024/01'），使用 strCache")
            
            # ⭐ 关键修复：查找catAx并转换为dateAx
            # PowerPoint 不允许在 catAx 中使用时间单位元素，必须使用 dateAx
            cat_ax_elements = chart_element.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}catAx')
            
            if not cat_ax_elements:
                print("  ⚠️ 未找到分类轴元素")
                return
            
            # ⭐ 新方案：不转换为 dateAx，保持 catAx
            # PowerPoint 对 dateAx 的时间单位元素支持不稳定
            # 改用 catAx + numRef/numCache + scaling(min/max) + numFmt(日期格式)
            # 这样更简单且兼容性更好
            
            cat_ax = cat_ax_elements[0]
            print(f"  ✅ 使用 catAx（分类轴）+ 格式化字符串标签")
            
            # ⭐ 设置标签位置为"低"（在坐标轴下方）
            self._set_tick_label_position(cat_ax, 'low')
            print(f"  - 标签位置: 低（坐标轴下方）")
            
            # ⭐ 设置标签间隔（控制显示数量）
            # 由于当前实现走的是 catAx + 格式化字符串标签，major_unit_scale 不能直接
            # 映射为 PowerPoint 的真正 month/year date axis 语义，所以这里按数据长度
            # 做一次显示密度折算。
            category_count = self._infer_category_count(chart_element)
            label_interval = self._resolve_label_interval(category_count)
            self._set_label_interval(cat_ax, label_interval)
            print(f"  - 标签间隔: 每 {label_interval} 个点显示一个标签")
            
            # ⭐ 设置字体大小为 9pt
            try:
                category_axis = chart.category_axis
                from pptx.util import Pt
                category_axis.tick_labels.font.size = Pt(9)
                category_axis.tick_labels.font.name = "黑体"
                print(f"  - 横轴字体: 黑体 9pt")
            except Exception as e:
                print(f"  ⚠️ 字体设置失败: {e}")
            
            print(f"  ✅ 日期轴配置完成")
            
        except Exception as e:
            print(f"  ❌ 日期轴配置失败: {e}")
            import traceback
            traceback.print_exc()

    def _resolve_label_interval(self, category_count: int | None) -> int:
        """Map semantic date intent onto string-label density control."""

        if category_count is None or category_count <= 0:
            return max(1, int(self.major_unit))

        if self.major_unit_scale == "months":
            # Long daily histories should land around 8-12 visible month labels.
            if category_count >= 360:
                return max(1, category_count // 10)
            if category_count >= 120:
                return max(1, category_count // 8)
            return max(1, int(self.major_unit))

        if self.major_unit_scale == "years":
            return max(1, category_count // 6)

        return max(1, int(self.major_unit))

    def _infer_category_count(self, chart_element) -> int | None:
        """Best-effort count of category points from the first category cache."""

        pt_count = chart_element.find('.//c:cat//c:ptCount', namespaces=NAMESPACES)
        if pt_count is not None:
            try:
                return int(pt_count.get('val'))
            except (TypeError, ValueError):
                return None

        pts = chart_element.findall('.//c:cat//c:pt', namespaces=NAMESPACES)
        return len(pts) if pts else None
    
    def _convert_cat_to_numcache(self, chart_element):
        """
        ⚠️ 备用方法：将分类数据从 strCache 转换为 numCache
        
        【重要】此方法已不再默认调用（自修复后）
        
        修复说明：
        - 在 api.py 修复后，python-pptx 会直接生成 numCache
        - 此方法仅在需要手动处理旧文件或特殊场景时使用
        - 如果未来遇到 strCache，此方法会将字符串日期转换为 Excel 序列号
        
        转换逻辑：
        - 从 strCache 读取 "YYYY-MM-DD" 格式的字符串
        - 调用 excel_date() 转换为 Excel 日期序列号（如 45567）
        - 写入 numCache，让 PowerPoint 正确识别为日期
        """
        from lxml import etree
        
        # 查找所有 c:cat 元素
        cat_elements = chart_element.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}cat')
        
        converted_count = 0
        for cat_elem in cat_elements:
            # 查找 strCache（直接子元素）
            str_cache = cat_elem.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}strCache')
            if str_cache is None:
                continue
            
            # 提取数据点
            pt_count_elem = str_cache.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}ptCount')
            pt_count = int(pt_count_elem.get('val')) if pt_count_elem is not None else 0
            
            pts = str_cache.findall('{http://schemas.openxmlformats.org/drawingml/2006/chart}pt')
            
            if not pts:
                continue
            
            # 创建新的 numCache 元素
            num_cache = etree.Element(f"{{{NAMESPACES['c']}}}numCache")
            
            # 添加格式代码
            format_code = etree.SubElement(num_cache, f"{{{NAMESPACES['c']}}}formatCode")
            format_code.text = self.number_format if self.number_format else "yyyy/mm"
            
            # 添加点计数
            new_pt_count = etree.SubElement(num_cache, f"{{{NAMESPACES['c']}}}ptCount")
            new_pt_count.set('val', str(pt_count))
            
            # 复制所有数据点
            for pt in pts:
                idx = pt.get('idx')
                v_elem = pt.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}v')
                if v_elem is not None and v_elem.text:
                    # 创建新的 pt 元素
                    new_pt = etree.SubElement(num_cache, f"{{{NAMESPACES['c']}}}pt")
                    new_pt.set('idx', idx)
                    
                    # 添加值（保持数字格式）
                    new_v = etree.SubElement(new_pt, f"{{{NAMESPACES['c']}}}v")
                    new_v.text = v_elem.text
            
            # 获取 strCache 在父元素中的位置
            str_cache_index = list(cat_elem).index(str_cache)
            
            # 删除旧的 strCache
            cat_elem.remove(str_cache)
            
            # 在相同位置插入新的 numCache
            cat_elem.insert(str_cache_index, num_cache)
            converted_count += 1
        
        if converted_count > 0:
            print(f"  → 已将 {converted_count} 个分类数据转换为 numCache（数值格式）")
    
    def _convert_catax_to_dateax(self, chart_element, cat_ax):
        """
        将 catAx (分类轴) 转换为 dateAx (日期轴)
        
        这是关键修复！PowerPoint 不允许在 catAx 中使用时间单位元素（baseTimeUnit, majorTimeUnit 等）。
        必须将整个 catAx 元素替换为 dateAx。
        
        Args:
            chart_element: 图表的根 XML 元素
            cat_ax: 要转换的 catAx 元素
            
        Returns:
            转换后的 dateAx 元素
        """
        from lxml import etree
        
        # 创建新的 dateAx 元素
        date_ax = etree.Element(f"{{{NAMESPACES['c']}}}dateAx")
        
        # 复制 catAx 的所有子元素到 dateAx
        # 移除不兼容的元素（如 lblAlgn, lblOffset, noMultiLvlLbl）
        # 保留重要元素（如 numFmt, txPr 等）
        skip_elements = ['lblAlgn', 'lblOffset', 'noMultiLvlLbl']
        
        for child in cat_ax:
            tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag_name not in skip_elements:
                # 深度复制元素（包括所有子元素）
                new_child = etree.Element(child.tag, child.attrib)
                new_child.text = child.text
                new_child.tail = child.tail
                
                # 复制所有子元素
                for subchild in child:
                    new_child.append(subchild)
                
                date_ax.append(new_child)
        
        # 在父元素中替换 catAx 为 dateAx
        parent = cat_ax.getparent()
        if parent is not None:
            parent_list = list(parent)
            cat_ax_index = parent_list.index(cat_ax)
            parent.remove(cat_ax)
            parent.insert(cat_ax_index, date_ax)
        
        return date_ax
    
    def _set_or_update_element(self, parent, tag_name: str, value: str):
        """
        设置或更新 XML 元素
        
        Args:
            parent: 父元素
            tag_name: 标签名（不含命名空间）
            value: 属性值
        """
        full_tag = f"{{{NAMESPACES['c']}}}{tag_name}"
        
        # 查找现有元素
        existing = parent.find(f'.//c:{tag_name}', namespaces=NAMESPACES)
        
        if existing is not None:
            # 更新现有元素
            existing.set('val', value)
        else:
            # 创建新元素
            new_element = etree.Element(full_tag)
            new_element.set('val', value)
            
            # ⭐ 关键修复：时间单位元素必须在最后（auto 之后）
            # OOXML 规范的 dateAx 元素顺序：
            # axId → scaling → delete → axPos → majorTickMark → minorTickMark → tickLblPos → numFmt → crossAx → crosses → auto → [时间单位]
            
            # 策略：在 auto 之后插入，如果没有 auto 则追加到最后
            auto_elements = parent.findall('.//c:auto', namespaces=NAMESPACES)
            if auto_elements and len(auto_elements) > 0:
                # 在 auto 之后插入
                auto_elem = auto_elements[0]
                parent_list = list(parent)
                auto_index = parent_list.index(auto_elem)
                parent.insert(auto_index + 1, new_element)
            else:
                # 如果没有 auto，追加到最后
                parent.append(new_element)
    
    def _set_date_range(self, cat_ax):
        """设置日期范围（最小值/最大值）"""
        # 查找或创建 scaling 元素
        scaling = cat_ax.find('.//c:scaling', namespaces=NAMESPACES)
        
        if scaling is None:
            scaling = etree.Element(f"{{{NAMESPACES['c']}}}scaling")
            # 插入到合适位置（通常在 axId 之前）
            ax_id_elements = cat_ax.findall('.//c:axId', namespaces=NAMESPACES)
            if ax_id_elements:
                ax_id_elements[0].addprevious(scaling)
            else:
                cat_ax.insert(0, scaling)
        
        # 设置最小值
        if self.min_date:
            min_val = excel_date(self.min_date)
            self._set_scaling_value(scaling, 'min', str(min_val))
            print(f"  - 最小日期: {self.min_date.strftime('%Y-%m-%d')} (Excel: {min_val})")
        
        # 设置最大值
        if self.max_date:
            max_val = excel_date(self.max_date)
            self._set_scaling_value(scaling, 'max', str(max_val))
            print(f"  - 最大日期: {self.max_date.strftime('%Y-%m-%d')} (Excel: {max_val})")
    
    def _set_scaling_value(self, scaling_element, tag_name: str, value: str):
        """在 scaling 元素中设置 min/max"""
        full_tag = f"{{{NAMESPACES['c']}}}{tag_name}"
        
        existing = scaling_element.find(f'.//c:{tag_name}', namespaces=NAMESPACES)
        
        if existing is not None:
            existing.set('val', value)
        else:
            new_element = etree.Element(full_tag)
            new_element.set('val', value)
            scaling_element.append(new_element)
    
    def _set_date_format(self, axis_element, format_code: str):
        """
        设置日期轴的数字格式（numFmt）
        
        Args:
            axis_element: 轴元素（dateAx 或 catAx）
            format_code: Excel 格式代码（如 "yyyy/mm", "mm-dd" 等）
        """
        from lxml import etree
        
        # 查找或创建 numFmt 元素
        numfmt = axis_element.find('.//c:numFmt', namespaces=NAMESPACES)
        
        if numfmt is None:
            # 创建新的 numFmt 元素
            numfmt = etree.Element(f"{{{NAMESPACES['c']}}}numFmt")
            
            # ⭐ 关键：numFmt 应该在 tickLblPos 之后、crossAx 之前插入
            # 正确顺序：axId → scaling → delete → axPos → majorTickMark → minorTickMark → tickLblPos → numFmt → crossAx → crosses → auto → [时间单位]
            tick_lbl_pos_elements = axis_element.findall('.//c:tickLblPos', namespaces=NAMESPACES)
            if tick_lbl_pos_elements:
                tick_lbl_pos = tick_lbl_pos_elements[0]
                parent_list = list(axis_element)
                pos_index = parent_list.index(tick_lbl_pos)
                axis_element.insert(pos_index + 1, numfmt)
            else:
                # Fallback：在 crossAx 之前
                cross_ax_elements = axis_element.findall('.//c:crossAx', namespaces=NAMESPACES)
                if cross_ax_elements:
                    cross_ax = cross_ax_elements[0]
                    parent_list = list(axis_element)
                    cross_ax_index = parent_list.index(cross_ax)
                    axis_element.insert(cross_ax_index, numfmt)
                else:
                    # 最后的fallback：追加到最后
                    axis_element.append(numfmt)
        
        # 设置格式代码和 sourceLinked
        numfmt.set('formatCode', format_code)
        numfmt.set('sourceLinked', '0')  # 不链接到数据源
    
    def _set_date_range_from_data(self, date_ax, chart_element):
        """
        从数据中提取日期范围并设置到 dateAx 的 scaling 中
        
        这是关键修复！PowerPoint 的 dateAx 默认从 0（1900-01-01）开始计算。
        必须明确设置 min 和 max 值来指定实际的日期范围。
        
        Args:
            date_ax: dateAx 元素
            chart_element: 图表根元素
        """
        from lxml import etree
        
        # 查找第一个 cat 元素中的 numCache 或 numRef
        cat_elements = chart_element.findall('.//c:cat', namespaces=NAMESPACES)
        
        if not cat_elements:
            print("  ⚠️ 未找到分类数据，跳过日期范围设置")
            return
        
        cat_elem = cat_elements[0]
        
        # 尝试从 numCache 中提取日期值
        num_cache = cat_elem.find('.//c:numCache', namespaces=NAMESPACES)
        if num_cache is not None:
            pt_elements = num_cache.findall('.//c:pt', namespaces=NAMESPACES)
            if pt_elements:
                # 获取第一个和最后一个日期值
                first_pt = pt_elements[0]
                last_pt = pt_elements[-1]
                
                first_v = first_pt.find('c:v', namespaces=NAMESPACES)
                last_v = last_pt.find('c:v', namespaces=NAMESPACES)
                
                if first_v is not None and last_v is not None:
                    try:
                        min_val = float(first_v.text)
                        max_val = float(last_v.text)
                        
                        # 转换为日期显示
                        from datetime import datetime, timedelta
                        min_date = datetime(1899, 12, 30) + timedelta(days=min_val)
                        max_date = datetime(1899, 12, 30) + timedelta(days=max_val)
                        
                        # 设置 scaling 的 min 和 max
                        self._set_scaling_min_max(date_ax, min_val, max_val)
                        
                        print(f"  ✅ 已设置日期范围: {min_date.strftime('%Y-%m-%d')} ~ {max_date.strftime('%Y-%m-%d')}")
                        print(f"     (Excel 序列号: {min_val} ~ {max_val})")
                        return
                    except (ValueError, AttributeError) as e:
                        print(f"  ⚠️ 提取日期范围失败: {e}")
        
        print("  ⚠️ 未能从数据中提取日期范围")
    
    def _set_scaling_min_max(self, axis_element, min_val, max_val):
        """
        设置 scaling 元素的 min 和 max 值
        
        Args:
            axis_element: 轴元素（dateAx 或 catAx）
            min_val: 最小值（Excel 日期序列号）
            max_val: 最大值（Excel 日期序列号）
        """
        from lxml import etree
        
        # 查找或创建 scaling 元素
        scaling = axis_element.find('.//c:scaling', namespaces=NAMESPACES)
        
        if scaling is None:
            # 创建 scaling 元素
            scaling = etree.Element(f"{{{NAMESPACES['c']}}}scaling")
            
            # scaling 应该在 axId 之后、delete 之前插入
            ax_id_elements = axis_element.findall('.//c:axId', namespaces=NAMESPACES)
            if ax_id_elements:
                ax_id = ax_id_elements[0]
                parent_list = list(axis_element)
                ax_id_index = parent_list.index(ax_id)
                axis_element.insert(ax_id_index + 1, scaling)
            else:
                axis_element.insert(0, scaling)
        
        # 设置 min 值
        self._set_scaling_value(scaling, 'min', str(min_val))
        
        # 设置 max 值
        self._set_scaling_value(scaling, 'max', str(max_val))
    
    def _set_tick_label_position(self, axis_element, position: str):
        """
        设置刻度标签位置
        
        Args:
            axis_element: 轴元素（catAx 或 dateAx）
            position: 位置值（'low', 'high', 'nextTo'）
        """
        # 查找 tickLblPos 元素
        tick_lbl_pos = axis_element.find('.//c:tickLblPos', namespaces=NAMESPACES)
        
        if tick_lbl_pos is not None:
            # 更新现有元素
            tick_lbl_pos.set('val', position)
        else:
            # 如果不存在，不创建（使用默认值）
            pass
    
    def _set_label_interval(self, axis_element, interval: int):
        """
        设置标签显示间隔
        
        Args:
            axis_element: 轴元素（catAx 或 dateAx）
            interval: 间隔值（每隔多少个数据点显示一个标签）
        """
        from lxml import etree
        
        # 查找或创建 tickLblSkip 元素
        tick_lbl_skip = axis_element.find('.//c:tickLblSkip', namespaces=NAMESPACES)
        
        if tick_lbl_skip is None:
            # 创建新的 tickLblSkip 元素
            tick_lbl_skip = etree.Element(f"{{{NAMESPACES['c']}}}tickLblSkip")
            
            # tickLblSkip 应该在 auto 之后插入
            auto_elements = axis_element.findall('.//c:auto', namespaces=NAMESPACES)
            if auto_elements:
                auto_elem = auto_elements[0]
                parent_list = list(axis_element)
                auto_index = parent_list.index(auto_elem)
                axis_element.insert(auto_index + 1, tick_lbl_skip)
            else:
                # Fallback：追加到最后
                axis_element.append(tick_lbl_skip)
        
        # 设置间隔值
        tick_lbl_skip.set('val', str(interval))




# ============================================================================
# 便捷预设
# ============================================================================

# 每日刻度（适用于短期数据：1-30天）
DAILY_TICKS = DateAxisConfig(
    base_unit="days",
    major_unit=1.0,
    major_unit_scale="days",
    number_format="mm-dd",
)

# 每周刻度（适用于中期数据：1-6个月）
WEEKLY_TICKS = DateAxisConfig(
    base_unit="days",
    major_unit=7.0,
    major_unit_scale="days",
    number_format="yyyy-mm-dd",
)

# 每双周刻度（适用于季度数据）
BIWEEKLY_TICKS = DateAxisConfig(
    base_unit="days",
    major_unit=14.0,
    major_unit_scale="days",
    number_format="yyyy-mm-dd",
)

# 每月刻度（适用于年度数据）
MONTHLY_TICKS = DateAxisConfig(
    base_unit="months",
    major_unit=1.0,
    major_unit_scale="months",
    number_format="yyyy-mm",
)

# 每季度刻度（适用于多年数据）
QUARTERLY_TICKS = DateAxisConfig(
    base_unit="months",
    major_unit=3.0,
    major_unit_scale="months",
    number_format="yyyy-mm",
)

# 每年刻度（适用于长期数据）
YEARLY_TICKS = DateAxisConfig(
    base_unit="years",
    major_unit=1.0,
    major_unit_scale="years",
    number_format="yyyy",
)
