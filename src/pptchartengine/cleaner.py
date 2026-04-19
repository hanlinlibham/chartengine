"""ChartJunkCleaner — 自动清洗图表，去除默认 PPT 味

在 ChartBuilder.build() 完成后自动执行：
- 移除图表外边框
- Y 轴网格线 -> 极浅虚线或隐藏
- 坐标轴刻度线 -> inside 或 none
- 图例框边框 -> 移除
- 默认线宽 -> 1.5pt
"""

from lxml import etree

from .oxml_ns import NAMESPACES


class ChartJunkCleaner:
    """清洗图表默认样式，向 JP 标准靠拢"""

    # 极浅灰色（用于残留网格线）
    GRID_COLOR = "E8E8E8"
    GRID_WIDTH = 6350     # 0.5pt in EMUs
    LINE_WIDTH = 19050    # 1.5pt in EMUs

    def __init__(self, chart):
        self.chart = chart
        self.chartSpace = chart._chartSpace
        self.plotArea = self.chartSpace.plotArea

    def clean(self):
        """执行全部清洗步骤"""
        self._remove_chart_border()
        self._clean_gridlines()
        self._clean_tick_marks()
        self._clean_legend_border()
        self._clean_plot_area_border()
        return self.chart

    def _remove_chart_border(self):
        """移除图表外边框"""
        spPr = self.chartSpace.find('c:spPr', namespaces=NAMESPACES)
        if spPr is None:
            spPr = etree.SubElement(self.chartSpace, f"{{{NAMESPACES['c']}}}spPr")

        # 移除线条
        ln = spPr.find('a:ln', namespaces=NAMESPACES)
        if ln is not None:
            spPr.remove(ln)
        ln = etree.SubElement(spPr, f"{{{NAMESPACES['a']}}}ln")
        etree.SubElement(ln, f"{{{NAMESPACES['a']}}}noFill")

    def _clean_gridlines(self):
        """将 Y 轴网格线设为极浅虚线或隐藏"""
        # 处理所有值轴的主要网格线
        for axis_tag in ['valAx', 'catAx']:
            for axis in self.plotArea.findall(f'c:{axis_tag}', namespaces=NAMESPACES):
                major_gl = axis.find('c:majorGridlines', namespaces=NAMESPACES)
                if major_gl is not None:
                    self._set_light_gridline(major_gl)

                # 移除次要网格线
                minor_gl = axis.find('c:minorGridlines', namespaces=NAMESPACES)
                if minor_gl is not None:
                    axis.remove(minor_gl)

    def _set_light_gridline(self, gridline_elem):
        """将网格线设为极浅灰色虚线"""
        # 清除现有 spPr
        spPr = gridline_elem.find('c:spPr', namespaces=NAMESPACES)
        if spPr is not None:
            gridline_elem.remove(spPr)

        spPr = etree.SubElement(gridline_elem, f"{{{NAMESPACES['c']}}}spPr")
        ln = etree.SubElement(spPr, f"{{{NAMESPACES['a']}}}ln")
        ln.set('w', str(self.GRID_WIDTH))

        # 颜色
        solidFill = etree.SubElement(ln, f"{{{NAMESPACES['a']}}}solidFill")
        srgbClr = etree.SubElement(solidFill, f"{{{NAMESPACES['a']}}}srgbClr")
        srgbClr.set('val', self.GRID_COLOR)

        # 虚线样式
        prstDash = etree.SubElement(ln, f"{{{NAMESPACES['a']}}}prstDash")
        prstDash.set('val', 'dot')

    def _clean_tick_marks(self):
        """坐标轴刻度线 -> none"""
        for axis_tag in ['valAx', 'catAx']:
            for axis in self.plotArea.findall(f'c:{axis_tag}', namespaces=NAMESPACES):
                # 主刻度线
                major_tick = axis.find('c:majorTickMark', namespaces=NAMESPACES)
                if major_tick is not None:
                    major_tick.set('val', 'none')
                else:
                    mt = etree.SubElement(axis, f"{{{NAMESPACES['c']}}}majorTickMark")
                    mt.set('val', 'none')

                # 次刻度线
                minor_tick = axis.find('c:minorTickMark', namespaces=NAMESPACES)
                if minor_tick is not None:
                    minor_tick.set('val', 'none')
                else:
                    mt = etree.SubElement(axis, f"{{{NAMESPACES['c']}}}minorTickMark")
                    mt.set('val', 'none')

    def _clean_legend_border(self):
        """移除图例框边框"""
        legend = self.chartSpace.find('.//c:legend', namespaces=NAMESPACES)
        if legend is None:
            return

        spPr = legend.find('c:spPr', namespaces=NAMESPACES)
        if spPr is None:
            spPr = etree.SubElement(legend, f"{{{NAMESPACES['c']}}}spPr")

        # 移除线条
        ln = spPr.find('a:ln', namespaces=NAMESPACES)
        if ln is not None:
            spPr.remove(ln)
        ln = etree.SubElement(spPr, f"{{{NAMESPACES['a']}}}ln")
        etree.SubElement(ln, f"{{{NAMESPACES['a']}}}noFill")

        # 移除填充
        for fill_tag in ['solidFill', 'gradFill', 'pattFill']:
            fill = spPr.find(f'a:{fill_tag}', namespaces=NAMESPACES)
            if fill is not None:
                spPr.remove(fill)
        etree.SubElement(spPr, f"{{{NAMESPACES['a']}}}noFill")

    def _clean_plot_area_border(self):
        """移除 plotArea 边框"""
        spPr = self.plotArea.find('c:spPr', namespaces=NAMESPACES)
        if spPr is None:
            spPr = etree.SubElement(self.plotArea, f"{{{NAMESPACES['c']}}}spPr")

        ln = spPr.find('a:ln', namespaces=NAMESPACES)
        if ln is not None:
            spPr.remove(ln)
        ln = etree.SubElement(spPr, f"{{{NAMESPACES['a']}}}ln")
        etree.SubElement(ln, f"{{{NAMESPACES['a']}}}noFill")


def clean_chart(chart):
    """便捷函数: 对 chart 执行全部清洗"""
    return ChartJunkCleaner(chart).clean()
