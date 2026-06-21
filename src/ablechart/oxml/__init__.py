"""
OXML 包 - 底层 XML 操作

这个包负责所有的 lxml 操作，不关心业务逻辑（主轴/次轴）。
只知道如何根据指令创建标准的 OOXML 结构。
"""

from .axes import extract_axis_ids, create_value_axis, optimize_axis_labels, optimize_category_axis
from .plots import create_plot_element, add_axis_refs, add_plot_categories
from .series import add_series_to_plot

__all__ = [
    'extract_axis_ids',
    'create_value_axis',
    'optimize_axis_labels',
    'optimize_category_axis',
    'create_plot_element',
    'add_axis_refs',
    'add_plot_categories',
    'add_series_to_plot',
]

