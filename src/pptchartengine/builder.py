"""
图表构建器 - 有状态的编排器

这是重构后的核心，使用有状态的构建器模式来编排 oxml 层的调用。
解决 P(n,2) 组合爆炸问题。
"""

from typing import List, Dict, Optional
from pptx.chart.chart import Chart
import pandas as pd
from collections import defaultdict

from .oxml import (
    extract_axis_ids,
    create_value_axis,
    optimize_axis_labels,
    optimize_category_axis,
    create_plot_element,
    add_axis_refs,
    add_plot_categories,
    add_series_to_plot,
)

# 导入样式模块
try:
    from .styles import DEFAULT_STYLE_CONFIG
except ImportError:
    # 如果样式模块不存在，设置为 None
    DEFAULT_STYLE_CONFIG = None


class ChartBuilder:
    """
    有状态的图表构建器
    
    职责：
    1. 管理坐标轴（主轴/次轴）
    2. 按 (type, axis) 分组系列
    3. 委托 oxml 层创建 XML 元素
    
    优势：
    - 解决 P(n,2) 组合问题：不需要为每种组合写代码
    - 只需为每种图表类型实现一次 XML 生成
    - 主函数动态组合它们
    """
    
    def __init__(self, chart: Chart, df: pd.DataFrame, categories_col: str, style_config=None, layout_config=None):
        """
        初始化构建器
        
        Args:
            chart: python-pptx 创建的基础图表（用于激活 XML 结构）
            df: 数据源 DataFrame
            categories_col: 分类列名
            style_config: 样式配置对象（可选，默认使用 DEFAULT_STYLE_CONFIG）
            layout_config: 布局配置对象（可选，包含图例、轴配置）
        """
        self.chart = chart
        self.df = df
        self.categories_col = categories_col
        self.style_config = style_config if style_config is not None else DEFAULT_STYLE_CONFIG
        self.layout_config = layout_config  # 布局配置
        
        # 访问 XML 结构
        self.chartSpace = chart._chartSpace
        self.plotArea = self.chartSpace.plotArea
        
        # 提取现有坐标轴 ID
        self.cat_ax_id, self.pri_val_ax_id = extract_axis_ids(self.plotArea)
        self.sec_val_ax_id = None
        
        # 系列计数器（用于 Excel 列索引）
        self._series_counter = 0
        
        print(f"\n📊 ChartBuilder 初始化完成")
        print(f"  - 分类轴 ID: {self.cat_ax_id}")
        print(f"  - 主值轴 ID: {self.pri_val_ax_id}")
        if self.style_config is not None:
            print(f"  - 样式配置: 已启用")
        else:
            print(f"  - 样式配置: 默认")
        if self.layout_config is not None:
            print(f"  - 布局配置: 已启用")
    
    def ensure_secondary_axis(self) -> int:
        """
        确保次值轴存在（只创建一次）
        
        Returns:
            次值轴 ID
        """
        if self.sec_val_ax_id is None:
            # 生成新的轴 ID
            new_ax_id = max(self.cat_ax_id, self.pri_val_ax_id) + 1000
            
            # ⭐ 创建次值轴（右侧，标签在右边，crosses='max'）
            self.sec_val_ax_id = create_value_axis(
                self.plotArea,
                ax_id=new_ax_id,
                cross_ax_id=self.cat_ax_id,
                position='r',
                tick_label_position='high',
                crosses_at='max'  # 右轴线在图表右边
            )
            
            # ⭐ 优化主值轴（左侧，标签在左边，crosses='min'，移除网格线）
            optimize_axis_labels(
                self.plotArea,
                self.pri_val_ax_id,
                tick_label_position='low',
                crosses_at='min',  # 左轴线在图表左边
                remove_gridlines=True  # 移除内部横框
            )
            
            # ⭐ 优化分类轴（移除日期间的小竖线）
            optimize_category_axis(
                self.plotArea,
                self.cat_ax_id,
                remove_tick_marks=True  # 移除底部日期间的小竖线
            )
            
            print(f"\n⭐ 次值轴已创建")
            print(f"  - 次值轴 ID: {self.sec_val_ax_id}")
            print(f"  - 位置: 右侧 (position='r', crosses='max')")
            print(f"  - 标签: 右侧 (tickLblPos='high')")
            print(f"  - 主值轴已优化: 左侧 (tickLblPos='low', crosses='min')")
        
        return self.sec_val_ax_id
    
    def clear_bootstrap_chart(self):
        """
        清理引导时创建的图表元素
        
        Note: 
            使用 python-pptx 创建基础图表时，会自动创建一个图表元素。
            我们需要清理它，以便完全通过 XML 自定义。
        """
        # 查找并删除 python-pptx 自动创建的图表元素
        # 这些通常是 <c:barChart>, <c:lineChart> 等
        
        # 获取所有可能的图表类型元素
        chart_types = ['barChart', 'lineChart', 'areaChart', 'scatterChart', 'bubbleChart', 'pieChart']
        
        for chart_type in chart_types:
            # 查找所有该类型的图表元素
            elements = self.plotArea.xpath(f'./c:{chart_type}')
            for elem in elements:
                # 删除这个元素
                self.plotArea.remove(elem)
                print(f"  → 清理引导图表元素: <c:{chart_type}>")
        
        print(f"  → 引导图表已清理")
    
    def add_plot(self, series_group: List[Dict], plot_order_index: int = 0):
        """
        添加一组系列（同类型、同轴）
        
        Args:
            series_group: 系列配置列表
                [
                    {"key": "col1", "name": "系列1", "type": "bar", "axis": "primary"},
                    {"key": "col2", "name": "系列2", "type": "bar", "axis": "primary"}
                ]
            plot_order_index: 绘图顺序索引（0=最底层，1=上一层，依此类推）
                
        Notes:
            - 组内所有系列必须有相同的 type 和 axis
            - 会创建一个新的绘图元素 (<c:barChart>, <c:lineChart> 等)
            - 并为每个系列添加 <c:ser> 元素
        """
        if not series_group:
            return
        
        # 从第一个系列获取共享属性
        first_cfg = series_group[0]
        chart_type = first_cfg.get("type", "bar")
        axis_type = first_cfg.get("axis", "primary")
        grouping = first_cfg.get("grouping")
        
        # 决定使用哪个值轴
        if axis_type == 'primary':
            val_ax_id = self.pri_val_ax_id
        elif axis_type == 'secondary':
            val_ax_id = self.ensure_secondary_axis()
        else:
            raise ValueError(f"未知的轴类型: {axis_type}")
        
        print(f"\n➕ 添加绘图组: type={chart_type}, axis={axis_type}, order={plot_order_index}")
        print(f"  - 使用值轴 ID: {val_ax_id}")
        
        # 创建绘图元素（不包含轴引用）
        plot_element = create_plot_element(
            self.plotArea,
            chart_type,
            self.cat_ax_id,
            val_ax_id,
            order_index=plot_order_index,  # ⭐ 传递 order_index
            grouping=grouping,
        )
        
        # ⭐ 新方案：不在 plot 级别添加共享分类数据
        # 每个系列有自己的 <c:cat> 元素，避免重复
        # categories = self.df[self.categories_col].tolist()
        # add_plot_categories(plot_element, categories)
        # print(f"  - 添加共享分类数据（{len(categories)} 个分类）")
        
        # 为每个系列添加 <c:ser>
        for series_cfg in series_group:
            add_series_to_plot(
                plot_element,
                chart_type,
                series_cfg,
                self._series_counter,
                self.df,
                self.categories_col,
                self.style_config  # ⭐ 传递样式配置
            )
            print(f"  - 添加系列: '{series_cfg['name']}' (索引 {self._series_counter})")
            self._series_counter += 1
        
        # ⭐ 关键修复：在所有系列添加完成后，再添加轴引用
        # 确保 XML 元素顺序正确：<c:cat> <c:ser> ... <c:ser> <c:axId> <c:axId>
        add_axis_refs(plot_element, self.cat_ax_id, val_ax_id)
        print(f"  - 添加轴引用: cat_ax={self.cat_ax_id}, val_ax={val_ax_id}")
    
    def build(self, series_config: List[Dict]):
        """
        构建完整的组合图
        
        Args:
            series_config: 系列配置列表
                [
                    {"key": "col1", "name": "系列1", "type": "bar", "axis": "primary"},
                    {"key": "col2", "name": "系列2", "type": "line", "axis": "secondary"},
                ]
                
        Returns:
            构建后的 Chart 对象
        """
        print("\n" + "=" * 80)
        print("🔨 开始构建组合图")
        print("=" * 80)
        
        # 1. 按 (type, axis) 分组
        plot_groups = self._group_series(series_config)
        
        print(f"\n📦 系列分组结果:")
        for key, group in plot_groups.items():
            print(f"  - {key}: {len(group)} 个系列")
        
        # 2. 清理引导图表（可选）
        self.clear_bootstrap_chart()
        
        # 3. ⭐ 按堆叠顺序添加绘图组
        # 规则：先添加的绘图组在底层，后添加的在上层
        # 策略：先画"背景"（柱状图/面积图），再画"前景"（折线图/散点图）
        plot_order_counter = 0
        
        # 3.1 先添加所有"背景"图表 (bar, area, bubble)
        for (plot_type, axis_type, grouping), series_group in plot_groups.items():
            if plot_type in ('bar', 'column', 'area', 'bubble'):
                self.add_plot(series_group, plot_order_counter)
                plot_order_counter += 1
        
        # 3.2 再添加所有"前景"图表 (line, scatter)
        for (plot_type, axis_type, grouping), series_group in plot_groups.items():
            if plot_type in ('line', 'scatter'):
                self.add_plot(series_group, plot_order_counter)
                plot_order_counter += 1
        
        # 4. ⭐ 应用布局配置（图例、轴格式等）
        if self.layout_config is not None:
            self.layout_config.apply_to_chart(self.chart)

        # 5. ⭐ ChartJunkCleaner: 自动清洗默认 PPT 样式
        try:
            from .cleaner import clean_chart
            clean_chart(self.chart)
            print(f"  → ChartJunkCleaner applied")
        except Exception as e:
            print(f"  → ChartJunkCleaner skipped: {e}")

        print("\n" + "=" * 80)
        print("✅ 组合图构建完成！")
        print("=" * 80)

        return self.chart
    
    @staticmethod
    def _group_series(series_config: List[Dict]) -> Dict[tuple, List[Dict]]:
        """
        按 (type, axis) 分组系列
        
        Args:
            series_config: 系列配置列表
            
        Returns:
            分组后的字典 {(type, axis): [series_cfg, ...]}
            
        Examples:
            输入:
            [
                {"key": "s1", "type": "bar", "axis": "primary"},
                {"key": "s2", "type": "bar", "axis": "primary"},
                {"key": "s3", "type": "line", "axis": "secondary"},
            ]
            
            输出:
            {
                ("bar", "primary"): [{"key": "s1", ...}, {"key": "s2", ...}],
                ("line", "secondary"): [{"key": "s3", ...}]
            }
        """
        groups = defaultdict(list)
        for cfg in series_config:
            key = (
                cfg.get("type", "bar"),
                cfg.get("axis", "primary"),
                cfg.get("grouping"),
            )
            groups[key].append(cfg)
        return dict(groups)
