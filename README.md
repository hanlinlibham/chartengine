# pptchartengine

`pptchartengine` 是从当前 `ppt-st` 项目中提取出来的图表核心包，定位为：

- 原生可编辑 `.pptx` 图表引擎
- 金融时间序列友好的双轴组合图
- 支持 `bar` / `line` / `area` 组合
- 支持 `clustered` / `stacked` / `percent_stacked` 分组模式（当前以 bar/area 为主）
- 支持日期轴控制、图表解析、金融预设配置

## 当前范围

已提取的核心模块：

- `create_combo_chart()`：创建可编辑双轴组合图
- `ChartBuilder` / `oxml`：底层图表 XML 生成能力
- `ChartParser`：从现有 PPT 图表反向提取配置
- `StyleConfig` / `ChartLayoutConfig` / `DateAxisConfig`
- `presets.py`：4 个金融友好的预设图表配置
- `create_waterfall_chart()`：第一版瀑布桥图（stacked bridge）

## 示例

```python
import pandas as pd
from pptx import Presentation
from pptchartengine import create_combo_chart

df = pd.DataFrame(
    {
        "日期": pd.date_range("2024-01-01", periods=12, freq="M"),
        "指数": [3000, 3050, 3100, 3150, 3200, 3250, 3300, 3320, 3350, 3380, 3400, 3450],
        "收益率": [0.01, 0.02, 0.03, 0.025, 0.04, 0.05, 0.048, 0.055, 0.06, 0.058, 0.065, 0.07],
    }
)

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

create_combo_chart(
    slide=slide,
    df=df,
    categories_col="日期",
    series_config=[
        {"key": "指数", "name": "指数", "type": "bar", "axis": "secondary"},
        {"key": "收益率", "name": "收益率", "type": "line", "axis": "primary"},
    ],
)

prs.save("demo.pptx")
```

## 预设

当前导出的金融预设：

- `get_chart1_config`
- `get_chart2_config`
- `get_chart3_config`
- `get_chart4_config`

对应场景集中在养老金/收益率/仓位/久期类报告图。

## 命名建议

- 核心包：`pptchartengine`
- 上层 skill：`pptfi`
