# Chart Spec 参考

`render_chart(slide, spec)` 的声明式接口：**一个 JSON dict 进、可编辑图表出**。

设计给两类调用方：
1. **LLM**（尤其是能力较弱的小模型）：只要会吐 JSON 就能画图，不需要会写 Python、不需要 import 配置类、不需要 pptx 枚举。把本文（或 `chart_spec_reference()` 的输出）拼进提示词即可。
2. **人**：少记 API，错了有"是否想用 xxx"的提示。

## 核心 API

```python
from ablechart import render_chart, validate_spec, chart_spec_reference

chart = render_chart(slide, spec)            # 渲染（默认静默，quiet=False 打开引擎日志）
chart = render_chart(slide, spec, df=df)     # 数据也可以从外部 DataFrame 传入

report = validate_spec(spec)                  # 只校验不渲染 → {ok, errors, warnings, normalized}
print(chart_spec_reference())                 # spec 速查表，可拼进 LLM 提示词
```

校验失败抛出 `SpecError`，**一次性列出全部问题**，列名拼错会给出建议：

```
Chart spec 校验失败 (2 处):
  1. series[0].column: 找不到列 '利浬' — 是否想用 '利润'? 可用列: 年份, 营收, 利润, 增速
  2. series[1].axis: 无法识别 'middle' — 允许的值如: 1/2/first/l/left/main/primary/r/right/second/secondary/y
```

## 容错规则（全局）

- 字段值大小写不敏感，`-`/空格/`_` 等价
- 类型、轴、图例位置、grouping 均接受中英文别名
- 颜色接受 `#C00000` / `C00000` / `#f00` / `red` / `红`
- 未知字段**只警告不报错**（带"是否想用 xxx"提示）
- 数字格式接受别名：`percent`→`0%`、`thousands`→`#,##0`、`decimal2`→`0.00`，或直接传 Excel 格式码

## 智能推断（缺什么补什么）

| 缺省字段 | 推断规则 |
|---------|---------|
| `chart` | `combo` |
| `categories` | 第一个非数值列（或日期列） |
| `series` | 除 categories 外的全部数值列，默认柱状图 |
| 日期轴 | categories 是日期类型 → 自动日期轴 + 自动标签间隔（约 7 个刻度） |
| waterfall `values` | 第一个数值列 |
| waterfall `measures` | 自动识别取值全为 relative/total 的列 |
| scatter/bubble `x`/`y`/`size` | 按顺序取数值列 |

每次推断都会写进 `warnings`，调用方可见。

## 数据形态（`data` 字段）

```jsonc
{"年份": [2021, 2022], "营收": [100, 110]}                  // dict-of-lists
[{"年份": 2021, "营收": 100}, {"年份": 2022, "营收": 110}]   // records
{"columns": ["年份", "营收"], "rows": [[2021, 100]]}         // columns/rows
"data/quarterly.csv"                                         // CSV / XLSX 路径
```

省略 `data` 时使用 `render_chart(slide, spec, df=...)` 传入的 DataFrame。

## combo（柱 / 线 / 面积，单轴或双轴）

```jsonc
{
  "chart": "combo",                  // 可省略；bar/line/area/双轴 等别名均可
  "title": "营收与增速",
  "data": {...},
  "categories": "年份",
  "series": [
    "营收",                           // 字符串 = 默认柱状、主轴
    {"column": "增速", "type": "line", "axis": "right",   // type: bar|line|area（含中文别名）
     "color": "#C00000", "line_width": 2, "marker": "circle"}  // 逐系列覆盖
  ],
  "stacked": true,                   // 便捷开关：全部柱状系列堆叠
  "grouping": "percent",             // 或百分比堆叠
  "style": {
    "theme": "able_finance",           // 10 套主题
    "colors": ["#1B3D6E", "gold"],   // 自定义调色板（优先于 theme）
    "line_width": 1.5, "marker": "none"
  },
  "layout": {
    "legend": "top",                 // bottom|top|left|right|corner|none，或 dict
    "y_axis":  {"format": "percent", "min": 0, "max": 1, "unit": 0.2, "gridlines": false},
    "y2_axis": {"format": "#,##0"},
    "x_axis":  {"date_format": "yyyy/mm", "max_ticks": 7}   // 或 {"interval": 5}
  },
  "position": [1, 2],                // 英寸；也接受 "2.5cm" / "36pt"
  "chart_size": [8, 4.5]
}
```

说明：
- `x_axis.date_format` 会把字符串类目自动转成日期（转不动则回退普通分类轴并警告）
- `x_axis.max_ticks` / `interval` 在普通分类轴上同样生效（tickLblSkip）
- 逐系列 `color`/`line_width`/`marker` 覆盖在主题配色之上，只改指定的系列

## waterfall（瀑布桥图）

```jsonc
{
  "chart": "waterfall",              // bridge / 瀑布图 / 桥图 均可
  "title": "收益归因",
  "data": {"阶段": ["期初", "权益", "债券", "汇率", "期末"],
            "贡献": [8.5, 2.1, 1.3, -1.8, 10.1]},
  "categories": "阶段",
  "values": "贡献",
  "totals": ["期初", "期末"],         // 合计类目；或用 measures 列（relative/total）
  "colors": {"positive": "#1A5C2A", "negative": "#C00000", "total": "#1B3D6E"},
  "value_labels": true,              // 默认 true
  "connectors": false                // 默认 false（Windows 上有已知偏移）
}
```

## scatter / bubble

```jsonc
{
  "chart": "bubble",                 // 散点图 / 气泡图 均可
  "data": {"波动率": [8.1, 9.2], "收益率": [10.5, 12.0], "规模": [50, 80]},
  "x": "波动率", "y": "收益率", "size": "规模",   // bubble 才需要 size
  "name": "风险收益分布",
  "color": "#1E2761", "marker_size": 9
}
```

## range（估值区间图）— GTM 估值页标志图型

```jsonc
{
  "chart": "range",                  // 区间 / 估值区间 均可
  "data": {"行业": [...], "十年最低": [...], "十年最高": [...], "十年均值": [...], "当前PE": [...]},
  "low": "十年最低", "high": "十年最高",
  "average": "十年均值", "current": "当前PE"   // 可选
}
```

灰色浮动条 = 历史区间，绿色横杠 = 均值，蓝色菱形 = 当前。适用 PE/PB 历史分位、
利差区间、波动率锥、仓位历史区间等场景。全部原生可编辑。

## contribution（贡献分解）— GTM 宏观分解图标配

```jsonc
{"chart": "contribution", "total": "GDP同比", "data": {...}}
```

合计列自动画成橙色折线，其余数值列自动正负堆叠，分项配色自动跳过橙色。

## GTM 元素（combo 通用）

```jsonc
{
  "orientation": "horizontal",                         // 横向排名条
  "series": [{"column": "ROE",
              "labels": {"format": "0.0%"},            // 柱上数值标签（原生 dLbls）
              "last_point_label": {"format": "0.0%"}}],// 末点圆点+标注
  "highlight": {"category": "全部A股", "color": "#6BA43A"},  // 单类目高亮（原生 dPt）
  "forecast_from": "2025F",                            // 该类目起斜纹填充+虚线分隔
  "annotations": [
    {"type": "average", "value": 0.024, "label": "长期均值 2.4%", "label_at": "left"},
    {"type": "band", "from": 0.02, "to": 0.03, "label": "目标区间"},
    {"type": "vline", "category": "2024Q1", "label": "政策转向"}
  ]
}
```

注释通过 manualLayout 钉定绘图区实现精确对位；数值标签/高亮/斜纹是原生
XML（dLbls/dPt），在 PowerPoint 里依旧可编辑。

## 结构化输出（schema 约束解码）

`chart_spec_schema()` 返回 spec 的 JSON Schema（draft-07，~6KB），可直接作为
Claude tool use 的 `input_schema` 或 `response_format: json_schema` —— 图表类型、
轴、图例、grouping 等枚举在解码层锁死，语法/字段层错误归零；列名等语义部分
配合 `validate_spec()` 闭环。

## 降低模型门槛的机制

- **few-shot 示例库**：`chart_spec_examples("估值")` 按场景返回最小可用 spec（12 个场景），
  直接拼进弱模型的提示词
- **注释免算数**：`{"type": "average", "series": "利差"}` 引擎算均值；
  `{"type": "band", "series": "利差", "quantiles": [0.25, 0.75]}` 引擎算分位 —— 不要让模型做算术
- **自动双轴**：series 省略且两列量纲差 >50 倍 → 小量纲列自动改右轴折线（warning 提示，显式 series 可关）

## 推荐的 LLM 工作流

```python
report = validate_spec(spec, df)     # 1. 模型产出 spec 后先自检
if not report["ok"]:
    ...                              # 2. 把 report["errors"]（含建议）喂回模型修正
chart = render_chart(slide, spec, df)  # 3. 通过后渲染
```

`validate_spec` 不产生任何副作用，`report["normalized"]` 里有归一化后的渲染计划摘要，
可用来让模型确认"引擎理解的"和"它想要的"是否一致。
