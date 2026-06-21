"""场景化最小 spec 示例库 — 给 LLM 的 few-shot 弹药。

对能力较弱的模型，「按场景检索一个能跑的最小示例」远比读完整文档可靠。
用法：把 ``chart_spec_examples()`` 的输出拼进提示词，或按场景取单条::

    from ablechart import chart_spec_examples
    prompt += chart_spec_examples("估值分位")     # 单场景
    prompt += chart_spec_examples()               # 全部

每条示例都保证：字段最少、可直接渲染、命中引擎的智能默认值。
"""

from __future__ import annotations

import json
from typing import Optional

# (场景关键词, 说明, 最小 spec)
SCENARIO_EXAMPLES = [
    ("基础柱状图|排名|对比", "只传数据，类目/系列/类型全自动推断", {
        "title": "分公司营收",
        "data": {"分公司": ["华东", "华南", "华北"], "营收(亿)": [128, 96, 87]},
    }),
    ("量价|双轴|股价成交量", "第二个系列指定 折线+右轴 即可，中文别名均可", {
        "title": "营收与净利率",
        "data": {"季度": ["24Q1", "24Q2", "24Q3"], "营收": [100, 110, 120], "净利率": [0.12, 0.13, 0.15]},
        "series": ["营收", {"column": "净利率", "type": "折线", "axis": "右轴"}],
        "layout": {"y2_axis": {"format": "percent"}},
    }),
    ("净值曲线|走势|时间序列", "日期列自动启用日期轴并抽稀标签；末点标注一个开关", {
        "chart": "line",
        "title": "组合净值走势",
        "data": "nav.csv",
        "legend": "none",
        "series": [{"column": "净值", "last_point_label": {"format": "0.000"}}],
    }),
    ("收入结构|堆叠", "stacked 一个开关，全部数值列自动堆叠", {
        "title": "业务线收入结构",
        "stacked": True,
        "data": {"年份": ["2023", "2024"], "企业年金": [55, 63], "职业年金": [33, 40]},
    }),
    ("占比演变|百分比堆叠|资产配置", "grouping 写 '100%' 即百分比堆叠", {
        "title": "资产配置演变",
        "grouping": "100%",
        "data": {"季度": ["24Q4", "25Q1"], "股票": [24, 26], "债券": [56, 54], "现金": [20, 20]},
    }),
    ("增长归因|贡献分解|GDP分解", "chart=contribution + total 指定合计列，其余自动", {
        "chart": "contribution",
        "title": "营收增速贡献分解",
        "total": "营收同比",
        "data": {"季度": ["24Q1", "24Q2"], "利息": [0.04, 0.03], "手续费": [0.01, 0.02], "营收同比": [0.05, 0.05]},
    }),
    ("收益归因|瀑布|桥图", "totals 列出合计类目即可，正负配色自动", {
        "chart": "瀑布图",
        "title": "年度收益归因（bp）",
        "data": {"项目": ["期初", "配置", "选股", "成本", "期末"], "贡献": [420, 85, 112, -38, 579]},
        "totals": ["期初", "期末"],
    }),
    ("估值分位|历史区间|PE区间", "chart=range，low/high/average/current 四列", {
        "chart": "range",
        "title": "行业PE：当前 vs 十年区间",
        "data": {"行业": ["白酒", "银行"], "低": [18, 4], "高": [55, 9], "均值": [32, 6], "当前": [24, 5]},
        "low": "低", "high": "高", "average": "均值", "current": "当前",
    }),
    ("业绩排名|横向条形图", "orientation=horizontal；labels 加数值标签；highlight 高亮基准", {
        "title": "近一年收益率排名",
        "orientation": "horizontal",
        "legend": "none",
        "data": {"产品": ["A", "基准", "B"], "收益": [0.17, 0.073, -0.038]},
        "series": [{"column": "收益", "labels": {"format": "0.0%"}}],
        "highlight": {"category": "基准"},
    }),
    ("盈利预测|一致预期|预测期", "forecast_from 指定预测起点 → 斜纹+分隔线自动", {
        "title": "EPS增速与一致预期",
        "legend": "none",
        "data": {"年度": ["2024", "2025E", "2026E"], "EPS增速": [0.26, 0.21, 0.16]},
        "series": [{"column": "EPS增速", "labels": {"format": "0%"}}],
        "forecast_from": "2025E",
    }),
    ("均值线|目标区间|监测图", "average 给 series 引擎自己算均值；band 可给分位数自动算", {
        "chart": "line",
        "title": "信用利差监测",
        "data": "spread.csv",
        "legend": "none",
        "series": [{"column": "利差", "last_point_label": True}],
        "annotations": [
            {"type": "average", "series": "利差", "label": "区间均值"},
            {"type": "band", "series": "利差", "quantiles": [0.25, 0.75], "label": "正常区间"},
        ],
    }),
    ("风险收益|散点|气泡", "数值列按顺序自动当 x/y/size", {
        "chart": "bubble",
        "title": "基金风险收益分布",
        "data": {"波动率": [8.1, 9.2], "收益率": [10.5, 12.0], "规模": [50, 80]},
    }),
]


def chart_spec_examples(scenario: Optional[str] = None) -> str:
    """返回场景化最小 spec 示例（markdown）。scenario 支持关键词模糊匹配。"""
    rows = SCENARIO_EXAMPLES
    if scenario:
        token = str(scenario).strip().lower()
        matched = [r for r in rows if token in r[0].lower() or token in r[1].lower()]
        rows = matched or rows

    parts = ["# Chart Spec 场景示例（最小可用）\n"]
    for keywords, note, spec in rows:
        parts.append(f"## {keywords.split('|')[0]}（{note}）")
        parts.append("```json")
        parts.append(json.dumps(spec, ensure_ascii=False, indent=2))
        parts.append("```\n")
    return "\n".join(parts)
