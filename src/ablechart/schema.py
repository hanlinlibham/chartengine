"""Chart Spec 的 JSON Schema — 面向模型结构化输出（constrained decoding）。

把 spec 固化成 JSON Schema 后，可直接用于：

- Claude tool use 的 ``input_schema``（模型被 API 层强制产出合法 spec）
- OpenAI ``response_format: json_schema`` / 任何支持 schema 约束解码的推理栈
- 产出后再过 ``validate_spec()`` 做语义校验（列名存在性等 schema 管不到的部分）

设计原则：
1. **约束有价值的部分**：chart/type/axis/legend/grouping 等用 enum 锁死到
   规范值（结构化输出下模型不可能再拼错），数值/列名留自由度。
2. **不约束容错层已兜底的部分**：``additionalProperties: true``，未知字段
   照旧只警告；schema 是"地板"而不是"天花板"。
3. **enum 用规范值而非全部别名**：结构化输出场景模型按 schema 生成，不需要
   中文别名（别名层继续服务手写/弱约束场景）。
"""

from __future__ import annotations

from typing import Any, Dict


def _series_item() -> Dict[str, Any]:
    return {
        "anyOf": [
            {"type": "string", "description": "列名（默认柱状、主轴）"},
            {
                "type": "object",
                "properties": {
                    "column": {"type": "string", "description": "数据列名"},
                    "name": {"type": "string", "description": "图例显示名"},
                    "type": {"enum": ["bar", "line", "area"]},
                    "axis": {"enum": ["left", "right"]},
                    "stacked": {"type": "boolean"},
                    "color": {"type": "string", "description": "#RRGGBB"},
                    "line_width": {"type": "number", "description": "pt"},
                    "marker": {"enum": ["none", "circle", "square", "diamond", "triangle"]},
                    "labels": {
                        "anyOf": [
                            {"type": "boolean"},
                            {"type": "object", "properties": {
                                "format": {"type": "string"},
                                "position": {"enum": ["outside", "inside", "center"]},
                                "color": {"type": "string"},
                            }, "additionalProperties": True},
                        ],
                        "description": "柱上数值标签",
                    },
                    "last_point_label": {
                        "anyOf": [
                            {"type": "boolean"},
                            {"type": "object", "properties": {
                                "format": {"type": "string"},
                                "above": {"type": "boolean"},
                            }, "additionalProperties": True},
                        ],
                        "description": "末点圆点+「日期 数值」标注",
                    },
                },
                "required": ["column"],
                "additionalProperties": True,
            },
        ]
    }


def _annotation_item() -> Dict[str, Any]:
    return {
        "type": "object",
        "properties": {
            "type": {"enum": ["average", "hline", "band", "vline", "vband", "last_point"]},
            "value": {"type": "number", "description": "hline/average 的 y 值；average 给 series 时可省略（引擎算均值）"},
            "series": {"type": "string", "description": "引用的数据列：average 算均值 / band 配 quantiles 算分位 / last_point 的目标系列"},
            "label": {"type": "string"},
            "label_at": {"enum": ["left", "center", "right"]},
            "from": {"type": "number"},
            "to": {"type": "number"},
            "quantiles": {"type": "array", "items": {"type": "number"}, "minItems": 2, "maxItems": 2},
            "category": {"type": "string", "description": "vline/vband 锚定的类目值"},
            "from_category": {"type": "string"},
            "to_category": {"type": "string"},
            "index": {"type": "integer"},
            "color": {"type": "string"},
            "style": {"enum": ["dashed", "solid"]},
            "format": {"type": "string"},
            "axis": {"enum": ["left", "right"]},
        },
        "required": ["type"],
        "additionalProperties": True,
    }


def chart_spec_schema() -> Dict[str, Any]:
    """返回 chart spec 的 JSON Schema（draft-07 兼容）。"""
    from .styles import COLOR_SCHEMES

    axis_config = {
        "type": "object",
        "properties": {
            "format": {"type": "string", "description": '如 "0%" / "#,##0" / "0\\"x\\"" 或别名 percent/times'},
            "min": {"type": "number"},
            "max": {"type": "number"},
            "unit": {"type": "number"},
            "gridlines": {"type": "boolean"},
        },
        "additionalProperties": True,
    }

    return {
        "$schema": "http://json-schema.org/draft-07/schema#",
        "title": "ablechart chart spec",
        "type": "object",
        "properties": {
            "chart": {
                "enum": ["combo", "bar", "line", "area", "contribution",
                         "waterfall", "range", "scatter", "bubble"],
                "description": "图表类型；省略=combo（柱状）",
            },
            "title": {"type": "string"},
            "subtitle": {"type": "string", "description": "灰色单位行（GTM 面板惯例）"},
            "data": {
                "anyOf": [
                    {"type": "string", "description": "CSV/XLSX 路径"},
                    {"type": "object", "description": "dict-of-lists：{列名: [值...]}"},
                    {"type": "array", "description": "records：[{列: 值}...]"},
                ],
            },
            "categories": {"type": "string", "description": "X 轴列名；省略自动推断"},
            "series": {"type": "array", "items": _series_item()},
            "stacked": {"type": "boolean"},
            "grouping": {"enum": ["stacked", "percent_stacked", "clustered"]},
            "orientation": {"enum": ["vertical", "horizontal"]},
            "sort": {"enum": ["asc", "desc"]},
            "legend": {
                "anyOf": [
                    {"enum": ["bottom", "top", "left", "right", "corner", "none"]},
                    {"type": "object", "additionalProperties": True},
                ],
            },
            "style": {
                "type": "object",
                "properties": {
                    "theme": {"enum": sorted(COLOR_SCHEMES)},
                    "colors": {"type": "array", "items": {"type": "string"}},
                    "line_width": {"type": "number"},
                    "marker": {"type": "string"},
                },
                "additionalProperties": True,
            },
            "layout": {
                "type": "object",
                "properties": {
                    "legend": {"anyOf": [{"type": "string"}, {"type": "object", "additionalProperties": True}]},
                    "y_axis": axis_config,
                    "y2_axis": axis_config,
                    "x_axis": {
                        "type": "object",
                        "properties": {
                            "date_format": {"type": "string"},
                            "max_ticks": {"type": "integer"},
                            "interval": {"type": "integer"},
                        },
                        "additionalProperties": True,
                    },
                },
                "additionalProperties": True,
            },
            "annotations": {"type": "array", "items": _annotation_item()},
            "highlight": {
                "type": "object",
                "properties": {
                    "category": {"type": "string"},
                    "color": {"type": "string"},
                    "series": {"type": "string"},
                },
                "required": ["category"],
                "additionalProperties": True,
            },
            "forecast_from": {"type": "string", "description": "该类目起柱体斜纹+分隔虚线"},
            "forecast_label": {"type": "string"},
            # contribution
            "total": {"type": "string", "description": "contribution：合计列 → 橙色折线"},
            # waterfall
            "values": {"type": "string"},
            "measures": {"type": "string"},
            "totals": {"type": "array", "items": {"type": "string"},
                       "description": "waterfall：合计类目列表"},
            # range
            "low": {"type": "string"},
            "high": {"type": "string"},
            "current": {"type": "string"},
            "average": {"type": "string"},
            "format": {"type": "string", "description": "range 值轴格式，如 times"},
            # scatter / bubble
            "x": {"type": "string"},
            "y": {"type": "string"},
            "size": {"anyOf": [{"type": "string"}, {"type": "array"}]},
            "name": {"type": "string"},
            # 几何
            "position": {"type": "array", "items": {"type": "number"}, "minItems": 2, "maxItems": 2},
            "chart_size": {"type": "array", "items": {"type": "number"}, "minItems": 2, "maxItems": 2},
            "polish": {"type": "boolean"},
        },
        "additionalProperties": True,
    }
