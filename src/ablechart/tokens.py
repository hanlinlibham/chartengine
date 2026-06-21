"""配色 Token 唯一权威（颜色一致性的单一真源）。

引擎的核心价值 = 图表配色一致性。所有系列色都经
``StyleConfig.apply_to_series → get_color_for_series → COLOR_SCHEMES``
这一个收口；所有图表语义色（轴/网格/正负/合计/区间…）都经 ``CHART_TOKENS``。

设计师对接面（改色只动这里，无需懂 OXML 渲染）：
- 系列盘        → ``register_scheme(name, [...])`` / 直接改 ``COLOR_SCHEMES``
- 图表语义色    → ``set_chart_tokens(axis_text=..., wf_negative=...)`` / 改 ``CHART_TOKENS``

引擎自带 **3 套按用途的默认**（+ 金融预设用的 aim00）；品牌盘由上层（pptfi）
在 import 时通过 ``register_schemes`` 注入，引擎源码无需改动。
"""

import warnings
from typing import Dict, List

# ============================================================================
# 具名颜色常量（向后兼容 re-export）
# ============================================================================

# 深色系列（主色）
DARK_RED = "C00000"
DARK_GRAY = "595959"
DARK_BLUE = "0070C0"
DARK_ORANGE = "ED7D31"

# 浅色系列（辅助色）
LIGHT_RED = "FF9999"
LIGHT_GRAY = "BFBFBF"
LIGHT_BLUE = "9DC3E6"
LIGHT_ORANGE = "F4B183"

# aim00 风格（对比色：浅灰柱 + 深红线）—— 引擎金融预设专用
AIM00_LIGHT_GRAY = "C0C0C0"
AIM00_DARK_RED = "C00000"
AIM00_DARK_BLUE = "305496"
AIM00_GRAY = "808080"

# ============================================================================
# 系列盘注册表（COLOR_SCHEMES）—— 唯一真源
# ============================================================================

# 引擎自带默认：3 套按用途 + 金融预设盘。品牌盘由上层注册。
_BUILTIN_SCHEMES: Dict[str, List[str]] = {
    # 单色序列：深色优先，保证单系列折线/柱可读（历史 default/advisory 同值）
    "default":     ["1F3864", "2E9BD6", "00A398", "8C8C8C", "9DC3E6", "C9A84C", "5C7A93", "404040"],
    # 多系列分类：灰+青为主对，distinct hues，适合堆叠/分组多系列
    "categorical": ["595959", "29ABE2", "F5821F", "1F3864", "7B5EA7", "6BA43A", "00838F", "A6A6A6"],
    # gtm：引擎旗舰图案（contribution/range 等 GTM 模式）的市场指南配色，与 categorical 同源
    "gtm":         ["595959", "29ABE2", "F5821F", "1F3864", "7B5EA7", "6BA43A", "00838F", "A6A6A6"],
    # 涨跌正负：绿/红双极交替（数据含符号时用），可被 set_chart_tokens 派生覆盖
    "diverging":   ["588157", "B0413E", "8FB98A", "D98C8C", "2E7D32", "A83A38", "BFBFBF", "595959"],
    # aim00：引擎金融预设（presets.py）专用——浅灰柱 + 深红线
    "aim00":       [AIM00_LIGHT_GRAY, AIM00_DARK_RED, AIM00_DARK_BLUE, AIM00_GRAY],
}

# 运行时注册表（可变 dict）。上层 register_scheme(s) 往这里加品牌盘。
# 深拷贝每个色列表，避免 in-place 改 COLOR_SCHEMES[name] 污染 _BUILTIN_SCHEMES（restore 源）。
COLOR_SCHEMES: Dict[str, List[str]] = {k: list(v) for k, v in _BUILTIN_SCHEMES.items()}

# 唯一兜底：未知方案名一律落到 default（消除"两个打架的 default"）
DEFAULT_COLOR_SEQUENCE: List[str] = COLOR_SCHEMES["default"]

_warned_unknown: set = set()


def register_scheme(name: str, palette: List[str]) -> None:
    """注册（或覆盖）一个系列盘。palette = 6 位十六进制 RGB 列表。"""
    COLOR_SCHEMES[str(name)] = [c.lstrip("#").upper() for c in palette]


def register_schemes(mapping: Dict[str, List[str]]) -> None:
    """批量注册系列盘。"""
    for name, palette in dict(mapping).items():
        register_scheme(name, palette)


def get_scheme(name: str) -> List[str]:
    """取系列盘；未知名 → 落到 default 并告警一次（不静默）。"""
    if name in COLOR_SCHEMES:
        return COLOR_SCHEMES[name]
    if name not in _warned_unknown:
        _warned_unknown.add(name)
        warnings.warn(
            f"未知配色方案 '{name}'，已回退到 'default'。"
            f"可用: {', '.join(sorted(COLOR_SCHEMES))}",
            stacklevel=2,
        )
    return COLOR_SCHEMES["default"]


def list_schemes() -> List[str]:
    """当前注册表里的全部方案名（引擎独立=自带默认；import pptfi 后含品牌盘）。"""
    return sorted(COLOR_SCHEMES)


# ============================================================================
# 图表语义色（CHART_TOKENS）—— 轴/网格/正负/合计/区间…的唯一真源
# ============================================================================

_BUILTIN_CHART_TOKENS: Dict[str, str] = {
    # 坐标轴 / 网格（polish）
    "axis_text": "595959",
    "axis_line": "BFBFBF",
    "gridline": "E8E8E8",
    "zero_axis": "404040",      # 数据跨零时的深色零轴线
    # 标题 / 单位行（polish）
    "title": "262626",
    "subtitle": "595959",
    # 注释层（annotations）
    "annotation_line": "29ABE2",
    "band": "BFBFBF",
    # 瀑布（waterfall）
    "wf_positive": "588157",
    "wf_negative": "B0413E",
    "wf_total": "1F3864",
    "wf_connector": "A6A6A6",
    # 估值区间（range_chart）
    "range_band": "595959",
    "range_avg": "6BA43A",
    "range_current": "29ABE2",
    # 散点（scatter）默认色
    "scatter_default": "1F3864",
    # 贡献分解（contribution）：合计/净值橙线
    "contribution_line": "F5821F",
}

# contribution 分项堆叠配色（跳过橙色，橙色专属合计线）。列表型 token 单列管理。
_BUILTIN_CHART_PALETTE_TOKENS: Dict[str, List[str]] = {
    "contribution_parts": ["595959", "29ABE2", "1F3864", "7B5EA7", "6BA43A", "00838F", "8C8C8C", "A6A6A6"],
}
CHART_PALETTE_TOKENS: Dict[str, List[str]] = {k: list(v) for k, v in _BUILTIN_CHART_PALETTE_TOKENS.items()}


def get_chart_palette(key: str) -> List[str]:
    """取列表型图表 token（如 contribution_parts）。"""
    return CHART_PALETTE_TOKENS[key]

CHART_TOKENS: Dict[str, str] = dict(_BUILTIN_CHART_TOKENS)


def set_chart_tokens(**overrides: str) -> None:
    """覆盖图表语义色（如 set_chart_tokens(axis_text="FF0000", wf_negative="C00000")）。"""
    for key, value in overrides.items():
        if value is not None:
            CHART_TOKENS[key] = str(value).lstrip("#").upper()


def get_chart_token(key: str, default: str = None) -> str:
    """取图表语义色（调用时读取，使运行时覆盖能生效）。"""
    return CHART_TOKENS.get(key, default)


def reset_tokens() -> None:
    """恢复引擎自带默认（清掉所有注册/覆盖）。主要用于测试。"""
    COLOR_SCHEMES.clear()
    COLOR_SCHEMES.update({k: list(v) for k, v in _BUILTIN_SCHEMES.items()})
    CHART_TOKENS.clear()
    CHART_TOKENS.update(_BUILTIN_CHART_TOKENS)
    CHART_PALETTE_TOKENS.clear()
    CHART_PALETTE_TOKENS.update({k: list(v) for k, v in _BUILTIN_CHART_PALETTE_TOKENS.items()})
    _warned_unknown.clear()
