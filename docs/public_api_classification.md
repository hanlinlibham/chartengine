# Public API Classification (ADR-0007 §1)

> ADR-0007 §1 工程约束:每个进入 ``__init__.py`` 的 public name 必须能说出
> 它属于 ``create / parse / metadata / inspect / replace`` 中哪一类。本文档
> 是 PCE-006 audit 的产物(2026-05-24,123 个 public name)。
>
> **保持同步**:`__init__.py` `__all__` 任何增删 → 同步更新本文件。出现
> 不能归入任一类的 public name 时,考虑(a)增补新生命周期类(需 ADR);
> (b)把该 name 降为 internal (`_*` 前缀,从 `__all__` 移除)。

## 当前覆盖摘要 (Updated 2026-05-24: metadata 4 公开 — PCE-009 收口后)

| 类别 | 数量 | 备注 |
|---|---|---|
| **create** | 26 | family + combo + scatter/bubble + waterfall + range_snapshot |
| **parse** | 23 | parse_* + restore_*_dataframe + get_*_spec + ParseResult dataclasses |
| **metadata** | **4** ✅ | gap closed by PCE-009: `ChartMetadataV1` + `write_chart_metadata` + `METADATA_SHEET_NAME` + `METADATA_SCHEMA_VERSION` (layer 1 only;layer 2 via `semantic_anchor.py`,layer 3/4 pending PRD) |
| **inspect** | 3 | inspect_pptx_charts + ChartInventoryItem + ChartSelector |
| **replace** | 3 | replace_pptx_chart_data + ReplaceResult + SeriesData |
| **support** | 68 | configs / themes / presets / registries / cleanup utilities |
| **合计** | **127** | (`len(pptchartengine.__all__) == 127`) |

## 1. CREATE — chart 生成

从结构化 spec / DataFrame 生成原生 PowerPoint chart。

```
create_combo_chart                    # api.py
create_scatter_chart                  # scatter.py
create_bubble_chart                   # scatter.py
create_waterfall_chart                # waterfall.py
create_range_snapshot_chart           # range_snapshot.py
# semantic_family (20):
create_attribution_decomposition_chart
create_award_timeline_panel
create_concentration_chart
create_distribution_history_chart
create_distribution_snapshot_chart
create_dual_chart_panel
create_event_timeline_chart
create_factor_attribution_panel
create_factor_exposure_chart
create_heatmap_matrix_chart
create_holding_detail_panel
create_manager_timeline_profile
create_performance_compare_chart
create_ranked_tile_matrix_chart
create_regime_table_panel
create_score_overlay_chart
create_selection_timing_grid
create_semantic_chart                  # dispatcher
create_style_allocation_chart
create_style_box_chart
create_table_plus_chart_composite
```

## 2. PARSE — chart 反解析、数据恢复

从 `.pptx` / chart part 恢复 chart data, series config, layout info, metadata。

```
parse_chart_from_pptx                  # parser.py
parse_all_charts_from_pptx
parse_semantic_component_from_pptx
parse_all_semantic_components_from_pptx
parse_semantic_chart_from_layout_info
parse_scatter_chart                    # scatter.py
parse_scatter_from_pptx
parse_bubble_chart
parse_bubble_from_pptx
parse_waterfall_chart                  # waterfall.py
parse_waterfall_from_pptx
parse_range_snapshot_chart             # range_snapshot.py
parse_range_snapshot_from_pptx
ChartParser                            # parser.py — class
# Parse result dataclasses (4):
SemanticChartParseResult
ScatterParseResult
WaterfallParseResult
RangeSnapshotParseResult
# Data recovery helpers (3):
restore_range_snapshot_dataframe
restore_waterfall_dataframe
# Spec recovery from layout_info (3):
get_semantic_chart_spec
get_range_snapshot_spec
get_waterfall_spec
```

## 3. METADATA — chart-level semantic metadata persistence (PCE-009 ✓)

```
ChartMetadataV1                        # metadata.py — schema dataclass
write_chart_metadata                   # metadata.py — public writer
METADATA_SHEET_NAME                    # metadata.py — authoritative constant
METADATA_SCHEMA_VERSION                # metadata.py — schema version
```

**Implementation note**: layer 1 (embedded workbook hidden sheet) is the
default backend. Layer 2 (semantic_anchor for shape-composition families)
lives in `semantic_anchor.py` and is not in this module's public surface —
those families have no chart container so layer 1 doesn't apply. Layers
3+ (custom XML part / shape alt text) require ADR-0007 §3 PRD before
implementation. `_write_embedded_metadata` is kept as backward-compat
alias (signature preserved) so `api.py` / `scatter.py` / `semantic_family.py`
need no code change.

## 4. INSPECT — chart inventory(无 mutation)

```
inspect_pptx_charts                    # inspect.py
ChartInventoryItem
ChartSelector
```

## 5. REPLACE — chart 原位数据替换(保留 shape identity)

```
replace_pptx_chart_data                # replace.py
ReplaceResult
SeriesData
```

## 6. SUPPORT — 服务于以上五类(不属于生命周期,但 public)

ADR-0007 §1 五类生命周期描述的是 *能力* 维度;以下 name 是 *支撑* 维度
(配置、预设、主题、注册表、清理工具),它们存在的目的是让上面五类好用。

### 6.1 Configuration dataclass types

```
ChartLayoutConfig                      # layout.py
CategoryAxisConfig
ValueAxisConfig
LegendConfig
DateAxisConfig                         # date_axis.py
StyleConfig                            # styles.py
```

### 6.2 Default configuration instances

```
DEFAULT_CATEGORY_AXIS_CONFIG
DEFAULT_VALUE_AXIS_CONFIG
DEFAULT_LEGEND_CONFIG
DEFAULT_STYLE_CONFIG
```

### 6.3 Date-axis tick presets

```
DAILY_TICKS / WEEKLY_TICKS / BIWEEKLY_TICKS / MONTHLY_TICKS
QUARTERLY_TICKS / YEARLY_TICKS
```

### 6.4 Color constants & schemes

```
DARK_BLUE / DARK_GRAY / DARK_ORANGE / DARK_RED
LIGHT_BLUE / LIGHT_GRAY / LIGHT_ORANGE / LIGHT_RED
COLOR_SCHEMES                          # dict
```

### 6.5 Family identifier constants (string registry keys)

```
ATTRIBUTION_DECOMPOSITION_FAMILY / AWARD_TIMELINE_PANEL_FAMILY
CONCENTRATION_FAMILY / DISTRIBUTION_PLUS_HISTORY_FAMILY
DUAL_CHART_PANEL_FAMILY / EVENT_TIMELINE_FAMILY
FACTOR_ATTRIBUTION_PANEL_FAMILY / FACTOR_EXPOSURE_FAMILY
HEATMAP_MATRIX_FAMILY / HOLDING_DETAIL_FAMILY
MANAGER_TIMELINE_PROFILE_FAMILY / PERFORMANCE_COMPARE_FAMILY
RANKED_TILE_MATRIX_FAMILY / REGIME_TABLE_PANEL_FAMILY
SCORE_OVERLAY_FAMILY / SELECTION_TIMING_GRID_FAMILY
STYLE_ALLOCATION_FAMILY / STYLE_BOX_FAMILY
TABLE_PLUS_CHART_COMPOSITE_FAMILY
```

### 6.6 Registries / catalogs

```
SEMANTIC_FAMILY_REGISTRY               # dict[family_id → create_func]
list_semantic_families                 # function
CHART_PRESET_FUNCTIONS                 # dict
FINANCE_PRESET_FUNCTIONS               # dict
RANGE_SNAPSHOT_PRESET_FUNCTIONS        # dict
VERTICAL_VALUATION_PRESET_FUNCTIONS    # dict
```

### 6.7 Preset factories (build ready-to-use chart config)

```
get_chart_config / get_chart1_config / get_chart2_config / get_chart3_config / get_chart4_config
get_asx200_sector_valuation_snapshot_preset
get_msci_emu_sector_valuation_snapshot_preset
get_msci_japan_sector_valuation_snapshot_preset
get_sp500_sector_valuation_snapshot_preset
get_vertical_global_valuation_snapshot_preset
get_vertical_sector_valuation_snapshot_preset
```

### 6.8 Spec builders (assemble create_* input)

```
build_range_snapshot_preset
build_range_snapshot_spec
build_waterfall_spec
```

### 6.9 DataFrame helpers (data prep for create_*)

```
prepare_range_snapshot_dataframe
prepare_waterfall_dataframe
```

### 6.10 Chart post-processing utility

```
ChartJunkCleaner                       # cleaner.py
clean_chart
```

---

## Audit observations

1. **Lifecycle balance is healthy**: create (26) + parse (23) is the bulk;
   inspect (3) and replace (3) are intentionally small public surfaces.
2. **metadata = 0 public**: the only structural gap. Closing via
   `pptchartengine.metadata` module will raise the count and is the right
   trigger for PCE-LIM-003 work.
3. **support is ~55% of __all__** (68 / 123). Acceptable for a kernel that
   serves multiple upper-layer consumers (`pptfi`, `ablemind` tools, MCP)
   needing presets / themes / registries.
4. **No "uncategorisable" name** in current `__all__` — every entry fits
   one of the six buckets. ADR-0007 §1 工程约束 satisfied.
