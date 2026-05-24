# PRD-0002: Template-safe chart inspect + replace 实施

- Status: completed (archived)
- Created: 2026-05-24
- Completed: 2026-05-24
- Target: ADR-0006 §3 first-batch (bar / line / combo / area / pie / scatter / bubble)
- Related: ADR-0004, ADR-0006, ADR-0007, ADR-0005

> **流程说明**(透明记录,后续 review 用):本 PRD 反向补写。ADR-0006 §工程约束第 1 条 + ADR-0005 §2 PRD 触发条件(新 public API + metadata schema 演进)要求实施前先 PRD。实际开发走的是「直接开干」路径,实施 + 测试 + 文档化在 2026-05-24 一日内闭环。本 PRD 反映实施成果作为目的地文档,供后续 review / 引用 / 扩展(area / pie 补全 / semantic_family 接入)使用。"已完成即立即 archive" 符合 ADR-0005 §3 反文档腐烂规则。

## 问题陈述

ADR-0001 把项目定位为「专业数据驱动的可编辑 PowerPoint 更新与生成平台」,北极星场景是**模板更新**而非从零生成。这要求 `pptchartengine` 提供两个核心能力:

1. **inspect**:扫描任意 `.pptx`(包括非引擎生成的),输出 chart inventory + 稳定 selector,让上层(`pptfi` / `ablemind`)能定位 chart、判断可替换性。
2. **replace**:对已定位的 chart 原位替换数据,保留 shape identity、位置、尺寸、主题、可编辑性。

ADR-0004 round-trip 5 件套约束已经覆盖「引擎自己生成的 chart 还能 parse 回来」,但不覆盖「非引擎生成的 chart 的 inspect / replace」。ADR-0006 显式补齐这块,本 PRD 是 ADR-0006 §1/§2 的具体实施。

## 影响范围

### 影响的 ADR

- **ADR-0004**:round-trip 5 件套现多了 inspect/replace 两类 public API,需 round-trip harness 覆盖 — PCE-007 / 任务 #23 已落地 v1 矩阵
- **ADR-0006**:本 PRD 直接实施 §1 / §2 / §3(首批 7 类) / §5(8 条 invariant 中 6 条直接验证,2 条 PowerPoint open 间接保证)
- **ADR-0007**:本 PRD 引入两个新 public API,正式属于五类生命周期中的 `inspect` 和 `replace`(每类各 1 个);ChartSpec / SeriesData / ReplaceResult 保持技术层 only(无 `user_id` / 业务 slot / prompt)
- **ADR-0005**:本 PRD 自身是反向补写,流程不规范但内容补齐;后续工作应"先 PRD 再开干"恢复纪律

### 影响的 family

实施初版覆盖 ADR-0006 §3 **first-batch 7 个 chart_type**:
- 类目型(`CategoryChartData`):`bar` / `line` / `combo` / `area` / `pie`
- XY 型(`XyChartData`):`scatter`
- 三轴型(`BubbleChartData`):`bubble`

**不在本 PRD 范围**(后续 PRD 处理):
- 各 `semantic_family`(`performance_compare` / `event_timeline` / `attribution_decomposition` / `style_box` 等)— 多数是 shape composition,不走 chart_part 路径,需另外的 inspect/replace 适配
- linked external workbook / pivot chart / SmartArt / 复杂 plugin chart
- 没有 embedded workbook 的 chart(返回 `missing_embedded_workbook` fail-loud)

### 影响的依赖项目

- `python-pptx>=1.0.2,<1.1` — `chart.replace_data` API 是核心实现路径
- `pandas>=2.2,<3` — fixture / 测试用;**pin 上限**因为 pandas 3.0 breaking round-trip(详见 ISSUES.md `PCE-LIM-002`)
- `numpy>=1.26,<2` — pin 上限因为 ABI 兼容
- `lxml>=5,<6`、`openpyxl>=3.1,<4` — 显式收紧

## 解决方案概要

新增两个模块,严格遵守 ADR-0007 五类生命周期分类。

### `src/pptchartengine/inspect.py`(**inspect** lifecycle)

**Public**:
- `ChartSelector` dataclass(selector 优先级 ADR-0006 §1:`explicit_name` > `shape_id + chart_part` > `(slide_index, chart_index_on_slide)` fallback)
- `ChartInventoryItem` dataclass(技术层字段:slide_index / shape_id / chart_part / chart_type / category_count / series_count / series_names / has_embedded_workbook / replaceable / warnings)
- `inspect_pptx_charts(pptx_path) -> List[ChartInventoryItem]` — read-only;空 pptx 返回 `[]`(不抛);按 `(slide_index, chart_index_on_slide)` 排序;**不依赖 engine-written metadata**

**实现要点**:
- 用 `python-pptx` 遍历 `slide.shapes` 找 `has_chart=True` 的 shape
- chart_type 分类按 `plot.__class__.__name__`(bar / line / area / pie / scatter / bubble),多 plot 或混合 → `"combo"`
- chart_part 从 `chart.part.partname` 取(strip 前导 `/`)
- explicit_name 检测排除 python-pptx 默认 prefix(`Chart ` / `Placeholder ` / `Picture ` / ...)
- replaceable 由 `chart_type ∈ first-batch ∧ has_embedded_workbook` 决定

### `src/pptchartengine/replace.py`(**replace** lifecycle)

**Public**:
- `SeriesData` dataclass(name + values + 可选 x_values / size_values,字段需求按 chart_type 变化)
- `ReplaceResult` dataclass(status + selector_resolved + chart_part + series_replaced + categories_replaced + data_points_replaced + warnings + error_code + error_detail)
- `replace_pptx_chart_data(input_pptx, output_pptx, selector, categories, series) -> ReplaceResult`

**实现要点**(对齐 ADR-0006 §5.8 + ADR-0007 §4):
- **预校验在 mutation 前**:categories 长度 / series 字段 / chart_type 支持 / has_embedded_workbook;预校验失败必定在 `prs.save` 前 → 不产生半成品文件
- **复用 inspect** 做 chart_type / replaceable 判断,单向 import,无循环
- **按 chart_type dispatch**:`bar/line/combo/area/pie → CategoryChartData`;`scatter → XyChartData`;`bubble → BubbleChartData`
- **`chart.replace_data` 不退化为重建**:python-pptx 该 API 原位修改 chart XML + embedded workbook,不动 shape — 天然满足 ADR-0007 §4

### 错误码清单(ADR-0006 §3 + 本 PRD 扩展)

| 错误码 | 来源 | 触发条件 |
|---|---|---|
| `ambiguous_selector` | ADR-0006 §3 | selector 未匹配任何 chart |
| `unsupported_chart_type` | ADR-0006 §3 | chart_type 不在 first-batch |
| `missing_embedded_workbook` | ADR-0006 §3 | chart 无 embedded workbook |
| `category_length_mismatch` | ADR-0006 §3 | series.values 长度 ≠ categories 长度 |
| `categories_required_for_category_chart` | **本 PRD 新增** | category 型 chart 调用时 categories=None |
| `x_values_required_for_scatter` | **本 PRD 新增** | scatter chart series 缺 x_values |
| `xy_length_mismatch` | **本 PRD 新增** | scatter chart series x/y 长度不一致 |
| `size_values_required_for_bubble` | **本 PRD 新增** | bubble chart series 缺 x_values 或 size_values |
| `bubble_length_mismatch` | **本 PRD 新增** | bubble chart series x/y/size 长度不一致 |
| `series_required` | **本 PRD 新增** | series 输入为空 |

## 测试策略

按 ADR-0002 三层质量闸门映射:

### L1 — Unit / Contract tests(已覆盖,**62 / 62 GREEN**)

| 测试文件 | 数量 | 覆盖 |
|---|---|---|
| `tests/test_inspect.py` | 6 | 空 .pptx / 无 chart slides / 单 combo / 多 chart 顺序 / selector 稳定性 / read-only 不修改输入 |
| `tests/test_replace.py` | 14 | happy path / input 不修改 / shape identity / chart_type 不变 / categories 长度校验 / data 回流 / selector 回传 / scatter happy / bubble happy / scatter 缺 x_values 失败 / bubble 缺 size_values 失败 / category chart 缺 categories 失败 / 1 类原 happy / scatter_preserve / bubble_preserve |
| `tests/test_external_chart.py` | 4 | fixture 用 python-pptx 原生 API 生成(**绕开 pptchartengine code path**),验证 inspect/replace 在 non-engine-authored chart 上仍 work + sample template 0-chart sanity |
| `tests/test_round_trip_matrix.py` | 9 | parametrize 3 chart_type × 3 contract — 系统化保证 ADR-0006 §5 invariant 在每个 first-batch chart_type 上一致 |
| `tests/test_package_contract.py` | 29 | 历史 contract test(create + parse round-trip),作为 baseline 回归保护 |

### L2 — Round-trip / Cross-family safety net

- 已落地 `tests/test_round_trip_matrix.py` 作为 v1 矩阵
- 后续 `PCE-007` 任务扩展覆盖 area / pie / 平 bar / 平 line,以及 semantic_family 接入

### L3 — Visual diff / Golden(推迟到 P2 末)

- 待 ADR-0003 golden reference 启动 Slice 1(`performance_attribution`)时联调
- 本 PRD 阶段**不要求** L3 覆盖

## 参考材料

- `pptchartengine/docs/adr/0004-round-trip-metadata-principle.md` — round-trip 5 件套强约束
- `pptchartengine/docs/adr/0006-template-safe-chart-data-update.md` — chart inventory / data replacement / 首批 chart 类型 / 八条 invariant
- `pptchartengine/docs/adr/0007-chart-asset-kernel-boundary.md` — kernel 五类生命周期 / ChartSpec 技术层禁污染 / replace 不退化为重建
- `design/ppt-specialist/adr/0001-professional-data-update-ppt-platform.md` — 上层产品定位(模板更新北极星)
- `design/ppt-specialist/ablemind-integration-review.md` — ablemind 接口层评审,确认 inspect/replace tool surface 跟 capability layer 对齐
- `pptchartengine/ISSUES.md` — PCE-004 / 005 / LIM-001 / LIM-002 沉淀

## 完成定义

- [x] `inspect_pptx_charts` 公开,返回 `ChartInventoryItem` 列表,满足 ADR-0006 §1 inventory 字段
- [x] `replace_pptx_chart_data` 公开,first-batch 7 个 chart_type 全部支持,满足 ADR-0006 §5 invariant 1, 3, 4, 5, 7, 8(2 和 6 由 `python-pptx` 间接保证)
- [x] 6 个公开名字(`ChartSelector` / `ChartInventoryItem` / `SeriesData` / `ReplaceResult` / `inspect_pptx_charts` / `replace_pptx_chart_data`)加入 `__init__.py` 的 `__all__`,符合 ADR-0007 §1 五类生命周期标注
- [x] L1 测试 62 / 62 PASSED
- [x] `pyproject.toml` pin pandas / numpy / python-pptx / lxml / openpyxl 上限,baseline 跨机器可复现
- [x] `ISSUES.md` 沉淀:done(PCE-004 / 005),open(PCE-006 / 007 / 008),known limits(LIM-001 / 002 / 003)
- [x] 本 PRD 立即归档(ADR-0005 §3)
