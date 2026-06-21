# ADR-0007: PowerPoint chart asset kernel 的边界与 API 定位

- Status: Proposed
- Date: 2026-05-24
- Related:
  - ADR-0001
  - ADR-0002
  - ADR-0004
  - ADR-0006
  - ../../../design/ppt-specialist/adr/0001-professional-data-update-ppt-platform.md

## Context

AI presentation 项目已经很多。很多项目都能做到:

```text
prompt / document -> structured JSON -> slides -> PPTX or HTML export
```

如果 `pptchartengine` 也把自己定义成"AI 生成 PPT 的 Python 包",它会和通用 PPT 生成器、HTML-to-PPTX 项目、组件化 slide renderer 正面竞争。这个方向不是本仓库的优势。

本仓库已经通过 ADR-0001 明确自己是独立 chart engine,通过 ADR-0004 明确 round-trip metadata 是核心约束,通过 ADR-0006 明确 template-safe chart data update 是新增契约。现在需要把更上位的边界写清楚:

`pptchartengine` 的价值不是生成整份报告,而是把专业数据变成 **PowerPoint 原生、可编辑、可更新、可解析、可验证的 chart asset**。

这要求引擎长期死守确定性、可测试性和 PowerPoint 原生语义,而不是向上吸收 LLM、业务 slot、HTML 渲染或模板库产品逻辑。

## Decision

`pptchartengine` 定位为 **PowerPoint chart asset kernel**。

它不是:

- AI PPT 生成器
- 报告编排系统
- 模板管理系统
- HTML/交互式 Web Artifact 渲染器
- LLM prompt / agent runtime

它是:

> 结构化数据与 PowerPoint 原生 chart asset 之间的确定性转换层。

### 1. 引擎 API 围绕 chart asset 生命周期设计

Public API 应围绕五类生命周期能力组织:

1. `create`  
   从结构化 chart spec / DataFrame 生成原生 PowerPoint chart。

2. `parse`  
   从 `.pptx` / chart part 恢复 chart data、series config、layout info 和 metadata。

3. `metadata`  
   持久化 chart family、series mapping、data schema、style intent、engine version 等语义信息。

4. `inspect`  
   从用户已有 `.pptx` 中提取 chart inventory 和 replaceability facts。

5. `replace`  
   对已有 chart 原位替换数据,并保留 chart shape、样式、位置、尺寸和可编辑性。

报告级 API、LLM API、模板库 API 不进入 `pptchartengine` public surface。

### 2. ChartSpec 是技术 schema,不是业务 report spec

引擎可以引入统一 `ChartSpec` / `ChartData` / `ChartMetadata` 概念,但这些 schema 必须保持技术层语义:

- chart family / chart type
- categories
- series
- axis roles
- grouping / plot type
- style hints / style overrides
- metadata payload
- source data shape

不得包含:

- 用户 prompt
- 业务 slot 名称,例如 `nav_performance_chart`
- report section / slide narrative
- template_id / user_id / tenant_id
- 数据源连接配置
- HTML interaction config

`pptfi` 可以定义更高层的 `SemanticReportSpec`,再翻译成一个或多个 engine `ChartSpec`。`ablemind` 可以用 Pydantic 定义 tool schema。`pptchartengine` 本身应保持轻量依赖:如果需要强校验,优先使用标准库 dataclass / typed dict / JSON-schema 文档;是否引入 Pydantic 必须单独决策,不能因为上层 tool schema 使用 Pydantic 就让 kernel 直接依赖它。

### 3. Metadata 存储采用分层策略

metadata 的目标是让 chart 成为可继承资产,支撑 round-trip、模板更新和审计。

当前已接受的最低保障仍是 ADR-0004: metadata 写入 embedded workbook 隐藏 sheet。

未来 metadata 存储策略按优先级演进:

1. embedded workbook hidden sheet  
   当前默认路径,与 chart data 生命周期绑定,最适合 chart-level round-trip。

2. chart semantic anchor / invisible shape  
   用于纯 shape composition 或没有标准 chart container 的 family。

3. custom XML part  
   作为未来增强方向,用于 deck-level 或跨 chart metadata registry。引入前必须先写 PRD,明确 PowerPoint/Keynote/Google Slides 兼容性和 round-trip 迁移策略。

4. shape alternative text / name  
   只作为 selector hint 或人类可读辅助,不得作为唯一 metadata source of truth。

### 4. Replace in place 不能退化成重建

`replace` 的语义必须严格区别于 `create`:

- `create`: 生成新的 chart asset。
- `replace`: 在已有 chart asset 上替换数据。

`replace` 默认不得删除原 chart 后重建新 chart。只有在调用方显式选择 rebuild mode,且返回结果清楚标记为 `rebuilt`,才允许 fallback 重建。

template-safe replace 的默认验收来自 ADR-0006:

- shape identity 不变
- position / size 不变
- chart type 不变或显式 mapping
- workbook 数据与输入一致
- chart XML cache 与 workbook 一致
- parse after replace 成功
- unsupported case fail-loud

### 5. HTML 双输出属于上层,但可复用引擎语义

本项目可以支持"同一语义 spec 同时输出可编辑 PPTX + 高交互 HTML",但 HTML renderer 不属于 `pptchartengine`。

边界如下:

- `pptchartengine` 维护 chart-level semantic model 和 PowerPoint renderer。
- `pptfi` 负责把 report-level semantic spec 翻译成 PPTX renderer 调用和 HTML renderer 调用。
- HTML renderer 可以复用 engine 的 chart family catalog、data normalization、theme tokens,但不能让 engine 直接输出 ECharts/HTML。

这样可以保持 kernel 的确定性,同时让上层产品做交互式 HTML。

### 6. 垂直优先,不追求通用 chart 大而全

首批优势应集中在专业报告高频图表:

- line / bar / combo
- waterfall
- scatter / bubble
- range snapshot
- performance compare
- event timeline
- attribution / factor / style / score / concentration 类金融语义图族

不以通用图表库为目标,不追求覆盖所有 PowerPoint chart type。每个进入 public API 的 family 都必须符合 ADR-0004 round-trip 约束和 ADR-0002 质量闸门。

## Consequences

正面:

- 项目避开通用 AI PPT 生成器竞争,聚焦可编辑 PowerPoint chart asset 这一深水区。
- Kernel API 更稳定,可被 `pptfi`、skills、MCP、ablemind 等多个上层复用。
- ChartSpec 不被 prompt、业务 slot、HTML 交互污染,长期更可维护。
- Replace-in-place 与 inspect 能成为开源项目的关键差异化。
- 与金融、咨询、投研、投行等重度 PowerPoint 场景更匹配。

负面:

- 短期看起来不如"输入一句话生成整套 PPT"有展示冲击力。
- 需要上层 `pptfi` / `ablemind` 才能呈现完整产品体验。
- ChartSpec 过窄会让上层翻译成本增加;过宽又会污染 kernel,需要持续校准。
- 不引入 Pydantic 等重校验依赖会让 kernel 需要自己维护清晰的 schema 文档和测试。

工程约束:

- 新 public API 必须说明属于 create / parse / metadata / inspect / replace 哪一类。
- 新 schema 字段必须判断是否属于 chart 技术层;如果是业务语义,应上移到 `pptfi`。
- 新 metadata 位置必须说明 source of truth 和迁移策略。
- 新 HTML 或 report-level 需求不得直接改 `pptchartengine` public API,除非能证明它是 chart-level 语义。

## Alternatives Considered

### A. 把 `pptchartengine` 做成完整 AI PPT 生成器

否决。该方向已有大量项目和商业产品,且会把 LLM、模板、页面布局、HTML、用户体验逻辑全部压入 kernel,破坏本仓库的可测试性和可复用性。

### B. 把 `pptchartengine` 做成通用 chart library

否决。通用 chart library 会追求更多 chart type、更多视觉主题、更多输出格式,但本项目真正壁垒是 PowerPoint 原生可编辑 chart 的 round-trip 和 data update。

### C. 让 `pptfi` 直接承担 chart create / replace

否决。`pptfi` 应该做 report runtime 和 semantic spec orchestration。如果 chart 内核逻辑上移,会导致 `pptfi` 和未来其他上层重复实现 chart OOXML / metadata / parser 逻辑。

### D. 所有 schema 都用上层 Pydantic model

否决作为默认内核策略。Pydantic 适合 `pptfi` 和 `ablemind` 的 API/tool schema,但 kernel 需要控制依赖面。若未来引入 Pydantic,必须证明收益超过依赖和 public model 兼容成本。

### E. 用 custom XML part 立即替代 embedded workbook metadata

暂不采用。custom XML part 可能更稳定,但需要兼容性验证和迁移策略。当前 hidden workbook sheet 已支撑 chart-level round-trip;custom XML part 作为后续 PRD 评估项。
