# ADR-0006: 模板安全的 chart 数据更新作为引擎契约

- Status: Proposed
- Date: 2026-05-24
- Related:
  - ADR-0001
  - ADR-0002
  - ADR-0004
  - ../../../design/ppt-specialist/adr/0001-professional-data-update-ppt-platform.md

## Context

上层产品定位已经从"LLM 生成 PPT"收敛为"专业数据驱动的可编辑 PowerPoint 更新与生成平台"。这对 `pptchartengine` 的含义很具体:

- 用户经常已有企业 PPT 模板,并希望更新已有 chart 的数据。
- 模板里的 chart 往往没有业务命名,shape 历史复杂,不能让 LLM 靠视觉猜测。
- 专业用户要保留原 chart 的位置、大小、主题、样式、legend、axis 配置和可编辑性。
- 更新后的 chart 必须仍然是原生 PowerPoint chart,双击后可以看到新的嵌入 workbook 数据。
- `pptfi` / `ablemind` 需要一个确定性底层能力: inspect chart、定位 chart、替换数据、验证结果,而不是每次通过 OOXML 临时手写补丁。

ADR-0004 已经规定新 family 必须具备 round-trip metadata。它解决的是"引擎自己生成的 chart 未来还能 parse/restore"。但模板更新场景还多了一个问题:很多 chart 不是引擎生成的,没有 metadata,甚至不是我们熟悉的 family。引擎仍需要对常见原生 chart 提供安全的 technical inventory 与 data replacement 能力。

因此需要把"模板安全 chart 数据更新"提升为引擎级契约。

## Decision

`pptchartengine` 将把 **template-safe chart data update** 作为一等底层能力,但它只解决 chart 技术层问题,不接管产品层模板资产管理。

### 1. 引擎必须提供 chart inventory

引擎需要能从 `.pptx` 中提取 chart 技术清单,供 `pptfi` / `ablemind` 构建模板 manifest。

inventory 至少包含:

- slide index
- shape id / shape name
- chart index on slide
- chart part path
- embedded workbook relationship / part path
- chart type / plot types
- category count
- series count
- series names
- whether embedded workbook exists
- whether chart is replaceable by current engine
- warnings / risk flags

稳定 selector 的优先级:

1. explicit shape name / business tag,如果存在
2. shape id + chart part path
3. slide index + chart index,只作为 fallback

不得把"截图位置"或"视觉猜测"作为稳定 selector。

### 2. 引擎必须提供 chart data replacement

对可支持的原生 PowerPoint chart,引擎需要提供受控的 data replacement API。该 API 的核心语义是:

- 保留原 chart shape、位置、大小、主题、样式和 formatting。
- 只替换 categories / series values / embedded workbook 数据。
- 更新 chart XML cache 与 workbook 数据保持一致。
- 保留或重写 engine metadata,使更新后仍可 round-trip parse。
- 无法安全替换时 fail-loud,不得删除 chart 后静默重建。

推荐 public API 形态:

```python
inventory = inspect_pptx_charts("template.pptx")

result = replace_pptx_chart_data(
    input_pptx="template.pptx",
    output_pptx="updated.pptx",
    selector={"slide_index": 3, "shape_id": 17, "chart_part": "ppt/charts/chart2.xml"},
    categories=[...],
    series=[
        {"name": "Fund", "values": [...]},
        {"name": "Benchmark", "values": [...]},
    ],
)
```

实际命名可在实现 PRD 中调整,但语义必须保持: **定位已有 chart,原位替换数据,保留样式,结构化返回结果**。

### 3. 支持范围按垂直切片声明

首批只承诺常见原生 chart:

- line
- bar / column
- combo bar+line
- area
- pie
- scatter
- bubble

以下不作为首批承诺:

- linked external workbook chart
- pivot chart
- SmartArt / image chart
- 复杂 PowerPoint 插件图表
- 没有嵌入 workbook 且无法恢复数据结构的 chart
- 需要完整重建 plot area 的 custom overlay family

不支持时必须返回明确原因,例如:

- `unsupported_chart_type`
- `external_workbook_link`
- `missing_embedded_workbook`
- `series_shape_mismatch`
- `category_length_mismatch`
- `ambiguous_selector`

### 4. 引擎不拥有业务 slot 和模板资产

`pptchartengine` 输出 technical inventory,不输出业务 manifest。

不属于引擎的职责:

- user_id / tenant / sandbox / artifact 权限
- 模板库、模板版本、模板发布状态
- 业务 slot 命名,例如 `nav_performance_chart`
- required columns / 数据列语义映射
- 用户校准 UI
- 跨 slide 的报告级 workflow
- 生成 `template_manifest.json` 的产品 schema

这些由 `ablemind` / `pptfi` 负责。引擎只需要提供足够稳定、可验证的 chart 技术事实。

### 5. 验收标准

每个进入 public API 的 data replacement 能力必须满足:

1. inspect 能找到目标 chart,并返回稳定 selector。
2. replacement 后 `.pptx` 能被 PowerPoint 打开且不报修复警告。
3. chart shape id、position、size 在替换前后保持不变。
4. chart type、series count、series names 符合输入或显式 mapping。
5. embedded workbook 中的数据与输入一致。
6. chart XML cache 与 workbook 数据一致。
7. parse after replace 能恢复新 categories / series。
8. unsupported chart fail-loud,错误结构化,不产生伪成功 output。

这些断言并入 ADR-0002 的 L2/L3 质量闸门。涉及已有模板更新的 golden 应加入 pptfi `goldens/` 体系,但底层断言放在 `pptchartengine` 测试中。

## Consequences

正面:

- 上层 `ppt-specialist` 不需要靠视觉猜 chart,可以依赖确定性 inventory。
- 模板清洗 workflow 有稳定底座,能把脏 PPT 转成可更新模板资产。
- 引擎从"生成 chart"扩展为"维护 chart 数据资产",与 ADR-0004 的 round-trip 原则一致。
- 用户已有模板的样式和版式可以被保留,项目差异化更强。

负面:

- 引擎需要处理更多非本引擎生成的 chart,OOXML 兼容面扩大。
- 测试复杂度提高,需要覆盖真实 PowerPoint chart 变体。
- selector 设计若不严谨,后续模板版本迁移会有成本。
- 部分 chart 类型无法安全替换,必须接受 fail-loud 而不是过度承诺。

工程约束:

- 新增 template update 能力前必须先写 PRD,遵守 ADR-0005。
- 首批实现必须围绕一个真实模板 smoke test,不要先做泛化大框架。
- 不允许以"删除原 chart 后重建一个新 chart"冒充 template-safe update。
- `python-pptx` 可作为首批实现依赖,但 public contract 不能泄漏为 `python-pptx` 的对象模型。

## Alternatives Considered

### A. 把 chart 数据替换完全放到 `ablemind` tool 中

否决。`ablemind` 可以承载 tool schema 和权限,但 chart XML / workbook / parser / round-trip 是引擎知识。若完全放在 `ablemind`,会复制底层 chart 逻辑并绕开 ADR-0004。

### B. 每次更新都由 `pptfi` 重新生成页面

否决。重新生成适合 composer 工作流,不适合用户模板更新。它会破坏用户原模板中的样式、位置、动画、手工微调和机构品牌细节。

### C. 只支持引擎自己生成过且带 metadata 的 chart

否决作为最终方向。metadata chart 是最稳路径,但真实模板常常来自用户历史 PPT。引擎必须至少对常见原生 chart 提供 best-effort inventory 和替换能力。

### D. 让 LLM 直接编辑 chart OOXML

否决。OOXML 可以作为实现细节,不能作为 agent 常规接口。直接让 LLM 写 XML 不可审计、难以稳定复现,且容易生成 PowerPoint 打不开的文件。

### E. 使用外部 PowerPoint MCP / COM 自动化作为核心能力

否决作为核心依赖。桌面 PowerPoint 自动化可作为人工验证或企业内网增强,但本项目需要 headless、SaaS-safe、可测试的 Python library contract。
