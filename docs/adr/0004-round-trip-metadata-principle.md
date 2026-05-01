# ADR-0004: Round-trip metadata 作为核心约束

- Status: Accepted (formalizes existing design)
- Date: 2026-05-01

## Context

引擎当前已经按 round-trip 设计：图表生成时把语义 metadata 写入嵌入 workbook 的隐藏 sheet，解析端优先读 metadata 而非依赖 XML 结构猜测。

这个设计支撑：

- "图表是可继承的中间资产"——下次生成可基于 parse 结果继续修改
- LLM 驱动的图表编辑流（Claude 读 → 改 → 写回）
- `pptfi` 模板替换、报告再加工、增量更新

但新加 family 时容易忽略 round-trip 约束，只实现 `create_*`。这种 family 一旦合入会破坏整体闭环，导致 parser 出现 family-specific 启发式分支，长期维护成本上升。

需要把这个事实上的设计原则提升为显式约束。

## Decision

将 round-trip 提升为引擎核心约束。每个进入 public API 的图族必须实现完整闭环：

1. `create_<family>_chart()` —— 生成端
2. metadata 写入 workbook 隐藏 sheet —— 语义持久化
3. `parse_<family>_chart()` / `parse_<family>_from_pptx()` —— 解析端
4. `restore_<family>_dataframe()` —— 数据反向恢复（适用时）
5. 契约测试覆盖闭环：`create → save → reload → parse → 语义等价`

不满足该闭环的 family 只能进 internal API（`_*` 前缀），不导出在 `__init__.py` 的 `__all__` 中。

**例外**：

- 纯 shape composition（如 `ranked_tile_matrix`、`heatmap_matrix`、各类 panel）没有标准 chart container 的 family，可只保留 layout_info 级别的 round-trip
- candlestick / OHLC 等 README 已声明为"暂未支持"的图类，本就不在引擎范围
- 新 family 的实验阶段允许只有 `create_*`，但必须在 internal API 内部，并标记为 experimental

## Consequences

正面：

- 引擎对外 API 始终满足"生成后还能继续改"
- 与 LLM-driven 编辑路径天然兼容
- parser 不依赖脆弱的 XML 启发式
- 新人对"什么算 first-class family"有清晰判定

负面：

- 新 family 开发成本提高约 30-50%
- metadata schema 演进需要 backward compatibility 策略
- 部分 OOXML-only 高级图表（如原生 sankey）短期进不来
- experimental → public 的晋升流程需要明确

## Alternatives Considered

**只保证 create，parse 由消费方实现**：
违反"chart as data asset"原则，外部消费方会重复实现 parse 逻辑、且容易写出与生成端不一致的解析。否决。

**双 metadata schema（旧 / 新）并存**：
增加复杂度，没有实际收益。当前单一 schema 已能满足需求；未来确实需要演进时，再走 schema versioning。
