# Architecture Decision Records (ADR)

本目录记录 `pptchartengine` 的架构决策。

## 什么是 ADR

ADR 记录"为什么这样设计"，不是"代码长什么样"。当一个决策同时具备以下特征时，就值得写成 ADR：

- 影响整个 repo 或长期演进方向
- 选择有 trade-off，未来需要被解释
- 推翻或修改时需要新的 ADR superseded 它

不要为每个小改动写 ADR；常规重构、bug 修复、依赖升级都不需要。

## 索引

| # | 标题 | 状态 | 日期 |
|---|------|------|------|
| 0001 | [pptchartengine 与 skill 的分离](0001-engine-skill-separation.md) | Accepted | 2026-05-01 |
| 0002 | [发版前的垂直质量闸门](0002-vertical-first-quality-gates.md) | Accepted | 2026-05-01 |
| 0003 | [Golden reference reports 作为质量北极星](0003-golden-reference-reports.md) | Accepted | 2026-05-01 |
| 0004 | [Round-trip metadata 作为核心约束](0004-round-trip-metadata-principle.md) | Accepted | 2026-05-01 |
| 0005 | [AI 协作工程纪律](0005-ai-collaboration-discipline.md) | Accepted | 2026-05-01 |
| 0006 | [模板安全的 chart 数据更新作为引擎契约](0006-template-safe-chart-data-update.md) | Proposed | 2026-05-24 |
| 0007 | [PowerPoint chart asset kernel 的边界与 API 定位](0007-chart-asset-kernel-boundary.md) | Proposed | 2026-05-24 |
| 0008 | [单人项目的轻量测试纪律](0008-lightweight-test-discipline.md) | Proposed | 2026-05-24 |

## 模板

新建 ADR 时复制以下骨架：

```markdown
# ADR-XXXX: 标题

- Status: Proposed / Accepted / Superseded by ADR-YYYY
- Date: YYYY-MM-DD
- Related: ADR-XXXX (if any)

## Context
为什么需要这个决策？背景是什么？

## Decision
决定做什么？要具体到可被验证的程度。

## Consequences
正面和负面影响分别列出。

## Alternatives Considered
还考虑过什么方案、为什么不选。
```

## 流程

- 重大决策先开 PR 提交 `Proposed` 状态的 ADR
- review + 讨论后 merge 时改为 `Accepted`
- 推翻旧 ADR 时新写一份并把旧的标记为 `Superseded by ADR-XXXX`，不要直接删除旧 ADR
- 编号单调递增，已发布的编号不能复用
