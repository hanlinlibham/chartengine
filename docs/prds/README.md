# Product Requirement Documents (PRD)

本目录记录 `pptchartengine` 的产品需求文档。

PRD 是 [ADR-0005](../adr/0005-ai-collaboration-discipline.md) 定义的"目的地文档"，用于在重大变更前与 AI 达成共享设计概念，避免静默误对齐。

## 什么时候写 PRD

参考 ADR-0005 触发条件：

- 新 family
- breaking API change
- 新 preset 系列
- metadata schema 演进
- 跨多个模块的重构

不需要 PRD：bug 修复、依赖升级、文档调整、单 family 内部小改。

## 状态与生命周期

```
in-progress → completed → archive 或删除
```

完成的 PRD 立即移到 [archive/](archive/)（参见 ADR-0005 反文档腐烂规则）。已 archive 的 PRD 不再被 AI 协作时读入工作集。

## 索引

| # | 标题 | 状态 | 创建日期 | 计划 |
|---|------|------|---------|------|
| [0001](0001-range-snapshot-visual-polish.md) | Range Snapshot 视觉质量精修 | Deferred to Slice 3 | 2026-04-19 | v0.3.0 |

## 模板

```markdown
# PRD-XXXX: 标题

- Status: in-progress / completed / deferred
- Created: YYYY-MM-DD
- Target: vX.Y.Z 或 Slice N
- Related: ADR-XXXX (if any)

## 问题陈述

## 影响范围
影响哪些 ADR、family、依赖项目

## 解决方案概要

## 测试策略
必须能映射到 L1 / L2 / L3 中具体哪些断言

## 参考材料
原始资料、相关 ADR、依赖文件

## 完成定义
明确"做完了"的判定标准
```

## 流程

- 触发条件之一时新建 PRD（命名 `NNNN-slug.md`）
- 实施过程中保持 PRD 与代码一致；如发现初始判断错误，先改 PRD 再改代码
- 完成后 review 完成情况、关键判断回写到对应 ADR
- 立即移到 `archive/`，并在本 README 索引中改状态
